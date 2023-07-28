VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchemaMsSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbSchema class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' This requires us to use all the public methods and properties of the implemented class
' which keeps all the server classes consistent in how they are used in the export
' process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbSchema


' Handle local variables
Private Type udtThis
    blnInitialized As Boolean
    strName As String
    strBaseFolder As String
    strConnect As String
    strUserID As String
    strPassword As String
    varFilters As Variant
End Type
Private this As udtThis


' Dictionaries representing modified items, and all items
Private m_AllItems As Dictionary
Private m_ModifiedItems As Dictionary
Private m_Index As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Export
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Export(blnFullExport As Boolean, Optional strAlternatePath As String)

    Dim rstObjects As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim strItem As String
    Dim dItem As Dictionary
    Dim strHash As String
    Dim varItem As Variant
    Dim dblStart As Double

    ' Make sure we initialize before running the export
    If Not this.blnInitialized Then Exit Sub

    ' Make sure we have already performed a scan of the database objects
    If m_AllItems Is Nothing Then ScanDatabaseObjects

    ' If there are no new changes from the database side, we may not need to update the index.
    If (m_ModifiedItems.Count = 0) And (m_Index.Count = m_AllItems.Count) Then
        ' Database matches the current index.
    Else
        If m_ModifiedItems.Count = 0 Then
            Log.Add "     Updating index", , , , , True
        Else
            Log.Add "     Exporting " & m_ModifiedItems.Count & " objects...", , , , , True
            Log.ProgMax = m_ModifiedItems.Count
            Log.Flush
        End If

        ' Open database connection
        Set conn = New ADODB.Connection
        conn.Open this.strConnect, this.strUserID, this.strPassword

        ' Export the modified items
        For Each varItem In m_ModifiedItems.Keys
            dblStart = Perf.MicroTimer
            Set dItem = m_ModifiedItems(varItem)
            strItem = varItem
            strHash = ExportObject(dItem("type_desc"), dItem("schema"), dItem("name"), this.strBaseFolder & varItem, conn)
            If Len(strHash) Then dItem("hash") = strHash
            Log.Add "    Exported " & varItem & " in " & Round(Perf.MicroTimer - dblStart, 2) & " seconds.", Options.ShowDebug
            ' Update record in index
            UpdateIndex strItem, dItem
            Log.Increment
            ' Check for canceled operation
            If Log.ErrorLevel = eelCritical Then Exit For
        Next varItem

        ' Close database connection
        conn.Close
        Set conn = Nothing

        ' Purge missing index entries
        For Each varItem In m_Index
            If Not m_AllItems.Exists(varItem) Then
                m_Index.Remove varItem
            End If
        Next varItem

        ' For performance reasons (especially on large databases with thousands of objects) we only want
        ' to convert this back to JSON and write the file if we actually made changes.
        WriteFile BuildJsonFile(TypeName(Me), m_Index, "Version Control System Schema Index"), this.strBaseFolder & "vcs-index.json"
    End If

    ' Now, loop back through the files and remove any file that is not represented by
    ' the list of objects returned from the server.
    VerifyFiles

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Export the object definition and return a hash of the content.
'---------------------------------------------------------------------------------------
'
Private Function ExportObject(strType, strSchema As String, strName As String, strPath As String, ByRef oConn As ADODB.Connection) As String

    Static intUseSP As Integer

    Dim strSqlDef
    Dim strDefinition As String
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim strFullName As String
    Dim cmd As ADODB.Command
    Dim strCmdTemplate As String

    ' Prepare template statement for sp_GetDDL to work around VARCHAR(MAX) issue
    ' with many SQL Server ODBC drivers.
    strCmdTemplate = _
        "DECLARE @table TABLE (item text) " & _
        "INSERT INTO @table exec sp_GetDDL N'{name}' " & _
        "SELECT * FROM @table"

    ' Cache whether or not
    'if intusesp=0 then intusesp =

    strFullName = "[" & strSchema & "].[" & strName & "]"

    ' Determine how to export this type of object
    Select Case strType
        Case "USER_TABLE", "VIEW", "SYNONYM", "SQL_STORED_PROCEDURE", _
            "SQL_SCALAR_FUNCTION", "SQL_INLINE_TABLE_VALUED_FUNCTION", "SQL_TABLE_VALUED_FUNCTION"
            strSqlDef = Replace(strCmdTemplate, "{name}", strFullName)
            'strSqlDef = "exec sp_help " & strFullName & ""

        Case "TYPE_TABLE", "SEQUENCE_OBJECT", "SERVICE_QUEUE", "SYSTEM_TABLE", "INTERNAL_TABLE"
            ' Unsupported non-dependent objects

    End Select

    ' Sanity check
    If Len(strSqlDef) Then
        Perf.OperationStart "Run sp_GetDDL on " & strType
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = oConn
            .CommandText = strSqlDef
            Set rst = .Execute
        End With

        ' Get secondary recordset with object definition record
        Set rst2 = rst.NextRecordset
        With rst2
            If Not .EOF Then strDefinition = Nz(.Fields(0))
            If strDefinition = vbNullString Then
                If FSO.FileExists(strPath) Then DeleteFile strPath
            Else
                ' Export to file
                WriteFile strDefinition, strPath
            End If
            .Close
        End With
        Perf.OperationEnd
    End If

    ' Return hash from object definition (if found)
    ExportObject = GetSimpleHash(strDefinition)

End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateIndex
' Author    : Adam Waller
' Date      : 7/19/2023
' Purpose   : Update the entry in the index using values from the recordset
'---------------------------------------------------------------------------------------
'
Private Sub UpdateIndex(strItem As String, dObject As Dictionary)

    Dim dItem As Dictionary

    Perf.OperationStart "Update Schema Index"
    ' See if we are working with an existing entry

    If m_Index.Exists(strItem) Then
        Set dItem = m_Index(strItem)
    Else
        ' Add the new entry
        Set dItem = New Dictionary
        m_Index.Add strItem, dItem
    End If

    ' Update the meta values
    dItem("FileDate") = Now()
    dItem("LastModified") = dObject("LastModified")
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PassesFilter
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Returns true if this item passed any user-defined filter rules.
'           : The current implementation processes rules sequentially, applying each
'           : rule in order. Last matching rule will apply to the object.
'---------------------------------------------------------------------------------------
'
Private Function PassesFilter(strItem As String) As Boolean

    Dim blnPass As Boolean
    Dim lngRule As Long
    Dim strRule As String

    ' Loop through rules
    For lngRule = 0 To UBound(this.varFilters)
        strRule = Trim(this.varFilters(lngRule))
        Select Case Left(strRule, 1)
            Case "#", vbNullString
                ' Ignore comments and blank lines
            Case "!"
                ' Negative rule (do not include)
                If strItem Like Mid(strRule, 2) Then blnPass = False
            Case Else
                ' Positive rule
                If strItem Like strRule Then blnPass = True
        End Select
    Next lngRule

    ' Return final result
    PassesFilter = blnPass

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanDatabaseObjects
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Scan the database objects for any changed items
'---------------------------------------------------------------------------------------
'
Private Sub ScanDatabaseObjects()

    Dim dFile As Dictionary
    Dim dObjects As Dictionary
    Dim dteLast As Date
    Dim rstObjects As ADODB.Recordset
    Dim blnUseGetDDL As Boolean
    Dim conn As ADODB.Connection
    Dim strItem As String
    Dim strPath As String
    Dim blnExport As Boolean
    Dim blnModified As Boolean
    Dim dItem As Dictionary
    Dim blnHasChanges As Boolean

    ' Clear module level objects
    Set m_AllItems = Nothing
    Set m_ModifiedItems = Nothing
    Set m_Index = Nothing

    ' Make sure we initialize before running the scan
    If Not this.blnInitialized Then Exit Sub

    ' Load index
    Set dFile = ReadJsonFile(IndexPath)
    If Not dFile Is Nothing Then
        Set m_Index = dFile("Items")
    Else
        Set m_Index = New Dictionary
    End If

    ' Return list of objects from the server
    Set conn = New ADODB.Connection
    Perf.OperationStart "Retrieve SQL Objects"
    conn.Open this.strConnect, this.strUserID, this.strPassword
    Set rstObjects = conn.Execute(CodeDb.QueryDefs("qryMsSqlServerObjects").SQL)
    Perf.OperationEnd

    ' Initialize dictionaries
    Set m_AllItems = New Dictionary
    Set m_ModifiedItems = New Dictionary

    ' Loop through objects, building dictionary of items that match our filter.
    Perf.OperationStart "Loop through SQL objects"
    With rstObjects
        Do While Not .EOF

            ' Build item path and full path to source file
            strItem = Nz(!Folder) & PathSep & GetSafeFileName(Nz(!Name)) & ".sql"
            strPath = this.strBaseFolder & strItem

            ' See if we pass the filter
            If PassesFilter(strItem) Then

                ' Add all objects to full collection
                Set dItem = New Dictionary
                dItem("LastModified") = Nz(!last_modified)
                m_AllItems.Add strItem, dItem

                ' Reset flag
                blnModified = True

                ' See if this item has been modified
                If m_Index.Exists(strItem) Then
                    blnModified = (dItem("LastModified") <> Nz(!last_modified))
                End If

                ' Build dictionary of modified objects
                If blnModified Then
                    Set dItem = CloneDictionary(dItem)
                    dItem("type_desc") = Nz(!type_desc)
                    dItem("schema") = Nz(!schema)
                    dItem("name") = Nz(!Name)
                    m_ModifiedItems.Add strItem, dItem
                End If
            End If

            ' Move to next object
            .MoveNext
        Loop
        .Close
    End With
    Perf.OperationEnd

    ' Close connection
    conn.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyFiles
' Author    : Adam Waller
' Date      : 7/24/2023
' Purpose   : Compare the files to the index, ensuring that we don't have any orphaned
'           : files, or files that are now different from the last exported version.
'           : (Comparing index hashes of files where the modified dates differ from the
'           :  the export date.)
'---------------------------------------------------------------------------------------
'
Public Function VerifyFiles()

    Dim oFld As Scripting.Folder
    Dim oFile As Scripting.File
    Dim dFolders As Dictionary
    Dim varKey As Variant
    Dim strFolder As String
    Dim strName As String
    Dim strItem As String
    Dim strPath As String
    Dim lngTotal As Long
    Dim dFile As Dictionary
    Dim dIndex As Dictionary
    Dim strHash As String

    ' Build a collection of subfolders to check
    Set dFolders = New Dictionary
    For Each oFld In FSO.GetFolder(this.strBaseFolder).SubFolders
        Select Case oFld.Name
            Case "views", "tables", "procedures", "functions", "types", "sequences", "synonymns"
                ' Keep this list current with the pass-through query that defines the folder names.
                dFolders.Add oFld.Name, oFld.Files.Count
                lngTotal = lngTotal + dFolders(oFld.Name)
        End Select
    Next oFld

    ' Load the current index, falling back to a new dictionary object
    strPath = this.strBaseFolder & "vcs-index.json"
    If FSO.FileExists(strPath) Then
        Set dFile = ReadJsonFile(strPath)
        If Not dFile Is Nothing Then Set dIndex = dFile("Items")
    End If
    If dIndex Is Nothing Then Set dIndex = New Dictionary

    ' Loop through subfolders, checking files
    For Each varKey In dFolders.Keys
        strFolder = varKey
        For Each oFile In FSO.GetFolder(this.strBaseFolder & strFolder).Files
            strName = oFile.Name
            ' Filter by file extension
            If StrComp(FSO.GetExtensionName(strName), "sql", vbTextCompare) = 0 Then
                ' Check for index entry
                strItem = strFolder & PathSep & strName
                If dIndex.Exists(strItem) Then
                    ' Check the modified date
                    If DatesClose(dNZ(dIndex(strItem), "FileDate"), oFile.DateLastModified) Then
                        ' Looks like we match the dates from the last export. Should be pretty safe
                        ' to assume that this file is the same as when we exported it.
                    Else
                        ' The export dates don't match. Check the content hash.
                        ' (It should be faster to compare the file hash with the index
                        '  than to export the object again from the server.)
                        strHash = GetSimpleHash(ReadFile(oFile.Path))
                        If strHash = dNZ(dIndex(strItem), "FileHash") Then
                            ' Update modified date in index to match file.
                            ' (No need to export again)
                            dIndex(strItem)("FileModified") = oFile.DateLastModified
                        End If
                    End If
                Else
                    ' File does not exist in the index. Purge the orphaned file.
                    oFile.Delete
                End If
            End If
        Next oFile
    Next varKey

End Function




Private Function PurgeOrphanedObjects()

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_GetChangeCount
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Return count of modified objects
'---------------------------------------------------------------------------------------
'
Private Function IDbSchema_ObjectCount(blnModifiedOnly As Boolean) As Long
    If m_AllItems Is Nothing Then ScanDatabaseObjects
    If m_AllItems Is Nothing Then Exit Function
    IDbSchema_ObjectCount = IIf(blnModifiedOnly, m_ModifiedItems.Count, m_AllItems.Count)
End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Initialize
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Initialize the database schema
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Initialize(dInstance As Scripting.IDictionary)

    Dim strFilter As String
    Dim varRules As Variant

    ' Build filters
    strFilter = dNZ(dInstance, "Filter")
    If Trim(Replace(strFilter, vbCrLf, vbNullString)) = vbNullString Then
        ' Blank lines, or nothing defined
        varRules = Array("*")
    Else
        ' Create array of rules (we will skip comments and blank lines later)
        varRules = Split(strFilter, vbCrLf)
    End If

    With this
        .strName = dNZ(dInstance, "Name")
        .strConnect = dNZ(dInstance, "Connect")
        .strBaseFolder = Options.GetExportFolder & "databases\" & GetSafeFileName(.strName) & PathSep
        .varFilters = varRules
        .blnInitialized = (Len(.strConnect))
    End With
End Sub



Private Property Get IDbSchema_BaseFolder() As String
    IDbSchema_BaseFolder = this.strBaseFolder
End Property


Private Function IndexPath() As String
    IndexPath = this.strBaseFolder & "vcs-index.json"
End Function


Private Property Get IDbSchema_Name() As String

End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_ServerType
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Return server type
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_ServerType() As eDatabaseServerType
    IDbSchema_ServerType = estMsSql
End Property


Private Property Get IDbSchema_TypeDescription() As Variant
    IDbSchema_TypeDescription = "Microsoft SQL Server"
End Property


Private Property Get IDbSchema_ServerVersion() As String

End Property
