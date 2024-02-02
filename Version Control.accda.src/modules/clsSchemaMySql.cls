VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchemaMySql"
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


' Handle local variables
Private Type udtThis
    blnInitialized As Boolean
    strName As String
    strBaseFolder As String
    strConnect As String
    blnUtcTime As Boolean
    strUserID As String
    strPassword As String
    colFilters As Collection
End Type
Private this As udtThis

' Dictionaries representing modified items, and all items
Private m_Files As Dictionary
Private m_AllItems As Dictionary
Private m_ModifiedItems As Dictionary

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the server classes consistent in how they are used in the export
' process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbSchema


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Export
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Export DDL representations of the external database objects.
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Export(blnFullExport As Boolean, Optional strAlternatePath As String)

    Dim conn As ADODB.Connection
    Dim strItem As String
    Dim dItem As Dictionary
    Dim varItem As Variant
    Dim dblStart As Double
    Dim strPath As String
    Dim blnChanges As Boolean
    Dim dFolders As Dictionary

    ' Make sure we initialize before running the export
    If Not this.blnInitialized Then Exit Sub

    ' Make sure we have already performed a scan of the database objects
    If m_Files Is Nothing Then ScanFiles
    If m_AllItems Is Nothing Then ScanDatabaseObjects

    ' If there are no new changes found, we may not need to export anything
    If (m_ModifiedItems.Count = 0) And (m_Files.Count = m_AllItems.Count) Then
        ' Database matches the current set of files
    Else
        blnChanges = True
        If m_ModifiedItems.Count = 0 Then
            Log.Add "     Verifying files...", , , , , True
        Else
            Log.Add "     Exporting " & m_ModifiedItems.Count & " objects...", , , , , True
            Log.ProgMax = m_ModifiedItems.Count
            Log.Flush
        End If

        ' Open database connection
        Set conn = GetNewOpenConnection
        If conn Is Nothing Then Exit Sub

        ' --------------------------------------------------
        ' FIRST PASS - Export changed/new database objects
        ' --------------------------------------------------
        For Each varItem In m_ModifiedItems.Keys

            ' Time the export of each item
            dblStart = Perf.MicroTimer
            Set dItem = m_ModifiedItems(varItem)
            strItem = varItem
            ExportObject dItem("folder"), dItem("schema"), dItem("name"), dItem("last_modified"), dItem("hash"), CStr(varItem), conn
            Log.Add "    Exported " & varItem & " in " & Round(Perf.MicroTimer - dblStart, 2) & " seconds.", Options.ShowDebug
            Log.Increment
            ' Check for canceled operation
            If Log.ErrorLevel = eelCritical Then Exit For
        Next varItem

        ' Close database connection
        conn.Close
        Set conn = Nothing
    End If

    ' --------------------------------------------------
    ' SECOND PASS - Remove orphaned files
    ' --------------------------------------------------
    Perf.OperationStart "Clear Orphaned Schema Files"
    For Each varItem In m_Files
        If Not m_AllItems.Exists(varItem) Then
            strPath = this.strBaseFolder & varItem
            If FSO.FileExists(strPath) Then
                Log.Add " - Removed orphaned file: " & varItem, False
                DeleteFile strPath
                blnChanges = True
            End If
        End If
    Next varItem
    Perf.OperationEnd

    ' --------------------------------------------------
    ' THIRD PASS - Remove empty source folders
    ' --------------------------------------------------
    If blnChanges Then
        Set dFolders = GetBaseFolders
        For Each varItem In dFolders
            strPath = this.strBaseFolder & varItem
            If FSO.FolderExists(strPath) Then
                If FSO.GetFolder(strPath).Files.Count = 0 Then
                    ' Remove empty component subfolders
                    FSO.DeleteFolder strPath
                End If
            End If
        Next varItem
    End If

    ' Save updated index
    SaveUpdatedIndex

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Export the object definition to a file
'---------------------------------------------------------------------------------------
'
Private Function ExportObject(strType, strSchema As String, strName As String, dteModified As Date, strHash As String, strFile As String, ByRef oConn As ADODB.Connection) As String

    Dim strSqlDef
    Dim strDefinition As String
    Dim rst As ADODB.Recordset
    Dim strFullName As String
    Dim cmd As ADODB.Command
    Dim intField As Integer
    Dim strPath As String

    ' Build full name of SQL object
    strFullName = "`" & strSchema & "`.`" & strName & "`"

    ' Determine how to export this type of object
    Select Case strType
        Case "tables":      strSqlDef = "show create table " & strFullName
        Case "views":       strSqlDef = "show create view " & strFullName
        Case "procedures":  strSqlDef = "show create procedure " & strFullName
        Case "functions":   strSqlDef = "show create function " & strFullName
        Case "triggers":    strSqlDef = "show create trigger " & strFullName
        Case Else
            ' Unsupported type
            Log.Error eelError, "Unsupported object type: " & strType, ModuleName(Me) & ".ExportObject"
            Exit Function
    End Select

    ' Sanity check
    If Len(strSqlDef) Then
        Perf.OperationStart "Get DDL for " & strType
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = oConn
            .CommandText = strSqlDef
            Set rst = .Execute
        End With

        ' Look up definition from recordset
        With rst
            If Not .EOF Then
                ' Definition might be in second or third column
                For intField = 1 To 2
                    If StartsWith(.Fields(intField).Name, "Create ") Then
                        strDefinition = Nz(.Fields(intField))
                        Exit For
                    End If
                Next intField
            End If
            .Close
        End With

        ' Write object definition to file
        strPath = this.strBaseFolder & strFile
        If strDefinition = vbNullString Then
            If FSO.FileExists(strPath) Then DeleteFile strPath
        Else
            ' Export to file
            WriteFile strDefinition, strPath
            If strType = "views" Then
                ' Use hash index
                m_AllItems(strFile) = strHash
            Else
                ' Set file modified date to match SQL object
                SetFileDate strPath, dteModified, Not this.blnUtcTime
            End If
        End If

        Perf.OperationEnd
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanDatabaseObjects
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Scan the database objects for any changed items
'---------------------------------------------------------------------------------------
'
Private Sub ScanDatabaseObjects()

    Dim rstObjects As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim strItem As String
    Dim strSchema As String
    Dim strPath As String
    Dim blnModified As Boolean
    Dim dItem As Dictionary
    Dim lngCount As Long
    Dim strHash As String

    ' Clear module level objects
    Set m_AllItems = Nothing
    Set m_ModifiedItems = Nothing

    ' Make sure we initialize before running the scan
    If Not this.blnInitialized Then Exit Sub

    ' Open database connection
    Set conn = GetNewOpenConnection
    If conn Is Nothing Then Exit Sub

    ' Return list of objects from the server
    Perf.OperationStart "Retrieve SQL Objects"
    Set rstObjects = conn.Execute(CodeDb.QueryDefs("qryMySqlServerObjects").SQL)
    Perf.OperationEnd

    ' Initialize dictionaries
    Set m_AllItems = New Dictionary
    Set m_ModifiedItems = New Dictionary

    ' Loop through objects, building dictionary of items that match our filter.
    Perf.OperationStart "Loop through MySQL objects"
    With rstObjects
        Do While Not .EOF

            ' Build item path and full path to source file
            strSchema = Nz(!schema, "schema") & "."
            strItem = Nz(!Folder) & PathSep & GetSafeFileName(strSchema & Nz(!Name)) & ".sql"
            strPath = this.strBaseFolder & strItem
            strHash = vbNullString

            ' See if we pass the filter
            If PassesSchemaFilter(strItem, this.colFilters) Then

                ' Use modified date to match file, if possible.
                ' Some objects don't store a modified date, so use a hash for those.
                If Nz(!last_modified) = vbNullString Then
                    ' Create hash from definition
                    strHash = "hash:" & GetStringHash(Nz(!definition))
                End If

                ' Check for modification/new item
                If m_Files.Exists(strItem) Then
                    ' Flag as modified if the dates or hash don't match
                    blnModified = (Nz2(strHash, Nz(!last_modified)) <> CStr(m_Files(strItem)))
                Else
                    ' File does not yet exist
                    blnModified = True
                End If

                ' Add all objects to full collection
                m_AllItems.Add strItem, Nz2(strHash, Nz(!last_modified))

                ' Build dictionary of modified objects
                If blnModified Then
                    Set dItem = New Dictionary
                    dItem("folder") = Nz(!Folder)
                    dItem("schema") = Nz(!schema)
                    dItem("name") = Nz(!Name)
                    dItem("hash") = strHash
                    dItem("last_modified") = CDate(Nz(!last_modified, 0))
                    m_ModifiedItems.Add strItem, dItem
                End If
            End If

            ' Move to next object
            lngCount = lngCount + 1
            .MoveNext
        Loop
        .Close
    End With
    Perf.OperationEnd lngCount

    ' Close connection
    conn.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveUpdatedIndex
' Author    : Adam Waller
' Date      : 8/7/2023
' Purpose   : Save an updated version of the index (after scanning for changes and
'           : exporting any changed objects)
'---------------------------------------------------------------------------------------
'
Private Function SaveUpdatedIndex()

    Dim varKey As Variant
    Dim strValue As String
    Dim dIndex As Dictionary
    Dim strPath As String

    ' Only run this if we actually have an index to process
    If m_AllItems Is Nothing Then Exit Function

    ' Loop through all items, building a dictionary of hashed items.
    ' (These items don't have modified dates in MySQL, and must be tracked via hash.)
    Set dIndex = New Dictionary
    For Each varKey In m_AllItems.Keys
        strValue = m_AllItems(varKey)
        If StartsWith(strValue, "hash:") Then
            ' Add to index
            dIndex(varKey) = strValue
        End If
    Next varKey

    ' Update the saved index file
    strPath = this.strBaseFolder & "vcs-index.json"
    If dIndex.Count = 0 Then
        ' Remove index when no longer needed.
        If FSO.FileExists(strPath) Then DeleteFile strPath
    Else
        ' Save the rebuilt index
        WriteFile BuildJsonFile(TypeName(Me), dIndex, "Version Control System Schema Index for MySQL"), strPath
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNewOpenConnection
' Author    : Adam Waller
' Date      : 8/3/2023
' Purpose   : Return a new, open ADODB connection
'---------------------------------------------------------------------------------------
'
Private Function GetNewOpenConnection() As ADODB.Connection

    Dim oConn As ADODB.Connection

    Set oConn = New ADODB.Connection
    With oConn
        LogUnhandledErrors
        On Error Resume Next
        If Len(this.strUserID) Then
            .Open this.strConnect, this.strUserID, this.strPassword
        Else
            .Open this.strConnect
        End If
    End With

    ' Check for any connection error
    If CatchAny(eelError, "Unable to connect to " & this.strName, ModuleName(Me)) Then
        Log.Add "Connection string: " & this.strConnect, False
    Else
        ' Return open connection
        Set GetNewOpenConnection = oConn
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanFiles
' Author    : Adam Waller
' Date      : 7/28/2023
' Purpose   : Scan the files to get a list of source files that should match the current
'           : database objects. For performance reasons, we won't apply the filter here
'           : but will check the filter later when removing orphaned objects.
'---------------------------------------------------------------------------------------
'
Private Function ScanFiles()

    Dim oFld As Scripting.Folder
    Dim dBaseFolders As Dictionary
    Dim dFiles As Dictionary
    Dim varKey As Variant
    Dim strKey As String
    Dim strValue As String
    Dim strFolder As String
    Dim dFile As Dictionary
    Dim dIndex As Dictionary
    Dim strPath As String

    ' Reset module-level dictionary
    Set m_Files = New Dictionary

    ' Load any existing index file of hashes for object that don't store modified dates.
    strPath = this.strBaseFolder & "vcs-index.json"
    If FSO.FileExists(strPath) Then
        Set dFile = ReadJsonFile(strPath)
        If Not dFile Is Nothing Then Set dIndex = dFile("Items")
    End If
    If dIndex Is Nothing Then Set dIndex = New Dictionary

    ' Build a collection of subfolders and files with modified dates
    ' (Using the Windows API for faster scanning and more accurate dates)
    Set dBaseFolders = GetBaseFolders
    If FSO.FolderExists(this.strBaseFolder) Then
        For Each oFld In FSO.GetFolder(this.strBaseFolder).SubFolders
            strFolder = oFld.Name
            If dBaseFolders.Exists(strFolder) Then
                ' Get dictionary of files with modified dates
                Set dFiles = GetFileList(oFld.Path, "*.sql", Not this.blnUtcTime)
                ' Loop through files, adding to index
                For Each varKey In dFiles.Keys
                    strKey = strFolder & "\" & varKey
                    ' For most objects, the key value will be the modified date.
                    ' For objects (like views) that don't have a modified date, look up
                    ' the hash from the index. (Or fall back to the file date, if no
                    ' index entry is found.)
                    If dIndex.Exists(strKey) Then
                        strValue = dIndex(strKey)
                    Else
                        ' Use modified date from file
                        strValue = dFiles(varKey)
                    End If
                    ' Add each file with a key that matches the database object, and the
                    ' file modified date as the value for each item.
                    m_Files.Add strKey, strValue
                Next varKey
            End If
        Next oFld
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBaseFolders
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Return a dictionary of base folders used for component types
'---------------------------------------------------------------------------------------
'
Private Function GetBaseFolders() As Dictionary
    Set GetBaseFolders = New Dictionary
    With GetBaseFolders
        .CompareMode = TextCompare
        .Add "tables", Null
        .Add "views", Null
        .Add "procedures", Null
        .Add "functions", Null
        .Add "triggers", Null
        '.Add "indexes", Null
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_GetChangeCount
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Return count of modified objects
'---------------------------------------------------------------------------------------
'
Private Function IDbSchema_ObjectCount(blnModifiedOnly As Boolean) As Long
    If m_Files Is Nothing Then ScanFiles
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

    Dim colRules As Collection

    ' Map filter to collection
    If dInstance.Exists("Filter") Then
        ' Create collection of rules (we will skip comments and blank lines later)
        Set colRules = dInstance("Filter")
    Else
        ' No rules defined
        Set colRules = New Collection
    End If

    ' Set class values
    With this
        .strName = dNZ(dInstance, "Name")
        .strConnect = dNZ(dInstance, "Connect")
        .strBaseFolder = Options.GetExportFolder & "databases\" & GetSafeFileName(.strName) & PathSep
        Set .colFilters = colRules
        If dInstance.Exists("UtcDateTime") Then .blnUtcTime = dInstance("UtcDateTime")
        .blnInitialized = (Len(.strConnect))
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_BaseFolder
' Author    : Adam Waller
' Date      : 8/2/2023
' Purpose   : Return base folder for this schema export
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_BaseFolder() As String
    IDbSchema_BaseFolder = this.strBaseFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Name
' Author    : Adam Waller
' Date      : 8/2/2023
' Purpose   : Return the name of this schema
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_Name() As String
    IDbSchema_Name = this.strName
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_ServerType
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Return server type
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_ServerType() As eDatabaseServerType
    IDbSchema_ServerType = estMySql
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_TypeDescription
' Author    : Adam Waller
' Date      : 8/2/2023
' Purpose   : Return type description
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_TypeDescription() As Variant
    IDbSchema_TypeDescription = "MySQL Server"
End Property
