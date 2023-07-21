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
            ExportObject dItem("type_desc"), dItem("schema"), dItem("name"), this.strBaseFolder & varItem, conn
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

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Export the object definition and update the index
'---------------------------------------------------------------------------------------
'
Private Sub ExportObject(strType, strSchema As String, strName As String, strPath As String, ByRef oConn As ADODB.Connection)

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

End Sub


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
    dItem("ExportDate") = Now()
    dItem("LastModified") = dObject("LastModified")
    Perf.OperationEnd

End Sub


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

            ' Move to next object
            .MoveNext
        Loop
        .Close
    End With
    Perf.OperationEnd

    ' Close connection
    conn.Close

End Sub


Private Function PurgeOrphanedObjects()

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_GetChangeCount
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Return count of modified objects
'---------------------------------------------------------------------------------------
'
Private Function IDbSchema_GetChangeCount() As Long
    If m_AllItems Is Nothing Then ScanDatabaseObjects
    IDbSchema_GetChangeCount = m_ModifiedItems.Count
End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Initialize
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Initialize the database schema
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Initialize(dInstance As Scripting.IDictionary)
    With this
        .strName = dNZ(dInstance, "Name")
        .strConnect = dNZ(dInstance, "Connect")
        .strBaseFolder = Options.GetExportFolder & "databases\" & GetSafeFileName(.strName) & PathSep
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
