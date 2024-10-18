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


' How many times to retry connecting to ADODB.
Private Const ADODB_MAX_RETRIES As Long = 5

' Status of GetDDL stored procedure
Private Enum eSpStatus
    essUnknown
    essUnavailable
    essInstalled
End Enum

' Handle local variables
Private Type udtThis
    blnInitialized As Boolean
    strName As String
    strBaseFolder As String
    strConnect As String
    blnUtcTime As Boolean
    strUserID As String
    strPassword As String
    blnAttemptConvert As Boolean
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
Private Sub IDbSchema_Export(blnFullExport As Boolean _
                            , Optional strAlternatePath As String)

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

    ' Bail out if critical Error
    If Log.ErrorLevel >= eelError Then Exit Sub

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
            ExportObject dItem("type_desc"), dItem("schema"), dItem("name"), dItem("last_modified"), this.strBaseFolder & varItem, conn
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

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Export the object definition to a file
'---------------------------------------------------------------------------------------
'
Private Function ExportObject(strType _
                            , strSchema As String _
                            , strName As String _
                            , dteModified As Date _
                            , strPath As String _
                            , ByRef oConn As ADODB.Connection) As String

    Dim strSqlDef
    Dim strDefinition As String
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim strFullName As String
    Dim cmd As ADODB.Command
    Dim strCmdTemplate As String

    ' Attempt to use the sp_GetDDL SP if possible
    If CanUseGetDDL Then
        ' Prepare template statement for sp_GetDDL to work around VARCHAR(MAX) issue
        ' with many SQL Server ODBC drivers.
        strCmdTemplate = _
            "DECLARE @table TABLE (item text) " & _
            "INSERT INTO @table exec sp_GetDDL N'{name}' " & _
            "SELECT * FROM @table"
    Else
        ' Fall back to built-in SQL statements
        strCmdTemplate = _
            "DECLARE @table TABLE (item text) " & _
            "INSERT INTO @table SELECT object_definition (OBJECT_ID(N'{name}')) " & _
            "SELECT * FROM @table"

        'strCmdTemplate = "SELECT object_definition (OBJECT_ID(N'{name}'))"
        'strCmdTemplate = "test"
    End If

    ' Build full name of SQL object
    strFullName = "[" & strSchema & "].[" & strName & "]"

    ' Determine how to export this type of object
    Select Case strType
        Case "USER_TABLE", "VIEW", "SYNONYM", "SQL_STORED_PROCEDURE", _
            "SQL_SCALAR_FUNCTION", "SQL_INLINE_TABLE_VALUED_FUNCTION", "SQL_TABLE_VALUED_FUNCTION"
            strSqlDef = Replace(strCmdTemplate, "{name}", strFullName)

        Case "TYPE_TABLE", "SEQUENCE_OBJECT", "SERVICE_QUEUE", "SYSTEM_TABLE", "INTERNAL_TABLE"
            ' Unsupported non-dependent objects

    End Select

    ' Sanity check
    If Len(strSqlDef) Then
        If CanUseGetDDL Then
            Perf.OperationStart "Run sp_GetDDL on " & strType
        Else
            Perf.OperationStart "Get DDL for " & strType
        End If
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = oConn
            .CommandText = strSqlDef
            Set rst = .Execute
        End With

        ' Get secondary recordset with object definition record
        If strType = "USER_TABLE" And Not CanUseGetDDL Then
            strDefinition = GetTableDefFallback(strFullName, oConn)
        Else
            Set rst2 = rst.NextRecordset
            With rst2
                If Not .EOF Then strDefinition = Nz(.Fields(0))
                .Close
            End With
        End If

        ' Write object definition to file
        If strDefinition = vbNullString Then
            If FSO.FileExists(strPath) Then DeleteFile strPath
        Else
            ' Export to file, and set modified date to match SQL object
            WriteFile strDefinition, strPath
            SetFileDate strPath, dteModified, Not this.blnUtcTime
        End If

        Perf.OperationEnd
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanUseGetDDL
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Returns true if we can use the system SP `sp_GetDDL`
'---------------------------------------------------------------------------------------
'
Private Function CanUseGetDDL() As Boolean

    Static intUseSP As eSpStatus

    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset

    ' Cache whether or not sp_GetDDL exists
    If intUseSP = essUnknown Then

        ' Open connection, if not provided
        Set conn = GetNewOpenConnection
        If conn Is Nothing Then Exit Function

        Set cmd = New ADODB.Command
        With cmd
            ' Check in master database (System SP)
            .CommandText = "select OBJECT_ID('master.dbo.sp_GetDDL')"
            Set .ActiveConnection = conn
            Set rst = .Execute
            If Nz(rst.Fields(0).Value, 0) = 0 Then
                ' Nothing found on the master DB. Try this db.
                .CommandText = "select OBJECT_ID('sp_GetDDL')"
                Set rst = .Execute
                If Nz(rst.Fields(0).Value, 0) = 0 Then
                    ' Still not available
                    intUseSP = essUnavailable
                Else
                    ' Works on the local database
                    intUseSP = essInstalled
                End If
            Else
                ' Found an object ID. Should be available
                intUseSP = essInstalled
            End If
        End With

        ' Close connection
        conn.Close

        ' Add log entries if the tool is not available
        If intUseSP = essUnavailable Then
            Log.Add "   Note: sp_GetDDL was not available for generating object definitions. Using built-in SQL functions instead.", False
            Log.Add "   This system stored procedure can be found at: http://www.stormrage.com/2022/03/16/sp_getddla-and-sp_getddl/", False
        End If
    End If

    ' Return current status
    CanUseGetDDL = (intUseSP = essInstalled)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableDefFallback
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Return a simplified fall-back version of a table definition without using
'           : sp_GetDDL.
'---------------------------------------------------------------------------------------
'
Private Function GetTableDefFallback(strTable As String _
                                    , oConn As ADODB.Connection) As String

    Dim strSql As String
    Dim rst As ADODB.Recordset
    Dim intRst As Integer
    Dim fld As ADODB.Field
    Dim colText As New clsConcat

    ' Initialize counter
    intRst = 2

    ' Get initial table information
    strSql = "exec sp_help N'" & strTable & "'"
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set rst = oConn.Execute(strSql)
    colText.Add "-- sp_help Recordset 1" & vbCrLf & vbCrLf
    For Each fld In rst.Fields
        colText.Add fld.Name
        colText.Add vbTab
    Next fld
    colText.Add vbCrLf
    colText.Add rst.GetString(, , vbTab, vbCrLf)

    ' Loop through additional recordsets for columns, keys and other data
    Do
        Set rst = rst.NextRecordset
        If rst Is Nothing Then Exit Do
        If rst.State = adStateClosed Then Exit Do

        colText.Add vbCrLf & vbCrLf & "-- sp_help Recordset " & intRst & vbCrLf & vbCrLf
        For Each fld In rst.Fields
            colText.Add fld.Name
            colText.Add vbTab
        Next fld
        If Not rst.EOF Then
            colText.Add vbCrLf
            colText.Add rst.GetString(, , vbTab, vbCrLf)
        End If

        intRst = intRst + 1
    Loop

    ' Clear references
    Set fld = Nothing
    Set rst = Nothing

    ' Return SQL content
    GetTableDefFallback = colText.GetStr

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

    ' Clear module level objects
    Set m_AllItems = Nothing
    Set m_ModifiedItems = Nothing

    ' Make sure we initialize before running the scan
    If Not this.blnInitialized Then Exit Sub

    ' Initialize dictionaries
    Set m_AllItems = New Dictionary
    Set m_ModifiedItems = New Dictionary

    ' Open database connection
    Set conn = GetNewOpenConnection
    If conn Is Nothing Then Exit Sub

    ' Return list of objects from the server
    Perf.OperationStart "Retrieve SQL Objects"
    Set rstObjects = conn.Execute(CodeDb.QueryDefs("qryMsSqlServerObjects").SQL)
    Perf.OperationEnd

    ' Loop through objects, building dictionary of items that match our filter.
    Perf.OperationStart "Loop through SQL objects"
    With rstObjects
        Do While Not .EOF

            ' Build item path and full path to source file
            strSchema = Nz(!schema, "dbo") & "."
            strItem = Nz(!Folder) & PathSep & GetSafeFileName(strSchema & Nz(!Name)) & ".sql"
            strPath = this.strBaseFolder & strItem

            ' See if we pass the filter
            If PassesSchemaFilter(strItem, this.colFilters) Then

                ' Add all objects to full collection
                m_AllItems.Add strItem, Nz(!last_modified)

                ' Check for modification/new item
                If m_Files.Exists(strItem) Then
                    ' Flag as modified if the dates don't match
                    blnModified = (Nz(!last_modified) <> CStr(m_Files(strItem)))
                Else
                    ' File does not yet exist
                    blnModified = True
                End If

                ' Build dictionary of modified objects
                If blnModified Then
                    Set dItem = New Dictionary
                    dItem("type_desc") = Nz(!type_desc)
                    dItem("schema") = Nz(!schema)
                    dItem("name") = Nz(!Name)
                    dItem("last_modified") = CDate(!last_modified)
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
' Procedure : GetNewOpenConnection
' Author    : Adam Waller, hecon5
' Date      : 8/3/2023
' Purpose   : Return a new, open ADODB connection
'---------------------------------------------------------------------------------------
'
Private Function GetNewOpenConnection(Optional ByVal Reset As Boolean = False) As ADODB.Connection

    Static lngFailCount As Long

    Dim FunctionName As String
    Dim oConn As ADODB.Connection

    FunctionName = ModuleName(Me) & ".GetNewOpenConnection"

    LogUnhandledErrors FunctionName
    On Error Resume Next
    Perf.OperationStart FunctionName

Rebuild_Connection:
    Perf.OperationStart FunctionName & ".Create"

    Set oConn = New ADODB.Connection

Retry_Connection:
    With oConn
        If Len(this.strUserID) Then
            .Open GetADOConnectionString(this.strConnect, lngFailCount), this.strUserID, this.strPassword
        Else
            .Open GetADOConnectionString(this.strConnect, lngFailCount)
        End If

        CatchAny eelError, "Error Connecting to " & this.strName, FunctionName
        If (.State And adStateOpen) <> adStateOpen Then

            If lngFailCount <= ADODB_MAX_RETRIES Then
                lngFailCount = lngFailCount + 1
                GoTo Retry_Connection
            Else
                Log.Error eelError, "Could not establish connection to " & this.strName, FunctionName
                Perf.OperationEnd ' Close out the Create operation.
                Perf.OperationEnd ' Close out the Function operation.
                Exit Function
            End If
        End If
    End With

    Perf.OperationEnd

    If lngFailCount <= ADODB_MAX_RETRIES And (oConn.State And adStateOpen) <> adStateOpen Then GoTo Rebuild_Connection

    ' Check for any connection error
    If CatchAny(eelError, "Unable to connect to " & this.strName, FunctionName) Then
        Log.Add "Connection string: " & this.strConnect, False

    Else
        ' Return open connection
        Set GetNewOpenConnection = oConn
    End If

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetADOConnectionString
' Author    : hecon5
' Date      : 10/17/2024
' Purpose   : Returns with an ADODB connection string based on existing ODBC Connection.
'---------------------------------------------------------------------------------------
'
Private Function GetADOConnectionString(ByRef ODBCConnString As String _
                                        , Optional ByRef FailCount As Long = 0) As String

    Dim FunctionName As String
    Dim currConnString As String
    Dim ConnectionProperties() As String
    Dim PropertyCount As Long
    Dim strProperty As String
    Dim ConStrParsed As New clsConnectionODBC

    If Not this.blnAttemptConvert Then
        GetADOConnectionString = ODBCConnString
        Exit Function
    End If

    FunctionName = ModuleName(Me) & ".GetADOConnectionString"
    Perf.OperationStart FunctionName

    ConStrParsed.ParseOdbcConnectionString ODBCConnString

    currConnString = ConStrParsed.OriginalConnectionString
    ConnectionProperties = Split(currConnString, ";")

    With New clsConcat
        .AppendOnAdd = ";"
        ' Add some beginning properties that are specific to ADO (OLE) connection strings.
            Select Case FailCount
                Case 0 To 2
                    .Add "Provider=MSOLEDBSQL19"
                    .Add "Connect Retry Count=3"
                    If Len(ConStrParsed.Encrypt) Then .Add "Encrypt=yes"

                Case 3
                    .Add "Provider=MSOLEDBSQL"
                    If Len(ConStrParsed.Encrypt) Then .Add "Encrypt=yes"
                    .Add "Trusted_Connection=yes"
                    .Add "Connect Retry Count=3"

                Case 4
                    .Add "Provider=SQLNCLI11"
                    .Add "Trusted_Connection=yes"
                    If Len(ConStrParsed.Encrypt) Then .Add "Encrypt=yes"

                Case ADODB_MAX_RETRIES
                    .Add "Provider=SQLOLEDB"
                    .Add "Trusted_Connection=True"
                    If Len(ConStrParsed.Encrypt) Then .Add "Encrypt=yes"

                Case Else
                    .Add "Provider=SQLOLEDB"
                    If Len(ConStrParsed.Encrypt) Then .Add "Encrypt=true"
                    .Add "Trusted_Connection=True"
            End Select

        For PropertyCount = 0 To UBound(ConnectionProperties)
            strProperty = ConnectionProperties(PropertyCount)
            Select Case True ' This looks odd, yes, but it really just means we can use the comma delimited feature of "case"
                Case StartsWith(strProperty, "Provider=", vbTextCompare), _
                     StartsWith(strProperty, "ODBC", vbTextCompare), _
                     StartsWith(strProperty, "DataTypeCompatibility=", vbTextCompare), _
                     StartsWith(strProperty, "Driver=", vbTextCompare), _
                     StartsWith(strProperty, "Trusted_Connection=", vbTextCompare), _
                     StartsWith(strProperty, "Encrypt=", vbTextCompare), _
                     Len(strProperty) < 1
                    ' Don't include these; they aren't needed (or are specified elsewhere).

                Case Else
                    .Add strProperty
            End Select
        Next PropertyCount

        ' These properties should go at the end.
        If Len(ConStrParsed.Encrypt) Then .Add "Use Encryption for Data=true"
        .Add "MARS Connection=True"
        .Add "Integrated Security=SSPI"
        .Add "DataTypeCompatibility=80"

        Log.Add FunctionName & "ConnectionString: " & .GetStr, False
        GetADOConnectionString = .GetStr
    End With
    Perf.OperationEnd

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
    Dim strFolder As String

    ' Reset module-level dictionary
    Set m_Files = New Dictionary

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
                    ' Add each file with a key that matches the database object, and the
                    ' file modified date as the value for each item.
                    m_Files.Add strFolder & "\" & CStr(varKey), dFiles(varKey)
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
        .Add "views", Null
        .Add "tables", Null
        .Add "procedures", Null
        .Add "functions", Null
        .Add "types", Null
        .Add "sequences", Null
        .Add "synonymns", Null
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
        If dInstance.Exists("AttemptConvert") Then .blnAttemptConvert = dInstance("AttemptConvert")
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
    IDbSchema_ServerType = estMsSql
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_TypeDescription
' Author    : Adam Waller
' Date      : 8/2/2023
' Purpose   : Return type description
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_TypeDescription() As Variant
    IDbSchema_TypeDescription = "Microsoft SQL Server"
End Property
