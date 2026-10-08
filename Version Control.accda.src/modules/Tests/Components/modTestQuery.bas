Attribute VB_Name = "modTestQuery"
'---------------------------------------------------------------------------------------
' Module    : modTestQuery
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Legacy query import tests. Paired .sql overlays QueryDefs.SQL only when
'           : ForceImportOriginalQuerySQL is enabled (issue #769). Also covers query
'           : properties that only the .json carries back in (issue #813).
'           : Run: ?VCS.RunTests("modTestQuery")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("integration")


Private Const TEST_QUERY As String = "vcs_test_qry_sql_override"
Private Const TEST_PROP_SRC As String = "vcs_test_qprop_src"
Private Const TEST_PROP_DST As String = "vcs_test_qprop_dst"
Private Const TEST_PROP_UPDATE As String = "vcs_test_qprop_update"
Private Const TEST_PROP_APPEND As String = "vcs_test_qprop_append"
Private Const TEST_PROP_UNION As String = "vcs_test_qprop_union"
Private Const TEST_PROP_PASSTHROUGH As String = "vcs_test_qprop_passthrough"


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyQueryForceSqlOverride
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : A divergent .bas + .sql pair imports the .bas definition unless Force
'           : original SQL is on, in which case the .sql definition wins. Also covers
'           : the Load Selected missing-source contract: format 4.1.2 resolves to .bas,
'           : so a leftover .qdef path is the one that used to fail silently.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyQueryForceSqlOverride()

    Dim cQuery As IDbComponent
    Dim strFolder As String
    Dim strBas As String
    Dim strSql As String
    Dim blnSavedForce As Boolean
    Dim blnIndexDisabled As Boolean
    Dim strImported As String
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    blnSavedForce = Options.ForceImportOriginalQuerySQL
    blnIndexDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True
    ' No operation begins in this project during a test run, so an error level left
    ' by an earlier test persists, and at eelCritical LoadComponentFromText reports
    ' failure even after the query loads.
    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError
    On Error GoTo ErrHandler

    DeleteObjectIfExists acQuery, TEST_QUERY
    strFolder = GetTempFolder("VCS") & PathSep
    strBas = strFolder & TEST_QUERY & ".bas"
    strSql = strFolder & TEST_QUERY & ".sql"

    WriteFile LegacyBasDefinition("FromBas"), strBas
    WriteFile "SELECT 2 AS FromSql;" & vbCrLf, strSql

    Set cQuery = New clsDbQuery

    Options.ForceImportOriginalQuerySQL = False
    DeleteObjectIfExists acQuery, TEST_QUERY
    cQuery.Import strBas
    strImported = CurrentDb.QueryDefs(TEST_QUERY).SQL
    TestAssert InStr(1, strImported, "FromBas", vbTextCompare) > 0, _
        "Force SQL off imports the .bas definition"
    TestAssert InStr(1, strImported, "FromSql", vbTextCompare) = 0, _
        "Force SQL off does not apply the paired .sql"

    Options.ForceImportOriginalQuerySQL = True
    DeleteObjectIfExists acQuery, TEST_QUERY
    Set cQuery = New clsDbQuery
    cQuery.Import strBas
    strImported = CurrentDb.QueryDefs(TEST_QUERY).SQL
    TestAssert InStr(1, strImported, "FromSql", vbTextCompare) > 0, _
        "Force SQL on overlays the paired .sql"
    TestAssert InStr(1, strImported, "FromBas", vbTextCompare) = 0, _
        "Force SQL on does not keep the .bas SQL"

    TestAssert SourceFileIsMissing(strFolder & TEST_QUERY & ".qdef"), _
        "missing .qdef is reported instead of treated as present"

CleanUp:
    On Error Resume Next
    Options.ForceImportOriginalQuerySQL = blnSavedForce
    VCSIndex.Disabled = blnIndexDisabled
    Operation.ErrorLevel = eelSavedLevel
    DeleteObjectIfExists acQuery, TEST_QUERY
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True
    End If
    If lngErr <> 0 Then TestAssert False, _
        "unexpected legacy query import error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLoadSingleObjectMissingSource
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : The Load Selected missing-file decision must be true for an empty path
'           : and a nonexistent file, and false when the file exists. LoadSingleObject
'           : itself logs eelError on this path, so the harness asserts the decision
'           : rather than invoking that side effect.
'---------------------------------------------------------------------------------------
'
Public Sub TestLoadSingleObjectMissingSource()

    Dim strFolder As String
    Dim strFile As String

    strFolder = GetTempFolder("VCS") & PathSep
    strFile = strFolder & "present.bas"

    TestAssert SourceFileIsMissing(vbNullString), "empty path is missing"
    TestAssert SourceFileIsMissing(strFolder & "no-such-object.bas"), _
        "nonexistent file is missing"

    WriteFile "x", strFile
    TestAssert Not SourceFileIsMissing(strFile), "existing file is not missing"

    If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True

End Sub


'---------------------------------------------------------------------------------------
' Function  : LegacyBasDefinition
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Minimal SQL-view SaveAsText body with a distinctive output alias.
'---------------------------------------------------------------------------------------
'
Private Function LegacyBasDefinition(strAlias As String) As String
    LegacyBasDefinition = _
        "dbMemo ""SQL"" =""SELECT 1 AS " & strAlias & ";""" & vbCrLf & _
        "dbBoolean ""ReturnsRecords"" =""-1""" & vbCrLf & _
        "dbInteger ""ODBCTimeout"" =""60""" & vbCrLf & _
        "Begin" & vbCrLf & _
        "End" & vbCrLf
End Function


'---------------------------------------------------------------------------------------
' Procedure : TestImportRestoresRecordLocksAndNoFormat
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/8/2026
' Purpose   : RecordLocks and NoFormat from the .json QueryProperties must survive
'           : an import, on SQL View and Design View queries alike. SaveAsText does
'           : not write RecordLocks, so the .qdef the composer generates is the only
'           : way back in. The Design View case also checks that the query keeps its
'           : layout: a .qdef that LoadFromText rejected would fall back to SQL View
'           : with only a warning and lose it. The pass-through case checks that the
'           : extra line does not make LoadFromText reject the query outright.
'---------------------------------------------------------------------------------------
'
Public Sub TestImportRestoresRecordLocksAndNoFormat()

    Dim strFolder As String
    Dim blnIndexDisabled As Boolean
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    blnIndexDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True
    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError
    On Error GoTo ErrHandler

    DropRecordLocksFixtures
    CurrentDb.Execute "CREATE TABLE " & TEST_PROP_SRC & " (Id LONG, ItemName TEXT(50))", dbFailOnError
    CurrentDb.Execute "CREATE TABLE " & TEST_PROP_DST & " (Id LONG, ItemName TEXT(50))", dbFailOnError
    strFolder = GetTempFolder("VCS") & PathSep

    ImportQueryFixture strFolder, TEST_PROP_UPDATE, _
        "UPDATE " & TEST_PROP_SRC & " SET ItemName = 'x' WHERE Id = 0;", _
        """QueryType"": 48, ""QueryProperties"": {" & _
        """RecordLocks"": {""Type"": 2, ""Value"": 2}, " & _
        """FailOnError"": {""Type"": 1, ""Value"": true}}"
    TestAssert QueryPropertyText(TEST_PROP_UPDATE, "RecordLocks") = "2", _
        "update query (SQL View) keeps RecordLocks = 2"
    TestAssert QueryDocumentPropertyText(TEST_PROP_UPDATE, "RecordLocks") = "2", _
        "update query keeps RecordLocks = 2 on its document in the Tables container"
    TestAssert QueryPropertyText(TEST_PROP_UPDATE, "FailOnError") = CStr(True), _
        "update query keeps FailOnError = True (got " & QueryPropertyText(TEST_PROP_UPDATE, "FailOnError") & ")"

    ImportQueryFixture strFolder, TEST_PROP_APPEND, _
        "INSERT INTO " & TEST_PROP_DST & " (ItemName)" & vbCrLf & _
        "SELECT " & TEST_PROP_SRC & ".ItemName" & vbCrLf & _
        "FROM " & TEST_PROP_SRC & ";", _
        """QueryType"": 64, ""QueryProperties"": {" & _
        """RecordLocks"": {""Type"": 2, ""Value"": 2}}, " & _
        """DesignLayout"": {""State"": 0, " & _
        """Window"": {""left"": 0, ""top"": 0, ""right"": 1000, ""bottom"": 700}, " & _
        """DesignerPane"": {""left"": -1, ""top"": -1, ""right"": 980, ""bottom"": 300}, " & _
        """GridLeft"": 0, ""GridTop"": 0, ""ColumnsShown"": 651, " & _
        """Tables"": [{""Name"": """ & TEST_PROP_SRC & """, ""Alias"": """", " & _
        """Left"": 48, ""Top"": 12, ""Right"": 192, ""Bottom"": 156, ""ScrollTop"": 0}]}"
    TestAssert QueryPropertyText(TEST_PROP_APPEND, "RecordLocks") = "2", _
        "append query (Design View) keeps RecordLocks = 2"
    TestAssert QueryHasDesignLayout(TEST_PROP_APPEND), _
        "append query is still imported in Design View"

    ImportQueryFixture strFolder, TEST_PROP_UNION, _
        "SELECT Id, ItemName FROM " & TEST_PROP_SRC & vbCrLf & _
        "UNION SELECT Id, ItemName FROM " & TEST_PROP_DST & ";", _
        """QueryType"": 128, ""QueryProperties"": {" & _
        """NoFormat"": {""Type"": 1, ""Value"": false}}"
    TestAssert QueryPropertyText(TEST_PROP_UNION, "NoFormat") = CStr(False), _
        "union query keeps NoFormat = False (got " & QueryPropertyText(TEST_PROP_UNION, "NoFormat") & ")"
    TestAssert QueryDocumentPropertyText(TEST_PROP_UNION, "NoFormat") = CStr(False), _
        "union query keeps NoFormat = False on its document (got " & _
        QueryDocumentPropertyText(TEST_PROP_UNION, "NoFormat") & ")"

    ImportQueryFixture strFolder, TEST_PROP_PASSTHROUGH, _
        "SELECT 1 AS One;", _
        """QueryType"": 112, ""Connect"": ""ODBC;"", ""QueryProperties"": {" & _
        """RecordLocks"": {""Type"": 2, ""Value"": 2}}"
    TestAssert ObjectExists(acQuery, TEST_PROP_PASSTHROUGH), _
        "pass-through query with RecordLocks is imported"
    TestAssert QueryPropertyText(TEST_PROP_PASSTHROUGH, "RecordLocks") = "2", _
        "pass-through query keeps RecordLocks = 2"

CleanUp:
    On Error Resume Next
    VCSIndex.Disabled = blnIndexDisabled
    Operation.ErrorLevel = eelSavedLevel
    DropRecordLocksFixtures
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True
    End If
    If lngErr <> 0 Then TestAssert False, _
        "unexpected query property import error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportQueryFixture
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/8/2026
' Purpose   : Write a .sql and its .json (strItems is the body of "Items") and import
'           : the pair through clsDbQuery, as a build does.
'---------------------------------------------------------------------------------------
'
Private Sub ImportQueryFixture(strFolder As String, strName As String, _
    strSql As String, strItems As String)

    Dim cQuery As IDbComponent
    Dim strSqlFile As String

    strSqlFile = strFolder & strName & ".sql"
    WriteFile strSql & vbCrLf, strSqlFile
    WriteFile "{""Info"": {""Class"": ""clsDbQuery"", ""Description"": """ & strName & """}, " & _
        """Items"": {" & strItems & "}}", strFolder & strName & ".json"
    DeleteObjectIfExists acQuery, strName
    Set cQuery = New clsDbQuery
    cQuery.Import strSqlFile

End Sub


Private Function QueryPropertyText(strQuery As String, strProperty As String) As String
    On Error Resume Next
    QueryPropertyText = "(missing)"
    Dim dbs As DAO.Database
    Dim prp As DAO.Property
    Set dbs = CurrentDb
    Set prp = dbs.QueryDefs(strQuery).Properties(strProperty)
    ' Callers compare against CStr(True) or CStr(False): CStr of a Boolean is
    ' localized (a Spanish Access returns "Verdadero"), so never a literal "True".
    If prp.Type = dbBoolean Then QueryPropertyText = CStr(CBool(prp.Value)) Else QueryPropertyText = CStr(prp.Value)
End Function


Private Function QueryDocumentPropertyText(strQuery As String, strProperty As String) As String
    ' The same read on the query's document in the Tables container, which keeps
    ' its own copy of these properties.
    On Error Resume Next
    QueryDocumentPropertyText = "(missing)"
    Dim dbs As DAO.Database
    Dim prp As DAO.Property
    Set dbs = CurrentDb
    Set prp = dbs.Containers("Tables").Documents(strQuery).Properties(strProperty)
    If prp.Type = dbBoolean Then
        QueryDocumentPropertyText = CStr(CBool(prp.Value))
    Else
        QueryDocumentPropertyText = CStr(prp.Value)
    End If
End Function


Private Function QueryHasDesignLayout(strQuery As String) As Boolean
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset("SELECT LvExtra FROM MSysObjects WHERE Name = '" & _
        strQuery & "' AND Type = 5", dbOpenSnapshot)
    If Not rst.EOF Then QueryHasDesignLayout = Not IsNull(rst!LvExtra)
    rst.Close
End Function


Private Sub DropRecordLocksFixtures()
    On Error Resume Next
    DeleteObjectIfExists acQuery, TEST_PROP_UPDATE
    DeleteObjectIfExists acQuery, TEST_PROP_APPEND
    DeleteObjectIfExists acQuery, TEST_PROP_UNION
    DeleteObjectIfExists acQuery, TEST_PROP_PASSTHROUGH
    DeleteObjectIfExists acTable, TEST_PROP_SRC
    DeleteObjectIfExists acTable, TEST_PROP_DST
End Sub
