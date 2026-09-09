Attribute VB_Name = "modTestQuery"
'---------------------------------------------------------------------------------------
' Module    : modTestQuery
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Legacy query import tests. Paired .sql overlays QueryDefs.SQL only when
'           : ForceImportOriginalQuerySQL is enabled (issue #769).
'           : Run: ?VCS.RunTests("modTestQuery")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("integration")


Private Const TEST_QUERY As String = "vcs_test_qry_sql_override"


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
    Dim lngErr As Long
    Dim strErr As String

    blnSavedForce = Options.ForceImportOriginalQuerySQL
    blnIndexDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True
    On Error GoTo CleanUp

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
    lngErr = Err.Number
    strErr = Err.Description
    Options.ForceImportOriginalQuerySQL = blnSavedForce
    VCSIndex.Disabled = blnIndexDisabled
    DeleteObjectIfExists acQuery, TEST_QUERY
    If Len(strFolder) > 0 Then
        If FSO.FolderExists(StripSlash(strFolder)) Then FSO.DeleteFolder StripSlash(strFolder), True
    End If
    If lngErr <> 0 Then Err.Raise lngErr, , strErr

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
