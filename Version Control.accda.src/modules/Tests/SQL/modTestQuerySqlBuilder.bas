Attribute VB_Name = "modTestQuerySqlBuilder"
'---------------------------------------------------------------------------------------
' Module    : modTestQuerySqlBuilder
' Author    : Adam Waller
' Date      : 4/29/2026
' Purpose   : Diagnostic harness that validates the fast MSysQueries SQL builder
'           : against Access's own QueryDefs.SQL text for the current database.
'           :
'           : This is intentionally an on-demand Advanced Tools workflow. Normal
'           : exports keep using the fast builder; this module pays the per-query
'           : QueryDefs.SQL cost only when a developer/user asks for validation.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.SQL")
'@Tag("integration")

Private Const ModuleName As String = "modTestQuerySqlBuilder"
Private Const STATUS_PASS_EXACT As String = "PASS_EXACT"
Private Const STATUS_PASS_CANONICAL As String = "PASS_CANONICAL"
Private Const STATUS_REVIEW As String = "REVIEW"
Private Const STATUS_FAIL_WARNING As String = "FAIL_WARNING"
Private Const STATUS_ERROR As String = "ERROR"
Private Const DIFF_CONTEXT_LINES As Long = 2
Private Const DIFF_MAX_CHARS_FOR_JSON As Long = 4000


'---------------------------------------------------------------------------------------
' Procedure : RunQuerySqlBuilderValidation
' Author    : Adam Waller
' Date      : 4/29/2026
' Purpose   : Compare reconstructed SQL to Access QueryDefs.SQL for selected
'           : queries or all queries in the current database. Returns a JSON
'           : summary and writes a durable log plus per-query artifacts for
'           : non-exact/non-safe results.
'---------------------------------------------------------------------------------------
'
Public Function RunQuerySqlBuilderValidation(Optional ByVal varQueryNames As Variant) As String

    Dim dResult As Dictionary
    Dim dStats As Dictionary
    Dim colResults As Collection
    Dim colQueryNames As Collection
    Dim strExportRoot As String
    Dim strLogFolder As String
    Dim strArtifactRoot As String
    Dim strLogPath As String
    Dim blnRootOwned As Boolean
    Dim cRoot As clsRootOperationLease
    Dim sngStart As Single
    Dim lngCurrent As Long
    Dim lngTotal As Long
    Dim varName As Variant
    Dim dQueryResult As Dictionary
    Dim blnCanceled As Boolean
    Dim strQueryFilterDisplay As String
    Dim lngPriorFormat As Long
    Dim blnFormatPinned As Boolean
    Dim dCanonStats As Dictionary
    Dim cLedger As clsConcat
    Dim strLedgerPath As String

    LogUnhandledErrors ModuleName & ".RunQuerySqlBuilderValidation"
    On Error GoTo ErrHandler

    Set dResult = New Dictionary
    Set dStats = NewValidationStats()
    Set colResults = New Collection

    If Not DatabaseFileOpen Then
        dResult.Add "success", False
        dResult.Add "error", "No database is open."
        RunQuerySqlBuilderValidation = ConvertToJson(dResult)
        Exit Function
    End If

    Set colQueryNames = GetValidationQueryNames(varQueryNames)
    strQueryFilterDisplay = QueryFilterDisplay(colQueryNames)

    ' This harness is called both on its own and from inside a test run. Nested, it owns
    ' nothing: the enclosing run holds the root, the console, and the log file.
    If Not (Operation.Status = eosRunning And Operation.OperationType = eotTestRun) Then
        If Operation.AutomationSource Then Operation.ForceUnattended = True
        Set cRoot = Operation.TryBeginRoot(eotOther)
        If cRoot Is Nothing Then
            dResult.Add "success", False
            dResult.Add "error", "Could not begin validation (another operation may be running)."
            RunQuerySqlBuilderValidation = ConvertToJson(dResult)
            Exit Function
        End If
        blnRootOwned = True
        If Operation.AutomationSource Or Operation.ForceUnattended Then
            Operation.InteractionMode = eimSilent
        End If
    End If

    If blnRootOwned Then PrepareValidationConsole

    Set Options = Nothing
    Options.LoadProjectOptions
    If Operation.Source = eosMCPTool Or Operation.Source = eosExternalAPI Then
        Options.LoadOptionOverrides
    End If

    ' Validate the builder as the current release writes it, whatever format this
    ' project pins; the project's setting is restored afterwards.
    lngPriorFormat = Options.ExportFormatVersion
    blnFormatPinned = True
    Options.ExportFormatVersion = LatestExportFormat

    Set dCanonStats = NewCanonicalizerStats()
    Set cLedger = New clsConcat
    cLedger.Add "# Canonicalizer report", vbCrLf, vbCrLf
    cLedger.Add "Builder SQL before and after clsSqlCanonicalizer. Report only: export output is unchanged.", vbCrLf, vbCrLf
    cLedger.Add "| Query | Result | Idempotent | Parse | Structure | Behavior |", vbCrLf
    cLedger.Add "|---|---|---|---|---|---|", vbCrLf

    strExportRoot = Options.GetExportFolder
    If Len(strExportRoot) = 0 Then strExportRoot = CurrentProject.Path & PathSep
    strLogFolder = FSO.BuildPath(strExportRoot, "logs")
    strArtifactRoot = FSO.BuildPath(strLogFolder, "SqlBuilderValidation_" & Log.OperationId) & PathSep
    strLogPath = FSO.BuildPath(strLogFolder, "SqlBuilderValidation_" & Log.OperationId & ".log")
    strLedgerPath = strArtifactRoot & "canonicalizer.md"
    VerifyPath strArtifactRoot

    ' Only the owner of the root starts a fresh log; a nested run appends to the
    ' enclosing test run's console and log file.
    If blnRootOwned Then
        Log.Clear
        Log.KeepProgressVisible = True
        Log.SourcePath = strExportRoot
        Log.Active = True
        Perf.StartTiming
    End If
    sngStart = Perf.MicroTimer

    With Log
        .Spacer
        .Add T("Query SQL Builder Validation")
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Database: {0}", var0:=CurrentProject.FullName), False
        .Add T("Artifact folder: {0}", var0:=strArtifactRoot), False
        .Add T("Export format: {0}", var0:=ExportFormatToVersion(LatestExportFormat))
        If Len(strQueryFilterDisplay) > 0 Then .Add T("Query filter: {0}", var0:=strQueryFilterDisplay)
        .Add T("Started: {0}", var0:=Format$(Now, "yyyy-mm-dd hh:nn:ss"))
        .Spacer
        .Flush
    End With

    lngTotal = colQueryNames.Count
    If lngTotal = 0 Then
        Log.Add T("No queries found to validate.")
        GoTo CleanUp
    End If
    Log.Progress 0, lngTotal, T("Starting validation")

    Log.Add T("Validating {0} quer{1}...", var0:=CStr(lngTotal), _
        var1:=IIf(lngTotal = 1, "y", "ies"))

    For Each varName In colQueryNames
        lngCurrent = lngCurrent + 1
        Log.Progress lngCurrent, lngTotal, CStr(varName)

        ' Print the query name padded to fixed width, leaving the line open.
        ' This matches the export routine's "PadRight, work, Add result" pattern.
        Log.PadRight ConsoleQueryName(CStr(varName)), True, False

        ' Process the query
        Set dQueryResult = ValidateOneQuery(CStr(varName), strArtifactRoot)
        colResults.Add dQueryResult
        IncrementStatus dStats, CStr(dQueryResult("status"))
        If dQueryResult.Exists("canonicalizer") Then
            TallyCanonicalizer dCanonStats, dQueryResult("canonicalizer")
            cLedger.Add CanonicalizerLedgerLine(CStr(varName), dQueryResult("canonicalizer"))
        End If

        ' Now print the result, completing the line
        Log.Add ResultLabel(CStr(dQueryResult("status"))), , , _
            ResultColor(CStr(dQueryResult("status")))
        LogQueryDetailFileOnly dQueryResult
        Log.Flush
        Operation.Pulse
        DoEvents
        If Operation.Status <> eosRunning Then
            blnCanceled = True
            Exit For
        End If
    Next varName

CleanUp:
    dStats("total") = lngTotal
    dStats("elapsedSeconds") = Round(Perf.MicroTimer - sngStart, 3)

    With Log
        .Spacer
        .Add vbNullString
        .Add T("SQL builder validation summary"), True, , "blue", True
        LogSummaryLine "  Total queries:      ", CStr(dStats("total")), vbNullString
        LogSummaryLine "  Exact pass:         ", CStr(dStats(STATUS_PASS_EXACT)), _
            IIf(dStats(STATUS_PASS_EXACT) > 0, "green", vbNullString)
        LogSummaryLine "  Canonical pass:     ", CStr(dStats(STATUS_PASS_CANONICAL)), _
            IIf(dStats(STATUS_PASS_CANONICAL) > 0, "green", vbNullString)
        LogSummaryLine "  Review:             ", CStr(dStats(STATUS_REVIEW)), _
            IIf(dStats(STATUS_REVIEW) > 0, "orange", vbNullString)
        LogSummaryLine "  Warning failures:   ", CStr(dStats(STATUS_FAIL_WARNING)), _
            IIf(dStats(STATUS_FAIL_WARNING) > 0, "red", vbNullString)
        LogSummaryLine "  Errors:             ", CStr(dStats(STATUS_ERROR)), _
            IIf(dStats(STATUS_ERROR) > 0, "red", vbNullString)
        .Add T("  Elapsed (s):        {0}", var0:=CStr(dStats("elapsedSeconds")))
        .Add T("  Log file:           {0}", var0:=strLogPath)
        .Add T("  Artifact folder:    {0}", var0:=strArtifactRoot)
        If blnCanceled Then .Add T("  Result:             CANCELED"), , , "red", True
        .Add vbNullString
        .Add T("Canonicalizer report (export output unchanged)"), True, , "blue", True
        LogSummaryLine "  Checked:            ", CStr(dCanonStats("total")), vbNullString
        LogSummaryLine "  Changed:            ", CStr(dCanonStats("changed")), vbNullString
        LogSummaryLine "  Declined:           ", CStr(dCanonStats("declined")), _
            IIf(dCanonStats("declined") > 0, "orange", vbNullString)
        LogSummaryLine "  Failures:           ", CStr(dCanonStats("failures")), _
            IIf(dCanonStats("failures") > 0, "red", vbNullString)
        .Add T("  Ledger:             {0}", var0:=strLedgerPath)
        .Spacer
    End With
    WriteFile cLedger.GetStr, strLedgerPath

    Options.ExportFormatVersion = lngPriorFormat
    blnFormatPinned = False

    If blnRootOwned Then
        Perf.EndTiming
        On Error Resume Next
        Log.SaveFile strLogPath
        Log.Active = False
        Log.Flush
        On Error GoTo 0
    Else
        Log.Flush
    End If

    dResult.Add "success", (Not blnCanceled And ValidationSucceeded(dStats) _
        And CLng(dCanonStats("failures")) = 0)
    dResult.Add "canceled", blnCanceled
    dResult.Add "exportFormat", ExportFormatToVersion(LatestExportFormat)
    dResult.Add "database", CurrentProject.FullName
    dResult.Add "queryFilter", strQueryFilterDisplay
    dResult.Add "queryFilters", QueryFilterCollectionForJson(colQueryNames)
    dResult.Add "logPath", strLogPath
    dResult.Add "artifactRoot", strArtifactRoot
    dResult.Add "stats", dStats
    dResult.Add "canonicalizer", dCanonStats
    dResult.Add "canonicalizerLedger", strLedgerPath
    dResult.Add "results", colResults

    If blnRootOwned Then UpdateValidationConsoleComplete strLogPath, strArtifactRoot, blnCanceled, dResult

    If blnRootOwned Then
        If blnCanceled Then
            cRoot.Complete eorCanceled
        ElseIf dResult("success") Then
            cRoot.Complete eorSuccess
        Else
            cRoot.Complete eorFailed
        End If
        Log.KeepProgressVisible = False
    End If
    RunQuerySqlBuilderValidation = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    On Error Resume Next
    If blnFormatPinned Then Options.ExportFormatVersion = lngPriorFormat
    If blnRootOwned Then
        cRoot.Complete eorFailed
        Log.KeepProgressVisible = False
        Log.Active = False
    End If
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    RunQuerySqlBuilderValidation = ConvertToJson(dResult)
End Function


'---------------------------------------------------------------------------------------
' Procedure : InspectQuerySqlStorage
' Author    : Adam Waller
' Date      : 4/30/2026
' Purpose   : Create a sandbox query from SQL and return Access-authored
'           : QueryDefs.SQL plus MSysObjects/MSysQueries storage rows. This is
'           : a diagnostic helper for building the join reconstruction spec.
'---------------------------------------------------------------------------------------
'
Public Function InspectQuerySqlStorage(ByVal strSql As String, _
    Optional ByVal blnForceDesignView As Boolean = False) As String

    Dim dResult As Dictionary
    Dim dMetadata As Dictionary
    Dim colRows As Collection
    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strQueryName As String
    Dim blnCreated As Boolean

    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "forceDesignView", blnForceDesignView

    LogUnhandledErrors ModuleName & ".InspectQuerySqlStorage"
    On Error GoTo ErrHandler

    If Not DatabaseFileOpen Then
        dResult.Add "error", "No database is open."
        GoTo CleanUp
    End If

    strQueryName = UniqueInspectQueryName()
    dResult.Add "queryName", strQueryName
    dResult.Add "inputSql", strSql

    Set dbs = CurrentDb
    Set qdf = dbs.CreateQueryDef(strQueryName, strSql)
    blnCreated = True

    If blnForceDesignView Then
        On Error Resume Next
        DoCmd.OpenQuery strQueryName, acViewDesign
        DoCmd.Close acQuery, strQueryName, acSaveYes
        If Err.Number <> 0 Then
            dResult.Add "designViewError", CStr(Err.Number) & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If

    dResult.Add "accessSql", dbs.QueryDefs(strQueryName).SQL
    Set dMetadata = New Dictionary
    Set colRows = ReadMsysQueryRows(strQueryName, dMetadata)
    dResult.Add "metadata", dMetadata
    dResult.Add "msysqueries", colRows
    dResult("success") = True

CleanUp:
    On Error Resume Next
    If blnCreated Then dbs.QueryDefs.Delete strQueryName
    dResult.Add "cleanupError", IIf(Err.Number = 0, vbNullString, CStr(Err.Number) & " " & Err.Description)
    On Error GoTo 0
    InspectQuerySqlStorage = ConvertToJson(dResult, JSON_WHITESPACE)
    Exit Function

ErrHandler:
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    Resume CleanUp
End Function


'---------------------------------------------------------------------------------------
' Procedure : CompareSqlBuilderArtifacts
' Author    : Adam Waller
' Date      : 4/30/2026
' Purpose   : Compare SQL builder validation artifacts with DAO row counts and
'           : field schemas. This runs inside Access so VBA UDFs and linked
'           : tables work better than direct ODBC validation.
'---------------------------------------------------------------------------------------
'
Public Function CompareSqlBuilderArtifacts(ByVal strArtifactRoot As String) As String

    Dim dResult As Dictionary
    Dim dStats As Dictionary
    Dim colResults As Collection
    Dim strLedgerPath As String
    Dim strLedger As String
    Dim fldRoot As Object
    Dim fldArtifact As Object
    Dim dArtifactResult As Dictionary

    Set dResult = New Dictionary
    Set dStats = NewArtifactCompareStats()
    Set colResults = New Collection

    LogUnhandledErrors ModuleName & ".CompareSqlBuilderArtifacts"
    On Error GoTo ErrHandler

    If Right$(strArtifactRoot, 1) <> PathSep Then strArtifactRoot = strArtifactRoot & PathSep
    If Not FSO.FolderExists(strArtifactRoot) Then
        dResult.Add "success", False
        dResult.Add "error", "Artifact folder not found: " & strArtifactRoot
        CompareSqlBuilderArtifacts = ConvertToJson(dResult, JSON_WHITESPACE)
        Exit Function
    End If

    strLedger = "# DAO SQL Builder Artifact Ledger" & vbCrLf & vbCrLf & _
        "Artifact root: `" & strArtifactRoot & "`" & vbCrLf & vbCrLf & _
        "| Query | Status | Access Count | Generated Count | Detail |" & vbCrLf & _
        "|---|---|---:|---:|---|" & vbCrLf

    Set fldRoot = FSO.GetFolder(strArtifactRoot)
    For Each fldArtifact In fldRoot.SubFolders
        If FSO.FileExists(fldArtifact.Path & PathSep & "access.sql") And _
           FSO.FileExists(fldArtifact.Path & PathSep & "generated.sql") Then
            Set dArtifactResult = CompareOneSqlBuilderArtifact(CStr(fldArtifact.Path))
            colResults.Add dArtifactResult
            IncrementArtifactCompareStats dStats, CStr(dArtifactResult("status"))
            strLedger = strLedger & ArtifactCompareLedgerLine(dArtifactResult)
        End If
    Next fldArtifact

    strLedger = strLedger & vbCrLf & "## Summary" & vbCrLf & _
        "- Total: " & CStr(dStats("total")) & vbCrLf & _
        "- SAFE_DAO_COMPARE: " & CStr(dStats("SAFE_DAO_COMPARE")) & vbCrLf & _
        "- MISMATCH: " & CStr(dStats("MISMATCH")) & vbCrLf & _
        "- BLOCKED: " & CStr(dStats("BLOCKED")) & vbCrLf & _
        "- ACTION_OR_UNSUPPORTED: " & CStr(dStats("ACTION_OR_UNSUPPORTED")) & vbCrLf

    strLedgerPath = strArtifactRoot & "dao_compare_ledger.md"
    WriteFile strLedger, strLedgerPath

    dResult.Add "success", (CLng(dStats("MISMATCH")) = 0)
    dResult.Add "artifactRoot", strArtifactRoot
    dResult.Add "ledgerPath", strLedgerPath
    dResult.Add "stats", dStats
    dResult.Add "results", colResults
    CompareSqlBuilderArtifacts = ConvertToJson(dResult, JSON_WHITESPACE)
    Exit Function

ErrHandler:
    dResult.RemoveAll
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    CompareSqlBuilderArtifacts = ConvertToJson(dResult, JSON_WHITESPACE)
End Function


Private Sub UpdateValidationConsoleComplete(ByVal strLogPath As String, ByVal strArtifactRoot As String, _
    ByVal blnCanceled As Boolean, ByVal dResult As Dictionary)

    Dim blnSuccess As Boolean

    blnSuccess = False
    If Not dResult Is Nothing Then
        If dResult.Exists("success") Then blnSuccess = dResult("success")
    End If

    On Error Resume Next
    With Form_frmVCSMain
        .txtLog.ScrollBars = 2
        .strLastLogFilePath = strLogPath
        .cmdOpenLogFile.Visible = (Len(strLogPath) > 0)
        .cmdCopyPath.Visible = .cmdOpenLogFile.Visible
        .strCopyToClipboardPath = strLogPath
        If blnCanceled Then
            .SetStatusText T("Canceled"), T("Validation Canceled"), _
                T("The SQL builder validation was canceled. Partial results are available in the log file.")
        ElseIf blnSuccess Then
            .SetStatusText T("Finished"), T("SQL Builder Validation Passed"), _
                T("All validated queries passed. Click the log button to view the full results.")
        Else
            .SetStatusText T("Finished"), T("SQL Builder Validation Needs Review"), _
                T("Some queries need review. Click the log button to view details and diagnostic artifact paths.")
        End If
    End With
    Err.Clear
End Sub


Private Sub PrepareValidationConsole()
    On Error Resume Next
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    Dim frm As Form_frmVCSMain
    Set frm = Form_frmVCSMain
    With frm
        .ResetForOperation
        .SetStatusText T("Validating query SQL builder..."), _
            T("Comparing reconstructed SQL to Access SQL"), _
            T("A summary is shown here; review and warning cases are written to the validation log folder.")
    End With
    Err.Clear
End Sub


Private Function GetValidationQueryNames(Optional ByVal varQueryNames As Variant) As Collection
    Dim colNames As New Collection
    Dim qry As AccessObject
    Dim varName As Variant

    If IsMissing(varQueryNames) Or IsEmpty(varQueryNames) Then
        For Each qry In CurrentData.AllQueries
            colNames.Add qry.Name
        Next qry
    ElseIf IsArray(varQueryNames) Then
        For Each varName In varQueryNames
            AddValidationQueryName colNames, varName
        Next varName
        If colNames.Count = 0 Then
            For Each qry In CurrentData.AllQueries
                colNames.Add qry.Name
            Next qry
        End If
    Else
        AddValidationQueryName colNames, varQueryNames
        If colNames.Count = 0 Then
            For Each qry In CurrentData.AllQueries
                colNames.Add qry.Name
            Next qry
        End If
    End If

    Set GetValidationQueryNames = colNames
End Function


Private Sub AddValidationQueryName(ByVal colNames As Collection, ByVal varName As Variant)
    If IsObject(varName) Then Exit Sub
    If IsNull(varName) Then Exit Sub

    Dim strName As String
    strName = Trim$(CStr(varName))
    If Len(strName) = 0 Then Exit Sub

    colNames.Add strName
End Sub


Private Function QueryFilterDisplay(ByVal colQueryNames As Collection) As String
    If colQueryNames Is Nothing Then Exit Function
    Select Case colQueryNames.Count
        Case 0
        Case CurrentData.AllQueries.Count
            QueryFilterDisplay = vbNullString
        Case 1
            QueryFilterDisplay = CStr(colQueryNames(1))
        Case Else
            QueryFilterDisplay = CStr(colQueryNames.Count) & " selected queries"
    End Select
End Function


Private Function QueryFilterCollectionForJson(ByVal colQueryNames As Collection) As Collection
    Dim colFilters As New Collection
    Dim varName As Variant

    If Not colQueryNames Is Nothing Then
        If colQueryNames.Count <> CurrentData.AllQueries.Count Then
            For Each varName In colQueryNames
                colFilters.Add CStr(varName)
            Next varName
        End If
    End If

    Set QueryFilterCollectionForJson = colFilters
End Function


Private Function CompareOneSqlBuilderArtifact(ByVal strArtifactFolder As String) As Dictionary
    Dim dResult As Dictionary
    Dim dbs As DAO.Database
    Dim strName As String
    Dim strAccessSql As String
    Dim strGeneratedSql As String
    Dim strAccessSelect As String
    Dim strGeneratedSelect As String
    Dim strAccessQueryName As String
    Dim strGeneratedQueryName As String
    Dim lngAccessCount As Long
    Dim lngGeneratedCount As Long
    Dim strAccessSchema As String
    Dim strGeneratedSchema As String
    Dim strDetail As String

    Set dResult = New Dictionary
    Set dbs = CurrentDb
    strName = FSO.GetFileName(strArtifactFolder)
    strAccessQueryName = "vcs_cmp_access_" & UniqueHashSuffix(strArtifactFolder)
    strGeneratedQueryName = "vcs_cmp_generated_" & UniqueHashSuffix(strArtifactFolder)

    dResult.Add "name", strName
    dResult.Add "artifactFolder", strArtifactFolder

    LogUnhandledErrors ModuleName & ".CompareOneSqlBuilderArtifact"
    On Error GoTo ErrHandler

    strAccessSql = ReadFile(strArtifactFolder & PathSep & "access.sql")
    strGeneratedSql = ReadFile(strArtifactFolder & PathSep & "generated.sql")
    strAccessSelect = ExtractSourceSelect(strAccessSql)
    strGeneratedSelect = ExtractSourceSelect(strGeneratedSql)

    If Len(strAccessSelect) = 0 Or Len(strGeneratedSelect) = 0 Then
        dResult.Add "status", "ACTION_OR_UNSUPPORTED"
        dResult.Add "detail", "No read-only SELECT source extracted."
        GoTo CleanUp
    End If

    DeleteQueryDefIfExists dbs, strAccessQueryName
    DeleteQueryDefIfExists dbs, strGeneratedQueryName

    dbs.CreateQueryDef strAccessQueryName, strAccessSelect
    dbs.CreateQueryDef strGeneratedQueryName, strGeneratedSelect

    lngAccessCount = QueryDefRowCount(dbs, strAccessQueryName)
    lngGeneratedCount = QueryDefRowCount(dbs, strGeneratedQueryName)
    strAccessSchema = QueryDefSchemaSignature(dbs, strAccessQueryName)
    strGeneratedSchema = QueryDefSchemaSignature(dbs, strGeneratedQueryName)

    dResult.Add "accessCount", lngAccessCount
    dResult.Add "generatedCount", lngGeneratedCount
    dResult.Add "schemaMatch", (strAccessSchema = strGeneratedSchema)

    If lngAccessCount = lngGeneratedCount And strAccessSchema = strGeneratedSchema Then
        dResult.Add "status", "SAFE_DAO_COMPARE"
        dResult.Add "detail", "row count and schema match"
    Else
        strDetail = "accessSchema=" & strAccessSchema & " generatedSchema=" & strGeneratedSchema
        dResult.Add "status", "MISMATCH"
        dResult.Add "detail", strDetail
    End If

CleanUp:
    On Error Resume Next
    DeleteQueryDefIfExists dbs, strAccessQueryName
    DeleteQueryDefIfExists dbs, strGeneratedQueryName
    Err.Clear
    On Error GoTo 0
    Set CompareOneSqlBuilderArtifact = dResult
    Exit Function

ErrHandler:
    strDetail = CStr(Err.Number) & " " & Err.Description
    Err.Clear
    On Error GoTo -1
    dResult.Add "status", "BLOCKED"
    dResult.Add "detail", strDetail
    Resume CleanUp
End Function


Private Function ExtractSourceSelect(ByVal strSql As String) As String
    strSql = Replace(strSql, ChrW(&HFEFF), vbNullString)
    strSql = Trim$(strSql)
    Do While Right$(strSql, 1) = ";"
        strSql = Trim$(Left$(strSql, Len(strSql) - 1))
    Loop

    If UCase$(Left$(strSql, 11)) = "INSERT INTO" Then
        Dim lngSelect As Long
        lngSelect = InStr(1, strSql, "SELECT ", vbTextCompare)
        If lngSelect > 0 Then strSql = Mid$(strSql, lngSelect)
    End If

    If UCase$(Left$(Trim$(strSql), 6)) = "SELECT" Then ExtractSourceSelect = strSql
End Function


Private Sub DeleteQueryDefIfExists(ByVal dbs As DAO.Database, ByVal strQueryName As String)
    Dim qdf As DAO.QueryDef
    Dim blnDelete As Boolean

    For Each qdf In dbs.QueryDefs
        If qdf.Name = strQueryName Then
            blnDelete = True
            Exit For
        End If
    Next qdf

    If blnDelete Then dbs.QueryDefs.Delete strQueryName
End Sub


Private Function QueryDefRowCount(ByVal dbs As DAO.Database, ByVal strQueryName As String) As Long
    Dim rst As DAO.Recordset
    Set rst = dbs.OpenRecordset("SELECT Count(*) AS Cnt FROM [" & strQueryName & "]", dbOpenSnapshot)
    QueryDefRowCount = CLng(rst!cnt)
    rst.Close
End Function


Private Function QueryDefSchemaSignature(ByVal dbs As DAO.Database, ByVal strQueryName As String) As String
    Dim rst As DAO.Recordset
    Dim lngIndex As Long
    Dim strSig As String

    Set rst = dbs.OpenRecordset(strQueryName, dbOpenSnapshot)
    For lngIndex = 0 To rst.Fields.Count - 1
        strSig = strSig & rst.Fields(lngIndex).Name & ":" & _
            CStr(rst.Fields(lngIndex).Type) & ";"
    Next lngIndex
    rst.Close
    QueryDefSchemaSignature = strSig
End Function


Private Function NewArtifactCompareStats() As Dictionary
    Dim dStats As New Dictionary
    dStats.Add "total", CLng(0)
    dStats.Add "SAFE_DAO_COMPARE", CLng(0)
    dStats.Add "MISMATCH", CLng(0)
    dStats.Add "BLOCKED", CLng(0)
    dStats.Add "ACTION_OR_UNSUPPORTED", CLng(0)
    Set NewArtifactCompareStats = dStats
End Function


Private Sub IncrementArtifactCompareStats(ByVal dStats As Dictionary, ByVal strStatus As String)
    If Not dStats.Exists(strStatus) Then dStats.Add strStatus, CLng(0)
    dStats(strStatus) = CLng(dStats(strStatus)) + 1
    dStats("total") = CLng(dStats("total")) + 1
End Sub


Private Function ArtifactCompareLedgerLine(ByVal dResult As Dictionary) As String
    ArtifactCompareLedgerLine = "| `" & Replace(CStr(dResult("name")), "|", "\|") & "` | " & _
        CStr(dResult("status")) & " | " & _
        Nz(dResult("accessCount"), vbNullString) & " | " & _
        Nz(dResult("generatedCount"), vbNullString) & " | " & _
        Replace(Replace(Nz(dResult("detail"), vbNullString), vbCrLf, " "), "|", "\|") & _
        " |" & vbCrLf
End Function


Private Function UniqueInspectQueryName() As String
    Randomize
    UniqueInspectQueryName = "vcs_inspect_" & Format$(Now, "yyyymmddhhnnss") & _
        "_" & Right$("000000" & CStr(CLng(Rnd() * 1000000)), 6)
End Function


Private Function ValidateOneQuery(ByVal strQueryName As String, ByVal strArtifactRoot As String) As Dictionary
    Dim dResult As Dictionary
    Dim dMetadata As Dictionary
    Dim colRows As Collection
    Dim cComposer As clsQueryComposer
    Dim strGeneratedRaw As String
    Dim strGenerated As String
    Dim strAccessRaw As String
    Dim strAccess As String
    Dim strStatus As String
    Dim strReason As String
    Dim strArtifactFolder As String
    Dim strDiff As String
    Dim blnPassThrough As Boolean
    Dim colWarnings As Collection
    ' strCheckpoint records the last successful step. When a runtime error
    ' fires, the per-query handler reports it in the JSON result and writes
    ' it to error.txt so we can pinpoint exactly which step exploded
    ' without re-running with a debugger attached.
    Dim strCheckpoint As String

    Set dResult = New Dictionary
    dResult.Add "name", strQueryName

    LogUnhandledErrors ModuleName & ".ValidateOneQuery"
    On Error GoTo ErrHandler

    strCheckpoint = "init"
    Set dMetadata = New Dictionary
    strCheckpoint = "ReadMsysQueryRows"
    Set colRows = ReadMsysQueryRows(strQueryName, dMetadata)
    strCheckpoint = "HasConnectRow"
    ' A pass-through query may keep its connect string in Attribute 1, not Attribute 4.
    blnPassThrough = HasConnectRow(colRows) _
        Or (CurrentDb.QueryDefs(strQueryName).Type = dbQSQLPassThrough)

    strCheckpoint = "new clsQueryComposer"
    Set cComposer = New clsQueryComposer
    strCheckpoint = "ReconstructSQL"
    strGeneratedRaw = cComposer.ReconstructSQL(colRows)
    strCheckpoint = "QueryDefs.SQL"
    strAccessRaw = CurrentDb.QueryDefs(strQueryName).SQL

    ' Snapshot Log.ErrorCount before formatting so we can detect if the
    ' SQL formatter logged errors (e.g. "Unable to parse SQL after position N").
    Dim lngErrsBefore As Long
    lngErrsBefore = Log.ErrorCount

    strCheckpoint = "FormatSqlForComparison(generated)"
    strGenerated = FormatSqlForComparison(strGeneratedRaw, blnPassThrough)
    strCheckpoint = "FormatSqlForComparison(access)"
    strAccess = FormatSqlForComparison(strAccessRaw, blnPassThrough)
    strCheckpoint = "cComposer.Warnings"
    Set colWarnings = cComposer.Warnings

    strCheckpoint = "classify"
    If Log.ErrorCount > lngErrsBefore Then
        strStatus = STATUS_ERROR
        strReason = "SQL formatter reported error(s)"
    ElseIf cComposer.WarningCount > 0 Then
        strStatus = STATUS_FAIL_WARNING
        strReason = "builder emitted warning(s)"
    ElseIf NormalizeLineEndings(strGenerated) = NormalizeLineEndings(strAccess) Then
        strStatus = STATUS_PASS_EXACT
    ElseIf CanonicalSql(strGenerated) = CanonicalSql(strAccess) Then
        strStatus = STATUS_PASS_CANONICAL
        strReason = "canonical SQL matched"
    Else
        strStatus = STATUS_REVIEW
        strReason = "formatted SQL differs"
    End If

    strCheckpoint = "populate dResult"
    dResult.Add "status", strStatus
    If Len(strReason) > 0 Then dResult.Add "reason", strReason
    dResult.Add "warnings", colWarnings

    ' Report only: the builder comparison above never sees canonicalized SQL.
    strCheckpoint = "CheckCanonicalizer"
    If Not blnPassThrough Then dResult.Add "canonicalizer", CheckCanonicalizer(strQueryName, strGeneratedRaw)

    If strStatus = STATUS_REVIEW Or strStatus = STATUS_FAIL_WARNING Or strStatus = STATUS_ERROR Then
        strCheckpoint = "WriteQueryArtifacts"
        strArtifactFolder = WriteQueryArtifacts(strArtifactRoot, strQueryName, _
            strGeneratedRaw, strGenerated, strAccessRaw, strAccess, colRows, _
            dMetadata, colWarnings)
        dResult.Add "artifactFolder", strArtifactFolder
        strCheckpoint = "MakeUnifiedDiff"
        strDiff = MakeUnifiedDiff(strAccess, strGenerated, "access.sql", "generated.sql")
        strCheckpoint = "diff Add"
        If Len(strDiff) > DIFF_MAX_CHARS_FOR_JSON Then
            dResult.Add "diff", Left$(strDiff, DIFF_MAX_CHARS_FOR_JSON)
            dResult.Add "diffTruncated", True
        Else
            dResult.Add "diff", strDiff
            dResult.Add "diffTruncated", False
        End If
    End If

    Set ValidateOneQuery = dResult
    Exit Function

ErrHandler:
    Dim lngErrNumber As Long
    Dim strErrDescription As String
    lngErrNumber = Err.Number
    strErrDescription = Err.Description
    ' From here on, swallow secondary errors so a failure inside the
    ' artifact-writing fallback can't replace the original error info.
    On Error Resume Next
    If lngErrNumber = 3334 And strCheckpoint = "QueryDefs.SQL" Then
        strErrDescription = strErrDescription & _
            " (Access cannot render SQL for this query -- it may have corrupted internal storage.)"
    End If
    strArtifactFolder = WriteErrorArtifact(strArtifactRoot, strQueryName, _
        lngErrNumber, strErrDescription, strCheckpoint, _
        strGeneratedRaw, strGenerated, _
        strAccessRaw, strAccess, colRows, dMetadata)
    dResult.RemoveAll
    dResult.Add "name", strQueryName
    dResult.Add "status", STATUS_ERROR
    dResult.Add "reason", strErrDescription
    dResult.Add "errorNumber", CLng(lngErrNumber)
    dResult.Add "checkpoint", strCheckpoint
    If Len(strArtifactFolder) > 0 Then dResult.Add "artifactFolder", strArtifactFolder
    On Error GoTo 0
    Set ValidateOneQuery = dResult
End Function


Private Function ReadMsysQueryRows(ByVal strQueryName As String, ByVal dMetadata As Dictionary) As Collection
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim lngObjectId As Long
    Dim colRows As New Collection
    Dim dRow As Dictionary

    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset( _
        "SELECT Id, Flags, LvProp, LvExtra FROM MSysObjects " & _
        "WHERE Name=""" & DblQ(strQueryName) & """ AND Type=5", _
        dbOpenSnapshot)
    If rst.EOF Then Err.Raise vbObjectError + 601, ModuleName, _
        "Query '" & strQueryName & "' not found in MSysObjects."

    lngObjectId = CLng(rst!Id)
    dMetadata.Add "Name", strQueryName
    dMetadata.Add "ObjectId", lngObjectId
    dMetadata.Add "Flags", Nz(rst!Flags, 0)
    dMetadata.Add "HasLvProp", Not IsNull(rst!LvProp)
    dMetadata.Add "HasLvExtra", Not IsNull(rst!LvExtra)
    rst.Close

    Set rst = dbs.OpenRecordset( _
        "SELECT Attribute, [Order], Flag, Expression, Name1, Name2 " & _
        "FROM MSysQueries WHERE ObjectId=" & lngObjectId & _
        " ORDER BY Attribute, [Order]", _
        dbOpenSnapshot)
    Do While Not rst.EOF
        Set dRow = New Dictionary
        dRow.Add "Attribute", CLng(rst!Attribute)
        dRow.Add "OrderSeq", OrderSequence(rst.Fields("Order").Value)
        dRow.Add "OrderHex", OrderHexPrefix(rst.Fields("Order").Value)
        dRow.Add "Flag", NullIfNull(rst!Flag)
        dRow.Add "Expression", NullIfNull(rst!Expression)
        dRow.Add "Name1", NullIfNull(rst!Name1)
        dRow.Add "Name2", NullIfNull(rst!Name2)
        colRows.Add dRow
        rst.MoveNext
    Loop
    rst.Close

    Set ReadMsysQueryRows = colRows
End Function


Private Function HasConnectRow(ByVal colRows As Collection) As Boolean
    Dim dRow As Dictionary
    For Each dRow In colRows
        If CLng(dRow("Attribute")) = 4 Then
            If Not IsNull(dRow("Expression")) Then
                HasConnectRow = (Len(CStr(dRow("Expression"))) > 0)
                Exit Function
            End If
        End If
    Next dRow
End Function


Private Function FormatSqlForComparison(ByVal strSql As String, ByVal blnPassThrough As Boolean) As String
    ' Keep validation output quiet while a query name is printed with the
    ' line left open. clsSqlFormatter logs parse failures directly to the
    ' console, which interrupts the "query name ... [RESULT]" display.
    ' The canonical comparison below already normalizes whitespace outside
    ' literals, so raw line-ending normalization is sufficient here.
    FormatSqlForComparison = NormalizeLineEndings(strSql)
End Function


Private Function WriteQueryArtifacts(ByVal strArtifactRoot As String, ByVal strQueryName As String, _
    ByVal strGeneratedRaw As String, ByVal strGenerated As String, _
    ByVal strAccessRaw As String, ByVal strAccess As String, _
    ByVal colRows As Collection, ByVal dMetadata As Dictionary, _
    ByVal colWarnings As Collection) As String

    Dim strFolder As String
    Dim strDiff As String

    strFolder = strArtifactRoot & GetSafeFileName(strQueryName) & PathSep
    VerifyPath strFolder

    WriteFile strGenerated, strFolder & "generated.sql"
    WriteFile strAccess, strFolder & "access.sql"
    WriteFile strGeneratedRaw, strFolder & "generated.raw.sql"
    WriteFile strAccessRaw, strFolder & "access.raw.sql"
    WriteFile ConvertToJson(colRows, JSON_WHITESPACE), strFolder & "msysqueries.json"
    WriteFile ConvertToJson(dMetadata, JSON_WHITESPACE), strFolder & "metadata.json"
    WriteOptionalArtifact JoinCollection(colWarnings, vbCrLf), strFolder & "warnings.txt"
    strDiff = MakeUnifiedDiff(strAccess, strGenerated, "access.sql", "generated.sql")
    WriteOptionalArtifact strDiff, strFolder & "diff.txt"

    WriteQueryArtifacts = strFolder
End Function


Private Sub WriteOptionalArtifact(ByVal strText As String, ByVal strPath As String)
    If Len(strText) = 0 Then
        If FSO.FileExists(strPath) Then DeleteFile strPath
    Else
        WriteFile strText, strPath
    End If
End Sub


Private Function WriteErrorArtifact(ByVal strArtifactRoot As String, ByVal strQueryName As String, _
    ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
    ByVal strCheckpoint As String, _
    ByVal strGeneratedRaw As String, ByVal strGenerated As String, _
    ByVal strAccessRaw As String, ByVal strAccess As String, _
    ByVal colRows As Collection, ByVal dMetadata As Dictionary) As String

    Dim strFolder As String
    Dim strErrText As String
    strFolder = strArtifactRoot & GetSafeFileName(strQueryName) & PathSep
    VerifyPath strFolder

    strErrText = "Error " & CStr(lngErrNumber) & ": " & strErrDescription & vbCrLf & _
        "Checkpoint (last successful step): " & strCheckpoint & vbCrLf & _
        "(Failure occurred during the next operation after this checkpoint.)"
    WriteFile strErrText, strFolder & "error.txt"
    If Len(strGenerated) > 0 Then WriteFile strGenerated, strFolder & "generated.sql"
    If Len(strAccess) > 0 Then WriteFile strAccess, strFolder & "access.sql"
    If Len(strGeneratedRaw) > 0 Then WriteFile strGeneratedRaw, strFolder & "generated.raw.sql"
    If Len(strAccessRaw) > 0 Then WriteFile strAccessRaw, strFolder & "access.raw.sql"
    If Not colRows Is Nothing Then WriteFile ConvertToJson(colRows, JSON_WHITESPACE), strFolder & "msysqueries.json"
    If Not dMetadata Is Nothing Then WriteFile ConvertToJson(dMetadata, JSON_WHITESPACE), strFolder & "metadata.json"

    WriteErrorArtifact = strFolder
End Function


Private Function CanonicalSql(ByVal strSql As String) As String
    CanonicalSql = NormalizeOnAndOrder(TightenParenSpacing( _
        CollapseWhitespaceOutsideLiterals(NormalizeLineEndings(strSql))))
End Function


'---------------------------------------------------------------------------------------
' Procedure : TightenParenSpacing
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Drop a space after "(" and before ")" or ",", outside strings and
'           : brackets. Access stores `(a, b)` whether it was typed `( a, b )` or
'           : `DESC ,`, so the spacing is presentation only.
'---------------------------------------------------------------------------------------
'
Private Function TightenParenSpacing(ByVal strText As String) As String
    Dim lngI As Long
    Dim strCh As String
    Dim strNext As String
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean
    Dim cOut As New clsConcat

    For lngI = 1 To Len(strText)
        strCh = Mid$(strText, lngI, 1)

        If blnInString Then
            cOut.Add strCh
            If strCh = """" Then
                If lngI < Len(strText) And Mid$(strText, lngI + 1, 1) = """" Then
                    lngI = lngI + 1
                    cOut.Add """"
                Else
                    blnInString = False
                End If
            End If
        ElseIf blnInBracket Then
            cOut.Add strCh
            If strCh = "]" Then blnInBracket = False
        Else
            Select Case strCh
                Case """"
                    blnInString = True
                    cOut.Add strCh
                Case "["
                    blnInBracket = True
                    cOut.Add strCh
                Case " "
                    strNext = Mid$(strText, lngI + 1, 1)
                    If cOut.RightStr(1) <> "(" And strNext <> ")" And strNext <> "," Then cOut.Add strCh
                Case Else
                    cOut.Add strCh
            End Select
        End If
    Next lngI

    TightenParenSpacing = cOut.GetStr
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckCanonicalizer
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Report what clsSqlCanonicalizer would do to the builder's SQL for one
'           : query, and prove the rewrite safe: idempotent, parsed by Access, same
'           : decomposed structure, and (for read-only SELECTs) the same rows and
'           : columns. Changes nothing outside temporary query definitions.
'---------------------------------------------------------------------------------------
'
Private Function CheckCanonicalizer(ByVal strQueryName As String, ByVal strBuilderSql As String) As Dictionary

    Dim dCheck As Dictionary
    Dim cCanon As clsSqlCanonicalizer
    Dim strCanonical As String
    Dim strAgain As String
    Dim blnFailed As Boolean

    Set dCheck = New Dictionary
    Set cCanon = New clsSqlCanonicalizer

    strCanonical = cCanon.Canonicalize(strBuilderSql)
    If cCanon.Changed Then
        dCheck.Add "result", "changed"
    ElseIf Len(cCanon.DeclineReason) > 0 Then
        dCheck.Add "result", "declined"
        dCheck.Add "declineReason", cCanon.DeclineReason
    Else
        dCheck.Add "result", "unchanged"
    End If

    If cCanon.Changed Then
        strAgain = cCanon.Canonicalize(strCanonical)
        dCheck.Add "idempotent", (Not cCanon.Changed And Len(cCanon.DeclineReason) = 0 _
            And StrComp(strAgain, strCanonical, vbBinaryCompare) = 0)
        dCheck.Add "parse", CanonicalParseCheck(strBuilderSql, strCanonical)
        dCheck.Add "structure", CanonicalStructureCheck(strBuilderSql, strCanonical)
        dCheck.Add "behavior", CanonicalBehaviorCheck(strQueryName, strBuilderSql, strCanonical)
        dCheck.Add "canonicalSql", strCanonical
        blnFailed = (Not dCheck("idempotent")) _
            Or StartsWith(CStr(dCheck("parse")), "fail") _
            Or StartsWith(CStr(dCheck("structure")), "fail") _
            Or StartsWith(CStr(dCheck("behavior")), "fail")
    End If
    dCheck.Add "failed", blnFailed

    Set CheckCanonicalizer = dCheck

End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalParseCheck
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : The paste test: Access must accept the canonical SQL as a query
'           : definition. Skipped when the builder SQL itself does not parse.
'---------------------------------------------------------------------------------------
'
Private Function CanonicalParseCheck(ByVal strBefore As String, ByVal strAfter As String) As String
    Dim strError As String
    strError = SqlParseError(strBefore)
    If Len(strError) > 0 Then
        CanonicalParseCheck = "skipped: builder SQL does not parse (" & strError & ")"
        Exit Function
    End If
    strError = SqlParseError(strAfter)
    If Len(strError) > 0 Then
        CanonicalParseCheck = "fail: " & strError
    Else
        CanonicalParseCheck = "pass"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SqlParseError
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return the error Access raises when creating a temporary (unsaved)
'           : query definition from strSql, or an empty string when it parses.
'---------------------------------------------------------------------------------------
'
Public Function SqlParseError(ByVal strSql As String) As String
    Dim qdf As DAO.QueryDef
    LogUnhandledErrors ModuleName & ".SqlParseError"
    On Error Resume Next
    Set qdf = CurrentDb.CreateQueryDef(vbNullString, strSql)
    If Err.Number <> 0 Then SqlParseError = CStr(Err.Number) & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    Set qdf = Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalStructureCheck
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Decompose both SQL strings and compare the structure the parser reads
'           : from them. The generated .qdef text may differ only by parentheses
'           : (the WHERE/HAVING expressions); anything else is a failure.
'---------------------------------------------------------------------------------------
'
Private Function CanonicalStructureCheck(ByVal strBefore As String, ByVal strAfter As String) As String

    Dim cBefore As clsQueryComposer
    Dim cAfter As clsQueryComposer
    Dim strQdefBefore As String
    Dim strQdefAfter As String

    LogUnhandledErrors ModuleName & ".CanonicalStructureCheck"
    On Error GoTo ErrHandler

    Set cBefore = New clsQueryComposer
    Set cAfter = New clsQueryComposer
    cBefore.DecomposeSQL strBefore
    cAfter.DecomposeSQL strAfter

    If cBefore.QueryType <> cAfter.QueryType Then
        CanonicalStructureCheck = "fail: query type differs"
    ElseIf cBefore.IsDesignerCompatible <> cAfter.IsDesignerCompatible Then
        CanonicalStructureCheck = "fail: designer compatibility differs"
    Else
        strQdefBefore = cBefore.GenerateQdef(True, New Dictionary, New Dictionary)
        strQdefAfter = cAfter.GenerateQdef(True, New Dictionary, New Dictionary)
        If StrComp(WithoutParensOrSpace(strQdefBefore), WithoutParensOrSpace(strQdefAfter), vbBinaryCompare) = 0 Then
            CanonicalStructureCheck = "pass"
        Else
            CanonicalStructureCheck = "fail: decomposed structure differs"
        End If
    End If
    Exit Function

ErrHandler:
    CanonicalStructureCheck = "fail: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
End Function


' Quotes go too: the .qdef wraps long expressions into fixed-width quoted chunks,
' so removing a parenthesis moves every chunk boundary after it.
Private Function WithoutParensOrSpace(ByVal strText As String) As String
    Dim varChar As Variant
    For Each varChar In Array("(", ")", " ", vbTab, vbCr, vbLf, """")
        strText = Replace$(strText, CStr(varChar), vbNullString)
    Next varChar
    WithoutParensOrSpace = strText
End Function


'---------------------------------------------------------------------------------------
' Procedure : CanonicalBehaviorCheck
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : For read-only SELECTs, run both SQL strings and compare row count and
'           : column schema. Skipped when the builder SQL cannot run on its own
'           : (parameters, form references, unavailable links).
'---------------------------------------------------------------------------------------
'
Private Function CanonicalBehaviorCheck(ByVal strQueryName As String, _
    ByVal strBefore As String, ByVal strAfter As String) As String

    Dim dbs As DAO.Database
    Dim strSelectBefore As String
    Dim strSelectAfter As String
    Dim strNameBefore As String
    Dim strNameAfter As String
    Dim lngCountBefore As Long
    Dim lngCountAfter As Long
    Dim strSchemaBefore As String
    Dim strSchemaAfter As String
    Dim blnAfterStage As Boolean

    strSelectBefore = ExtractSourceSelect(strBefore)
    strSelectAfter = ExtractSourceSelect(strAfter)
    If Len(strSelectBefore) = 0 Or Len(strSelectAfter) = 0 Then
        CanonicalBehaviorCheck = "skipped: not a read-only SELECT"
        Exit Function
    End If

    Set dbs = CurrentDb
    strNameBefore = "vcs_canon_before_" & UniqueHashSuffix(strQueryName)
    strNameAfter = "vcs_canon_after_" & UniqueHashSuffix(strQueryName)

    LogUnhandledErrors ModuleName & ".CanonicalBehaviorCheck"
    On Error GoTo ErrHandler

    ' Both definitions exist before either is measured: saving a query adds rows to
    ' system tables such as MSysQueries, which a query may itself be reading.
    DeleteQueryDefIfExists dbs, strNameBefore
    DeleteQueryDefIfExists dbs, strNameAfter
    dbs.CreateQueryDef strNameBefore, strSelectBefore
    blnAfterStage = True
    dbs.CreateQueryDef strNameAfter, strSelectAfter
    blnAfterStage = False
    lngCountBefore = QueryDefRowCount(dbs, strNameBefore)
    strSchemaBefore = QueryDefSchemaSignature(dbs, strNameBefore)

    blnAfterStage = True
    lngCountAfter = QueryDefRowCount(dbs, strNameAfter)
    strSchemaAfter = QueryDefSchemaSignature(dbs, strNameAfter)

    If lngCountBefore <> lngCountAfter Then
        CanonicalBehaviorCheck = "fail: row count " & lngCountBefore & " vs " & lngCountAfter
    ElseIf strSchemaBefore <> strSchemaAfter Then
        CanonicalBehaviorCheck = "fail: column schema differs"
    Else
        CanonicalBehaviorCheck = "pass"
    End If

CleanUp:
    On Error Resume Next
    DeleteQueryDefIfExists dbs, strNameBefore
    DeleteQueryDefIfExists dbs, strNameAfter
    Err.Clear
    On Error GoTo 0
    Exit Function

ErrHandler:
    If blnAfterStage Then
        CanonicalBehaviorCheck = "fail: " & CStr(Err.Number) & " " & Err.Description
    Else
        CanonicalBehaviorCheck = "skipped: builder SQL does not run (" & CStr(Err.Number) & " " & Err.Description & ")"
    End If
    Resume CleanUp
End Function


Private Function NewCanonicalizerStats() As Dictionary
    Dim dStats As New Dictionary
    dStats.Add "total", CLng(0)
    dStats.Add "changed", CLng(0)
    dStats.Add "unchanged", CLng(0)
    dStats.Add "declined", CLng(0)
    dStats.Add "failures", CLng(0)
    Set NewCanonicalizerStats = dStats
End Function


Private Sub TallyCanonicalizer(ByVal dStats As Dictionary, ByVal dCheck As Dictionary)
    dStats("total") = CLng(dStats("total")) + 1
    dStats(CStr(dCheck("result"))) = CLng(dStats(CStr(dCheck("result")))) + 1
    If dCheck("failed") Then dStats("failures") = CLng(dStats("failures")) + 1
End Sub


Private Function CanonicalizerLedgerLine(ByVal strQueryName As String, ByVal dCheck As Dictionary) As String
    Dim strResult As String
    strResult = CStr(dCheck("result"))
    If dCheck.Exists("declineReason") Then strResult = strResult & ": " & CStr(dCheck("declineReason"))
    If dCheck("failed") Then strResult = "**FAILED** " & strResult
    CanonicalizerLedgerLine = "| `" & Replace(strQueryName, "|", "\|") & "` | " & _
        LedgerCell(strResult) & " | " & _
        LedgerCell(DictValue(dCheck, "idempotent")) & " | " & _
        LedgerCell(DictValue(dCheck, "parse")) & " | " & _
        LedgerCell(DictValue(dCheck, "structure")) & " | " & _
        LedgerCell(DictValue(dCheck, "behavior")) & " |" & vbCrLf
End Function


Private Function DictValue(ByVal dItems As Dictionary, ByVal strKey As String) As String
    If dItems.Exists(strKey) Then DictValue = CStr(dItems(strKey))
End Function


Private Function LedgerCell(ByVal strText As String) As String
    LedgerCell = Replace(Replace(strText, vbCrLf, " "), "|", "\|")
End Function


Private Function NormalizeOnAndOrder(ByVal strSql As String) As String
    Dim lngOn As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strOut As String
    Dim strCondition As String
    Dim strCanonical As String

    lngOn = InStr(1, strSql, " ON ", vbTextCompare)
    Do While lngOn > 0
        lngStart = lngOn + 4
        lngEnd = FindOnConditionEnd(strSql, lngStart)
        strOut = strOut & Left$(strSql, lngStart - 1)
        strCondition = Mid$(strSql, lngStart, lngEnd - lngStart)
        strCanonical = CanonicalAndCondition(strCondition)
        strOut = strOut & strCanonical
        strSql = Mid$(strSql, lngEnd)
        lngOn = InStr(1, strSql, " ON ", vbTextCompare)
    Loop

    NormalizeOnAndOrder = strOut & strSql
End Function


Private Function FindOnConditionEnd(ByVal strSql As String, ByVal lngStart As Long) As Long
    Dim lngI As Long
    Dim lngDepth As Long
    Dim strCh As String
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean

    For lngI = lngStart To Len(strSql)
        strCh = Mid$(strSql, lngI, 1)

        If blnInString Then
            If strCh = """" Then
                If lngI < Len(strSql) And Mid$(strSql, lngI + 1, 1) = """" Then
                    lngI = lngI + 1
                Else
                    blnInString = False
                End If
            End If
        ElseIf blnInBracket Then
            If strCh = "]" Then blnInBracket = False
        Else
            Select Case strCh
                Case """"
                    blnInString = True
                Case "["
                    blnInBracket = True
                Case "("
                    lngDepth = lngDepth + 1
                Case ")"
                    If lngDepth = 0 Then
                        FindOnConditionEnd = lngI
                        Exit Function
                    End If
                    lngDepth = lngDepth - 1
                Case ";"
                    If lngDepth = 0 Then
                        FindOnConditionEnd = lngI
                        Exit Function
                    End If
                Case " "
                    If lngDepth = 0 And IsOnConditionBoundary(strSql, lngI) Then
                        FindOnConditionEnd = lngI
                        Exit Function
                    End If
            End Select
        End If
    Next lngI

    FindOnConditionEnd = Len(strSql) + 1
End Function


Private Function IsOnConditionBoundary(ByVal strSql As String, ByVal lngPos As Long) As Boolean
    Dim strRest As String
    strRest = Mid$(strSql, lngPos)
    IsOnConditionBoundary = _
        StartsWith(strRest, " INNER JOIN ", vbTextCompare) Or _
        StartsWith(strRest, " LEFT JOIN ", vbTextCompare) Or _
        StartsWith(strRest, " RIGHT JOIN ", vbTextCompare) Or _
        StartsWith(strRest, " WHERE ", vbTextCompare) Or _
        StartsWith(strRest, " GROUP BY ", vbTextCompare) Or _
        StartsWith(strRest, " HAVING ", vbTextCompare) Or _
        StartsWith(strRest, " ORDER BY ", vbTextCompare) Or _
        StartsWith(strRest, " PIVOT ", vbTextCompare)
End Function


Private Function CanonicalAndCondition(ByVal strCondition As String) As String
    Dim asParts() As String
    Dim lngCount As Long
    Dim colLeaves As New Collection
    Dim varLeaf As Variant

    ' Export encloses each table pair's ON conditions as one term at 5.2.0
    ' (`((a) AND (b)) AND (c)`); QueryDefs.SQL writes a flat `(a) AND (b) AND (c)`.
    AddAndConditionLeaves colLeaves, strCondition
    If colLeaves.Count <= 1 Then
        CanonicalAndCondition = Trim$(strCondition)
    Else
        ReDim asParts(0 To colLeaves.Count - 1)
        For Each varLeaf In colLeaves
            asParts(lngCount) = CStr(varLeaf)
            lngCount = lngCount + 1
        Next varLeaf
        SortStrings asParts, lngCount
        CanonicalAndCondition = JoinFirstN(asParts, lngCount, " AND ")
    End If
End Function


Private Sub AddAndConditionLeaves(colLeaves As Collection, ByVal strCondition As String)
    Dim asParts() As String
    Dim lngCount As Long
    Dim lngPart As Long

    With New clsSqlSyntax
        strCondition = .StripOuterParens(strCondition)
    End With
    asParts = SplitTopLevelAnd(strCondition, lngCount)
    If lngCount <= 1 Then
        colLeaves.Add strCondition
    Else
        For lngPart = 0 To lngCount - 1
            AddAndConditionLeaves colLeaves, asParts(lngPart)
        Next lngPart
    End If
End Sub


Private Function SplitTopLevelAnd(ByVal strText As String, ByRef lngCount As Long) As String()
    Dim asParts() As String
    Dim lngStart As Long
    Dim lngI As Long
    Dim lngDepth As Long
    Dim strCh As String
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean

    ReDim asParts(0 To 0)
    lngStart = 1

    For lngI = 1 To Len(strText)
        strCh = Mid$(strText, lngI, 1)

        If blnInString Then
            If strCh = """" Then
                If lngI < Len(strText) And Mid$(strText, lngI + 1, 1) = """" Then
                    lngI = lngI + 1
                Else
                    blnInString = False
                End If
            End If
        ElseIf blnInBracket Then
            If strCh = "]" Then blnInBracket = False
        Else
            Select Case strCh
                Case """"
                    blnInString = True
                Case "["
                    blnInBracket = True
                Case "("
                    lngDepth = lngDepth + 1
                Case ")"
                    If lngDepth > 0 Then lngDepth = lngDepth - 1
                Case " "
                    If lngDepth = 0 And Mid$(strText, lngI, 5) = " AND " Then
                        AddSplitPart asParts, lngCount, Mid$(strText, lngStart, lngI - lngStart)
                        lngI = lngI + 4
                        lngStart = lngI + 1
                    End If
            End Select
        End If
    Next lngI

    AddSplitPart asParts, lngCount, Mid$(strText, lngStart)
    SplitTopLevelAnd = asParts
End Function


Private Sub AddSplitPart(ByRef asParts() As String, ByRef lngCount As Long, ByVal strPart As String)
    If lngCount > UBound(asParts) Then ReDim Preserve asParts(0 To lngCount)
    asParts(lngCount) = Trim$(strPart)
    lngCount = lngCount + 1
End Sub


Private Sub SortStrings(ByRef asValues() As String, ByVal lngCount As Long)
    Dim lngI As Long
    Dim lngJ As Long
    Dim strTmp As String

    For lngI = 0 To lngCount - 2
        For lngJ = lngI + 1 To lngCount - 1
            If StrComp(asValues(lngJ), asValues(lngI), vbTextCompare) < 0 Then
                strTmp = asValues(lngI)
                asValues(lngI) = asValues(lngJ)
                asValues(lngJ) = strTmp
            End If
        Next lngJ
    Next lngI
End Sub


Private Function JoinFirstN(ByRef asValues() As String, ByVal lngCount As Long, _
    ByVal strDelimiter As String) As String

    Dim lngI As Long
    Dim strOut As String

    For lngI = 0 To lngCount - 1
        If lngI > 0 Then strOut = strOut & strDelimiter
        strOut = strOut & asValues(lngI)
    Next lngI
    JoinFirstN = strOut
End Function


Private Function CollapseWhitespaceOutsideLiterals(ByVal strText As String) As String
    Dim lngI As Long
    Dim strCh As String
    Dim strOut As String
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean
    Dim blnLastWasSpace As Boolean

    For lngI = 1 To Len(strText)
        strCh = Mid$(strText, lngI, 1)

        If blnInString Then
            strOut = strOut & strCh
            If strCh = """" Then
                If lngI < Len(strText) And Mid$(strText, lngI + 1, 1) = """" Then
                    lngI = lngI + 1
                    strOut = strOut & """"
                Else
                    blnInString = False
                End If
            End If
        ElseIf blnInBracket Then
            strOut = strOut & strCh
            If strCh = "]" Then blnInBracket = False
        Else
            Select Case strCh
                Case """"
                    blnInString = True
                    blnLastWasSpace = False
                    strOut = strOut & strCh
                Case "["
                    blnInBracket = True
                    blnLastWasSpace = False
                    strOut = strOut & strCh
                Case " ", vbTab, vbCr, vbLf
                    If Not blnLastWasSpace Then
                        strOut = strOut & " "
                        blnLastWasSpace = True
                    End If
                Case Else
                    blnLastWasSpace = False
                    strOut = strOut & strCh
            End Select
        End If
    Next lngI

    CollapseWhitespaceOutsideLiterals = Trim$(strOut)
End Function


Private Function MakeUnifiedDiff(ByVal strA As String, ByVal strB As String, _
    ByVal strLabelA As String, ByVal strLabelB As String) As String

    Dim aA() As String
    Dim aB() As String
    Dim lngMax As Long
    Dim lngI As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim cc As New clsConcat
    Dim blnInHunk As Boolean
    Dim blnDiffer As Boolean

    aA = Split(NormalizeLineEndings(strA), vbLf)
    aB = Split(NormalizeLineEndings(strB), vbLf)
    lngMax = IIf(UBound(aA) > UBound(aB), UBound(aA), UBound(aB))

    cc.Add "--- ", strLabelA, vbCrLf
    cc.Add "+++ ", strLabelB, vbCrLf

    For lngI = 0 To lngMax
        blnDiffer = (SafeLine(aA, lngI) <> SafeLine(aB, lngI))
        If blnDiffer Then
            If Not blnInHunk Then
                lngStart = MaxLong(0, lngI - DIFF_CONTEXT_LINES)
                lngEnd = MinLong(lngMax, lngI + DIFF_CONTEXT_LINES)
                cc.Add "@@ lines ", CStr(lngStart + 1), "-", CStr(lngEnd + 1), " @@", vbCrLf
                blnInHunk = True
            Else
                lngEnd = MinLong(lngMax, lngI + DIFF_CONTEXT_LINES)
            End If
        End If
        If blnInHunk And lngI <= lngEnd Then
            If SafeLine(aA, lngI) <> SafeLine(aB, lngI) Then
                cc.Add "-", SafeLine(aA, lngI), vbCrLf
                cc.Add "+", SafeLine(aB, lngI), vbCrLf
            Else
                cc.Add " ", SafeLine(aA, lngI), vbCrLf
            End If
        ElseIf blnInHunk And lngI > lngEnd Then
            blnInHunk = False
        End If
    Next lngI

    MakeUnifiedDiff = cc.GetStr
End Function


Private Function NormalizeLineEndings(ByVal strText As String) As String
    NormalizeLineEndings = Replace$(Replace$(strText, vbCrLf, vbLf), vbCr, vbLf)
End Function


Private Function SafeLine(ByRef arr() As String, ByVal idx As Long) As String
    If idx < LBound(arr) Or idx > UBound(arr) Then
        SafeLine = vbNullString
    Else
        SafeLine = arr(idx)
    End If
End Function


Private Function NewValidationStats() As Dictionary
    Dim d As New Dictionary
    d.Add "total", CLng(0)
    d.Add STATUS_PASS_EXACT, CLng(0)
    d.Add STATUS_PASS_CANONICAL, CLng(0)
    d.Add STATUS_REVIEW, CLng(0)
    d.Add STATUS_FAIL_WARNING, CLng(0)
    d.Add STATUS_ERROR, CLng(0)
    d.Add "elapsedSeconds", CDbl(0)
    Set NewValidationStats = d
End Function


Private Sub IncrementStatus(ByVal dStats As Dictionary, ByVal strStatus As String)
    If Not dStats.Exists(strStatus) Then dStats.Add strStatus, CLng(0)
    dStats(strStatus) = CLng(dStats(strStatus)) + 1
End Sub


Private Function ValidationSucceeded(ByVal dStats As Dictionary) As Boolean
    ValidationSucceeded = (CLng(dStats(STATUS_REVIEW)) = 0 And _
                           CLng(dStats(STATUS_FAIL_WARNING)) = 0 And _
                           CLng(dStats(STATUS_ERROR)) = 0)
End Function


Private Sub LogQueryDetailFileOnly(ByVal dQueryResult As Dictionary)
    Dim strStatus As String
    Dim strLine As String

    strStatus = CStr(dQueryResult("status"))
    If strStatus = STATUS_PASS_EXACT Or strStatus = STATUS_PASS_CANONICAL Then Exit Sub

    strLine = strStatus & " " & CStr(dQueryResult("name"))
    If dQueryResult.Exists("reason") Then strLine = strLine & " reason=""" & CStr(dQueryResult("reason")) & """"
    If dQueryResult.Exists("checkpoint") Then strLine = strLine & " checkpoint=""" & CStr(dQueryResult("checkpoint")) & """"
    If dQueryResult.Exists("artifactFolder") Then strLine = strLine & " artifact=""" & CStr(dQueryResult("artifactFolder")) & """"
    Log.Add strLine, False
End Sub


Private Sub LogSummaryLine(ByVal strLabel As String, ByVal strValue As String, ByVal strColor As String)
    If Len(strColor) > 0 Then
        Log.Add strLabel, True, False
        Log.Add strValue, True, True, strColor, True
    Else
        Log.Add strLabel & strValue
    End If
End Sub


Private Function ConsoleQueryName(ByVal strQueryName As String) As String
    Dim strName As String
    Dim lngWidth As Long

    lngWidth = Log.PadLength
    If lngWidth < 5 Then lngWidth = 30
    strName = strQueryName
    If Len(strName) > lngWidth - 1 Then
        strName = Left$(strName, lngWidth - 4) & "..."
    End If
    ConsoleQueryName = strName
End Function


Private Function ResultLabel(ByVal strStatus As String) As String
    Select Case strStatus
        Case STATUS_PASS_EXACT, STATUS_PASS_CANONICAL
            ResultLabel = "[PASS]"
        Case STATUS_REVIEW
            ResultLabel = "[REVIEW]"
        Case STATUS_FAIL_WARNING
            ResultLabel = "[FAIL]"
        Case STATUS_ERROR
            ResultLabel = "[ERROR]"
        Case Else
            ResultLabel = "[" & strStatus & "]"
    End Select
End Function


Private Function ResultColor(ByVal strStatus As String) As String
    Select Case strStatus
        Case STATUS_PASS_EXACT, STATUS_PASS_CANONICAL
            ResultColor = vbNullString
        Case STATUS_REVIEW
            ResultColor = "orange"
        Case STATUS_FAIL_WARNING, STATUS_ERROR
            ResultColor = "red"
        Case Else
            ResultColor = vbNullString
    End Select
End Function


Private Function JoinCollection(ByVal col As Collection, ByVal strDelimiter As String) As String
    Dim lngI As Long
    Dim strOut As String
    If col Is Nothing Then Exit Function
    For lngI = 1 To col.Count
        If lngI > 1 Then strOut = strOut & strDelimiter
        strOut = strOut & CStr(col(lngI))
    Next lngI
    JoinCollection = strOut
End Function


Private Function NullIfNull(ByVal varValue As Variant) As Variant
    If IsNull(varValue) Then
        NullIfNull = Null
    Else
        NullIfNull = varValue
    End If
End Function


Private Function OrderSequence(ByVal varOrder As Variant) As Variant
    On Error GoTo Failed
    If IsNull(varOrder) Then
        OrderSequence = Null
    ElseIf IsArray(varOrder) Then
        OrderSequence = CLng(varOrder(LBound(varOrder) + 3))
    Else
        OrderSequence = Null
    End If
    Exit Function
Failed:
    OrderSequence = Null
End Function


Private Function OrderHexPrefix(ByVal varOrder As Variant) As Variant
    On Error GoTo Failed
    Dim lngI As Long
    Dim lngEnd As Long
    Dim strOut As String
    If IsNull(varOrder) Or Not IsArray(varOrder) Then
        OrderHexPrefix = Null
        Exit Function
    End If
    lngEnd = MinLong(UBound(varOrder), LBound(varOrder) + 15)
    For lngI = LBound(varOrder) To lngEnd
        If Len(strOut) > 0 Then strOut = strOut & " "
        strOut = strOut & Right$("0" & Hex$(CLng(varOrder(lngI))), 2)
    Next lngI
    OrderHexPrefix = strOut
    Exit Function
Failed:
    OrderHexPrefix = Null
End Function


Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function


Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function


Public Sub TestCanonicalSql_InsertColumnSpacing()
    TestAssert CanonicalSql("INSERT INTO t ( a, b ) SELECT u.a, u.b FROM u;") = _
        CanonicalSql("INSERT INTO t (a, b) SELECT u.a, u.b FROM u;"), _
        "Access's ( a, b ) matches (a, b)"
End Sub


Public Sub TestCanonicalSql_OrderBySpacing()
    TestAssert CanonicalSql("SELECT a FROM t ORDER BY a DESC , b;") = _
        CanonicalSql("SELECT a FROM t ORDER BY a DESC, b;"), _
        "Access's DESC , matches DESC,"
End Sub


Public Sub TestCanonicalSql_UpdateSetLine()
    TestAssert CanonicalSql("UPDATE t SET t.a = 1;") = _
        CanonicalSql("UPDATE t" & vbCrLf & "SET t.a = 1;"), _
        "SET on the UPDATE line matches SET on its own line"
End Sub


Public Sub TestCanonicalSql_DroppedAliasBracketIsReview()
    TestAssert CanonicalSql("SELECT t.a AS [Option] FROM t;") <> _
        CanonicalSql("SELECT t.a AS Option FROM t;"), _
        "a dropped alias bracket is still a difference to review"
End Sub


Public Sub TestCheckCanonicalizer_KnownGoodQuery()
    Dim dCheck As Dictionary
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    Set dCheck = CheckCanonicalizer("vcsTestCanonicalizer", _
        "SELECT MSysObjects.Name FROM MSysObjects WHERE ((MSysObjects.Type)=1);")
    TestAssert dCheck("result") = "changed", "designer wrappers are removed"
    TestAssert dCheck("idempotent"), "idempotent"
    TestAssert dCheck("parse") = "pass", "parse: " & dCheck("parse")
    TestAssert dCheck("structure") = "pass", "structure: " & dCheck("structure")
    TestAssert dCheck("behavior") = "pass", "behavior: " & dCheck("behavior")
    TestAssert dCheck("canonicalSql") = "SELECT MSysObjects.Name FROM MSysObjects WHERE MSysObjects.Type=1;", _
        "canonical SQL"
    TestAssert Not dCheck("failed"), "not a failure"

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "CheckCanonicalizer " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp
End Sub


Public Sub TestCheckCanonicalizer_UnchangedQuery()
    Dim dCheck As Dictionary
    Set dCheck = CheckCanonicalizer("vcsTestCanonicalizer", "SELECT MSysObjects.Name FROM MSysObjects;")
    TestAssert dCheck("result") = "unchanged", "nothing to remove"
    TestAssert Not dCheck.Exists("parse"), "no checks run for an unchanged query"
    TestAssert Not dCheck("failed"), "not a failure"
End Sub


Public Sub TestCanonicalizerReport_CountsAndFlagsFailure()
    Dim dStats As Dictionary
    Dim dGood As Dictionary
    Dim dBad As Dictionary
    Dim strLine As String

    Set dStats = NewCanonicalizerStats()
    Set dGood = New Dictionary
    dGood.Add "result", "unchanged"
    dGood.Add "failed", False
    Set dBad = New Dictionary
    dBad.Add "result", "changed"
    dBad.Add "idempotent", True
    dBad.Add "parse", "pass"
    dBad.Add "structure", "fail: decomposed structure differs"
    dBad.Add "behavior", "pass"
    dBad.Add "failed", True

    TallyCanonicalizer dStats, dGood
    TallyCanonicalizer dStats, dBad
    TestAssert dStats("total") = 2, "both queries counted"
    TestAssert dStats("unchanged") = 1 And dStats("changed") = 1, "results tallied by kind"
    TestAssert dStats("failures") = 1, "the failed check is counted"

    strLine = CanonicalizerLedgerLine("qry|Example", dBad)
    TestAssert InStr(1, strLine, "**FAILED** changed", vbBinaryCompare) > 0, "ledger flags the failure"
    TestAssert InStr(1, strLine, "qry\|Example", vbBinaryCompare) > 0, "pipe in a name is escaped"
    TestAssert InStr(1, strLine, "fail: decomposed structure differs", vbBinaryCompare) > 0, _
        "ledger shows which check failed"
End Sub
