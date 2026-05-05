Attribute VB_Name = "modTestRoundtrip"
'---------------------------------------------------------------------------------------
' Module    : modTestRoundtrip
' Author    : Adam Waller / VCS contributors
' Date      : 4/24/2026
' Purpose   : Generalized object round-trip regression harness.
'           :
'           : Iterates over a folder of fixture files (one subfolder per
'           : IDbComponent type) and, for each fixture, performs:
'           :
'           :   Pass 1  Import (sandbox name) -> Export A -> compare to fixture
'           :   Pass 2  Re-export from the same in-memory object -> compare to A
'           :
'           : Mismatches are reported as inline unified diffs (text content) or
'           : structured key/value diffs (JSON metadata). Fixture files in the
'           : repo are the source of truth -- they are never modified except via
'           : an explicit rebaseline pass (blnRebaseline:=True).
'           :
'           : v1 only routes .sql query fixtures through clsDbQuery. The
'           : enumeration / dispatcher / sandboxing / cleanup machinery is
'           : intentionally generic so additional IDbComponent types
'           : (forms, reports, modules, table data ...) can be added by
'           : implementing a single per-type round-trip helper.
'           :
'           : Output:
'           :  - Live progress through Log singleton (frmVCSMain console).
'           :  - Per-session log file: <SourcePath>\logs\ObjectRoundtrip_<opId>.log
'           :  - Return value is a JSON document summarizing every fixture.
'           :
'           : This harness is meant to be run from the development copy of the
'           : add-in (not as a loaded add-in inside a host database). It uses
'           : the Operation/Log infrastructure but begins its own operation,
'           : so it should not be invoked while another VCS operation is
'           : in progress.
'           :
'           : External callers (Immediate Window in a host database, MCP, CI)
'           : invoke this through the public API: VCS.RunRoundtripTests().
'           : Option Private Module hides the internals from cross-project
'           : Application.Run lookups; the public functions in this module
'           : are still callable by other modules within the add-in project
'           : (notably clsVersionControl.RunRoundtripTests).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module

Private Const ModuleName As String = "modTestRoundtrip"

' Prefixes used to make sandboxed objects easy to spot and clean up.
Private Const TEST_PREFIX As String = "vcs_test_"
Private Const SCAFFOLD_PREFIX As String = "vcs_scaffold_"

' Default fixture root, relative to the repo root (i.e. CodeProject.Path's
' parent when running directly from "<repo>\Version Control.accda").
Private Const FIXTURE_RELATIVE_PATH As String = "Testing\Fixtures"

' Number of context lines emitted around each diff hunk.
Private Const DIFF_CONTEXT_LINES As Long = 2

' Maximum diff payload (per fixture) carried back in the JSON result. Full
' diffs are always written to the per-session log file regardless of this cap.
Private Const DIFF_MAX_CHARS_FOR_JSON As Long = 4000

' Names of scaffold objects loaded for the duration of the run, in the order
' they were imported. Used by UnloadScaffold to drop them in reverse.
Private m_colScaffoldQueries As Collection


'---------------------------------------------------------------------------------------
' Procedure : RunObjectRoundtripTests
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Main entry point for the harness. Iterates fixtures under
'           : strFixtureFolder and returns a JSON result document.
'           :
'           : The folder layout is:
'           :   <strFixtureFolder>\
'           :     _scaffold\        ' supporting objects (loaded once, dropped at end)
'           :     queries\          ' .sql + .json pairs (one fixture per pair)
'           :       <category>\     ' optional grouping subfolders
'           :     forms\, reports\, ... (future)
'           :
'           : When blnRebaseline is True, the harness writes any newly produced
'           : "after" file back over its fixture counterpart instead of failing
'           : the comparison. Use with care -- inspect git diff afterwards.
'---------------------------------------------------------------------------------------
'
Public Function RunObjectRoundtripTests(Optional ByVal strFixtureFolder As String = vbNullString, _
    Optional ByVal blnRebaseline As Boolean = False) As String

    Dim dResult As Dictionary
    Dim dStats As Dictionary
    Dim colResults As Collection
    Dim strScratch As String
    Dim blnIndexWasDisabled As Boolean
    Dim blnOperationOwned As Boolean
    Dim eimPriorMode As eInteractionMode
    Dim sngStart As Single

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary
    Set dStats = New Dictionary
    Set colResults = New Collection
    Set m_colScaffoldQueries = New Collection

    ' Resolve fixture folder (default to our shipped corpus).
    If Len(strFixtureFolder) = 0 Then strFixtureFolder = GetDefaultFixtureRoot()
    If Right$(strFixtureFolder, 1) <> PathSep Then strFixtureFolder = strFixtureFolder & PathSep

    If Not FSO.FolderExists(strFixtureFolder) Then
        dResult.Add "success", False
        dResult.Add "error", "Fixture folder not found: " & strFixtureFolder
        RunObjectRoundtripTests = ConvertToJson(dResult)
        Exit Function
    End If

    ' Begin our own operation, then flip the singleton InteractionMode to
    ' eimSilent for the duration of the run. Cache the prior mode so we can
    ' restore it in CleanUp / ErrHandler. Without silent mode, any MsgBox2
    ' call (including the modal prompts clsLog raises for eelCritical /
    ' eelError on an import failure) would block the harness waiting for a
    ' user click and stall an unattended test run.
    If Not Operation.Begin(eotOther) Then
        dResult.Add "success", False
        dResult.Add "error", "Could not begin test operation (another operation may be running)."
        RunObjectRoundtripTests = ConvertToJson(dResult)
        Exit Function
    End If
    blnOperationOwned = True
    eimPriorMode = Operation.InteractionMode
    Operation.InteractionMode = eimSilent

    ' Disable the index for the duration of the run so test imports/exports
    ' do not pollute vcs-index.idx. Restore the prior value at the end.
    blnIndexWasDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True

    ' Configure logging: route output through the main console form (if open)
    ' and write a dedicated session log file alongside the fixture folder.
    Log.Clear
    Log.SourcePath = strFixtureFolder
    Log.Active = True
    Perf.StartTiming
    sngStart = Perf.MicroTimer

    With Log
        .Spacer
        .Add T("Object Round-Trip Regression Harness")
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Fixture Folder: {0}", var0:=strFixtureFolder)
        If blnRebaseline Then .Add T("MODE: REBASELINE (fixtures may be overwritten)")
        .Add T("Started: {0}", var0:=Format$(Now, "yyyy-mm-dd hh:nn:ss"))
        .Spacer
    End With

    ' Sanity check: this harness only handles the new SQL/JSON query format.
    If Options.ExportFormatVersion < EFV_5_0_0 Then
        Log.Error eelError, T( _
            "Round-trip harness requires Export Format Version {0} or later (current: {1}).", _
            var0:=CStr(EFV_5_0_0), var1:=CStr(Options.ExportFormatVersion)), _
            ModuleName & ".RunObjectRoundtripTests"
        GoTo CleanUp
    End If

    ' Aggressive cleanup: drop any leftover sandbox objects from a prior run
    ' that may have crashed before its own cleanup ran.
    CleanupStaleObjects

    ' Provision a per-run scratch folder. All Pass 1 / Pass 2 outputs land here
    ' so they can be diffed against the canonical fixtures and inspected after
    ' a failure.
    strScratch = ProvisionScratchFolder(strFixtureFolder)
    Log.Add T("Scratch folder: {0}", var0:=strScratch), False

    ' Pre-load shared supporting objects (if any).
    LoadScaffold strFixtureFolder & "_scaffold" & PathSep

    ' Currently only queries are supported; forms/reports/etc. would be added
    ' here as additional Run<Type>Fixtures calls populating colResults.
    RunQueryFixtures strFixtureFolder & "queries" & PathSep, strScratch, blnRebaseline, colResults

CleanUp:
    ' Drop scaffold objects.
    UnloadScaffold

    ' Compute summary statistics.
    BuildStatsDict colResults, dStats
    dStats.Add "elapsedSeconds", Round(Perf.MicroTimer - sngStart, 3)

    With Log
        .Spacer
        .Add T("Round-trip summary"), True, , "blue", True
        .Add T("  Total fixtures:    {0}", var0:=CStr(dStats("total")))
        .Add T("  Passed:            {0}", var0:=CStr(dStats("passed")))
        .Add T("  Failed:            {0}", var0:=CStr(dStats("failed"))), , , _
            IIf(dStats("failed") > 0, "red", vbNullString), (dStats("failed") > 0)
        .Add T("  Skipped:           {0}", var0:=CStr(dStats("skipped")))
        .Add T("  Errors:            {0}", var0:=CStr(dStats("errors"))), , , _
            IIf(dStats("errors") > 0, "red", vbNullString), (dStats("errors") > 0)
        .Add T("  Elapsed (s):       {0}", var0:=CStr(dStats("elapsedSeconds")))
        .Spacer
    End With

    Perf.EndTiming

    ' Persist the per-session log file with our custom prefix so it is easy
    ' to distinguish from Export/Build/Merge logs.
    On Error Resume Next
    Log.SaveFile FSO.BuildPath(strFixtureFolder & "logs", _
        "ObjectRoundtrip_" & Log.OperationId & ".log")
    Log.Active = False
    Log.Flush
    On Error GoTo 0

    ' Restore VCSIndex disabled state and the prior InteractionMode. We
    ' restore InteractionMode here (rather than relying on Operation.Finish)
    ' because it is now a sticky property on the Operation singleton --
    ' callers that override it own restoring it.
    VCSIndex.Disabled = blnIndexWasDisabled
    If blnOperationOwned Then Operation.InteractionMode = eimPriorMode

    ' Build final JSON.
    dResult.Add "success", (dStats("failed") = 0 And dStats("errors") = 0)
    dResult.Add "fixtureFolder", strFixtureFolder
    dResult.Add "scratchFolder", strScratch
    dResult.Add "rebaseline", blnRebaseline
    dResult.Add "logPath", FSO.BuildPath(strFixtureFolder & "logs", _
        "ObjectRoundtrip_" & Log.OperationId & ".log")
    dResult.Add "stats", dStats
    dResult.Add "results", CollectionToJsonArray(colResults)

    If blnOperationOwned Then
        If dResult("success") Then
            Operation.Finish eorSuccess
        Else
            Operation.Finish eorFailed
        End If
    End If

    RunObjectRoundtripTests = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    ' Best-effort cleanup so we do not leave the index disabled, the
    ' operation hanging in eosRunning state, or the singleton stuck in
    ' silent mode if something blew up unexpectedly.
    On Error Resume Next
    UnloadScaffold
    VCSIndex.Disabled = blnIndexWasDisabled
    If blnOperationOwned Then
        Operation.InteractionMode = eimPriorMode
        Operation.Finish eorFailed
    End If
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    RunObjectRoundtripTests = ConvertToJson(dResult)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunQueryFixtures
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Enumerate every .sql fixture under strQueriesFolder (recursive)
'           : and run the round-trip on each one. Adds a result Dictionary to
'           : colResults for every fixture (including skipped/errored ones).
'---------------------------------------------------------------------------------------
'
Private Sub RunQueryFixtures(ByVal strQueriesFolder As String, ByVal strScratch As String, _
    ByVal blnRebaseline As Boolean, ByVal colResults As Collection)

    Dim colFiles As Collection
    Dim varFile As Variant
    Dim dFixtureResult As Dictionary
    Dim lngTotal As Long

    If Not FSO.FolderExists(strQueriesFolder) Then
        Log.Add T("No 'queries' folder under fixture root; skipping queries."), False
        Exit Sub
    End If

    Set colFiles = EnumerateSqlFixtures(strQueriesFolder)
    lngTotal = colFiles.Count

    If lngTotal = 0 Then
        Log.Add T("No query fixtures found under {0}", var0:=strQueriesFolder), False
        Exit Sub
    End If

    Log.Add T("Running {0} query fixture(s)...", var0:=CStr(lngTotal))

    For Each varFile In colFiles
        Set dFixtureResult = RunQueryRoundtrip(CStr(varFile), strScratch, blnRebaseline)
        colResults.Add dFixtureResult
        LogFixtureResult dFixtureResult
    Next varFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunQueryRoundtrip
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Execute the two-pass round-trip for a single .sql fixture and
'           : return a result dictionary.
'           :
'           : Pass 1: copy fixture to a sandbox name, import via clsDbQuery,
'           :         export back to scratch\pass1\, compare to fixture.
'           : Pass 2: export the same in-memory query to scratch\pass2\,
'           :         compare to pass1 (idempotency check).
'           :
'           : Fixture content is compared logically:
'           :  - .sql is compared byte-for-byte (no name embedded).
'           :  - .json is compared structurally, ignoring header.FileName which
'           :    legitimately differs because the in-memory query is sandboxed.
'           :
'           : The sandbox query is dropped in CleanUp regardless of pass/fail.
'---------------------------------------------------------------------------------------
'
Private Function RunQueryRoundtrip(ByVal strFixtureSql As String, ByVal strScratch As String, _
    ByVal blnRebaseline As Boolean) As Dictionary

    Dim dResult As Dictionary
    Dim strFixtureJson As String
    Dim strFixtureBase As String
    Dim strOriginalName As String
    Dim strSandboxName As String
    Dim strPass1Folder As String
    Dim strPass2Folder As String
    Dim strSandboxSqlIn As String
    Dim strSandboxJsonIn As String
    Dim strPass1Sql As String
    Dim strPass1Json As String
    Dim strPass2Sql As String
    Dim strPass2Json As String
    Dim cQuery As clsDbQuery
    Dim cComponent As IDbComponent
    Dim colChecks As Collection
    Dim blnSandboxImported As Boolean
    Dim lngErrCountBefore As Long
    Dim eelPriorErrorLevel As eErrorLevel

    Set dResult = New Dictionary
    Set colChecks = New Collection

    strOriginalName = FSO.GetBaseName(strFixtureSql)
    strFixtureBase = Left$(strFixtureSql, Len(strFixtureSql) - 4) ' strip ".sql"
    strFixtureJson = strFixtureBase & ".json"

    dResult.Add "fixture", strFixtureSql
    dResult.Add "name", strOriginalName
    dResult.Add "type", "query"
    dResult.Add "checks", colChecks

    LogUnhandledErrors
    On Error GoTo FixtureErrHandler

    ' --- Per-fixture isolation ----------------------------------------------
    ' modLoadSaveText.LoadComponentFromText returns False whenever
    ' Operation.ErrorLevel = eelCritical, and Operation.ErrorLevel only resets
    ' inside Operation.Begin -- which the harness calls once per *run*, not
    ' once per fixture. Without resetting it here, the first fixture that
    ' tickles a CRITICAL log entry (typically an outright import failure)
    ' poisons every subsequent LoadComponentFromText for the rest of the
    ' run, producing a cascade of false "Import logged 1 error(s)" failures.
    ' Cache the prior level so the run-level result still reflects whether
    ' anything went critical during the session.
    eelPriorErrorLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError

    ' Sanity: the .json companion is required.
    If Not FSO.FileExists(strFixtureJson) Then
        dResult("status") = "skip"
        dResult("reason") = "Missing companion .json: " & strFixtureJson
        Set RunQueryRoundtrip = dResult
        Exit Function
    End If

    ' Build sandbox names + paths.
    strSandboxName = TEST_PREFIX & strOriginalName & "_" & ShortHash(strFixtureSql & Now)
    dResult.Add "sandboxName", strSandboxName

    strPass1Folder = strScratch & "pass1" & PathSep
    strPass2Folder = strScratch & "pass2" & PathSep
    VerifyPath strPass1Folder
    VerifyPath strPass2Folder

    strSandboxSqlIn = strPass1Folder & strSandboxName & ".sql"
    strSandboxJsonIn = strPass1Folder & strSandboxName & ".json"
    strPass1Sql = strPass1Folder & strSandboxName & ".out.sql"
    strPass1Json = strPass1Folder & strSandboxName & ".out.json"
    strPass2Sql = strPass2Folder & strSandboxName & ".out.sql"
    strPass2Json = strPass2Folder & strSandboxName & ".out.json"

    ' Stage the fixture under its sandbox name. The .sql is copied as-is.
    ' The .json is also copied verbatim: clsDbQuery.ImportNewFormat reads
    ' query name from the filename (not from JSON), and the "Info" block
    ' inside the JSON is purely descriptive metadata that the import path
    ' ignores. Re-export will rewrite Info.Description from the in-memory
    ' query name (which is the sandbox name).
    FSO.CopyFile strFixtureSql, strSandboxSqlIn, True
    FSO.CopyFile strFixtureJson, strSandboxJsonIn, True

    ' Track Log error counts so we can detect import-time failures even if
    ' clsDbQuery swallows them with eelError but does not raise an exception.
    lngErrCountBefore = Log.ErrorCount

    ' --- Import (sandbox name) ---
    Set cQuery = New clsDbQuery
    Set cComponent = cQuery
    cComponent.Import strSandboxSqlIn
    blnSandboxImported = ObjectExists(acQuery, strSandboxName)

    If Not blnSandboxImported Then
        dResult("status") = "fail"
        dResult("reason") = "Import did not produce a query named '" & strSandboxName & "'"
        GoTo FixtureCleanUp
    End If

    ' Satisfy the IDbComponent.Export precondition. clsDbQuery.ImportNewFormat
    ' only sets m_Query on its success path -- a partial import that creates
    ' the qdef but errors out before the binding line would leave m_Query
    ' Nothing and crash the subsequent Export. Bind explicitly here, the same
    ' way IDbComponent_GetAllFromDB binds objects before normal exports.
    BindComponentAfterImport cComponent, acQuery, strSandboxName

    If Log.ErrorCount > lngErrCountBefore Then
        ' Import logged an eelError; treat as failure even though the qdef
        ' may have been created via the SQL View fallback path.
        AddCheck colChecks, "import", "fail", _
            "Import logged " & (Log.ErrorCount - lngErrCountBefore) & " error(s); see log."
    Else
        AddCheck colChecks, "import", "pass", vbNullString
    End If

    ' --- Pass 1 export (re-export the sandbox query) ---
    cComponent.Export strPass1Sql

    If Not FSO.FileExists(strPass1Sql) Then
        dResult("status") = "fail"
        dResult("reason") = "Pass 1 export did not produce: " & strPass1Sql
        GoTo FixtureCleanUp
    End If

    ' --- Pass 1 comparison (re-export vs canonical fixture) ---
    ComparePass1ToFixture strFixtureSql, strFixtureJson, strPass1Sql, strPass1Json, _
        strOriginalName, strSandboxName, blnRebaseline, colChecks

    ' --- Pass 2 export (idempotency) ---
    cComponent.Export strPass2Sql

    If Not FSO.FileExists(strPass2Sql) Then
        AddCheck colChecks, "pass2_export", "fail", _
            "Pass 2 export did not produce: " & strPass2Sql
    Else
        ComparePass2Idempotency strPass1Sql, strPass1Json, strPass2Sql, strPass2Json, colChecks
    End If

    ' Roll up final status from the recorded checks.
    dResult("status") = RollUpStatus(colChecks)
    If dResult("status") = "fail" Then dResult("reason") = "One or more comparisons failed."

FixtureCleanUp:
    On Error Resume Next
    If blnSandboxImported Then
        DeleteSandboxObject acQuery, strSandboxName
        DBEngine.Idle dbRefreshCache
    End If
    Set cComponent = Nothing
    Set cQuery = Nothing

    ' Restore the prior error level if this fixture didn't itself escalate
    ' it. Preserves CRITICAL across the run if some fixture went critical
    ' before us, while still letting this fixture's own escalation propagate.
    If Operation.ErrorLevel < eelPriorErrorLevel Then
        Operation.ErrorLevel = eelPriorErrorLevel
    End If

    Set RunQueryRoundtrip = dResult
    Exit Function

FixtureErrHandler:
    dResult("status") = "error"
    dResult("reason") = "Unhandled error: " & Err.Number & " " & Err.Description
    AddCheck colChecks, "exception", "error", Err.Description
    Resume FixtureCleanUp

End Function


'---------------------------------------------------------------------------------------
' Procedure : BindComponentAfterImport
' Author    : VCS contributors
' Date      : 4/25/2026
' Purpose   : Satisfy the IDbComponent.Export precondition by Set'ing DbObject
'           : to the freshly-imported AccessObject. Test code imports under a
'           : sandbox name and calls this helper before invoking Export. This
'           : mirrors what IDbComponent_GetAllFromDB does for normal exports
'           : (enumerate the container, Set DbObject on each instance) -- the
'           : harness has to do it explicitly because it bypasses that
'           : enumeration path.
'           :
'           : Extend the Select Case as new component types gain round-trip
'           : helpers. See IDbComponent.cls (Export procedure header) for the
'           : full contract.
'---------------------------------------------------------------------------------------
'
Private Sub BindComponentAfterImport(ByVal cComponent As IDbComponent, _
    ByVal intType As AcObjectType, ByVal strName As String)

    Select Case intType
        Case acQuery:  Set cComponent.DbObject = CurrentData.AllQueries(strName)
        ' Future round-trip helpers (uncomment as each type is added):
        ' Case acForm:   Set cComponent.DbObject = CurrentProject.AllForms(strName)
        ' Case acReport: Set cComponent.DbObject = CurrentProject.AllReports(strName)
        ' Case acModule: Set cComponent.DbObject = CurrentProject.AllModules(strName)
        ' Case acMacro:  Set cComponent.DbObject = CurrentProject.AllMacros(strName)
        ' Case acTable:  Set cComponent.DbObject = CurrentData.AllTables(strName)
        Case Else
            Log.Error eelError, _
                "BindComponentAfterImport: unsupported AcObjectType " & intType, _
                ModuleName & ".BindComponentAfterImport"
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ComparePass1ToFixture
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Compare the Pass 1 export against the canonical fixture.
'           :  - .sql is compared byte-for-byte.
'           :  - .json is compared structurally with the entire "Info" block
'           :    ignored (Info.Description carries the query name, which
'           :    legitimately differs because the in-memory query was renamed
'           :    to its sandbox name; the rest of Info is fixed metadata).
'           : When blnRebaseline is True, mismatches overwrite the fixture
'           : with the Pass 1 output (after rewriting Info.Description back
'           : to the original name).
'---------------------------------------------------------------------------------------
'
Private Sub ComparePass1ToFixture(ByVal strFixtureSql As String, ByVal strFixtureJson As String, _
    ByVal strPass1Sql As String, ByVal strPass1Json As String, _
    ByVal strOriginalName As String, ByVal strSandboxName As String, _
    ByVal blnRebaseline As Boolean, ByVal colChecks As Collection)

    Dim strFixtureSqlText As String
    Dim strPass1SqlText As String
    Dim strDiff As String

    ' --- SQL compare (byte-for-byte via hash) ---
    If GetFileHash(strFixtureSql) = GetFileHash(strPass1Sql) Then
        AddCheck colChecks, "sql_vs_fixture", "pass", vbNullString
    Else
        strFixtureSqlText = ReadFile(strFixtureSql)
        strPass1SqlText = ReadFile(strPass1Sql)
        strDiff = MakeUnifiedDiff(strFixtureSqlText, strPass1SqlText, _
            "fixture/" & strOriginalName & ".sql", _
            "pass1/" & strSandboxName & ".out.sql")
        AddCheckWithDiff colChecks, "sql_vs_fixture", "fail", _
            "SQL differs from fixture", strDiff
        If blnRebaseline Then
            WriteFile strPass1SqlText, strFixtureSql
            Log.Add T("REBASELINE: overwrote {0}", var0:=strFixtureSql), False
        End If
    End If

    ' --- JSON compare (structural, ignore header.FileName) ---
    CompareJsonFiles strFixtureJson, strPass1Json, strOriginalName, strSandboxName, _
        "json_vs_fixture", "pass1/" & strSandboxName & ".out.json", _
        blnRebaseline, strFixtureJson, colChecks

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ComparePass2Idempotency
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Compare Pass 2 against Pass 1. Both passes export the same
'           : in-memory query under the same sandbox name, so even
'           : header.FileName must match -- this is a strict idempotency check.
'---------------------------------------------------------------------------------------
'
Private Sub ComparePass2Idempotency(ByVal strPass1Sql As String, ByVal strPass1Json As String, _
    ByVal strPass2Sql As String, ByVal strPass2Json As String, ByVal colChecks As Collection)

    Dim strDiff As String

    If GetFileHash(strPass1Sql) = GetFileHash(strPass2Sql) Then
        AddCheck colChecks, "sql_pass2_idempotent", "pass", vbNullString
    Else
        strDiff = MakeUnifiedDiff(ReadFile(strPass1Sql), ReadFile(strPass2Sql), _
            "pass1.sql", "pass2.sql")
        AddCheckWithDiff colChecks, "sql_pass2_idempotent", "fail", _
            "Pass 2 SQL differs from Pass 1 (export is not idempotent)", strDiff
    End If

    If FSO.FileExists(strPass1Json) And FSO.FileExists(strPass2Json) Then
        If GetFileHash(strPass1Json) = GetFileHash(strPass2Json) Then
            AddCheck colChecks, "json_pass2_idempotent", "pass", vbNullString
        Else
            strDiff = MakeUnifiedDiff(ReadFile(strPass1Json), ReadFile(strPass2Json), _
                "pass1.json", "pass2.json")
            AddCheckWithDiff colChecks, "json_pass2_idempotent", "fail", _
                "Pass 2 JSON differs from Pass 1 (export is not idempotent)", strDiff
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompareJsonFiles
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Structural comparison of two .json files, ignoring the entire
'           : Info block (which contains Description == query name and
'           : legitimately differs between fixture and Pass 1 output because
'           : the imported query is sandboxed).
'           :
'           : Strategy: parse both, drop the Info section from each,
'           : re-serialize with ConvertToJson, hash and compare. On mismatch,
'           : fall back to a line-level unified diff over the re-serialized
'           : text so the report is human-readable.
'---------------------------------------------------------------------------------------
'
Private Sub CompareJsonFiles(ByVal strExpectedFile As String, ByVal strActualFile As String, _
    ByVal strOriginalName As String, ByVal strSandboxName As String, _
    ByVal strCheckId As String, ByVal strActualLabel As String, _
    ByVal blnRebaseline As Boolean, ByVal strRebaselineTarget As String, _
    ByVal colChecks As Collection)

    Dim dExpected As Dictionary
    Dim dActual As Dictionary
    Dim strNormExpected As String
    Dim strNormActual As String
    Dim strDiff As String

    If Not FSO.FileExists(strActualFile) Then
        AddCheck colChecks, strCheckId, "fail", _
            "Expected .json output file not produced: " & strActualFile
        Exit Sub
    End If

    Set dExpected = ReadJsonFile(strExpectedFile)
    Set dActual = ReadJsonFile(strActualFile)

    If dExpected Is Nothing Then
        AddCheck colChecks, strCheckId, "fail", _
            "Could not parse expected .json: " & strExpectedFile
        Exit Sub
    End If
    If dActual Is Nothing Then
        AddCheck colChecks, strCheckId, "fail", _
            "Could not parse actual .json: " & strActualFile
        Exit Sub
    End If

    ' Normalize both: strip the Info section before hashing. Info contains
    ' name-derived metadata (Description == query name) and is purely
    ' descriptive -- the import path does not consume it.
    StripInfoSection dExpected
    StripInfoSection dActual

    strNormExpected = ConvertToJson(dExpected, JSON_WHITESPACE)
    strNormActual = ConvertToJson(dActual, JSON_WHITESPACE)

    If GetStringHash(strNormExpected) = GetStringHash(strNormActual) Then
        AddCheck colChecks, strCheckId, "pass", vbNullString
    Else
        strDiff = MakeUnifiedDiff(strNormExpected, strNormActual, _
            "fixture/" & strOriginalName & ".json (normalized)", _
            strActualLabel & " (normalized)")
        AddCheckWithDiff colChecks, strCheckId, "fail", _
            "JSON differs from fixture (excluding Info section)", strDiff
        If blnRebaseline Then
            ' Re-serialize the actual JSON with Info.Description rewritten back
            ' to the canonical (original) name, so the rebaselined fixture is
            ' name-agnostic and matches what the next non-rebaseline run will
            ' compare against.
            RebaselineFixtureJson strActualFile, strRebaselineTarget, strOriginalName
            Log.Add T("REBASELINE: overwrote {0}", var0:=strRebaselineTarget), False
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StripInfoSection
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Remove the entire "Info" block from a parsed .json document, in
'           : place. The Info block (Class + Description) is purely descriptive
'           : metadata for human readability and is not consumed during import,
'           : so dropping it makes round-trip JSON comparisons name-agnostic.
'---------------------------------------------------------------------------------------
'
Private Sub StripInfoSection(ByVal dJson As Dictionary)
    If dJson Is Nothing Then Exit Sub
    If dJson.Exists("Info") Then dJson.Remove "Info"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RebaselineFixtureJson
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Read the actual sandbox-export .json, rewrite Info.Description
'           : back to the original (canonical) query name, and write the result
'           : over the fixture file. Used by the rebaseline path so updated
'           : fixtures embed the canonical name rather than a sandbox name.
'---------------------------------------------------------------------------------------
'
Private Sub RebaselineFixtureJson(ByVal strActualFile As String, ByVal strFixturePath As String, _
    ByVal strOriginalName As String)

    Dim dJson As Dictionary
    Dim dInfo As Dictionary

    Set dJson = ReadJsonFile(strActualFile)
    If dJson Is Nothing Then
        FSO.CopyFile strActualFile, strFixturePath, True
        Exit Sub
    End If

    If dJson.Exists("Info") Then
        If TypeOf dJson("Info") Is Dictionary Then
            Set dInfo = dJson("Info")
            If dInfo.Exists("Description") Then dInfo("Description") = strOriginalName
        End If
    End If

    WriteFile ConvertToJson(dJson, JSON_WHITESPACE), strFixturePath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnumerateSqlFixtures
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Recursively enumerate .sql files under strRoot (skipping any
'           : "_scaffold" subfolders, which are handled separately). Returns a
'           : Collection of full file paths sorted by their relative path.
'---------------------------------------------------------------------------------------
'
Private Function EnumerateSqlFixtures(ByVal strRoot As String) As Collection
    Dim col As Collection
    Set col = New Collection
    EnumerateSqlFixturesRecurse strRoot, col
    Set EnumerateSqlFixtures = col
End Function

Private Sub EnumerateSqlFixturesRecurse(ByVal strFolder As String, ByVal col As Collection)
    Dim oFolder As Object
    Dim oFile As Object
    Dim oSub As Object

    If Not FSO.FolderExists(strFolder) Then Exit Sub
    Set oFolder = FSO.GetFolder(strFolder)

    For Each oFile In oFolder.Files
        If LCase$(FSO.GetExtensionName(oFile.Name)) = "sql" Then
            col.Add oFile.Path
        End If
    Next oFile

    For Each oSub In oFolder.SubFolders
        ' Skip scaffold (handled separately) and dotfile / underscore-prefixed
        ' housekeeping folders.
        If LCase$(oSub.Name) <> "_scaffold" And Left$(oSub.Name, 1) <> "." Then
            EnumerateSqlFixturesRecurse oSub.Path & PathSep, col
        End If
    Next oSub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadScaffold / UnloadScaffold
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Pre-load supporting objects (currently queries only) from the
'           : "_scaffold" folder of the fixture root and register them for
'           : end-of-session cleanup. Scaffold objects use their original names
'           : (no prefix) because fixtures may reference them by name in JOINs
'           : / subqueries.
'---------------------------------------------------------------------------------------
'
Private Sub LoadScaffold(ByVal strScaffoldFolder As String)

    Dim oFolder As Object
    Dim oFile As Object
    Dim cQuery As clsDbQuery
    Dim cComponent As IDbComponent
    Dim strName As String
    Dim lngLoaded As Long

    If Not FSO.FolderExists(strScaffoldFolder) Then Exit Sub

    Set oFolder = FSO.GetFolder(strScaffoldFolder)
    For Each oFile In oFolder.Files
        If LCase$(FSO.GetExtensionName(oFile.Name)) = "sql" Then
            strName = FSO.GetBaseName(oFile.Name)

            ' If a same-named query already exists, warn and skip to avoid
            ' clobbering the user's database with scaffold content.
            If ObjectExists(acQuery, strName) Then
                Log.Error eelWarning, T( _
                    "Scaffold query '{0}' already exists in the database; skipping.", _
                    var0:=strName), ModuleName & ".LoadScaffold"
            Else
                Set cQuery = New clsDbQuery
                Set cComponent = cQuery
                cComponent.Import oFile.Path
                If ObjectExists(acQuery, strName) Then
                    m_colScaffoldQueries.Add strName
                    lngLoaded = lngLoaded + 1
                End If
                Set cComponent = Nothing
                Set cQuery = Nothing
            End If
        End If
    Next oFile

    If lngLoaded > 0 Then Log.Add T("Loaded {0} scaffold object(s).", var0:=CStr(lngLoaded)), False

End Sub

Private Sub UnloadScaffold()
    Dim varName As Variant
    If m_colScaffoldQueries Is Nothing Then Exit Sub
    For Each varName In m_colScaffoldQueries
        DeleteSandboxObject acQuery, CStr(varName)
    Next varName
    DBEngine.Idle dbRefreshCache
    Set m_colScaffoldQueries = New Collection
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CleanupStaleObjects
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Drop any leftover sandbox objects (vcs_test_* / vcs_scaffold_*)
'           : at the start of a run. These can accumulate if a previous run
'           : crashed before its own cleanup.
'           : Currently scans queries only (the only object type supported
'           : in v1); extend per-type as new components are added.
'---------------------------------------------------------------------------------------
'
Private Sub CleanupStaleObjects()
    Dim qdf As DAO.QueryDef
    Dim colVictims As Collection
    Dim varName As Variant

    Set colVictims = New Collection

    For Each qdf In CurrentDb.QueryDefs
        If Left$(qdf.Name, Len(TEST_PREFIX)) = TEST_PREFIX _
            Or Left$(qdf.Name, Len(SCAFFOLD_PREFIX)) = SCAFFOLD_PREFIX Then
            colVictims.Add qdf.Name
        End If
    Next qdf

    If colVictims.Count = 0 Then Exit Sub
    Log.Add T("Cleaning up {0} stale test object(s) from a prior run.", _
        var0:=CStr(colVictims.Count)), False

    For Each varName In colVictims
        DeleteSandboxObject acQuery, CStr(varName)
    Next varName
    DBEngine.Idle dbRefreshCache
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DeleteSandboxObject
' Author    : VCS contributors
' Date      : 4/25/2026
' Purpose   : Drop a sandboxed (or scaffold) object via the lowest-level DAO/VBE
'           : API for its type. Bypasses modDatabase.DeleteObjectIfExists, which
'           : guards against an add-in/CurrentDb name collision by renaming
'           : the victim to "<name>_DELETE_<7-char-hash>" before deleting --
'           : that suffix overflows Access's 64-char object name limit when the
'           : sandbox name is already long (e.g. vcs_test_<35-char fixture>_<7>
'           : = 52 chars + 15-char delete suffix = 67), causing error 3125 and
'           : leaving the test object behind permanently.
'           :
'           : The collision protection is unnecessary for harness-managed names
'           : because TEST_PREFIX / SCAFFOLD_PREFIX are unique to this module
'           : and cannot match anything in the add-in's own object set. Direct
'           : DAO deletion succeeds regardless of name length.
'           :
'           : Returns True if the object no longer exists when the call returns
'           : (including the case where it never existed).
'---------------------------------------------------------------------------------------
'
Private Function DeleteSandboxObject(ByVal intType As AcObjectType, _
    ByVal strName As String) As Boolean

    LogUnhandledErrors
    On Error Resume Next

    Select Case intType
        Case acQuery
            CurrentDb.QueryDefs.Delete strName
        ' Future component types (uncomment as round-trip helpers are added):
        ' Case acTable:  CurrentDb.TableDefs.Delete strName
        ' Case acModule: CurrentVBProject.VBComponents.Remove _
        '                  CurrentVBProject.VBComponents(strName)
        ' Case acForm, acReport, acMacro:
        '     DoCmd.DeleteObject intType, strName    ' no add-in collision risk
        '                                            ' for vcs_test_/vcs_scaffold_
        Case Else
            Log.Error eelError, _
                "DeleteSandboxObject: unsupported AcObjectType " & intType, _
                ModuleName & ".DeleteSandboxObject"
            DeleteSandboxObject = False
            Exit Function
    End Select

    ' Swallow "object not found" -- callers treat it as a successful no-op.
    If Err.Number <> 0 Then Err.Clear
    DeleteSandboxObject = Not ObjectExists(intType, strName)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ProvisionScratchFolder
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Create a fresh per-run scratch folder under the fixture root's
'           : "scratch\" subfolder, named after the current operation ID so
'           : multiple consecutive runs do not clobber each other's outputs.
'---------------------------------------------------------------------------------------
'
Private Function ProvisionScratchFolder(ByVal strFixtureFolder As String) As String
    Dim strFolder As String
    strFolder = strFixtureFolder & "scratch" & PathSep & Log.OperationId & PathSep
    VerifyPath strFolder
    ProvisionScratchFolder = strFolder
End Function


'---------------------------------------------------------------------------------------
' Procedure : MakeUnifiedDiff
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Produce a small unified-diff-style report between two text blobs.
'           : Not a full Myers diff -- this is a line-by-line scanner that
'           : groups consecutive non-matching lines into hunks with
'           : DIFF_CONTEXT_LINES of context. Adequate for highlighting where
'           : a round-trip diverged; for deeper inspection developers should
'           : use the preserved scratch files with their own diff tool.
'---------------------------------------------------------------------------------------
'
Private Function MakeUnifiedDiff(ByVal strA As String, ByVal strB As String, _
    ByVal strLabelA As String, ByVal strLabelB As String) As String

    Dim aA() As String
    Dim aB() As String
    Dim cc As clsConcat
    Dim lngI As Long
    Dim lngMax As Long
    Dim blnInHunk As Boolean
    Dim blnDiffer As Boolean

    aA = Split(NormalizeLineEndings(strA), vbLf)
    aB = Split(NormalizeLineEndings(strB), vbLf)
    lngMax = IIf(UBound(aA) > UBound(aB), UBound(aA), UBound(aB))

    Set cc = New clsConcat
    cc.AppendOnAdd = vbCrLf
    cc.Add "--- " & strLabelA
    cc.Add "+++ " & strLabelB

    For lngI = 0 To lngMax
        blnDiffer = (SafeLine(aA, lngI) <> SafeLine(aB, lngI))
        If blnDiffer Then
            If Not blnInHunk Then
                cc.Add "@@ line " & (lngI + 1) & " @@"
                EmitContext cc, aA, lngI - DIFF_CONTEXT_LINES, lngI - 1, " "
                blnInHunk = True
            End If
            If lngI <= UBound(aA) Then cc.Add "-" & SafeLine(aA, lngI)
            If lngI <= UBound(aB) Then cc.Add "+" & SafeLine(aB, lngI)
        Else
            If blnInHunk Then
                EmitContext cc, aA, lngI, lngI + DIFF_CONTEXT_LINES - 1, " "
                blnInHunk = False
            End If
        End If
    Next lngI

    MakeUnifiedDiff = cc.GetStr
End Function

Private Sub EmitContext(ByVal cc As clsConcat, ByRef arr() As String, _
    ByVal lngFrom As Long, ByVal lngTo As Long, ByVal strPrefix As String)
    Dim i As Long
    If lngFrom < 0 Then lngFrom = 0
    If lngTo > UBound(arr) Then lngTo = UBound(arr)
    For i = lngFrom To lngTo
        cc.Add strPrefix & arr(i)
    Next i
End Sub

Private Function SafeLine(ByRef arr() As String, ByVal idx As Long) As String
    If idx < 0 Or idx > UBound(arr) Then
        SafeLine = vbNullString
    Else
        SafeLine = arr(idx)
    End If
End Function

Private Function NormalizeLineEndings(ByVal s As String) As String
    NormalizeLineEndings = Replace$(Replace$(s, vbCrLf, vbLf), vbCr, vbLf)
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddCheck / AddCheckWithDiff
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Append a check result dictionary to colChecks.
'           : "diff" is omitted when there is none; "diffTruncated" is set when
'           : the diff exceeds DIFF_MAX_CHARS_FOR_JSON (the full diff is still
'           : written to the log file).
'---------------------------------------------------------------------------------------
'
Private Sub AddCheck(ByVal colChecks As Collection, ByVal strId As String, _
    ByVal strStatus As String, ByVal strMessage As String)
    AddCheckWithDiff colChecks, strId, strStatus, strMessage, vbNullString
End Sub

Private Sub AddCheckWithDiff(ByVal colChecks As Collection, ByVal strId As String, _
    ByVal strStatus As String, ByVal strMessage As String, ByVal strDiff As String)

    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "id", strId
    d.Add "status", strStatus
    If Len(strMessage) > 0 Then d.Add "message", strMessage
    If Len(strDiff) > 0 Then
        If Len(strDiff) > DIFF_MAX_CHARS_FOR_JSON Then
            d.Add "diff", Left$(strDiff, DIFF_MAX_CHARS_FOR_JSON)
            d.Add "diffTruncated", True
        Else
            d.Add "diff", strDiff
            d.Add "diffTruncated", False
        End If
    End If
    colChecks.Add d

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RollUpStatus
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Reduce a collection of check dictionaries to a single status:
'           : "error" > "fail" > "skip" > "pass" (left wins).
'---------------------------------------------------------------------------------------
'
Private Function RollUpStatus(ByVal colChecks As Collection) As String
    Dim varCheck As Variant
    Dim strStatus As String
    Dim blnHasFail As Boolean
    Dim blnHasSkip As Boolean

    If colChecks.Count = 0 Then
        RollUpStatus = "skip"
        Exit Function
    End If

    For Each varCheck In colChecks
        strStatus = CStr(varCheck("status"))
        Select Case strStatus
            Case "error"
                RollUpStatus = "error"
                Exit Function
            Case "fail"
                blnHasFail = True
            Case "skip"
                blnHasSkip = True
        End Select
    Next varCheck

    If blnHasFail Then
        RollUpStatus = "fail"
    ElseIf blnHasSkip Then
        RollUpStatus = "skip"
    Else
        RollUpStatus = "pass"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : LogFixtureResult
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Write a single-line summary for a fixture to the console log,
'           : plus full diff text into the log file when there were failures.
'---------------------------------------------------------------------------------------
'
Private Sub LogFixtureResult(ByVal dResult As Dictionary)

    Dim strStatus As String
    Dim strColor As String
    Dim blnBold As Boolean
    Dim varCheck As Variant
    Dim dCheck As Dictionary
    Dim strBadge As String

    strStatus = CStr(dResult("status"))
    Select Case strStatus
        Case "pass":  strBadge = "[PASS]":  strColor = "green"
        Case "fail":  strBadge = "[FAIL]":  strColor = "red":     blnBold = True
        Case "error": strBadge = "[ERROR]": strColor = "red":     blnBold = True
        Case "skip":  strBadge = "[SKIP]":  strColor = "gray"
        Case Else:    strBadge = "[" & UCase$(strStatus) & "]"
    End Select

    Log.Add strBadge & " " & CStr(dResult("name")), True, , strColor, blnBold

    If strStatus = "skip" And dResult.Exists("reason") Then
        Log.Add "       reason: " & CStr(dResult("reason")), False
    End If

    ' Emit per-check details (and full diffs) into the log file, but only the
    ' headlines onto the console form.
    If strStatus = "fail" Or strStatus = "error" Then
        For Each varCheck In dResult("checks")
            Set dCheck = varCheck
            If CStr(dCheck("status")) <> "pass" Then
                Log.Add "       " & CStr(dCheck("id")) & ": " & _
                    CStr(dCheck("status")) & _
                    IIf(dCheck.Exists("message"), " - " & CStr(dCheck("message")), ""), _
                    True, , strColor
                If dCheck.Exists("diff") Then
                    ' File-only (blnPrint:=False); console stays readable.
                    Log.Add CStr(dCheck("diff")), False
                End If
            End If
        Next varCheck
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildStatsDict
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Tally pass/fail/skip/error counts across all fixture results.
'---------------------------------------------------------------------------------------
'
Private Sub BuildStatsDict(ByVal colResults As Collection, ByVal dStats As Dictionary)
    Dim varResult As Variant
    Dim lngTotal As Long, lngPassed As Long, lngFailed As Long
    Dim lngSkipped As Long, lngErrored As Long

    For Each varResult In colResults
        lngTotal = lngTotal + 1
        Select Case CStr(varResult("status"))
            Case "pass":  lngPassed = lngPassed + 1
            Case "fail":  lngFailed = lngFailed + 1
            Case "skip":  lngSkipped = lngSkipped + 1
            Case "error": lngErrored = lngErrored + 1
        End Select
    Next varResult

    dStats.Add "total", lngTotal
    dStats.Add "passed", lngPassed
    dStats.Add "failed", lngFailed
    dStats.Add "skipped", lngSkipped
    dStats.Add "errors", lngErrored
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CollectionToJsonArray
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Convert a Collection into a Collection that ConvertToJson will
'           : serialize as a JSON array (it already does -- this helper exists
'           : so the call site reads as "this is the array I want returned").
'           : Also flattens nested check Collections into Collections.
'---------------------------------------------------------------------------------------
'
Private Function CollectionToJsonArray(ByVal col As Collection) As Collection
    ' modJsonConverter.ConvertToJson serializes Collection -> JSON array, so
    ' returning the collection as-is is sufficient. Kept as a named helper to
    ' make intent obvious at the call site.
    Set CollectionToJsonArray = col
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDefaultFixtureRoot
' Author    : VCS contributors
' Date      : 4/24/2026
' Purpose   : Resolve the path to the harness's own fixture corpus shipped in
'           : the add-in repo. Assumes the running .accda lives at the repo
'           : root: <repo>\Version Control.accda. If the harness is being run
'           : from an installed add-in (where this folder will not exist), the
'           : caller is expected to pass an explicit fixture folder instead.
'---------------------------------------------------------------------------------------
'
Private Function GetDefaultFixtureRoot() As String
    Dim strRoot As String
    strRoot = CodeProject.Path
    If Right$(strRoot, 1) <> PathSep Then strRoot = strRoot & PathSep
    GetDefaultFixtureRoot = strRoot & FIXTURE_RELATIVE_PATH & PathSep
End Function

