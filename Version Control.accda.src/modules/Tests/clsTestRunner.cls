VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestRunner
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Test runner engine for the TestAssert framework. Discovers test subs in
'           : CurrentVBProject, drives execution, records assertion results, and streams
'           : progress to frmVCSMain (console fallback) and/or frmVCSTestRunner (web UI).
'           :
'           : This class is a singleton managed through modObjects.TestRunner.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests")

Private Const ModuleName As String = "clsTestRunner"
Private Const GlobalTestSetupProc As String = "GlobalTestSetup"
Private Const GlobalTestTeardownProc As String = "GlobalTestTeardown"

' Runner state machine
Public Enum eTestRunnerState
    etrsIdle = 0
    etrsScanning = 1
    etrsReady = 2
    etrsRunning = 3
    etrsComplete = 4
    etrsCancelled = 5
End Enum

' Test result status
Public Enum eTestStatus
    etsPending = 0
    etsPassed = 1
    etsFailed = 2
    etsErrored = 3
    etsEmpty = 4
End Enum

' Private UDT for internal state
Private Type udtTestRunner
    State As eTestRunnerState
    Tests As Dictionary             ' Key: "ModuleName.ProcName", Value: Dictionary (test record)
    SuiteMap As Dictionary          ' Key: folder path, Value: Collection of test keys
    CurrentTestKey As String
    AssertionCounter As Long
    AssertionResults As Collection  ' Collection of assertion result Dictionaries for current test
    CancelRequested As Boolean
    RunStartTime As Currency
    PassedCount As Long
    FailedCount As Long
    ErroredCount As Long
    EmptyCount As Long
    TotalAssertions As Long
    PassedAssertions As Long
    FailedAssertions As Long
    ScannedAssertionCount As Long
    LastRunDurationMs As Long
    LastRunKeys As Collection
End Type
Private this As udtTestRunner


'---------------------------------------------------------------------------------------
' Procedure : State
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Current state of the test runner.
'---------------------------------------------------------------------------------------
'
Public Property Get State() As eTestRunnerState
    State = this.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : Tests
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Dictionary of discovered test records (populated after scan).
'---------------------------------------------------------------------------------------
'
Public Property Get Tests() As Dictionary
    Set Tests = this.Tests
End Property


'---------------------------------------------------------------------------------------
' Procedure : SuiteMap
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Dictionary mapping folder/suite names to collections of test keys.
'---------------------------------------------------------------------------------------
'
Public Property Get SuiteMap() As Dictionary
    Set SuiteMap = this.SuiteMap
End Property


'---------------------------------------------------------------------------------------
' Procedure : PassedCount / FailedCount / ErroredCount / EmptyCount
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Summary counts from the most recent run.
'---------------------------------------------------------------------------------------
'
Public Property Get PassedCount() As Long
    PassedCount = this.PassedCount
End Property
Public Property Get FailedCount() As Long
    FailedCount = this.FailedCount
End Property
Public Property Get ErroredCount() As Long
    ErroredCount = this.ErroredCount
End Property
Public Property Get EmptyCount() As Long
    EmptyCount = this.EmptyCount
End Property
Public Property Get ScannedAssertionCount() As Long
    ScannedAssertionCount = this.ScannedAssertionCount
End Property


'---------------------------------------------------------------------------------------
' Procedure : LastRunDurationMs
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Duration of the most recent RunSelected/RunAll execution in milliseconds.
'---------------------------------------------------------------------------------------
'
Public Property Get LastRunDurationMs() As Long
    LastRunDurationMs = this.LastRunDurationMs
End Property


'---------------------------------------------------------------------------------------
' Procedure : LastRunKeys
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Test keys executed in the most recent RunSelected/RunAll invocation.
'---------------------------------------------------------------------------------------
'
Public Property Get LastRunKeys() As Collection
    Set LastRunKeys = this.LastRunKeys
End Property


'---------------------------------------------------------------------------------------
' Procedure : Scan
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Discover test modules and test subs in CurrentVBProject.
'           : Populates the Tests dictionary and SuiteMap.
'---------------------------------------------------------------------------------------
'
Public Sub Scan()

    this.State = etrsScanning
    ScanDiscover
    this.State = etrsReady

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanMergingPriorResults
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Rediscover tests without wiping pass/fail state for unchanged Module.Proc
'           : keys. Used when reopening the warm web runner or refreshing the tree.
'---------------------------------------------------------------------------------------
'
Public Sub ScanMergingPriorResults()

    Dim dPrior As Dictionary

    this.State = etrsScanning
    Set dPrior = SnapshotTestsByModuleProc()
    ScanDiscover
    MergePriorResultsIntoDiscovered dPrior
    RebuildCountsFromTests
    this.State = etrsReady

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunAll
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Run every discovered test.
'---------------------------------------------------------------------------------------
'
Public Sub RunAll()

    ' Scan when there are no discovered tests. Covers a page-initiated (bridge) RunAll
    ' after the singleton was reset (empty Tests) -- previously this exited silently.
    If this.Tests Is Nothing Then Scan
    If this.Tests.Count = 0 Then Scan
    modTestRunnerDiag.Diag "runall", "tests=" & this.Tests.Count
    If this.Tests.Count = 0 Then
        this.State = etrsComplete
        Exit Sub
    End If

    Dim colKeys As New Collection
    Dim varKey As Variant
    For Each varKey In this.Tests.Keys
        colKeys.Add CStr(varKey)
    Next varKey
    RunSelected colKeys

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunSelected
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Run a specific set of tests identified by their keys.
'---------------------------------------------------------------------------------------
'
Public Sub RunSelected(colTestKeys As Collection)

    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strCmd As String
    Dim lngErrNum As Long
    Dim strErrDesc As String
    Dim strTestKey As String
    Dim strProcName As String
    Dim strCurrentModule As String
    Dim curStart As Currency
    Dim intErrorTrapping As Integer
    Dim blnBreakOnError As Boolean
    Dim obj As Object
    Dim lngCurrent As Long
    Dim lngTotal As Long
    Dim lngErrAtStart As Long
    Dim lngErrJournalAtStart As Long
    Dim eelAtStart As eErrorLevel

    ' Re-entrancy guard: a run must never nest. The web runner streams progress with
    ' DoEvents between tests, which can dispatch a queued bridge call (RunAll/etc.).
    ' A nested run would call Scan/ResetCounts and repoint this.Tests, corrupting the
    ' outer loop (stale keys -> Scripting.Dictionary auto-vivifies Empty entries ->
    ' "Object required" in the summary). Ignore any invocation while already running.
    If this.State = etrsRunning Then
        modTestRunnerDiag.Diag "run.skip", "re-entrant (already running)"
        Exit Sub
    End If

    ' Ensure tests are discovered. A targeted rerun (per-test/suite button or "Rerun
    ' failed") can arrive long after the initial scan, by which point a VBA state reset
    ' may have cleared the singleton's Tests. Unlike RunAll, RunSelected does not scan,
    ' so every supplied key would miss this.Tests and the run would stream nothing
    ' (onRunStart -> onRunComplete with no tests, leaving the rows blank). Self-heal by
    ' scanning first; the deterministic keys (Module.Proc) still match the UI's keys.
    If this.Tests Is Nothing Then Scan
    If this.Tests.Count = 0 Then Scan

    modTestRunnerDiag.Diag "run.begin", "tests=" & colTestKeys.Count & " known=" & this.Tests.Count
    this.State = etrsRunning
    this.CancelRequested = False
    Set this.LastRunKeys = New Collection
    Dim lngKeyIdx As Long
    For lngKeyIdx = 1 To colTestKeys.Count
        this.LastRunKeys.Add CStr(colTestKeys(lngKeyIdx))
    Next lngKeyIdx
    ResetCounts
    Perf.OperationStart "Run Tests"
    this.RunStartTime = Perf.MicroTimer

    ' Suppress debugger breaks so tests that intentionally raise errors don't halt
    ' in the debugger. EXCEPTION: when the developer is debugging the add-in itself
    ' (Options.BreakOnError enabled), leave the current VBE Error Trapping mode
    ' (e.g. "Break in Class Modules") and BreakOnError intact so runtime errors in
    ' the runner/bridge break at their actual source line instead of being masked
    ' by "Break on Unhandled Errors".
    intErrorTrapping = Application.GetOption("Error Trapping")
    blnBreakOnError = Options.BreakOnError
    If Not blnBreakOnError Then
        Application.SetOption "Error Trapping", 2  ' Break on Unhandled Errors
        Options.BreakOnError = False
    End If

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    ' Compile project before running tests
    On Error Resume Next
    DoCmd.RunCommand acCmdCompileAllModules
    DoEvents
    On Error GoTo 0
    If modTestRunnerUI.WebRunnerActive() Then modTestRunnerUI.RefocusWebRunner
    If Not Application.IsCompiled Then
        modTestRunnerDiag.Diag "run.skip", "not compiled"
        Log.Add T("Project has compile errors. Please resolve before running tests."), , , "red"
        this.State = etrsCancelled
        GoTo CleanUp
    End If

    ' Set up progress bar (absent when hosted in the web runner, where frmVCSMain
    ' and its progress control are never opened).
    lngTotal = colTestKeys.Count
    Log.KeepProgressVisible = True
    If Not Log.ProgressBar Is Nothing Then
        Log.ProgressBar.Reset
        Log.ProgressBar.Max = lngTotal
    End If

    modTestRunnerUI.StreamRunStart colTestKeys

    ' Reset status on all selected tests
    For Each varKey In colTestKeys
        If this.Tests.Exists(CStr(varKey)) Then
            Set dTest = this.Tests(CStr(varKey))
            dTest("status") = etsPending
            Set dTest("assertionResults") = New Collection
        End If
    Next varKey

    DoEvents
    modTestRunnerUI.PollBridgeCancel
    If this.CancelRequested Or (Operation.Status <> eosRunning) Then
        this.State = etrsCancelled
        GoTo AfterTestLoop
    End If

    For Each varKey In colTestKeys

        ' Allow pending UI events (e.g. form close) to fire
        DoEvents
        modTestRunnerUI.PollBridgeCancel

        ' Check both internal flag and operation-level cancellation
        If this.CancelRequested Or (Operation.Status <> eosRunning) Then
            this.State = etrsCancelled
            Exit For
        End If

        strTestKey = CStr(varKey)
        If Not this.Tests.Exists(strTestKey) Then GoTo NextTest

        Set dTest = this.Tests(strTestKey)
        strProcName = CStr(dTest("procName"))
        this.CurrentTestKey = strTestKey
        this.AssertionCounter = 0
        Set this.AssertionResults = New Collection
        lngErrAtStart = Log.ErrorCount
        lngErrJournalAtStart = Log.ErrorJournalCount
        eelAtStart = Operation.ErrorLevel

        modTestRunnerUI.StreamTestStart strTestKey, dTest

        ' Emit module heading when the module changes
        If StrComp(CStr(dTest("moduleName")), strCurrentModule, vbTextCompare) <> 0 Then
            strCurrentModule = CStr(dTest("moduleName"))
            Log.Add strCurrentModule, True, True, vbNullString, True
            Log.Flush   ' Update the console
        End If

        If dTest("sourceType") = "class" Then
            ' --- Class-based test: instantiate via factory, call via CallByName ---
            curStart = Perf.MicroTimer
            On Error Resume Next
            Set obj = Application.Run(BuildRunCmd("TestClassFactory"), CStr(dTest("moduleName")))
            lngErrNum = Err.Number
            strErrDesc = Err.Description
            If lngErrNum <> 0 Then Err.Clear
            On Error GoTo 0

            If lngErrNum <> 0 Then
                dTest("status") = etsErrored
                dTest("errorMessage") = "Factory error " & lngErrNum & ": " & strErrDesc
                this.ErroredCount = this.ErroredCount + 1
                LogTestResult strProcName, "ERROR", "red"
                Log.Add "    " & dTest("errorMessage"), False
                GoTo NextTest
            End If

            On Error Resume Next
            CallByName obj, strProcName, VbMethod
            lngErrNum = Err.Number
            strErrDesc = Err.Description
            If lngErrNum <> 0 Then Err.Clear
            On Error GoTo 0

            ' Release instance (fires Class_Terminate for teardown)
            Set obj = Nothing
        Else
            ' --- Standard module test: Application.Run ---
            If Not GlobalProcExists(strProcName) Then
                dTest("status") = etsErrored
                dTest("errorMessage") = "Not callable (Option Private Module or not found): " & strProcName
                this.ErroredCount = this.ErroredCount + 1
                LogTestResult strProcName, "SKIP", "orange"
                Log.Add "    " & dTest("errorMessage"), False
                GoTo NextTest
            End If

            strCmd = BuildRunCmd(strProcName)
            curStart = Perf.MicroTimer
            On Error Resume Next
            Application.Run strCmd
            lngErrNum = Err.Number
            strErrDesc = Err.Description
            If lngErrNum <> 0 Then Err.Clear
            On Error GoTo 0
        End If

        ' Store assertion results and duration
        Set dTest("assertionResults") = this.AssertionResults
        dTest("durationMs") = CLng((Perf.MicroTimer - curStart) * 1000)

        ' Classify outcome and log result
        If lngErrNum <> 0 Then
            dTest("status") = etsErrored
            dTest("errorMessage") = "Error " & lngErrNum & ": " & strErrDesc
            this.ErroredCount = this.ErroredCount + 1
            LogTestResult strProcName, "ERROR", "red", dTest("durationMs")
            Log.Add "    " & dTest("errorMessage"), False
        ElseIf Log.ErrorCount > lngErrAtStart Or Operation.ErrorLevel > eelAtStart Then
            dTest("status") = etsErrored
            AttachLoggedErrors dTest, lngErrJournalAtStart
            If Not dTest.Exists("errorMessage") Then
                dTest("errorMessage") = T("Logged error(s) during test (see console or TestRun log)")
            End If
            this.ErroredCount = this.ErroredCount + 1
            LogTestResult strProcName, "ERROR", "red", dTest("durationMs")
            Log.Add "    " & dTest("errorMessage"), False
        ElseIf this.AssertionCounter = 0 Then
            dTest("status") = etsEmpty
            this.EmptyCount = this.EmptyCount + 1
            LogTestResult strProcName, "EMPTY", "gray", dTest("durationMs")
        ElseIf AnyAssertionFailed(this.AssertionResults) Then
            dTest("status") = etsFailed
            this.FailedCount = this.FailedCount + 1
            LogTestResult strProcName, "FAIL", "red", dTest("durationMs")
            LogAssertionDetails this.AssertionResults
        Else
            dTest("status") = etsPassed
            this.PassedCount = this.PassedCount + 1
            LogTestResult strProcName, "PASS", "green", dTest("durationMs")
            LogAssertionDetails this.AssertionResults
        End If

        modTestRunnerUI.StreamTestComplete strTestKey, dTest

        lngCurrent = lngCurrent + 1
        Log.Progress lngCurrent, lngTotal, strProcName
        Operation.Pulse

NextTest:
    Next varKey

AfterTestLoop:
    Log.KeepProgressVisible = False
    If Not Log.ProgressBar Is Nothing Then Log.ProgressBar.Hide
    Perf.OperationEnd
    this.CurrentTestKey = vbNullString
    If this.State <> etrsCancelled Then this.State = etrsComplete

    this.LastRunDurationMs = CLng((Perf.MicroTimer - this.RunStartTime) * 1000)
    If this.State = etrsCancelled Then
        modTestRunnerUI.StreamRunCancelled
    Else
        modTestRunnerUI.StreamRunComplete this.LastRunDurationMs
    End If

    LogRunSummary

CleanUp:
    ' Restore original error trapping settings
    Options.BreakOnError = blnBreakOnError
    Application.SetOption "Error Trapping", intErrorTrapping

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunFailed
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Re-run only tests that failed or errored in the last run.
'---------------------------------------------------------------------------------------
'
Public Sub RunFailed()

    If this.Tests Is Nothing Then Exit Sub

    Dim colKeys As New Collection
    Dim varKey As Variant
    Dim dTest As Dictionary
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        If dTest("status") = etsFailed Or dTest("status") = etsErrored Then
            colKeys.Add CStr(varKey)
        End If
    Next varKey

    If colKeys.Count > 0 Then RunSelected colKeys

End Sub


'---------------------------------------------------------------------------------------
' Procedure : InvokeGlobalTestSetup
' Author    : Adam Waller
' Date      : 6/2/2026
' Purpose   : Call the target project's optional GlobalTestSetup hook once per test run.
'           : No-op when the procedure is absent. Hook errors are logged and non-fatal.
'---------------------------------------------------------------------------------------
'
Public Sub InvokeGlobalTestSetup()
    InvokeOptionalGlobalHook GlobalTestSetupProc
End Sub


'---------------------------------------------------------------------------------------
' Procedure : InvokeGlobalTestTeardown
' Author    : Adam Waller
' Date      : 6/2/2026
' Purpose   : Call the target project's optional GlobalTestTeardown hook once per test
'           : run. No-op when the procedure is absent. Hook errors are logged and
'           : non-fatal.
'---------------------------------------------------------------------------------------
'
Public Sub InvokeGlobalTestTeardown()
    InvokeOptionalGlobalHook GlobalTestTeardownProc
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Cancel
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Request cancellation of the current run. The current test will finish
'           : before the run stops.
'---------------------------------------------------------------------------------------
'
Public Sub Cancel()
    this.CancelRequested = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveFilters
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Resolve ParamArray filter arguments into a collection of test keys to run.
'           : Each argument is resolved in priority order:
'           :   1. Module name (exact match)
'           :   2. Suite/folder name (exact or final-segment match on SuiteMap keys)
'           :   3. Procedure name (match on procName or full Module.Procedure key)
'           :   4. Tag (match against test tags)
'           : Prefix with "-" to exclude. Inclusions combine with OR; exclusions with AND.
'           : If only exclusions are specified, the base set is all tests.
'---------------------------------------------------------------------------------------
'
Public Function ResolveFilters(varFilters As Variant) As Collection

    Dim colIncludeKeys As New Collection
    Dim dExcludeKeys As New Dictionary
    Dim blnHasIncludes As Boolean
    Dim lngIdx As Long
    Dim strFilter As String
    Dim strRaw As String
    Dim blnExclude As Boolean
    Dim colMatched As Collection
    Dim varKey As Variant
    Dim i As Long

    For lngIdx = LBound(varFilters) To UBound(varFilters)
        strFilter = Trim$(CStr(varFilters(lngIdx)))
        If Len(strFilter) = 0 Then GoTo NextFilter

        ' Determine if exclusion
        If Left$(strFilter, 1) = "-" Then
            blnExclude = True
            strRaw = Mid$(strFilter, 2)
        Else
            blnExclude = False
            strRaw = strFilter
        End If
        If Len(strRaw) = 0 Then GoTo NextFilter

        ' Resolve the filter value to a set of test keys
        Set colMatched = ResolveOneFilter(strRaw)

        ' Add to inclusion or exclusion set
        If blnExclude Then
            For i = 1 To colMatched.Count
                If Not dExcludeKeys.Exists(CStr(colMatched(i))) Then
                    dExcludeKeys.Add CStr(colMatched(i)), True
                End If
            Next i
        Else
            blnHasIncludes = True
            For i = 1 To colMatched.Count
                On Error Resume Next
                colIncludeKeys.Add CStr(colMatched(i)), CStr(colMatched(i))
                On Error GoTo 0
            Next i
        End If

NextFilter:
    Next lngIdx

    ' Build final result: start from includes (or all tests), then remove exclusions
    Dim colResult As New Collection
    If blnHasIncludes Then
        ' Use the inclusion set as the base
        For i = 1 To colIncludeKeys.Count
            If Not dExcludeKeys.Exists(CStr(colIncludeKeys(i))) Then
                colResult.Add CStr(colIncludeKeys(i))
            End If
        Next i
    Else
        ' No inclusions specified — start with all tests
        For Each varKey In this.Tests.Keys
            If Not dExcludeKeys.Exists(CStr(varKey)) Then
                colResult.Add CStr(varKey)
            End If
        Next varKey
    End If

    Set ResolveFilters = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveOneFilter
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Resolve a single filter value (without the "-" prefix) to matching test
'           : keys using the 4-level priority: module, suite, procedure, tag.
'---------------------------------------------------------------------------------------
'
Private Function ResolveOneFilter(strValue As String) As Collection

    Dim colResult As New Collection
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim varSuiteKey As Variant
    Dim colSuiteTests As Collection
    Dim varParts As Variant
    Dim strLastSegment As String
    Dim i As Long
    Dim colTags As Collection

    ' 1. Try as module name
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        If StrComp(CStr(dTest("moduleName")), strValue, vbTextCompare) = 0 Then
            colResult.Add CStr(varKey)
        End If
    Next varKey
    If colResult.Count > 0 Then
        Set ResolveOneFilter = colResult
        Exit Function
    End If

    ' 2. Try as suite/folder name (exact match or final-segment match)
    For Each varSuiteKey In this.SuiteMap.Keys
        If StrComp(CStr(varSuiteKey), strValue, vbTextCompare) = 0 Then
            ' Exact match on suite key
            Set colSuiteTests = this.SuiteMap(CStr(varSuiteKey))
            For i = 1 To colSuiteTests.Count
                colResult.Add CStr(colSuiteTests(i))
            Next i
        Else
            ' Check if the last dot-segment matches
            varParts = Split(CStr(varSuiteKey), ".")
            strLastSegment = CStr(varParts(UBound(varParts)))
            If StrComp(strLastSegment, strValue, vbTextCompare) = 0 Then
                Set colSuiteTests = this.SuiteMap(CStr(varSuiteKey))
                For i = 1 To colSuiteTests.Count
                    colResult.Add CStr(colSuiteTests(i))
                Next i
            End If
        End If
    Next varSuiteKey
    If colResult.Count > 0 Then
        Set ResolveOneFilter = colResult
        Exit Function
    End If

    ' 3. Try as procedure name or full Module.Procedure key
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        If StrComp(CStr(dTest("procName")), strValue, vbTextCompare) = 0 Then
            colResult.Add CStr(varKey)
        ElseIf StrComp(CStr(varKey), strValue, vbTextCompare) = 0 Then
            colResult.Add CStr(varKey)
        End If
    Next varKey
    If colResult.Count > 0 Then
        Set ResolveOneFilter = colResult
        Exit Function
    End If

    ' 4. Treat as tag
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        Set colTags = dTest("tags")
        If TestHasTag(colTags, strValue) Then
            colResult.Add CStr(varKey)
        End If
    Next varKey

    Set ResolveOneFilter = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : TestHasTag
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Check if a test's tag collection contains the specified tag (case-insensitive).
'---------------------------------------------------------------------------------------
'
Private Function TestHasTag(colTags As Collection, strTag As String) As Boolean

    Dim i As Long
    Dim strLower As String

    If colTags Is Nothing Then Exit Function
    strLower = LCase$(strTag)

    For i = 1 To colTags.Count
        If StrComp(CStr(colTags(i)), strLower, vbTextCompare) = 0 Then
            TestHasTag = True
            Exit Function
        End If
    Next i

End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordAssertion
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Called by HandleTestAssertion when a TestAssert fires during a test run.
'           : Records the result for the current test.
'---------------------------------------------------------------------------------------
'
Public Sub RecordAssertion(ByVal blnCondition As Boolean, ByVal varContext As Variant)

    Dim dResult As Dictionary

    this.AssertionCounter = this.AssertionCounter + 1
    this.TotalAssertions = this.TotalAssertions + 1
    If blnCondition Then
        this.PassedAssertions = this.PassedAssertions + 1
    Else
        this.FailedAssertions = this.FailedAssertions + 1
    End If

    Set dResult = New Dictionary
    dResult.Add "seq", this.AssertionCounter
    dResult.Add "passed", blnCondition
    dResult.Add "context", IIf(IsMissing(varContext) Or IsEmpty(varContext), vbNullString, CStr(varContext))

    this.AssertionResults.Add dResult

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetResultsAsJson
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Serialize the current test results to a JSON string for persistence.
'---------------------------------------------------------------------------------------
'
Public Function GetResultsAsJson() As String

    Dim dRoot As Dictionary
    Dim dSummary As Dictionary
    Dim dTests As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim dTestOut As Dictionary
    Dim colAssertions As Collection
    Dim colAssertOut As Collection
    Dim dA As Dictionary
    Dim dAOut As Dictionary
    Dim i As Long

    Set dRoot = New Dictionary
    Set dSummary = New Dictionary
    Set dTests = New Dictionary

    dRoot.Add "runAt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    dRoot.Add "databasePath", CurrentProject.FullName
    dRoot.Add "addinVersion", GetVCSVersion
    dRoot.Add "durationMs", RunDurationMs

    dSummary.Add "subs", this.PassedCount + this.FailedCount + this.ErroredCount + this.EmptyCount
    dSummary.Add "assertions", this.TotalAssertions
    dSummary.Add "passed", this.PassedAssertions
    dSummary.Add "failed", this.FailedAssertions
    dSummary.Add "errored", this.ErroredCount
    dSummary.Add "empty", this.EmptyCount
    Set dRoot("summary") = dSummary

    If Not this.Tests Is Nothing Then
        Dim colTags As Collection
        Dim colTagsOut As Collection
        For Each varKey In this.Tests.Keys
            Set dTest = this.Tests(CStr(varKey))
            If CLng(dTest("status")) = etsPending Then GoTo NextJsonTest
            Set dTestOut = New Dictionary
            dTestOut.Add "status", StatusToString(CLng(dTest("status")))
            dTestOut.Add "durationMs", dTest("durationMs")
            If dTest.Exists("errorMessage") Then
                dTestOut.Add "errorMessage", dTest("errorMessage")
            End If
            If dTest.Exists("loggedErrors") Then
                Dim colLoggedErrors As Collection
                Dim colLoggedOut As Collection
                Dim dErr As Dictionary
                Dim dErrOut As Dictionary
                Set colLoggedOut = New Collection
                Set colLoggedErrors = dTest("loggedErrors")
                For i = 1 To colLoggedErrors.Count
                    Set dErr = colLoggedErrors(i)
                    Set dErrOut = New Dictionary
                    dErrOut.Add "level", dErr("level")
                    dErrOut.Add "message", dErr("message")
                    If Len(CStr(dErr("source"))) > 0 Then dErrOut.Add "source", dErr("source")
                    If CLng(dErr("errNumber")) <> 0 Then dErrOut.Add "errNumber", dErr("errNumber")
                    If Len(CStr(dErr("errDescription"))) > 0 Then
                        dErrOut.Add "errDescription", dErr("errDescription")
                    End If
                    colLoggedOut.Add dErrOut
                Next i
                Set dTestOut("loggedErrors") = colLoggedOut
            End If

            ' Include tags array
            Set colTagsOut = New Collection
            If dTest.Exists("tags") Then
                Set colTags = dTest("tags")
                For i = 1 To colTags.Count
                    colTagsOut.Add CStr(colTags(i))
                Next i
            End If
            Set dTestOut("tags") = colTagsOut

            Set colAssertOut = New Collection
            Set colAssertions = dTest("assertionResults")
            For i = 1 To colAssertions.Count
                Set dA = colAssertions(i)
                Set dAOut = New Dictionary
                dAOut.Add "seq", dA("seq")
                dAOut.Add "passed", dA("passed")
                If Len(dA("context")) > 0 Then dAOut.Add "context", dA("context")
                colAssertOut.Add dAOut
            Next i
            Set dTestOut("assertions") = colAssertOut
            Set dTests(CStr(varKey)) = dTestOut
NextJsonTest:
        Next varKey
    End If

    Set dRoot("tests") = dTests
    GetResultsAsJson = ConvertToJson(dRoot, JSON_WHITESPACE)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunDurationMs
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Overflow-safe total run duration in ms. Perf.MicroTimer is seconds since
'           : boot; if RunStartTime is 0 (no run on this instance, or state was reset)
'           : then (MicroTimer - 0) * 1000 overflows Long on a machine with meaningful
'           : uptime. Fall back to the last computed duration in that case.
'---------------------------------------------------------------------------------------
'
Private Function RunDurationMs() As Long

    Dim dblMs As Double

    If this.RunStartTime <= 0 Then
        RunDurationMs = this.LastRunDurationMs
        Exit Function
    End If

    dblMs = (Perf.MicroTimer - this.RunStartTime) * 1000
    If dblMs < 0 Or dblMs > 2147483647# Then
        RunDurationMs = this.LastRunDurationMs
    Else
        RunDurationMs = CLng(dblMs)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTestTreeAsJson
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Serialize the scanned test tree for the web runner sidebar (module suites).
'---------------------------------------------------------------------------------------
'
Public Function GetTestTreeAsJson() As String

    Dim dTree As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strModule As String
    Dim dSuite As Dictionary
    Dim colTests As Collection
    Dim dNode As Dictionary
    Dim colTags As Collection
    Dim colTagsOut As Collection
    Dim i As Long

    If this.Tests Is Nothing Then
        GetTestTreeAsJson = "{}"
        Exit Function
    End If

    Set dTree = New Dictionary
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        strModule = CStr(dTest("moduleName"))

        If Not dTree.Exists(strModule) Then
            Set dSuite = New Dictionary
            dSuite.Add "name", strModule
            dSuite.Add "folder", CStr(dTest("folder"))
            Set colTests = New Collection
            Set dSuite("tests") = colTests
            Set dTree(strModule) = dSuite
        End If

        Set dSuite = dTree(strModule)
        Set colTests = dSuite("tests")
        Set dNode = New Dictionary
        dNode.Add "key", CStr(varKey)
        dNode.Add "name", CStr(dTest("procName"))
        dNode.Add "module", strModule
        dNode.Add "procName", CStr(dTest("procName"))
        If dTest.Exists("line") Then dNode.Add "lineNumber", CLng(dTest("line"))
        Set colTagsOut = New Collection
        If dTest.Exists("tags") Then
            Set colTags = dTest("tags")
            For i = 1 To colTags.Count
                colTagsOut.Add CStr(colTags(i))
            Next i
        End If
        dNode.Add "tags", colTagsOut
        colTests.Add dNode
    Next varKey

    GetTestTreeAsJson = ConvertToJson(dTree, JSON_WHITESPACE)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveResults
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Write test results to a timestamped JSON file in the logs folder.
'---------------------------------------------------------------------------------------
'
Public Function SaveResults() As String

    Dim strFolder As String
    Dim strFile As String

    strFolder = Options.GetExportFolder & "logs" & PathSep
    VerifyPath strFolder
    strFile = strFolder & "TestResults_" & _
        Format$(Now, "yyyymmdd\_hhnnss") & "_" & _
        Format$(Timer * 1000 Mod 1000, "000") & ".json"

    WriteFile GetResultsAsJson, strFile
    SaveResults = strFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadStateTests
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Reconstruct the Tests dictionary and summary counts from durable state.
'---------------------------------------------------------------------------------------
'
Public Sub LoadStateTests(dStateTests As Dictionary)

    Dim varKey As Variant
    Dim dSrc As Dictionary
    Dim dTest As Dictionary
    Dim strFolder As String
    Dim colSuite As Collection

    Set this.Tests = New Dictionary
    Set this.SuiteMap = New Dictionary
    ResetCounts

    For Each varKey In dStateTests.Keys
        Set dSrc = dStateTests(CStr(varKey))
        Set dTest = DictionaryFromStateEntry(CStr(varKey), dSrc)
        this.Tests.Add CStr(varKey), dTest

        strFolder = CStr(dTest("folder"))
        If Not this.SuiteMap.Exists(strFolder) Then
            Set colSuite = New Collection
            this.SuiteMap.Add strFolder, colSuite
        End If
        this.SuiteMap(strFolder).Add CStr(varKey)

        AccumulateCountsFromTest dTest
    Next varKey

    this.State = etrsComplete

End Sub


' ===================== Private helpers =====================


'---------------------------------------------------------------------------------------
' Procedure : ProjectHasFolderAnnotations
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Quick check whether any module in the project contains @Folder annotations.
'---------------------------------------------------------------------------------------
'
Private Function ProjectHasFolderAnnotations(proj As VBIDE.VBProject) As Boolean

    Dim cmp As VBIDE.VBComponent
    Dim strCode As String

    On Error Resume Next
    For Each cmp In proj.VBComponents
        If cmp.Type = vbext_ct_StdModule Or cmp.Type = vbext_ct_ClassModule Then
            If cmp.CodeModule.CountOfLines > 0 Then
                strCode = UCase$(cmp.CodeModule.Lines(1, IIf(cmp.CodeModule.CountOfLines < 30, cmp.CodeModule.CountOfLines, 30)))
                If InStr(strCode, "'@FOLDER(") > 0 Then
                    ProjectHasFolderAnnotations = True
                    Exit Function
                End If
            End If
        End If
    Next cmp
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsTestModule
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Determine whether a VBComponent is a test module using discovery rules:
'           :   1. Standard modules or standalone class modules (not forms/reports)
'           :   2. @Folder annotation containing "Tests" segment (if project uses folders)
'           :   3. Name containing "Test" (fallback if no @Folder tags in project)
'---------------------------------------------------------------------------------------
'
Private Function IsTestModule(cmp As VBIDE.VBComponent, _
                              blnHasFolderTags As Boolean) As Boolean

    Dim strCode As String
    Dim strFolder As String

    ' Only scan standard modules and standalone class modules
    Select Case cmp.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule
        Case Else: Exit Function
    End Select

    On Error Resume Next

    If blnHasFolderTags Then
        ' Use @Folder annotation: look for a "Tests" segment
        If cmp.CodeModule.CountOfLines > 0 Then
            strCode = cmp.CodeModule.Lines(1, IIf(cmp.CodeModule.CountOfLines < 30, cmp.CodeModule.CountOfLines, 30))
            strFolder = ExtractFolderAnnotation(strCode)
            If Len(strFolder) > 0 Then
                IsTestModule = FolderContainsTestsSegment(strFolder)
                Exit Function
            End If
        End If
    End If

    ' Fallback: name-based discovery (*Test* in the name)
    If InStr(1, cmp.Name, "Test", vbTextCompare) > 0 Then
        IsTestModule = True
    End If

    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractFolderAnnotation
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Extract the folder path from an @Folder annotation in module source.
'           : Returns empty string if not found.
'---------------------------------------------------------------------------------------
'
Private Function ExtractFolderAnnotation(strCode As String) As String

    Dim strUpper As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    Const TAG As String = "'@FOLDER("

    strUpper = UCase$(strCode)
    lngPos = InStr(1, strUpper, TAG)
    If lngPos = 0 Then Exit Function

    lngStart = InStr(lngPos, strCode, """")
    If lngStart > 0 Then
        lngEnd = InStr(lngStart + 1, strCode, """")
        If lngEnd > lngStart + 1 Then
            ExtractFolderAnnotation = Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderContainsTestsSegment
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if a folder path (dot-separated) contains "Tests" as a segment.
'---------------------------------------------------------------------------------------
'
Private Function FolderContainsTestsSegment(strFolder As String) As Boolean

    Dim varParts As Variant
    Dim i As Long
    varParts = Split(strFolder, ".")
    For i = LBound(varParts) To UBound(varParts)
        If StrComp(varParts(i), "Tests", vbTextCompare) = 0 Then
            FolderContainsTestsSegment = True
            Exit Function
        End If
    Next i

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractTags
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Parse all '@Tag("...") annotations from a block of text. Returns a
'           : collection of tag name strings (lowercased for case-insensitive matching).
'---------------------------------------------------------------------------------------
'
Private Function ExtractTags(strCode As String) As Collection

    Dim colResult As New Collection
    Dim strUpper As String
    Dim lngPos As Long
    Dim lngQuoteStart As Long
    Dim lngQuoteEnd As Long
    Dim strTag As String

    Const TAG_MARKER As String = "'@TAG("

    strUpper = UCase$(strCode)
    lngPos = 1

    Do
        lngPos = InStr(lngPos, strUpper, TAG_MARKER)
        If lngPos = 0 Then Exit Do

        ' Find the quoted tag name
        lngQuoteStart = InStr(lngPos, strCode, """")
        If lngQuoteStart = 0 Then Exit Do

        lngQuoteEnd = InStr(lngQuoteStart + 1, strCode, """")
        If lngQuoteEnd <= lngQuoteStart Then Exit Do

        strTag = LCase$(Mid$(strCode, lngQuoteStart + 1, lngQuoteEnd - lngQuoteStart - 1))
        If Len(strTag) > 0 Then
            ' Avoid duplicates
            On Error Resume Next
            colResult.Add strTag, strTag
            On Error GoTo 0
        End If

        lngPos = lngQuoteEnd + 1
    Loop

    Set ExtractTags = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractProcedureTags
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Scan forward from the procedure declaration line to find '@Tag("...")
'           : annotations inside the procedure body. Collects tags from comment and
'           : blank lines at the top of the body; stops at the first executable line.
'---------------------------------------------------------------------------------------
'
Private Function ExtractProcedureTags(varLines As Variant, ByVal lngDeclLine As Long) As Collection

    Dim colResult As New Collection
    Dim lngScan As Long
    Dim strTrimmed As String
    Dim strBlock As String

    strBlock = vbNullString

    ' Walk forward from the line after the declaration
    For lngScan = lngDeclLine + 1 To UBound(varLines)
        strTrimmed = Trim$(varLines(lngScan))

        ' Blank line — continue scanning
        If Len(strTrimmed) = 0 Then
            ' Skip
        ElseIf Left$(strTrimmed, 1) = "'" Then
            ' Comment line — accumulate for tag extraction
            strBlock = strBlock & strTrimmed & vbCrLf
        ElseIf Left$(UCase$(strTrimmed), 10) = "ATTRIBUTE " Then
            ' VBA member attribute line — skip (appears after declaration in .cls files)
        Else
            ' Hit executable code — stop
            Exit For
        End If
    Next lngScan

    If Len(strBlock) > 0 Then
        Set colResult = ExtractTags(strBlock)
    End If

    Set ExtractProcedureTags = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeTagCollections
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Combine two tag collections into one, de-duplicating by key.
'---------------------------------------------------------------------------------------
'
Private Function MergeTagCollections(col1 As Collection, col2 As Collection) As Collection

    Dim colResult As New Collection
    Dim i As Long
    Dim strTag As String

    On Error Resume Next
    If Not col1 Is Nothing Then
        For i = 1 To col1.Count
            strTag = col1(i)
            colResult.Add strTag, strTag
        Next i
    End If
    If Not col2 Is Nothing Then
        For i = 1 To col2.Count
            strTag = col2(i)
            colResult.Add strTag, strTag
        Next i
    End If
    On Error GoTo 0

    Set MergeTagCollections = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : SyncFactoryEntries
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Ensure the TestClassFactory function in modTestAssert has Select Case
'           : entries matching the currently discovered test classes. Only modifies
'           : the CodeModule if the entries are out of date.
'---------------------------------------------------------------------------------------
'
Private Sub SyncFactoryEntries(proj As VBIDE.VBProject)

    Dim cmp As VBIDE.VBComponent
    Dim colClassNames As Collection
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strClassName As String
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim lngLine As Long
    Dim strLine As String
    Dim strExpected As String
    Dim strExisting As String
    Dim i As Long

    ' Collect unique class names from discovered tests
    Set colClassNames = New Collection
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        If dTest("sourceType") = "class" Then
            strClassName = CStr(dTest("moduleName"))
            ' Use dictionary-style uniqueness check via error trapping
            On Error Resume Next
            colClassNames.Add strClassName, strClassName
            On Error GoTo 0
        End If
    Next varKey

    ' No class-based tests — nothing to sync
    If colClassNames.Count = 0 Then Exit Sub

    ' Sort for deterministic output
    Set colClassNames = SortCollectionByValue(colClassNames)

    ' Locate modTestAssert in the project
    On Error Resume Next
    Set cmp = proj.VBComponents("modTestAssert")
    On Error GoTo 0
    If cmp Is Nothing Then Exit Sub

    ' Find the BEGIN/END markers
    lngBegin = 0
    lngEnd = 0
    For lngLine = 1 To cmp.CodeModule.CountOfLines
        strLine = cmp.CodeModule.Lines(lngLine, 1)
        If InStr(1, strLine, "BEGIN TEST CLASS ENTRIES", vbTextCompare) > 0 Then
            lngBegin = lngLine
        ElseIf InStr(1, strLine, "END TEST CLASS ENTRIES", vbTextCompare) > 0 Then
            lngEnd = lngLine
            Exit For
        End If
    Next lngLine

    ' Markers not found — append the full TestClassFactory template
    If lngBegin = 0 Or lngEnd = 0 Then
        AppendFactoryTemplate cmp, colClassNames
        Exit Sub
    End If

    ' Build expected content between markers
    strExpected = BuildFactorySelectCase(colClassNames)

    ' Read existing content between markers
    strExisting = vbNullString
    If lngEnd > lngBegin + 1 Then
        strExisting = cmp.CodeModule.Lines(lngBegin + 1, lngEnd - lngBegin - 1)
    End If

    ' Only modify if different
    If StrComp(strExisting, strExpected, vbBinaryCompare) = 0 Then Exit Sub

    ' Replace lines between markers
    If lngEnd > lngBegin + 1 Then
        cmp.CodeModule.DeleteLines lngBegin + 1, lngEnd - lngBegin - 1
    End If
    cmp.CodeModule.InsertLines lngBegin + 1, strExpected

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildFactorySelectCase
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Generate the Select Case block for TestClassFactory given a collection
'           : of class names.
'---------------------------------------------------------------------------------------
'
Private Function BuildFactorySelectCase(colClassNames As Collection) As String

    Dim strLines As String
    Dim i As Long

    strLines = "    Select Case strClassName" & vbCrLf
    For i = 1 To colClassNames.Count
        strLines = strLines & "        Case """ & colClassNames(i) & """: Set TestClassFactory = New " & colClassNames(i) & vbCrLf
    Next i
    strLines = strLines & "    End Select"

    BuildFactorySelectCase = strLines

End Function


'---------------------------------------------------------------------------------------
' Procedure : AppendFactoryTemplate
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Append the full TestClassFactory function to modTestAssert when the
'           : BEGIN/END markers are not yet present.
'---------------------------------------------------------------------------------------
'
Private Sub AppendFactoryTemplate(cmp As VBIDE.VBComponent, colClassNames As Collection)

    Dim strTemplate As String

    strTemplate = vbCrLf & vbCrLf & _
        "'---------------------------------------------------------------------------------------" & vbCrLf & _
        "' Procedure : TestClassFactory" & vbCrLf & _
        "' Purpose   : Create a new instance of a test class by name. The Select Case entries" & vbCrLf & _
        "'           : below are auto-generated by the VCS test runner during test discovery." & vbCrLf & _
        "'           : Do not manually edit the lines between the BEGIN/END markers." & vbCrLf & _
        "'---------------------------------------------------------------------------------------" & vbCrLf & _
        "'" & vbCrLf & _
        "Public Function TestClassFactory(ByVal strClassName As String) As Object" & vbCrLf & _
        "'--- BEGIN TEST CLASS ENTRIES (auto-generated by VCS test runner) ---" & vbCrLf & _
        BuildFactorySelectCase(colClassNames) & vbCrLf & _
        "'--- END TEST CLASS ENTRIES ---" & vbCrLf & _
        "End Function"

    cmp.CodeModule.InsertLines cmp.CodeModule.CountOfLines + 1, strTemplate

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanModuleForTests
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Scan a test module for parameterless Public Subs (and Functions for class
'           : modules) and register each one as a test. Populates the Tests dictionary.
'---------------------------------------------------------------------------------------
'
Private Sub ScanModuleForTests(cmp As VBIDE.VBComponent, blnHasFolderTags As Boolean)

    Dim strAllCode As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim strTrimmed As String
    Dim strProcName As String
    Dim strFolder As String
    Dim strTestKey As String
    Dim dTest As Dictionary
    Dim blnIsClass As Boolean
    Dim blnMatched As Boolean
    Dim colModuleTags As Collection
    Dim colProcTags As Collection
    Dim colCombined As Collection
    Dim strHeader As String

    On Error Resume Next
    If cmp.CodeModule.CountOfLines = 0 Then Exit Sub
    strAllCode = cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
    On Error GoTo 0
    If Len(strAllCode) = 0 Then Exit Sub

    blnIsClass = (cmp.Type = vbext_ct_ClassModule)

    ' Determine folder annotation for suite grouping
    If blnHasFolderTags Then
        strFolder = ExtractFolderAnnotation(strAllCode)
    End If
    If Len(strFolder) = 0 Then strFolder = cmp.Name

    varLines = Split(strAllCode, vbCrLf)

    ' Extract module-level tags from the header area (~30 lines)
    Dim lngHeaderEnd As Long
    lngHeaderEnd = UBound(varLines)
    If lngHeaderEnd > 29 Then lngHeaderEnd = 29
    strHeader = vbNullString
    For lngLine = LBound(varLines) To lngHeaderEnd
        strHeader = strHeader & varLines(lngLine) & vbCrLf
    Next lngLine
    Set colModuleTags = ExtractTags(strHeader)

    ' Count TestAssert calls to determine if this is a test module and how many assertions it contains
    Dim lngAssertCount As Long
    For lngLine = LBound(varLines) To UBound(varLines)
        If Left$(UCase$(Trim$(varLines(lngLine))), 11) = "TESTASSERT " Then
            lngAssertCount = lngAssertCount + 1
        End If
    Next lngLine
    If lngAssertCount = 0 Then Exit Sub
    this.ScannedAssertionCount = this.ScannedAssertionCount + lngAssertCount

    For lngLine = LBound(varLines) To UBound(varLines)
        strTrimmed = Trim$(varLines(lngLine))

        ' For class modules, match both Sub and Function
        If blnIsClass Then
            blnMatched = IsParameterlessMethod(strTrimmed, strProcName)
        Else
            blnMatched = IsParameterlessSub(strTrimmed, strProcName)
        End If

        If blnMatched Then
            ' Skip class lifecycle methods
            If blnIsClass And Left$(LCase$(strProcName), 6) = "class_" Then GoTo NextLine

            strTestKey = cmp.Name & "." & strProcName

            ' Collect procedure-level tags from the preceding comment block
            Set colProcTags = ExtractProcedureTags(varLines, lngLine)

            ' Combine module-level and procedure-level tags
            Set colCombined = MergeTagCollections(colModuleTags, colProcTags)

            Set dTest = New Dictionary
            dTest.Add "moduleName", cmp.Name
            dTest.Add "procName", strProcName
            dTest.Add "line", CLng(lngLine + 1)
            dTest.Add "folder", strFolder
            dTest.Add "sourceType", IIf(blnIsClass, "class", "module")
            dTest.Add "status", CLng(etsPending)
            dTest.Add "durationMs", CLng(0)
            Set dTest("tags") = colCombined
            Set dTest("assertionResults") = New Collection

            this.Tests.Add strTestKey, dTest

            ' Add to suite map
            If Not this.SuiteMap.Exists(strFolder) Then
                this.SuiteMap.Add strFolder, New Collection
            End If
            this.SuiteMap(strFolder).Add strTestKey
        End If

NextLine:
    Next lngLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsParameterlessSub
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if a line declares a parameterless Public Sub. Returns the sub name
'           : via the ByRef parameter if matched.
'---------------------------------------------------------------------------------------
'
Private Function IsParameterlessSub(strTrimmed As String, ByRef strName As String) As Boolean

    Dim strUpper As String
    Dim lngSubPos As Long
    Dim lngParenOpen As Long
    Dim lngParenClose As Long
    Dim strBetween As String

    strUpper = UCase$(strTrimmed)

    ' Match "Public Sub X()" or "Sub X()" (Public is default)
    If Left$(strUpper, 11) = "PUBLIC SUB " Then
        lngSubPos = 12
    ElseIf Left$(strUpper, 4) = "SUB " Then
        ' Exclude "Private Sub"
        If Left$(strUpper, 12) = "PRIVATE SUB " Then Exit Function
        lngSubPos = 5
    Else
        Exit Function
    End If

    ' Extract name (everything up to the opening paren)
    lngParenOpen = InStr(lngSubPos, strTrimmed, "(")
    If lngParenOpen = 0 Then Exit Function

    strName = Trim$(Mid$(strTrimmed, lngSubPos, lngParenOpen - lngSubPos))
    If Len(strName) = 0 Then Exit Function

    ' Check for empty parameter list
    lngParenClose = InStr(lngParenOpen + 1, strTrimmed, ")")
    If lngParenClose = 0 Then Exit Function

    strBetween = Trim$(Mid$(strTrimmed, lngParenOpen + 1, lngParenClose - lngParenOpen - 1))
    IsParameterlessSub = (Len(strBetween) = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsParameterlessMethod
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Check if a line declares a parameterless Public Sub or Public Function.
'           : Returns the method name via the ByRef parameter if matched. Used for
'           : class-based test discovery where both Subs and Functions are valid tests.
'---------------------------------------------------------------------------------------
'
Private Function IsParameterlessMethod(strTrimmed As String, ByRef strName As String) As Boolean

    Dim strUpper As String
    Dim lngPos As Long
    Dim lngParenOpen As Long
    Dim lngParenClose As Long
    Dim strBetween As String

    strUpper = UCase$(strTrimmed)

    ' Match Public Sub/Function or bare Sub/Function (implicit Public)
    If Left$(strUpper, 11) = "PUBLIC SUB " Then
        lngPos = 12
    ElseIf Left$(strUpper, 16) = "PUBLIC FUNCTION " Then
        lngPos = 17
    ElseIf Left$(strUpper, 4) = "SUB " Then
        If Left$(strUpper, 12) = "PRIVATE SUB " Then Exit Function
        lngPos = 5
    ElseIf Left$(strUpper, 9) = "FUNCTION " Then
        If Left$(strUpper, 17) = "PRIVATE FUNCTION " Then Exit Function
        lngPos = 10
    Else
        Exit Function
    End If

    lngParenOpen = InStr(lngPos, strTrimmed, "(")
    If lngParenOpen = 0 Then Exit Function

    strName = Trim$(Mid$(strTrimmed, lngPos, lngParenOpen - lngPos))
    If Len(strName) = 0 Then Exit Function

    lngParenClose = InStr(lngParenOpen + 1, strTrimmed, ")")
    If lngParenClose = 0 Then Exit Function

    strBetween = Trim$(Mid$(strTrimmed, lngParenOpen + 1, lngParenClose - lngParenOpen - 1))
    IsParameterlessMethod = (Len(strBetween) = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindTestAssertCall
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Find "TestAssert " at the start of a statement (case-insensitive).
'           : Returns the position of the character after "TestAssert " (the expression
'           : start), or 0 if not found.
'---------------------------------------------------------------------------------------
' Procedure : BuildRunCmd
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Build a project-qualified Application.Run command for a procedure in
'           : CurrentVBProject. Follows the RunSubInCurrentProject pattern from
'           : modDatabase.bas.
'---------------------------------------------------------------------------------------
'
Private Function BuildRunCmd(strProcName As String) As String

    If CurrentVBProject.Name = PROJECT_NAME Then
        With CurrentProject
            BuildRunCmd = .Path & PathSep & FSO.GetBaseName(.Name) & "." & strProcName
        End With
    Else
        BuildRunCmd = "[" & CurrentVBProject.Name & "]." & strProcName
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : InvokeOptionalGlobalHook
' Author    : Adam Waller
' Date      : 6/2/2026
' Purpose   : Invoke an optional parameterless public hook in the target database via
'           : Application.Run. Silent when the procedure does not exist. Logs hook
'           : failures via Log.Add (never Log.Error) so they do not affect test counts.
'---------------------------------------------------------------------------------------
'
Private Sub InvokeOptionalGlobalHook(strProcName As String)

    Dim lngErrNum As Long
    Dim strErrDesc As String

    On Error Resume Next
    Set VBE.ActiveVBProject = CurrentVBProject
    On Error GoTo 0

    If Not GlobalProcExists(strProcName) Then Exit Sub

    On Error Resume Next
    Application.Run BuildRunCmd(strProcName)
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If lngErrNum <> 0 Then
        Log.Add T("Global hook {0} error {1}: {2}", _
            var0:=strProcName, var1:=lngErrNum, var2:=strErrDesc), False, , "orange"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AttachLoggedErrors
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Copy Log.Error journal entries from a test run into the test dictionary
'           : for TestResults JSON export (loggedErrors + errorMessage).
'---------------------------------------------------------------------------------------
'
Private Sub AttachLoggedErrors(dTest As Dictionary, lngJournalAtStart As Long)

    Dim colNew As Collection
    Dim colStored As Collection
    Dim dErr As Dictionary
    Dim varItem As Variant
    Dim strFirst As String

    If Log.ErrorJournalCount <= lngJournalAtStart Then Exit Sub

    Set colNew = Log.GetErrorJournalSince(lngJournalAtStart)
    Set colStored = New Collection
    For Each varItem In colNew
        Set dErr = varItem
        colStored.Add dErr
        If Len(strFirst) = 0 Then strFirst = CStr(dErr("message"))
    Next varItem

    Set dTest("loggedErrors") = colStored
    If colNew.Count > 1 Then
        dTest("errorMessage") = strFirst & T(" (+{0} more logged error(s))", var0:=colNew.Count - 1)
    Else
        dTest("errorMessage") = strFirst
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AnyAssertionFailed
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if any assertion in a results collection failed.
'---------------------------------------------------------------------------------------
'
Private Function AnyAssertionFailed(colResults As Collection) As Boolean

    Dim dResult As Dictionary
    Dim i As Long
    For i = 1 To colResults.Count
        Set dResult = colResults(i)
        If Not CBool(dResult("passed")) Then
            AnyAssertionFailed = True
            Exit Function
        End If
    Next i

End Function


'---------------------------------------------------------------------------------------
' Procedure : StatusToString
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Convert a numeric test status to a display string.
'---------------------------------------------------------------------------------------
'
Private Function StatusToString(lngStatus As Long) As String
    Select Case lngStatus
        Case etsPassed:  StatusToString = "PASSED"
        Case etsFailed:  StatusToString = "FAILED"
        Case etsErrored: StatusToString = "ERRORED"
        Case etsEmpty:   StatusToString = "EMPTY"
        Case Else:       StatusToString = "PENDING"
    End Select
End Function


Private Function StringToStatus(ByVal strStatus As String) As eTestStatus

    Select Case UCase$(Trim$(strStatus))
        Case "PASSED":  StringToStatus = etsPassed
        Case "FAILED":  StringToStatus = etsFailed
        Case "ERRORED": StringToStatus = etsErrored
        Case "EMPTY":   StringToStatus = etsEmpty
        Case Else:      StringToStatus = etsPending
    End Select

End Function


Private Function DictionaryFromStateEntry(ByVal strKey As String, ByVal dSrc As Dictionary) As Dictionary

    Dim dTest As Dictionary
    Dim colTags As Collection
    Dim colTagsOut As Collection
    Dim colAssertOut As Collection
    Dim colAssertions As Collection
    Dim colLoggedOut As Collection
    Dim dA As Dictionary
    Dim dAOut As Dictionary
    Dim dErr As Dictionary
    Dim dErrOut As Dictionary
    Dim i As Long

    Set dTest = New Dictionary
    dTest.Add "moduleName", CStr(dSrc("moduleName"))
    dTest.Add "procName", CStr(dSrc("procName"))
    dTest.Add "folder", CStr(dSrc("folder"))
    dTest.Add "line", CLng(Nz(dSrc("line"), 0))
    dTest.Add "sourceType", CStr(Nz(dSrc("sourceType"), "module"))
    dTest.Add "status", CLng(StringToStatus(CStr(Nz(dSrc("status"), "PENDING"))))
    dTest.Add "durationMs", CLng(Nz(dSrc("durationMs"), 0))

    If dSrc.Exists("errorMessage") Then
        If Len(CStr(dSrc("errorMessage"))) > 0 Then
            dTest.Add "errorMessage", CStr(dSrc("errorMessage"))
        End If
    End If

    Set colTagsOut = New Collection
    If dSrc.Exists("tags") Then
        If TypeName(dSrc("tags")) = "Collection" Then
            Set colTags = dSrc("tags")
            For i = 1 To colTags.Count
                colTagsOut.Add CStr(colTags(i))
            Next i
        End If
    End If
    Set dTest("tags") = colTagsOut

    Set colAssertOut = New Collection
    If dSrc.Exists("assertions") Then
        If TypeName(dSrc("assertions")) = "Collection" Then
            Set colAssertions = dSrc("assertions")
            For i = 1 To colAssertions.Count
                Set dA = colAssertions(i)
                Set dAOut = New Dictionary
                dAOut.Add "seq", CLng(Nz(dA("seq"), i))
                dAOut.Add "passed", CBool(Nz(dA("passed"), False))
                If dA.Exists("context") Then
                    If Len(CStr(dA("context"))) > 0 Then
                        dAOut.Add "context", CStr(dA("context"))
                    End If
                End If
                colAssertOut.Add dAOut
            Next i
        End If
    End If
    Set dTest("assertionResults") = colAssertOut

    If dSrc.Exists("loggedErrors") Then
        If TypeName(dSrc("loggedErrors")) = "Collection" Then
            Set colLoggedOut = New Collection
            Set colAssertions = dSrc("loggedErrors")
            For i = 1 To colAssertions.Count
                Set dErr = colAssertions(i)
                Set dErrOut = New Dictionary
                dErrOut.Add "level", dErr("level")
                dErrOut.Add "message", dErr("message")
                If dErr.Exists("source") Then dErrOut.Add "source", CStr(dErr("source"))
                If dErr.Exists("errNumber") Then dErrOut.Add "errNumber", CLng(dErr("errNumber"))
                If dErr.Exists("errDescription") Then
                    dErrOut.Add "errDescription", CStr(dErr("errDescription"))
                End If
                colLoggedOut.Add dErrOut
            Next i
            Set dTest("loggedErrors") = colLoggedOut
        End If
    End If

    Set DictionaryFromStateEntry = dTest

End Function


Private Sub AccumulateCountsFromTest(ByVal dTest As Dictionary)

    Dim colAssertions As Collection
    Dim dA As Dictionary
    Dim i As Long

    Select Case CLng(dTest("status"))
        Case etsPassed
            this.PassedCount = this.PassedCount + 1
        Case etsFailed
            this.FailedCount = this.FailedCount + 1
        Case etsErrored
            this.ErroredCount = this.ErroredCount + 1
        Case etsEmpty
            this.EmptyCount = this.EmptyCount + 1
        Case Else
            Exit Sub
    End Select

    If dTest.Exists("assertionResults") Then
        Set colAssertions = dTest("assertionResults")
        For i = 1 To colAssertions.Count
            Set dA = colAssertions(i)
            this.TotalAssertions = this.TotalAssertions + 1
            If CBool(Nz(dA("passed"), False)) Then
                this.PassedAssertions = this.PassedAssertions + 1
            Else
                this.FailedAssertions = this.FailedAssertions + 1
            End If
        Next i
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogTestResult
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Log a single test result line with the status left-justified at a
'           : fixed column position so all result labels start at the same place,
'           : followed by the test duration.
'---------------------------------------------------------------------------------------
'
Private Sub LogTestResult(strProcName As String, strStatus As String, strColor As String, _
    Optional lngDurationMs As Long = -1)
    Const STATUS_COL As Long = 46
    Dim strLine As String
    Dim lngPad As Long
    strLine = "  " & strProcName
    lngPad = STATUS_COL - Len(strLine)
    If lngPad < 2 Then lngPad = 2
    Log.Add strLine & Space$(lngPad), True, False
    If lngDurationMs >= 0 Then
        Log.Add strStatus & "  " & FormatDuration(lngDurationMs), True, True, strColor
    Else
        Log.Add strStatus, True, True, strColor
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogRunSummary
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Print an end-of-run summary with each result category on its own line,
'           : non-zero counts bold in the appropriate color.
'---------------------------------------------------------------------------------------
'
Private Sub LogRunSummary()

    Const SPACER_LEN As Long = 61

    Dim lngTotal As Long
    Dim dblElapsed As Double

    lngTotal = this.PassedCount + this.FailedCount + this.ErroredCount + this.EmptyCount
    dblElapsed = Round((Perf.MicroTimer - this.RunStartTime), 3)

    With Log
        .Spacer True, SPACER_LEN
        .Add "Test summary", True, , "blue", True
        .Add "  Subs:         " & lngTotal
        .Add "  Assertions:   " & this.TotalAssertions
        LogSummaryLine "    Passed:     ", CStr(this.PassedAssertions), _
            IIf(this.PassedAssertions > 0, "green", vbNullString)
        LogSummaryLine "    Failed:     ", CStr(this.FailedAssertions), _
            IIf(this.FailedAssertions > 0, "red", vbNullString)
        LogSummaryLine "  Errored:      ", CStr(this.ErroredCount), _
            IIf(this.ErroredCount > 0, "red", vbNullString)
        LogSummaryLine "  Empty:        ", CStr(this.EmptyCount), _
            IIf(this.EmptyCount > 0, "gray", vbNullString)
        .Add "  Elapsed:      " & FormatDuration(CLng(dblElapsed * 1000))
        .Spacer True, SPACER_LEN
    End With

    LogTestTimingTable SPACER_LEN

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogSummaryLine
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Print a summary label + value, with the value bold and colored when
'           : a color is specified (non-zero count).
'---------------------------------------------------------------------------------------
'
Private Sub LogSummaryLine(ByVal strLabel As String, ByVal strValue As String, ByVal strColor As String)
    If Len(strColor) > 0 Then
        Log.Add strLabel, True, False
        Log.Add strValue, True, True, strColor, True
    Else
        Log.Add strLabel & strValue
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogTestTimingTable
' Author    : Adam Waller
' Date      : 5/11/2026
' Purpose   : Print a slowest-first table of per-test durations after the summary.
'           : Only shown when 2+ tests were run.
'---------------------------------------------------------------------------------------
'
Private Sub LogTestTimingTable(ByVal lngSpacerLen As Long)

    Dim lngCount As Long
    Dim lngExecuted As Long
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim i As Long
    Dim lngMaxName As Long
    Dim strProcName As String
    Dim strDur As String
    Dim lngPad As Long
    Dim varParts As Variant

    If this.Tests Is Nothing Then Exit Sub

    ' Count only tests that actually ran (skip pending/unexecuted). Guard against
    ' non-object entries (a defensive backstop should this.Tests ever hold a stray
    ' Empty value from an auto-vivified key).
    For Each varKey In this.Tests.Keys
        If IsObject(this.Tests(CStr(varKey))) Then
            Set dTest = this.Tests(CStr(varKey))
            If CLng(dTest("status")) <> etsPending Then lngExecuted = lngExecuted + 1
        End If
    Next varKey
    If lngExecuted < 2 Then Exit Sub

    ' Build sortable records: "00000000|procName|STATUS"
    ' Same pattern as clsPerformance.SortItemsByTime -- fixed-width key
    ' on the left so QuickSort produces ascending order, then walk
    ' backwards for descending (slowest first).
    Dim varRecords() As Variant
    ReDim varRecords(0 To lngExecuted - 1)

    i = 0
    For Each varKey In this.Tests.Keys
        If Not IsObject(this.Tests(CStr(varKey))) Then GoTo NextTimingEntry
        Set dTest = this.Tests(CStr(varKey))
        If CLng(dTest("status")) = etsPending Then GoTo NextTimingEntry
        strProcName = CStr(dTest("procName"))
        varRecords(i) = Format$(CLng(dTest("durationMs")), "00000000") & "|" & _
            strProcName & "|" & StatusToString(CLng(dTest("status")))
        If Len(strProcName) > lngMaxName Then lngMaxName = Len(strProcName)
        i = i + 1
NextTimingEntry:
    Next varKey

    QuickSort varRecords

    Const DUR_COL As Long = 10

    With Log
        .Add T("Test timing (slowest first)"), False
        For i = lngExecuted - 1 To 0 Step -1
            varParts = Split(CStr(varRecords(i)), "|")
            strProcName = CStr(varParts(1))
            strDur = FormatDuration(CLng(varParts(0)))
            lngPad = (lngMaxName + 2) - Len(strProcName)
            If lngPad < 2 Then lngPad = 2
            .Add "  " & strProcName & Space$(lngPad) & _
                Space$(DUR_COL - Len(strDur)) & strDur & "  " & CStr(varParts(2)), False
        Next i
        .Spacer False, lngSpacerLen
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogAssertionDetails
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Log per-assertion detail lines to the console with color coding.
'---------------------------------------------------------------------------------------
'
Private Sub LogAssertionDetails(colResults As Collection)
    Dim dResult As Dictionary
    Dim i As Long
    Dim strTag As String
    Dim strColor As String
    For i = 1 To colResults.Count
        Set dResult = colResults(i)
        If CBool(dResult("passed")) Then
            strTag = "PASS"
            strColor = "green"
        Else
            strTag = "FAIL"
            strColor = "red"
        End If
        Log.Add "    #" & dResult("seq") & " ", True, False, "gray"
        Log.Add strTag, True, IIf(Len(dResult("context")) = 0, True, False), strColor
        If Len(dResult("context")) > 0 Then
            Log.Add ": " & dResult("context"), True, True, "gray"
        End If
    Next i
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FormatDuration
' Author    : Adam Waller
' Date      : 5/11/2026
' Purpose   : Convert milliseconds to a compact human-readable string.
'           :   < 1000  -> "25ms"
'           :   < 60000 -> "12.5s"
'           :   else    -> "2.1m"
'---------------------------------------------------------------------------------------
'
Private Function FormatDuration(ByVal lngMs As Long) As String
    If lngMs < 1000 Then
        FormatDuration = lngMs & "ms"
    ElseIf lngMs < 60000 Then
        FormatDuration = Format$(lngMs / 1000, "0.00") & "s"
    Else
        FormatDuration = Format$(lngMs / 60000, "0.0") & "m"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetCounts
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Reset run summary counters.
'---------------------------------------------------------------------------------------
'
Private Sub ResetCounts()
    this.PassedCount = 0
    this.FailedCount = 0
    this.ErroredCount = 0
    this.EmptyCount = 0
    this.TotalAssertions = 0
    this.PassedAssertions = 0
    this.FailedAssertions = 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanDiscover
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Discover test modules and procedures into fresh Tests/SuiteMap dictionaries.
'           : Caller sets etrsScanning and final state.
'---------------------------------------------------------------------------------------
'
Private Sub ScanDiscover()

    Dim cmp As VBIDE.VBComponent
    Dim proj As VBIDE.VBProject
    Dim blnHasFolderTags As Boolean
    Dim dModules As Dictionary
    Dim varKey As Variant

    Set this.Tests = New Dictionary
    Set this.SuiteMap = New Dictionary
    this.ScannedAssertionCount = 0
    ResetCounts

    Set proj = CurrentVBProject

    blnHasFolderTags = ProjectHasFolderAnnotations(proj)

    Set dModules = New Dictionary
    For Each cmp In proj.VBComponents
        If IsTestModule(cmp, blnHasFolderTags) Then
            dModules.Add cmp.Name, cmp
        End If
    Next cmp

    For Each varKey In dModules.Keys
        ScanModuleForTests dModules(varKey), blnHasFolderTags
    Next varKey

    SyncFactoryEntries proj

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SnapshotTestsByModuleProc
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Snapshot the current Tests dictionary keyed by Module.Proc for merge-scan.
'---------------------------------------------------------------------------------------
'
Private Function SnapshotTestsByModuleProc() As Dictionary

    Dim dPrior As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strMatchKey As String

    Set dPrior = New Dictionary
    If this.Tests Is Nothing Then
        Set SnapshotTestsByModuleProc = dPrior
        Exit Function
    End If

    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        strMatchKey = CStr(dTest("moduleName")) & "." & CStr(dTest("procName"))
        If Not dPrior.Exists(strMatchKey) Then
            Set dPrior(strMatchKey) = dTest
        End If
    Next varKey

    Set SnapshotTestsByModuleProc = dPrior

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergePriorResultsIntoDiscovered
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Copy prior pass/fail state onto freshly discovered tests when Module.Proc
'           : still matches.
'---------------------------------------------------------------------------------------
'
Private Sub MergePriorResultsIntoDiscovered(ByVal dPrior As Dictionary)

    Dim varKey As Variant
    Dim dNew As Dictionary
    Dim dOld As Dictionary
    Dim strMatchKey As String

    If dPrior Is Nothing Then Exit Sub
    If dPrior.Count = 0 Then Exit Sub

    For Each varKey In this.Tests.Keys
        Set dNew = this.Tests(CStr(varKey))
        strMatchKey = CStr(dNew("moduleName")) & "." & CStr(dNew("procName"))
        If dPrior.Exists(strMatchKey) Then
            Set dOld = dPrior(strMatchKey)
            MergePriorTestResult dNew, dOld
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergePriorTestResult
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Copy result fields from a prior test record onto a newly discovered one.
'---------------------------------------------------------------------------------------
'
Private Sub MergePriorTestResult(ByVal dNew As Dictionary, ByVal dOld As Dictionary)

    If CLng(Nz(dOld("status"), etsPending)) = etsPending Then Exit Sub

    dNew("status") = dOld("status")
    dNew("durationMs") = CLng(Nz(dOld("durationMs"), 0))

    If dOld.Exists("errorMessage") Then
        If Len(CStr(dOld("errorMessage"))) > 0 Then
            If dNew.Exists("errorMessage") Then
                dNew("errorMessage") = CStr(dOld("errorMessage"))
            Else
                dNew.Add "errorMessage", CStr(dOld("errorMessage"))
            End If
        End If
    End If

    If dOld.Exists("assertionResults") Then
        Set dNew("assertionResults") = CloneAssertionResults(dOld("assertionResults"))
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloneAssertionResults
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Deep-copy an assertion-results collection for merge-scan preservation.
'---------------------------------------------------------------------------------------
'
Private Function CloneAssertionResults(ByVal colSrc As Collection) As Collection

    Dim colOut As Collection
    Dim dA As Dictionary
    Dim dCopy As Dictionary
    Dim i As Long

    Set colOut = New Collection
    If colSrc Is Nothing Then
        Set CloneAssertionResults = colOut
        Exit Function
    End If

    For i = 1 To colSrc.Count
        Set dA = colSrc(i)
        Set dCopy = New Dictionary
        dCopy.Add "seq", CLng(Nz(dA("seq"), i))
        dCopy.Add "passed", CBool(Nz(dA("passed"), False))
        If dA.Exists("context") Then
            If Len(CStr(dA("context"))) > 0 Then
                dCopy.Add "context", CStr(dA("context"))
            End If
        End If
        If dA.Exists("sourceText") Then
            If Len(CStr(dA("sourceText"))) > 0 Then
                dCopy.Add "sourceText", CStr(dA("sourceText"))
            End If
        End If
        If dA.Exists("lineNumber") Then
            dCopy.Add "lineNumber", CLng(Nz(dA("lineNumber"), 0))
        End If
        colOut.Add dCopy
    Next i

    Set CloneAssertionResults = colOut

End Function


'---------------------------------------------------------------------------------------
' Procedure : RebuildCountsFromTests
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Recompute run summary counters from the current Tests dictionary.
'---------------------------------------------------------------------------------------
'
Private Sub RebuildCountsFromTests()

    Dim varKey As Variant
    Dim dTest As Dictionary

    ResetCounts
    If this.Tests Is Nothing Then Exit Sub

    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        AccumulateCountsFromTest dTest
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Initialize with empty collections.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set this.Tests = New Dictionary
    Set this.SuiteMap = New Dictionary
    Set this.AssertionResults = New Collection
End Sub
