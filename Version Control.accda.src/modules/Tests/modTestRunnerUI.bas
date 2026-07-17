Attribute VB_Name = "modTestRunnerUI"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")
'---------------------------------------------------------------------------------------
' Module    : modTestRunnerUI
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Web test-runner host helpers: Edge capability detection, HTML path
'           : resolution, VBA<->JS bridge dispatch, and streaming adapters for TestUI.
'---------------------------------------------------------------------------------------

Private Const ModuleName As String = "modTestRunnerUI"
Private Const ALLOWED_CALLBACKS As String = "RunAll,RunSelected,RunFailed,Cancel,OpenTestSource,ReportJsError,RefreshTests,OpenResultsReport"
Private Const JS_RETRIEVE_TIMEOUT_SENTINEL As String = "RetrieveJavascriptValue timed out"
Private Const MIN_EDGE_BUILD As Long = 16327
Private Const WEB_RUNNER_CACHE_FOLDER As String = "TestRunnerCache"

' Keep aligned with VBA_CALL_TIMEOUT_MS in runner.html (values cannot be shared
' across the JS/VBA boundary).
Public Const WEB_RUNNER_READY_TIMEOUT_MS As Long = 30000

Private m_blnWebRunnerActive As Boolean
Private m_blnDocumentReady As Boolean
Private m_blnTreePublished As Boolean
Private m_blnStandalone As Boolean      ' opened to view last results (not a fresh run)
Private m_strPendingTreeJson As String
Private m_strPendingDefaultFilter As String
Private m_varEdgeBuildCached As Variant
Private m_strHtmlCacheFolder As String
Private m_curLastCancelPoll As Currency

' Throttle mid-run Cancel reads (PollBridgeCancel is called once per test).
Private Const CANCEL_POLL_MS As Long = 1000

' Run command accepted (ack sent to JS) and awaiting execution on the timer stack.
Private m_strPendingRunFn As String
Private m_colPendingRunKeys As Collection
Private m_blnPendingRunSetup As Boolean ' a new Operation was begun; invoke global setup hook
Private m_eimPriorMode As eInteractionMode ' cached before bridge run; restored on teardown


'---------------------------------------------------------------------------------------
' Procedure : EdgeTestRunnerSupported
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : True when this Access build can host the EdgeBrowserControl (M365 2304+).
'---------------------------------------------------------------------------------------
'
Public Function EdgeTestRunnerSupported() As Boolean

    Dim lngBuild As Long

    If IsEmpty(m_varEdgeBuildCached) Then
        lngBuild = GetAccessFileBuild()
        m_varEdgeBuildCached = (lngBuild = 0 Or lngBuild >= MIN_EDGE_BUILD)
    End If

    EdgeTestRunnerSupported = CBool(m_varEdgeBuildCached)

End Function


'---------------------------------------------------------------------------------------
' Procedure : WebRunnerActive
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : True while the web test runner form is hosting a run.
'---------------------------------------------------------------------------------------
'
Public Function WebRunnerActive() As Boolean
    WebRunnerActive = m_blnWebRunnerActive And IsLoaded(acForm, "frmVCSTestRunner", False)
End Function


'---------------------------------------------------------------------------------------
' Procedure : WebRunnerDocumentReady
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : True after the Edge document has finished loading.
'---------------------------------------------------------------------------------------
'
Public Function WebRunnerDocumentReady() As Boolean
    WebRunnerDocumentReady = m_blnDocumentReady
End Function


'---------------------------------------------------------------------------------------
' Procedure : PollBridgeCancel
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Drain Cancel commands from the JS outbox during a test run. The form
'           : timer is suspended while a run started from DrainOutbox is on the stack,
'           : so this is the sole mid-run cancel path. Throttled to CANCEL_POLL_MS so
'           : fast test bursts do not issue one RetrieveJavascriptValue per test.
'---------------------------------------------------------------------------------------
'
Public Sub PollBridgeCancel()

    Dim frm As Object
    Dim dblElapsedMs As Double

    If Not WebRunnerActive() Then Exit Sub
    If Not WebRunnerDocumentReady() Then Exit Sub

    If m_curLastCancelPoll <> 0 Then
        dblElapsedMs = (Perf.MicroTimer - m_curLastCancelPoll) * 1000
        If dblElapsedMs >= 0 And dblElapsedMs < CANCEL_POLL_MS Then Exit Sub
    End If

    m_curLastCancelPoll = Perf.MicroTimer
    Set frm = RunnerForm()
    If frm Is Nothing Then Exit Sub
    frm.DrainCancelOutbox

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RefocusWebRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Return OS focus to Access and the web test runner after compile or
'           : other VBE operations that stole the foreground window.
'---------------------------------------------------------------------------------------
'
Public Sub RefocusWebRunner()

    Dim frm As Object

    If Not WebRunnerActive() Then Exit Sub
    Set frm = RunnerForm()
    If frm Is Nothing Then Exit Sub
    If Not frm.Visible Then Exit Sub

    modVCSUtility.BringAccessToForeground
    frm.ShowRunner

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunnerForm
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Return the OPEN frmVCSTestRunner instance from the Forms collection (the
'           : one DoCmd.OpenForm created), not the Form_frmVCSTestRunner default auto-
'           : instance. The default instance can differ from the opened one, and its
'           : Form_Load may never have bound the Edge control (Nothing), so always talk
'           : to the loaded instance. Returns Nothing when the form is not open.
'---------------------------------------------------------------------------------------
'
Private Function RunnerForm() As Object
    If IsLoaded(acForm, "frmVCSTestRunner", False) Then
        Set RunnerForm = Forms("frmVCSTestRunner")
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenWebTestRunner
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Open (or reuse) frmVCSTestRunner and navigate to runner HTML. The form
'           : hides instead of unloading when the user clicks X or presses Escape,
'           : keeping the WebView2 control warm for the next open. ShowRunner re-enables
'           : the poll timer that was disabled while hidden.
'---------------------------------------------------------------------------------------
'
Public Sub OpenWebTestRunner()

    Dim frm As Object

    DoCmd.Hourglass True
    DoEvents

    ' Clear any leftover test-run Operation from a previous hide/cancel so reopen
    ' (and the next Run click) are not blocked by "Another Operation Already Running".
    ClearOrphanedTestOperation

    m_blnWebRunnerActive = True
    m_blnStandalone = False

    Set frm = RunnerForm()
    modTestRunnerDiag.DiagStart "OpenWebTestRunner reuse=" & (Not frm Is Nothing) & _
        " edgeBuild>=min=" & EdgeTestRunnerSupported()

    If frm Is Nothing Then
        m_blnDocumentReady = False
        m_blnTreePublished = False
        m_strPendingTreeJson = vbNullString
        modTestRunnerDiag.Diag "open.new"
        DoCmd.OpenForm "frmVCSTestRunner", , , , , acNormal
        ' OpenForm resets the hourglass; keep it on until the Edge page is ready.
        DoCmd.Hourglass True
    Else
        ReuseOrReloadRunner frm, False
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedTestOperation
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Finish a leftover eotTestRun Operation when no test run is actually in
'           : progress. Hide-on-close can leave Operation.Status = eosRunning if the
'           : user closed mid-run (Cancel is cooperative) or if an earlier open path
'           : began an Operation without finishing it.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedTestOperation()

    If TestRunner.State = etrsRunning Then Exit Sub
    If Operation.Status <> eosRunning Then Exit Sub
    If Operation.OperationType <> eotTestRun Then Exit Sub

    modTestRunnerDiag.Diag "op.clear_orphan", "state=" & TestRunner.State
    Log.SuppressDebugOutput = False
    Operation.InteractionMode = m_eimPriorMode
    Operation.Finish eorCanceled

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenTestRunnerForResults
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Open (or reuse) the web runner to VIEW the last run's results without
'           : running anything. On document-ready the host rehydrates the UI from the
'           : TestRunner singleton (see NotifyDocumentReady -> RehydrateWebRunner), so a
'           : reopened runner shows the prior results and its failed set to re-run.
'---------------------------------------------------------------------------------------
'
Public Sub OpenTestRunnerForResults()

    Dim frm As Object

    DoCmd.Hourglass True
    DoEvents

    ClearOrphanedTestOperation

    m_blnWebRunnerActive = True
    m_blnStandalone = True
    m_strPendingDefaultFilter = vbNullString

    Set frm = RunnerForm()
    modTestRunnerDiag.DiagStart "OpenTestRunnerForResults reuse=" & (Not frm Is Nothing)

    If frm Is Nothing Then
        m_blnDocumentReady = False
        m_blnTreePublished = False
        m_strPendingTreeJson = vbNullString
        DoCmd.OpenForm "frmVCSTestRunner", , , , , acNormal
        ' OpenForm resets the hourglass; keep it on until the Edge page is ready.
        DoCmd.Hourglass True
    Else
        ReuseOrReloadRunner frm, True
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RehydrateWebRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Repopulate the web UI from the TestRunner singleton's last results
'           : (tree + each completed test), so a reopened runner shows prior state.
'           : No-op when there are no discovered tests/results.
'---------------------------------------------------------------------------------------
'
Public Sub RehydrateWebRunner()

    If Not EnsureRunnerHasTests() Then Exit Sub

    modTestRunnerDiag.Diag "rehydrate", "tests=" & TestRunner.Tests.Count
    PublishTestTree TestRunner.GetTestTreeAsJson()
    StreamCompletedTestResults

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseWebTestRunner
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Close the web runner form and reset host state.
'---------------------------------------------------------------------------------------
'
Public Sub CloseWebTestRunner()

    Dim frm As Object

    ResetWebRunnerHostState
    On Error Resume Next
    Set frm = RunnerForm()
    If Not frm Is Nothing Then
        frm.AllowClose = True
        DoCmd.Close acForm, "frmVCSTestRunner", acSaveNo
    End If
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResetWebRunnerHostState
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Clear host flags without closing the form (used from Form_Unload).
'---------------------------------------------------------------------------------------
'
Public Sub ResetWebRunnerHostState()
    DoCmd.Hourglass False
    m_blnWebRunnerActive = False
    m_blnDocumentReady = False
    m_blnTreePublished = False
    m_strPendingTreeJson = vbNullString
    m_strPendingDefaultFilter = vbNullString
    m_blnStandalone = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetWebRunnerDefaultFilter
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Seed the web test runner filter box when opening from VCS.RunTests or the
'           : ribbon (DefaultTestFilter). Cleared when the filter string is empty.
'---------------------------------------------------------------------------------------
'
Public Sub SetWebRunnerDefaultFilter(ByVal strFilter As String)
    m_strPendingDefaultFilter = Trim$(strFilter)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PublishTestTree
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Queue or push the pre-run suite tree JSON to the web UI.
'---------------------------------------------------------------------------------------
'
Public Sub PublishTestTree(ByVal strTreeJson As String)

    m_strPendingTreeJson = strTreeJson
    TryPublishPendingTree

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WaitForWebRunnerReady
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Pump messages until the Edge document is ready (or timeout). The default
'           : is generous because a cold WebView2 first-init (first open after launch)
'           : can take well over 15s; starting the run before the page is ready streams
'           : results into a blank page.
'---------------------------------------------------------------------------------------
'
Public Sub WaitForWebRunnerReady(Optional ByVal lngTimeoutMs As Long = WEB_RUNNER_READY_TIMEOUT_MS)

    Dim curStart As Currency
    Dim curLimit As Currency

    curStart = Perf.MicroTimer
    curLimit = curStart + (lngTimeoutMs / 1000#)

    modTestRunnerDiag.Diag "wait.begin", "timeoutMs=" & lngTimeoutMs

    Do While Not m_blnDocumentReady
        ' Access clears the hourglass on OpenForm / DoEvents; re-assert while waiting.
        DoCmd.Hourglass True
        DoEvents
        If Perf.MicroTimer > curLimit Then Exit Do
        If Not IsLoaded(acForm, "frmVCSTestRunner", False) Then Exit Do
    Loop

    If m_blnDocumentReady Then
        modTestRunnerDiag.Diag "wait.ready", _
            "afterMs=" & CLng((Perf.MicroTimer - curStart) * 1000)
    Else
        modTestRunnerDiag.Diag "wait.timeout", _
            "afterMs=" & CLng((Perf.MicroTimer - curStart) * 1000) & _
            " loaded=" & IsLoaded(acForm, "frmVCSTestRunner", False)
    End If

    TryPublishPendingTree
    DoCmd.Hourglass False

End Sub


'---------------------------------------------------------------------------------------
' Procedure : NotifyDocumentReady
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Called from frmVCSTestRunner when DocumentComplete fires.
'---------------------------------------------------------------------------------------
'
Public Sub NotifyDocumentReady()
    modTestRunnerDiag.Diag "documentcomplete"
    m_blnDocumentReady = True
    ' Every DocumentComplete is a FRESH page (initial load OR a spurious reload -- the
    ' Edge control can re-fire DocumentComplete without our Navigate). The retained tree
    ' JSON must be (re)sent, so reset the published flag; otherwise a reload leaves the
    ' page with an empty testTree and every later onTestComplete has no row to update.
    m_blnTreePublished = False
    ConnectBridgeInPage
    TryPublishPendingTree
    RefreshWebTestTreeDeferred
    If HasCompletedTests() Then StreamCompletedTestResults
    DoCmd.Hourglass False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NotifyDocumentLost
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Reset readiness when the browser navigates away or the form closes.
'---------------------------------------------------------------------------------------
'
Public Sub NotifyDocumentLost()
    m_blnDocumentReady = False
    m_blnTreePublished = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamRunStart
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push TestUI.onRunStart to the web view.
'---------------------------------------------------------------------------------------
'
Public Sub StreamRunStart(ByVal colTestKeys As Collection)

    Dim dPayload As Dictionary
    Dim colKeysOut As Collection
    Dim varKey As Variant

    If Not WebRunnerReady() Then Exit Sub

    Set dPayload = New Dictionary
    Set colKeysOut = New Collection
    For Each varKey In colTestKeys
        colKeysOut.Add CStr(varKey)
    Next varKey
    dPayload.Add "totalTests", colKeysOut.Count
    Set dPayload("testKeys") = colKeysOut

    PushTestUI "onRunStart", ConvertToJson(dPayload)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamTestStart
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push TestUI.onTestStart for a single test.
'---------------------------------------------------------------------------------------
'
Public Sub StreamTestStart(ByVal strTestKey As String, ByVal dTest As Dictionary)

    Dim dPayload As Dictionary

    If Not WebRunnerReady() Then Exit Sub

    Set dPayload = New Dictionary
    dPayload.Add "key", strTestKey
    dPayload.Add "name", CStr(dTest("procName"))
    dPayload.Add "module", CStr(dTest("moduleName"))
    dPayload.Add "procName", CStr(dTest("procName"))

    PushTestUI "onTestStart", ConvertToJson(dPayload)
    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamTestComplete
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push TestUI.onTestComplete for a single finished test.
'---------------------------------------------------------------------------------------
'
Public Sub StreamTestComplete(ByVal strTestKey As String, ByVal dTest As Dictionary)

    If Not WebRunnerReady() Then Exit Sub
    PushTestUI "onTestComplete", AdaptTestResultJson(strTestKey, dTest)
    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamRunComplete
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push TestUI.onRunComplete after the runner loop finishes.
'---------------------------------------------------------------------------------------
'
Public Sub StreamRunComplete(Optional ByVal lngTotalMs As Long = 0)

    Dim dPayload As Dictionary

    If Not WebRunnerReady() Then Exit Sub

    Set dPayload = New Dictionary
    dPayload.Add "totalMs", lngTotalMs
    PushTestUI "onRunComplete", ConvertToJson(dPayload)
    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamRunCancelled
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push TestUI.onRunCancelled when the user stops a run.
'---------------------------------------------------------------------------------------
'
Public Sub StreamRunCancelled()

    If Not WebRunnerReady() Then Exit Sub
    PushTestUI "onRunCancelled", "null"
    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StreamRunError
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Push TestUI.onRunError when a run fails before or outside the normal
'           : completion path (compile errors, failures after the acceptance ack).
'           : No-ops in console mode.
'---------------------------------------------------------------------------------------
'
Public Sub StreamRunError(ByVal strMessage As String)

    Dim dPayload As Dictionary

    If Not WebRunnerReady() Then Exit Sub

    Set dPayload = New Dictionary
    dPayload.Add "message", strMessage
    PushTestUI "onRunError", ConvertToJson(dPayload)
    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PushTestUI
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : ExecuteJavascript wrapper for TestUI event handlers.
'---------------------------------------------------------------------------------------
'
Public Sub PushTestUI(ByVal strHandler As String, ByVal strJsonPayload As String)

    Dim frm As Object

    On Error Resume Next
    If Not WebRunnerReady() Then
        modTestRunnerDiag.Diag "push.dropped", strHandler & " (runner not ready)"
        Exit Sub
    End If
    Set frm = RunnerForm()
    If Not frm Is Nothing Then
        modTestRunnerDiag.Diag "push", strHandler
        frm.ExecuteRunnerScript "window.TestUI." & strHandler & "(" & strJsonPayload & ")"
    Else
        modTestRunnerDiag.Diag "push.noform", strHandler
    End If
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveRunnerHtmlPath
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Return a local path to the runner HTML. The HTML is not co-located with
'           : the compiled add-in, so it is delivered as an embedded resource
'           : (tblResources) and extracted to a stable subfolder under the add-in
'           : install path (CodeProject.Path\TestRunnerCache\). A dev fallback copies
'           : straight from the source tree when the resource has not been embedded
'           : yet. A stable path keeps the https://msaccess/ navigate URL constant
'           : across sessions so WebView2 localStorage (Recent filters, column widths)
'           : persists. ResolveRunnerNavigateUrl applies GetShortPath when the path
'           : contains spaces.
'---------------------------------------------------------------------------------------
'
Public Function ResolveRunnerHtmlPath() As String

    Const FunctionName As String = ModuleName & ".ResolveRunnerHtmlPath"

    Dim strFileName As String
    Dim strKey As String
    Dim strTarget As String
    Dim strSource As String

    strFileName = "runner.html"
    strKey = "Test Runner HTML"

    ' Stable extraction folder under the add-in install path (constant navigate URL).
    If Len(m_strHtmlCacheFolder) = 0 Or Not FSO.FolderExists(m_strHtmlCacheFolder) Then
        m_strHtmlCacheFolder = CodeProject.Path & PathSep & WEB_RUNNER_CACHE_FOLDER & PathSep
        VerifyPath m_strHtmlCacheFolder & "placeholder"
    End If
    strTarget = m_strHtmlCacheFolder & strFileName

    ' Preferred: extract the embedded resource (works for installed/compiled add-in)
    If modResource.GetResourceHash(strKey) <> vbNullString Then
        modResource.ExtractResource strKey, m_strHtmlCacheFolder
    End If

    ' Locate the on-disk source (repo-root TestRunner/, parallel to Ribbon/) for
    ' fallback / live-edit refresh when building from source.
    strSource = CodeProject.Path & PathSep & "TestRunner" & PathSep & strFileName

    ' Copy from source when the resource is missing, or when the source is newer
    ' (lets a developer iterate on runner.html without rebuilding the add-in).
    If FSO.FileExists(strSource) Then
        If Not FSO.FileExists(strTarget) Then
            FSO.CopyFile strSource, strTarget, True
        ElseIf FSO.GetFile(strSource).DateLastModified > FSO.GetFile(strTarget).DateLastModified Then
            FSO.CopyFile strSource, strTarget, True
        End If
    End If

    If Not FSO.FileExists(strTarget) Then
        Log.Add T("Test runner HTML not found (resource '{0}' and source both missing).", _
            var0:=strKey), , , "red"
        Exit Function
    End If

    ResolveRunnerHtmlPath = strTarget

    CatchAny eelWarning, vbNullString, FunctionName

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveRunnerNavigateUrl
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Build the https://msaccess/ URL required by EdgeBrowserControl local files.
'---------------------------------------------------------------------------------------
'
Public Function ResolveRunnerNavigateUrl() As String

    Dim strPath As String

    strPath = ResolveRunnerHtmlPath()
    If Len(strPath) = 0 Then Exit Function

    ' The Edge control silently fails to load https://msaccess/ URLs containing
    ' spaces. Convert to the 8.3 short path when the extracted path has one (e.g.,
    ' a Windows profile folder with a space).
    If InStr(1, strPath, " ", vbBinaryCompare) > 0 Then strPath = GetShortPath(strPath)

    ResolveRunnerNavigateUrl = "https://msaccess/" & Replace(strPath, PathSep, "/")
    modTestRunnerDiag.Diag "navigate.url", ResolveRunnerNavigateUrl

End Function


'---------------------------------------------------------------------------------------
' Procedure : DispatchBridgeCall
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Allowlisted inbound bridge entry point for JS->VBA calls.
'---------------------------------------------------------------------------------------
'
Public Function DispatchBridgeCall(ByVal strFnName As String, ByVal strRequestId As String, _
    ByVal strPayloadJson As String) As String

    Const FunctionName As String = ModuleName & ".DispatchBridgeCall"

    modTestRunnerDiag.Diag "dispatch.begin", strFnName & " id=" & strRequestId

    If Not IsAllowedBridgeCallback(strFnName) Then
        modTestRunnerDiag.Diag "dispatch.disallowed", strFnName
        Err.Raise vbObjectError + 513, FunctionName, "Unknown or disallowed function: " & strFnName
    End If

    ' Run commands (RunAll/RunSelected/RunFailed) are dispatched via AcceptBridgeRun /
    ' ExecutePendingBridgeRun so the JS promise is resolved at acceptance, not after
    ' the (long, blocking) run completes. See frmVCSTestRunner.DispatchRequest.
    Select Case strFnName
        Case "Cancel"
            DispatchBridgeCall = BridgeCancel(strPayloadJson)
        Case "OpenTestSource"
            DispatchBridgeCall = BridgeOpenTestSource(strPayloadJson)
        Case "ReportJsError"
            DispatchBridgeCall = BridgeReportJsError(strPayloadJson)
        Case "RefreshTests"
            DispatchBridgeCall = BridgeRefreshTests(strPayloadJson)
        Case "OpenResultsReport"
            DispatchBridgeCall = BridgeOpenResultsReport(strPayloadJson)
        Case Else
            Err.Raise vbObjectError + 513, FunctionName, "Unhandled bridge callback: " & strFnName
    End Select

    modTestRunnerDiag.Diag "dispatch.end", strFnName & " id=" & strRequestId

End Function


'---------------------------------------------------------------------------------------
' Procedure : BridgeReportJsError
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Log a JS-side UI error reported from runner.html (fire-and-forget).
'---------------------------------------------------------------------------------------
'
Public Function BridgeReportJsError(ByVal strPayloadJson As String) As String

    Dim dPayload As Dictionary
    Dim strMsg As String
    Dim strDetail As String

    Set dPayload = ParseJson(IIf(Len(strPayloadJson) > 0, strPayloadJson, "{}"))
    strMsg = vbNullString
    strDetail = vbNullString

    If TypeName(dPayload) = "Dictionary" Then
        If dPayload.Exists("message") Then strMsg = CStr(dPayload("message"))
        If dPayload.Exists("detail") Then strDetail = CStr(dPayload("detail"))
    End If

    If Len(strMsg) = 0 Then strMsg = "(no message)"
    If Len(strDetail) > 0 Then strMsg = strMsg & " | " & strDetail

    modTestRunnerDiag.Diag "js.ui.error", strMsg
    Log.Add strMsg, , , "orange"
    BridgeReportJsError = "{""ok"":true}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsRunCommand
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : True for bridge commands that start a (long, blocking) test run. These
'           : are acknowledged at acceptance rather than resolved after completion.
'---------------------------------------------------------------------------------------
'
Public Function IsRunCommand(ByVal strFnName As String) As Boolean
    Select Case strFnName
        Case "RunAll", "RunSelected", "RunFailed"
            IsRunCommand = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : AcceptBridgeRun
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Validate and accept a run command from the web UI, returning the ack JSON
'           : that resolves the JS promise BEFORE the run executes. Raises (-> promise
'           : reject) on validation failure. Only the fast part of run startup happens
'           : here; the global test setup hook (unbounded user code) is deferred to
'           : ExecutePendingBridgeRun so the ack cannot time out.
'---------------------------------------------------------------------------------------
'
Public Function AcceptBridgeRun(ByVal strFnName As String, ByVal strPayloadJson As String) As String

    Const FunctionName As String = ModuleName & ".AcceptBridgeRun"

    Dim dPayload As Dictionary
    Dim colKeys As Collection
    Dim varItem As Variant

    If TestRunner.State = etrsRunning Then
        Err.Raise vbObjectError + 515, FunctionName, T("Tests are already running")
    End If

    ' Validate the request before accepting it
    Select Case strFnName
        Case "RunSelected"
            Set dPayload = ParseJson(IIf(Len(strPayloadJson) > 0, strPayloadJson, "{}"))
            Set colKeys = New Collection
            If TypeName(dPayload) = "Dictionary" Then
                If dPayload.Exists("testKeys") Then
                    If TypeName(dPayload("testKeys")) = "Collection" Then
                        For Each varItem In dPayload("testKeys")
                            colKeys.Add CStr(varItem)
                        Next varItem
                    End If
                End If
            End If
            If colKeys.Count = 0 Then
                Err.Raise vbObjectError + 514, FunctionName, T("No test keys supplied")
            End If
        Case "RunFailed"
            If Not TestRunner.HasFailedTests Then
                Err.Raise vbObjectError + 514, FunctionName, T("No failed tests to run")
            End If
    End Select

    ' Suppress Immediate window output while the web runner hosts the run.
    Log.SuppressDebugOutput = True

    ClearOrphanedTestOperation

    m_blnPendingRunSetup = False
    If Operation.Status <> eosRunning Then
        If Not Operation.Begin(eotTestRun) Then
            Log.SuppressDebugOutput = False
            Err.Raise vbObjectError + 515, FunctionName, T("Could not begin test operation")
        End If
        Log.Active = True
        m_eimPriorMode = Operation.InteractionMode
        Operation.InteractionMode = eimSilent
        Log.ClearErrorJournal
        m_blnPendingRunSetup = True
    End If

    m_strPendingRunFn = strFnName
    Set m_colPendingRunKeys = colKeys

    modTestRunnerDiag.Diag "run.accept", strFnName
    m_curLastCancelPoll = 0
    AcceptBridgeRun = "{""ok"":true,""accepted"":true}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExecutePendingBridgeRun
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Execute the run accepted by AcceptBridgeRun. The JS promise was already
'           : resolved with the acceptance ack, so failures from here on are streamed
'           : to the page as onRunError (never raised back through the bridge), and
'           : teardown is guaranteed so the next run is not blocked by an orphaned
'           : operation.
'---------------------------------------------------------------------------------------
'
Public Sub ExecutePendingBridgeRun()

    Const FunctionName As String = ModuleName & ".ExecutePendingBridgeRun"

    Dim strFnName As String
    Dim colKeys As Collection
    Dim blnInvokeSetup As Boolean
    Dim strErrMsg As String

    If Len(m_strPendingRunFn) = 0 Then Exit Sub

    ' Copy and clear the pending state up front so a failure cannot leave a stale run.
    strFnName = m_strPendingRunFn
    Set colKeys = m_colPendingRunKeys
    blnInvokeSetup = m_blnPendingRunSetup
    m_strPendingRunFn = vbNullString
    Set m_colPendingRunKeys = Nothing
    m_blnPendingRunSetup = False

    LogUnhandledErrors
    On Error GoTo ErrHandler

    If blnInvokeSetup Then TestRunner.InvokeGlobalTestSetup

    Select Case strFnName
        Case "RunAll"
            TestRunner.RunAll
        Case "RunSelected"
            TestRunner.RunSelected colKeys
        Case "RunFailed"
            TestRunner.RunFailed
    End Select

    EndInteractiveBridgeRun
    Exit Sub

ErrHandler:
    strErrMsg = Err.Description
    CatchAny eelError, T("Test run failed"), FunctionName
    On Error Resume Next
    modTestRunnerDiag.Diag "run.execute.error", strFnName & " | " & strErrMsg
    StreamRunError strErrMsg
    If Operation.Status = eosRunning Then Operation.Finish eorFailed
    Log.SuppressDebugOutput = False
    Operation.InteractionMode = m_eimPriorMode
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BridgeCancel
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Cancel the active test run from the web UI.
'---------------------------------------------------------------------------------------
'
Public Function BridgeCancel(ByVal strPayloadJson As String) As String

    If (Operation.Status = eosRunning And Operation.OperationType = eotTestRun) _
        Or TestRunner.State = etrsRunning Then
        TestRunner.Cancel
    End If
    BridgeCancel = "{""ok"":true}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : BridgeRefreshTests
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Rediscover tests and republish the tree without reloading the page.
'---------------------------------------------------------------------------------------
'
Public Function BridgeRefreshTests(ByVal strPayloadJson As String) As String

    If TestRunner.State = etrsRunning Then
        Err.Raise vbObjectError + 518, ModuleName & ".BridgeRefreshTests", T("Tests are already running")
    End If

    RefreshWebTestTreeDeferred
    BridgeRefreshTests = "{""ok"":true,""testCount"":" & TestRunner.Tests.Count & "}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : BridgeOpenResultsReport
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Regenerate test-results.html and open it in the default browser.
'---------------------------------------------------------------------------------------
'
Public Function BridgeOpenResultsReport(ByVal strPayloadJson As String) As String

    Dim strPath As String

    strPath = modTestReport.ExportResultsHtml()
    If Len(strPath) = 0 Then
        Err.Raise vbObjectError + 519, ModuleName & ".BridgeOpenResultsReport", _
            T("No durable test state file was found to export.")
    End If

    Application.FollowHyperlink strPath
    BridgeOpenResultsReport = "{""ok"":true,""path"":""" & Replace(strPath, "\", "\\") & """}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : BridgeOpenTestSource
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Navigate the VBE to a test module/procedure (optional line).
'---------------------------------------------------------------------------------------
'
Public Function BridgeOpenTestSource(ByVal strPayloadJson As String) As String

    Dim dPayload As Dictionary
    Dim strModule As String
    Dim strProc As String
    Dim lngLine As Long
    Dim cmp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule

    Set dPayload = ParseJson(IIf(Len(strPayloadJson) > 0, strPayloadJson, "{}"))
    strModule = CStr(Nz(dPayload("module"), vbNullString))
    strProc = CStr(Nz(dPayload("procName"), vbNullString))
    lngLine = CLng(Nz(dPayload("lineNumber"), 0))

    If Len(strModule) = 0 Then
        Err.Raise vbObjectError + 516, ModuleName & ".BridgeOpenTestSource", "Module name required"
    End If

    Set VBE.ActiveVBProject = CurrentVBProject
    Set cmp = CurrentVBProject.VBComponents(strModule)
    Set cm = cmp.CodeModule

    ' Resolve the line to place the cursor on: scan the code for the procedure's
    ' declaration line (lands exactly on the "Sub/Function" statement, not the leading
    ' comment block that ProcStartLine returns). Fall back to ProcBodyLine, then any
    ' caller-supplied line, then the top of the module.
    If Len(strProc) > 0 Then
        lngLine = FindProcedureLine(cm, strProc)
    End If
    If lngLine <= 0 Then lngLine = 1

    Application.VBE.MainWindow.Visible = True
    cm.CodePane.Show
    cm.CodePane.SetSelection lngLine, 1, lngLine, 1
    cm.CodePane.Show   ' second Show scrolls the selection into view

    BridgeOpenTestSource = "{""ok"":true}"

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindProcedureLine
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Return the 1-based line of a procedure's declaration by scanning the code
'           : module's lines (as requested). Matches an optional scope/Static prefix
'           : followed by Sub / Function / Property Get|Let|Set and the procedure name.
'           : Falls back to the VBE's ProcBodyLine if no line matches.
'---------------------------------------------------------------------------------------
'
Private Function FindProcedureLine(ByVal cm As VBIDE.CodeModule, ByVal strProc As String) As Long

    Dim i As Long
    Dim strUpper As String
    Dim strName As String

    strName = UCase$(Trim$(strProc))
    If Len(strName) = 0 Then Exit Function

    For i = 1 To cm.CountOfLines
        strUpper = UCase$(Trim$(cm.Lines(i, 1)))
        If IsProcDeclarationLine(strUpper, strName) Then
            FindProcedureLine = i
            Exit Function
        End If
    Next i

    ' Fallback: the VBE's own body-line lookup.
    On Error Resume Next
    FindProcedureLine = cm.ProcBodyLine(strProc, vbext_pk_Proc)
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsProcDeclarationLine
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : True if strUpper (already UCase+Trim'd) is a declaration of procedure
'           : strName (already UCase'd). Handles Public/Private/Friend/Static prefixes
'           : and Sub/Function/Property Get|Let|Set; the name must be followed by "(",
'           : a space, or end of line (so "TestFoo" does not match "TestFooBar").
'---------------------------------------------------------------------------------------
'
Private Function IsProcDeclarationLine(ByVal strUpper As String, ByVal strName As String) As Boolean

    Dim s As String
    Dim varKind As Variant
    Dim strRest As String
    Dim strKind As String

    s = strUpper
    Do
        If Left$(s, 7) = "PUBLIC " Then
            s = Trim$(Mid$(s, 8))
        ElseIf Left$(s, 8) = "PRIVATE " Then
            s = Trim$(Mid$(s, 9))
        ElseIf Left$(s, 7) = "FRIEND " Then
            s = Trim$(Mid$(s, 8))
        ElseIf Left$(s, 7) = "STATIC " Then
            s = Trim$(Mid$(s, 8))
        Else
            Exit Do
        End If
    Loop

    For Each varKind In Array("SUB ", "FUNCTION ", "PROPERTY GET ", "PROPERTY LET ", "PROPERTY SET ")
        strKind = CStr(varKind)
        If Left$(s, Len(strKind)) = strKind Then
            strRest = Trim$(Mid$(s, Len(strKind) + 1))
            If strRest = strName Then
                IsProcDeclarationLine = True
            ElseIf Left$(strRest, Len(strName) + 1) = strName & "(" Then
                IsProcDeclarationLine = True
            ElseIf Left$(strRest, Len(strName) + 1) = strName & " " Then
                IsProcDeclarationLine = True
            End If
            Exit Function
        End If
    Next varKind

End Function


'---------------------------------------------------------------------------------------
' Procedure : AdaptTestResultJson
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Map a clsTestRunner test dictionary to TestUI.onTestComplete JSON.
'---------------------------------------------------------------------------------------
'
Public Function AdaptTestResultJson(ByVal strTestKey As String, ByVal dTest As Dictionary) As String
    AdaptTestResultJson = ConvertToJson(AdaptTestResult(strTestKey, dTest))
End Function


'---------------------------------------------------------------------------------------
' Procedure : AdaptTestResult
' Author    : Adam Waller
' Date      : 7/10/2026
' Purpose   : Build the web-shaped result Dictionary for a single test (testKey, status,
'           : durationMs, optional errorMessage, assertions). Returned as an object so
'           : callers can either serialize it alone (AdaptTestResultJson) or collect many
'           : and serialize once for a batch replay (StreamCompletedTestResults).
'---------------------------------------------------------------------------------------
'
Private Function AdaptTestResult(ByVal strTestKey As String, ByVal dTest As Dictionary) As Dictionary

    Dim dOut As Dictionary
    Dim colAssertions As Collection
    Dim colOut As Collection
    Dim dA As Dictionary
    Dim dAOut As Dictionary
    Dim i As Long

    Set dOut = New Dictionary
    dOut.Add "testKey", strTestKey
    dOut.Add "status", WebStatusFromRunnerStatus(CLng(dTest("status")))
    dOut.Add "durationMs", CLng(Nz(dTest("durationMs"), 0))

    If dTest.Exists("errorMessage") Then
        If Len(CStr(dTest("errorMessage"))) > 0 Then
            dOut.Add "errorMessage", CStr(dTest("errorMessage"))
        End If
    End If

    Set colOut = New Collection
    If dTest.Exists("assertionResults") Then
        Set colAssertions = dTest("assertionResults")
        For i = 1 To colAssertions.Count
            Set dA = colAssertions(i)
            Set dAOut = New Dictionary
            dAOut.Add "seq", dA("seq")
            dAOut.Add "passed", CBool(dA("passed"))
            If Len(CStr(Nz(dA("context"), vbNullString))) > 0 Then
                dAOut.Add "context", CStr(dA("context"))
            End If
            colOut.Add dAOut
        Next i
    End If
    Set dOut("assertions") = colOut

    Set AdaptTestResult = dOut

End Function


'---------------------------------------------------------------------------------------
' Procedure : RetrieveJsValue
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : RetrieveJavascriptValue wrapper with timeout sentinel detection.
'---------------------------------------------------------------------------------------
'
Public Function RetrieveJsValue(ByVal ctl As Object, ByVal strExpression As String) As String

    Dim strResult As String
    Dim intAttempt As Integer

    ' RetrieveJavascriptValue is intermittently slow / hits its internal timeout in
    ' the Edge control (confirmed in the diagnostic trace). Retry once after pumping
    ' messages, which usually clears the transient case. Each attempt is timed via Perf
    ' ("RetrieveJavascriptValue") so its cost shows in the run's performance report.
    For intAttempt = 1 To 2
        Perf.OperationStart "RetrieveJavascriptValue"
        strResult = CStr(ctl.RetrieveJavascriptValue(strExpression))
        Perf.OperationEnd
        If InStr(1, strResult, JS_RETRIEVE_TIMEOUT_SENTINEL, vbTextCompare) = 0 Then
            RetrieveJsValue = strResult
            Exit Function
        End If
        DoEvents
    Next intAttempt

    ' Both attempts timed out.
    Err.Raise vbObjectError + 517, ModuleName & ".RetrieveJsValue", strResult

End Function


Private Function WebRunnerReady() As Boolean
    WebRunnerReady = WebRunnerActive() And m_blnDocumentReady
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReuseOrReloadRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Show a hidden warm form and skip HTML reload when the page is still healthy.
'---------------------------------------------------------------------------------------
'
Private Sub ReuseOrReloadRunner(ByVal frm As Object, ByVal blnStandalone As Boolean)

    m_blnStandalone = blnStandalone
    modTestRunnerDiag.Diag "open.reuse"

    On Error Resume Next
    frm.ShowRunner
    Err.Clear
    On Error GoTo 0

    If WebRunnerPageHealthy() Then
        modTestRunnerDiag.Diag "open.reuse.warm"
        DoCmd.Hourglass False
        RefreshWebTestTreeDeferred
        If blnStandalone And HasCompletedTests() Then StreamCompletedTestResults
    Else
        modTestRunnerDiag.Diag "open.reuse.reload"
        m_blnDocumentReady = False
        m_blnTreePublished = False
        frm.ReloadRunnerHtml
        DoCmd.Hourglass True
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WebRunnerPageHealthy
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : True when the warm WebView2 page is still loaded and TestUI is present.
'---------------------------------------------------------------------------------------
'
Private Function WebRunnerPageHealthy() As Boolean

    Dim frm As Object
    Dim strResult As String

    If Not m_blnDocumentReady Then Exit Function
    If Not WebRunnerActive() Then Exit Function

    On Error Resume Next
    Set frm = RunnerForm()
    If frm Is Nothing Then Exit Function
    strResult = frm.RetrieveRunnerJsValue("typeof window.TestUI")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    WebRunnerPageHealthy = (strResult = "object")

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefreshWebTestTreeDeferred
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Paint the warm UI first, then merge-scan and republish the tree only.
'---------------------------------------------------------------------------------------
'
Private Sub RefreshWebTestTreeDeferred()

    If TestRunner.State = etrsRunning Then Exit Sub
    If Not WebRunnerReady() Then Exit Sub

    DoEvents
    If TestRunner.State = etrsRunning Then Exit Sub

    modTestRunnerDiag.Diag "refresh.tree"
    If m_blnStandalone Then
        modTestState.LoadInto TestRunner
    End If
    TestRunner.ScanMergingPriorResults
    PublishTestTree TestRunner.GetTestTreeAsJson()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureRunnerHasTests
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Ensure the TestRunner singleton has discovered tests, loading durable
'           : state from disk when the in-memory dictionary is empty.
'---------------------------------------------------------------------------------------
'
Private Function EnsureRunnerHasTests() As Boolean

    If TestRunner Is Nothing Then Exit Function
    If TestRunner.Tests Is Nothing Then Exit Function
    If TestRunner.Tests.Count = 0 Then
        modTestState.LoadInto TestRunner
    End If
    If TestRunner.Tests Is Nothing Then Exit Function
    EnsureRunnerHasTests = (TestRunner.Tests.Count > 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : HasCompletedTests
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : True when the singleton has at least one non-pending test result.
'---------------------------------------------------------------------------------------
'
Private Function HasCompletedTests() As Boolean

    Dim varKey As Variant
    Dim dTest As Dictionary

    If Not EnsureRunnerHasTests() Then Exit Function

    For Each varKey In TestRunner.Tests.Keys
        Set dTest = TestRunner.Tests(CStr(varKey))
        If CLng(Nz(dTest("status"), etsPending)) <> etsPending Then
            HasCompletedTests = True
            Exit Function
        End If
    Next varKey

End Function


'---------------------------------------------------------------------------------------
' Procedure : StreamCompletedTestResults
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Replay completed test results to the web UI without republishing the tree.
'           : Pushed as a SINGLE onResultsBatch payload rather than one onTestComplete
'           : per test: the reopen/reload replay of a full run (~187 results) was ~187
'           : synchronous ExecuteJavascript round-trips, each forcing a full re-render,
'           : which spiked the WebView2 and stalled the next bridge poll (a 2s
'           : poll.timeout in the diagnostic trace). One call, one re-render.
'---------------------------------------------------------------------------------------
'
Private Sub StreamCompletedTestResults()

    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim colResults As Collection
    Dim dPayload As Dictionary

    If Not EnsureRunnerHasTests() Then Exit Sub
    If Not WebRunnerReady() Then Exit Sub

    Set colResults = New Collection
    For Each varKey In TestRunner.Tests.Keys
        Set dTest = TestRunner.Tests(CStr(varKey))
        If CLng(Nz(dTest("status"), etsPending)) <> etsPending Then
            colResults.Add AdaptTestResult(CStr(varKey), dTest)
        End If
    Next varKey

    If colResults.Count = 0 Then Exit Sub

    Set dPayload = New Dictionary
    Set dPayload("results") = colResults
    PushTestUI "onResultsBatch", ConvertToJson(dPayload)

End Sub


Private Sub ConnectBridgeInPage()
    Dim frm As Object
    On Error Resume Next
    If WebRunnerReady() Then
        Set frm = RunnerForm()
        If Not frm Is Nothing Then frm.ExecuteRunnerScript "window.VBA._connected = true;"
    End If
    Err.Clear
End Sub


Private Sub TryPublishPendingTree()

    If Not WebRunnerReady() Then Exit Sub
    If m_blnTreePublished Then Exit Sub
    If Len(m_strPendingTreeJson) = 0 Then Exit Sub

    PushTestUI "setContext", GetRunnerContextJson()
    PushTestUI "onReady", m_strPendingTreeJson
    m_blnTreePublished = True

End Sub


Private Function GetRunnerContextJson() As String

    Dim dCtx As Dictionary

    Set dCtx = New Dictionary
    dCtx.Add "projectName", RunnerProjectDisplayName()
    If Len(m_strPendingDefaultFilter) > 0 Then
        dCtx.Add "defaultFilter", m_strPendingDefaultFilter
    End If
    GetRunnerContextJson = ConvertToJson(dCtx)

End Function


Private Function RunnerProjectDisplayName() As String

    Dim strName As String

    strName = CurrentProject.Name
    If Len(strName) = 0 Then
        On Error Resume Next
        strName = FSO.GetFileName(CurrentProject.FullName)
        Err.Clear
    End If
    RunnerProjectDisplayName = strName

End Function


Private Function IsAllowedBridgeCallback(ByVal strFnName As String) As Boolean
    IsAllowedBridgeCallback = (InStr(1, "," & ALLOWED_CALLBACKS & ",", _
        "," & strFnName & ",", vbTextCompare) > 0)
End Function


Private Function WebStatusFromRunnerStatus(ByVal lngStatus As Long) As String
    Select Case lngStatus
        Case etsPassed:  WebStatusFromRunnerStatus = "pass"
        Case etsFailed:  WebStatusFromRunnerStatus = "fail"
        Case etsErrored: WebStatusFromRunnerStatus = "error"
        Case etsEmpty:   WebStatusFromRunnerStatus = "skip"
        Case Else:       WebStatusFromRunnerStatus = "pending"
    End Select
End Function


Private Function GetAccessFileBuild() As Long

    Dim strExe As String
    Dim strVersion As String
    Dim varParts As Variant

    On Error GoTo CleanUp

    strExe = SysCmd(acSysCmdAccessDir) & "MSACCESS.EXE"
    strVersion = FSO.GetFileVersion(strExe)
    If Len(strVersion) > 0 Then
        varParts = Split(strVersion, ".")
        If UBound(varParts) >= 2 Then
            GetAccessFileBuild = CLng(varParts(2))
        End If
    End If

CleanUp:
    ' Fail-open: return 0 on error; do not warn (WMI path raised a one-shot dialog).
    CatchAny eelNoError, vbNullString, vbNullString, False, True

End Function


Private Sub EndInteractiveBridgeRun()

    If Operation.Status = eosRunning Then
        TestRunner.SaveResults
        modTestState.PersistAfterRun
        TestRunner.InvokeGlobalTestTeardown
        Operation.Finish IIf(TestRunner.State = etrsCancelled, eorCanceled, eorSuccess)
    End If

    Log.SuppressDebugOutput = False
    Operation.InteractionMode = m_eimPriorMode

    ' Teardown (and ActiveVBProject switches during the run) can leave the VBE in
    ' the foreground when it was already open. Return focus to the runner form.
    RefocusWebRunner

End Sub
