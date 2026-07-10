Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSTestRunner
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : EdgeBrowserControl host for the web test runner UI and JS<->VBA bridge.
'           : INBOUND (JS->VBA) is a lightweight poll: JS enqueues commands in
'           : window.__vbaOutbox (no navigation), and this form's timer drains the queue
'           : via RetrieveJavascriptValue on an adaptive interval and dispatches them.
'           : OUTBOUND (VBA->JS) and replies use ExecuteJavascript.
'           :
'           : Poll cadence: adaptive when idle (100ms floor, backs off with measured
'           : read cost up to 2000ms); fixed 1000ms while etrsRunning (mostly moot --
'           : the timer is suspended for the duration of a run dispatched from DrainOutbox
'           : because Access will not re-enter Form_Timer while its handler is on the
'           : stack). Mid-run Cancel is handled by modTestRunnerUI.PollBridgeCancel from
'           : the clsTestRunner loop (throttled to 1s), not by this timer.
'           :
'           : Why polling, not events: exhaustive testing showed the only inbound signal
'           : that reaches this control's BeforeNavigate is a MAIN-FRAME navigation
'           : (iframe sub-frame navigations and the "vba://" scheme never fire it), and a
'           : scripted main-frame navigation is not cleanly cancelable -- it reloads the
'           : whole page. Polling is the only inbound path that never reloads the page.
'           : See DECISIONS.md (2026-07-08).
'           :
'           : DocumentComplete is still used (readiness + cold-start re-navigate). The
'           : Edge event property tokens are "OnBeforeNavigation"/"OnDocumentComplete"
'           : (NOT "OnBeforeNavigate" - rejected on import); wire via the Property Sheet.
'           : The BeforeNavigate handler's Cancel arg is As Integer (the real signature).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ModuleName As String = "Form_frmVCSTestRunner"
Private Const RUN_POLL_MS As Long = 1000          ' fixed interval while a run is active
Private Const IDLE_MIN_MS As Long = 100           ' fastest idle poll (button responsiveness)
Private Const IDLE_MAX_MS As Long = 2000          ' slowest idle poll (heavy interaction)
Private Const IDLE_INTERVAL_FACTOR As Long = 10   ' interval = readEwma x factor (clamped)
Private Const EWMA_ALPHA As Double = 0.3          ' smoothing for measured read cost

' Legacy inbound bridge marker (https://msaccess/__vba__/<fn>/<id>). The navigation-
' based inbound path was replaced by outbox polling (see header); BeforeNavigate now
' only logs page navigations. The marker is retained so any stray such navigation is
' still recognized rather than treated as a page load.
Private Const BRIDGE_MARKER As String = "/__vba__/"

Private m_ctlWeb As Object
Private m_dHandledRequests As Object
Private m_blnDispatchBusy As Boolean
Private m_blnAllowClose As Boolean

' Poll performance monitoring (the RetrieveJavascriptValue drain cost). Aggregated over
' a window and reported to the diagnostic trace so we can watch the polling overhead.
Private Const POLL_SLOW_MS As Double = 100          ' log any single drain slower than this
Private Const POLL_REPORT_EVERY As Long = 20        ' emit a poll.stats summary every N drains
Private m_lngPollCount As Long
Private m_dblPollTotalMs As Double
Private m_dblPollMaxMs As Double
Private m_lngDrainDepth As Long        ' >1 means a drain re-entered via DoEvents
Private m_lngPollSeq As Long           ' monotonically increasing tick counter
Private m_curLastDrainEnd As Currency  ' Perf.MicroTimer at end of previous drain
Private m_dblGapTotalMs As Double
Private m_dblGapMaxMs As Double
Private m_lngQueueMax As Long
Private m_lngReentrantCount As Long    ' drains in the current poll.stats window where depth>1
Private m_dblEwmaReadMs As Double    ' EWMA of RetrieveJavascriptValue drain cost (ms)
Private m_lngLastInterval As Long      ' last Me.TimerInterval set by DrainOutbox
Private m_blnEwmaInited As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Bind the Edge control (late-bound) and navigate to runner HTML.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    modTestRunnerDiag.Diag "form.load"
    EnsureInit
    NavigateRunnerHtml
    ' Start the inbound poll. It no-ops until the runner document is ready (guarded in
    ' DrainOutbox), so it is safe to run from load through the about:blank cold-start.
    Me.TimerInterval = IDLE_MIN_MS
    ' OpenForm clears the hourglass; keep it on until DocumentComplete / WaitForWebRunnerReady.
    DoCmd.Hourglass True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Timer
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Inbound bridge pump. Drains window.__vbaOutbox and dispatches commands.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Timer()
    EnsureInit
    DrainOutbox
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureInit
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Lazily bind the control and the handled-requests set. Called from every
'           : event entry point because Access does not guarantee that the instance
'           : whose control event fires is the same instance whose Form_Load ran; using
'           : Me here always binds to the firing instance's own control.
'---------------------------------------------------------------------------------------
'
Private Sub EnsureInit()
    If m_dHandledRequests Is Nothing Then Set m_dHandledRequests = CreateObject("Scripting.Dictionary")
    If m_ctlWeb Is Nothing Then Set m_ctlWeb = Me.webTestRunner
    If Not m_blnEwmaInited Then
        m_dblEwmaReadMs = EwmaSeedMs()
        m_blnEwmaInited = True
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EwmaSeedMs
' Purpose   : Initial EWMA read-cost estimate (yields IDLE_MIN_MS when clamped).
'---------------------------------------------------------------------------------------
'
Private Function EwmaSeedMs() As Double
    EwmaSeedMs = IDLE_MIN_MS / IDLE_INTERVAL_FACTOR
End Function


'---------------------------------------------------------------------------------------
' Procedure : ComputeNextPollInterval
' Purpose   : Next form-timer interval: fixed during a run, adaptive when idle.
'---------------------------------------------------------------------------------------
'
Private Function ComputeNextPollInterval() As Long

    Dim dblInterval As Double

    If TestRunner.State = etrsRunning Then
        ComputeNextPollInterval = RUN_POLL_MS
    Else
        dblInterval = m_dblEwmaReadMs * IDLE_INTERVAL_FACTOR
        If dblInterval < IDLE_MIN_MS Then dblInterval = IDLE_MIN_MS
        If dblInterval > IDLE_MAX_MS Then dblInterval = IDLE_MAX_MS
        ComputeNextPollInterval = CLng(dblInterval)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : When the user closes the form (X button or Escape), hide it instead of
'           : unloading to keep the WebView2 control warm. Programmatic callers set
'           : m_blnAllowClose via AllowClose before issuing DoCmd.Close, which lets the
'           : unload proceed so Access shutdown is never blocked.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)

    ' Allow close when: (a) programmatic callers set AllowClose, or (b) the form is
    ' already hidden (Access shutdown closing a hidden form — no user to hide for).
    If Not m_blnAllowClose And Me.Visible Then
        Cancel = True
        HideRunner
        Exit Sub
    End If

    If TestRunner.State = etrsRunning Then TestRunner.Cancel

    modTestRunnerDiag.Diag "form.unload"
    Me.TimerInterval = 0

    ' Best effort: release the loaded document before the control is destroyed.
    On Error Resume Next
    If Not m_ctlWeb Is Nothing Then m_ctlWeb.Navigate "about:blank"
    Err.Clear
    On Error GoTo 0

    modTestRunnerUI.NotifyDocumentLost
    modTestRunnerUI.ResetWebRunnerHostState
    Set m_ctlWeb = Nothing
    Set m_dHandledRequests = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_KeyDown
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Hide the form when the user presses Escape (KeyPreview is on).
'---------------------------------------------------------------------------------------
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        HideRunner
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AllowClose
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Set by programmatic callers before DoCmd.Close to bypass the hide-on-
'           : close behavior and allow a real unload (e.g., application shutdown).
'---------------------------------------------------------------------------------------
'
Public Property Let AllowClose(ByVal blnValue As Boolean)
    m_blnAllowClose = blnValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : HideRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Hide the form and disable the timer. Cancels any running test run.
'---------------------------------------------------------------------------------------
'
Public Sub HideRunner()

    If TestRunner.State = etrsRunning Then TestRunner.Cancel
    ' If a prior bridge run left Operation stuck (e.g. hide mid-run before Finish),
    ' clear it so the next Run Tests / Run click is not blocked.
    modTestRunnerUI.ClearOrphanedTestOperation
    modTestRunnerDiag.Diag "form.hide"
    Me.TimerInterval = 0
    Me.Visible = False

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShowRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Make the form visible again and re-enable the poll timer.
'---------------------------------------------------------------------------------------
'
Public Sub ShowRunner()

    modTestRunnerDiag.Diag "form.show"
    Me.Visible = True
    Me.TimerInterval = IDLE_MIN_MS
    Me.SetFocus

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DrainCancelOutbox
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Pull Cancel commands from window.__vbaOutbox without disturbing other
'           : queued commands. Called from modTestRunnerUI.PollBridgeCancel during a
'           : test run (the form timer is suspended while a run started from DrainOutbox
'           : is on the stack, so this is the sole mid-run cancel path).
'---------------------------------------------------------------------------------------
'
Public Sub DrainCancelOutbox()

    Const FunctionName As String = ModuleName & ".DrainCancelOutbox"

    Dim strJson As String
    Dim colQueue As Collection
    Dim dItem As Dictionary
    Dim i As Long
    Dim strId As String

    If m_ctlWeb Is Nothing Then Exit Sub
    If Not modTestRunnerUI.WebRunnerActive() Then Exit Sub
    If Not modTestRunnerUI.WebRunnerDocumentReady() Then Exit Sub

    On Error Resume Next
    strJson = modTestRunnerUI.RetrieveJsValue(m_ctlWeb, _
        "(function(){var o=window.__vbaOutbox||[],c=[],r=[];" & _
        "for(var i=0;i<o.length;i++){if(o[i].fn==='Cancel')c.push(o[i]);else r.push(o[i]);}" & _
        "window.__vbaOutbox=r;return JSON.stringify(c);})()")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo CleanUp

    If Len(strJson) = 0 Or strJson = "[]" Then GoTo CleanUp

    Set colQueue = ParseJson(strJson)
    For i = 1 To colQueue.Count
        Set dItem = colQueue(i)
        strId = CStr(dItem("id"))
        DispatchRequest "Cancel", strId, "{}"
    Next i

CleanUp:
    CatchAny eelWarning, vbNullString, FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReloadRunnerHtml
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Re-navigate the control (singleton reuse).
'---------------------------------------------------------------------------------------
'
Public Sub ReloadRunnerHtml()
    NavigateRunnerHtml
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExecuteRunnerScript
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Push JavaScript to the loaded page (VBA->JS). Works from the control's
'           : own events (confirmed: DocumentComplete pushes render).
'---------------------------------------------------------------------------------------
'
Public Sub ExecuteRunnerScript(ByVal strScript As String)

    If m_ctlWeb Is Nothing Then Exit Sub
    On Error Resume Next
    m_ctlWeb.ExecuteJavascript strScript
    If Err.Number <> 0 Then
        modTestRunnerDiag.Diag "push.exec.error", Err.Description
    End If
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RetrieveRunnerJsValue
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Read a JavaScript expression from the loaded page (used for warm-reuse
'           : health checks outside control events).
'---------------------------------------------------------------------------------------
'
Public Function RetrieveRunnerJsValue(ByVal strExpression As String) As String

    If m_ctlWeb Is Nothing Then Exit Function
    RetrieveRunnerJsValue = modTestRunnerUI.RetrieveJsValue(m_ctlWeb, strExpression)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DrainDiag
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Drain JS-side breadcrumbs into the trace. Must be called OUTSIDE the
'           : control's own events (RetrieveJavascriptValue times out inside them).
'---------------------------------------------------------------------------------------
'
Public Sub DrainDiag()
    modTestRunnerDiag.DiagDrainJs m_ctlWeb
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NavigateRunnerHtml
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Navigate to local runner HTML via the https://msaccess/ prefix.
'---------------------------------------------------------------------------------------
'
Private Sub NavigateRunnerHtml()

    Dim strUrl As String

    ' Self-heal: re-bind the control to this instance if it was never set.
    If m_ctlWeb Is Nothing Then Set m_ctlWeb = Me.webTestRunner
    If m_ctlWeb Is Nothing Then Exit Sub

    modTestRunnerUI.NotifyDocumentLost
    If Not m_dHandledRequests Is Nothing Then m_dHandledRequests.RemoveAll

    strUrl = modTestRunnerUI.ResolveRunnerNavigateUrl()
    If Len(strUrl) = 0 Then
        modTestRunnerDiag.Diag "navigate.abort", "empty URL (HTML not resolved)"
        DoCmd.Hourglass False
        Exit Sub
    End If

    modTestRunnerDiag.Diag "navigate.call"
    m_ctlWeb.Navigate strUrl

End Sub


'---------------------------------------------------------------------------------------
' Procedure : webTestRunner_BeforeNavigate
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Logs every navigation (diagnostics). Inbound bridge calls now arrive via
'           : outbox polling (DrainOutbox), not navigation, so the BRIDGE_MARKER branch
'           : below is a defensive no-op in practice (JS no longer navigates to signal).
'---------------------------------------------------------------------------------------
'
Private Sub webTestRunner_BeforeNavigate(Cancel As Integer, URL As String)

    Dim strUrl As String
    Dim lngPos As Long
    Dim strPath As String
    Dim varParts As Variant
    Dim strFnName As String
    Dim strRequestId As String

    EnsureInit
    strUrl = CStr(URL)

    ' Log EVERY navigation so we can see exactly what reaches this event.
    modTestRunnerDiag.Diag "beforenavigate.raw", strUrl

    ' A bridge signal is a main-frame navigation to https://msaccess/__vba__/<fn>/<id>.
    lngPos = InStr(1, strUrl, BRIDGE_MARKER, vbTextCompare)
    If lngPos = 0 Then Exit Sub   ' normal navigation (page load) -- allow it

    Cancel = True

    ' Everything after the marker is "<fn>/<id>" (ignore any trailing slash/query).
    strPath = Mid$(strUrl, lngPos + Len(BRIDGE_MARKER))
    strPath = Split(strPath, "?")(0)
    strPath = Split(strPath, "#")(0)
    If Right$(strPath, 1) = "/" Then strPath = Left$(strPath, Len(strPath) - 1)

    varParts = Split(strPath, "/")
    strFnName = varParts(0)
    If UBound(varParts) >= 1 Then
        strRequestId = varParts(1)
    Else
        strRequestId = vbNullString
    End If

    modTestRunnerDiag.Diag "beforenavigate", strFnName & " id=" & strRequestId

    DispatchRequest strFnName, strRequestId, "{}"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : webTestRunner_DocumentComplete
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Page loaded; mark ready and connect the bridge (VBA->JS pushes work here).
'---------------------------------------------------------------------------------------
'
Private Sub webTestRunner_DocumentComplete(URL As Variant)

    Dim strUrl As String
    EnsureInit
    strUrl = CStr(URL)
    modTestRunnerDiag.Diag "documentcomplete.event", strUrl

    ' Only the runner page (https://msaccess/...) counts as "ready". On a COLD first
    ' open the WebView2 control isn't initialized when Form_Load calls Navigate, so
    ' that early Navigate is lost and the control lands on about:blank (its
    ' ControlSource). That about:blank DocumentComplete is the reliable signal that the
    ' control is now initialized -- so (re)navigate to the runner from here. Event-
    ' driven, no timer.
    If InStr(1, strUrl, "msaccess", vbTextCompare) > 0 _
        And InStr(1, strUrl, BRIDGE_MARKER, vbTextCompare) = 0 Then
        modTestRunnerUI.NotifyDocumentReady
    Else
        ' about:blank (cold-init), a bridge nav that slipped through un-cancelled, or
        ' any stray landing -> (re)load the runner page.
        modTestRunnerDiag.Diag "renavigate", strUrl
        NavigateRunnerHtml
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DispatchRequest
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Dispatch an allowlisted bridge call and reply. Dedupes by request id.
'---------------------------------------------------------------------------------------
'
Private Sub DispatchRequest(ByVal strFnName As String, ByVal strRequestId As String, _
    ByVal strPayloadJson As String)

    Dim strResult As String

    If Len(strRequestId) > 0 Then
        If m_dHandledRequests.Exists(strRequestId) Then
            modTestRunnerDiag.Diag "dispatch.dedup.skip", strFnName & " id=" & strRequestId
            Exit Sub
        End If
        m_dHandledRequests(strRequestId) = True
    End If

    modTestRunnerDiag.Diag "dispatch", strFnName & " id=" & strRequestId

    On Error GoTo DispatchErr
    If modTestRunnerUI.IsRunCommand(strFnName) Then
        ' Run commands block for the whole run, so resolve the JS promise with an
        ' acceptance ack BEFORE executing; completion streams via onRunComplete /
        ' onRunCancelled / onRunError. A raise from AcceptBridgeRun (validation
        ' failure) still rejects the promise below.
        strResult = modTestRunnerUI.AcceptBridgeRun(strFnName, strPayloadJson)
        ResolveBridgeCall strRequestId, strResult
        modTestRunnerUI.ExecutePendingBridgeRun
    Else
        strResult = modTestRunnerUI.DispatchBridgeCall(strFnName, strRequestId, strPayloadJson)
        ResolveBridgeCall strRequestId, strResult
    End If
    Exit Sub

DispatchErr:
    modTestRunnerDiag.Diag "dispatch.error", strFnName & " id=" & strRequestId & " | " & Err.Description
    RejectBridgeCall strRequestId, Err.Description
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DrainOutbox
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Pull queued commands from window.__vbaOutbox and dispatch them. Runs on the
'           : form timer (NOT a control event), so RetrieveJavascriptValue works here.
'           : While a run is executing on the timer stack, Access does not re-enter
'           : Form_Timer; mid-run Cancel is handled by PollBridgeCancel instead. Other
'           : commands only dispatch when idle. Request-id dedup prevents double-runs.
'---------------------------------------------------------------------------------------
'
Private Sub DrainOutbox()

    Const FunctionName As String = ModuleName & ".DrainOutbox"

    Dim strJson As String
    Dim strExpr As String
    Dim strOutboxSplice As String
    Dim colQueue As Collection
    Dim dRoot As Dictionary
    Dim dItem As Dictionary
    Dim i As Long
    Dim strFn As String
    Dim strId As String
    Dim strPayload As String
    Dim lngQueueCount As Long
    Dim lngAgeMs As Long

    Dim curStart As Currency
    Dim dblDrainMs As Double
    Dim dblGapMs As Double
    Dim blnRetrieveErr As Boolean
    Dim blnDiagEnabled As Boolean
    Dim blnReentrant As Boolean
    Dim blnDidDrain As Boolean

    If m_ctlWeb Is Nothing Then Exit Sub
    If Not modTestRunnerUI.WebRunnerActive() Then Exit Sub
    If Not modTestRunnerUI.WebRunnerDocumentReady() Then Exit Sub

    m_lngPollSeq = m_lngPollSeq + 1
    m_lngDrainDepth = m_lngDrainDepth + 1
    blnReentrant = (m_lngDrainDepth > 1)
    If blnReentrant Then
        modTestRunnerDiag.Diag "poll.reentrant", "depth=" & m_lngDrainDepth & " seq=" & m_lngPollSeq
        GoTo CleanUp
    End If

    blnDiagEnabled = modTestRunnerDiag.DiagEnabled
    ' Stamp each drained item with age (ms in JS clock) so queue latency is visible in the trace.
    strOutboxSplice = "(function(){var a=window.__vbaOutbox||[],b=a.splice(0,20);" & _
        "return b.map(function(o){var c=Object.assign({},o);" & _
        "c.age=Date.now()-(o.ts||Date.now());return c;});})()"
    If blnDiagEnabled Then
        strExpr = "JSON.stringify({o:" & strOutboxSplice & _
            ",d:(window.__diag?window.__diag.splice(0):[])})"
    Else
        strExpr = "JSON.stringify(" & strOutboxSplice & ")"
    End If

    ' Time the drain's RetrieveJavascriptValue (the expensive part) even when it fails,
    ' so a timeout -- the slowest case -- is still recorded in the poll stats.
    curStart = Perf.MicroTimer
    If m_curLastDrainEnd <> 0 Then dblGapMs = (curStart - m_curLastDrainEnd) * 1000
    If dblGapMs < 0 Then dblGapMs = 0

    On Error Resume Next
    strJson = modTestRunnerUI.RetrieveJsValue(m_ctlWeb, strExpr)
    blnRetrieveErr = (Err.Number <> 0)
    Err.Clear
    On Error GoTo CleanUp

    dblDrainMs = (Perf.MicroTimer - curStart) * 1000
    If dblDrainMs < 0 Then dblDrainMs = 0
    blnDidDrain = True

    If Not blnRetrieveErr And Len(strJson) > 0 Then
        If blnDiagEnabled Then
            Set dRoot = ParseJson(strJson)
            If TypeName(dRoot) = "Dictionary" Then
                If dRoot.Exists("d") Then
                    If TypeName(dRoot("d")) = "Collection" Then
                        modTestRunnerDiag.DiagAppendItems dRoot("d")
                    End If
                End If
                If dRoot.Exists("o") Then
                    If TypeName(dRoot("o")) = "Collection" Then Set colQueue = dRoot("o")
                End If
            End If
        ElseIf strJson <> "[]" Then
            Set colQueue = ParseJson(strJson)
        End If
    End If

    If colQueue Is Nothing Then
        lngQueueCount = 0
    Else
        lngQueueCount = colQueue.Count
    End If

    MonitorPoll dblDrainMs, dblGapMs, lngQueueCount, blnRetrieveErr, blnReentrant

    If blnRetrieveErr Then GoTo CleanUp
    If lngQueueCount = 0 Then GoTo CleanUp

    For i = 1 To colQueue.Count
        Set dItem = colQueue(i)
        strFn = CStr(dItem("fn"))
        strId = CStr(dItem("id"))
        strPayload = "{}"
        If dItem.Exists("params") Then strPayload = ConvertToJson(dItem("params"))

        If dItem.Exists("age") Then
            lngAgeMs = CLng(dItem("age"))
            If lngAgeMs > POLL_SLOW_MS Then
                modTestRunnerDiag.Diag "outbox.age", strFn & " id=" & strId & " age=" & lngAgeMs & "ms"
            End If
        End If

        If StrComp(strFn, "Cancel", vbTextCompare) = 0 _
            Or StrComp(strFn, "ReportJsError", vbTextCompare) = 0 Then
            DispatchRequest strFn, strId, strPayload
        ElseIf Not m_blnDispatchBusy Then
            m_blnDispatchBusy = True
            DispatchRequest strFn, strId, strPayload
            m_blnDispatchBusy = False
        Else
            modTestRunnerDiag.Diag "outbox.busy.skip", strFn & " id=" & strId
        End If
    Next i

CleanUp:
    If blnDidDrain Then
        If TestRunner.State <> etrsRunning And m_lngLastInterval = RUN_POLL_MS Then
            m_dblEwmaReadMs = EwmaSeedMs()
        Else
            m_dblEwmaReadMs = (EWMA_ALPHA * dblDrainMs) + ((1 - EWMA_ALPHA) * m_dblEwmaReadMs)
        End If
    End If
    m_lngLastInterval = ComputeNextPollInterval()
    Me.TimerInterval = m_lngLastInterval
    m_curLastDrainEnd = Perf.MicroTimer
    m_lngDrainDepth = m_lngDrainDepth - 1
    CatchAny eelWarning, vbNullString, FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveBridgeCall
'---------------------------------------------------------------------------------------
'
Private Sub ResolveBridgeCall(ByVal strRequestId As String, ByVal strResultJson As String)
    If Len(strRequestId) = 0 Then Exit Sub
    modTestRunnerDiag.Diag "resolve", "id=" & strRequestId
    ExecuteRunnerScript "window.__vbaResolve('" & JsEscape(strRequestId) & "'," & _
        SanitizeJsonLiteral(strResultJson) & ");"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RejectBridgeCall
'---------------------------------------------------------------------------------------
'
Private Sub RejectBridgeCall(ByVal strRequestId As String, ByVal strErrMsg As String)
    If Len(strRequestId) = 0 Then Exit Sub
    modTestRunnerDiag.Diag "reject", "id=" & strRequestId & " | " & strErrMsg
    ExecuteRunnerScript "window.__vbaReject('" & JsEscape(strRequestId) & "'," & _
        QuoteJsString(strErrMsg) & ");"
End Sub


Private Function JsEscape(ByVal strValue As String) As String
    JsEscape = Replace(strValue, "\", "\\")
    JsEscape = Replace(JsEscape, "'", "\'")
End Function


Private Function QuoteJsString(ByVal strValue As String) As String
    QuoteJsString = "'" & JsEscape(strValue) & "'"
End Function


Private Function SanitizeJsonLiteral(ByVal strJson As String) As String
    If Len(strJson) = 0 Then
        SanitizeJsonLiteral = "null"
    ElseIf Left$(Trim$(strJson), 1) = "{" Or Left$(Trim$(strJson), 1) = "[" Or _
        Left$(Trim$(strJson), 1) = """" Or IsNumeric(strJson) Then
        SanitizeJsonLiteral = strJson
    Else
        SanitizeJsonLiteral = QuoteJsString(strJson)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : MonitorPoll
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Track the per-drain RetrieveJavascriptValue cost, inter-drain idle gap,
'           : queue depth, and re-entrancy. Emits an immediate "poll.slow" line for
'           : outliers and a rolling "poll.stats" summary every POLL_REPORT_EVERY drains,
'           : without a line per tick. Aggregate totals also appear in the Perf
'           : report via the "RetrieveJavascriptValue" operation.
'---------------------------------------------------------------------------------------
'
Private Sub MonitorPoll(ByVal dblDrainMs As Double, ByVal dblGapMs As Double, _
    ByVal lngQueueCount As Long, ByVal blnTimedOut As Boolean, ByVal blnReentrant As Boolean)

    Dim strDetail As String

    If dblDrainMs < 0 Then dblDrainMs = 0
    If dblGapMs < 0 Then dblGapMs = 0

    m_lngPollCount = m_lngPollCount + 1
    m_dblPollTotalMs = m_dblPollTotalMs + dblDrainMs
    If dblDrainMs > m_dblPollMaxMs Then m_dblPollMaxMs = dblDrainMs
    m_dblGapTotalMs = m_dblGapTotalMs + dblGapMs
    If dblGapMs > m_dblGapMaxMs Then m_dblGapMaxMs = dblGapMs
    If lngQueueCount > m_lngQueueMax Then m_lngQueueMax = lngQueueCount
    If blnReentrant Then m_lngReentrantCount = m_lngReentrantCount + 1

    strDetail = Format$(CLng(dblDrainMs), "0") & "ms gap=" & Format$(CLng(dblGapMs), "0") & _
        "ms q=" & lngQueueCount

    If blnTimedOut Then
        modTestRunnerDiag.Diag "poll.timeout", strDetail
    ElseIf dblDrainMs > POLL_SLOW_MS Then
        modTestRunnerDiag.Diag "poll.slow", strDetail
    End If

    If m_lngPollCount >= POLL_REPORT_EVERY Then
        modTestRunnerDiag.Diag "poll.stats", "n=" & m_lngPollCount & _
            " avg=" & Format$(m_dblPollTotalMs / m_lngPollCount, "0.0") & "ms" & _
            " max=" & Format$(m_dblPollMaxMs, "0") & "ms" & _
            " gap.avg=" & Format$(m_dblGapTotalMs / m_lngPollCount, "0.0") & "ms" & _
            " gap.max=" & Format$(m_dblGapMaxMs, "0") & "ms" & _
            " qmax=" & m_lngQueueMax & " reentrant=" & m_lngReentrantCount & _
            " int=" & m_lngLastInterval & "ms"
        m_lngPollCount = 0
        m_dblPollTotalMs = 0
        m_dblPollMaxMs = 0
        m_dblGapTotalMs = 0
        m_dblGapMaxMs = 0
        m_lngQueueMax = 0
        m_lngReentrantCount = 0
    End If

End Sub
