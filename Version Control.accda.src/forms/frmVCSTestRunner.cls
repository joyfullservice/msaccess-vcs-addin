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
'           : via RetrieveJavascriptValue every POLL_INTERVAL_MS and dispatches them.
'           : OUTBOUND (VBA->JS) and replies use ExecuteJavascript.
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
Private Const POLL_INTERVAL_MS As Long = 500

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
    Me.TimerInterval = POLL_INTERVAL_MS
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
End Sub


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
    Me.TimerInterval = POLL_INTERVAL_MS
    Me.SetFocus

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DrainCancelOutbox
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Pull Cancel commands from window.__vbaOutbox without disturbing other
'           : queued commands. Callable from the test run loop (Form_Timer cannot
'           : re-enter while a run started from the timer is still on the stack).
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
    strResult = modTestRunnerUI.DispatchBridgeCall(strFnName, strRequestId, strPayloadJson)
    ResolveBridgeCall strRequestId, strResult
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
'           : A Cancel is dispatched even while a run is in progress (this timer tick can
'           : re-enter during the run's DoEvents) so Stop interrupts immediately; other
'           : commands only dispatch when idle. Request-id dedup prevents double-runs.
'---------------------------------------------------------------------------------------
'
Private Sub DrainOutbox()

    Const FunctionName As String = ModuleName & ".DrainOutbox"

    Dim strJson As String
    Dim colQueue As Collection
    Dim dItem As Dictionary
    Dim i As Long
    Dim strFn As String
    Dim strId As String
    Dim strPayload As String

    Dim curStart As Currency
    Dim blnRetrieveErr As Boolean

    If m_ctlWeb Is Nothing Then Exit Sub
    If Not modTestRunnerUI.WebRunnerActive() Then Exit Sub
    If Not modTestRunnerUI.WebRunnerDocumentReady() Then Exit Sub

    ' Time the drain's RetrieveJavascriptValue (the expensive part) even when it fails,
    ' so a timeout -- the slowest case -- is still recorded in the poll stats.
    curStart = Perf.MicroTimer
    On Error Resume Next
    strJson = modTestRunnerUI.RetrieveJsValue(m_ctlWeb, _
        "JSON.stringify((window.__vbaOutbox || []).splice(0, 20))")
    blnRetrieveErr = (Err.Number <> 0)
    Err.Clear
    On Error GoTo CleanUp

    MonitorPoll (Perf.MicroTimer - curStart) * 1000, blnRetrieveErr
    If blnRetrieveErr Then GoTo CleanUp
    If Len(strJson) = 0 Or strJson = "[]" Then GoTo CleanUp

    Set colQueue = ParseJson(strJson)
    For i = 1 To colQueue.Count
        Set dItem = colQueue(i)
        strFn = CStr(dItem("fn"))
        strId = CStr(dItem("id"))
        strPayload = "{}"
        If dItem.Exists("params") Then strPayload = ConvertToJson(dItem("params"))

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
    DrainDiag
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
' Purpose   : Track the per-drain RetrieveJavascriptValue cost. Emits an immediate
'           : "poll.slow" line for outliers and a rolling "poll.stats" summary every
'           : POLL_REPORT_EVERY drains, so the polling overhead is visible in the trace
'           : without a line per (500ms) tick. Aggregate totals also appear in the Perf
'           : report via the "RetrieveJavascriptValue" operation.
'---------------------------------------------------------------------------------------
'
Private Sub MonitorPoll(ByVal dblMs As Double, ByVal blnTimedOut As Boolean)

    If dblMs < 0 Then dblMs = 0

    m_lngPollCount = m_lngPollCount + 1
    m_dblPollTotalMs = m_dblPollTotalMs + dblMs
    If dblMs > m_dblPollMaxMs Then m_dblPollMaxMs = dblMs

    If blnTimedOut Then
        modTestRunnerDiag.Diag "poll.timeout", Format$(CLng(dblMs), "0") & "ms"
    ElseIf dblMs > POLL_SLOW_MS Then
        modTestRunnerDiag.Diag "poll.slow", Format$(CLng(dblMs), "0") & "ms"
    End If

    If m_lngPollCount >= POLL_REPORT_EVERY Then
        modTestRunnerDiag.Diag "poll.stats", "n=" & m_lngPollCount & _
            " avg=" & Format$(m_dblPollTotalMs / m_lngPollCount, "0.0") & "ms" & _
            " max=" & Format$(m_dblPollMaxMs, "0") & "ms"
        m_lngPollCount = 0
        m_dblPollTotalMs = 0
        m_dblPollMaxMs = 0
    End If

End Sub
