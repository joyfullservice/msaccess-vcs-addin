Attribute VB_Name = "modTestRunnerDiag"
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")
'---------------------------------------------------------------------------------------
' Module    : modTestRunnerDiag
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Diagnostic trace log for the web test runner bridge and lifecycle. Writes
'           : a single, agent-readable text file capturing the real sequence of events
'           : (navigate, DocumentComplete, BeforeNavigate, RetrieveJavascriptValue
'           : timing/timeouts, deferred dispatch, resolve/reject, readiness wait) plus
'           : JS-side breadcrumbs drained from window.__diag. Also a phase-timing
'           : profiler: nested DiagBegin/DiagEnd spans, payload sizes, and a per-phase
'           : summary table. The point is to close the feedback loop: when the page
'           : fails to load, a call times out, or the runner feels slow, the log shows
'           : exactly where the flow diverged or the time went.
'           :
'           : Tracing is OFF by default (DiagEnabled = False). Set DiagEnabled = True in
'           : the Immediate Window (or VCS.TestRunnerDiag True) and reopen the runner
'           : to capture a session. All Diag / window.diag call sites remain in place
'           : for future debugging; they no-op when tracing is off.
'           :
'           : Location: <ExportFolder>\logs\TestRunnerDiag_<timestamp>.log (falls back
'           : to a temp folder when Options are not loaded). Each session gets its own
'           : timestamped file, matching the other logs' naming; the resolved path is
'           : written in the header.
'           :
'           : Columns: [+elapsed ms ?ms-since-previous] TAG | detail
'---------------------------------------------------------------------------------------

Private Const ModuleName As String = "modTestRunnerDiag"
Private Const DIAG_PREFIX As String = "TestRunnerDiag_"
Private Const ForAppending As Long = 8
Private Const FLUSH_EVERY As Long = 32
Private Const FLUSH_INTERVAL_MS As Long = 250
Private Const MAX_LONG As Double = 2147483647#

' Default False = tracing off. VBA initializes Boolean module variables to False.
Private m_blnEnabled As Boolean
Private m_curStart As Currency
Private m_strPath As String

' Buffered writer (open/append/close once per flush, not per line).
Private m_buf As clsConcat
Private m_lngBuffered As Long
Private m_curLastFlush As Currency

' Nested spans and per-phase totals.
Private m_colSpans As Collection
Private m_dPhases As Dictionary

' Delta column and idle-gap marker.
Private m_curLastLine As Currency
Private m_dblLastElapsed As Double
Private m_curLastIdle As Currency

' JS clock anchor: one Date.now() mapped onto the session MicroTimer.
Private m_blnJsClockSet As Boolean
Private m_dblJsNow As Double
Private m_curJsAnchor As Currency


'---------------------------------------------------------------------------------------
' Procedure : DiagEnabled
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Whether diagnostic tracing is active (off by default). Turning it off
'           : flushes the buffer and writes the phase-summary table.
'---------------------------------------------------------------------------------------
'
Public Property Get DiagEnabled() As Boolean
    DiagEnabled = m_blnEnabled
End Property
Public Property Let DiagEnabled(ByVal blnValue As Boolean)
    If m_blnEnabled And Not blnValue Then
        DiagWriteSummary
        DiagFlush
    End If
    m_blnEnabled = blnValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : DiagLogPath
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Full path to the current diagnostic log file (creating a session if none).
'---------------------------------------------------------------------------------------
'
Public Function DiagLogPath() As String
    EnsureSession
    DiagLogPath = m_strPath
End Function


'---------------------------------------------------------------------------------------
' Procedure : CurrentLogPath
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Path of the active session file, or empty when none has been started.
'           : Does not create a session (unlike DiagLogPath).
'---------------------------------------------------------------------------------------
'
Public Function CurrentLogPath() As String
    CurrentLogPath = m_strPath
End Function


'---------------------------------------------------------------------------------------
' Procedure : DiagFolder
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Folder where session files are written (trailing PathSep).
'---------------------------------------------------------------------------------------
'
Public Function DiagFolder() As String
    DiagFolder = ResolveDiagFolder()
End Function


'---------------------------------------------------------------------------------------
' Procedure : DiagStart
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Begin a FRESH diagnostic session (new timestamped file + header). Called
'           : when the runner form opens; no-ops when tracing is disabled.
'---------------------------------------------------------------------------------------
'
Public Sub DiagStart(ByVal strContext As String)

    If Not m_blnEnabled Then Exit Sub

    If Len(m_strPath) > 0 Then
        DiagWriteSummary
        DiagFlush
    End If

    ResetSessionState
    m_curStart = Perf.MicroTimer
    m_curLastLine = m_curStart
    m_strPath = ResolveDiagFolder() & DIAG_PREFIX & Format$(Now, "yyyymmdd\_hhnnss") & ".log"
    WriteHeader strContext
    DiagFlush

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureSession
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Lazily start a session if none is active (e.g. after a VBA state reset
'           : cleared the module variables). Only runs when tracing is enabled.
'---------------------------------------------------------------------------------------
'
Private Sub EnsureSession()

    If Not m_blnEnabled Then Exit Sub
    If Len(m_strPath) > 0 Then Exit Sub
    ResetSessionState
    m_curStart = Perf.MicroTimer
    m_curLastLine = m_curStart
    m_strPath = ResolveDiagFolder() & DIAG_PREFIX & Format$(Now, "yyyymmdd\_hhnnss") & ".log"
    WriteHeader "auto-started (no explicit DiagStart, or VBA state was reset)"
    DiagFlush

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResetSessionState
'---------------------------------------------------------------------------------------
'
Private Sub ResetSessionState()
    Set m_buf = New clsConcat
    m_lngBuffered = 0
    m_curLastFlush = 0
    Set m_colSpans = New Collection
    Set m_dPhases = New Dictionary
    m_dblLastElapsed = 0
    m_curLastIdle = 0
    m_blnJsClockSet = False
    m_dblJsNow = 0
    m_curJsAnchor = 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteHeader
'---------------------------------------------------------------------------------------
'
Private Sub WriteHeader(ByVal strContext As String)
    WriteRaw "======================================================================"
    WriteRaw "TestRunner diagnostic session  " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteRaw "Context      : " & strContext
    WriteRaw "VCS version  : " & SafeStr(GetVCSVersion())
    WriteRaw "Access ver   : " & SafeStr(Application.Version)
    WriteRaw "Log path     : " & m_strPath
    WriteRaw "Columns      : [+elapsed ms +delta] TAG | detail"
    WriteRaw "======================================================================"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Diag
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Append a single trace line (relative timestamp + tag + optional detail).
'           : Self-initializes a session if needed, is overflow-safe on the clock, and
'           : never raises (diagnostics must not perturb the flow they observe).
'---------------------------------------------------------------------------------------
'
Public Sub Diag(ByVal strTag As String, Optional ByVal strDetail As String = vbNullString)

    If Not m_blnEnabled Then Exit Sub

    On Error Resume Next
    EnsureSession
    WriteTimedLine ElapsedNowMs(), SpanIndent() & strTag, strDetail
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagBegin
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Open a nested timing span. Pair with DiagEnd using the same phase name.
'---------------------------------------------------------------------------------------
'
Public Sub DiagBegin(ByVal strPhase As String, Optional ByVal strDetail As String = vbNullString)

    Dim dSpan As Dictionary
    Dim lngDepth As Long

    If Not m_blnEnabled Then Exit Sub

    On Error Resume Next
    EnsureSession
    If m_colSpans Is Nothing Then Set m_colSpans = New Collection
    lngDepth = m_colSpans.Count
    Set dSpan = New Dictionary
    dSpan.Add "phase", strPhase
    dSpan.Add "start", Perf.MicroTimer
    dSpan.Add "depth", lngDepth
    dSpan.Add "minMs", 0#
    m_colSpans.Add dSpan
    WriteTimedLine ElapsedNowMs(), String$(lngDepth * 2, " ") & strPhase & ".begin", strDetail
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagBeginQuiet
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Open a timing span that always accrues to the phase summary but only prints
'           : when it runs longer than dblMinMs, as a single line. For phases called too
'           : often to print (console flushes, bridge value retrieves) -- printing every
'           : one buries the trace and the file I/O starts perturbing what it measures.
'---------------------------------------------------------------------------------------
'
Public Sub DiagBeginQuiet(ByVal strPhase As String, Optional ByVal dblMinMs As Double = 10)

    Dim dSpan As Dictionary

    If Not m_blnEnabled Then Exit Sub

    On Error Resume Next
    EnsureSession
    If m_colSpans Is Nothing Then Set m_colSpans = New Collection
    Set dSpan = New Dictionary
    dSpan.Add "phase", strPhase
    dSpan.Add "start", Perf.MicroTimer
    dSpan.Add "depth", m_colSpans.Count
    dSpan.Add "minMs", dblMinMs
    m_colSpans.Add dSpan
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagEnd
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Close the matching DiagBegin span and accrue its duration.
'---------------------------------------------------------------------------------------
'
Public Sub DiagEnd(ByVal strPhase As String, Optional ByVal strDetail As String = vbNullString)

    Dim dSpan As Dictionary
    Dim dblMs As Double
    Dim dblMinMs As Double
    Dim strExtra As String
    Dim lngDepth As Long
    Dim strClosed As String
    Dim blnMatched As Boolean

    If Not m_blnEnabled Then Exit Sub

    On Error Resume Next
    EnsureSession
    If m_colSpans Is Nothing Then Exit Sub
    If m_colSpans.Count = 0 Then
        WriteTimedLine ElapsedNowMs(), strPhase & ".end.unmatched", strDetail
        Err.Clear
        Exit Sub
    End If

    ' Pop from the top until the named phase is closed (or the stack is empty).
    Do While m_colSpans.Count > 0
        Set dSpan = m_colSpans(m_colSpans.Count)
        m_colSpans.Remove m_colSpans.Count
        strClosed = CStr(dSpan("phase"))
        dblMs = (Perf.MicroTimer - CCur(dSpan("start"))) * 1000
        If dblMs < 0 Then dblMs = 0
        AccruePhase strClosed, dblMs
        lngDepth = CLng(Nz(dSpan("depth"), 0))
        dblMinMs = CDbl(Nz(dSpan("minMs"), 0))
        blnMatched = (StrComp(strClosed, strPhase, vbTextCompare) = 0)
        strExtra = "ms=" & Format$(dblMs, "0.0")
        If blnMatched Then
            If Len(strDetail) > 0 Then strExtra = strExtra & " " & strDetail
        Else
            strExtra = strExtra & " implicit"
        End If
        If dblMinMs = 0 Then
            WriteTimedLine ElapsedNowMs(), String$(lngDepth * 2, " ") & strClosed & ".end", strExtra
        ElseIf dblMs >= dblMinMs Then
            ' Quiet span: no ".begin" was printed, so print the whole span on one line.
            WriteTimedLine ElapsedNowMs(), String$(lngDepth * 2, " ") & strClosed, strExtra
        End If
        If blnMatched Then Exit Do
    Loop
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagSize
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Record a payload size (bytes or character count) under a tag.
'---------------------------------------------------------------------------------------
'
Public Sub DiagSize(ByVal strTag As String, ByVal lngBytes As Long)
    If Not m_blnEnabled Then Exit Sub
    Diag strTag, "bytes=" & lngBytes
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagNoteVbaIdle
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Mark the moment VBA returned to the message loop. DiagSinceIdleMs at a
'           : later handler entry is the gap that separates "input was queued" from
'           : "handler was slow".
'---------------------------------------------------------------------------------------
'
Public Sub DiagNoteVbaIdle()
    If Not m_blnEnabled Then Exit Sub
    On Error Resume Next
    m_curLastIdle = Perf.MicroTimer
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagSinceIdleMs
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Milliseconds since DiagNoteVbaIdle (0 when never marked).
'---------------------------------------------------------------------------------------
'
Public Function DiagSinceIdleMs() As Double

    Dim dblMs As Double

    If Not m_blnEnabled Then Exit Function
    If m_curLastIdle = 0 Then Exit Function
    On Error Resume Next
    dblMs = (Perf.MicroTimer - m_curLastIdle) * 1000
    If dblMs < 0 Then dblMs = 0
    DiagSinceIdleMs = dblMs
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : DiagSetJsClock
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Anchor JS Date.now() onto the session MicroTimer so drained breadcrumbs
'           : can be placed at the time they happened rather than at drain time.
'---------------------------------------------------------------------------------------
'
Public Sub DiagSetJsClock(ByVal dblJsNowMs As Double)
    If Not m_blnEnabled Then Exit Sub
    On Error Resume Next
    EnsureSession
    m_dblJsNow = dblJsNowMs
    m_curJsAnchor = Perf.MicroTimer
    m_blnJsClockSet = True
    Diag "js.clock", "jsNow=" & Format$(dblJsNowMs, "0")
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagAppendItems
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Fold already-parsed JS-side breadcrumbs (from the combined outbox/diag
'           : poll) into the trace file. Avoids a second RetrieveJavascriptValue per
'           : timer tick. When a JS clock anchor is set and an item has ts, the line
'           : is stamped with the mapped session elapsed.
'---------------------------------------------------------------------------------------
'
Public Sub DiagAppendItems(ByVal colItems As Collection)

    Dim dItem As Dictionary
    Dim i As Long
    Dim dblElapsed As Double
    Dim strTag As String
    Dim strDetail As String

    If Not m_blnEnabled Then Exit Sub
    If colItems Is Nothing Then Exit Sub
    If colItems.Count = 0 Then Exit Sub

    On Error Resume Next
    EnsureSession
    For i = 1 To colItems.Count
        Set dItem = colItems(i)
        strTag = "js." & CStr(dItem("t"))
        strDetail = CStr(Nz(dItem("m"), vbNullString))
        If m_blnJsClockSet And dItem.Exists("ts") Then
            dblElapsed = MappedJsElapsed(CDbl(Nz(dItem("ts"), 0)))
        Else
            dblElapsed = ElapsedNowMs()
        End If
        WriteTimedLine dblElapsed, strTag, strDetail
    Next i
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagDrainJs
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Pull queued JS-side breadcrumbs from window.__diag and fold them into the
'           : same trace file. Prefer DiagAppendItems when the poll already retrieved
'           : __diag in the same RetrieveJavascriptValue as the outbox.
'---------------------------------------------------------------------------------------
'
Public Sub DiagDrainJs(ByVal ctl As Object)

    Dim strJson As String
    Dim colItems As Collection

    If Not m_blnEnabled Then Exit Sub
    If ctl Is Nothing Then Exit Sub

    On Error GoTo Bail
    strJson = modTestRunnerUI.RetrieveJsValue(ctl, _
        "JSON.stringify(window.__diag ? window.__diag.splice(0) : [])")

    If Len(strJson) = 0 Or strJson = "[]" Then Exit Sub

    Set colItems = ParseJson(strJson)
    DiagAppendItems colItems
    Exit Sub

Bail:
    Diag "js.drain.error", Err.Description
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagFlush
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Write any buffered lines to disk.
'---------------------------------------------------------------------------------------
'
Public Sub DiagFlush()

    If Not m_blnEnabled And m_lngBuffered = 0 Then Exit Sub
    FlushBuffer

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagWriteSummary
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Append a per-phase table (count / total / max / avg), sorted by total ms.
'---------------------------------------------------------------------------------------
'
Public Sub DiagWriteSummary()

    Dim varKey As Variant
    Dim dPhase As Dictionary
    Dim lngCount As Long
    Dim i As Long
    Dim varRecords() As Variant
    Dim varParts As Variant
    Dim strName As String
    Dim dblTotal As Double
    Dim dblMax As Double
    Dim dblAvg As Double
    Dim lngNameWidth As Long

    If m_dPhases Is Nothing Then Exit Sub
    If m_dPhases.Count = 0 Then Exit Sub

    On Error Resume Next
    lngCount = m_dPhases.Count
    ReDim varRecords(0 To lngCount - 1)
    lngNameWidth = 12
    i = 0
    For Each varKey In m_dPhases.Keys
        Set dPhase = m_dPhases(CStr(varKey))
        strName = CStr(varKey)
        If Len(strName) > lngNameWidth Then lngNameWidth = Len(strName)
        varRecords(i) = Format$(CLng(CDbl(dPhase("totalMs")) * 10), "000000000") & "|" & strName
        i = i + 1
    Next varKey
    QuickSort varRecords

    WriteRaw "----------------------------------------------------------------------"
    WriteRaw "Phase summary (sorted by total ms)"
    WriteRaw PadRight("phase", lngNameWidth) & "  count     total       max       avg"
    For i = lngCount - 1 To 0 Step -1
        varParts = Split(CStr(varRecords(i)), "|")
        strName = CStr(varParts(1))
        Set dPhase = m_dPhases(strName)
        dblTotal = CDbl(dPhase("totalMs"))
        dblMax = CDbl(dPhase("maxMs"))
        If CLng(dPhase("count")) > 0 Then
            dblAvg = dblTotal / CLng(dPhase("count"))
        Else
            dblAvg = 0
        End If
        WriteRaw PadRight(strName, lngNameWidth) & "  " & _
            PadLeft(CStr(CLng(dPhase("count"))), 5) & "  " & _
            PadLeft(Format$(dblTotal, "0.0"), 8) & "  " & _
            PadLeft(Format$(dblMax, "0.0"), 8) & "  " & _
            PadLeft(Format$(dblAvg, "0.0"), 8)
    Next i
    WriteRaw "----------------------------------------------------------------------"
    FlushBuffer
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveDiagFolder
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Choose the log folder: the standard export "logs" folder when Options are
'           : loaded, otherwise a temp folder (so form-lifecycle events logged before a
'           : run still land somewhere). Returns a path ending in PathSep.
'---------------------------------------------------------------------------------------
'
Private Function ResolveDiagFolder() As String

    Dim strFolder As String

    On Error Resume Next
    If OptionsLoaded Then strFolder = Options.GetExportFolder & "logs" & PathSep
    On Error GoTo 0

    If Len(strFolder) = 0 Then strFolder = GetTempFolder("MSAccessVCS_Diag") & PathSep
    VerifyPath strFolder & "placeholder"
    ResolveDiagFolder = strFolder

End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteTimedLine
'---------------------------------------------------------------------------------------
'
Private Sub WriteTimedLine(ByVal dblElapsed As Double, ByVal strTag As String, _
    ByVal strDetail As String)

    Dim dblDelta As Double
    Dim strLine As String

    If dblElapsed < 0 Then dblElapsed = 0
    If dblElapsed > MAX_LONG Then dblElapsed = 0
    dblDelta = dblElapsed - m_dblLastElapsed
    m_dblLastElapsed = dblElapsed
    m_curLastLine = Perf.MicroTimer

    strLine = "[+" & Format$(CLng(dblElapsed), "00000") & "ms " & _
        FormatSignedDelta(dblDelta) & "] " & strTag
    If Len(strDetail) > 0 Then strLine = strLine & " | " & strDetail
    WriteRaw strLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteRaw
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Buffer a raw line; flush every FLUSH_EVERY lines so tracing stays cheap.
'           : Also flush once FLUSH_INTERVAL_MS has passed. A VBA state reset discards
'           : the buffer along with the rest of the module state, and the tail is the
'           : part worth reading, so the time bound caps what a reset can take with it.
'---------------------------------------------------------------------------------------
'
Private Sub WriteRaw(ByVal strLine As String)

    On Error Resume Next
    If m_buf Is Nothing Then Set m_buf = New clsConcat
    m_buf.Add strLine, vbCrLf
    m_lngBuffered = m_lngBuffered + 1
    If m_lngBuffered >= FLUSH_EVERY Then
        FlushBuffer
    ElseIf MsSinceFlush() >= FLUSH_INTERVAL_MS Then
        FlushBuffer
    End If
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FlushBuffer
'---------------------------------------------------------------------------------------
'
Private Sub FlushBuffer()

    Dim ts As Object
    Dim strChunk As String

    On Error Resume Next
    If m_buf Is Nothing Then Exit Sub
    If m_lngBuffered = 0 Then Exit Sub
    If Len(m_strPath) = 0 Then m_strPath = DiagLogPath()
    strChunk = m_buf.GetStr
    Set ts = FSO.OpenTextFile(m_strPath, ForAppending, True)
    ts.Write strChunk
    ts.Close
    m_buf.Clear
    m_lngBuffered = 0
    m_curLastFlush = Perf.MicroTimer
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MsSinceFlush
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Milliseconds since the last flush, overflow-safe like ElapsedNowMs. Treats
'           : "never flushed" as due, so the first line lands on disk immediately.
'---------------------------------------------------------------------------------------
'
Private Function MsSinceFlush() As Double

    Dim dblMs As Double

    If m_curLastFlush = 0 Then
        MsSinceFlush = FLUSH_INTERVAL_MS
        Exit Function
    End If
    dblMs = (Perf.MicroTimer - m_curLastFlush) * 1000
    If dblMs < 0 Or dblMs > MAX_LONG Then dblMs = FLUSH_INTERVAL_MS
    MsSinceFlush = dblMs

End Function


'---------------------------------------------------------------------------------------
' Procedure : AccruePhase
'---------------------------------------------------------------------------------------
'
Private Sub AccruePhase(ByVal strPhase As String, ByVal dblMs As Double)

    Dim dPhase As Dictionary

    If m_dPhases Is Nothing Then Set m_dPhases = New Dictionary
    If m_dPhases.Exists(strPhase) Then
        Set dPhase = m_dPhases(strPhase)
    Else
        Set dPhase = New Dictionary
        dPhase.Add "count", 0&
        dPhase.Add "totalMs", 0#
        dPhase.Add "maxMs", 0#
        Set m_dPhases(strPhase) = dPhase
    End If
    dPhase("count") = CLng(dPhase("count")) + 1
    dPhase("totalMs") = CDbl(dPhase("totalMs")) + dblMs
    If dblMs > CDbl(dPhase("maxMs")) Then dPhase("maxMs") = dblMs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ElapsedNowMs
'---------------------------------------------------------------------------------------
'
Private Function ElapsedNowMs() As Double

    Dim dblMs As Double

    If m_curStart = 0 Then
        ElapsedNowMs = 0
        Exit Function
    End If
    dblMs = (Perf.MicroTimer - m_curStart) * 1000
    If dblMs < 0 Or dblMs > MAX_LONG Then dblMs = 0
    ElapsedNowMs = dblMs

End Function


'---------------------------------------------------------------------------------------
' Procedure : MappedJsElapsed
'---------------------------------------------------------------------------------------
'
Private Function MappedJsElapsed(ByVal dblJsTs As Double) As Double

    Dim dblMs As Double

    dblMs = (dblJsTs - m_dblJsNow) + ((m_curJsAnchor - m_curStart) * 1000)
    If dblMs < 0 Then dblMs = 0
    If dblMs > MAX_LONG Then dblMs = 0
    MappedJsElapsed = dblMs

End Function


'---------------------------------------------------------------------------------------
' Procedure : SpanIndent
'---------------------------------------------------------------------------------------
'
Private Function SpanIndent() As String
    If m_colSpans Is Nothing Then Exit Function
    If m_colSpans.Count = 0 Then Exit Function
    SpanIndent = String$(m_colSpans.Count * 2, " ")
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatSignedDelta
'---------------------------------------------------------------------------------------
'
Private Function FormatSignedDelta(ByVal dblDelta As Double) As String
    If dblDelta < 0 Then
        FormatSignedDelta = "-" & Format$(CLng(-dblDelta), "00000")
    Else
        FormatSignedDelta = "+" & Format$(CLng(dblDelta), "00000")
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : PadRight / PadLeft
'---------------------------------------------------------------------------------------
'
Private Function PadRight(ByVal strValue As String, ByVal lngWidth As Long) As String
    If Len(strValue) >= lngWidth Then
        PadRight = strValue
    Else
        PadRight = strValue & Space$(lngWidth - Len(strValue))
    End If
End Function

Private Function PadLeft(ByVal strValue As String, ByVal lngWidth As Long) As String
    If Len(strValue) >= lngWidth Then
        PadLeft = strValue
    Else
        PadLeft = Space$(lngWidth - Len(strValue)) & strValue
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SafeStr
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Null/empty-safe string coercion for header fields.
'---------------------------------------------------------------------------------------
'
Private Function SafeStr(ByVal varValue As Variant) As String
    On Error Resume Next
    SafeStr = CStr(Nz(varValue, vbNullString))
    Err.Clear
End Function
