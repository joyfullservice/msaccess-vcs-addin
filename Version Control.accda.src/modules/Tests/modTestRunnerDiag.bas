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
'           : JS-side breadcrumbs drained from window.__diag. The point is to close the
'           : feedback loop: when the page fails to load or a call times out, the log
'           : shows exactly where the flow diverged.
'           :
'           : Tracing is OFF by default (DiagEnabled = False). Set DiagEnabled = True in
'           : the Immediate Window and reopen the runner to capture a session. All Diag
'           : / window.diag call sites remain in place for future debugging.
'           :
'           : Location: <ExportFolder>\logs\TestRunnerDiag_<timestamp>.log (falls back
'           : to a temp folder when Options are not loaded). Each session gets its own
'           : timestamped file, matching the other logs' naming; the resolved path is
'           : written in the header.
'---------------------------------------------------------------------------------------

Private Const ModuleName As String = "modTestRunnerDiag"
Private Const DIAG_PREFIX As String = "TestRunnerDiag_"
Private Const ForAppending As Long = 8

Private Const MAX_LONG As Double = 2147483647#

' Default False = tracing off. VBA initializes Boolean module variables to False.
Private m_blnEnabled As Boolean
Private m_curStart As Currency
Private m_strPath As String


'---------------------------------------------------------------------------------------
' Procedure : DiagEnabled
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Whether diagnostic tracing is active (off by default).
'---------------------------------------------------------------------------------------
'
Public Property Get DiagEnabled() As Boolean
    DiagEnabled = m_blnEnabled
End Property
Public Property Let DiagEnabled(ByVal blnValue As Boolean)
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
' Procedure : DiagStart
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Begin a FRESH diagnostic session (new timestamped file + header). Called
'           : when the runner form opens; no-ops when tracing is disabled.
'---------------------------------------------------------------------------------------
'
Public Sub DiagStart(ByVal strContext As String)

    If Not m_blnEnabled Then Exit Sub

    m_curStart = Perf.MicroTimer
    m_strPath = ResolveDiagFolder() & DIAG_PREFIX & Format$(Now, "yyyymmdd\_hhnnss") & ".log"
    WriteHeader strContext

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
    m_curStart = Perf.MicroTimer
    m_strPath = ResolveDiagFolder() & DIAG_PREFIX & Format$(Now, "yyyymmdd\_hhnnss") & ".log"
    WriteHeader "auto-started (no explicit DiagStart, or VBA state was reset)"

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
    WriteRaw "Columns      : [+elapsed ms] TAG | detail"
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

    Dim dblMs As Double
    Dim strLine As String

    If Not m_blnEnabled Then Exit Sub

    On Error Resume Next
    EnsureSession
    ' Overflow-safe elapsed: m_curStart is 0 only if EnsureSession failed; guard anyway
    ' (MicroTimer is seconds-since-boot, so *1000 with a 0 start overflows Long).
    dblMs = (Perf.MicroTimer - m_curStart) * 1000
    If dblMs < 0 Or dblMs > MAX_LONG Then dblMs = 0
    strLine = "[+" & Format$(CLng(dblMs), "00000") & "ms] " & strTag
    If Len(strDetail) > 0 Then strLine = strLine & " | " & strDetail
    WriteRaw strLine
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiagAppendItems
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Fold already-parsed JS-side breadcrumbs (from the combined outbox/diag
'           : poll) into the trace file. Avoids a second RetrieveJavascriptValue per
'           : timer tick.
'---------------------------------------------------------------------------------------
'
Public Sub DiagAppendItems(ByVal colItems As Collection)

    Dim dItem As Dictionary
    Dim i As Long

    If Not m_blnEnabled Then Exit Sub
    If colItems Is Nothing Then Exit Sub
    If colItems.Count = 0 Then Exit Sub

    On Error Resume Next
    For i = 1 To colItems.Count
        Set dItem = colItems(i)
        Diag "js." & CStr(dItem("t")), CStr(Nz(dItem("m"), vbNullString))
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
' Procedure : WriteRaw
' Author    : Adam Waller
' Date      : 7/7/2026
' Purpose   : Append a raw line to the diagnostic file (open/append/close per line;
'           : diagnostics are not performance-critical and this is lock-friendly).
'---------------------------------------------------------------------------------------
'
Private Sub WriteRaw(ByVal strLine As String)

    Dim ts As Object

    On Error Resume Next
    If Len(m_strPath) = 0 Then m_strPath = DiagLogPath()
    Set ts = FSO.OpenTextFile(m_strPath, ForAppending, True)
    ts.WriteLine strLine
    ts.Close
    Err.Clear

End Sub


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
