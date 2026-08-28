VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMCP
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Manage HTTP callbacks to MCP server for async operations with progress reporting.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Integration")

' Throttle cancellation checks to avoid excessive HTTP requests
Private Const CANCEL_CHECK_INTERVAL_MS As Long = 500  ' Check every 500ms

' Use a private type to manage local variables
Private Type udtThis
    CallbackUrl As String
    OperationId As String
    IsActive As Boolean
    LastCancelCheck As Double
    LogPostAttempts As Long
    ProgressPostAttempts As Long
    TerminalPostAttempts As Long
    FailedPostAttempts As Long
    RequestsPruned As Long
    MaxInFlight As Long
End Type
Private this As udtThis

' Live XMLHTTP objects for async log/progress posts. Releasing the object
' before readyState=4 aborts the request.
Private m_colInFlight As Collection


'---------------------------------------------------------------------------------------
' Procedure : RegisterCallback
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Parse callback info JSON and store URL/operation_id for MCP callbacks.
'           : This method must be called before using PostCallback or CheckCancelled.
'---------------------------------------------------------------------------------------
'
Public Sub RegisterCallback(strCallbackInfo As String)

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Dim dCallback As Dictionary
    Dim strJson As String

    ResetDiagnostics

    ' Parse the callback info JSON
    strJson = Trim(strCallbackInfo)
    MCPDebugLog "clsMCP.RegisterCallback: Input length=" & Len(strJson)
    If Len(strJson) = 0 Then
        If this.IsActive Then MCPDebugLog "clsMCP.RegisterCallback: Empty input, setting IsActive=False"
        this.IsActive = False
        Exit Sub
    End If

    Set dCallback = modJsonConverter.ParseJson(strJson)
    MCPDebugLog "clsMCP.RegisterCallback: Parsed JSON successfully"

    ' Extract callback_url and operation_id
    If dCallback.Exists("callback_url") Then
        this.CallbackUrl = dCallback("callback_url")
        MCPDebugLog "clsMCP.RegisterCallback: callback_url=" & this.CallbackUrl
    Else
        this.CallbackUrl = vbNullString
        MCPDebugLog "clsMCP.RegisterCallback: No callback_url in JSON"
    End If

    If dCallback.Exists("operation_id") Then
        this.OperationId = dCallback("operation_id")
        MCPDebugLog "clsMCP.RegisterCallback: operation_id=" & this.OperationId
    Else
        this.OperationId = vbNullString
        MCPDebugLog "clsMCP.RegisterCallback: No operation_id in JSON"
    End If

    ' Set active flag if we have both required fields
    this.IsActive = (Len(this.CallbackUrl) > 0 And Len(this.OperationId) > 0)
    MCPDebugLog "clsMCP.RegisterCallback: IsActive=" & this.IsActive

    ' An MCP callback means this process is doing agent work. COM-launched
    ' Access has no windowing bias toward a P-core; skip user-owned sessions.
    If this.IsActive Then
        If OpenedByAutomation Then PreferFullPowerCurrentProcess
    End If

    ' Reset cancel check timer
    this.LastCancelCheck = 0

    Exit Sub

ErrHandler:
    ' If parsing fails, disable callbacks but don't break the operation
    MCPDebugLog "clsMCP.RegisterCallback: ERROR - " & Err.Description
    this.IsActive = False
    this.CallbackUrl = vbNullString
    this.OperationId = vbNullString

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PostCallback
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Send callback to MCP server via HTTP POST.
'           : Callback types: "progress", "log", "complete", "error", "cancelled"
'           : log/progress are async (fire-and-forget) so a large build is not
'           : blocked on localhost round-trips. complete/error/cancelled stay
'           : synchronous so the client sees a real finish before Access continues.
'---------------------------------------------------------------------------------------
'
Public Sub PostCallback(strType As String, _
    Optional lngProgress As Long = -1, _
    Optional lngTotal As Long = -1, _
    Optional strMessage As String = "", _
    Optional dExtraData As Dictionary = Nothing)

    Dim http As Object
    Dim dPayload As Dictionary
    Dim strJson As String
    Dim varKey As Variant
    Dim blnAsync As Boolean
    Dim strSubmitOperation As String
    Dim lngSendError As Long

    If Not this.IsActive Or Len(this.CallbackUrl) = 0 Then Exit Sub

    LogUnhandledErrors
    On Error Resume Next  ' Don't let callback failures break the operation

    Perf.OperationStart "MCP Callback"
    PruneCompletedRequests

    Set http = CreateObject("MSXML2.XMLHTTP")

    Set dPayload = New Dictionary
    dPayload.Add "operation_id", this.OperationId
    dPayload.Add "type", strType

    If lngProgress >= 0 Then dPayload.Add "progress", lngProgress
    If lngTotal >= 0 Then dPayload.Add "total", lngTotal
    If Len(strMessage) > 0 Then dPayload.Add "message", strMessage

    If Not dExtraData Is Nothing Then
        For Each varKey In dExtraData.Keys
            dPayload.Add varKey, dExtraData(varKey)
        Next varKey
    End If

    ' Left to report itself. Perf records operations exclusive of their parent, so the
    ' payload cost lands under "Convert to JSON" and this line covers transport only.
    ' Do not suppress the nested report to consolidate the two: that buys one readable
    ' number by corrupting the serializer's call count, which is the figure that shows
    ' how much JSON work the add-in does in total.
    strJson = modJsonConverter.ConvertToJson(dPayload)

    blnAsync = IsStreamingCallback(strType)
    strSubmitOperation = CallbackSubmitOperation(strType)
    CountPostAttempt strType

    Perf.OperationStart strSubmitOperation
    Err.Clear
    http.Open "POST", this.CallbackUrl, blnAsync
    http.setRequestHeader "Content-Type", "application/json"
    http.send strJson
    lngSendError = Err.Number
    If lngSendError <> 0 Then
        this.FailedPostAttempts = this.FailedPostAttempts + 1
        Err.Clear
    ElseIf blnAsync Then
        If Not http Is Nothing Then
            If m_colInFlight Is Nothing Then Set m_colInFlight = New Collection
            m_colInFlight.Add http
            If m_colInFlight.Count > this.MaxInFlight Then _
                this.MaxInFlight = m_colInFlight.Count
        End If
    End If
    Perf.OperationEnd

    Perf.OperationEnd
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsStreamingCallback
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : True for routine stream events that may be posted asynchronously.
'---------------------------------------------------------------------------------------
'
Private Function IsStreamingCallback(strType As String) As Boolean
    IsStreamingCallback = (StrComp(strType, "log", vbTextCompare) = 0) Or _
        (StrComp(strType, "progress", vbTextCompare) = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : CallbackSubmitOperation
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Separate log, progress, and terminal submission costs in Perf reports.
'---------------------------------------------------------------------------------------
'
Private Function CallbackSubmitOperation(strType As String) As String
    Select Case LCase$(strType)
        Case "log":       CallbackSubmitOperation = "MCP Log Submit"
        Case "progress":  CallbackSubmitOperation = "MCP Progress Submit"
        Case Else:        CallbackSubmitOperation = "MCP Terminal Submit"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : CountPostAttempt
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Count stream volume independently of timer nesting and JSON conversion.
'---------------------------------------------------------------------------------------
'
Private Sub CountPostAttempt(strType As String)
    Select Case LCase$(strType)
        Case "log"
            this.LogPostAttempts = this.LogPostAttempts + 1
        Case "progress"
            this.ProgressPostAttempts = this.ProgressPostAttempts + 1
        Case Else
            this.TerminalPostAttempts = this.TerminalPostAttempts + 1
    End Select
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PruneCompletedRequests
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : Drop finished async XMLHTTP objects so they do not accumulate, without
'           : waiting on in-flight posts.
'---------------------------------------------------------------------------------------
'
Private Sub PruneCompletedRequests()

    Dim lngIdx As Long
    Dim http As Object

    If m_colInFlight Is Nothing Then Exit Sub
    If m_colInFlight.Count = 0 Then Exit Sub

    Perf.OperationStart "MCP Request Prune"
    For lngIdx = m_colInFlight.Count To 1 Step -1
        Set http = Nothing
        Set http = m_colInFlight(lngIdx)
        If http Is Nothing Then
            m_colInFlight.Remove lngIdx
            this.RequestsPruned = this.RequestsPruned + 1
        ElseIf http.readyState >= 4 Then
            m_colInFlight.Remove lngIdx
            this.RequestsPruned = this.RequestsPruned + 1
        End If
    Next lngIdx
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDiagnosticData
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Return stream volume and async-request lifecycle counters for perf JSON.
'---------------------------------------------------------------------------------------
'
Public Function GetDiagnosticData() As Dictionary

    Dim dData As New Dictionary
    Dim lngPending As Long

    If Not m_colInFlight Is Nothing Then lngPending = m_colInFlight.Count

    dData.Add "active", this.IsActive
    dData.Add "logPostAttempts", this.LogPostAttempts
    dData.Add "progressPostAttempts", this.ProgressPostAttempts
    dData.Add "terminalPostAttempts", this.TerminalPostAttempts
    dData.Add "failedPostAttempts", this.FailedPostAttempts
    dData.Add "requestsPruned", this.RequestsPruned
    dData.Add "pendingRequests", lngPending
    dData.Add "maxInFlight", this.MaxInFlight

    Set GetDiagnosticData = dData

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetDiagnostics
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Start each registered callback operation with independent counters.
'---------------------------------------------------------------------------------------
'
Private Sub ResetDiagnostics()
    this.LogPostAttempts = 0
    this.ProgressPostAttempts = 0
    this.TerminalPostAttempts = 0
    this.FailedPostAttempts = 0
    this.RequestsPruned = 0
    If m_colInFlight Is Nothing Then Set m_colInFlight = New Collection
    this.MaxInFlight = m_colInFlight.Count
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Post a simple log message to MCP. Convenience wrapper for PostCallback.
'           : Use this for status updates like "Beginning export of 45 queries..."
'---------------------------------------------------------------------------------------
'
Public Sub Log(strMessage As String)
    PostCallback "log", , , strMessage
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckCancelled
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Check if operation should be cancelled by polling MCP server.
'           : Uses throttling to avoid excessive HTTP requests.
'---------------------------------------------------------------------------------------
'
Public Function CheckCancelled() As Boolean

    ' Return False if not active
    If Not this.IsActive Or Len(this.CallbackUrl) = 0 Then
        CheckCancelled = False
        Exit Function
    End If

    ' Throttle checks to avoid excessive HTTP requests
    Dim dblNow As Double
    dblNow = Timer * 1000  ' Convert to milliseconds

    If dblNow - this.LastCancelCheck < CANCEL_CHECK_INTERVAL_MS Then
        CheckCancelled = False
        Exit Function
    End If
    this.LastCancelCheck = dblNow

    LogUnhandledErrors
    On Error Resume Next  ' Don't let cancellation check failures break the operation

    ' Build cancel-status URL from callback URL
    ' callback_url: http://127.0.0.1:54321/callback
    ' cancel_url:   http://127.0.0.1:54321/cancel-status/{operation_id}
    Dim strCancelUrl As String
    Dim http As Object
    Dim dResponse As Dictionary

    strCancelUrl = Replace(this.CallbackUrl, "/callback", "/cancel-status/" & this.OperationId)

    ' This is the one MCP round-trip that still blocks, so it needs its own line rather
    ' than disappearing into whichever operation happened to be running. Timed after the
    ' throttle above, so the calls that return early cost nothing and count nothing.
    Perf.OperationStart "MCP Cancel Check"

    ' Check cancel status
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", strCancelUrl, False
    http.send

    If http.Status = 200 Then
        Set dResponse = modJsonConverter.ParseJson(http.responseText)
        If dResponse.Exists("cancelled") Then
            CheckCancelled = CBool(dResponse("cancelled"))
        Else
            CheckCancelled = False
        End If
    Else
        ' If request fails, assume not cancelled
        CheckCancelled = False
    End If

    Perf.OperationEnd

    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : CallbackUrl
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Get the callback URL
'---------------------------------------------------------------------------------------
'
Public Property Get CallbackUrl() As String
    CallbackUrl = this.CallbackUrl
End Property


'---------------------------------------------------------------------------------------
' Procedure : OperationId
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Get the operation ID
'---------------------------------------------------------------------------------------
'
Public Property Get OperationId() As String
    OperationId = this.OperationId
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsActive
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Check if callbacks are configured and active
'---------------------------------------------------------------------------------------
'
Public Property Get IsActive() As Boolean
    IsActive = this.IsActive
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Initialize class variables
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    this.CallbackUrl = vbNullString
    this.OperationId = vbNullString
    this.IsActive = False
    this.LastCancelCheck = 0
    Set m_colInFlight = New Collection
End Sub
