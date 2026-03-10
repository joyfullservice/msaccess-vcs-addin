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
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : RegisterCallback
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Parse callback info JSON and store URL/operation_id for MCP callbacks.
'           : This method must be called before using PostCallback or CheckCancelled.
'---------------------------------------------------------------------------------------
'
Public Sub RegisterCallback(strCallbackInfo As String)

    On Error GoTo ErrHandler

    Dim dCallback As Dictionary
    Dim strJson As String

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
'---------------------------------------------------------------------------------------
'
Public Sub PostCallback(strType As String, _
    Optional lngProgress As Long = -1, _
    Optional lngTotal As Long = -1, _
    Optional strMessage As String = "", _
    Optional dExtraData As Dictionary = Nothing)

    ' Skip if not active or no callback URL
    MCPDebugLog "clsMCP.PostCallback: Type=" & strType & ", IsActive=" & this.IsActive & ", URL=" & this.CallbackUrl
    If Not this.IsActive Or Len(this.CallbackUrl) = 0 Then
        MCPDebugLog "clsMCP.PostCallback: Skipping (not active or no URL)"
        Exit Sub
    End If

    On Error Resume Next  ' Don't let callback failures break the operation

    Dim http As Object
    Dim dPayload As Dictionary
    Dim strJson As String
    Dim varKey As Variant

    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Build JSON payload
    Set dPayload = New Dictionary
    dPayload.Add "operation_id", this.OperationId
    dPayload.Add "type", strType

    ' Add optional fields
    If lngProgress >= 0 Then dPayload.Add "progress", lngProgress
    If lngTotal >= 0 Then dPayload.Add "total", lngTotal
    If Len(strMessage) > 0 Then dPayload.Add "message", strMessage

    ' Merge extra data if provided
    If Not dExtraData Is Nothing Then
        For Each varKey In dExtraData.Keys
            dPayload.Add varKey, dExtraData(varKey)
        Next varKey
    End If

    ' Convert to JSON
    strJson = modJsonConverter.ConvertToJson(dPayload)
    MCPDebugLog "clsMCP.PostCallback: Sending to " & this.CallbackUrl

    ' Send POST request
    http.Open "POST", this.CallbackUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send strJson
    MCPDebugLog "clsMCP.PostCallback: HTTP Status=" & http.Status

    ' Note: We don't check the response - callbacks are fire-and-forget
    ' Errors are silently ignored to prevent breaking the main operation

    On Error GoTo 0

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

    On Error Resume Next  ' Don't let cancellation check failures break the operation

    ' Build cancel-status URL from callback URL
    ' callback_url: http://127.0.0.1:54321/callback
    ' cancel_url:   http://127.0.0.1:54321/cancel-status/{operation_id}
    Dim strCancelUrl As String
    Dim http As Object
    Dim dResponse As Dictionary

    strCancelUrl = Replace(this.CallbackUrl, "/callback", "/cancel-status/" & this.OperationId)

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
End Sub
