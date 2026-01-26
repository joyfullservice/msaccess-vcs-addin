Attribute VB_Name = "modTimer"
'---------------------------------------------------------------------------------------
' Module    : modTimer
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : API timer functions for callbacks
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

' Windows API calls for Timer functionality
Private Declare PtrSafe Function ApiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function ApiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private m_lngTimerID As LongPtr


'---------------------------------------------------------------------------------------
' Procedure : WinAPITimerCallback
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Generic callback function to handle timer requests to resume operations.
'---------------------------------------------------------------------------------------
'
Public Sub WinAPITimerCallback()

    Dim strParam1 As String
    Dim strParam2 As String
    Dim strCommand As String

    ' First, make sure we kill the timer!
    KillTimer

    ' Read in parameter values
    strCommand = GetSetting(PROJECT_NAME, "Timer", "Operation")
    strParam1 = GetSetting(PROJECT_NAME, "Timer", "Param1")
    strParam2 = GetSetting(PROJECT_NAME, "Timer", "Param2")

    ' Read callback info before clearing (needed for APIAsyncOperation)
    Dim strCallbackInfo As String
    strCallbackInfo = GetSetting(PROJECT_NAME, "Timer", "CallbackInfo")
    MCPDebugLog "WinAPITimerCallback: Command=" & strCommand & ", CallbackInfo length=" & Len(strCallbackInfo)

    ' Clear values from registry (In case an operation sets another timer)
    SaveSetting PROJECT_NAME, "Timer", "Operation", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param1", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param2", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", vbNullString

    ' Unstage the current operation
    If Operation.Status = eosStaged Then Operation.Restore

    ' Now, run the desired operation
    Select Case strCommand

        Case "HandleRibbonCommand"
            HandleRibbonCommand strParam1

        Case "Build"
            ' Build from source (full or merge build)
            Build strParam1, CBool(strParam2)

        Case "APIAsyncOperation"
            ' Handle async operation with MCP callbacks
            HandleAPIAsyncOperation strParam1, strParam2, strCallbackInfo

        Case Else
            ' Use the Run command to execute the specified operation with supplied parameters
            If strParam2 <> vbNullString Then
                Application.Run strCommand, strParam1, strParam2
            ElseIf strParam1 <> vbNullString Then
                Application.Run strCommand, strParam1
            Else
                Application.Run strCommand
            End If

    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Set the API timer to trigger the desired operation
'---------------------------------------------------------------------------------------
'
Public Sub SetTimer(strOperation As String, _
    Optional strParam1 As String, Optional strParam2 As String, _
    Optional sngSeconds As Single = 0.5)

    ' Make sure we are not trying to stack timer operations
    If m_lngTimerID <> 0 Then
        MsgBox2 "Failed to Set Callback Timer", _
            "Multiple callback timers are not currently supported.", _
            "Please ensure that any previous timer was completed or killed first.", vbExclamation
        Exit Sub
    End If

    ' Save parameter values
    SaveSetting PROJECT_NAME, "Timer", "Param1", strParam1
    SaveSetting PROJECT_NAME, "Timer", "Param2", strParam2

    ' Save ID to registry before setting the timer
    SaveSetting PROJECT_NAME, "Timer", "Operation", strOperation
    SaveSetting PROJECT_NAME, "Timer", "TimerID", m_lngTimerID
    m_lngTimerID = ApiSetTimer(0, 0, 1000 * sngSeconds, AddressOf WinAPITimerCallback)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : KillTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Kill any existing timer
'---------------------------------------------------------------------------------------
'
Private Sub KillTimer()
    If m_lngTimerID = 0 Then m_lngTimerID = GetSetting(PROJECT_NAME, "Timer", "TimerID", 0)
    If m_lngTimerID <> 0 Then
        ApiKillTimer 0, m_lngTimerID
        Debug.Print "Killed API Timer " & m_lngTimerID
        m_lngTimerID = 0
        SaveSetting PROJECT_NAME, "Timer", "TimerID", 0
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HandleAPIAsyncOperation
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Handle async operation with MCP callbacks. Reads callback info from
'           : registry, registers with MCP, then starts the operation.
'---------------------------------------------------------------------------------------
'
Private Sub HandleAPIAsyncOperation(strMethod As String, strArgs As String, strCallbackInfo As String)

    On Error GoTo ErrHandler

    Dim strArg1 As String
    Dim strArg2 As String
    Dim lngPipePos As Long

    ' Register callback with MCP if provided
    MCPDebugLog "HandleAPIAsyncOperation: Method=" & strMethod & ", CallbackInfo length=" & Len(strCallbackInfo)
    If Len(strCallbackInfo) > 0 Then
        MCPDebugLog "HandleAPIAsyncOperation: Registering callback..."
        MCP.RegisterCallback strCallbackInfo
        MCPDebugLog "HandleAPIAsyncOperation: MCP.IsActive=" & MCP.IsActive
        Operation.Source = eosMCPTool
    Else
        MCPDebugLog "HandleAPIAsyncOperation: No callback info, using External API source"
        Operation.Source = eosExternalAPI
    End If

    ' Parse arguments (format: "arg1|arg2" or just "arg1")
    If Len(strArgs) > 0 Then
        lngPipePos = InStr(strArgs, "|")
        If lngPipePos > 0 Then
            strArg1 = Left(strArgs, lngPipePos - 1)
            strArg2 = Mid(strArgs, lngPipePos + 1)
        Else
            strArg1 = strArgs
        End If
    End If

    ' Start the operation via API
    ' Log.Add automatically routes to MCP when MCP.IsActive
    If Len(strArg2) > 0 Then
        API strMethod, strArg1, strArg2
    ElseIf Len(strArg1) > 0 Then
        API strMethod, strArg1
    Else
        API strMethod
    End If

    ' Completion callback is now sent from Operation.Finish() before ReleaseObjects
    MCPDebugLog "HandleAPIAsyncOperation: Operation complete, Result=" & Operation.Result

    Exit Sub

ErrHandler:
    ' Post error callback if MCP is active
    If MCP.IsActive Then
        MCP.PostCallback "error", -1, -1, strMethod & " failed: " & Err.Description
    End If

    ' Re-throw error
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub
