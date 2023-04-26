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

    ' Clear values from registry (In case an operation sets another timer)
    SaveSetting PROJECT_NAME, "Timer", "Operation", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param1", vbNullString
    SaveSetting PROJECT_NAME, "Timer", "Param2", vbNullString

    ' Now, run the desired operation
    Select Case strCommand

        Case "HandleRibbonCommand"
            HandleRibbonCommand strParam1

        Case "Build"
            ' Build from source (full or merge build)
            Build strParam1, CBool(strParam2)

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
