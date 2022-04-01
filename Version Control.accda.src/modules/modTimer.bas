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

' Types of operations to resume
Public Enum eResumeOperation
    roUnspecified
    roExportCurrentDatabase
    roFullBuildFromSource
    roLocalizeLibRefs
End Enum


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

    Dim strFolder As String
    
    ' First, make sure we kill the timer!
    KillTimer
    
    ' Now, run the desired operation
    Select Case GetSetting(PROJECT_NAME, "Timer", "Operation", 0)
    
        Case roUnspecified
            ' Operation type not specified or not found.
        
        Case roExportCurrentDatabase
            ' Launch the code export process. (See modAddIn.RunExportForCurrentDB)
            ' Only supports exporting all database objects
            VCS.Export
                
        Case roFullBuildFromSource
            ' Build from source (Full build)
            strFolder = GetSetting(PROJECT_NAME, "Build", "SourceFolder")
            SaveSetting PROJECT_NAME, "Build", "SourceFolder", vbNullString
            If strFolder <> vbNullString Then Build strFolder, True
        
        Case roLocalizeLibRefs
        
    End Select
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetTimer
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Set the API timer to trigger the desired operation
'---------------------------------------------------------------------------------------
'
Public Sub SetTimer(intOperation As eResumeOperation, Optional strParam As String, Optional sngSeconds As Single = 0.5)

    Dim strPath As String
    
    ' Make sure we are not trying to stack timer operations
    If m_lngTimerID <> 0 Then
        MsgBox2 "Failed to Set Callback Timer", _
            "Multiple callback timers are not currently supported.", _
            "Please ensure that any previous timer was completed or killed first.", vbExclamation
        Exit Sub
    End If

    ' Set any additional parameters here
    Select Case intOperation

        Case roFullBuildFromSource
            ' Save build path
            SaveSetting PROJECT_NAME, "Build", "SourceFolder", strParam

    End Select

    ' Save ID to registry before setting the timer
    SaveSetting PROJECT_NAME, "Timer", "Operation", intOperation
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
        Debug.Print "Killed timer " & m_lngTimerID
        m_lngTimerID = 0
        SaveSetting PROJECT_NAME, "Timer", "TimerID", 0
    End If
End Sub
