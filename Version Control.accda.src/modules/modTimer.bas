Option Explicit
Option Private Module

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long) As Long

Private m_lngBuildTimerID As Long


'---------------------------------------------------------------------------------------
' Procedure : RunBuildAfterClose
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Schedule a timer to fire 1 second after closing the current database.
'---------------------------------------------------------------------------------------
'
Public Sub RunBuildAfterClose(strSourceFolder As String)
    m_lngBuildTimerID = SetTimer(0, 0, 1000, AddressOf BuildTimerCallback)
    ' We will also lose the TimerID private variable value, so save it to registry as well.
    SaveSetting GetCodeVBProject.Name, "Build", "TimerID", m_lngBuildTimerID
    SaveSetting GetCodeVBProject.Name, "Build", "SourceFolder", strSourceFolder
    ' Now we should be ready to close the current database
    Application.CloseCurrentDatabase
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildTimerCallback
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : This is called by the API to resume our build process after closing the
'           : current database. (CloseCurrentDatabase ends all executing code.)
'---------------------------------------------------------------------------------------
'
Public Sub BuildTimerCallback()

    Dim strFolder As String
    
    ' Look up the existing timer to make sure we kill it properly.
    If m_lngBuildTimerID = 0 Then m_lngBuildTimerID = GetSetting(GetCodeVBProject.Name, "Build", "TimerID", 0)
    If m_lngBuildTimerID <> 0 Then
        KillTimer 0, m_lngBuildTimerID
        Debug.Print "Killed build timer " & m_lngBuildTimerID
        m_lngBuildTimerID = 0
        SaveSetting GetCodeVBProject.Name, "Build", "TimerID", 0
    End If
    
    ' Now, with the timer killed, we can clear the saved value and relaunch the build.
    strFolder = GetSetting(GetCodeVBProject.Name, "Build", "SourceFolder")
    SaveSetting GetCodeVBProject.Name, "Build", "SourceFolder", vbNullString
    If strFolder <> vbNullString Then
        Build strFolder
    End If
    
End Sub