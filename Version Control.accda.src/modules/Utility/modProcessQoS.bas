Attribute VB_Name = "modProcessQoS"
'---------------------------------------------------------------------------------------
' Module    : modProcessQoS
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Prefer full-power cores for MCP-launched Access. Access is
'           : single-threaded; EcoQoS can park that thread on an LP-E core.
'           : This turns execution-speed throttling off and raises the process
'           : to Above Normal. It does not set CPU affinity.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

Private Const ModuleName = "modProcessQoS"

' PROCESS_INFORMATION_CLASS.ProcessPowerThrottling
Private Const PROCESS_POWER_THROTTLING As Long = 4
Private Const PROCESS_POWER_THROTTLING_CURRENT_VERSION As Long = 1
Private Const PROCESS_POWER_THROTTLING_EXECUTION_SPEED As Long = 1
Private Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000&

Private Type PROCESS_POWER_THROTTLING_STATE
    Version As Long
    ControlMask As Long
    StateMask As Long
End Type

Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function SetProcessInformation Lib "kernel32" ( _
    ByVal hProcess As LongPtr, _
    ByVal ProcessInformationClass As Long, _
    ByRef ProcessInformation As PROCESS_POWER_THROTTLING_STATE, _
    ByVal ProcessInformationSize As Long) As Long
Private Declare PtrSafe Function SetPriorityClass Lib "kernel32" ( _
    ByVal hProcess As LongPtr, _
    ByVal dwPriorityClass As Long) As Long


'---------------------------------------------------------------------------------------
' Procedure : PreferFullPowerCurrentProcess
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Best-effort: disable EcoQoS on this process and raise it to
'           : Above Normal so the scheduler prefers a P-core. Safe to call
'           : more than once. Swallows errors on older Windows.
'---------------------------------------------------------------------------------------
'
Public Sub PreferFullPowerCurrentProcess()

    Dim typState As PROCESS_POWER_THROTTLING_STATE
    Dim hProcess As LongPtr

    LogUnhandledErrors
    On Error Resume Next

    hProcess = GetCurrentProcess

    typState.Version = PROCESS_POWER_THROTTLING_CURRENT_VERSION
    typState.ControlMask = PROCESS_POWER_THROTTLING_EXECUTION_SPEED
    typState.StateMask = 0

    SetProcessInformation hProcess, PROCESS_POWER_THROTTLING, typState, LenB(typState)
    SetPriorityClass hProcess, ABOVE_NORMAL_PRIORITY_CLASS

    Err.Clear

End Sub
