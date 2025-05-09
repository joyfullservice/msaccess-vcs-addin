VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOperation
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : This class manages the overall state of user operations such as Export,
'           : Merge, Import, custom action, etc...
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' This constant determines whether the current operation is still valid or if it is
' considered to have timed out. (Or exited without cleaning up properly.)
Private Const OPERATION_TIMEOUT_SECONDS As Integer = 60 * 10


' Periodically update this to Now() to manage timeouts.
Public Heartbeat As Date


' Use a private type to manage local variables
Private Type udtThis
    StartTime As Date
    FinishTime As Date
    State As eOperationState
    Result As eOperationResult

    ' Since multiple Access instances may be running, we need to
    ' track which instance initiated the operation.
    ' (We don't want multiple operations happening on different
    ' instances at the same time, since we also interact with
    ' potentially shared registry values.)
    AppHwnd As Long

    ' Other variables we are caching
    ErrorTrapping As Integer
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Pulse
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Pulse the heartbeat
'---------------------------------------------------------------------------------------
'
Public Function Pulse()
    Heartbeat = Now
End Function


'---------------------------------------------------------------------------------------
' Procedure : Begin
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Begin an operation (I.e. Export, Merge, Import, custom action, etc...)
'           : Returns true if we are safe to begin a new operation. (No forms open,
'           : no operations currently in progress.)
'---------------------------------------------------------------------------------------
'
Public Function Begin() As Boolean
    this.State = eosRunning
    this.StartTime = Now
    Me.Pulse
End Function


Public Function Stage()
    this.State = eosStaged
    Me.Pulse
End Function


Public Function Restore()
    this.State = eosRunning
    Me.Pulse
End Function


'---------------------------------------------------------------------------------------
' Procedure : Finish
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Complete the current operation, release all objects.
'---------------------------------------------------------------------------------------
'
Public Function Finish(Optional intResult As eOperationResult)
    this.State = eosReady
    If intResult > eorUnknown Then this.Result = intResult
    this.FinishTime = Now
End Function


'---------------------------------------------------------------------------------------
' Procedure : Status
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Return the current status of the operation.
'---------------------------------------------------------------------------------------
'
Public Property Get Status() As eOperationState
    ' Check for timeout
    If this.State <> eosReady Then
        If DateAdd("s", OPERATION_TIMEOUT_SECONDS, Me.Heartbeat) < Now Then
            ' Timeout exceeded from last heartbeat
        End If
    End If
    Status = this.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : UpdateRegistry
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Updates values in registry to reflect current class values
'---------------------------------------------------------------------------------------
'
Private Sub UpdateRegistry()
    ' Make sure we still have a valid cache
    If this.AppHwnd <> 0 Then
        CheckSetting "StartTime", this.StartTime
        CheckSetting "FinishTime", this.FinishTime
        CheckSetting "State", this.State
        CheckSetting "Result", this.Result
        CheckSetting "AppHwnd", this.AppHwnd
        CheckSetting "ErrorTrapping", this.AppHwnd
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckSetting
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Update the setting if it has changed. (Faster to read than write)
'---------------------------------------------------------------------------------------
'
Private Sub CheckSetting(strKey As String, varValue)
    If GetSetting(PROJECT_NAME, "State", strKey) <> CStr(varValue) Then
        SaveSetting PROJECT_NAME, "State", strKey, CStr(varValue)
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearRegistryCache
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Clear the entire key for the state.
'---------------------------------------------------------------------------------------
'
Private Sub ClearRegistryCache()
    DeleteSetting PROJECT_NAME, "State"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreFromRegistry
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Restore the operation state from the registry cache. Returns true if
'           : successful.
'---------------------------------------------------------------------------------------
'
Private Function RestoreFromRegistry() As Boolean
    ' Only restore to the same Access instance
    If GetSetting(PROJECT_NAME, "State", "AppHwnd") = hWndAccessApp Then
        this.StartTime = ReadReg("StartTime")
        this.FinishTime = ReadReg("FinishTime")
        this.State = ReadReg("State")
        this.Result = ReadReg("Result")
        this.AppHwnd = ReadReg("AppHwnd")
        this.ErrorTrapping = ReadReg("ErrorTrapping")
        RestoreFromRegistry = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadReg
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Read a setting from the registry (wrapper for readability)
'---------------------------------------------------------------------------------------
'
Private Function ReadReg(strKey As String) As String
    ReadReg = GetSetting(PROJECT_NAME, "State", strKey)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasFormOpen
' Author    : Adam Waller
' Date      : 3/29/2024
' Purpose   : Make sure no other forms are currently loaded before initiating a new
'           : operation. (Force linear usage of one operation at a time.)
'           : Also warns user that they must close the form before they can proceed.
'
'---------------------------------------------------------------------------------------
'
Private Function HasFormOpen(Optional blnWarnUser As Boolean = True) As Boolean

    Dim objForm As AccessObject
    Dim frm As Access.Form
    Dim strCaption As String

    ' Cache collection of form names used by this add-in
    Static dFormNames As Dictionary

    ' Build cached list of form names
    If dFormNames Is Nothing Then
        Set dFormNames = New Dictionary
        For Each objForm In CodeProject.AllForms
            dFormNames.Add objForm.Name, vbNullString
        Next objForm
    End If

    ' Check for open forms (from current database, add-in, or library database)
    For Each frm In Application.Forms
        If dFormNames.Exists(frm.Name) Then
            ' Found a matching form name.
            If frm.CurrentView = acCurViewDesign Then
                ' Design view is okay, since the add-in won't have a form in design
                ' view when running as an add-in. (No conflict)
            Else
                ' A form with the same name. This could be a problem.
                If blnWarnUser Then
                    If frm.Caption = vbNullString Then
                        strCaption = frm.Name
                    Else
                        ' Users may recognize caption better than name
                        strCaption = frm.Caption & " (" & frm.Name & ")"
                    End If
                    MsgBox2 "Add-in Form Already Open", _
                        "Please close '" & strCaption & "' before running this action.", , vbInformation
                    HasFormOpen = True
                    Exit Function
                End If
            End If
        End If
    Next

End Function




'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Release any outstanding objects and restore the session state
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    modObjects.ReleaseObjects
    RestoreState
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveState
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Save any user settings before running our code.
'---------------------------------------------------------------------------------------
'
Private Sub SaveState()

    Dim strValue As String

    ' Just in case the user terminated the code before the settings could be restored,
    ' check for any existing user settings that might need to be restored.
    If ReadSetting("Save Time") <> vbNullString Then RestoreState

    ' Save a value so we can delete this section later
    PreserveSetting "Save Time", Now()

    ' Error trapping setting. (We need this to "Break in Class Modules" for this add-in)
    strValue = Application.GetOption("Error Trapping")
    If strValue <> "1" Then
        PreserveSetting "Error Trapping", strValue
        Application.SetOption "Error Trapping", 1
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreState
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Restore user settings.
'---------------------------------------------------------------------------------------
'
Private Sub RestoreState()

    Dim strValue As String

    strValue = ReadSetting("Error Trapping")
    If strValue <> vbNullString Then
        If Application.GetOption("Error Trapping") <> strValue Then
            Application.SetOption "Error Trapping", CLng(strValue)
        End If
    End If

    ' Remove temporary registry section
    DeleteSetting PROJECT_NAME, "State"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreserveSetting
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Wrapper functions for saving and reading session state settings.
'---------------------------------------------------------------------------------------
'
Private Sub PreserveSetting(strName As String, strSetting As String)
    SaveSetting PROJECT_NAME, "State", strName, strSetting
End Sub
Private Function ReadSetting(strName As String, Optional strDefault As String) As String
    ReadSetting = GetSetting(PROJECT_NAME, "State", strName, strDefault)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Load initial settings for class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ' Attempt to restore from registry
    If Not RestoreFromRegistry Then
        ' Set initial values
        this.AppHwnd = Application.hWndAccessApp
    End If
End Sub
