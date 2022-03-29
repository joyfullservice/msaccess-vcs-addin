VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVersionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsVersionControl
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This class is publicly exposed through the VCS object in modAPI and can
'           : be used to automate add-in functionality from other systems.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Show
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Show the Version Control System main form
'---------------------------------------------------------------------------------------
'
Public Sub Show()
    AddInMenuItemLaunch
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShowOptions
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Show the options form
'---------------------------------------------------------------------------------------
'
Public Sub ShowOptions()
    ' Force reload of options from current project before opening the form.
    Set modObjects.Options = Nothing
    Form_frmVCSOptions.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Export the source code for the current database
'---------------------------------------------------------------------------------------
'
Public Sub Export()
    AddInMenuItemExport
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetOptions
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Access the options
'---------------------------------------------------------------------------------------
'
Public Function Options() As clsOptions
    Set Options = modObjects.Options
End Function


'---------------------------------------------------------------------------------------
' Procedure : RepairColors
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapply the color properties in the current database
'---------------------------------------------------------------------------------------
'
Public Sub RepairColors()
    RepairColorDefinitionBlocks
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Version
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Return the version of this (CodeDB) instance of VCS.
'---------------------------------------------------------------------------------------
'
Public Property Get Version() As String
    Version = GetVCSVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : GoToLink...
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Wrapper functions to visit online links (called from Ribbon menu)
'---------------------------------------------------------------------------------------
'
Public Sub GoToLinkHome()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration"
End Sub
Public Sub GoToLinkDocumentation()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/wiki"
End Sub
Public Sub GoToLinkSupport()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/issues"
End Sub
Public Sub GoToLinkDownload()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-integration/releases"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Save the current state so we can restore it after the current operation.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    SaveState
End Sub


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
