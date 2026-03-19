Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsDefaults
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Defaults section: manage project default settings and add-in installation.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load install settings into controls.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    LoadOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IOptionsSection_LoadOptions / IOptionsSection_SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : IOptionsSection interface implementation.
'---------------------------------------------------------------------------------------
'
Private Sub IOptionsSection_LoadOptions()
    LoadOptions
End Sub

Private Sub IOptionsSection_SaveOptions()
    SaveOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load installation settings from the registry.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()
    With GetInstallSettings(False)
        Me.txtInstallFolder = .strInstallFolder
        Me.chkTrustAddInFolder = .blnTrustAddInFolder
        Me.chkUseRibbonAddIn = .blnUseRibbonAddIn
        Me.chkOpenAfterInstall = .blnOpenAfterInstall
        Me.chkUseCompiledAddIn = .blnUseCompiledAddIn
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Install settings are saved immediately via button actions, not deferred.
'           : (Project defaults are managed via dedicated buttons, not the main Save.)
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()
    ' Install and default settings are action-based (buttons), not deferred to Save.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAsDefault_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Save current options as default for new projects.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAsDefault_Click()

    Dim strPath As String
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Validate export folder for defaults
    strPath = Nz(opt.ExportFolder)
    If strPath <> vbNullString Then
        If InStr(1, strPath, "%dbName%", vbTextCompare) < 1 Then
            MsgBox2 T("Invalid Export Path for Default"), _
                T("If you specify an absolute or relative Export Path as a default option," & vbCrLf & _
                "you must include the %dbName% placeholder to keep the paths unique."), _
                T("Please update the Export Path and try again."), vbExclamation
            Exit Sub
        End If
    End If

    ' Save current working options as defaults
    ' (Ensure current subform state is saved first)
    Me.Parent.Form.SaveCurrentSubform
    opt.SaveOptionsAsDefault
    MsgBox2 T("Options Saved"), T("The current options have been saved as default for new projects."), , vbInformation

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdRestoreDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Restore the default options to this project.
'---------------------------------------------------------------------------------------
'
Private Sub cmdRestoreDefaults_Click()

    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance
    opt.LoadDefaultOptions

    ' Reload the current subform to reflect restored defaults
    Dim sec As IOptionsSection
    Set sec = Me
    sec.LoadOptions

    MsgBox2 T("Default Options Restored"), _
        T("The system default options have been restored to this project."), _
        T("Be sure to click Save and Close when you are finished making changes."), _
        vbInformation, T("Version Control System")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdClearDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the default and project settings.
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearDefaults_Click()
    Dim cDefaults As clsOptions
    Set cDefaults = New clsOptions
    cDefaults.SaveOptionsAsDefault
    If MsgBox2(T("Apply to this Project?"), _
        T("The default options have been reset. Would you like these applied to this project as well?"), _
        T("Click 'Yes' to apply the default options, or 'No' to leave the current options."), _
        vbQuestion + vbYesNo, T("Version Control System")) = VbMsgBoxResult.vbYes Then
        cmdRestoreDefaults_Click
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdOpenInstallFolder_Click
' Author    : Adam Waller
' Date      : 7/6/2023
' Purpose   : Open the installation folder.
'---------------------------------------------------------------------------------------
'
Private Sub cmdOpenInstallFolder_Click()
    Application.FollowHyperlink modInstall.GetInstallSettings.strInstallFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdUninstall_Click
' Author    : Adam Kauffman
' Date      : 5/25/2020
' Purpose   : Remove the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub cmdUninstall_Click()
    UninstallVCSAddin
End Sub
