Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : chkAdvancedOptions_Click
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Toggle to show advanced install options.
'---------------------------------------------------------------------------------------
'
Private Sub chkAdvancedOptions_Click()
    tabInstallType = Not (chkAdvancedOptions - 1)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 2/4/2021
' Purpose   : Close form if user presses the escape key. (Cancel property = True)
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdInstall_Click
' Author    : Adam Waller
' Date      : 2/4/2021
' Purpose   : Install add-in using selected options.
'---------------------------------------------------------------------------------------
'
Private Sub cmdInstall_Click()

    Dim strFolder As String
    Dim strFile As String
    Dim strMsg As String

    ' Activate the hourglass before loading installer VBA module for somoother experince.
    DoCmd.Hourglass True
    DoEvents

    ' Validate the folder path
    strFolder = StripSlash(Nz(txtInstallFolder))
    If StrComp(GetInstallSettings.strInstallFolder, strFolder, vbTextCompare) <> 0 Then
        ' Using a custom install folder
        If Not FSO.FolderExists(strFolder) Then
            strMsg = T("Folder does not exist: {0}", var0:=strFolder)
        Else
            ' Test writing a file to make sure we have write access to this folder.
            LogUnhandledErrors
            On Error Resume Next
            strFile = strFolder & PathSep & "WriteTest.txt"
            WriteFile "Test", strFile
            CatchAny eelNoError, vbNullString
            If ReadFile(strFile) <> "Test" & vbCrLf Then strMsg = T("Unable to write to folder: {0}", var0:=strFolder)
        End If
    End If

    ' Resume normal error handling
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Bail out if we have a problem with the install.
    If strMsg <> vbNullString Then
        DoCmd.Hourglass False
        MsgBox2 T("Unable to Install"), strMsg, , vbExclamation
        Exit Sub
    End If

    ' Run the installer
    modInstall.InstallVCSAddin chkAddTrustedLocation, chkUseRibbon, chkOpenAfterInstall, txtInstallFolder, chkCreateCompiledVersion
    DoCmd.Hourglass False

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdChangeInstallFolder_Click
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Allow the user to select another location for installation.
'---------------------------------------------------------------------------------------
'
Private Sub cmdChangeInstallFolder_Click()

    ' See if the add-in is already installed
    If FSO.FileExists(GetAddInFileName) Then
        MsgBox2 T("Please Uninstall First"), _
            T("Please uninstall (and delete) the add-in, then reinstall to the new location. " & _
            "If you have already uninstalled, delete the addin file. " & _
            "The install folder will open now to allow you to delete the file."), _
            T("You will have the option to keep your current settings during the uninstall process.") _
            , vbExclamation
        Application.FollowHyperlink GetInstallSettings.strInstallFolder
    Else
        ' Show a folder picker to select the desired location.
        ' (The path will be validated before installation, just in case it is changed direclty in the text box.)
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            .ButtonName = T("Select Folder")
            .InitialFileName = GetInstallSettings.strInstallFolder
            .Title = T("Select {0} Install Folder", var0:=PROJECT_NAME)
            .Show
            If .SelectedItems.Count > 0 Then
                ' Selected a folder
                txtInstallFolder = .SelectedItems(1)
            Else
                ' Canceled dialog
                Exit Sub
            End If
        End With
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Setting the control source causes delayed display. This way the display
'           : is instant when the form is opened.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Const STYLE_NO_TABS As Integer = 2

    ' Change install type tab control to no tabs and no border
    tabInstallType.Style = STYLE_NO_TABS
    tabInstallType.BorderStyle = 0

    ' Load translations
    Translation.ApplyTo Me

    ' Display version (better performance than bound control)
    lblVersion.Caption = T("Version {0}", var0:=GetVCSVersion())

    With GetInstallSettings
        chkAddTrustedLocation = .blnTrustAddInFolder
        chkOpenAfterInstall = .blnOpenAfterInstall
        chkUseRibbon = .blnUseRibbonAddIn
        With txtInstallFolder
            .Value = GetInstallSettings.strInstallFolder
            .Locked = True ' Only enable this text box if not installed.
            .BackColor = IIf(.Locked, 15921906, 16777215)
        End With
    End With

    ' Show installed version
    If InstalledVersion = vbNullString Then
        lblInstalled.Caption = T("(Add-in not currently installed)")
    Else
        lblInstalled.Caption = T("Version {0} currently installed.", var0:=InstalledVersion)
    End If

End Sub
