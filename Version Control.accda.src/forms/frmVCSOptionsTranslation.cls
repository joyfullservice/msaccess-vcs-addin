Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsTranslation
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Translation options section: manage language translations for the add-in.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load translation settings into controls.
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
' Purpose   : Load translation settings from the Translation object.
'           : (Translation settings are managed separately from clsOptions.)
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()
    With Translation
        Me.chkContributeTranslations = .Contribute
        Me.txtTranslationsPath = .TranslationsPath
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Translation settings are saved immediately on change (not deferred),
'           : so this is intentionally empty.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()
    ' Translation settings are saved immediately via event handlers.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkContributeTranslations_Click
' Author    : Adam Waller
' Date      : 9/2/2024
' Purpose   : Set flag on whether to actively contribute to translation files.
'---------------------------------------------------------------------------------------
'
Private Sub chkContributeTranslations_Click()
    Translation.Contribute = Me.chkContributeTranslations
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdBrowseTranslationsPath_Click
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Browse for the translations folder.
'---------------------------------------------------------------------------------------
'
Private Sub cmdBrowseTranslationsPath_Click()

    Dim strFolder As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .ButtonName = T("Select Translations Folder")
        .InitialFileName = StripSlash(txtTranslationsPath) & PathSep
        .Title = T("Select Translations Files Folder")
        .Show
        If .SelectedItems.Count > 0 Then
            strFolder = .SelectedItems(1) & PathSep
            If FSO.FileExists(strFolder & PROJECT_NAME & ".pot") Then
                txtTranslationsPath = strFolder
                Translation.TranslationsPath = strFolder
            Else
                MsgBox2 T("Translation files not found"), _
                    T("Required translation files were not found in this folder."), _
                    T("You selected: ") & .SelectedItems(1), vbExclamation
                DoCmd.Hourglass False
                Exit Sub
            End If
        End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSyncTranslations_Click
' Author    : Adam Waller
' Date      : 8/23/2024
' Purpose   : Sync translation files with the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSyncTranslations_Click()

    With Translation
        DoCmd.Hourglass True

        If CodeProject.FullName = GetInstalledAddInFileName Then
            If .StringCount > 200 Then .ExportTemplate
        End If

        .LoadTranslations
        .ExportTranslations

        DoCmd.Hourglass False
    End With

    MsgBox2 T("Translations Updated"), T("Translation files have been syncronized with the add-in."), , vbInformation

End Sub
