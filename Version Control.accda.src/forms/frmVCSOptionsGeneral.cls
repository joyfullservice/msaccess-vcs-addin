Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsGeneral
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : General options section: export folder, tools, and language.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values into controls on initial display.
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
' Purpose   : Load option values from the working options instance into form controls.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()

    With Form_frmVCSOptions
        Me.txtExportFolder = .OptionsInstance.ExportFolder

        ' Machine-level settings (pending registry values on main form)
        Me.cboDiffTool = .DiffTool
        Me.cboOpenRepository = .OpenRepository
    End With

    ' Language
    Me.cboLanguage = Translation.Language

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save form control values back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()

    With Form_frmVCSOptions
        .OptionsInstance.ExportFolder = Nz(Me.txtExportFolder)

        ' Store pending registry values on main form (committed on Save & Close)
        .DiffTool = Nz(Me.cboDiffTool)
        .OpenRepository = CLng(Nz(Me.cboOpenRepository, 0))
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtExportFolder_BeforeUpdate
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Validate the export folder path entry.
'---------------------------------------------------------------------------------------
'
Private Sub txtExportFolder_BeforeUpdate(Cancel As Integer)

    Dim strPath As String

    strPath = Nz(txtExportFolder)
    If strPath <> vbNullString Then
        If (Left(strPath, 1) = PathSep) Or _
            (InStr(2, strPath, ":" & PathSep) > 0) Then
            ' Looks like a valid path
        Else
            MsgBox2 T("Invalid Export Folder"), _
                T("This does not appear to be a valid relative or absolute path."), _
                T("Please see the wiki documentation for more detail and examples."), vbExclamation
            Cancel = True
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cboLanguage_Click
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Set language when changed.
'---------------------------------------------------------------------------------------
'
Private Sub cboLanguage_Click()
    With Translation
        .Language = Nz2(cboLanguage, "en_US")
        .SetLanguage .Language
        MsgBox2 T("Language Set"), _
            T("Changes to language settings will take effect the next time you open this form."), _
            , vbInformation
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdTranslations_Click
' Author    : Adam Waller
' Date      : 8/7/2024
' Purpose   : Navigate to the Translation section on the main options form.
'---------------------------------------------------------------------------------------
'
Private Sub cmdTranslations_Click()
    Me.Parent.Form.NavigateTo "Translation"
End Sub
