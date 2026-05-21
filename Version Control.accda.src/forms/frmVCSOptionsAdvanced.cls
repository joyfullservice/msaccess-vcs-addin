Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsAdvanced
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Advanced options section: debugging, hashing, export tweaks, and logging.
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

    With Form_frmVCSOptions.OptionsInstance
        ' Debugging
        Me.chkShowDebug = .ShowDebug
        Me.chkBreakOnError = .BreakOnError
        Me.chkShowVCSLegacy = .ShowVCSLegacy

        ' Hashing
        Me.cboHashAlgorithm = .HashAlgorithm
        Me.chkUseShortHash = .UseShortHash

        ' Export Tweaks
        Me.chkStripPublishOption = .StripPublishOption
        Me.chkPreserveRubberDuckID = .PreserveRubberDuckID
        Me.chkSaveAllDocumentProperties = .SaveAllDocumentProperties
        Me.chkSaveLinkedTableDef = .SaveLinkedTableDef

        ' System
        Me.chkUseGitIntegration = .UseGitIntegration

        ' Logging
        Me.txtMaxLogFiles = .MaxLogFiles

        ' Automated Testing
        Me.txtDefaultTestFilter = .DefaultTestFilter
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save form control values back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()

    With Form_frmVCSOptions.OptionsInstance
        ' Debugging
        .ShowDebug = Nz(Me.chkShowDebug)
        .BreakOnError = Nz(Me.chkBreakOnError)
        .ShowVCSLegacy = Nz(Me.chkShowVCSLegacy)

        ' Hashing
        .HashAlgorithm = Nz(Me.cboHashAlgorithm)
        .UseShortHash = Nz(Me.chkUseShortHash)

        ' Export Tweaks
        .StripPublishOption = Nz(Me.chkStripPublishOption)
        .PreserveRubberDuckID = Nz(Me.chkPreserveRubberDuckID)
        .SaveAllDocumentProperties = Nz(Me.chkSaveAllDocumentProperties)
        .SaveLinkedTableDef = Nz(Me.chkSaveLinkedTableDef)

        ' System
        .UseGitIntegration = Nz(Me.chkUseGitIntegration, 0)

        ' Logging
        .MaxLogFiles = Nz(Me.txtMaxLogFiles, 0)

        ' Automated Testing
        .DefaultTestFilter = Nz(Me.txtDefaultTestFilter, vbNullString)
    End With

End Sub
