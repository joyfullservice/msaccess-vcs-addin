Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsBuild
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Build options section: build behavior, git integration, and build hooks.
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
        Me.chkExportAfterMerge = .ExportAfterMerge
        Me.chkForceImportOriginalQuerySQL = .ForceImportOriginalQuerySQL
        Me.txtRunBeforeBuild = .RunBeforeBuild
        Me.txtRunAfterBuild = .RunAfterBuild
        Me.txtRunBeforeMerge = .RunBeforeMerge
        Me.txtRunAfterMerge = .RunAfterMerge
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
        .ExportAfterMerge = Nz(Me.chkExportAfterMerge)
        .ForceImportOriginalQuerySQL = Nz(Me.chkForceImportOriginalQuerySQL)
        .RunBeforeBuild = Nz(Me.txtRunBeforeBuild)
        .RunAfterBuild = Nz(Me.txtRunAfterBuild)
        .RunBeforeMerge = Nz(Me.txtRunBeforeMerge)
        .RunAfterMerge = Nz(Me.txtRunAfterMerge)
    End With

End Sub
