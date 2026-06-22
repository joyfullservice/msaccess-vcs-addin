Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsExport
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Export options section: source file settings, sanitization, content
'           : options, printer settings detail, and export hooks.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values and populate combo box row sources.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim intSanitizeLevel As eSanitizeLevel
    Dim intEnvLevel As eUseEnvConnections
    Dim lngFormat As eExportFormatVersion
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Populate sanitize level combo
    With Me.cboSanitizeLevel
        .RowSource = vbNullString
        For intSanitizeLevel = eslMinimal To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & opt.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Populate sanitize colors combo
    With Me.cboSanitizeColors
        .RowSource = vbNullString
        For intSanitizeLevel = eslNone To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & opt.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Populate .env connection string combo
    With Me.cboUseEnvForConnections
        .RowSource = vbNullString
        For intEnvLevel = uecAuto To eUseEnvConnections.[_Last] - 1
            .AddItem intEnvLevel & ";" & opt.GetUseEnvConnectionsName(intEnvLevel)
        Next intEnvLevel
    End With

    ' Populate export format version combo
    With Me.cboExportFormatVersion
        .RowSource = vbNullString
        For lngFormat = EFV_4_1_2 To eExportFormatVersion.[_Last]
            ' Packed version integers are sparse; whitelist real enum members only.
            Select Case lngFormat
                Case EFV_4_1_2, EFV_5_0_0
                    If lngFormat = LATEST_EXPORT_FORMAT Then
                        .AddItem lngFormat & ";" & ExportFormatToVersion(lngFormat) & " (latest)"
                    Else
                        .AddItem lngFormat & ";" & ExportFormatToVersion(lngFormat)
                    End If
            End Select
        Next lngFormat
    End With

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

    Dim varKey As Variant

    With Form_frmVCSOptions.OptionsInstance
        ' Source Files
        Me.chkUseFastSave = .UseFastSave
        Me.chkSplitLayoutFromVBA = .SplitLayoutFromVBA
        Me.cboExportFormatVersion = .ExportFormatVersion

        ' Sanitization
        Me.cboSanitizeLevel = .SanitizeLevel
        Me.cboSanitizeColors = .SanitizeColors

        ' Content
        Me.chkSaveQuerySQL = .SaveQuerySQL
        Me.chkFormatSQL = .FormatSQL
        Me.chkUseDeterministicQueryExport = .UseDeterministicQueryExport
        Me.chkSaveTableSQL = .SaveTableSQL
        Me.chkSavePrintVars = .SavePrintVars
        Me.chkDecodeConditionalFormatting = .DecodeConditionalFormatting
        Me.chkExtractThemeFiles = .ExtractThemeFiles
        Me.cboUseEnvForConnections = .UseEnvForConnections

        ' Printer settings detail
        For Each varKey In .ExportPrintSettings.Keys
            Me.Controls("chk" & varKey) = .ExportPrintSettings(varKey)
        Next varKey

        ' Hooks
        Me.txtRunBeforeExport = .RunBeforeExport
        Me.txtRunAfterExport = .RunAfterExport
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

    Dim varKey As Variant
    Dim dSettings As Dictionary

    With Form_frmVCSOptions.OptionsInstance
        ' Source Files
        .UseFastSave = Nz(Me.chkUseFastSave)
        .ExportFormatVersion = Nz(Me.cboExportFormatVersion)

        ' Check for SplitLayoutFromVBA change before saving
        OnSplitLayoutChange Nz(Me.chkSplitLayoutFromVBA)
        .SplitLayoutFromVBA = Nz(Me.chkSplitLayoutFromVBA)

        ' Sanitization
        .SanitizeLevel = Nz(Me.cboSanitizeLevel)
        .SanitizeColors = Nz(Me.cboSanitizeColors)

        ' Content
        .SaveQuerySQL = Nz(Me.chkSaveQuerySQL)
        .FormatSQL = Nz(Me.chkFormatSQL)
        .UseDeterministicQueryExport = Nz(Me.chkUseDeterministicQueryExport)
        .SaveTableSQL = Nz(Me.chkSaveTableSQL)
        .SavePrintVars = Nz(Me.chkSavePrintVars)
        .DecodeConditionalFormatting = Nz(Me.chkDecodeConditionalFormatting)
        .ExtractThemeFiles = Nz(Me.chkExtractThemeFiles)
        .UseEnvForConnections = Nz(Me.cboUseEnvForConnections)

        ' Printer settings detail
        Set dSettings = .ExportPrintSettings
        For Each varKey In dSettings.Keys
            dSettings(varKey) = CBool(Me.Controls("chk" & varKey))
        Next varKey

        ' Hooks
        .RunBeforeExport = Nz(Me.txtRunBeforeExport)
        .RunAfterExport = Nz(Me.txtRunAfterExport)
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OnSplitLayoutChange
' Author    : Adam Waller
' Date      : 11/9/2023
' Purpose   : When the user enables Split Layout from VBA, suggest the wiki page
'           : about preserving Git history for split files.
'---------------------------------------------------------------------------------------
'
Private Sub OnSplitLayoutChange(varNewValue As Variant)

    With Form_frmVCSOptions.OptionsInstance
        If .SplitLayoutFromVBA = varNewValue Then Exit Sub
    End With

    If varNewValue = True Then
        If Git.Installed Then
            If Git.IsInsideRepository Then
                If MsgBox2(T("May I make a Suggestion?"), _
                    T("This project appears to be within a Git repository. This add-in includes a special utility " & _
                    "that can split the files (layout and VBA) while preserving this history of previous changes in BOTH files."), _
                    T("Would you like to see additional information on this from the wiki?"), vbQuestion + vbYesNo) = vbYes Then
                    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Split-Files"
                End If
            End If
        End If
    End If

End Sub
