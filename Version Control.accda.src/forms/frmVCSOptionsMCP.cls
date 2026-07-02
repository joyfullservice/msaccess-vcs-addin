Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsMCP
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : MCP options section: controls which capabilities are available to AI
'           : agents connecting via the MCP server.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Load option values into controls on initial display.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    LoadOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IOptionsSection_LoadOptions / IOptionsSection_SaveOptions
' Author    : Adam Waller
' Date      : 4/14/2026
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
' Date      : 4/14/2026
' Purpose   : Load option values from the working options instance into form controls.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()

    With Form_frmVCSOptions.OptionsInstance
        Me.chkMcpAllowImport = .McpAllowImport
        Me.chkMcpAllowCallVBA = .McpAllowCallVBA
        Me.chkMcpAllowExecuteSQL = .McpAllowExecuteSQL
        Me.chkMcpAllowRunVBA = .McpAllowRunVBA
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Save form control values back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()

    With Form_frmVCSOptions.OptionsInstance
        .McpAllowImport = Nz(Me.chkMcpAllowImport)
        .McpAllowCallVBA = Nz(Me.chkMcpAllowCallVBA)
        .McpAllowExecuteSQL = Nz(Me.chkMcpAllowExecuteSQL)
        .McpAllowRunVBA = Nz(Me.chkMcpAllowRunVBA)
    End With

End Sub
