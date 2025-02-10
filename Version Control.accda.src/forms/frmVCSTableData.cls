Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/22/2024
' Purpose   : Load any initial settings
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Translation.ApplyTo Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : Adam Waller
' Date      : 3/22/2024
' Purpose   : Scale column sizes
'---------------------------------------------------------------------------------------
'
Public Sub Form_Resize()
    ScaleColumns Me, , Array(Me.txtTableIcon.Name, Me.cboFormatType.Name)

    ' Size to fit; don't rely on Access' saved settings to get this right.
    Me.txtTableIcon.ColumnWidth = -2
End Sub
