Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Sub Form_Resize()
    ScaleColumns Me, , Array(Me.txtTableIcon.Name, Me.cboFormatType.Name)

    ' Size to fit; don't rely on Access' saved settings to get this right.
    Me.txtTableIcon.ColumnWidth = -2
End Sub
