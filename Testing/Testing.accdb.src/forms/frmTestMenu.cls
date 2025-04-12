Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmdShowMenu_Click()
    Menu.PopupBelow cmdShowMenu, "Demo Popup"
    'PrintPos
End Sub

Private Sub cmdSpecial_Click()
    Menu.PopupBelow cmdSpecial, "Special \\..// Popup"
End Sub
