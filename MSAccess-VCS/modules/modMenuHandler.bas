Option Compare Database
Option Explicit

Private m_Menu As New clsVbeMenu


'---------------------------------------------------------------------------------------
' Procedure : LoadVBEMenuForVCS
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Initialize the menu for VCS
'---------------------------------------------------------------------------------------
'
Public Sub LoadVBEMenuForVCS(Optional strType As String = "TortoiseSVN")
    
    Dim cModel As IVersionControl
    
    Select Case strType
        Case "TortoiseSVN"
            Set cModel = New clsModelSVN
        Case Else
            MsgBox strType & " Version Control System not currently supported" & vbCrLf & _
                "Please contact your administrator for assistance", vbExclamation
            Exit Sub
    End Select
    m_Menu.Construct cModel
    
End Sub
