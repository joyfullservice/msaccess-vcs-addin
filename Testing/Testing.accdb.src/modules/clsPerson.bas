Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public FirstName As String
Public LastName As String


Public Sub Greet()
    MsgBox "Hello, " & Me.FirstName & "!"
End Sub