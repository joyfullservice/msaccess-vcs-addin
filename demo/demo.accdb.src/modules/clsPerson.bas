Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit


Public FirstName As String
Public LastName As String


Public Property Get FullName() As String
    FullName = FirstName & " " & LastName
End Property