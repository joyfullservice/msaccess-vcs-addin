Attribute VB_Name = "VCS_String"
Option Compare Database

Option Private Module
Option Explicit


'--------------------
' String Functions: String Builder,String Padding (right only), Substrings
'--------------------

' String builder: Init
Public Function VCS_Sb_Init() As String()
    Dim x(-1 To -1) As String
    VCS_Sb_Init = x
End Function

' String builder: Clear
Public Sub VCS_Sb_Clear(ByRef sb() As String)
    ReDim VCS_Sb_Init(-1 To -1)
End Sub

' String builder: Append
Public Sub VCS_Sb_Append(ByRef sb() As String, ByVal Value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = Value
End Sub

' String builder: Get value
Public Function VCS_Sb_Get(ByRef sb() As String) As String
    VCS_Sb_Get = Join(sb, "")
End Function


' Pad a string on the right to make it `count` characters long.
Public Function VCS_PadRight(ByVal Value As String, ByVal Count As Integer) As String
    VCS_PadRight = Value
    If Len(Value) < Count Then
        VCS_PadRight = VCS_PadRight & Space$(Count - Len(Value))
    End If
End Function
