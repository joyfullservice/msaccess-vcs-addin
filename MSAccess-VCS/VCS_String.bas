Attribute VB_Name = "VCS_String"
Option Compare Database
Option Private Module
Option Explicit

'--------------------
' String Functions: String Builder,String Padding (right only), Substrings
'--------------------

' String builder: Init
Public Function Sb_Init() As String()
    Dim x(-1 To -1) As String
    Sb_Init = x
End Function

' String builder: Clear
Public Sub Sb_Clear(ByRef sb() As String)
    ReDim sb(-1 To -1)
End Sub

' String builder: Append
Public Sub Sb_Append(ByRef sb() As String, Value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = Value
End Sub

' String builder: Get value
Public Function Sb_Get(ByRef sb() As String) As String
    Sb_Get = Join(sb, "")
End Function


' Pad a string on the right to make it `count` characters long.
Public Function PadRight(Value As String, Count As Integer)
    PadRight = Value
    If Len(Value) < Count Then
        PadRight = PadRight & Space(Count - Len(Value))
    End If
End Function

' returns substring between e.g. "(" and ")", internal brackets ar skippped
Public Function SubString(p As Integer, s As String, startsWith As String, endsWith As String)
    Dim start As Integer
    Dim cursor As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    Dim level As Integer
    start = InStr(p, s, startsWith)
    level = 1
    p1 = InStr(start + 1, s, startsWith)
    p2 = InStr(start + 1, s, endsWith)
    While level > 0
        If p1 > p2 And p2 > 0 Then
            cursor = p2
            level = level - 1
        ElseIf p2 > p1 And p1 > 0 Then
            cursor = p1
            level = level + 1
        ElseIf p2 > 0 And p1 = 0 Then
            cursor = p2
            level = level - 1
        ElseIf p1 > 0 And p1 = 0 Then
            cursor = p1
            level = level + 1
        ElseIf p1 = 0 And p2 = 0 Then
            SubString = ""
            Exit Function
        End If
        p1 = InStr(cursor + 1, s, startsWith)
        p2 = InStr(cursor + 1, s, endsWith)
    Wend
    SubString = Mid(s, start + 1, cursor - start - 1)
End Function



