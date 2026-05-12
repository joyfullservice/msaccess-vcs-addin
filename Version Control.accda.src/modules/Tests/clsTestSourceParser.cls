VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSourceParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSourceParser
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsSourceParser content loading, output, and sanitization.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Core")

Private m_parser As clsSourceParser


Private Sub Class_Initialize()
    Set m_parser = New clsSourceParser
End Sub


Private Sub Class_Terminate()
    Set m_parser = Nothing
End Sub


Public Sub TestLoadString_ReturnsOutput()
    Dim strContent As String
    strContent = "Begin Form" & vbCrLf & "  Caption =""Test""" & vbCrLf & "End"
    m_parser.LoadString strContent, edbForm
    TestAssert Len(m_parser.GetOutput) > 0, "GetOutput returns content"
End Sub


Public Sub TestHash_Deterministic()
    Dim strContent As String
    strContent = "Some content for hashing"
    m_parser.LoadString strContent, edbModule
    Dim strHash1 As String
    strHash1 = m_parser.Hash()
    m_parser.LoadString strContent, edbModule
    TestAssert m_parser.Hash() = strHash1, "same content produces same hash"
End Sub


Public Sub TestStripClassHeader()
    Dim strContent As String
    Dim strResult As String
    strContent = "VERSION 1.0 CLASS" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "  MultiUse = -1" & vbCrLf & _
        "END" & vbCrLf & _
        "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database"
    strResult = m_parser.StripClassHeader(strContent, False)
    TestAssert InStr(strResult, "VERSION 1.0 CLASS") = 0, "class header stripped"
    TestAssert InStr(strResult, "Option Compare Database") > 0, "code content preserved"
End Sub


Public Sub TestStripClassHeader_NameOnly()
    Dim strContent As String
    Dim strResult As String
    strContent = "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database" & vbCrLf & _
        "Option Explicit"
    strResult = m_parser.StripClassHeader(strContent, True)
    TestAssert InStr(strResult, "VB_Name") = 0, "VB_Name attribute stripped"
    TestAssert InStr(strResult, "Option Compare") > 0, "code preserved"
End Sub
