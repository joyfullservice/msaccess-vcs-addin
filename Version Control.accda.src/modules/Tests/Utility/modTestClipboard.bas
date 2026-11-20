Attribute VB_Name = "modTestClipboard"
'---------------------------------------------------------------------------------------
' Module    : modTestClipboard
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Unit tests for the modClipboard text helpers.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")
'@Tag("clipboard")


Public Sub TestClipboardRoundTrip()
    Dim strOriginal As String
    strOriginal = GetClipboardText()

    Dim strExpected As String
    strExpected = "C:\Test\Path\file.log"
    TestAssert SetClipboardText(strExpected), "Clipboard write should succeed"
    TestAssert GetClipboardText() = strExpected, "Round-trip text should match"

    SetClipboardText strOriginal
End Sub


Public Sub TestClipboardWithSpecialChars()
    Dim strOriginal As String
    strOriginal = GetClipboardText()

    Dim strExpected As String
    strExpected = "C:\Users\Name With Spaces\logs\TestResults_20260518_120000_001.json"
    TestAssert SetClipboardText(strExpected), "Clipboard write should succeed"
    TestAssert GetClipboardText() = strExpected, "Paths with spaces should round-trip"

    SetClipboardText strOriginal
End Sub
