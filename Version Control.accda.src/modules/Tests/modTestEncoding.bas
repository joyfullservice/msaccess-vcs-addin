Attribute VB_Name = "modTestEncoding"
'---------------------------------------------------------------------------------------
' Module    : modTestEncoding
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modEncoding functions not covered by clsTestEncoding.
'           : clsTestEncoding covers round-trip conversion and file hashing.
'           : This module covers pure/near-pure functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.FileIO")


Public Sub TestStringHasExtendedASCII()
    TestAssert StringHasExtendedASCII("café"), "accented char detected"
    TestAssert StringHasExtendedASCII("ÆØÅ"), "Scandinavian chars detected"
    TestAssert Not StringHasExtendedASCII("Hello World"), "ASCII only"
    TestAssert Not StringHasExtendedASCII(""), "empty string"
    TestAssert Not StringHasExtendedASCII("abc123!@#"), "ASCII symbols"
End Sub


Public Sub TestGetSystemEncoding()
    Dim strEnc As String
    strEnc = GetSystemEncoding()
    TestAssert Len(strEnc) > 0, "returns non-empty encoding name"
    ' On most Western systems this will be Windows-1252 or similar
    TestAssert strEnc <> "utf-8" Or GetSystemEncoding(True) = "utf-8", _
        "UTF-8 only returned when allowed"
End Sub
