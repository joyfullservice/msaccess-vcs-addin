Attribute VB_Name = "modTestHash"
'---------------------------------------------------------------------------------------
' Module    : modTestHash
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modHash hashing functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.FileIO")


Public Sub TestGetStringHash()
    Dim strHash1 As String
    Dim strHash2 As String
    strHash1 = GetStringHash("test content")
    strHash2 = GetStringHash("test content")
    TestAssert Len(strHash1) > 0, "returns non-empty hash"
    TestAssert strHash1 = strHash2, "deterministic (same input = same output)"
End Sub


Public Sub TestGetStringHash_DifferentInputs()
    Dim strHash1 As String
    Dim strHash2 As String
    strHash1 = GetStringHash("input A")
    strHash2 = GetStringHash("input B")
    TestAssert strHash1 <> strHash2, "different inputs produce different hashes"
End Sub


Public Sub TestGetDictionaryHash()
    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Set d1 = New Dictionary
    Set d2 = New Dictionary
    d1.Add "key", "value"
    d2.Add "key", "value"
    TestAssert Len(GetDictionaryHash(d1)) > 0, "returns non-empty hash"
    TestAssert GetDictionaryHash(d1) = GetDictionaryHash(d2), "identical dictionaries same hash"
End Sub


Public Sub TestGetSimpleHash()
    Dim strHash As String
    strHash = GetSimpleHash("hello world")
    TestAssert Len(strHash) = 7, "returns 7-char short hash"
    TestAssert GetSimpleHash("hello world") = strHash, "deterministic"
    TestAssert GetSimpleHash("different") <> strHash, "different inputs differ"
End Sub
