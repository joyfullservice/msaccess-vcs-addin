Attribute VB_Name = "modTestJsonConverter"
'---------------------------------------------------------------------------------------
' Module    : modTestJsonConverter
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modJsonConverter ParseJson/ConvertToJson.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.JSON")


Public Sub TestJsonRoundTrip_SimpleObject()
    Dim strJson As String
    Dim dResult As Dictionary
    strJson = "{""name"":""test"",""value"":42}"
    Set dResult = ParseJson(strJson)
    TestAssert dResult("name") = "test", "string value preserved"
    TestAssert dResult("value") = 42, "numeric value preserved"
End Sub


Public Sub TestJsonRoundTrip_NestedObject()
    Dim strJson As String
    Dim dResult As Dictionary
    strJson = "{""outer"":{""inner"":""deep""}}"
    Set dResult = ParseJson(strJson)
    TestAssert dResult("outer")("inner") = "deep", "nested value accessible"
End Sub


Public Sub TestJsonRoundTrip_Array()
    Dim strJson As String
    Dim colResult As Collection
    strJson = "[1,2,3]"
    Set colResult = ParseJson(strJson)
    TestAssert colResult.Count = 3, "array count"
    TestAssert colResult(1) = 1, "first element"
    TestAssert colResult(3) = 3, "last element"
End Sub


Public Sub TestJsonRoundTrip_EmptyObject()
    Dim dResult As Dictionary
    Set dResult = ParseJson("{}")
    TestAssert dResult.Count = 0, "empty object has no keys"
End Sub


Public Sub TestJsonRoundTrip_SpecialChars()
    Dim d As Dictionary
    Dim strJson As String
    Dim dResult As Dictionary
    Set d = New Dictionary
    d.Add "text", "line1" & vbCrLf & "line2"
    strJson = ConvertToJson(d)
    Set dResult = ParseJson(strJson)
    TestAssert dResult("text") = "line1" & vbCrLf & "line2", "newlines preserved"
End Sub


Public Sub TestJsonNewLineIssue()
    Dim strTest As String
    strTest = "Line1" & vbCrLf & "Line2" & vbCr & "Line3" & vbLf & "Line4" & vbCrLf

    Dim dTest As Dictionary
    Set dTest = New Dictionary
    dTest("Multiline") = strTest
    TestAssert dTest("Multiline") = strTest, "dictionary stores correctly"

    Dim strResult As String
    strResult = ParseJson(ConvertToJson(dTest, 2))("Multiline")
    TestAssert strResult = strTest, "round trip preserves all line ending types"
End Sub


Public Sub TestConvertToJson_NullValue()
    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "key", Null
    Dim strJson As String
    strJson = ConvertToJson(d)
    TestAssert InStr(strJson, "null") > 0, "null serialized"
End Sub
