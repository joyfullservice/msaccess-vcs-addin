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
'@Tag("unit")


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


Public Sub TestConvertToJson_ExactNestedPrettyOutput()

    Dim dRoot As Dictionary
    Dim dChild As Dictionary
    Dim colItems As Collection
    Dim strExpected As String

    Set dRoot = New Dictionary
    Set dChild = New Dictionary
    Set colItems = New Collection
    dChild.Add "enabled", True
    dChild.Add "count", 2&
    colItems.Add "alpha"
    colItems.Add Null
    dRoot.Add "name", "example"
    dRoot.Add "child", dChild
    dRoot.Add "items", colItems

    strExpected = "{" & vbCrLf & _
        "  ""name"": ""example""," & vbCrLf & _
        "  ""child"": {" & vbCrLf & _
        "    ""enabled"": true," & vbCrLf & _
        "    ""count"": 2" & vbCrLf & _
        "  }," & vbCrLf & _
        "  ""items"": [" & vbCrLf & _
        "    ""alpha""," & vbCrLf & _
        "    null" & vbCrLf & _
        "  ]" & vbCrLf & _
        "}"

    TestAssert StrComp(ConvertToJson(dRoot, 2), strExpected, vbBinaryCompare) = 0, _
        "nested pretty output is byte-identical"

End Sub


Public Sub TestConvertToJson_ExactArrayOutput()

    Dim avarOneDimensional As Variant
    Dim avarTwoDimensional(0 To 1, 0 To 1) As Variant

    avarOneDimensional = Array("alpha", 2&, False, Null)
    avarTwoDimensional(0, 0) = "a"
    avarTwoDimensional(0, 1) = 1&
    avarTwoDimensional(1, 0) = "b"
    avarTwoDimensional(1, 1) = 2&

    TestAssert StrComp(ConvertToJson(avarOneDimensional), _
        "[""alpha"",2,false,null]", vbBinaryCompare) = 0, "one-dimensional array"
    TestAssert StrComp(ConvertToJson(avarTwoDimensional), _
        "[[""a"",1],[""b"",2]]", vbBinaryCompare) = 0, "two-dimensional array"

End Sub


Public Sub TestConvertToJson_StringIndentation()

    Dim dValue As Dictionary
    Dim strExpected As String

    Set dValue = New Dictionary
    dValue.Add "key", "value"
    strExpected = "{" & vbCrLf & vbTab & """key"": ""value""" & vbCrLf & "}"

    TestAssert StrComp(ConvertToJson(dValue, vbTab), strExpected, vbBinaryCompare) = 0, _
        "string indentation output is byte-identical"

End Sub


Public Sub TestConvertToJson_ExactEscapeOptions()

    Dim blnOldAllowUnicode As Boolean
    Dim blnOldEscapeSolidus As Boolean
    Dim strInput As String
    Dim strExpected As String

    blnOldAllowUnicode = JsonOptions.AllowUnicodeChars
    blnOldEscapeSolidus = JsonOptions.EscapeSolidus
    strInput = "quote"" slash/ backslash\" & vbTab & ChrW$(233) & ChrW$(127)

    JsonOptions.AllowUnicodeChars = True
    JsonOptions.EscapeSolidus = False
    strExpected = """quote\"" slash/ backslash\\\t" & ChrW$(233) & "\u007F"""
    TestAssert StrComp(ConvertToJson(strInput), strExpected, vbBinaryCompare) = 0, _
        "raw Unicode and unescaped solidus"

    JsonOptions.AllowUnicodeChars = False
    JsonOptions.EscapeSolidus = True
    strExpected = """quote\"" slash\/ backslash\\\t\u00E9\u007F"""
    TestAssert StrComp(ConvertToJson(strInput), strExpected, vbBinaryCompare) = 0, _
        "escaped Unicode and solidus"

    JsonOptions.AllowUnicodeChars = blnOldAllowUnicode
    JsonOptions.EscapeSolidus = blnOldEscapeSolidus

End Sub


Public Sub TestConvertToJson_LargeNumberOption()

    Dim blnOldUseDouble As Boolean
    Dim strNumber As String

    blnOldUseDouble = JsonOptions.UseDoubleForLargeNumbers
    strNumber = "1234567890123456"

    JsonOptions.UseDoubleForLargeNumbers = False
    TestAssert ConvertToJson(strNumber) = strNumber, "large numeric string is unquoted"

    JsonOptions.UseDoubleForLargeNumbers = True
    TestAssert ConvertToJson(strNumber) = """" & strNumber & """", _
        "large numeric string stays quoted when doubles are allowed"

    JsonOptions.UseDoubleForLargeNumbers = blnOldUseDouble

End Sub


Public Sub TestConvertToJson_UndefinedAndEmptyContainers()

    Dim dValues As Dictionary
    Dim dEmpty As Dictionary
    Dim colValues As Collection
    Dim colEmpty As Collection

    Set dValues = New Dictionary
    Set dEmpty = New Dictionary
    Set colValues = New Collection
    Set colEmpty = New Collection

    dValues.Add "kept", 1&
    dValues.Add "omitted", Empty
    colValues.Add Empty

    TestAssert ConvertToJson(dValues) = "{""kept"":1}", _
        "undefined dictionary item is omitted"
    TestAssert ConvertToJson(colValues) = "[null]", _
        "undefined collection item becomes null"
    TestAssert ConvertToJson(dEmpty, 2) = "{" & vbCrLf & "}", _
        "pretty empty dictionary retains newline"
    TestAssert ConvertToJson(colEmpty, 2) = "[" & vbCrLf & "]", _
        "pretty empty collection retains newline"

End Sub


Public Sub TestConvertToJson_LocalDateOption()

    Dim blnOldConvertDate As Boolean
    Dim dtValue As Date

    blnOldConvertDate = JsonOptions.ConvertDateToIso
    JsonOptions.ConvertDateToIso = False
    dtValue = DateSerial(2026, 1, 2) + TimeSerial(3, 4, 5)

    TestAssert ConvertToJson(dtValue) = """" & CStr(dtValue) & """", _
        "local date serialization remains locale-compatible"

    JsonOptions.ConvertDateToIso = blnOldConvertDate

End Sub
