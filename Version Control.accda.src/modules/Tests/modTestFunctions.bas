Attribute VB_Name = "modTestFunctions"
'---------------------------------------------------------------------------------------
' Module    : modTestFunctions
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modFunctions general-purpose utility functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")


Public Sub TestGetSafeFileName()
    TestAssert GetSafeFileName("normal") = "normal", "no special chars"
    TestAssert InStr(GetSafeFileName("a<b"), "%3C") > 0, "encodes <"
    TestAssert InStr(GetSafeFileName("a>b"), "%3E") > 0, "encodes >"
    TestAssert InStr(GetSafeFileName("a:b"), "%3A") > 0, "encodes :"
    TestAssert InStr(GetSafeFileName("a""b"), "%22") > 0, "encodes double quote"
    TestAssert InStr(GetSafeFileName("a/b"), "%2F") > 0, "encodes /"
    TestAssert InStr(GetSafeFileName("a\b"), "%5C") > 0, "encodes \"
    TestAssert InStr(GetSafeFileName("a|b"), "%7C") > 0, "encodes |"
    TestAssert InStr(GetSafeFileName("a?b"), "%3F") > 0, "encodes ?"
    TestAssert InStr(GetSafeFileName("a*b"), "%2A") > 0, "encodes *"
    TestAssert InStr(GetSafeFileName("a%b"), "%25") > 0, "encodes % first"
End Sub


Public Sub TestGetSafeFileName_RoundTrip()
    Dim strOriginal As String
    strOriginal = "My<Query>:Test""File/Path\Pipe|Question?Star*Pct%"
    TestAssert GetOriginalFromSafeName(GetSafeFileName(strOriginal)) = strOriginal, _
        "round-trip preserves original"
End Sub


Public Sub TestNz2()
    TestAssert Nz2("value", "default") = "value", "non-empty returns first"
    TestAssert Nz2("", "default") = "default", "empty string returns second"
    TestAssert Nz2(0, 42) = 42, "zero returns second"
    TestAssert IsNull(Nz2(Null, Null)), "null with null"
    TestAssert Nz2(Null, "fallback") = "fallback", "null returns second"
    TestAssert Nz2("", "") = "", "both empty"
    TestAssert Nz2(99, 42) = 99, "non-zero returns first"
End Sub


Public Sub TestInArray()
    Dim varArray As Variant
    varArray = Array("a", "b", "c", 1, 2, 3)
    TestAssert InArray(varArray, "b"), "string match"
    TestAssert Not InArray(varArray, "B"), "case sensitive"
    TestAssert InArray(varArray, "B", vbTextCompare), "case insensitive"
    TestAssert InArray(varArray, 2), "numeric match"
    TestAssert Not InArray(varArray, "x"), "not found"
    TestAssert Not InArray(Null, "b"), "null array"
    TestAssert Not InArray(Array(), "b"), "empty array"
End Sub


Public Sub TestAddToArray()
    Dim varArr As Variant
    varArr = Array()
    AddToArray varArr, "first"
    TestAssert UBound(varArr) = 0, "first element at index 0"
    TestAssert varArr(0) = "first", "first element value"
    AddToArray varArr, "second"
    TestAssert UBound(varArr) = 1, "second element at index 1"
    TestAssert varArr(1) = "second", "second element value"
End Sub


Public Sub TestIsEmptyArray()
    Dim varEmpty As Variant
    varEmpty = Array()
    TestAssert IsEmptyArray(varEmpty), "empty array"
    TestAssert Not IsEmptyArray(Array(1, 2, 3)), "non-empty array"
    TestAssert Not IsEmptyArray("not an array"), "non-array returns false"
End Sub


Public Sub TestDatesClose()
    Dim dte1 As Date
    Dim dte2 As Date
    dte1 = Now
    dte2 = DateAdd("s", 1, dte1)
    TestAssert DatesClose(dte1, dte2), "1 second apart within default threshold"
    dte2 = DateAdd("s", 5, dte1)
    TestAssert Not DatesClose(dte1, dte2), "5 seconds apart outside default threshold"
    TestAssert DatesClose(dte1, dte2, 10), "5 seconds within custom 10s threshold"
    TestAssert DatesClose(dte1, dte1), "same date"
End Sub


Public Sub TestQuickSort()
    Dim astr() As String
    astr = Split("u i a")
    QuickSort astr
    TestAssert Join(astr, " ") = "a i u", "sorts alphabetically"

    Dim alng() As Variant
    alng = Array(3, 1, 4, 1, 5, 9, 2, 6)
    QuickSort alng
    TestAssert alng(0) = 1, "smallest first"
    TestAssert alng(UBound(alng)) = 9, "largest last"
End Sub


Public Sub TestBitSet()
    TestAssert BitSet(7, 1), "bit 1 set in 7"
    TestAssert BitSet(7, 2), "bit 2 set in 7"
    TestAssert BitSet(7, 4), "bit 4 set in 7"
    TestAssert Not BitSet(7, 8), "bit 8 not set in 7"
    TestAssert BitSet(0, 0), "zero flag in zero"
    TestAssert Not BitSet(0, 1), "bit 1 not set in 0"
End Sub


Public Sub TestSwapExtension()
    TestAssert SwapExtension("c:\test.bas", "cls") = "c:\test.cls", "swap .bas to .cls"
    TestAssert SwapExtension("file.txt", "json") = "file.json", "swap .txt to .json"
End Sub


Public Sub TestLargest()
    TestAssert Largest(1, 2, 3) = 3, "largest of ascending"
    TestAssert Largest(3, 2, 1) = 3, "largest of descending"
    TestAssert Largest(5) = 5, "single value"
    TestAssert Largest(-1, -2, -3) = -1, "negative values"
End Sub


Public Sub TestZN()
    TestAssert IsNull(ZN("")), "empty string returns null"
    TestAssert IsNull(ZN(0)), "zero returns null"
    TestAssert ZN("value") = "value", "non-empty passes through"
    TestAssert ZN(42) = 42, "non-zero passes through"
End Sub


Public Sub TestExpandEnvironmentVariables()
    Dim strResult As String
    strResult = ExpandEnvironmentVariables("%TEMP%")
    TestAssert Len(strResult) > 0, "TEMP expanded"
    TestAssert InStr(strResult, "%") = 0, "no percent signs remain"

    TestAssert ExpandEnvironmentVariables("no vars here") = "no vars here", "no vars unchanged"
    TestAssert ExpandEnvironmentVariables("") = "", "empty string"
End Sub
