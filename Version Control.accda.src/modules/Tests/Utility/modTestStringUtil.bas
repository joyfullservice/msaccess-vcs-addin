Attribute VB_Name = "modTestStringUtil"
'---------------------------------------------------------------------------------------
' Module    : modTestStringUtil
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modStringUtil string manipulation functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")


Public Sub TestStartsWith()
    TestAssert StartsWith("Hello", "He"), "match at start"
    TestAssert StartsWith("Hello", "Hello"), "exact match"
    TestAssert Not StartsWith("Hello", "Wo"), "no match"
    TestAssert Not StartsWith("Hello", "hello"), "case sensitive by default"
    TestAssert StartsWith("Hello", "hello", vbTextCompare), "case insensitive"
    TestAssert StartsWith("", ""), "both empty"
    TestAssert Not StartsWith("", "x"), "empty haystack"
    TestAssert StartsWith("x", ""), "empty needle"
End Sub


Public Sub TestEndsWith()
    TestAssert EndsWith("Hello", "lo"), "match at end"
    TestAssert EndsWith("Hello", "Hello"), "exact match"
    TestAssert Not EndsWith("Hello", "He"), "no match"
    TestAssert Not EndsWith("Hello", "LO"), "case sensitive by default"
    TestAssert EndsWith("Hello", "LO", vbTextCompare), "case insensitive"
    TestAssert EndsWith("", ""), "both empty"
    TestAssert Not EndsWith("", "x"), "empty haystack"
    TestAssert EndsWith("x", ""), "empty needle"
End Sub


Public Sub TestMultiReplace()
    TestAssert MultiReplace("abcabc", "a", "x", "c", "z") = "xbzxbz", "multiple replacements"
    TestAssert MultiReplace("test", "x", "y") = "test", "no match leaves unchanged"
    TestAssert MultiReplace("", "a", "b") = "", "empty string"
End Sub


Public Sub TestCoalesce()
    TestAssert Coalesce("first", "second") = "first", "returns first non-empty"
    TestAssert Coalesce("", "second") = "second", "skips empty"
    TestAssert Coalesce("", "", "third") = "third", "skips multiple empty"
    TestAssert Coalesce("", "") = "", "all empty returns empty"
End Sub


Public Sub TestDblQ()
    TestAssert DblQ("it's") = "it''s", "doubles single quotes"
    TestAssert DblQ("say ""hi""") = "say """"hi""""", "doubles double quotes"
    TestAssert DblQ("plain") = "plain", "no quotes unchanged"
    TestAssert DblQ("") = "", "empty string"
End Sub


Public Sub TestDeDupString()
    TestAssert DeDupString("aa", "a") = "a", "two to one"
    TestAssert DeDupString("aaaa", "a") = "a", "four to one"
    TestAssert DeDupString("abab", "ab") = "ab", "multi-char duplicate"
    TestAssert DeDupString("abc", "x") = "abc", "no duplicates"
    TestAssert DeDupString("", "a") = "", "empty string"
End Sub


Public Sub TestRepeat()
    TestAssert Repeat("ab", 3) = "ababab", "repeat 3 times"
    TestAssert Repeat("x", 1) = "x", "repeat once"
    TestAssert Repeat("x", 0) = "", "repeat zero times"
End Sub


Public Sub TestLikeAny()
    TestAssert LikeAny("test.bas", "*.bas", "*.cls"), "matches first pattern"
    TestAssert LikeAny("test.cls", "*.bas", "*.cls"), "matches second pattern"
    TestAssert Not LikeAny("test.txt", "*.bas", "*.cls"), "no match"
    TestAssert LikeAny("ABC", "[A-Z]*"), "wildcard match"
End Sub
