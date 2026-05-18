VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestRunnerFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestRunnerFilters
' Author    : Adam Waller
' Date      : 5/18/2026
' Purpose   : Tests for the test runner's unified filter resolution and tag extraction.
'           : This module intentionally carries '@Tag annotations so it can verify
'           : the scan phase detects them correctly (self-referential test).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")
'@Tag("meta")


Private Sub Class_Initialize()
    ' Ensure the runner has scanned before each test method
    If TestRunner.State < etrsReady Then TestRunner.Scan
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestScanDetectsModuleLevelTags
' Purpose   : Verifies that module-level '@Tag annotations are discovered during scan.
'---------------------------------------------------------------------------------------
'
Public Sub TestScanDetectsModuleLevelTags()
    Dim dTest As Dictionary
    Dim strKey As String
    Dim colTags As Collection

    strKey = "clsTestRunnerFilters.TestScanDetectsModuleLevelTags"
    TestAssert TestRunner.Tests.Exists(strKey), "self-test key exists in Tests"

    Set dTest = TestRunner.Tests(strKey)
    Set colTags = dTest("tags")
    TestAssert colTags.Count >= 2, "at least 2 tags (unit, meta)"
    TestAssert CollectionContains(colTags, "unit"), "has 'unit' tag"
    TestAssert CollectionContains(colTags, "meta"), "has 'meta' tag"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestScanDetectsProcedureLevelTags
' Purpose   : Verifies that procedure-level '@Tag annotations are merged with module tags.
'---------------------------------------------------------------------------------------
'
Public Sub TestScanDetectsProcedureLevelTags()
    '@Tag("proconly")
    Dim dTest As Dictionary
    Dim colTags As Collection

    Set dTest = TestRunner.Tests("clsTestRunnerFilters.TestScanDetectsProcedureLevelTags")
    Set colTags = dTest("tags")
    TestAssert CollectionContains(colTags, "unit"), "inherits module-level 'unit'"
    TestAssert CollectionContains(colTags, "meta"), "inherits module-level 'meta'"
    TestAssert CollectionContains(colTags, "proconly"), "has procedure-level 'proconly'"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_ByModuleName
' Purpose   : ResolveFilters matches by module name (priority 1).
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_ByModuleName()
    Dim colKeys As Collection
    Set colKeys = TestRunner.ResolveFilters(Array("clsTestRunnerFilters"))
    TestAssert colKeys.Count > 0, "module filter returns matches"

    ' All returned keys should belong to this module
    Dim i As Long
    For i = 1 To colKeys.Count
        TestAssert Left$(colKeys(i), Len("clsTestRunnerFilters.")) = "clsTestRunnerFilters.", _
            "key " & colKeys(i) & " belongs to this module"
    Next i
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_BySuiteName
' Purpose   : ResolveFilters matches by suite folder segment (priority 2).
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_BySuiteName()
    Dim colKeys As Collection
    Set colKeys = TestRunner.ResolveFilters(Array("Infrastructure"))
    TestAssert colKeys.Count > 0, "suite filter 'Infrastructure' returns matches"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_ByProcedureName
' Purpose   : ResolveFilters matches by procedure name (priority 3).
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_ByProcedureName()
    Dim colKeys As Collection
    Set colKeys = TestRunner.ResolveFilters(Array("TestResolveFilters_ByProcedureName"))
    TestAssert colKeys.Count = 1, "procedure filter returns exactly 1 match"
    TestAssert colKeys(1) = "clsTestRunnerFilters.TestResolveFilters_ByProcedureName", _
        "correct key returned"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_ByTag
' Purpose   : ResolveFilters matches by tag (priority 4).
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_ByTag()
    Dim colKeys As Collection
    Set colKeys = TestRunner.ResolveFilters(Array("proconly"))
    TestAssert colKeys.Count = 1, "'proconly' tag matches exactly 1 test"
    TestAssert colKeys(1) = "clsTestRunnerFilters.TestScanDetectsProcedureLevelTags", _
        "correct test matched by tag"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_Exclusion
' Purpose   : Dash-prefix excludes matching tests from the result set.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_Exclusion()
    Dim colAll As Collection
    Dim colExcluded As Collection

    ' Get all tests in this module
    Set colAll = TestRunner.ResolveFilters(Array("clsTestRunnerFilters"))

    ' Exclude by the 'proconly' tag (removes one test)
    Set colExcluded = TestRunner.ResolveFilters(Array("clsTestRunnerFilters", "-proconly"))
    TestAssert colExcluded.Count = colAll.Count - 1, _
        "excluding 'proconly' removes 1 test (got " & colExcluded.Count & " vs " & colAll.Count & ")"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_ExclusionOnly
' Purpose   : When only exclusions are specified, the base set is all tests.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_ExclusionOnly()
    Dim colResult As Collection
    Set colResult = TestRunner.ResolveFilters(Array("-proconly"))
    ' Should be all tests minus the one tagged "proconly"
    TestAssert colResult.Count = TestRunner.Tests.Count - 1, _
        "exclusion-only removes from all tests"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_TagIsCaseInsensitive
' Purpose   : Tag matching is case-insensitive.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_TagIsCaseInsensitive()
    Dim colLower As Collection
    Dim colUpper As Collection
    Set colLower = TestRunner.ResolveFilters(Array("meta"))
    Set colUpper = TestRunner.ResolveFilters(Array("META"))
    TestAssert colLower.Count = colUpper.Count, "tag matching is case-insensitive"
    TestAssert colLower.Count > 0, "tag 'meta' returns matches"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveFilters_UnmatchedFilterReturnsEmpty
' Purpose   : A filter that matches nothing returns an empty collection.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveFilters_UnmatchedFilterReturnsEmpty()
    Dim colKeys As Collection
    Set colKeys = TestRunner.ResolveFilters(Array("nonexistent_xyzzy_tag_12345"))
    TestAssert colKeys.Count = 0, "unmatched filter returns empty collection"
End Sub


' ===================== Private helpers =====================


Private Function CollectionContains(col As Collection, strValue As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(CStr(col(i)), strValue, vbTextCompare) = 0 Then
            CollectionContains = True
            Exit Function
        End If
    Next i
End Function
