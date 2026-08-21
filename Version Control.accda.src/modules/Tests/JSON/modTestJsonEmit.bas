Attribute VB_Name = "modTestJsonEmit"
'---------------------------------------------------------------------------------------
' Module    : modTestJsonEmit
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Unit tests for modJsonEmit fast serializers (test-state and results).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.JSON")
'@Tag("unit")


Public Sub TestEscapeJsonString_FastPath()
    TestAssert modJsonEmit.EscapeJsonString("modTestFoo.TestBar") = "modTestFoo.TestBar", _
        "simple identifier unchanged"
End Sub


Public Sub TestEscapeJsonString_QuotesAndBackslash()
    TestAssert modJsonEmit.EscapeJsonString("say ""hi""") = _
        "say \" & Chr(34) & "hi\" & Chr(34), "quotes escaped"
    TestAssert modJsonEmit.EscapeJsonString("C:\temp\file") = "C:\\temp\\file", "backslashes escaped"
End Sub


Public Sub TestEscapeJsonString_ControlChars()
    TestAssert modJsonEmit.EscapeJsonString("a" & vbLf & "b") = "a\nb", "newline escaped"
    TestAssert modJsonEmit.EscapeJsonString("") = vbNullString, "empty string"
End Sub


Public Sub TestEmitTestStateJson_RoundTrip()
    Dim dRoot As Dictionary
    Dim dParsed As Dictionary
    Dim dExpected As Dictionary

    Set dRoot = BuildSampleStateRoot()
    Set dExpected = ParseJson(ConvertToJson(dRoot, JSON_WHITESPACE))
    Set dParsed = ParseJson(modJsonEmit.EmitTestStateJson(dRoot))
    TestAssertJsonEqual dExpected, dParsed, "state emitter round trip"
End Sub


Public Sub TestEmitTestResultsJson_RoundTrip()
    Dim dTests As Dictionary
    Dim dParsed As Dictionary
    Dim dSummary As Dictionary
    Dim dExpectedTests As Dictionary
    Dim dExpectedSummary As Dictionary
    Dim dActualTests As Dictionary
    Dim dActualSummary As Dictionary

    Set dTests = BuildSampleRunnerTests()
    Set dExpectedTests = BuildSampleResultsTestsDict()
    Set dSummary = New Dictionary
    dSummary.Add "subs", 1&
    dSummary.Add "assertions", 2&
    dSummary.Add "passed", 2&
    dSummary.Add "failed", 0&
    dSummary.Add "errored", 0&
    dSummary.Add "empty", 0&
    Set dExpectedSummary = dSummary

    Set dParsed = ParseJson(modJsonEmit.EmitTestResultsJson(dTests, 1234, False, True, _
        1, 0, 0, 0, 2, 2, 0, "C:\out\test-results.xml"))
    TestAssert TypeName(dParsed("tests")) = "Dictionary", "results.tests present"
    TestAssert TypeName(dParsed("summary")) = "Dictionary", "results.summary present"
    Set dActualTests = dParsed("tests")
    Set dActualSummary = dParsed("summary")
    CompareJsonDict dExpectedTests, dActualTests, "results.tests"
    CompareJsonDict dExpectedSummary, dActualSummary, "results.summary"
    TestAssert dParsed("durationMs") = 1234, "durationMs"
    TestAssert dParsed("cancelled") = False, "cancelled"
    TestAssert dParsed("allPassed") = True, "allPassed"
    TestAssert dParsed("junitPath") = "C:\out\test-results.xml", "junitPath"
End Sub


Public Sub TestEmitTestTreeJson_Grouping()

    Dim dParsed As Dictionary
    Dim dSuite As Dictionary
    Dim dNode As Dictionary
    Dim colNodes As Collection
    Dim colTags As Collection
    Dim varKeys As Variant

    Set dParsed = ParseJson(modJsonEmit.EmitTestTreeJson(BuildSampleTreeTests()))
    TestAssert dParsed.Count = 2, "two module suites"

    varKeys = dParsed.Keys
    TestAssert CStr(varKeys(0)) = "modTestAlpha", "suites keep scan order"

    Set dSuite = dParsed("modTestAlpha")
    TestAssert dSuite.Count = 3, "suite key count"
    TestAssert dSuite("name") = "modTestAlpha", "suite name"
    TestAssert dSuite("folder") = "Tests.Alpha", "suite folder taken from first test"

    Set colNodes = dSuite("tests")
    TestAssert colNodes.Count = 2, "both tests grouped under one suite"

    Set dNode = colNodes(1)
    TestAssert dNode.Count = 6, "node key count"
    TestAssert dNode("key") = "modTestAlpha.TestOne", "node key"
    TestAssert dNode("name") = "TestOne", "node name"
    TestAssert dNode("module") = "modTestAlpha", "node module"
    TestAssert dNode("procName") = "TestOne", "node procName"
    TestAssert dNode("lineNumber") = 10, "node lineNumber"
    Set colTags = dNode("tags")
    TestAssert colTags.Count = 1, "node tag count"
    TestAssert colTags(1) = "unit", "node tag value"

    Set dNode = colNodes(2)
    TestAssert dNode("key") = "modTestAlpha.TestTwo", "second node key"
    TestAssert dNode.Count = 5, "node without a line number omits lineNumber"
    Set colTags = dNode("tags")
    TestAssert colTags.Count = 0, "missing tags emit an empty array"

    Set dSuite = dParsed("modTestBeta")
    TestAssert dSuite("folder") = "Tests.Beta", "second suite folder"
    Set colNodes = dSuite("tests")
    TestAssert colNodes.Count = 1, "second suite test count"

End Sub


Public Sub TestEmitTestTreeJson_Empty()
    Dim dNone As Dictionary
    TestAssert modJsonEmit.EmitTestTreeJson(dNone) = "{}", "no runner tests emits an empty object"
End Sub


Public Sub TestEmitResultsBatchJson()

    Dim dParsed As Dictionary
    Dim dRecord As Dictionary
    Dim colResults As Collection
    Dim colAssertions As Collection
    Dim dAssertion As Dictionary
    Dim lngCount As Long

    Set dParsed = ParseJson(modJsonEmit.EmitResultsBatchJson(BuildSampleRunnerTests(), lngCount))
    TestAssert lngCount = 1, "one completed test counted"

    Set colResults = dParsed("results")
    TestAssert colResults.Count = 1, "one record in the batch"

    Set dRecord = colResults(1)
    TestAssert dRecord.Count = 5, "record key count"
    TestAssert dRecord("testKey") = "modTestJsonEmit.TestEmitTestResultsJson_RoundTrip", "testKey"
    TestAssert dRecord("status") = "pass", "runner status mapped to the web vocabulary"
    TestAssert dRecord("durationMs") = 7, "durationMs"
    TestAssert dRecord("prior") = False, "prior defaults to false"

    Set colAssertions = dRecord("assertions")
    TestAssert colAssertions.Count = 2, "assertion count"
    Set dAssertion = colAssertions(1)
    TestAssert dAssertion("seq") = 1, "assertion seq"
    TestAssert dAssertion("passed") = True, "assertion passed"
    TestAssert dAssertion("context") = "ctx\path", "assertion context unescaped on the way back"

End Sub


Public Sub TestEmitResultsBatchJson_NoCompletedTests()

    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim lngCount As Long

    Set dTests = New Dictionary
    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestAlpha"
    dTest.Add "procName", "TestOne"
    dTest.Add "status", etsPending
    Set dTests("modTestAlpha.TestOne") = dTest

    TestAssert modJsonEmit.EmitResultsBatchJson(dTests, lngCount) = vbNullString, _
        "a pending-only run emits nothing"
    TestAssert lngCount = 0, "nothing counted"

End Sub


Public Sub TestEmitWebResultJson_MatchesBatchRecord()

    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim dSingle As Dictionary
    Dim colResults As Collection
    Dim dFromBatch As Dictionary
    Dim strKey As String
    Dim lngCount As Long

    ' One shape, two entry points: the per-test push and the hydrate replay must agree.
    strKey = "modTestJsonEmit.TestEmitTestResultsJson_RoundTrip"
    Set dTests = BuildSampleRunnerTests()
    Set dTest = dTests(strKey)
    Set dSingle = ParseJson(modJsonEmit.EmitWebResultJson(strKey, dTest))
    Set colResults = ParseJson(modJsonEmit.EmitResultsBatchJson(dTests, lngCount))("results")
    Set dFromBatch = colResults(1)
    CompareJsonDict dSingle, dFromBatch, "web result record"

End Sub


Private Function BuildSampleTreeTests() As Dictionary

    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim colTags As Collection

    Set dTests = New Dictionary

    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestAlpha"
    dTest.Add "procName", "TestOne"
    dTest.Add "folder", "Tests.Alpha"
    dTest.Add "line", 10&
    Set colTags = New Collection
    colTags.Add "unit"
    Set dTest("tags") = colTags
    Set dTests("modTestAlpha.TestOne") = dTest

    ' No line and no tags: both are optional on a scanned record.
    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestAlpha"
    dTest.Add "procName", "TestTwo"
    dTest.Add "folder", "Tests.Alpha"
    Set dTests("modTestAlpha.TestTwo") = dTest

    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestBeta"
    dTest.Add "procName", "TestThree"
    dTest.Add "folder", "Tests.Beta"
    dTest.Add "line", 20&
    Set dTests("modTestBeta.TestThree") = dTest

    Set BuildSampleTreeTests = dTests

End Function


Private Function BuildSampleStateRoot() As Dictionary

    Dim dRoot As Dictionary
    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim dSummary As Dictionary
    Dim colTags As Collection
    Dim colAssertions As Collection
    Dim dA As Dictionary

    Set dRoot = New Dictionary
    dRoot.Add "runAt", "2026-08-18 09:00:00"
    dRoot.Add "databasePath", "C:\db.accdb"
    dRoot.Add "addinVersion", "5.1.0"
    dRoot.Add "sessionRunAt", "2026-08-18 09:00:00"

    Set dSummary = New Dictionary
    dSummary.Add "subs", 1&
    dSummary.Add "assertions", 2&
    dSummary.Add "passed", 2&
    dSummary.Add "failed", 0&
    dSummary.Add "errored", 0&
    dSummary.Add "empty", 0&
    Set dRoot("summary") = dSummary

    Set dTests = New Dictionary
    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestJsonEmit"
    dTest.Add "procName", "TestEmitTestStateJson_RoundTrip"
    dTest.Add "folder", "Tests.JSON"
    dTest.Add "line", 42&
    dTest.Add "sourceType", "Module"
    dTest.Add "status", "PASSED"
    dTest.Add "durationMs", 5&
    dTest.Add "lastRunAt", "2026-08-18 09:00:00"
    dTest.Add "stale", False
    dTest.Add "errorMessage", "say ""oops"" and \back\"

    Set colTags = New Collection
    colTags.Add "unit"
    colTags.Add "slow"
    Set dTest("tags") = colTags

    Set colAssertions = New Collection
    Set dA = New Dictionary
    dA.Add "seq", 1&
    dA.Add "passed", True
    dA.Add "context", "line1" & vbCrLf & "line2"
    colAssertions.Add dA
    Set dA = New Dictionary
    dA.Add "seq", 2&
    dA.Add "passed", True
    colAssertions.Add dA
    Set dTest("assertions") = colAssertions

    Set dTests("modTestJsonEmit.TestEmitTestStateJson_RoundTrip") = dTest
    Set dRoot("tests") = dTests
    Set BuildSampleStateRoot = dRoot

End Function


Private Function BuildSampleRunnerTests() As Dictionary

    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim colTags As Collection
    Dim colAssertions As Collection
    Dim colLogged As Collection
    Dim dA As Dictionary

    Set dTests = New Dictionary
    Set dTest = New Dictionary
    dTest.Add "moduleName", "modTestJsonEmit"
    dTest.Add "procName", "TestEmitTestResultsJson_RoundTrip"
    dTest.Add "folder", "Tests.JSON"
    dTest.Add "line", 99&
    dTest.Add "status", etsPassed
    dTest.Add "durationMs", 7&
    dTest.Add "errorMessage", vbNullString

    Set colTags = New Collection
    colTags.Add "unit"
    Set dTest("tags") = colTags

    Set colAssertions = New Collection
    Set dA = New Dictionary
    dA.Add "seq", 1&
    dA.Add "passed", True
    dA.Add "context", "ctx\path"
    colAssertions.Add dA
    Set dA = New Dictionary
    dA.Add "seq", 2&
    dA.Add "passed", True
    colAssertions.Add dA
    Set dTest("assertionResults") = colAssertions

    Set colLogged = New Collection
    Set dTest("loggedErrors") = colLogged

    Set dTests("modTestJsonEmit.TestEmitTestResultsJson_RoundTrip") = dTest
    Set BuildSampleRunnerTests = dTests

End Function


Private Function BuildSampleResultsTestsDict() As Dictionary

    Dim dTests As Dictionary
    Dim dTest As Dictionary
    Dim colTags As Collection
    Dim colAssertions As Collection
    Dim dA As Dictionary
    Dim colLogged As Collection

    Set dTests = New Dictionary
    Set dTest = New Dictionary
    dTest.Add "status", "PASSED"
    dTest.Add "durationMs", 7&
    dTest.Add "errorMessage", vbNullString
    Set colTags = New Collection
    colTags.Add "unit"
    Set dTest("tags") = colTags
    Set colAssertions = New Collection
    Set dA = New Dictionary
    dA.Add "seq", 1&
    dA.Add "passed", True
    dA.Add "context", "ctx\path"
    colAssertions.Add dA
    Set dA = New Dictionary
    dA.Add "seq", 2&
    dA.Add "passed", True
    colAssertions.Add dA
    Set dTest("assertions") = colAssertions
    Set colLogged = New Collection
    Set dTest("loggedErrors") = colLogged
    Set dTests("modTestJsonEmit.TestEmitTestResultsJson_RoundTrip") = dTest
    Set BuildSampleResultsTestsDict = dTests

End Function


Private Sub TestAssertJsonEqual(ByVal dExpected As Dictionary, ByVal dActual As Dictionary, _
    ByVal strLabel As String)

    TestAssert Not dExpected Is Nothing, strLabel & " expected root"
    TestAssert Not dActual Is Nothing, strLabel & " actual root"
    CompareJsonDict dExpected, dActual, strLabel

End Sub


' Values are never copied into a Variant with `=`: a Dictionary or Collection landing in
' a Variant that way invokes its default `Item` member and raises error 450. Dispatch on
' TypeName at the point of access instead, passing objects straight to typed parameters.
Private Sub CompareJsonDict(ByVal dExpected As Dictionary, ByVal dActual As Dictionary, _
    ByVal strPath As String)

    Dim varKey As Variant
    Dim strKey As String
    Dim strItem As String
    Dim dExpChild As Dictionary
    Dim dActChild As Dictionary
    Dim colExpChild As Collection
    Dim colActChild As Collection

    ' Compared both ways round: an emitter that writes an extra field has to fail too.
    TestAssert dExpected.Count = dActual.Count, strPath & " key count"

    For Each varKey In dExpected.Keys
        strKey = CStr(varKey)
        strItem = strPath & "." & strKey
        If Not dActual.Exists(strKey) Then
            TestAssert False, strItem & " missing"
        ElseIf TypeName(dExpected(strKey)) = "Dictionary" Then
            TestAssert TypeName(dActual(strKey)) = "Dictionary", strItem & " type"
            If TypeName(dActual(strKey)) = "Dictionary" Then
                Set dExpChild = dExpected(strKey)
                Set dActChild = dActual(strKey)
                CompareJsonDict dExpChild, dActChild, strItem
            End If
        ElseIf TypeName(dExpected(strKey)) = "Collection" Then
            TestAssert TypeName(dActual(strKey)) = "Collection", strItem & " type"
            If TypeName(dActual(strKey)) = "Collection" Then
                Set colExpChild = dExpected(strKey)
                Set colActChild = dActual(strKey)
                CompareJsonCollection colExpChild, colActChild, strItem
            End If
        Else
            TestAssert CStr(dExpected(strKey)) = CStr(dActual(strKey)), strItem & " value"
        End If
    Next varKey

End Sub


Private Sub CompareJsonCollection(ByVal colExpected As Collection, ByVal colActual As Collection, _
    ByVal strPath As String)

    Dim i As Long
    Dim strItem As String
    Dim dExpChild As Dictionary
    Dim dActChild As Dictionary
    Dim colExpChild As Collection
    Dim colActChild As Collection

    TestAssert colExpected.Count = colActual.Count, strPath & " count"
    If colExpected.Count <> colActual.Count Then Exit Sub

    For i = 1 To colExpected.Count
        strItem = strPath & "(" & CStr(i) & ")"
        If TypeName(colExpected(i)) = "Dictionary" Then
            TestAssert TypeName(colActual(i)) = "Dictionary", strItem & " type"
            If TypeName(colActual(i)) = "Dictionary" Then
                Set dExpChild = colExpected(i)
                Set dActChild = colActual(i)
                CompareJsonDict dExpChild, dActChild, strItem
            End If
        ElseIf TypeName(colExpected(i)) = "Collection" Then
            TestAssert TypeName(colActual(i)) = "Collection", strItem & " type"
            If TypeName(colActual(i)) = "Collection" Then
                Set colExpChild = colExpected(i)
                Set colActChild = colActual(i)
                CompareJsonCollection colExpChild, colActChild, strItem
            End If
        Else
            TestAssert CStr(colExpected(i)) = CStr(colActual(i)), strItem & " value"
        End If
    Next i

End Sub
