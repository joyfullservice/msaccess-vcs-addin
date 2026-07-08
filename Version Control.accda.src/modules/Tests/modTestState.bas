Attribute VB_Name = "modTestState"
'---------------------------------------------------------------------------------------
' Module    : modTestState
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Durable merged test-state persistence for the VCS test runner. Maintains a
'           : single test-state.json under <export-folder>/test-results/ that reflects
'           : the latest known status of every test, merged across full and partial runs.
'           : Survives Access restarts and VBA state resets; the web runner reloads
'           : from this file when the in-memory singleton is empty.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Private Const ModuleName As String = "modTestState"
Private Const STATE_FILE As String = "test-state.json"


'---------------------------------------------------------------------------------------
' Procedure : GetTestResultsFolder
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Resolved path to the durable test-results folder (created if missing).
'---------------------------------------------------------------------------------------
'
Public Function GetTestResultsFolder() As String

    GetTestResultsFolder = Options.GetExportFolder & "test-results" & PathSep
    VerifyPath GetTestResultsFolder

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetStateFilePath
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Full path to the durable test-state.json file.
'---------------------------------------------------------------------------------------
'
Public Function GetStateFilePath() As String

    GetStateFilePath = GetTestResultsFolder() & STATE_FILE

End Function


'---------------------------------------------------------------------------------------
' Procedure : PersistAfterRun
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Merge the current run into test-state.json and optionally emit JUnit XML.
'---------------------------------------------------------------------------------------
'
Public Sub PersistAfterRun()

    MergeAndSave
    If Options.ExportTestResultsJUnit Then modTestJUnit.ExportFromState
    If Options.ExportTestResultsHtml Then modTestReport.ExportResultsHtml

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeAndSave
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Merge the current scan/run into test-state.json. Executed tests in the
'           : latest run are updated with a fresh lastRunAt; other known tests keep their
'           : prior status and are flagged stale.
'---------------------------------------------------------------------------------------
'
Public Sub MergeAndSave()

    Dim dRoot As Dictionary
    Dim dTestsOut As Dictionary
    Dim dSummary As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim dOldTests As Dictionary
    Dim dOldEntry As Dictionary
    Dim dOut As Dictionary
    Dim strSessionRunAt As String
    Dim strKey As String
    Dim blnExecuted As Boolean

    If TestRunner.Tests Is Nothing Then Exit Sub
    If TestRunner.Tests.Count = 0 Then Exit Sub

    strSessionRunAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set dOldTests = LoadStateTestsDict()

    Set dRoot = New Dictionary
    dRoot.Add "runAt", strSessionRunAt
    dRoot.Add "databasePath", CurrentProject.FullName
    dRoot.Add "addinVersion", GetVCSVersion
    dRoot.Add "sessionRunAt", strSessionRunAt

    Set dTestsOut = New Dictionary
    For Each varKey In TestRunner.Tests.Keys
        strKey = CStr(varKey)
        Set dTest = TestRunner.Tests(strKey)
        blnExecuted = WasExecutedThisRun(strKey, dTest)

        If blnExecuted Then
            Set dOut = SerializeTestRecord(strKey, dTest, strSessionRunAt, False)
            Set dTestsOut(strKey) = dOut
        ElseIf Not dOldTests Is Nothing And dOldTests.Exists(strKey) Then
            Set dOut = CopyStateEntry(dOldTests(strKey), True)
            Set dTestsOut(strKey) = dOut
        Else
            Set dOut = SerializePendingRecord(strKey, dTest)
            Set dTestsOut(strKey) = dOut
        End If
    Next varKey

    Set dSummary = BuildSummaryFromState(dTestsOut)
    Set dRoot("summary") = dSummary
    Set dRoot("tests") = dTestsOut

    WriteFile ConvertToJson(dRoot, JSON_WHITESPACE), GetStateFilePath()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadState
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Load the durable state root dictionary from disk (Nothing when absent).
'---------------------------------------------------------------------------------------
'
Public Function LoadState() As Dictionary

    Dim strPath As String
    Dim dRoot As Dictionary

    strPath = GetStateFilePath()
    If Not FSO.FileExists(strPath) Then Exit Function

    On Error GoTo LoadErr
    Set dRoot = ParseJson(ReadFile(strPath))
    If TypeName(dRoot) = "Dictionary" Then
        Set LoadState = dRoot
    End If
    Exit Function

LoadErr:
    Log.Add T("Could not load test state file: {0}", var0:=strPath), , , "orange"
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadInto
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Reconstruct the TestRunner singleton from test-state.json.
'---------------------------------------------------------------------------------------
'
Public Sub LoadInto(tr As clsTestRunner)

    Dim dRoot As Dictionary
    Dim dTests As Dictionary

    Set dRoot = LoadState()
    If dRoot Is Nothing Then Exit Sub
    If Not dRoot.Exists("tests") Then Exit Sub
    If TypeName(dRoot("tests")) <> "Dictionary" Then Exit Sub

    Set dTests = dRoot("tests")
    tr.LoadStateTests dTests

End Sub


' ===================== Private helpers =====================


Private Function LoadStateTestsDict() As Dictionary

    Dim dRoot As Dictionary

    Set dRoot = LoadState()
    If dRoot Is Nothing Then Exit Function
    If Not dRoot.Exists("tests") Then Exit Function
    If TypeName(dRoot("tests")) <> "Dictionary" Then Exit Function

    Set LoadStateTestsDict = dRoot("tests")

End Function


Private Function WasExecutedThisRun(ByVal strKey As String, ByVal dTest As Dictionary) As Boolean

    Dim colKeys As Collection
    Dim i As Long

    If CLng(Nz(dTest("status"), etsPending)) = etsPending Then Exit Function

    Set colKeys = TestRunner.LastRunKeys
    If colKeys Is Nothing Then Exit Function

    For i = 1 To colKeys.Count
        If StrComp(CStr(colKeys(i)), strKey, vbTextCompare) = 0 Then
            WasExecutedThisRun = True
            Exit Function
        End If
    Next i

End Function


Private Function SerializeTestRecord(ByVal strKey As String, ByVal dTest As Dictionary, _
    ByVal strLastRunAt As String, ByVal blnStale As Boolean) As Dictionary

    Dim dOut As Dictionary
    Dim colTagsOut As Collection
    Dim colAssertOut As Collection
    Dim colAssertions As Collection
    Dim colLoggedOut As Collection
    Dim colLoggedErrors As Collection
    Dim dA As Dictionary
    Dim dAOut As Dictionary
    Dim dErr As Dictionary
    Dim dErrOut As Dictionary
    Dim colTags As Collection
    Dim i As Long

    Set dOut = New Dictionary
    dOut.Add "moduleName", CStr(dTest("moduleName"))
    dOut.Add "procName", CStr(dTest("procName"))
    dOut.Add "folder", CStr(dTest("folder"))
    dOut.Add "line", CLng(Nz(dTest("line"), 0))
    If dTest.Exists("sourceType") Then
        dOut.Add "sourceType", CStr(dTest("sourceType"))
    End If
    dOut.Add "status", StatusToString(CLng(dTest("status")))
    dOut.Add "durationMs", CLng(Nz(dTest("durationMs"), 0))
    dOut.Add "lastRunAt", strLastRunAt
    dOut.Add "stale", blnStale

    If dTest.Exists("errorMessage") Then
        If Len(CStr(dTest("errorMessage"))) > 0 Then
            dOut.Add "errorMessage", CStr(dTest("errorMessage"))
        End If
    End If

    Set colTagsOut = New Collection
    If dTest.Exists("tags") Then
        Set colTags = dTest("tags")
        For i = 1 To colTags.Count
            colTagsOut.Add CStr(colTags(i))
        Next i
    End If
    Set dOut("tags") = colTagsOut

    Set colAssertOut = New Collection
    If dTest.Exists("assertionResults") Then
        Set colAssertions = dTest("assertionResults")
        For i = 1 To colAssertions.Count
            Set dA = colAssertions(i)
            Set dAOut = New Dictionary
            dAOut.Add "seq", dA("seq")
            dAOut.Add "passed", dA("passed")
            If Len(CStr(Nz(dA("context"), vbNullString))) > 0 Then
                dAOut.Add "context", CStr(dA("context"))
            End If
            colAssertOut.Add dAOut
        Next i
    End If
    Set dOut("assertions") = colAssertOut

    If dTest.Exists("loggedErrors") Then
        Set colLoggedOut = New Collection
        Set colLoggedErrors = dTest("loggedErrors")
        For i = 1 To colLoggedErrors.Count
            Set dErr = colLoggedErrors(i)
            Set dErrOut = New Dictionary
            dErrOut.Add "level", dErr("level")
            dErrOut.Add "message", dErr("message")
            If Len(CStr(Nz(dErr("source"), vbNullString))) > 0 Then
                dErrOut.Add "source", CStr(dErr("source"))
            End If
            If CLng(Nz(dErr("errNumber"), 0)) <> 0 Then
                dErrOut.Add "errNumber", CLng(dErr("errNumber"))
            End If
            If Len(CStr(Nz(dErr("errDescription"), vbNullString))) > 0 Then
                dErrOut.Add "errDescription", CStr(dErr("errDescription"))
            End If
            colLoggedOut.Add dErrOut
        Next i
        Set dOut("loggedErrors") = colLoggedOut
    End If

    Set SerializeTestRecord = dOut

End Function


Private Function SerializePendingRecord(ByVal strKey As String, ByVal dTest As Dictionary) As Dictionary

    Dim dOut As Dictionary
    Dim colTagsOut As Collection
    Dim colTags As Collection
    Dim i As Long

    Set dOut = New Dictionary
    dOut.Add "moduleName", CStr(dTest("moduleName"))
    dOut.Add "procName", CStr(dTest("procName"))
    dOut.Add "folder", CStr(dTest("folder"))
    dOut.Add "line", CLng(Nz(dTest("line"), 0))
    If dTest.Exists("sourceType") Then
        dOut.Add "sourceType", CStr(dTest("sourceType"))
    End If
    dOut.Add "status", "PENDING"
    dOut.Add "durationMs", CLng(0)
    dOut.Add "stale", False

    Set colTagsOut = New Collection
    If dTest.Exists("tags") Then
        Set colTags = dTest("tags")
        For i = 1 To colTags.Count
            colTagsOut.Add CStr(colTags(i))
        Next i
    End If
    Set dOut("tags") = colTagsOut
    Set dOut("assertions") = New Collection

    Set SerializePendingRecord = dOut

End Function


Private Function CopyStateEntry(ByVal dOld As Dictionary, ByVal blnStale As Boolean) As Dictionary

    Dim dOut As Dictionary
    Dim varKey As Variant

    Set dOut = New Dictionary
    For Each varKey In dOld.Keys
        dOut.Add CStr(varKey), dOld(CStr(varKey))
    Next varKey
    dOut("stale") = blnStale

    Set CopyStateEntry = dOut

End Function


Private Function BuildSummaryFromState(ByVal dTests As Dictionary) As Dictionary

    Dim dSummary As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim lngSubs As Long
    Dim lngPassed As Long
    Dim lngFailed As Long
    Dim lngErrored As Long
    Dim lngEmpty As Long
    Dim lngAssertions As Long
    Dim lngPassedAssertions As Long
    Dim lngFailedAssertions As Long
    Dim colAssertions As Collection
    Dim dA As Dictionary
    Dim i As Long
    Dim strStatus As String

    Set dSummary = New Dictionary

    For Each varKey In dTests.Keys
        Set dTest = dTests(CStr(varKey))
        strStatus = UCase$(CStr(Nz(dTest("status"), "PENDING")))
        If strStatus = "PENDING" Then GoTo NextSummaryTest

        lngSubs = lngSubs + 1
        Select Case strStatus
            Case "PASSED": lngPassed = lngPassed + 1
            Case "FAILED": lngFailed = lngFailed + 1
            Case "ERRORED": lngErrored = lngErrored + 1
            Case "EMPTY": lngEmpty = lngEmpty + 1
        End Select

        If dTest.Exists("assertions") Then
            If TypeName(dTest("assertions")) = "Collection" Then
                Set colAssertions = dTest("assertions")
                For i = 1 To colAssertions.Count
                    Set dA = colAssertions(i)
                    lngAssertions = lngAssertions + 1
                    If CBool(Nz(dA("passed"), False)) Then
                        lngPassedAssertions = lngPassedAssertions + 1
                    Else
                        lngFailedAssertions = lngFailedAssertions + 1
                    End If
                Next i
            End If
        End If
NextSummaryTest:
    Next varKey

    dSummary.Add "subs", lngSubs
    dSummary.Add "assertions", lngAssertions
    dSummary.Add "passed", lngPassedAssertions
    dSummary.Add "failed", lngFailedAssertions
    dSummary.Add "errored", lngErrored
    dSummary.Add "empty", lngEmpty

    Set BuildSummaryFromState = dSummary

End Function


Private Function StatusToString(lngStatus As Long) As String

    Select Case lngStatus
        Case etsPassed:  StatusToString = "PASSED"
        Case etsFailed:  StatusToString = "FAILED"
        Case etsErrored: StatusToString = "ERRORED"
        Case etsEmpty:   StatusToString = "EMPTY"
        Case Else:       StatusToString = "PENDING"
    End Select

End Function
