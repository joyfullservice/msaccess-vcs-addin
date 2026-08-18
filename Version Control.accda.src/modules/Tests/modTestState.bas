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

' Session cache for LoadState (keyed on path + file mtime + size).
Private m_dCachedState As Dictionary
Private m_strCachedPath As String
Private m_datCachedModified As Date
Private m_lngCachedSize As Long


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

    Dim dRoot As Dictionary

    modTestRunnerDiag.DiagBegin "persist"
    Set dRoot = MergeAndSave()
    If Options.ExportTestResultsJUnit Then
        modTestRunnerDiag.DiagBegin "persist.junit"
        modTestJUnit.ExportFromState , dRoot
        modTestRunnerDiag.DiagEnd "persist.junit"
    End If
    If Options.ExportTestResultsHtml Then
        modTestRunnerDiag.DiagBegin "persist.html"
        modTestReport.ExportResultsHtml
        modTestRunnerDiag.DiagEnd "persist.html"
    End If
    modTestRunnerDiag.DiagEnd "persist"

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
Public Function MergeAndSave() As Dictionary

    Dim dRoot As Dictionary
    Dim dTestsOut As Dictionary
    Dim dSummary As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim dOldTests As Dictionary
    Dim dOldEntry As Dictionary
    Dim dOut As Dictionary
    Dim dRunKeys As Dictionary
    Dim strSessionRunAt As String
    Dim strKey As String
    Dim blnExecuted As Boolean
    Dim strJson As String
    Dim strPath As String

    If TestRunner.Tests Is Nothing Then Exit Function
    If TestRunner.Tests.Count = 0 Then Exit Function

    modTestRunnerDiag.DiagBegin "state.merge"
    strSessionRunAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    modTestRunnerDiag.DiagBegin "state.load"
    Set dOldTests = LoadStateTestsDict()
    modTestRunnerDiag.DiagEnd "state.load"

    Set dRoot = New Dictionary
    dRoot.Add "runAt", strSessionRunAt
    dRoot.Add "databasePath", CurrentProject.FullName
    dRoot.Add "addinVersion", GetVCSVersion
    dRoot.Add "sessionRunAt", strSessionRunAt

    Set dTestsOut = New Dictionary
    Set dRunKeys = BuildLastRunKeySet()
    modTestRunnerDiag.DiagBegin "state.serialize"
    For Each varKey In TestRunner.Tests.Keys
        strKey = CStr(varKey)
        Set dTest = TestRunner.Tests(strKey)
        blnExecuted = WasExecutedThisRun(strKey, dTest, dRunKeys)

        If blnExecuted Then
            Set dOut = SerializeTestRecord(strKey, dTest, strSessionRunAt, False)
        ElseIf Not dOldTests Is Nothing Then
            If dOldTests.Exists(strKey) Then
                Set dOut = CopyStateEntry(dOldTests(strKey), True)
            Else
                Set dOut = SerializePendingRecord(strKey, dTest)
            End If
        Else
            Set dOut = SerializePendingRecord(strKey, dTest)
        End If
        Set dTestsOut(strKey) = dOut
    Next varKey
    modTestRunnerDiag.DiagEnd "state.serialize", "n=" & dTestsOut.Count

    modTestRunnerDiag.DiagBegin "state.summary"
    Set dSummary = BuildSummaryFromState(dTestsOut)
    modTestRunnerDiag.DiagEnd "state.summary"
    Set dRoot("summary") = dSummary
    Set dRoot("tests") = dTestsOut

    modTestRunnerDiag.DiagBegin "state.json"
    strJson = modJsonEmit.EmitTestStateJson(dRoot)
    modTestRunnerDiag.DiagEnd "state.json", "chars=" & Len(strJson)
    modTestRunnerDiag.DiagSize "state.json", Len(strJson)
    strPath = GetStateFilePath()
    modTestRunnerDiag.DiagBegin "state.write"
    WriteFile strJson, strPath
    modTestRunnerDiag.DiagEnd "state.write"
    SeedStateCache dRoot, strPath
    modTestRunnerDiag.DiagEnd "state.merge"
    Set MergeAndSave = dRoot

End Function


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
    Dim strJson As String

    strPath = GetStateFilePath()
    If Not FSO.FileExists(strPath) Then Exit Function

    If CacheMatches(strPath) Then
        Set LoadState = m_dCachedState
        Exit Function
    End If

    On Error GoTo LoadErr
    modTestRunnerDiag.DiagBegin "state.read"
    strJson = ReadFile(strPath)
    modTestRunnerDiag.DiagEnd "state.read", "chars=" & Len(strJson)
    modTestRunnerDiag.DiagSize "state.file", Len(strJson)
    modTestRunnerDiag.DiagBegin "state.parse"
    Set dRoot = ParseJson(strJson)
    modTestRunnerDiag.DiagEnd "state.parse"
    If TypeName(dRoot) = "Dictionary" Then
        SeedStateCache dRoot, strPath
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


'---------------------------------------------------------------------------------------
' Procedure : MergeInto
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Overlay durable state from test-state.json onto an already-scanned runner.
'---------------------------------------------------------------------------------------
'
Public Sub MergeInto(tr As clsTestRunner)

    Dim dTests As Dictionary

    Set dTests = LoadStateTestsDict()
    If dTests Is Nothing Then Exit Sub
    tr.MergeStateResults dTests

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PrefetchState
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Read and parse test-state.json into the session cache without applying it
'           : to anything. Called while the web runner waits on the Edge cold start,
'           : where VBA would otherwise only be pumping messages. The later MergeInto
'           : then hits the cache instead of paying for the parse.
'---------------------------------------------------------------------------------------
'
Public Sub PrefetchState()
    LoadState
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StateCached
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : True when the current test-state.json is already parsed in the session
'           : cache, so loading it is free. Lets the caller skip the deferred-hydrate
'           : indicator, which only exists to cover the cost of the parse.
'---------------------------------------------------------------------------------------
'
Public Function StateCached() As Boolean
    StateCached = CacheMatches(GetStateFilePath())
End Function


' ===================== Private helpers =====================


Private Function LoadStateTestsDict() As Dictionary

    Dim dRoot As Dictionary

    Set dRoot = LoadState()
    If dRoot Is Nothing Then Exit Function
    If Not dRoot.Exists("tests") Then Exit Function
    If TypeName(dRoot("tests")) <> "Dictionary" Then Exit Function

    Set LoadStateTestsDict = dRoot("tests")

End Function


Private Sub SeedStateCache(ByVal dRoot As Dictionary, ByVal strPath As String)

    Dim fil As Scripting.File

    On Error Resume Next
    Set m_dCachedState = dRoot
    m_strCachedPath = strPath
    If FSO.FileExists(strPath) Then
        Set fil = FSO.GetFile(strPath)
        m_datCachedModified = fil.DateLastModified
        m_lngCachedSize = fil.Size
    End If
    Err.Clear

End Sub


Private Function CacheMatches(ByVal strPath As String) As Boolean

    Dim fil As Scripting.File

    If m_dCachedState Is Nothing Then Exit Function
    If Len(m_strCachedPath) = 0 Then Exit Function
    If StrComp(m_strCachedPath, strPath, vbTextCompare) <> 0 Then Exit Function
    If Not FSO.FileExists(strPath) Then Exit Function

    Set fil = FSO.GetFile(strPath)
    CacheMatches = (fil.DateLastModified = m_datCachedModified And fil.Size = m_lngCachedSize)

End Function


Private Function WasExecutedThisRun(ByVal strKey As String, ByVal dTest As Dictionary, _
    dRunKeys As Dictionary) As Boolean

    If CLng(Nz(dTest("status"), etsPending)) = etsPending Then Exit Function
    If dRunKeys Is Nothing Then Exit Function

    WasExecutedThisRun = dRunKeys.Exists(strKey)

End Function


' Case-insensitive set of the keys executed in the last run. Built once per save: this is
' asked about every discovered test, and walking LastRunKeys per test made the merge
' quadratic in the size of the suite.
Private Function BuildLastRunKeySet() As Dictionary

    Dim colKeys As Collection
    Dim dKeys As Dictionary
    Dim strKey As String
    Dim i As Long

    Set dKeys = New Dictionary
    dKeys.CompareMode = TextCompare

    Set colKeys = TestRunner.LastRunKeys
    If Not colKeys Is Nothing Then
        For i = 1 To colKeys.Count
            strKey = CStr(colKeys(i))
            If Not dKeys.Exists(strKey) Then dKeys.Add strKey, True
        Next i
    End If

    Set BuildLastRunKeySet = dKeys

End Function


Private Function SerializeTestRecord(ByVal strKey As String, ByVal dTest As Dictionary, _
    ByVal strLastRunAt As String, ByVal blnStale As Boolean) As Dictionary

    Dim dOut As Dictionary
    Dim colTagsOut As Collection
    Dim colLoggedOut As Collection
    Dim colLoggedErrors As Collection
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

    ' Shared, not copied. The runner's assertion records carry exactly the fields the state
    ' shape needs, they are write-once, and rebuilding them here cost ~0.8 s per save: a
    ' Scripting.Dictionary costs ~0.4 ms to create in a live Access session, and a full run
    ' of this project produces over 2,000 assertions. See clsTestRunner.StateAssertions.
    If dTest.Exists("assertionResults") Then
        Set dOut("assertions") = dTest("assertionResults")
    Else
        Set dOut("assertions") = New Collection
    End If

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
