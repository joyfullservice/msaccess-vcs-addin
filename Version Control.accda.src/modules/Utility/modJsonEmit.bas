Attribute VB_Name = "modJsonEmit"
'---------------------------------------------------------------------------------------
' Module    : modJsonEmit
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Fast, schema-specific JSON emission for fixed test-runner payloads. Avoids
'           : the per-character json_Encode path in modJsonConverter for large artifacts.
' Layer     : Utility
' Depends on: clsConcat, and clsTestRunner for the eTestStatus enum
' Note      : Field order and the rules for which optional fields are omitted must stay in
'           : step with modTestState.SerializeTestRecord and clsTestRunner.GetResultsAsJson,
'           : since agents and the runner UI parse these files. modTestJsonEmit round-trips
'           : both emitters against ConvertToJson to hold that line.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

Private Const ModuleName As String = "modJsonEmit"


'---------------------------------------------------------------------------------------
' Procedure : EscapeJsonString
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Escape a JSON string value (without surrounding quotes). Matches
'           : JsonOptions.AllowUnicodeChars = True: raw Unicode, escape only ", \,
'           : and codepoints 0-31 and 127-159.
'---------------------------------------------------------------------------------------
'
Public Function EscapeJsonString(strText As String) As String

    Dim lngIndex As Long
    Dim lngLen As Long
    Dim lngAsc As Long
    Dim strChar As String
    Dim buf As clsConcat
    Dim bytText() As Byte

    lngLen = Len(strText)
    If lngLen = 0 Then Exit Function

    ' Most values (test names, folders, statuses, timestamps) need no escaping at all, and
    ' the scan that proves it is what has to be cheap. Copying to a byte array and reading
    ' the UTF-16 code units directly avoids a Mid$ string allocation per character; only
    ' the low byte matters, since every escaped codepoint is below 256.
    bytText = strText
    For lngIndex = 0 To UBound(bytText) - 1 Step 2
        If bytText(lngIndex + 1) = 0 Then
            lngAsc = bytText(lngIndex)
            If lngAsc < 32 Or lngAsc = 34 Or lngAsc = 92 _
                Or (lngAsc >= 127 And lngAsc <= 159) Then GoTo BuildEscaped
        End If
    Next lngIndex
    EscapeJsonString = strText
    Exit Function

BuildEscaped:
    Set buf = New clsConcat
    For lngIndex = 1 To lngLen
        strChar = Mid$(strText, lngIndex, 1)
        lngAsc = AscW(strChar)
        If lngAsc < 0 Then lngAsc = lngAsc + 65536
        Select Case lngAsc
            Case 34: buf.Add "\"""
            Case 92: buf.Add "\\"
            Case 8: buf.Add "\b"
            Case 12: buf.Add "\f"
            Case 10: buf.Add "\n"
            Case 13: buf.Add "\r"
            Case 9: buf.Add "\t"
            Case 0 To 31, 127 To 159
                buf.Add "\u", Right$("0000" & Hex$(lngAsc), 4)
            Case Else
                buf.Add strChar
        End Select
    Next lngIndex
    EscapeJsonString = buf.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : EmitTestStateJson
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Serialize a merged test-state root dictionary. One line per test record;
'           : no pretty-print indentation.
'---------------------------------------------------------------------------------------
'
Public Function EmitTestStateJson(dRoot As Dictionary) As String

    Dim buf As clsConcat
    Dim dTests As Dictionary
    Dim dSummary As Dictionary
    Dim dTest As Dictionary
    Dim varKey As Variant
    Dim strKey As String
    Dim blnFirst As Boolean

    Set buf = New clsConcat
    buf.Add "{"
    AppendStringField buf, "runAt", CStr(dRoot("runAt")), True
    AppendStringField buf, "databasePath", CStr(dRoot("databasePath")), False
    AppendStringField buf, "addinVersion", CStr(dRoot("addinVersion")), False
    AppendStringField buf, "sessionRunAt", CStr(dRoot("sessionRunAt")), False
    buf.Add ","
    buf.Add """summary"":"
    Set dSummary = dRoot("summary")
    AppendStateSummary buf, dSummary
    buf.Add ","
    buf.Add """tests"":{"

    Set dTests = dRoot("tests")
    blnFirst = True
    For Each varKey In dTests.Keys
        strKey = CStr(varKey)
        Set dTest = dTests(strKey)
        If Not blnFirst Then buf.Add ","
        blnFirst = False
        buf.Add vbCrLf
        AppendQuoted buf, strKey
        buf.Add ":"
        AppendStateTestRecord buf, dTest
    Next varKey

    buf.Add "}"
    buf.Add "}"
    EmitTestStateJson = buf.GetStr

End Function


' ===================== Private append helpers =====================


' Field-name arguments below are string literals declared in this module, never data, so
' they are emitted without an escape pass. Values always go through EscapeJsonString.
'
' Strings and objects are passed ByRef throughout: VBA copies the whole string on a ByVal
' String parameter, and this chain is entered on the order of ten thousand times per
' payload. ByVal is kept only for Long and Boolean, where passing the value beats
' dereferencing a pointer. Each helper emits its field in a single clsConcat.Add call,
' since Add tests ten optional arguments per invocation regardless of how many are used.
Private Sub AppendQuoted(buf As clsConcat, strText As String)
    buf.Add """", EscapeJsonString(strText), """"
End Sub


Private Sub AppendStringField(buf As clsConcat, strKey As String, _
    strValue As String, ByVal blnFirst As Boolean)

    If blnFirst Then
        buf.Add """", strKey, """:""", EscapeJsonString(strValue), """"
    Else
        buf.Add ",""", strKey, """:""", EscapeJsonString(strValue), """"
    End If

End Sub


Private Sub AppendLongField(buf As clsConcat, strKey As String, ByVal lngValue As Long)
    buf.Add ",""", strKey, """:", CStr(lngValue)
End Sub


Private Sub AppendBoolField(buf As clsConcat, strKey As String, ByVal blnValue As Boolean)

    If blnValue Then
        buf.Add ",""", strKey, """:true"
    Else
        buf.Add ",""", strKey, """:false"
    End If

End Sub


Private Sub AppendOptionalStringField(buf As clsConcat, strKey As String, _
    strValue As String)

    If Len(strValue) = 0 Then Exit Sub
    buf.Add ",""", strKey, """:""", EscapeJsonString(strValue), """"

End Sub


Private Sub AppendStateSummary(buf As clsConcat, dSummary As Dictionary)

    buf.Add "{""subs"":", CStr(CLng(dSummary("subs")))
    AppendLongField buf, "assertions", CLng(dSummary("assertions"))
    AppendLongField buf, "passed", CLng(dSummary("passed"))
    AppendLongField buf, "failed", CLng(dSummary("failed"))
    AppendLongField buf, "errored", CLng(dSummary("errored"))
    AppendLongField buf, "empty", CLng(dSummary("empty"))
    buf.Add "}"

End Sub


Private Sub AppendStateTestRecord(buf As clsConcat, dTest As Dictionary)

    Dim colAssertions As Collection
    Dim colLogged As Collection

    buf.Add "{"
    AppendStringField buf, "moduleName", CStr(dTest("moduleName")), True
    AppendStringField buf, "procName", CStr(dTest("procName")), False
    AppendStringField buf, "folder", CStr(dTest("folder")), False
    AppendLongField buf, "line", CLng(Nz(dTest("line"), 0))
    If dTest.Exists("sourceType") Then
        AppendStringField buf, "sourceType", CStr(dTest("sourceType")), False
    End If
    AppendStringField buf, "status", CStr(dTest("status")), False
    AppendLongField buf, "durationMs", CLng(Nz(dTest("durationMs"), 0))
    If dTest.Exists("lastRunAt") Then
        AppendStringField buf, "lastRunAt", CStr(dTest("lastRunAt")), False
    End If
    AppendBoolField buf, "stale", CBool(Nz(dTest("stale"), False))
    If dTest.Exists("errorMessage") Then
        AppendOptionalStringField buf, "errorMessage", CStr(Nz(dTest("errorMessage"), vbNullString))
    End If

    AppendStringArrayField buf, dTest, "tags"

    buf.Add ",""assertions"":"
    If dTest.Exists("assertions") Then
        If TypeName(dTest("assertions")) = "Collection" Then
            Set colAssertions = dTest("assertions")
            AppendAssertionArray buf, colAssertions
        Else
            buf.Add "[]"
        End If
    Else
        buf.Add "[]"
    End If

    ' Written whenever the key is present, empty array included, to match what
    ' modTestState.SerializeTestRecord put through ConvertToJson.
    If dTest.Exists("loggedErrors") Then
        buf.Add ",""loggedErrors"":"
        If TypeName(dTest("loggedErrors")) = "Collection" Then
            Set colLogged = dTest("loggedErrors")
            AppendLoggedErrorArray buf, colLogged
        Else
            buf.Add "[]"
        End If
    End If

    buf.Add "}"

End Sub


' Emits `,"<key>":[...]`, an empty array when the key is absent or holds anything other
' than a Collection. The item is resolved with Set rather than handed to a Variant.
Private Sub AppendStringArrayField(buf As clsConcat, dSource As Dictionary, _
    strKey As String)

    Dim colItems As Collection
    Dim i As Long

    buf.Add ",""", strKey, """:["
    If dSource.Exists(strKey) Then
        If TypeName(dSource(strKey)) = "Collection" Then
            Set colItems = dSource(strKey)
            For i = 1 To colItems.Count
                If i > 1 Then buf.Add ","
                AppendQuoted buf, CStr(colItems(i))
            Next i
        End If
    End If
    buf.Add "]"

End Sub


Private Sub AppendAssertionArray(buf As clsConcat, colAssertions As Collection)

    Dim dA As Dictionary
    Dim i As Long
    Dim blnFirst As Boolean

    buf.Add "["
    blnFirst = True
    For i = 1 To colAssertions.Count
        Set dA = colAssertions(i)
        If Not blnFirst Then buf.Add ","
        blnFirst = False
        buf.Add "{""seq"":", CStr(CLng(dA("seq")))
        AppendBoolField buf, "passed", CBool(dA("passed"))
        If dA.Exists("context") Then
            AppendOptionalStringField buf, "context", CStr(Nz(dA("context"), vbNullString))
        End If
        buf.Add "}"
    Next i
    buf.Add "]"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EmitTestResultsJson
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Serialize a per-run TestResults JSON payload from the in-memory runner
'           : Tests dictionary (not the durable state shape).
'---------------------------------------------------------------------------------------
'
Public Function EmitTestResultsJson(dTests As Dictionary, _
    ByVal lngDurationMs As Long, ByVal blnCancelled As Boolean, ByVal blnAllPassed As Boolean, _
    ByVal lngPassedCount As Long, ByVal lngFailedCount As Long, _
    ByVal lngErroredCount As Long, ByVal lngEmptyCount As Long, _
    ByVal lngTotalAssertions As Long, ByVal lngPassedAssertions As Long, _
    ByVal lngFailedAssertions As Long, _
    Optional strJUnitPath As String = vbNullString, _
    Optional strStatePath As String = vbNullString, _
    Optional strLogPath As String = vbNullString, _
    Optional strResultsPath As String = vbNullString) As String

    Dim buf As clsConcat
    Dim varKey As Variant
    Dim strKey As String
    Dim dTest As Dictionary
    Dim blnFirst As Boolean

    Set buf = New clsConcat
    buf.Add "{"
    AppendStringField buf, "runAt", Format$(Now, "yyyy-mm-dd hh:nn:ss"), True
    AppendStringField buf, "databasePath", CurrentProject.FullName, False
    AppendStringField buf, "addinVersion", GetVCSVersion, False
    buf.Add ",""durationMs"":", CStr(lngDurationMs)
    AppendBoolField buf, "cancelled", blnCancelled
    AppendBoolField buf, "allPassed", blnAllPassed
    AppendOptionalStringField buf, "junitPath", strJUnitPath
    AppendOptionalStringField buf, "statePath", strStatePath
    AppendOptionalStringField buf, "logPath", strLogPath
    AppendOptionalStringField buf, "resultsPath", strResultsPath
    buf.Add ",""summary"":"
    AppendResultsSummary buf, lngPassedCount + lngFailedCount + lngErroredCount + lngEmptyCount, _
        lngTotalAssertions, lngPassedAssertions, lngFailedAssertions, lngErroredCount, lngEmptyCount
    buf.Add ",""tests"":{"

    blnFirst = True
    If Not dTests Is Nothing Then
        For Each varKey In dTests.Keys
            strKey = CStr(varKey)
            Set dTest = dTests(strKey)
            If CLng(dTest("status")) = etsPending Then GoTo NextResultsTest
            If Not blnFirst Then buf.Add ","
            blnFirst = False
            AppendQuoted buf, strKey
            buf.Add ":"
            AppendResultsTestRecord buf, dTest
NextResultsTest:
        Next varKey
    End If

    buf.Add "}"
    buf.Add "}"
    EmitTestResultsJson = buf.GetStr

End Function


Private Sub AppendResultsSummary(buf As clsConcat, ByVal lngSubs As Long, _
    ByVal lngAssertions As Long, ByVal lngPassed As Long, ByVal lngFailed As Long, _
    ByVal lngErrored As Long, ByVal lngEmpty As Long)

    buf.Add "{""subs"":", CStr(lngSubs)
    AppendLongField buf, "assertions", lngAssertions
    AppendLongField buf, "passed", lngPassed
    AppendLongField buf, "failed", lngFailed
    AppendLongField buf, "errored", lngErrored
    AppendLongField buf, "empty", lngEmpty
    buf.Add "}"

End Sub


Private Sub AppendResultsTestRecord(buf As clsConcat, dTest As Dictionary)

    Dim colAssertions As Collection
    Dim colLoggedErrors As Collection

    buf.Add "{"
    AppendStringField buf, "status", StatusToResultsString(CLng(dTest("status"))), True
    AppendLongField buf, "durationMs", CLng(Nz(dTest("durationMs"), 0))
    ' Unlike the state record, this one keeps an empty errorMessage: the runner UI and the
    ' TestResults artifact have always carried the key whenever the runner recorded it.
    If dTest.Exists("errorMessage") Then
        buf.Add ",""errorMessage"":""", _
            EscapeJsonString(CStr(Nz(dTest("errorMessage"), vbNullString))), """"
    End If

    If dTest.Exists("loggedErrors") Then
        buf.Add ",""loggedErrors"":"
        If TypeName(dTest("loggedErrors")) = "Collection" Then
            Set colLoggedErrors = dTest("loggedErrors")
            AppendLoggedErrorArray buf, colLoggedErrors
        Else
            buf.Add "[]"
        End If
    End If

    AppendStringArrayField buf, dTest, "tags"

    buf.Add ",""assertions"":"
    If dTest.Exists("assertionResults") Then
        If TypeName(dTest("assertionResults")) = "Collection" Then
            Set colAssertions = dTest("assertionResults")
            AppendAssertionArray buf, colAssertions
        Else
            buf.Add "[]"
        End If
    Else
        buf.Add "[]"
    End If

    buf.Add "}"

End Sub


Private Function StatusToResultsString(ByVal lngStatus As Long) As String

    Select Case lngStatus
        Case etsPassed:  StatusToResultsString = "PASSED"
        Case etsFailed:  StatusToResultsString = "FAILED"
        Case etsErrored: StatusToResultsString = "ERRORED"
        Case etsEmpty:   StatusToResultsString = "EMPTY"
        Case Else:       StatusToResultsString = "PENDING"
    End Select

End Function


Private Sub AppendLoggedErrorArray(buf As clsConcat, colLogged As Collection)

    Dim dErr As Dictionary
    Dim i As Long
    Dim blnFirst As Boolean

    buf.Add "["
    blnFirst = True
    For i = 1 To colLogged.Count
        Set dErr = colLogged(i)
        If Not blnFirst Then buf.Add ","
        blnFirst = False
        buf.Add "{"
        AppendStringField buf, "level", CStr(dErr("level")), True
        AppendStringField buf, "message", CStr(dErr("message")), False
        If dErr.Exists("source") Then
            AppendOptionalStringField buf, "source", CStr(Nz(dErr("source"), vbNullString))
        End If
        If dErr.Exists("errNumber") Then
            If CLng(Nz(dErr("errNumber"), 0)) <> 0 Then
                AppendLongField buf, "errNumber", CLng(dErr("errNumber"))
            End If
        End If
        If dErr.Exists("errDescription") Then
            AppendOptionalStringField buf, "errDescription", CStr(Nz(dErr("errDescription"), vbNullString))
        End If
        buf.Add "}"
    Next i
    buf.Add "]"

End Sub


' ===================== Web runner bridge payloads =====================


'---------------------------------------------------------------------------------------
' Procedure : EmitTestTreeJson
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Serialize the scanned test tree for the runner sidebar, grouped by module
'           : in first-seen order. Each module's test nodes accumulate into their own
'           : buffer during the single pass over the tests, so no intermediate
'           : Dictionary/Collection tree is built.
'---------------------------------------------------------------------------------------
'
Public Function EmitTestTreeJson(dTests As Dictionary) As String

    Dim buf As clsConcat
    Dim bufTests As clsConcat
    Dim dModules As Dictionary
    Dim dFolders As Dictionary
    Dim dTest As Dictionary
    Dim varKey As Variant
    Dim strKey As String
    Dim strModule As String
    Dim blnFirst As Boolean

    If dTests Is Nothing Then
        EmitTestTreeJson = "{}"
        Exit Function
    End If

    Set dModules = New Dictionary
    Set dFolders = New Dictionary
    For Each varKey In dTests.Keys
        strKey = CStr(varKey)
        Set dTest = dTests(strKey)
        strModule = CStr(dTest("moduleName"))
        If dModules.Exists(strModule) Then
            Set bufTests = dModules(strModule)
            bufTests.Add ","
        Else
            Set bufTests = New clsConcat
            Set dModules(strModule) = bufTests
            dFolders.Add strModule, CStr(dTest("folder"))
        End If
        AppendTreeNode bufTests, strKey, dTest, strModule
    Next varKey

    Set buf = New clsConcat
    buf.Add "{"
    blnFirst = True
    For Each varKey In dModules.Keys
        strModule = CStr(varKey)
        If Not blnFirst Then buf.Add ","
        blnFirst = False
        AppendQuoted buf, strModule
        buf.Add ":{"
        AppendStringField buf, "name", strModule, True
        AppendStringField buf, "folder", CStr(dFolders(strModule)), False
        buf.Add ",""tests"":["
        Set bufTests = dModules(strModule)
        buf.Add bufTests.GetStr
        buf.Add "]}"
    Next varKey
    buf.Add "}"
    EmitTestTreeJson = buf.GetStr

End Function


Private Sub AppendTreeNode(buf As clsConcat, strTestKey As String, _
    dTest As Dictionary, strModule As String)

    buf.Add "{"
    AppendStringField buf, "key", strTestKey, True
    AppendStringField buf, "name", CStr(dTest("procName")), False
    AppendStringField buf, "module", strModule, False
    AppendStringField buf, "procName", CStr(dTest("procName")), False
    If dTest.Exists("line") Then
        AppendLongField buf, "lineNumber", CLng(Nz(dTest("line"), 0))
    End If
    AppendStringArrayField buf, dTest, "tags"
    buf.Add "}"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EmitTestStartJson
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Serialize the TestUI.onTestStart payload. Unlike the results, which the
'           : bridge delivers in batches, this one is pushed per test, so a full suite
'           : paid the generic converter's cost several hundred times.
'           : Field set matches AppendTreeNode's first four, which the runner UI keys on.
'---------------------------------------------------------------------------------------
'
Public Function EmitTestStartJson(strTestKey As String, dTest As Dictionary) As String

    Dim buf As clsConcat
    Dim strProcName As String

    ' name and procName carry the same value, so it is escaped once.
    Set buf = New clsConcat
    strProcName = EscapeJsonString(CStr(dTest("procName")))
    With buf
        .Add "{""key"":""", EscapeJsonString(strTestKey), """"
        .Add ",""name"":""", strProcName, """"
        .Add ",""module"":""", EscapeJsonString(CStr(dTest("moduleName"))), """"
        .Add ",""procName"":""", strProcName, """"
        .Add "}"
    End With
    EmitTestStartJson = buf.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : EmitResultsBatchJson
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Serialize every completed test as one onResultsBatch payload for the
'           : hydrate replay. Returns an empty string (and lngCount = 0) when no test
'           : has run, which is the caller's signal to skip the push.
'---------------------------------------------------------------------------------------
'
Public Function EmitResultsBatchJson(dTests As Dictionary, ByRef lngCount As Long) As String

    Dim buf As clsConcat
    Dim dTest As Dictionary
    Dim varKey As Variant
    Dim strKey As String

    lngCount = 0
    If dTests Is Nothing Then Exit Function

    Set buf = New clsConcat
    buf.Add "{""results"":["
    For Each varKey In dTests.Keys
        strKey = CStr(varKey)
        Set dTest = dTests(strKey)
        If CLng(Nz(dTest("status"), etsPending)) <> etsPending Then
            If lngCount > 0 Then buf.Add ","
            AppendWebResultRecord buf, strKey, dTest
            lngCount = lngCount + 1
        End If
    Next varKey

    If lngCount = 0 Then Exit Function
    buf.Add "]}"
    EmitResultsBatchJson = buf.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : EmitWebResultJson
' Author    : Adam Waller
' Date      : 8/18/2026
' Purpose   : Serialize one test result in the shape TestUI.onTestComplete expects. Shares
'           : AppendWebResultRecord with the batch replay so the two cannot drift.
'---------------------------------------------------------------------------------------
'
Public Function EmitWebResultJson(strTestKey As String, dTest As Dictionary) As String

    Dim buf As clsConcat

    Set buf = New clsConcat
    AppendWebResultRecord buf, strTestKey, dTest
    EmitWebResultJson = buf.GetStr

End Function


Private Sub AppendWebResultRecord(buf As clsConcat, strTestKey As String, _
    dTest As Dictionary)

    Dim colAssertions As Collection

    buf.Add "{"
    AppendStringField buf, "testKey", strTestKey, True
    AppendStringField buf, "status", WebStatusString(CLng(Nz(dTest("status"), etsPending))), False
    AppendLongField buf, "durationMs", CLng(Nz(dTest("durationMs"), 0))
    If dTest.Exists("errorMessage") Then
        AppendOptionalStringField buf, "errorMessage", CStr(Nz(dTest("errorMessage"), vbNullString))
    End If

    buf.Add ",""assertions"":"
    If dTest.Exists("assertionResults") Then
        If TypeName(dTest("assertionResults")) = "Collection" Then
            Set colAssertions = dTest("assertionResults")
            AppendAssertionArray buf, colAssertions
        Else
            buf.Add "[]"
        End If
    Else
        buf.Add "[]"
    End If

    If dTest.Exists("lastRunAt") Then
        AppendOptionalStringField buf, "lastRunAt", CStr(Nz(dTest("lastRunAt"), vbNullString))
    End If
    ' Read through Exists: a plain Dictionary read would add the key back to the live
    ' runner record with an Empty value.
    If dTest.Exists("fromPriorRun") Then
        AppendBoolField buf, "prior", CBool(Nz(dTest("fromPriorRun"), False))
    Else
        AppendBoolField buf, "prior", False
    End If

    buf.Add "}"

End Sub


Private Function WebStatusString(ByVal lngStatus As Long) As String

    Select Case lngStatus
        Case etsPassed:  WebStatusString = "pass"
        Case etsFailed:  WebStatusString = "fail"
        Case etsErrored: WebStatusString = "error"
        Case etsEmpty:   WebStatusString = "skip"
        Case Else:       WebStatusString = "pending"
    End Select

End Function
