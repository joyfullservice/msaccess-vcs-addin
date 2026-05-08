VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestRunner
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Test runner engine for the TestAssert framework. Discovers test subs in
'           : CurrentVBProject, drives execution, records assertion results, and streams
'           : progress to the frmVCSMain log console via the Log singleton.
'           :
'           : This class is a singleton managed through modObjects.TestRunner.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests")

Private Const ModuleName As String = "clsTestRunner"

' Runner state machine
Public Enum eTestRunnerState
    etrsIdle = 0
    etrsScanning = 1
    etrsReady = 2
    etrsRunning = 3
    etrsComplete = 4
    etrsCancelled = 5
End Enum

' Test result status
Public Enum eTestStatus
    etsPending = 0
    etsPassed = 1
    etsFailed = 2
    etsErrored = 3
    etsEmpty = 4
End Enum

' Private UDT for internal state
Private Type udtTestRunner
    State As eTestRunnerState
    Tests As Dictionary             ' Key: "ModuleName.ProcName", Value: Dictionary (test record)
    SuiteMap As Dictionary          ' Key: folder path, Value: Collection of test keys
    CurrentTestKey As String
    AssertionCounter As Long
    AssertionResults As Collection  ' Collection of assertion result Dictionaries for current test
    CancelRequested As Boolean
    RunStartTime As Currency
    PassedCount As Long
    FailedCount As Long
    ErroredCount As Long
    EmptyCount As Long
End Type
Private this As udtTestRunner


'---------------------------------------------------------------------------------------
' Procedure : State
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Current state of the test runner.
'---------------------------------------------------------------------------------------
'
Public Property Get State() As eTestRunnerState
    State = this.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : Tests
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Dictionary of discovered test records (populated after scan).
'---------------------------------------------------------------------------------------
'
Public Property Get Tests() As Dictionary
    Set Tests = this.Tests
End Property


'---------------------------------------------------------------------------------------
' Procedure : SuiteMap
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Dictionary mapping folder/suite names to collections of test keys.
'---------------------------------------------------------------------------------------
'
Public Property Get SuiteMap() As Dictionary
    Set SuiteMap = this.SuiteMap
End Property


'---------------------------------------------------------------------------------------
' Procedure : PassedCount / FailedCount / ErroredCount / EmptyCount
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Summary counts from the most recent run.
'---------------------------------------------------------------------------------------
'
Public Property Get PassedCount() As Long
    PassedCount = this.PassedCount
End Property
Public Property Get FailedCount() As Long
    FailedCount = this.FailedCount
End Property
Public Property Get ErroredCount() As Long
    ErroredCount = this.ErroredCount
End Property
Public Property Get EmptyCount() As Long
    EmptyCount = this.EmptyCount
End Property


'---------------------------------------------------------------------------------------
' Procedure : Scan
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Discover test modules and test subs in CurrentVBProject.
'           : Populates the Tests dictionary and SuiteMap.
'---------------------------------------------------------------------------------------
'
Public Sub Scan()

    Dim cmp As VBIDE.VBComponent
    Dim proj As VBIDE.VBProject
    Dim blnHasFolderTags As Boolean
    Dim dModules As Dictionary

    this.State = etrsScanning
    Set this.Tests = New Dictionary
    Set this.SuiteMap = New Dictionary
    ResetCounts

    Set proj = CurrentVBProject

    ' First pass: check if any module in the project uses @Folder annotations
    blnHasFolderTags = ProjectHasFolderAnnotations(proj)

    ' Second pass: identify test modules
    Set dModules = New Dictionary
    For Each cmp In proj.VBComponents
        If IsTestModule(cmp, blnHasFolderTags) Then
            dModules.Add cmp.Name, cmp
        End If
    Next cmp

    ' Third pass: scan each test module for test subs
    Dim varKey As Variant
    For Each varKey In dModules.Keys
        ScanModuleForTests dModules(varKey), blnHasFolderTags
    Next varKey

    this.State = etrsReady

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunAll
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Run every discovered test.
'---------------------------------------------------------------------------------------
'
Public Sub RunAll()

    If this.Tests Is Nothing Then Scan
    If this.Tests.Count = 0 Then
        this.State = etrsComplete
        Exit Sub
    End If

    Dim colKeys As New Collection
    Dim varKey As Variant
    For Each varKey In this.Tests.Keys
        colKeys.Add CStr(varKey)
    Next varKey
    RunSelected colKeys

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunSelected
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Run a specific set of tests identified by their keys.
'---------------------------------------------------------------------------------------
'
Public Sub RunSelected(colTestKeys As Collection)

    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strCmd As String
    Dim lngErrNum As Long
    Dim strErrDesc As String
    Dim strTestKey As String
    Dim strProcName As String
    Dim curStart As Currency
    Dim intErrorTrapping As Integer
    Dim blnBreakOnError As Boolean

    this.State = etrsRunning
    this.CancelRequested = False
    ResetCounts
    Perf.OperationStart "Run Tests"
    this.RunStartTime = Perf.MicroTimer

    ' Disable break-on-error so test errors don't stop in the debugger
    intErrorTrapping = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2  ' Break on Unhandled Errors
    blnBreakOnError = Options.BreakOnError
    Options.BreakOnError = False

    ' Reset status on all selected tests
    For Each varKey In colTestKeys
        If this.Tests.Exists(CStr(varKey)) Then
            Set dTest = this.Tests(CStr(varKey))
            dTest("status") = etsPending
            Set dTest("assertionResults") = New Collection
        End If
    Next varKey

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    For Each varKey In colTestKeys

        If this.CancelRequested Then
            this.State = etrsCancelled
            Exit For
        End If

        strTestKey = CStr(varKey)
        If Not this.Tests.Exists(strTestKey) Then GoTo NextTest

        Set dTest = this.Tests(strTestKey)
        strProcName = CStr(dTest("procName"))
        this.CurrentTestKey = strTestKey
        this.AssertionCounter = 0
        Set this.AssertionResults = New Collection

        ' Verify procedure is callable
        If Not GlobalProcExists(strProcName) Then
            dTest("status") = etsErrored
            dTest("errorMessage") = "Not callable (Option Private Module or not found): " & strProcName
            this.ErroredCount = this.ErroredCount + 1
            LogTestResult strTestKey, "SKIP", "orange"
            Log.Add "  " & dTest("errorMessage"), False
            GoTo NextTest
        End If

        ' Run the test sub
        strCmd = BuildRunCmd(strProcName)
        curStart = Perf.MicroTimer
        On Error Resume Next
        Application.Run strCmd
        lngErrNum = Err.Number
        strErrDesc = Err.Description
        If lngErrNum <> 0 Then Err.Clear
        On Error GoTo 0

        ' Store assertion results and duration
        Set dTest("assertionResults") = this.AssertionResults
        dTest("durationMs") = CLng((Perf.MicroTimer - curStart) * 1000)

        ' Classify outcome and log result
        If lngErrNum <> 0 Then
            dTest("status") = etsErrored
            dTest("errorMessage") = "Error " & lngErrNum & ": " & strErrDesc
            this.ErroredCount = this.ErroredCount + 1
            LogTestResult strTestKey, "ERROR", "red"
            Log.Add "  " & dTest("errorMessage"), False
        ElseIf this.AssertionCounter = 0 Then
            dTest("status") = etsEmpty
            this.EmptyCount = this.EmptyCount + 1
            LogTestResult strTestKey, "EMPTY", "gray"
        ElseIf AnyAssertionFailed(this.AssertionResults) Then
            dTest("status") = etsFailed
            this.FailedCount = this.FailedCount + 1
            LogTestResult strTestKey, "FAIL", "red"
            LogFailedAssertions this.AssertionResults
        Else
            dTest("status") = etsPassed
            this.PassedCount = this.PassedCount + 1
            LogTestResult strTestKey, "PASS", "green"
        End If

        DoEvents

NextTest:
    Next varKey

    Perf.OperationEnd
    this.CurrentTestKey = vbNullString
    If this.State <> etrsCancelled Then this.State = etrsComplete

    LogRunSummary

    ' Restore original error trapping settings
    Options.BreakOnError = blnBreakOnError
    Application.SetOption "Error Trapping", intErrorTrapping

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunFailed
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Re-run only tests that failed or errored in the last run.
'---------------------------------------------------------------------------------------
'
Public Sub RunFailed()

    If this.Tests Is Nothing Then Exit Sub

    Dim colKeys As New Collection
    Dim varKey As Variant
    Dim dTest As Dictionary
    For Each varKey In this.Tests.Keys
        Set dTest = this.Tests(CStr(varKey))
        If dTest("status") = etsFailed Or dTest("status") = etsErrored Then
            colKeys.Add CStr(varKey)
        End If
    Next varKey

    If colKeys.Count > 0 Then RunSelected colKeys

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Cancel
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Request cancellation of the current run. The current test will finish
'           : before the run stops.
'---------------------------------------------------------------------------------------
'
Public Sub Cancel()
    this.CancelRequested = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RecordAssertion
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Called by HandleTestAssertion when a TestAssert fires during a test run.
'           : Records the result for the current test.
'---------------------------------------------------------------------------------------
'
Public Sub RecordAssertion(ByVal blnCondition As Boolean, ByVal varContext As Variant)

    Dim dResult As Dictionary

    this.AssertionCounter = this.AssertionCounter + 1

    Set dResult = New Dictionary
    dResult.Add "seq", this.AssertionCounter
    dResult.Add "passed", blnCondition
    dResult.Add "context", IIf(IsMissing(varContext) Or IsEmpty(varContext), vbNullString, CStr(varContext))

    this.AssertionResults.Add dResult

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetResultsAsJson
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Serialize the current test results to a JSON string for persistence.
'---------------------------------------------------------------------------------------
'
Public Function GetResultsAsJson() As String

    Dim dRoot As Dictionary
    Dim dSummary As Dictionary
    Dim dTests As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim dTestOut As Dictionary
    Dim colAssertions As Collection
    Dim colAssertOut As Collection
    Dim dA As Dictionary
    Dim dAOut As Dictionary
    Dim i As Long

    Set dRoot = New Dictionary
    Set dSummary = New Dictionary
    Set dTests = New Dictionary

    dRoot.Add "runAt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    dRoot.Add "databasePath", CurrentProject.FullName
    dRoot.Add "addinVersion", GetVCSVersion
    dRoot.Add "durationMs", CLng((Perf.MicroTimer - this.RunStartTime) * 1000)

    dSummary.Add "total", this.PassedCount + this.FailedCount + this.ErroredCount + this.EmptyCount
    dSummary.Add "passed", this.PassedCount
    dSummary.Add "failed", this.FailedCount
    dSummary.Add "errored", this.ErroredCount
    dSummary.Add "empty", this.EmptyCount
    Set dRoot("summary") = dSummary

    If Not this.Tests Is Nothing Then
        For Each varKey In this.Tests.Keys
            Set dTest = this.Tests(CStr(varKey))
            Set dTestOut = New Dictionary
            dTestOut.Add "status", StatusToString(CLng(dTest("status")))
            dTestOut.Add "durationMs", dTest("durationMs")
            If dTest.Exists("errorMessage") Then
                dTestOut.Add "errorMessage", dTest("errorMessage")
            End If

            Set colAssertOut = New Collection
            Set colAssertions = dTest("assertionResults")
            For i = 1 To colAssertions.Count
                Set dA = colAssertions(i)
                Set dAOut = New Dictionary
                dAOut.Add "seq", dA("seq")
                dAOut.Add "passed", dA("passed")
                If Len(dA("context")) > 0 Then dAOut.Add "context", dA("context")
                colAssertOut.Add dAOut
            Next i
            Set dTestOut("assertions") = colAssertOut
            Set dTests(CStr(varKey)) = dTestOut
        Next varKey
    End If

    Set dRoot("tests") = dTests
    GetResultsAsJson = ConvertToJson(dRoot, JSON_WHITESPACE)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveResults
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Write test results to a timestamped JSON file in the logs folder.
'---------------------------------------------------------------------------------------
'
Public Sub SaveResults()

    Dim strFolder As String
    Dim strFile As String

    strFolder = Options.GetExportFolder & "logs" & PathSep
    VerifyPath strFolder
    strFile = strFolder & "TestResults_" & _
        Format$(Now, "yyyymmdd\_hhnnss") & "_" & _
        Format$(Timer * 1000 Mod 1000, "000") & ".json"

    WriteFile GetResultsAsJson, strFile

End Sub


' ===================== Private helpers =====================


'---------------------------------------------------------------------------------------
' Procedure : ProjectHasFolderAnnotations
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Quick check whether any module in the project contains @Folder annotations.
'---------------------------------------------------------------------------------------
'
Private Function ProjectHasFolderAnnotations(proj As VBIDE.VBProject) As Boolean

    Dim cmp As VBIDE.VBComponent
    Dim strCode As String

    On Error Resume Next
    For Each cmp In proj.VBComponents
        If cmp.Type = vbext_ct_StdModule Then
            If cmp.CodeModule.CountOfLines > 0 Then
                strCode = UCase$(cmp.CodeModule.Lines(1, IIf(cmp.CodeModule.CountOfLines < 30, cmp.CodeModule.CountOfLines, 30)))
                If InStr(strCode, "'@FOLDER(") > 0 Then
                    ProjectHasFolderAnnotations = True
                    Exit Function
                End If
            End If
        End If
    Next cmp
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsTestModule
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Determine whether a VBComponent is a test module using discovery rules:
'           :   1. @Folder annotation containing "Tests" segment (if project uses folders)
'           :   2. Name containing "Test" (fallback if no @Folder tags in project)
'           :   3. Standard modules only (not forms/reports)
'---------------------------------------------------------------------------------------
'
Private Function IsTestModule(cmp As VBIDE.VBComponent, _
                              blnHasFolderTags As Boolean) As Boolean

    Dim strCode As String
    Dim strFolder As String

    ' Only scan standard modules
    If cmp.Type <> vbext_ct_StdModule Then Exit Function

    On Error Resume Next

    If blnHasFolderTags Then
        ' Use @Folder annotation: look for a "Tests" segment
        If cmp.CodeModule.CountOfLines > 0 Then
            strCode = cmp.CodeModule.Lines(1, IIf(cmp.CodeModule.CountOfLines < 30, cmp.CodeModule.CountOfLines, 30))
            strFolder = ExtractFolderAnnotation(strCode)
            If Len(strFolder) > 0 Then
                IsTestModule = FolderContainsTestsSegment(strFolder)
                Exit Function
            End If
        End If
    End If

    ' Fallback: name-based discovery (*Test* in the name)
    If InStr(1, cmp.Name, "Test", vbTextCompare) > 0 Then
        IsTestModule = True
    End If

    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractFolderAnnotation
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Extract the folder path from an @Folder annotation in module source.
'           : Returns empty string if not found.
'---------------------------------------------------------------------------------------
'
Private Function ExtractFolderAnnotation(strCode As String) As String

    Dim strUpper As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    Const TAG As String = "'@FOLDER("

    strUpper = UCase$(strCode)
    lngPos = InStr(1, strUpper, TAG)
    If lngPos = 0 Then Exit Function

    lngStart = InStr(lngPos, strCode, """")
    If lngStart > 0 Then
        lngEnd = InStr(lngStart + 1, strCode, """")
        If lngEnd > lngStart + 1 Then
            ExtractFolderAnnotation = Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderContainsTestsSegment
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if a folder path (dot-separated) contains "Tests" as a segment.
'---------------------------------------------------------------------------------------
'
Private Function FolderContainsTestsSegment(strFolder As String) As Boolean

    Dim varParts As Variant
    Dim i As Long
    varParts = Split(strFolder, ".")
    For i = LBound(varParts) To UBound(varParts)
        If StrComp(varParts(i), "Tests", vbTextCompare) = 0 Then
            FolderContainsTestsSegment = True
            Exit Function
        End If
    Next i

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanModuleForTests
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Scan a test module for parameterless Public Subs and register each one
'           : as a test. Populates the Tests dictionary.
'---------------------------------------------------------------------------------------
'
Private Sub ScanModuleForTests(cmp As VBIDE.VBComponent, blnHasFolderTags As Boolean)

    Dim strAllCode As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim strTrimmed As String
    Dim strProcName As String
    Dim strFolder As String
    Dim strTestKey As String
    Dim dTest As Dictionary

    On Error Resume Next
    If cmp.CodeModule.CountOfLines = 0 Then Exit Sub
    strAllCode = cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
    On Error GoTo 0
    If Len(strAllCode) = 0 Then Exit Sub

    ' Determine folder annotation for suite grouping
    If blnHasFolderTags Then
        strFolder = ExtractFolderAnnotation(strAllCode)
    End If
    If Len(strFolder) = 0 Then strFolder = cmp.Name

    varLines = Split(strAllCode, vbCrLf)

    For lngLine = LBound(varLines) To UBound(varLines)
        strTrimmed = Trim$(varLines(lngLine))

        If IsParameterlessSub(strTrimmed, strProcName) Then
            strTestKey = cmp.Name & "." & strProcName

            Set dTest = New Dictionary
            dTest.Add "moduleName", cmp.Name
            dTest.Add "procName", strProcName
            dTest.Add "folder", strFolder
            dTest.Add "status", CLng(etsPending)
            dTest.Add "durationMs", CLng(0)
            Set dTest("assertionResults") = New Collection

            this.Tests.Add strTestKey, dTest

            ' Add to suite map
            If Not this.SuiteMap.Exists(strFolder) Then
                this.SuiteMap.Add strFolder, New Collection
            End If
            this.SuiteMap(strFolder).Add strTestKey
        End If

    Next lngLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsParameterlessSub
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if a line declares a parameterless Public Sub. Returns the sub name
'           : via the ByRef parameter if matched.
'---------------------------------------------------------------------------------------
'
Private Function IsParameterlessSub(strTrimmed As String, ByRef strName As String) As Boolean

    Dim strUpper As String
    Dim lngSubPos As Long
    Dim lngParenOpen As Long
    Dim lngParenClose As Long
    Dim strBetween As String

    strUpper = UCase$(strTrimmed)

    ' Match "Public Sub X()" or "Sub X()" (Public is default)
    If Left$(strUpper, 11) = "PUBLIC SUB " Then
        lngSubPos = 12
    ElseIf Left$(strUpper, 4) = "SUB " Then
        ' Exclude "Private Sub"
        If Left$(strUpper, 12) = "PRIVATE SUB " Then Exit Function
        lngSubPos = 5
    Else
        Exit Function
    End If

    ' Extract name (everything up to the opening paren)
    lngParenOpen = InStr(lngSubPos, strTrimmed, "(")
    If lngParenOpen = 0 Then Exit Function

    strName = Trim$(Mid$(strTrimmed, lngSubPos, lngParenOpen - lngSubPos))
    If Len(strName) = 0 Then Exit Function

    ' Check for empty parameter list
    lngParenClose = InStr(lngParenOpen + 1, strTrimmed, ")")
    If lngParenClose = 0 Then Exit Function

    strBetween = Trim$(Mid$(strTrimmed, lngParenOpen + 1, lngParenClose - lngParenOpen - 1))
    IsParameterlessSub = (Len(strBetween) = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindTestAssertCall
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Find "TestAssert " at the start of a statement (case-insensitive).
'           : Returns the position of the character after "TestAssert " (the expression
'           : start), or 0 if not found.
'---------------------------------------------------------------------------------------
' Procedure : BuildRunCmd
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Build a project-qualified Application.Run command for a procedure in
'           : CurrentVBProject. Follows the RunSubInCurrentProject pattern from
'           : modDatabase.bas.
'---------------------------------------------------------------------------------------
'
Private Function BuildRunCmd(strProcName As String) As String

    If CurrentVBProject.Name = PROJECT_NAME Then
        With CurrentProject
            BuildRunCmd = .Path & PathSep & FSO.GetBaseName(.Name) & "." & strProcName
        End With
    Else
        BuildRunCmd = "[" & CurrentVBProject.Name & "]." & strProcName
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : AnyAssertionFailed
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if any assertion in a results collection failed.
'---------------------------------------------------------------------------------------
'
Private Function AnyAssertionFailed(colResults As Collection) As Boolean

    Dim dResult As Dictionary
    Dim i As Long
    For i = 1 To colResults.Count
        Set dResult = colResults(i)
        If Not CBool(dResult("passed")) Then
            AnyAssertionFailed = True
            Exit Function
        End If
    Next i

End Function


'---------------------------------------------------------------------------------------
' Procedure : StatusToString
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Convert a numeric test status to a display string.
'---------------------------------------------------------------------------------------
'
Private Function StatusToString(lngStatus As Long) As String
    Select Case lngStatus
        Case etsPassed:  StatusToString = "PASSED"
        Case etsFailed:  StatusToString = "FAILED"
        Case etsErrored: StatusToString = "ERRORED"
        Case etsEmpty:   StatusToString = "EMPTY"
        Case Else:       StatusToString = "PENDING"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : LogTestResult
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Log a single test result line with the status left-justified at a
'           : fixed column position so all result labels start at the same place.
'---------------------------------------------------------------------------------------
'
Private Sub LogTestResult(strTestKey As String, strStatus As String, strColor As String)
    Const STATUS_COL As Long = 50
    Dim lngPad As Long
    lngPad = STATUS_COL - Len(strTestKey)
    If lngPad < 2 Then lngPad = 2
    Log.Add strTestKey & Space$(lngPad), True, False
    Log.Add strStatus, True, True, strColor
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogRunSummary
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Print an end-of-run summary with each result category on its own line,
'           : non-zero counts bold in the appropriate color.
'---------------------------------------------------------------------------------------
'
Private Sub LogRunSummary()

    Dim lngTotal As Long
    Dim dblElapsed As Double

    lngTotal = this.PassedCount + this.FailedCount + this.ErroredCount + this.EmptyCount
    dblElapsed = Round((Perf.MicroTimer - this.RunStartTime), 3)

    With Log
        .Spacer
        .Add "Test summary", True, , "blue", True
        .Add "  Total:     " & lngTotal
        LogSummaryLine "  Passed:    ", CStr(this.PassedCount), _
            IIf(this.PassedCount > 0, "green", vbNullString)
        LogSummaryLine "  Failed:    ", CStr(this.FailedCount), _
            IIf(this.FailedCount > 0, "red", vbNullString)
        LogSummaryLine "  Errored:   ", CStr(this.ErroredCount), _
            IIf(this.ErroredCount > 0, "red", vbNullString)
        LogSummaryLine "  Empty:     ", CStr(this.EmptyCount), _
            IIf(this.EmptyCount > 0, "gray", vbNullString)
        .Add "  Elapsed:   " & dblElapsed & "s"
        .Spacer
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogSummaryLine
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Print a summary label + value, with the value bold and colored when
'           : a color is specified (non-zero count).
'---------------------------------------------------------------------------------------
'
Private Sub LogSummaryLine(ByVal strLabel As String, ByVal strValue As String, ByVal strColor As String)
    If Len(strColor) > 0 Then
        Log.Add strLabel, True, False
        Log.Add strValue, True, True, strColor, True
    Else
        Log.Add strLabel & strValue
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogFailedAssertions
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Log details for failed assertions to the console (file-only).
'---------------------------------------------------------------------------------------
'
Private Sub LogFailedAssertions(colResults As Collection)
    Dim dResult As Dictionary
    Dim i As Long
    For i = 1 To colResults.Count
        Set dResult = colResults(i)
        If Not CBool(dResult("passed")) Then
            If Len(dResult("context")) > 0 Then
                Log.Add "        #" & dResult("seq") & " FAILED: " & dResult("context"), False
            Else
                Log.Add "        #" & dResult("seq") & " FAILED", False
            End If
        End If
    Next i
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResetCounts
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Reset run summary counters.
'---------------------------------------------------------------------------------------
'
Private Sub ResetCounts()
    this.PassedCount = 0
    this.FailedCount = 0
    this.ErroredCount = 0
    this.EmptyCount = 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Initialize with empty collections.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set this.Tests = New Dictionary
    Set this.SuiteMap = New Dictionary
    Set this.AssertionResults = New Collection
End Sub
