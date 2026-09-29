Attribute VB_Name = "modTestTestProcedures"
'---------------------------------------------------------------------------------------
' Module    : modTestTestProcedures
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Enforce the error-handling rules for test procedures. Err.Raise, or any
'           : failure after On Error GoTo 0, breaks into a modal VBA error dialog that
'           : stalls an unattended run. These tests scan every procedure the runner
'           : would discover and fail on either construct. Discovery mirrors the rules
'           : in clsTestRunner, which cannot be reused here: a fresh runner's Scan
'           : rewrites modTestAssert, and editing code resets the project mid-run.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private Const HEADER_LINES As Long = 30


'---------------------------------------------------------------------------------------
' Procedure : TestProceduresNeverRaiseOrDisableTrapping
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : No discovered test procedure uses Err.Raise or On Error GoTo 0.
'---------------------------------------------------------------------------------------
'
Public Sub TestProceduresNeverRaiseOrDisableTrapping()

    Dim dViolations As Dictionary
    Dim lngScanned As Long
    Dim varKey As Variant
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    Set dViolations = GetTestProcedureViolations(lngScanned)
    TestAssert lngScanned > 0, "scanned the project's test procedures"
    TestAssert dViolations.Count = 0, _
        dViolations.Count & " test procedure line(s) use Err.Raise or On Error GoTo 0"
    For Each varKey In dViolations.Keys
        TestAssert False, varKey & ": " & dViolations(varKey)
    Next varKey

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, _
        "unable to scan test procedures " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestForbiddenConstructDetection
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The line check finds both constructs as statements, and ignores them in
'           : comments and string literals, where tests legitimately mention them.
'---------------------------------------------------------------------------------------
'
Public Sub TestForbiddenConstructDetection()
    TestAssert FindForbiddenConstruct("    Err.Raise vbObjectError + 1") = "Err.Raise", _
        "Err.Raise statement"
    TestAssert FindForbiddenConstruct("    If lngErr <> 0 Then Err.Raise lngErr, , strErr") = "Err.Raise", _
        "inline Err.Raise"
    TestAssert FindForbiddenConstruct("    On Error GoTo 0") = "On Error GoTo 0", _
        "On Error GoTo 0 statement"
    TestAssert FindForbiddenConstruct(vbTab & "on  error   goto  0") = "On Error GoTo 0", _
        "case and spacing variations"
    TestAssert FindForbiddenConstruct("    On Error GoTo ErrHandler") = vbNullString, _
        "a handler target is allowed"
    TestAssert FindForbiddenConstruct("    ' Err.Raise opens a dialog") = vbNullString, _
        "comment is ignored"
    TestAssert FindForbiddenConstruct("    Rem On Error GoTo 0") = vbNullString, _
        "Rem comment is ignored"
    TestAssert FindForbiddenConstruct("    strCode = ""On Error GoTo 0""") = vbNullString, _
        "string literal is ignored"
    TestAssert FindForbiddenConstruct("    x = ""it's"": Err.Raise 5") = "Err.Raise", _
        "apostrophe inside a string does not start a comment"
    TestAssert FindForbiddenConstruct("    objMyErr.Raise") = vbNullString, _
        "Raise on another object is allowed"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDiscoveryMirrorsRunnerRules
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The scan checks exactly the procedures clsTestRunner would run.
'---------------------------------------------------------------------------------------
'
Public Sub TestDiscoveryMirrorsRunnerRules()

    Dim strName As String

    TestAssert IsDiscoverableTest("Public Sub TestA()", False, strName), "public sub"
    TestAssert strName = "TestA", "procedure name extracted"
    TestAssert IsDiscoverableTest("Sub TestB() ' note", False, strName), "implicitly public sub"
    TestAssert Not IsDiscoverableTest("Private Sub TestC()", False, strName), "private sub"
    TestAssert Not IsDiscoverableTest("Public Sub TestD(strArg As String)", False, strName), _
        "sub with a parameter"
    TestAssert Not IsDiscoverableTest("Public Function TestE() As Boolean", False, strName), _
        "function in a standard module"
    TestAssert IsDiscoverableTest("Public Function TestF() As Boolean", True, strName), _
        "function in a class module"
    TestAssert Not IsDiscoverableTest("Public Sub Class_Initialize()", True, strName), _
        "class lifecycle method"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTestProcedureViolations
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Return "Module.Procedure line N" -> construct for every forbidden line in
'           : a discovered test procedure, counting the procedures scanned.
'---------------------------------------------------------------------------------------
'
Private Function GetTestProcedureViolations(ByRef lngScanned As Long) As Dictionary

    Dim proj As VBIDE.VBProject
    Dim cmp As VBIDE.VBComponent
    Dim blnHasFolders As Boolean
    Dim dViolations As Dictionary

    Set dViolations = New Dictionary
    Set proj = GetCodeVBProject
    blnHasFolders = ProjectHasFolderAnnotations(proj)

    For Each cmp In proj.VBComponents
        If IsTestModule(cmp, blnHasFolders) Then ScanModule cmp, dViolations, lngScanned
    Next cmp

    Set GetTestProcedureViolations = dViolations

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanModule
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Check the body of each discovered test procedure in one module. Like the
'           : runner, a module without a TestAssert statement has no tests.
'---------------------------------------------------------------------------------------
'
Private Sub ScanModule(cmp As VBIDE.VBComponent, dViolations As Dictionary, _
    ByRef lngScanned As Long)

    Dim varLines As Variant
    Dim lngLine As Long
    Dim strTrimmed As String
    Dim strProc As String
    Dim strConstruct As String
    Dim blnIsClass As Boolean
    Dim blnInTest As Boolean

    varLines = ModuleLines(cmp)
    If Not ContainsTestAssert(varLines) Then Exit Sub
    blnIsClass = (cmp.Type = vbext_ct_ClassModule)

    For lngLine = 0 To UBound(varLines)
        strTrimmed = Trim$(varLines(lngLine))
        If blnInTest Then
            If IsProcedureEnd(strTrimmed) Then
                blnInTest = False
            Else
                strConstruct = FindForbiddenConstruct(strTrimmed)
                If Len(strConstruct) > 0 Then
                    dViolations.Add cmp.Name & "." & strProc & " line " & (lngLine + 1), strConstruct
                End If
            End If
        ElseIf IsDiscoverableTest(strTrimmed, blnIsClass, strProc) Then
            blnInTest = True
            lngScanned = lngScanned + 1
        End If
    Next lngLine

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindForbiddenConstruct
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Return the name of a forbidden construct on this line, or an empty string.
'---------------------------------------------------------------------------------------
'
Private Function FindForbiddenConstruct(strLine As String) As String

    Dim strCode As String

    If InStr(1, strLine, "Raise", vbTextCompare) = 0 And _
        InStr(1, strLine, "GoTo", vbTextCompare) = 0 Then Exit Function

    strCode = " " & UCase$(Replace(CodeWithoutLiterals(strLine), vbTab, " ")) & " "
    Do While InStr(strCode, "  ") > 0
        strCode = Replace(strCode, "  ", " ")
    Loop

    If strCode Like "*[!A-Z0-9_]ERR.RAISE[!A-Z0-9_]*" Then
        FindForbiddenConstruct = "Err.Raise"
    ElseIf strCode Like "*[!A-Z0-9_]ON ERROR GOTO 0[!A-Z0-9_]*" Then
        FindForbiddenConstruct = "On Error GoTo 0"
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CodeWithoutLiterals
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Blank out string literals and drop any trailing comment.
'---------------------------------------------------------------------------------------
'
Private Function CodeWithoutLiterals(strLine As String) As String

    Dim lngPos As Long
    Dim strChar As String
    Dim blnInString As Boolean
    Dim strUpper As String

    strUpper = UCase$(Trim$(strLine))
    If strUpper = "REM" Or strUpper Like "REM *" Then Exit Function

    For lngPos = 1 To Len(strLine)
        strChar = Mid$(strLine, lngPos, 1)
        If strChar = """" Then
            blnInString = Not blnInString
            strChar = " "
        ElseIf blnInString Then
            strChar = " "
        ElseIf strChar = "'" Then
            Exit For
        End If
        CodeWithoutLiterals = CodeWithoutLiterals & strChar
    Next lngPos

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDiscoverableTest
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Match clsTestRunner's test declaration rules: a parameterless public Sub,
'           : or in a class also a Function, excluding Class_ lifecycle methods.
'---------------------------------------------------------------------------------------
'
Private Function IsDiscoverableTest(strTrimmed As String, blnIsClass As Boolean, _
    ByRef strName As String) As Boolean

    Dim strUpper As String
    Dim lngStart As Long
    Dim lngOpen As Long
    Dim lngClose As Long

    strUpper = UCase$(strTrimmed)
    If strUpper Like "PUBLIC SUB *" Then
        lngStart = 12
    ElseIf strUpper Like "SUB *" Then
        lngStart = 5
    ElseIf blnIsClass And strUpper Like "PUBLIC FUNCTION *" Then
        lngStart = 17
    ElseIf blnIsClass And strUpper Like "FUNCTION *" Then
        lngStart = 10
    Else
        Exit Function
    End If

    lngOpen = InStr(lngStart, strTrimmed, "(")
    If lngOpen = 0 Then Exit Function
    lngClose = InStr(lngOpen + 1, strTrimmed, ")")
    If lngClose = 0 Then Exit Function
    If Len(Trim$(Mid$(strTrimmed, lngOpen + 1, lngClose - lngOpen - 1))) > 0 Then Exit Function

    strName = Trim$(Mid$(strTrimmed, lngStart, lngOpen - lngStart))
    If Len(strName) = 0 Then Exit Function
    If blnIsClass And StrComp(Left$(strName, 6), "Class_", vbTextCompare) = 0 Then Exit Function

    IsDiscoverableTest = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsProcedureEnd
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : True for the End Sub or End Function that closes a procedure.
'---------------------------------------------------------------------------------------
'
Private Function IsProcedureEnd(strTrimmed As String) As Boolean
    Dim strUpper As String
    strUpper = UCase$(strTrimmed)
    IsProcedureEnd = (strUpper Like "END SUB*" Or strUpper Like "END FUNCTION*")
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsTestModule
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Match clsTestRunner's module rules: a standard or class module whose
'           : @Folder has a Tests segment, falling back to "Test" in the name when the
'           : project or the module carries no @Folder annotation.
'---------------------------------------------------------------------------------------
'
Private Function IsTestModule(cmp As VBIDE.VBComponent, blnHasFolders As Boolean) As Boolean

    Dim strFolder As String
    Dim varPart As Variant

    If Not IsCodeModule(cmp) Then Exit Function

    If blnHasFolders Then strFolder = FolderAnnotation(HeaderText(cmp))
    If Len(strFolder) = 0 Then
        IsTestModule = (InStr(1, cmp.Name, "Test", vbTextCompare) > 0)
    Else
        For Each varPart In Split(strFolder, ".")
            If StrComp(varPart, "Tests", vbTextCompare) = 0 Then IsTestModule = True
        Next varPart
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectHasFolderAnnotations
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : True when any code module declares an @Folder annotation.
'---------------------------------------------------------------------------------------
'
Private Function ProjectHasFolderAnnotations(proj As VBIDE.VBProject) As Boolean

    Dim cmp As VBIDE.VBComponent

    For Each cmp In proj.VBComponents
        If IsCodeModule(cmp) Then
            If Len(FolderAnnotation(HeaderText(cmp))) > 0 Then
                ProjectHasFolderAnnotations = True
                Exit Function
            End If
        End If
    Next cmp

End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderAnnotation
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Return the quoted path from an @Folder annotation, or an empty string.
'---------------------------------------------------------------------------------------
'
Private Function FolderAnnotation(strCode As String) As String

    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    lngPos = InStr(1, strCode, "'@Folder(", vbTextCompare)
    If lngPos = 0 Then Exit Function
    lngStart = InStr(lngPos, strCode, """")
    If lngStart = 0 Then Exit Function
    lngEnd = InStr(lngStart + 1, strCode, """")
    If lngEnd > lngStart + 1 Then FolderAnnotation = Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ContainsTestAssert
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : True when any line begins with a TestAssert statement.
'---------------------------------------------------------------------------------------
'
Private Function ContainsTestAssert(varLines As Variant) As Boolean

    Dim lngLine As Long

    For lngLine = 0 To UBound(varLines)
        If Left$(UCase$(Trim$(varLines(lngLine))), 11) = "TESTASSERT " Then
            ContainsTestAssert = True
            Exit Function
        End If
    Next lngLine

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsCodeModule
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : True for standard and standalone class modules, the only discoverable types.
'---------------------------------------------------------------------------------------
'
Private Function IsCodeModule(cmp As VBIDE.VBComponent) As Boolean
    IsCodeModule = (cmp.Type = vbext_ct_StdModule Or cmp.Type = vbext_ct_ClassModule)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HeaderText
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The first lines of a module, where the runner reads annotations.
'---------------------------------------------------------------------------------------
'
Private Function HeaderText(cmp As VBIDE.VBComponent) As String

    Dim lngCount As Long

    lngCount = cmp.CodeModule.CountOfLines
    If lngCount > HEADER_LINES Then lngCount = HEADER_LINES
    If lngCount > 0 Then HeaderText = cmp.CodeModule.Lines(1, lngCount)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ModuleLines
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : All lines of a module, as a zero-based array.
'---------------------------------------------------------------------------------------
'
Private Function ModuleLines(cmp As VBIDE.VBComponent) As Variant
    If cmp.CodeModule.CountOfLines = 0 Then
        ModuleLines = Split(vbNullString, vbCrLf)
    Else
        ModuleLines = Split(cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines), vbCrLf)
    End If
End Function
