Attribute VB_Name = "modTestPromptDefaults"
'---------------------------------------------------------------------------------------
' Module    : modTestPromptDefaults
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : An unattended caller is never prompted: MsgBox2 logs the prompt and
'           : returns intDefaultResult in place of an answer. A question that passes no
'           : default gets vbOK, which is not a button of a Yes/No or Retry/Cancel
'           : question, and the call site then decides by whatever comparison it makes.
'           : So every question has to say what it answers when nobody is there. This
'           : scans the code of the project under test for MsgBox2 calls whose buttons
'           : ask a question, and fails on any that passes no intDefaultResult.
'           : It reads the buttons as written at the call, so a question whose buttons
'           : arrive in a variable is not seen.
'           : Run: ?VCS.RunTests("modTestPromptDefaults")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Private Const ModuleName As String = "modTestPromptDefaults"

' Positions of the MsgBox2 parameters, counting from 1.
Private Const ArgButtons As Long = 4
Private Const ArgDefaultResult As Long = 6


Public Sub TestEveryQuestionPassesAnUnattendedAnswer()

    Dim cmp As VBIDE.VBComponent
    Dim colMissing As Collection
    Dim lngQuestions As Long
    Dim strMissing As String
    Dim varSite As Variant
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    Set colMissing = New Collection
    For Each cmp In CurrentVBProject.VBComponents
        ScanComponent cmp, colMissing, lngQuestions
    Next cmp

    ' A parser that finds nothing would pass the real check below.
    TestAssert lngQuestions >= 10, "scan found the questions it checks (" & lngQuestions & ")"

    For Each varSite In colMissing
        strMissing = strMissing & IIf(Len(strMissing), ", ", vbNullString) & varSite
    Next varSite
    TestAssert colMissing.Count = 0, "questions without an unattended answer: " & strMissing

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "scanning the project " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestPromptScanReadsBothCallForms()

    Dim colMissing As Collection
    Dim lngQuestions As Long

    ' The scanner itself, on lines written for it.
    Set colMissing = New Collection
    ScanLogicalLine "If MsgBox2(T(""a, b""), ""c"", , vbYesNo + vbQuestion) = vbYes Then", "call", colMissing, lngQuestions
    ScanLogicalLine "MsgBox2 ""a"", ""b"", ""c"", vbOKCancel, , vbCancel", "statement", colMissing, lngQuestions
    ScanLogicalLine "x = MsgBox2(""a"", intButtons:=vbRetryCancel, intDefaultResult:=vbCancel)", "named", colMissing, lngQuestions
    ScanLogicalLine "MsgBox2 ""a"", ""vbYesNo in a string"", , vbInformation", "notice", colMissing, lngQuestions
    ScanLogicalLine "' MsgBox2 ""a"", , , vbYesNo", "comment", colMissing, lngQuestions
    ScanLogicalLine "x = modUIUtil.MsgBox2(""a"", , , vbYesNo, , vbNo)", "qualified", colMissing, lngQuestions
    ScanLogicalLine "x = cOther.MsgBox2(""a"", , , vbYesNo)", "other", colMissing, lngQuestions

    TestAssert lngQuestions = 4, "four questions found (" & lngQuestions & ")"
    TestAssert colMissing.Count = 1, "one question without a default (" & colMissing.Count & ")"
    If colMissing.Count = 1 Then TestAssert colMissing(1) = "call", "the call form without a default"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanComponent
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : Join continued lines into logical lines and scan each one.
'---------------------------------------------------------------------------------------
'
Private Sub ScanComponent(cmp As VBIDE.VBComponent, colMissing As Collection, _
    lngQuestions As Long)

    Dim lngLine As Long
    Dim lngStart As Long
    Dim strLine As String
    Dim strLogical As String

    With cmp.CodeModule
        For lngLine = 1 To .CountOfLines
            strLine = .Lines(lngLine, 1)
            If Len(strLogical) = 0 Then lngStart = lngLine
            If Right$(RTrim$(strLine), 2) = " _" Then
                strLogical = strLogical & Left$(RTrim$(strLine), Len(RTrim$(strLine)) - 1)
            Else
                strLogical = strLogical & strLine
                ScanLogicalLine strLogical, cmp.Name & ":" & lngStart, colMissing, lngQuestions
                strLogical = vbNullString
            End If
        Next lngLine
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanLogicalLine
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : Find the MsgBox2 calls on one logical line, and add strSite to colMissing
'           : for each question among them that passes no intDefaultResult.
'---------------------------------------------------------------------------------------
'
Private Sub ScanLogicalLine(ByVal strLine As String, ByVal strSite As String, colMissing As Collection, _
    lngQuestions As Long)

    Const Target As String = "MsgBox2"

    Dim strCode As String
    Dim lngPos As Long
    Dim colArgs As Collection
    Dim strButtons As String
    Dim blnHasDefault As Boolean
    Dim lngArg As Long

    strCode = CodeWithoutLiterals(strLine)
    lngPos = InStr(1, strCode, Target, vbBinaryCompare)
    Do While lngPos > 0
        If IsCallAt(strCode, lngPos, Len(Target)) Then
            Set colArgs = ArgumentsAfter(strCode, lngPos + Len(Target))
            strButtons = vbNullString
            blnHasDefault = False
            For lngArg = 1 To colArgs.Count
                Select Case ArgumentName(colArgs(lngArg))
                    Case "intButtons"
                        strButtons = colArgs(lngArg)
                    Case "intDefaultResult"
                        blnHasDefault = True
                    Case vbNullString
                        If lngArg = ArgButtons Then strButtons = colArgs(lngArg)
                        If lngArg = ArgDefaultResult Then blnHasDefault = (Len(Trim$(colArgs(lngArg))) > 0)
                End Select
            Next lngArg
            If IsQuestion(strButtons) Then
                lngQuestions = lngQuestions + 1
                If Not blnHasDefault Then colMissing.Add strSite
            End If
        End If
        lngPos = InStr(lngPos + Len(Target), strCode, Target, vbBinaryCompare)
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CodeWithoutLiterals
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : The line with every string literal emptied and any comment removed, so
'           : that neither can be read as code.
'---------------------------------------------------------------------------------------
'
Private Function CodeWithoutLiterals(ByVal strLine As String) As String

    Dim lngPos As Long
    Dim strChar As String
    Dim blnInString As Boolean

    With New clsConcat
        For lngPos = 1 To Len(strLine)
            strChar = Mid$(strLine, lngPos, 1)
            If blnInString Then
                ' A doubled quote stays inside the literal.
                If strChar = """" Then
                    If Mid$(strLine, lngPos + 1, 1) = """" Then
                        lngPos = lngPos + 1
                    Else
                        blnInString = False
                        .Add strChar
                    End If
                End If
            ElseIf strChar = """" Then
                blnInString = True
                .Add strChar
            ElseIf strChar = "'" Then
                Exit For
            Else
                .Add strChar
            End If
        Next lngPos
        CodeWithoutLiterals = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsCallAt
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : True when the identifier at lngPos is a whole word that is called, not
'           : declared or part of a longer name.
'---------------------------------------------------------------------------------------
'
Private Function IsCallAt(strCode As String, lngPos As Long, lngLen As Long) As Boolean

    Dim strBefore As String
    Dim strAfter As String

    If lngPos > 1 Then strBefore = Mid$(strCode, lngPos - 1, 1)
    strAfter = Mid$(strCode, lngPos + lngLen, 1)
    ' A call qualified by its module still counts.
    If strBefore = "." Then
        If Right$(Left$(strCode, lngPos - 1), Len("modUIUtil.")) <> "modUIUtil." Then Exit Function
    ElseIf strBefore Like "[A-Za-z0-9_]" Then
        Exit Function
    End If
    If strAfter Like "[A-Za-z0-9_]" Then Exit Function
    If InStr(1, strCode, "Function MsgBox2", vbTextCompare) > 0 Then Exit Function
    IsCallAt = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ArgumentsAfter
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : The top-level arguments of the call that starts at lngPos: inside the
'           : parentheses of a call form, or up to the end of a statement form.
'---------------------------------------------------------------------------------------
'
Private Function ArgumentsAfter(strCode As String, lngPos As Long) As Collection

    Dim colArgs As Collection
    Dim lngCursor As Long
    Dim lngDepth As Long
    Dim strChar As String
    Dim strArg As String
    Dim blnCallForm As Boolean

    Set colArgs = New Collection
    lngCursor = lngPos
    Do While Mid$(strCode, lngCursor, 1) = " "
        lngCursor = lngCursor + 1
    Loop
    blnCallForm = (Mid$(strCode, lngCursor, 1) = "(")
    If blnCallForm Then lngCursor = lngCursor + 1

    Do While lngCursor <= Len(strCode)
        strChar = Mid$(strCode, lngCursor, 1)
        Select Case strChar
            Case "("
                lngDepth = lngDepth + 1
                strArg = strArg & strChar
            Case ")"
                If lngDepth = 0 Then Exit Do
                lngDepth = lngDepth - 1
                strArg = strArg & strChar
            Case ","
                If lngDepth = 0 Then
                    colArgs.Add strArg
                    strArg = vbNullString
                Else
                    strArg = strArg & strChar
                End If
            Case ":"
                ' A statement separator ends a statement-form call; ":=" does not.
                If lngDepth = 0 And Not blnCallForm And Mid$(strCode, lngCursor + 1, 1) <> "=" Then Exit Do
                strArg = strArg & strChar
            Case Else
                strArg = strArg & strChar
        End Select
        lngCursor = lngCursor + 1
    Loop
    colArgs.Add strArg

    Set ArgumentsAfter = colArgs

End Function


'---------------------------------------------------------------------------------------
' Procedure : ArgumentName
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : The parameter name of a named argument (name:=value), or an empty string
'           : for a positional one. A ":=" nested deeper, as in T("{0}", var0:=x), does
'           : not make the argument a named one.
'---------------------------------------------------------------------------------------
'
Private Function ArgumentName(ByVal strArg As String) As String

    Dim strTrimmed As String
    Dim lngPos As Long

    strTrimmed = Trim$(strArg)
    For lngPos = 1 To Len(strTrimmed)
        If Not Mid$(strTrimmed, lngPos, 1) Like "[A-Za-z0-9_]" Then Exit For
    Next lngPos
    If lngPos > 1 And Mid$(strTrimmed, lngPos, 2) = ":=" Then
        ArgumentName = Left$(strTrimmed, lngPos - 1)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsQuestion
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/20/2026
' Purpose   : True when the buttons argument offers more than one answer.
'---------------------------------------------------------------------------------------
'
Private Function IsQuestion(ByVal strButtons As String) As Boolean
    IsQuestion = (strButtons Like "*vbYesNo*") _
        Or (strButtons Like "*vbOKCancel*") _
        Or (strButtons Like "*vbRetryCancel*") _
        Or (strButtons Like "*vbAbortRetryIgnore*")
End Function
