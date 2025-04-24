VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCaseChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsCaseChange
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Track and fix changes to variable casing in VBA.
'           : NOTE: This is not intended to be 100% comprehensive, addressing every
'           : possible edge case. It is designed to do the 99% in a highly performant
'           : way that doesn't negatively impact the overall performance of the tool.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Type udtThis
    ParsedSymbols As Dictionary
    SavedSymbols As Dictionary
    Changes As Dictionary
    TestMode As Boolean
    WatchName As String
    ModuleName As String
    CurrentLine As Long
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Scan
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Scan the current project to load symbols (variables, function names, etc.)
'---------------------------------------------------------------------------------------
'
Public Function Scan() As Boolean

    Dim cmp As VBComponent
    Dim strVBA As String

    Set this.ParsedSymbols = New Dictionary

    If this.TestMode Then Perf.StartTiming
    Perf.OperationStart "Scan VBA Symbols"

    If this.TestMode Then Perf.OperationStart "Scan Components"
    For Each cmp In CurrentVBProject.VBComponents
        If Not cmp.CodeModule Is Nothing Then
            If this.TestMode Then Perf.OperationStart "Read VBA"
            this.ModuleName = cmp.Name
            strVBA = cmp.CodeModule.Lines(1, 999999)
            If this.TestMode Then Perf.OperationEnd
            ParseVBA strVBA
        End If
    Next cmp
    If this.TestMode Then Perf.OperationEnd

    Perf.OperationEnd this.ParsedSymbols.Count

    If this.TestMode Then
        Perf.EndTiming
        Debug.Print Perf.GetReports
    End If

    WriteFile BuildJsonFile(TypeName(Me), SortDictionaryByKeys(this.ParsedSymbols), "Symbols (variable names) used in this project"), _
        Options.GetExportFolder & PathSep & "symbols.json"

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseVBA
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Parse the symbols from the VBA code
'---------------------------------------------------------------------------------------
'
Public Function ParseVBA(strCode As String)

    Dim varLines As Variant
    Dim lngLine As Long
    Dim strTLine As String
    Dim varWords As Variant
    Dim intWord As Integer
    Dim lngPos As Long
    Dim blnInBlock As Boolean
    Dim strTest As String

    ' Split code into lines
    varLines = Split(strCode, vbCrLf)

    ' Loop through lines
    Do While lngLine <= UBound(varLines)

        ' Note current line for debugging
        this.CurrentLine = lngLine + 1

        ' Get trimmmed line (right and left, in case of line continuation)
        If Mid$(varLines(lngLine), 1, 1) = " " Then
            strTLine = Trim$(varLines(lngLine))
        Else
            strTLine = varLines(lngLine)
        End If
        If strTLine = vbNullString Then GoTo NextLine

        ' Remove any line numbers
        If IsNumeric(Left(strTLine, 1)) Then
            ' Find first space
            lngPos = InStr(1, strTLine, " ")
            If lngPos > 1 Then
                strTLine = Mid$(strTLine, lngPos, Len(strTLine))
            Else
                ' No code we are interested in here. (Could be just a line number)
                GoTo NextLine
            End If
        End If

        ' Could be in an enum or type block
        If blnInBlock Then
            If StartsWith(strTLine, "End ") Then
                ' Exit block
                blnInBlock = False
            Else
                If Left(strTLine, 1) <> "'" Then
                    ' First word should be variable name
                    AddWord Split(strTLine, " ")(0)
                End If
            End If
            GoTo NextLine
        End If

        ' Check for line label
        If Right(strTLine, 1) = ":" Then
            strTLine = StripText(strTLine)
            If Right(strTLine, 1) = ":" Then
                If InStr(1, strTLine, " ") = 0 Then
                    AddWord strTLine
                    GoTo NextLine
                End If
            End If
        End If


        ' Check first letter before parsing into words
        Select Case Left$(strTLine, 1)
            Case "C", "D", "E", "F", "G", "I", "P", "R", "S", "T"

                ' Join any line continuations
                Do While Right(strTLine, 2) = " _"
                    ' Make sure this isn't in a commented line
                    strTLine = StripText(strTLine)
                    If Right(strTLine, 2) = " _" Then
                        ' Append next line as line continuation
                        lngLine = lngLine + 1
                        strTest = Trim(varLines(lngLine))
                        If Len(strTest) > 0 Then
                            strTLine = strTLine & " " & strTest
                        End If
                    End If
                Loop

                ' Split into words
                varWords = Split(strTLine, " ")
                If UBound(varWords) < 1 Then GoTo NextLine

                ' Check for enum/type blocks
                If varWords(0) = "Type" Or varWords(1) = "Type" _
                    Or varWords(0) = "Enum" Or varWords(1) = "Enum" Then
                    blnInBlock = True
                    GoTo NextLine
                End If

                ' Check for simple variables
                Select Case varWords(0)
                    Case "Const", "Dim", "Friend", "Global", "Implements", "Private", "Public", "ReDim", "Static"

                        ' Next word should be variable name.
                        ' Loop in case we have multiple variables declared on the same line.
                        For intWord = 1 To UBound(varWords)
                            AddWord varWords(intWord)

                            ' Check for comma separator if multiple variables were declared
                            If Right(varWords(1), 1) <> "," Then Exit For
                        Next intWord

                End Select

                ' Parse function and parameter names
                'Event|Function|Property|Sub



        End Select


NextLine:
        ' Increment line number
        lngLine = lngLine + 1
    Loop

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddWord
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Add word to dictionaries, also parsing out parentheses, commas, etc.
'---------------------------------------------------------------------------------------
'
Private Function AddWord(strRawWord As Variant) As String

    Dim lngEnd As Long
    Dim strWord As String

    ' Determine end of word (before parenthese, comma, colon, period, etc...)
    lngEnd = FirstPos(strRawWord, "(", ")", ",", ";", ":")

    ' Get end of Word, if no stop words found
    If lngEnd = 0 Then lngEnd = Len(strRawWord) + 1

    ' Get parsed word
    strWord = Mid$(strRawWord, 1, lngEnd - 1)

    ' Break if watched name is reached
    If strWord = this.WatchName Then
        Debug.Print "Found " & strWord & " on line " & this.CurrentLine & " of " & this.ModuleName
        Perf.PauseTiming
        ' Use ctl+L to view the call stack for debugging
        Stop
        Perf.ResumeTiming
    End If

    If this.ParsedSymbols.Exists(strWord) Then
        ' Increment usage count
        this.ParsedSymbols(strWord) = this.ParsedSymbols(strWord) + 1
    Else
        this.ParsedSymbols.Add strWord, 1
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : StripText
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Strip quoted text and comments from code line. (Replaces with space)
'           : (Used to improve parsing around comments and quoted text)
'---------------------------------------------------------------------------------------

Private Function StripText(strTrimmedLine As String) As String

    Dim lngQuoteStart As Long
    Dim lngQuoteEnd As Long
    Dim lngCommentStart As Long
    Dim lngPosition As Long

    ' Nothing to parse on a blank string
    If strTrimmedLine = vbNullString Then Exit Function

    ' Track performance when in testing mode
    If this.TestMode Then Perf.OperationStart "Strip Text & Comments"

    ' Run some initial quick checks on code line
    If Left(strTrimmedLine, 1) = "'" Then
        ' Just a comment
        StripText = vbNullString
    Else
        lngCommentStart = InStr(1, strTrimmedLine, "'")
        lngQuoteStart = InStr(1, strTrimmedLine, """")
        If lngQuoteStart = 0 Then
            ' No quoted text
            If lngCommentStart = 0 Then
                ' No comments either. Just plain code
                StripText = strTrimmedLine
            Else
                ' Trim off trailing comment
                StripText = RTrim(Left(strTrimmedLine, lngCommentStart - 1))
            End If
        Else
            ' See if the double quote is part of a comment
            If lngQuoteStart > lngCommentStart And lngCommentStart > 0 Then
                ' Trim off trailing comment
                StripText = RTrim(Left(strTrimmedLine, lngCommentStart - 1))
            Else
                ' We need to run the heavier parsing on this text
                With New clsConcat
                    .AppendOnAdd = " "

                    ' Add initial string before first quote
                    .Add Mid(strTrimmedLine, 1, lngQuoteStart - 1)
                    lngQuoteEnd = InStr(lngQuoteStart + 1, strTrimmedLine, """")
                    lngPosition = lngQuoteEnd + 1

                    ' Now, loop through quote pairs to remove quoted
                    ' sections and watch for comments outside quotes.
                    Do
                        ' Get next quote and comment positions
                        lngQuoteStart = InStr(lngPosition, strTrimmedLine, """")
                        lngCommentStart = InStr(lngPosition, strTrimmedLine, """")

                        ' Parse next string section
                        If lngQuoteStart = 0 Then
                            If lngCommentStart = 0 Then
                                ' No quotes or comments. Return remaining line.
                                .Add Mid(strTrimmedLine, lngPosition)
                                Exit Do
                            Else
                                ' Add up to comment
                                .Add Mid(strTrimmedLine, lngPosition, lngCommentStart - lngPosition - 1)
                                Exit Do
                            End If
                        Else
                            ' Find end of quoted section
                            lngQuoteEnd = InStr(lngQuoteStart + 1, strTrimmedLine, """")
                            If lngQuoteEnd = 0 Then
                                ' Unfinished quote. Ignore this text
                                Exit Do
                            Else
                                ' Move position to end of current quote
                                lngPosition = lngQuoteEnd + 1
                            End If
                        End If
                     Loop

                     ' Return result
                     .RTrim
                     StripText = .GetStr
                End With
            End If
        End If
    End If

    If this.TestMode Then Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : SelfTest
' Author    : Adam Waller
' Date      : 4/24/2025
' Purpose   : Test internal functions for expected output
'---------------------------------------------------------------------------------------
'
Public Sub SelfTest()

    ' StripText function
    Debug.Assert StripText("' Just a commment") = vbNullString
    Debug.Assert StripText("' Just a ""commment""") = vbNullString
    Debug.Assert StripText("Dim x As Long ' with a ""commment""") = "Dim x As Long"
    Debug.Assert StripText("Sub Creative(""Isn't this terrible!"")") = "Sub Creative( )"
    Debug.Assert StripText("Sub Creative(""Dim x, y As Long ' Because we can!"")") = "Sub Creative( )"

    this.TestMode = True
    'this.WatchName = "AfterSaveCallbackProcedureName"
    Me.Scan

    Debug.Print this.ParsedSymbols.Count & " Symbols Found"

End Sub
