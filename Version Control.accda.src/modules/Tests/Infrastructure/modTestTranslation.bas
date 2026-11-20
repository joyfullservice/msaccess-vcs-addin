Attribute VB_Name = "modTestTranslation"
'---------------------------------------------------------------------------------------
' Module    : modTestTranslation
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Tests for clsTranslation: harvesting T() strings from source, choosing a
'           : language from the Office UI language, and the translation files in the
'           : repository's Translation folder.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")


'---------------------------------------------------------------------------------------
' Procedure : TestExtractPlainLiteral
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : A literal first argument is harvested; the remaining arguments are not.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractPlainLiteral()
    TestAssert Extracted("Log.Add T(""Exported {0} files"", var0:=lngCount)") = _
        "Exported {0} files", "literal with placeholder argument"
    TestAssert Extracted("strText = T(""Done"")") = "Done", "literal as the only argument"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractEscapedQuote
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Doubled quotes inside a literal become one quote.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractEscapedQuote()
    TestAssert Extracted("x = T(""Click """"Yes"""" to continue"")") = _
        "Click ""Yes"" to continue", "doubled quotes unescaped"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractContinuedConcatenation
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Literals joined with & across continued lines form one string, and line
'           : break constants become LF.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractContinuedConcatenation()
    Dim strCode As String
    strCode = "MsgBox2 T(""Line one."" & vbCrLf & _" & vbCrLf & _
        "    ""Line two. "" & _" & vbCrLf & _
        "    ""Still two.""), vbInformation"
    TestAssert Extracted(strCode) = "Line one." & vbLf & "Line two. Still two.", _
        "continued concatenation"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractIIfBranches
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Both branches of IIf(condition, literal, literal) are harvested.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractIIfBranches()
    TestAssert Extracted("s = T(IIf(blnFullBuild, ""Build"", ""Merge""))") = "Build|Merge", _
        "IIf with two literals"
    TestAssert Extracted("s = T(IIf(Len(x) > 0, ""Some, but"", ""None"" & vbLf & ""yet""), var0:=1)") = _
        "Some, but|None" & vbLf & "yet", "IIf with a call in the condition and a comma in a literal"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractNestedCall
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : A T() call inside the arguments of another is found as well.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractNestedCall()
    TestAssert Extracted("Log.Add T(""{0} found"", var0:=T(""Forms""))") = "{0} found|Forms", _
        "nested T() call"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractSkipsNonLiteral
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : An argument computed at runtime is reported as skipped, not harvested.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractSkipsNonLiteral()

    Dim colSkipped As Collection
    Dim colStrings As Collection

    Set colSkipped = New Collection
    With New clsTranslation
        Set colStrings = .ExtractTranslatableStrings( _
            "x = T(LCase(strName))" & vbCrLf & _
            "y = T(""Prefix "" & strName)", colSkipped)
    End With

    TestAssert colStrings.Count = 0, "nothing harvested from computed arguments"
    TestAssert colSkipped.Count = 2, "both computed arguments reported, got " & colSkipped.Count

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExtractIgnoresNonCalls
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Text that only looks like a T() call is not harvested.
'---------------------------------------------------------------------------------------
'
Public Sub TestExtractIgnoresNonCalls()
    TestAssert Extracted("' Log.Add T(""In a comment"")") = vbNullString, "comment"
    TestAssert Extracted("x = 1 ' T(""Trailing comment"")") = vbNullString, "trailing comment"
    TestAssert Extracted("x = GetT(""Other function"")") = vbNullString, "longer name"
    TestAssert Extracted("x = obj.T(""Member"")") = vbNullString, "member of another object"
    TestAssert Extracted("s = ""T(""""Inside a literal"""")""") = vbNullString, "inside a literal"
    TestAssert Extracted("Public Function T(strText As String) As String") = vbNullString, _
        "declaration of T()"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveOsLanguageExactMatch
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : A translation for the exact Office language is used.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveOsLanguageExactMatch()
    With New clsTranslation
        TestAssert .ResolveOsLanguage("de_DE", Languages) = "de_DE", "German"
        TestAssert .ResolveOsLanguage("pt_BR", Languages) = "pt_BR", "Brazilian Portuguese"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveOsLanguageRegionalVariant
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : A regional variant uses another translation of the same language.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveOsLanguageRegionalVariant()
    With New clsTranslation
        TestAssert .ResolveOsLanguage("de_AT", Languages) = "de_DE", "Austrian German"
        TestAssert .ResolveOsLanguage("de_CH", Languages) = "de_DE", "Swiss German"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestResolveOsLanguageEnglishAndUnknown
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : English variants never resolve to en_TEST, and a language with no
'           : translation falls back to US English.
'---------------------------------------------------------------------------------------
'
Public Sub TestResolveOsLanguageEnglishAndUnknown()
    With New clsTranslation
        TestAssert .ResolveOsLanguage("en_GB", Languages) = "en_US", "British English"
        TestAssert .ResolveOsLanguage("en_US", Languages) = "en_US", "US English"
        TestAssert .ResolveOsLanguage("fr_FR", Languages) = "en_US", "no French translation"
        TestAssert .ResolveOsLanguage(vbNullString, Languages) = "en_US", "unknown Office language"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAvailableLanguagesReadFromTable
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : The languages loaded by the build can be read back. This is the input to
'           : ResolveOsLanguage whenever the user has not chosen a language.
'---------------------------------------------------------------------------------------
'
Public Sub TestAvailableLanguagesReadFromTable()

    Dim dLanguages As Dictionary
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler
    With New clsTranslation
        Set dLanguages = .GetAvailableLanguages
    End With
    TestAssert dLanguages.Exists("en_US"), "en_US loaded from tblLanguages"
    TestAssert dLanguages.Exists("de_DE"), "de_DE loaded from tblLanguages"

CleanUp:
    On Error Resume Next
    If lngErr <> 0 Then TestAssert False, "GetAvailableLanguages " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTranslationEntriesValid
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Every translated string in Translation\*.po keeps the {n} placeholders of
'           : its English original. A dropped placeholder loses the value silently; an
'           : extra one shows up as literal "{1}" in the interface. A translated form
'           : tooltip must also fit the 255 characters Access allows, or it is clipped.
'---------------------------------------------------------------------------------------
'
Public Sub TestTranslationEntriesValid()

    Dim strFolder As String
    Dim oFile As Scripting.File
    Dim lngFiles As Long

    strFolder = CodeProject.Path & PathSep & "Translation"
    If Not FSO.FolderExists(strFolder) Then
        TestAssert True, "skipped: no Translation folder at " & strFolder
        Exit Sub
    End If

    For Each oFile In FSO.GetFolder(strFolder).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "po", vbTextCompare) = 0 Then
            CheckPoFile oFile.Path, oFile.Name
            lngFiles = lngFiles + 1
        End If
    Next oFile

    TestAssert lngFiles > 0, "no .po files found in " & strFolder

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckPoFile
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Check each translated entry in one .po file.
'---------------------------------------------------------------------------------------
'
Private Sub CheckPoFile(strPath As String, strFileName As String)

    Dim varLines As Variant
    Dim lngLine As Long
    Dim strLine As String
    Dim strId As String
    Dim strStr As String
    Dim blnInStr As Boolean
    Dim blnTip As Boolean
    Dim blnNextTip As Boolean

    varLines = Split(Replace(ReadFile(strPath), vbCrLf, vbLf), vbLf)
    For lngLine = 0 To UBound(varLines)
        strLine = Trim$(varLines(lngLine))
        If Left$(strLine, 3) = "#: " Then
            If InStr(1, strLine, ".ControlTipText", vbBinaryCompare) > 0 Then blnNextTip = True
        ElseIf Left$(strLine, 6) = "msgid " Then
            CheckEntry strFileName, strId, strStr, blnTip
            strId = UnquotePo(Mid$(strLine, 7))
            strStr = vbNullString
            blnInStr = False
            blnTip = blnNextTip
            blnNextTip = False
        ElseIf Left$(strLine, 7) = "msgstr " Then
            strStr = UnquotePo(Mid$(strLine, 8))
            blnInStr = True
        ElseIf Left$(strLine, 1) = """" Then
            If blnInStr Then strStr = strStr & UnquotePo(strLine) Else strId = strId & UnquotePo(strLine)
        End If
    Next lngLine
    CheckEntry strFileName, strId, strStr, blnTip

End Sub


Private Sub CheckEntry(strFileName As String, strId As String, strStr As String, blnTip As Boolean)
    If Len(strId) = 0 Or Len(strStr) = 0 Then Exit Sub
    If PlaceholderSet(strId) <> PlaceholderSet(strStr) Then
        TestAssert False, strFileName & ": placeholders differ for """ & Left$(strId, 60) & """"
    End If
    If blnTip Then
        If Len(UnescapePo(strStr)) > 255 Then
            TestAssert False, strFileName & ": tooltip over 255 characters for """ & Left$(strId, 60) & """"
        End If
    End If
End Sub


Private Function UnescapePo(strText As String) As String
    UnescapePo = Replace(Replace(Replace(strText, "\n", vbLf), "\""", """"), "\\", "\")
End Function


Private Function PlaceholderSet(strText As String) As String
    Dim lngIndex As Long
    For lngIndex = 0 To 9
        If InStr(1, strText, "{" & lngIndex & "}") > 0 Then PlaceholderSet = PlaceholderSet & lngIndex
    Next lngIndex
End Function


Private Function UnquotePo(strQuoted As String) As String
    UnquotePo = strQuoted
    If Left$(UnquotePo, 1) = """" Then UnquotePo = Mid$(UnquotePo, 2)
    If Right$(UnquotePo, 1) = """" Then UnquotePo = Left$(UnquotePo, Len(UnquotePo) - 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Extracted
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Harvested strings from a block of source, joined with "|".
'---------------------------------------------------------------------------------------
'
Private Function Extracted(strCode As String) As String

    Dim colStrings As Collection
    Dim varString As Variant

    With New clsTranslation
        Set colStrings = .ExtractTranslatableStrings(strCode)
    End With
    For Each varString In colStrings
        If Len(Extracted) Then Extracted = Extracted & "|"
        Extracted = Extracted & varString
    Next varString

End Function


'---------------------------------------------------------------------------------------
' Procedure : Languages
' Author    : Adam Waller
' Date      : 11/20/2026
' Purpose   : Available languages as loaded from tblLanguages, en_TEST first so that a
'           : prefix match on English would find it.
'---------------------------------------------------------------------------------------
'
Private Function Languages() As Dictionary
    Set Languages = New Dictionary
    Languages.Add "en_TEST", "English (Localization Testing)"
    Languages.Add "en_US", "English"
    Languages.Add "de_DE", "Deutsch"
    Languages.Add "pt_BR", "Brazilian Portuguese"
End Function
