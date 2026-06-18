Attribute VB_Name = "modTestConditionalFormat"
'---------------------------------------------------------------------------------------
' Module    : modTestConditionalFormat
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Round-trip tests for clsConditionalFormat. The CF14 block is the
'           : authoritative copy and must rebuild byte-for-byte for every rule shape.
'           : The legacy ConditionalFormat block rebuilds byte-for-byte for single-rule
'           : shapes (the common case); multi-rule legacy is asserted semantically only,
'           : because its per-rule layout is not fully documented (see docs).
'           :
'           : Fixtures are the exact hex blocks emitted by Access SaveAsText for the
'           : controls in the conditional-formatting sample form.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")


' --- Text9: single expression rule (bold off), expr "[fraOption]=1" ---
Private Const TEXT9_CF14 As String = _
    "01000100000001000000000000000100000000000000ffffff000d0000005b00" & _
    "6600720061004f007000740069006f006e005d003d0031000000000000000000" & _
    "00000000000000000000000000"
Private Const TEXT9_LEGACY As String = _
    "010000007e000000010000000100000000000000000000000e00000001000000" & _
    "00000000ffffff00000000000000000000000000000000000000000000000000" & _
    "0000000000000000000000000000000000000000000000000000000000000000" & _
    "5b006600720061004f007000740069006f006e005d003d00310000000000"

' --- Text11: three rules (expression bold, expression, field-has-focus) ---
Private Const TEXT11_CF14 As String = _
    "01000300000001000000000000000101000000000000ffffff000d0000005b00" & _
    "6600720061004f007000740069006f006e005d003d0031000000000000000000" & _
    "0000000000000000000000000001000000000000000100000000000000ffffff" & _
    "000d0000005b006600720061004f007000740069006f006e005d003d00320000" & _
    "0000000000000000000000000000000000000000020000000000000001000000" & _
    "00000000ffffff00000000000000000000000000000000000000000000000000" & _
    "00"

' --- Text25: single field-value "between" rule, bounds "test" and "test" ---
Private Const TEXT25_CF14 As String = _
    "01000100000000000000000000000101000000000000ffffff00060000002200" & _
    "5400450053005400220006000000220054004500530054002200000000000000" & _
    "0000000000000000000000"
Private Const TEXT25_LEGACY As String = _
    "010000007c000000010000000000000000000000000000000700000001010000" & _
    "00000000ffffff00000000000000000000000000000000000000000000000000" & _
    "0000000000000000000000000000000000000000000000000000000000000000" & _
    "22005400450053005400220000002200540045005300540022000000"


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactExpression
' Purpose   : The authoritative CF14 block rebuilds byte-for-byte (single expression).
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactExpression()
    TestAssert RebuildCF14(TEXT9_CF14) = TEXT9_CF14, "Text9 CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactMultiRule
' Purpose   : CF14 rebuilds byte-for-byte for a multi-rule block (expression + focus).
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactMultiRule()
    TestAssert RebuildCF14(TEXT11_CF14) = TEXT11_CF14, "Text11 CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactBetween
' Purpose   : CF14 rebuilds byte-for-byte for a field-value "between" rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactBetween()
    TestAssert RebuildCF14(TEXT25_CF14) = TEXT25_CF14, "Text25 CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyByteExactExpression
' Purpose   : The legacy block rebuilds byte-for-byte for a single expression rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyByteExactExpression()
    TestAssert RebuildLegacy(TEXT9_CF14) = TEXT9_LEGACY, "Text9 legacy byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyByteExactBetween
' Purpose   : The legacy block rebuilds byte-for-byte for a single between rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyByteExactBetween()
    TestAssert RebuildLegacy(TEXT25_CF14) = TEXT25_LEGACY, "Text25 legacy byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeExpressionRule
' Purpose   : The decoded model captures the expected fields for an expression rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeExpressionRule()

    Dim cCF As clsConditionalFormat
    Dim dRule As Dictionary

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex TEXT9_CF14
    TestAssert RuleCount(cCF) = 1, "Text9 has one rule"
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("Type") = "Expression", "rule type is Expression"
    TestAssert dRule("Enabled") = True, "rule is enabled"
    TestAssert dRule("FontBold") = False, "rule is not bold"
    TestAssert dRule("Expression1") = "[fraOption]=1", "expression text decoded"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeBetweenRule
' Purpose   : The decoded model captures both bounds of a between rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeBetweenRule()

    Dim cCF As clsConditionalFormat
    Dim dRule As Dictionary

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex TEXT25_CF14
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("Type") = "FieldValue", "rule type is FieldValue"
    TestAssert dRule("Operator") = "Between", "operator is Between"
    TestAssert dRule("Expression1") = """test""", "first bound decoded"
    TestAssert dRule("Expression2") = """test""", "second bound decoded"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeMultiRule
' Purpose   : A three-rule block decodes to the expected rule types in order.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeMultiRule()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex TEXT11_CF14
    TestAssert RuleCount(cCF) = 3, "Text11 has three rules"
    TestAssert NthRule(cCF, 1)("Type") = "Expression", "rule 1 is Expression"
    TestAssert NthRule(cCF, 1)("FontBold") = True, "rule 1 is bold"
    TestAssert NthRule(cCF, 2)("Type") = "Expression", "rule 2 is Expression"
    TestAssert NthRule(cCF, 3)("Type") = "FieldHasFocus", "rule 3 is FieldHasFocus"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSemanticRoundTrip
' Purpose   : Decoding a rebuilt CF14 block yields the same model (stable round-trip),
'           : including the multi-rule case where bytes need not match exactly.
'---------------------------------------------------------------------------------------
'
Public Sub TestSemanticRoundTrip()
    TestAssert ModelSignature(TEXT11_CF14) = ModelSignature(RebuildCF14(TEXT11_CF14)), _
        "Text11 model is stable across rebuild"
    TestAssert ModelSignature(TEXT25_CF14) = ModelSignature(RebuildCF14(TEXT25_CF14)), _
        "Text25 model is stable across rebuild"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeStripsStaleInlineBlock
' Purpose   : When a source file has BOTH an inline binary block and a JSON entry for the
'           : same control, the JSON wins: the stale inline block is removed and a single
'           : rebuilt block is injected. A control with an inline block but NO JSON entry
'           : keeps its block untouched.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeStripsStaleInlineBlock()

    Dim cParser As clsSourceParser
    Dim strForm As String
    Dim strJson As String
    Dim strOut As String

    ' Text9 has a JSON entry (and a stale inline block); Text99 has only an inline block.
    strForm = BuildControl("Text9", "aaaa0000") & vbCrLf & BuildControl("Text99", "bbbb0000")
    strJson = BuildCFJson("Text9", TEXT9_CF14)

    Set cParser = New clsSourceParser
    cParser.LoadString strForm, edbForm
    cParser.MergeConditionalFormat strJson
    strOut = cParser.GetOutput

    TestAssert InStr(strOut, "aaaa0000") = 0, "Text9 stale inline block removed"
    TestAssert InStr(strOut, "bbbb0000") > 0, "Text99 inline block (no JSON) preserved"
    TestAssert CountOccurrences(strOut, "ConditionalFormat14 = Begin") = 2, _
        "no duplicate CF14 block (Text9 rebuilt + Text99 kept)"
    TestAssert CountOccurrences(strOut, "ConditionalFormat = Begin") = 1, _
        "single rebuilt legacy block for Text9"
    TestAssert InStr(strOut, "0d0000005b00") > 0, "Text9 rebuilt block content present"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeIsIdempotent
' Purpose   : Running the merge again on already-merged output yields identical content
'           : (the stale-strip removes the prior injected block before re-injecting), so
'           : repeated imports cannot accumulate duplicate blocks.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeIsIdempotent()

    Dim cParser As clsSourceParser
    Dim strForm As String
    Dim strJson As String
    Dim strOnce As String
    Dim strTwice As String

    strForm = BuildControl("Text9", "aaaa0000")
    strJson = BuildCFJson("Text9", TEXT9_CF14)

    Set cParser = New clsSourceParser
    cParser.LoadString strForm, edbForm
    cParser.MergeConditionalFormat strJson
    strOnce = cParser.GetOutput

    Set cParser = New clsSourceParser
    cParser.LoadString strOnce, edbForm
    cParser.MergeConditionalFormat strJson
    strTwice = cParser.GetOutput

    TestAssert strOnce = strTwice, "merge is idempotent (no block accumulation)"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeLeavesInlineWhenNoJson
' Purpose   : A control with an inline block but no JSON entry is left exactly as-is
'           : (so option-off / un-migrated source round-trips unchanged).
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeLeavesInlineWhenNoJson()

    Dim cParser As clsSourceParser
    Dim strForm As String
    Dim strJson As String

    ' JSON covers a different control, so Text99's inline block must be untouched.
    strForm = BuildControl("Text99", "bbbb0000")
    strJson = BuildCFJson("Text9", TEXT9_CF14)

    Set cParser = New clsSourceParser
    cParser.LoadString strForm, edbForm
    cParser.MergeConditionalFormat strJson
    TestAssert cParser.GetOutput = strForm, "inline block preserved when no JSON entry"

End Sub


' ===================================================================================
' Helpers (parameterized, so not discovered as tests)
' ===================================================================================

Private Function RebuildCF14(strHex As String) As String
    Dim cCF As clsConditionalFormat
    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex strHex
    RebuildCF14 = cCF.BuildCF14Hex
End Function

Private Function RebuildLegacy(strCF14Hex As String) As String
    Dim cCF As clsConditionalFormat
    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex strCF14Hex
    RebuildLegacy = cCF.BuildLegacyHex
End Function

Private Function RuleCount(cCF As clsConditionalFormat) As Long
    Dim dModel As Dictionary
    Set dModel = cCF.GetDictionary
    RuleCount = dModel("Rules").Count
End Function

Private Function NthRule(cCF As clsConditionalFormat, lngIndex As Long) As Dictionary
    Dim dModel As Dictionary
    Set dModel = cCF.GetDictionary
    Set NthRule = dModel("Rules")(lngIndex)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ModelSignature
' Purpose   : Build a deterministic string signature of the decoded model so two models
'           : can be compared for semantic equality regardless of byte layout.
'---------------------------------------------------------------------------------------
'
Private Function ModelSignature(strCF14Hex As String) As String

    Dim cCF As clsConditionalFormat
    Dim dModel As Dictionary
    Dim varRule As Variant
    Dim dRule As Dictionary
    Dim varKey As Variant
    Dim cData As clsConcat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex strCF14Hex
    Set dModel = cCF.GetDictionary
    Set cData = New clsConcat
    For Each varRule In dModel("Rules")
        Set dRule = varRule
        For Each varKey In dRule.Keys
            cData.Add CStr(varKey), "=", CStr(dRule(varKey)), ";"
        Next varKey
        cData.Add "|"
    Next varRule
    ModelSignature = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildControl
' Purpose   : Build a minimal control block (optionally with an inline CF14 block whose
'           : hex content is a recognizable marker so tests can detect strip/keep).
'---------------------------------------------------------------------------------------
'
Private Function BuildControl(strName As String, strInlineMarker As String) As String

    Dim cData As clsConcat

    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf
    cData.Add "    Begin TextBox"
    cData.Add "        Name =""" & strName & """"
    If Len(strInlineMarker) > 0 Then
        cData.Add "        ConditionalFormat14 = Begin"
        cData.Add "            0x" & strInlineMarker
        cData.Add "        End"
    End If
    cData.Add "    End"
    ' Drop the trailing line break so blocks join predictably
    cData.Remove Len(vbCrLf)
    BuildControl = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildCFJson
' Purpose   : Build companion-JSON content carrying one control's decoded rule model,
'           : matching the structure produced on export (Items.ConditionalFormatting).
'---------------------------------------------------------------------------------------
'
Private Function BuildCFJson(strName As String, strCF14Hex As String) As String

    Dim cCF As clsConditionalFormat
    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dControls As Dictionary

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex strCF14Hex
    Set dControls = New Dictionary
    dControls.Add strName, cCF.GetDictionary
    Set dItems = New Dictionary
    dItems.Add "ConditionalFormatting", dControls
    Set dFile = New Dictionary
    dFile.Add "Items", dItems
    BuildCFJson = ConvertToJson(dFile)

End Function


'---------------------------------------------------------------------------------------
' Procedure : CountOccurrences
' Purpose   : Count non-overlapping occurrences of a substring.
'---------------------------------------------------------------------------------------
'
Private Function CountOccurrences(strText As String, strFind As String) As Long

    Dim lngPos As Long

    lngPos = InStr(1, strText, strFind)
    Do While lngPos > 0
        CountOccurrences = CountOccurrences + 1
        lngPos = InStr(lngPos + Len(strFind), strText, strFind)
    Loop

End Function
