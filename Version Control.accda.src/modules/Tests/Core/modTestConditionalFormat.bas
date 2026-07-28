Attribute VB_Name = "modTestConditionalFormat"
'---------------------------------------------------------------------------------------
' Module    : modTestConditionalFormat
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Behavior tests for clsConditionalFormat: decoding, colors, operators,
'           : trailer echoes, graceful failure, and source-file merge semantics.
'           :
'           : The byte-exact regression corpus lives in modTestConditionalFormatCorpus.
'           : The fixtures here were captured from controls formatted in the Access
'           : design-view dialog, which encodes a set Boolean as 1. The dialog and VBA
'           : disagree: Access stores a Boolean copied from VBA as VBA True (&HFF), and
'           : preserves whichever encoding it finds rather than normalizing. It also
'           : allocates one fewer null unit per legacy expression slot for dialog-authored
'           : rules. The emitter targets the VBA encoding (DECISIONS.md 7/28/2026), so
'           : these fixtures are deliberately NOT asserted byte-exact - they are kept to
'           : prove the other encoding still decodes and survives a rebuild unchanged in
'           : meaning. See docs/access-conditional-format.md section 6.
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

' --- AlertText: 4 expression rules, non-white BackColors (from rAlertList report) ---
Private Const ALERTTEXT_CF14 As String = _
    "010004000000010000000000000001010000ff804000dbdbb7001d0000004900" & _
    "6e0053007400720028005b0041006c0065007200740050006100720061006d00" & _
    "73005d002c0027005b0021005d003b00270029003e0030000000000000000000" & _
    "00dbdbb700000000000000000001000000000000000100010000000000dbdbb7" & _
    "001d00000049006e0053007400720028005b0041006c00650072007400500061" & _
    "00720061006d0073005d002c0027005b0069005d003b00270029003e00300000" & _
    "0000000000000000dbdbb7000000000000000000010000000000000001010000" & _
    "00000000dbdbb7001d00000049006e0053007400720028005b0041006c006500" & _
    "7200740050006100720061006d0073005d002c0027005b0062005d003b002700" & _
    "29003e003000000000000000000000dbdbb70000000000000000000100000000" & _
    "00000001010000ffffff00ba1419001f00000049006e0053007400720028005b" & _
    "0041006c0065007200740050006100720061006d0073005d002c0027005b0021" & _
    "00210021005d003b00270029003e003000000000000000000000000000000000" & _
    "000000000000"
Private Const ALERTTEXT_LEGACY As String = _
    "010000001a010000030000000100000000000000000000001e00000001010000" & _
    "ff804000dbdbb70001000000000000001f0000003d0000000100010000000000" & _
    "dbdbb70001000000000000003e0000005c0000000101000000000000dbdbb700" & _
    "49006e0053007400720028005b0041006c006500720074005000610072006100" & _
    "6d0073005d002c0027005b0021005d003b00270029003e003000000000004900" & _
    "6e0053007400720028005b0041006c0065007200740050006100720061006d00" & _
    "73005d002c0027005b0069005d003b00270029003e0030000000000049006e00" & _
    "53007400720028005b0041006c0065007200740050006100720061006d007300" & _
    "5d002c0027005b0062005d003b00270029003e00300000000000"

' --- Field-value operator fixtures (issue #725) ---
' Captured from Access SaveAsText for a single field-value rule on one text box, one
' capture per AcFormatConditionOperator. BackColor = red RGB(255,0,0), Expression1 = "1"
' throughout; Between/NotBetween add Expression2 = "2" (single-value operators leave it
' empty). The operator is the 2-byte value at CF14 offset 10 and the legacy dword at
' offset 16. See docs/access-conditional-format.md section 4.3.
Private Const OP_BETWEEN_CF14 As String = _
    "01000100000000000000000000000100000000000000ff000000010000003100" & _
    "0100000032000000000000ff0000000000000000000000"
Private Const OP_NOTBETWEEN_CF14 As String = _
    "01000100000000000000010000000100000000000000ff000000010000003100" & _
    "0100000032000000000000ff0000000000000000000000"
Private Const OP_EQUAL_CF14 As String = _
    "01000100000000000000020000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"
Private Const OP_NOTEQUAL_CF14 As String = _
    "01000100000000000000030000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"
Private Const OP_GREATERTHAN_CF14 As String = _
    "01000100000000000000040000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"
Private Const OP_LESSTHAN_CF14 As String = _
    "01000100000000000000050000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"
Private Const OP_GREATERTHANOREQUAL_CF14 As String = _
    "01000100000000000000060000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"
Private Const OP_LESSTHANOREQUAL_CF14 As String = _
    "01000100000000000000070000000100000000000000ff000000010000003100" & _
    "000000000000000000ff0000000000000000000000"

' Legacy blocks for the "equal" single-value capture (operator at offset 16, empty
' Expression2 slot) and the "between" capture (both slots filled).
Private Const OP_BETWEEN_LEGACY As String = _
    "0100000068000000010000000000000000000000000000000200000001000000" & _
    "00000000ff000000000000000000000000000000000000000000000000000000" & _
    "0000000000000000000000000000000000000000000000000000000000000000" & _
    "3100000032000000"
Private Const OP_EQUAL_LEGACY As String = _
    "0100000068000000010000000000000002000000000000000200000001000000" & _
    "00000000ff000000000000000000000000000000000000000000000000000000" & _
    "0000000000000000000000000000000000000000000000000000000000000000" & _
    "3100000000000000"

' --- Multi-rule field-value block with mixed operators (issue #725 real-world shape) ---
' Three field-value rules: Between "11".."22" (red), Equal "55" (green), GreaterThan "99"
' (blue). Each later rule's operator is the second dword of its 8-byte prefix.
Private Const MULTI_OP_CF14 As String = _
    "01000300000000000000000000000100000000000000ff000000020000003100" & _
    "310002000000320032000000000000ff00000000000000000000000000000002" & _
    "000000010000000000000000ff00000200000035003500000000000000000000" & _
    "00ff00000000000000000000000000000400000001000000000000000000ff00" & _
    "02000000390039000000000000000000000000ff000000000000000000"

' --- Data bar fixtures (issue #730) ---
' Single data bar, automatic shortest/longest limits (Lowest/Highest Value). Offset 14 is
' unk1 (=1), not the type dword — rule 0 has no per-rule prefix.
Private Const DATABAR_AUTO_CF14 As String = _
    "01000100000003000000000000000100000000000000ffffff000000000000000000" & _
    "01000000001c2400000000000000000000"

' Number limits with typed bounds "10" and "100".
Private Const DATABAR_NUMBER_CF14 As String = _
    "01000100000003000000000000000100000000000000ffffff000200000031003000" & _
    "0300000031003000300001000000000070c0000100000001000000"

' Percent limits with typed bounds "25" and "75".
Private Const DATABAR_PERCENT_CF14 As String = _
    "01000100000003000000000000000100000000000000ffffff000200000032003500" & _
    "020000003700350001000000001c2400000200000002000000"

' ShowBarOnly with number limits and a custom bar color. Dialog-authored, so ShowBarOnly
' is stored as 1 rather than VBA True - see the module header.
Private Const DATABAR_SHOWONLY_CF14 As String = _
    "01000100000003000000000000000100000000000000ffffff000100000030000200" & _
    "0000350030000100000001ff0000000100000001000000"

' --- Font flag fixture (issue #730) ---
' Single expression rule "[n0]>1" with FontBold and FontUnderline set from VBA, so the
' flags record is 01 ff 00 ff. The &HFF in the high byte is what overflowed ReadLong.
Private Const FLAGS_BOLD_UNDERLINE_CF14 As String = _
    "010001000000010000000000000001ff00ff00000000ffffff00060000005b00" & _
    "6e0030005d003e003100000000000000000000000000000000000000000000"

' Expression rule followed by a data bar (data bar carries the 8-byte type prefix).
Private Const EXPR_DATABAR_CF14 As String = _
    "01000200000001000000000000000100000000000000ffffff000d0000005b006600720061" & _
    "004f007000740069006f006e005d003d0031000000000000000000000000000000000000" & _
    "0000000003000000000000000100000000000000ffffff0000000000000000000100000000" & _
    "1c2400000000000000000000"


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactExpression
' Purpose   : The authoritative CF14 block rebuilds byte-for-byte (single expression).
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactExpression()
    TestAssert RebuildCF14(TEXT9_CF14) = TEXT9_CF14, "Text9 CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDialogBooleanEncodingDecodes
' Purpose   : A set Boolean stored by the design-view dialog as 1 decodes to True, just
'           : like the VBA encoding (&HFF). Any non-zero byte means True; only the
'           : emitter has to pick one value to write.
'---------------------------------------------------------------------------------------
'
Public Sub TestDialogBooleanEncodingDecodes()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex TEXT11_CF14
    TestAssert NthRule(cCF, 1)("FontBold") = True, "FontBold decoded from 0x01"

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex DATABAR_SHOWONLY_CF14
    TestAssert NthRule(cCF, 1)("ShowBarOnly") = True, "ShowBarOnly decoded from 0x01"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDialogEncodingSurvivesRebuild
' Purpose   : Dialog-authored blocks are re-emitted in the VBA encoding, so they are not
'           : byte-identical to the original. What must hold is that nothing is lost: the
'           : rebuilt block decodes to the same model, and rebuilding again is stable.
'           : (Byte-exactness is asserted against the captured corpus instead.)
'---------------------------------------------------------------------------------------
'
Public Sub TestDialogEncodingSurvivesRebuild()

    TestAssert ModelSignature(TEXT11_CF14) = ModelSignature(RebuildCF14(TEXT11_CF14)), _
        "Text11 model unchanged by rebuild"
    TestAssert ModelSignature(TEXT25_CF14) = ModelSignature(RebuildCF14(TEXT25_CF14)), _
        "Text25 model unchanged by rebuild"
    TestAssert ModelSignature(ALERTTEXT_CF14) = ModelSignature(RebuildCF14(ALERTTEXT_CF14)), _
        "AlertText model unchanged by rebuild"
    TestAssert ModelSignature(DATABAR_SHOWONLY_CF14) = _
        ModelSignature(RebuildCF14(DATABAR_SHOWONLY_CF14)), _
        "ShowBarOnly model unchanged by rebuild"

    ' A second pass must be byte-stable, so an import/export cycle cannot keep drifting.
    TestAssert RebuildCF14(RebuildCF14(ALERTTEXT_CF14)) = RebuildCF14(ALERTTEXT_CF14), _
        "AlertText rebuild is idempotent"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLegacySlotSizingUsesVbaConvention
' Purpose   : Pin the legacy expression-slot convention so a change to it is visible.
'           : Every rule owns two slots sized max(len + 1, 2) code units, which is what
'           : Access writes for VBA-authored rules. The design-view dialog allocates one
'           : null unit less per expression rule, so a rebuilt block is exactly 2 bytes
'           : per rule longer than the dialog original. Field-value rules already carry
'           : two slots in both encodings and so come out the same size.
'           :
'           : Both layouts are valid Access output; the emitter targets the VBA convention
'           : (DECISIONS.md 7/28/2026).
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacySlotSizingUsesVbaConvention()

    ' Expression rules: +1 null unit (2 bytes) per rule versus the dialog original
    TestAssert Len(RebuildLegacy(TEXT9_CF14)) \ 2 = (Len(TEXT9_LEGACY) \ 2) + 2, _
        "1 expression rule is 2 bytes longer than the dialog block"
    TestAssert Len(RebuildLegacy(ALERTTEXT_CF14)) \ 2 = (Len(ALERTTEXT_LEGACY) \ 2) + 6, _
        "3 legacy expression rules are 6 bytes longer than the dialog block"

    ' Field-value rules already have two slots in both encodings, so the size matches
    TestAssert Len(RebuildLegacy(TEXT25_CF14)) \ 2 = Len(TEXT25_LEGACY) \ 2, _
        "field-value block size matches the dialog block"

    ' The rebuilt block must stay internally consistent: Access silently discards the
    ' legacy blocks of every control on the form if any one of them is malformed.
    TestAssert LegacyBlockSizeField(RebuildLegacy(TEXT9_CF14)) = _
        Len(RebuildLegacy(TEXT9_CF14)) \ 2, "single-rule blockSize field is consistent"
    TestAssert LegacyBlockSizeField(RebuildLegacy(ALERTTEXT_CF14)) = _
        Len(RebuildLegacy(ALERTTEXT_CF14)) \ 2, "multi-rule blockSize field is consistent"
    TestAssert LegacyBlockSizeField(RebuildLegacy(TEXT25_CF14)) = _
        Len(RebuildLegacy(TEXT25_CF14)) \ 2, "field-value blockSize field is consistent"

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
    TestAssert dRule("ForeColor") = "RGB(0,0,0)", "ForeColor is black"
    TestAssert dRule("BackColor") = "RGB(255,255,255)", "BackColor is white"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestColorRoundTrip
' Purpose   : RGB color strings round-trip through decode/rebuild and accept legacy
'           : numeric Long values on import.
'---------------------------------------------------------------------------------------
'
Public Sub TestColorRoundTrip()

    Dim cCF As clsConditionalFormat
    Dim cCF2 As clsConditionalFormat
    Dim dControl As Dictionary
    Dim dRule As Dictionary
    Dim colRules As Collection

    ' Decode known fixture colors to RGB strings
    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex TEXT9_CF14
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("ForeColor") = "RGB(0,0,0)", "LongToRGB black"
    TestAssert dRule("BackColor") = "RGB(255,255,255)", "LongToRGB white"

    ' RGB string model round-trips byte-for-byte
    Set dControl = cCF.GetDictionary
    Set cCF2 = New clsConditionalFormat
    cCF2.LoadFromDictionary dControl
    TestAssert cCF2.BuildCF14Hex() = TEXT9_CF14, "RGB model rebuilds byte-exact"

    ' RGB string import round-trips through decode/rebuild
    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare
    dRule.Add "Type", "Expression"
    dRule.Add "Enabled", True
    dRule.Add "FontBold", False
    dRule.Add "FontItalic", False
    dRule.Add "FontUnderline", False
    dRule.Add "ForeColor", "RGB(255,0,0)"
    dRule.Add "BackColor", "RGB(0,128,255)"
    dRule.Add "Expression1", "[x]=1"
    Set colRules = New Collection
    colRules.Add dRule
    Set dControl = New Dictionary
    dControl.Add "Rules", colRules
    Set cCF = New clsConditionalFormat
    cCF.LoadFromDictionary dControl
    Set cCF2 = New clsConditionalFormat
    cCF2.LoadFromCF14Hex cCF.BuildCF14Hex()
    Set dRule = NthRule(cCF2, 1)
    TestAssert dRule("ForeColor") = "RGB(255,0,0)", "RGB(255,0,0) round-trips"
    TestAssert dRule("BackColor") = "RGB(0,128,255)", "RGB(0,128,255) round-trips"

    ' Legacy numeric Long values still import
    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare
    dRule.Add "Type", "Expression"
    dRule.Add "Enabled", True
    dRule.Add "FontBold", False
    dRule.Add "FontItalic", False
    dRule.Add "FontUnderline", False
    dRule.Add "ForeColor", 0
    dRule.Add "BackColor", 16777215
    dRule.Add "Expression1", "[fraOption]=1"
    Set colRules = New Collection
    colRules.Add dRule
    Set dControl = New Dictionary
    dControl.Add "Rules", colRules
    Set cCF = New clsConditionalFormat
    cCF.LoadFromDictionary dControl
    TestAssert cCF.BuildCF14Hex() = TEXT9_CF14, "numeric legacy colors rebuild byte-exact"

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
' Procedure : TestDecodeAllOperators
' Purpose   : Every AcFormatConditionOperator on a field-value rule decodes to the correct
'           : operator name (issue #725: operators other than Between were hardcoded).
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeAllOperators()

    TestAssert OperatorOf(OP_BETWEEN_CF14) = "Between", "operator 0 = Between"
    TestAssert OperatorOf(OP_NOTBETWEEN_CF14) = "NotBetween", "operator 1 = NotBetween"
    TestAssert OperatorOf(OP_EQUAL_CF14) = "Equal", "operator 2 = Equal"
    TestAssert OperatorOf(OP_NOTEQUAL_CF14) = "NotEqual", "operator 3 = NotEqual"
    TestAssert OperatorOf(OP_GREATERTHAN_CF14) = "GreaterThan", "operator 4 = GreaterThan"
    TestAssert OperatorOf(OP_LESSTHAN_CF14) = "LessThan", "operator 5 = LessThan"
    TestAssert OperatorOf(OP_GREATERTHANOREQUAL_CF14) = "GreaterThanOrEqual", "operator 6 = GreaterThanOrEqual"
    TestAssert OperatorOf(OP_LESSTHANOREQUAL_CF14) = "LessThanOrEqual", "operator 7 = LessThanOrEqual"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSingleValueOperatorHasEmptyExpression2
' Purpose   : Single-value operators (Equal, GreaterThan, ...) decode with an empty
'           : Expression2, while Between keeps both bounds.
'---------------------------------------------------------------------------------------
'
Public Sub TestSingleValueOperatorHasEmptyExpression2()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex OP_EQUAL_CF14
    TestAssert NthRule(cCF, 1)("Expression1") = "1", "Equal keeps Expression1"
    TestAssert NthRule(cCF, 1)("Expression2") = "", "Equal has empty Expression2"

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex OP_BETWEEN_CF14
    TestAssert NthRule(cCF, 1)("Expression2") = "2", "Between keeps Expression2"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactAllOperators
' Purpose   : The authoritative CF14 block rebuilds byte-for-byte for every operator,
'           : proving the operator (header offset 10) survives decode/rebuild.
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactAllOperators()

    TestAssert RebuildCF14(OP_BETWEEN_CF14) = OP_BETWEEN_CF14, "Between CF14 byte-exact"
    TestAssert RebuildCF14(OP_NOTBETWEEN_CF14) = OP_NOTBETWEEN_CF14, "NotBetween CF14 byte-exact"
    TestAssert RebuildCF14(OP_EQUAL_CF14) = OP_EQUAL_CF14, "Equal CF14 byte-exact"
    TestAssert RebuildCF14(OP_NOTEQUAL_CF14) = OP_NOTEQUAL_CF14, "NotEqual CF14 byte-exact"
    TestAssert RebuildCF14(OP_GREATERTHAN_CF14) = OP_GREATERTHAN_CF14, "GreaterThan CF14 byte-exact"
    TestAssert RebuildCF14(OP_LESSTHAN_CF14) = OP_LESSTHAN_CF14, "LessThan CF14 byte-exact"
    TestAssert RebuildCF14(OP_GREATERTHANOREQUAL_CF14) = OP_GREATERTHANOREQUAL_CF14, "GreaterThanOrEqual CF14 byte-exact"
    TestAssert RebuildCF14(OP_LESSTHANOREQUAL_CF14) = OP_LESSTHANOREQUAL_CF14, "LessThanOrEqual CF14 byte-exact"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestLegacyByteExactOperator
' Purpose   : The legacy block rebuilds byte-for-byte with the operator at offset 16, for
'           : both a two-bound Between rule and a single-value Equal rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestLegacyByteExactOperator()
    TestAssert RebuildLegacy(OP_BETWEEN_CF14) = OP_BETWEEN_LEGACY, "Between legacy byte-exact"
    TestAssert RebuildLegacy(OP_EQUAL_CF14) = OP_EQUAL_LEGACY, "Equal legacy byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestFieldValueTrailerEcho
' Purpose   : A colored field-value rule decodes its trailer BackColor echo (at trailer
'           : offset +5, not +9) and rebuilds it byte-for-byte.
'---------------------------------------------------------------------------------------
'
Public Sub TestFieldValueTrailerEcho()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex OP_EQUAL_CF14
    TestAssert NthRule(cCF, 1)("BackColor") = "RGB(255,0,0)", "field-value BackColor decoded"
    TestAssert NthRule(cCF, 1)("TrailerColor") = "RGB(255,0,0)", "field-value trailer echo at +5 decoded"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeMultiOperatorRules
' Purpose   : A three-rule field-value block with mixed operators decodes each rule's
'           : operator, expression, and color (issue #725 real-world shape). Rule 0's
'           : operator lives in the header; later rules' operators live in their prefix.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeMultiOperatorRules()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex MULTI_OP_CF14
    TestAssert RuleCount(cCF) = 3, "three field-value rules"

    TestAssert NthRule(cCF, 1)("Operator") = "Between", "rule 1 operator (header) = Between"
    TestAssert NthRule(cCF, 1)("Expression1") = "11", "rule 1 Expression1"
    TestAssert NthRule(cCF, 1)("Expression2") = "22", "rule 1 Expression2"
    TestAssert NthRule(cCF, 1)("BackColor") = "RGB(255,0,0)", "rule 1 red"

    TestAssert NthRule(cCF, 2)("Operator") = "Equal", "rule 2 operator (prefix) = Equal"
    TestAssert NthRule(cCF, 2)("Expression1") = "55", "rule 2 Expression1"
    TestAssert NthRule(cCF, 2)("Expression2") = "", "rule 2 Expression2 empty"
    TestAssert NthRule(cCF, 2)("BackColor") = "RGB(0,255,0)", "rule 2 green"

    TestAssert NthRule(cCF, 3)("Operator") = "GreaterThan", "rule 3 operator (prefix) = GreaterThan"
    TestAssert NthRule(cCF, 3)("Expression1") = "99", "rule 3 Expression1"
    TestAssert NthRule(cCF, 3)("BackColor") = "RGB(0,0,255)", "rule 3 blue"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactMultiOperator
' Purpose   : The mixed-operator multi-rule CF14 block rebuilds byte-for-byte (proves the
'           : per-rule prefix operator dword survives decode/rebuild).
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactMultiOperator()
    TestAssert RebuildCF14(MULTI_OP_CF14) = MULTI_OP_CF14, "multi-operator CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactDataBarAuto
' Purpose   : CF14 rebuilds byte-for-byte for a single data bar with automatic limits
'           : (issue #730 — Lowest Value / Highest Value).
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactDataBarAuto()
    TestAssert RebuildCF14(DATABAR_AUTO_CF14) = DATABAR_AUTO_CF14, "data bar auto CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactDataBarLimits
' Purpose   : CF14 rebuilds byte-for-byte for number and percent data bar limits.
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactDataBarLimits()
    TestAssert RebuildCF14(DATABAR_NUMBER_CF14) = DATABAR_NUMBER_CF14, "data bar number CF14 byte-exact"
    TestAssert RebuildCF14(DATABAR_PERCENT_CF14) = DATABAR_PERCENT_CF14, "data bar percent CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCF14ByteExactDataBarMultiRule
' Purpose   : CF14 rebuilds byte-for-byte when a data bar follows another rule (per-rule
'           : type prefix on rules after the first). Back-to-back data bars are covered by
'           : the captured corpus.
'---------------------------------------------------------------------------------------
'
Public Sub TestCF14ByteExactDataBarMultiRule()
    TestAssert RebuildCF14(EXPR_DATABAR_CF14) = EXPR_DATABAR_CF14, "expression + data bar CF14 byte-exact"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeDataBarAuto
' Purpose   : Automatic-limit data bar decodes limit types and empty value fields.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeDataBarAuto()

    Dim cCF As clsConditionalFormat
    Dim dRule As Dictionary

    Set cCF = New clsConditionalFormat
    TestAssert cCF.LoadFromCF14Hex(DATABAR_AUTO_CF14), "auto data bar decodes"
    TestAssert Not cCF.DecodeFailed, "auto data bar DecodeFailed is false"
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("Type") = "DataBar", "rule type is DataBar"
    TestAssert dRule("ShortestLimit") = "automatic", "shortest limit automatic"
    TestAssert dRule("LongestLimit") = "automatic", "longest limit automatic"
    TestAssert dRule("ShortestValue") = "", "automatic shortest has empty value"
    TestAssert dRule("LongestValue") = "", "automatic longest has empty value"
    TestAssert dRule("ShowBarOnly") = False, "ShowBarOnly default"
    TestAssert dRule("BarColor") = "RGB(28,36,0)", "bar color decoded from BGR bytes"
    TestAssert dRule("FillColor") = "RGB(255,255,255)", "fill color white"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeDataBarNumberPercent
' Purpose   : Number and percent limit types decode with their typed value strings.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeDataBarNumberPercent()

    Dim cCF As clsConditionalFormat
    Dim dRule As Dictionary

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex DATABAR_NUMBER_CF14
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("ShortestLimit") = "number", "number shortest limit"
    TestAssert dRule("LongestLimit") = "number", "number longest limit"
    TestAssert dRule("ShortestValue") = "10", "shortest value 10"
    TestAssert dRule("LongestValue") = "100", "longest value 100"

    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex DATABAR_PERCENT_CF14
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("ShortestLimit") = "percent", "percent shortest limit"
    TestAssert dRule("LongestLimit") = "percent", "percent longest limit"
    TestAssert dRule("ShortestValue") = "25", "shortest value 25"
    TestAssert dRule("LongestValue") = "75", "longest value 75"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeFailsOnTruncatedBlock
' Purpose   : A truncated CF14 block reports decode failure instead of raising an error.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeFailsOnTruncatedBlock()

    Dim cCF As clsConditionalFormat

    Set cCF = New clsConditionalFormat
    TestAssert Not cCF.LoadFromCF14Hex(Left$(DATABAR_AUTO_CF14, 80)), "truncated block fails decode"
    TestAssert cCF.DecodeFailed, "truncated block sets DecodeFailed"
    TestAssert Not cCF.HasRules, "truncated block loads no rules"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecodeFailsOnMalformedHex
' Purpose   : Hex that is not a whole number of bytes, or carries stray characters, is
'           : reported as a decode failure. Splitting an odd-length string into bytes used
'           : to raise "subscript out of range" from inside HexToBytes, which would lose
'           : the control's formatting rather than keep the inline block (issue #730).
'---------------------------------------------------------------------------------------
'
Public Sub TestDecodeFailsOnMalformedHex()

    Dim cCF As clsConditionalFormat

    ' Odd length: a byte split in half
    Set cCF = New clsConditionalFormat
    TestAssert Not cCF.LoadFromCF14Hex(Left$(DATABAR_AUTO_CF14, 81)), "odd-length hex fails decode"
    TestAssert cCF.DecodeFailed, "odd-length hex sets DecodeFailed"

    ' Non-hex characters
    Set cCF = New clsConditionalFormat
    TestAssert Not cCF.LoadFromCF14Hex("0100zz00"), "non-hex characters fail decode"
    TestAssert cCF.DecodeFailed, "non-hex characters set DecodeFailed"

    ' Embedded whitespace (a hand-wrapped block that was not re-joined cleanly)
    Set cCF = New clsConditionalFormat
    TestAssert Not cCF.LoadFromCF14Hex("0100 0100"), "embedded whitespace fails decode"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUnderlineFlagDoesNotOverflow
' Purpose   : A rule with FontUnderline puts &HFF in the high byte of the format-flags
'           : record. Reading that record as a single dword overflowed a signed Long and
'           : raised error 6, so every form with an underline rule failed to decode. The
'           : flags are four independent bytes and must be read as such (issue #730).
'---------------------------------------------------------------------------------------
'
Public Sub TestUnderlineFlagDoesNotOverflow()

    Dim cCF As clsConditionalFormat
    Dim dRule As Dictionary

    Set cCF = New clsConditionalFormat
    TestAssert cCF.LoadFromCF14Hex(FLAGS_BOLD_UNDERLINE_CF14), "underline rule decodes without error"
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule("FontBold") = True, "bold decoded"
    TestAssert dRule("FontItalic") = False, "italic not set"
    TestAssert dRule("FontUnderline") = True, "underline decoded"
    TestAssert cCF.BuildCF14Hex() = FLAGS_BOLD_UNDERLINE_CF14, "underline rule rebuilds byte-exact"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSanitizeKeepsInlineOnDecodeFailure
' Purpose   : When CF14 cannot be decoded during export sanitize, the inline binary block
'           : is preserved and no JSON entry is produced for that control.
'---------------------------------------------------------------------------------------
'
Public Sub TestSanitizeKeepsInlineOnDecodeFailure()

    Dim cParser As clsSourceParser
    Dim strForm As String
    Dim strOut As String

    strForm = BuildControlWithCF14("TextBad", Left$(DATABAR_AUTO_CF14, 80))
    Set cParser = New clsSourceParser
    cParser.LoadString strForm, edbForm
    cParser.ObjectName = "frmTest"
    strOut = cParser.Sanitize(ectObjectDefinition)

    TestAssert InStr(strOut, "ConditionalFormat14 = Begin") > 0, "failed decode keeps inline CF14"
    TestAssert cParser.GetConditionalFormats.Count = 0, "failed decode omits JSON entry"

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
' Procedure : TestTrailerColorRoundTrip
' Purpose   : The TrailerColor field survives a decode/JSON/rebuild round-trip, preserving
'           : the non-zero trailer echo bytes.
'---------------------------------------------------------------------------------------
'
Public Sub TestTrailerColorRoundTrip()

    Dim cCF As clsConditionalFormat
    Dim cCF2 As clsConditionalFormat
    Dim dModel As Dictionary
    Dim dRule As Dictionary

    ' Decode AlertText and verify TrailerColor was parsed
    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex ALERTTEXT_CF14
    Set dRule = NthRule(cCF, 1)
    TestAssert dRule.Exists("TrailerColor"), "rule 0 has TrailerColor"
    TestAssert dRule("TrailerColor") = "RGB(219,219,183)", "rule 0 TrailerColor value"
    Set dRule = NthRule(cCF, 4)
    TestAssert Not dRule.Exists("TrailerColor"), "rule 3 (CF14-only) has no TrailerColor"

    ' Round-trip through the dictionary model (simulates JSON save/load). Compared against
    ' a direct rebuild rather than the dialog-authored original, since the emitter writes
    ' the VBA Boolean encoding - the point here is that the JSON hop loses nothing.
    Set dModel = cCF.GetDictionary
    Set cCF2 = New clsConditionalFormat
    cCF2.LoadFromDictionary dModel
    TestAssert cCF2.BuildCF14Hex() = RebuildCF14(ALERTTEXT_CF14), _
        "AlertText CF14 unchanged by JSON round-trip"
    TestAssert InStr(cCF2.BuildCF14Hex(), "dbdbb7") > 0, "trailer echo bytes preserved"

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

'---------------------------------------------------------------------------------------
' Procedure : LegacyBlockSizeField
' Purpose   : Read the blockSize dword the legacy block declares at byte offset 4, so a
'           : test can check it against the block's actual length.
'---------------------------------------------------------------------------------------
'
Private Function LegacyBlockSizeField(strLegacyHex As String) As Long

    If Len(strLegacyHex) < 16 Then Exit Function
    ' Bytes 4-7 are chars 9-16; reverse them into one hex literal to read little-endian
    LegacyBlockSizeField = CLng("&h" & Mid$(strLegacyHex, 15, 2) & Mid$(strLegacyHex, 13, 2) & _
        Mid$(strLegacyHex, 11, 2) & Mid$(strLegacyHex, 9, 2))

End Function


Private Function RuleCount(cCF As clsConditionalFormat) As Long
    Dim dModel As Dictionary
    Set dModel = cCF.GetDictionary
    RuleCount = dModel("Rules").Count
End Function

Private Function OperatorOf(strCF14Hex As String) As String
    Dim cCF As clsConditionalFormat
    Set cCF = New clsConditionalFormat
    cCF.LoadFromCF14Hex strCF14Hex
    OperatorOf = NthRule(cCF, 1)("Operator")
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
' Procedure : BuildControlWithCF14
' Purpose   : Build a minimal control block with a full inline CF14 hex payload.
'---------------------------------------------------------------------------------------
'
Private Function BuildControlWithCF14(strName As String, strCF14Hex As String) As String
    BuildControlWithCF14 = BuildControl(strName, strCF14Hex)
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
