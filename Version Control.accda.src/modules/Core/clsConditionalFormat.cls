VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConditionalFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsConditionalFormat
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse and rebuild the undocumented ConditionalFormat / ConditionalFormat14
'           : binary blocks emitted by Access SaveAsText for text box / combo box
'           : controls. Converts the blocks to a human-readable rule model (for storage
'           : in the companion .json) and rebuilds both blocks on import.
'           :
'           : The CF14 block is the authoritative, complete copy (it carries every rule
'           : including data bars and per-rule types). The legacy ConditionalFormat block
'           : is single-type and may omit rules; it is regenerated from the same model to
'           : keep both blocks consistent.
'           :
'           : See docs/access-conditional-format.md for the full binary specification.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

' AcFormatConditionType values (also stored in the binary blocks)
Private Const CF_TYPE_FIELDVALUE As Long = 0
Private Const CF_TYPE_EXPRESSION As Long = 1
Private Const CF_TYPE_FIELDHASFOCUS As Long = 2
Private Const CF_TYPE_DATABAR As Long = 3

' Trailing/base size of a non-data-bar rule body (excludes the optional 8-byte
' per-rule type prefix). Verified across the fixture: body length always equals
' BASE_BODY + 2 * (total UTF-16 code units across the rule's expression fields).
Private Const BASE_BODY As Long = 37

' Length (in bytes) of the per-rule type prefix that precedes every rule after
' the first (a conditionType dword followed by a reserved dword).
Private Const RULE_PREFIX_LEN As Long = 8

' Byte offset at which the legacy block's concatenated expression buffer begins.
' Verified across the single-rule fixtures (expression, focus, between).
Private Const LEGACY_EXPR_OFFSET As Long = 96

' Access 2007 legacy format caps at 3 rules per block.
Private Const LEGACY_RULE_CAP As Long = 3

' Minimum size Access allocates for a legacy expression slot, even when unused.
Private Const LEGACY_MIN_SLOT_UNITS As Long = 2

' Fixed trailer size on a data bar record (showBarOnly through longestLimit padding).
Private Const DATABAR_TRAILER_LEN As Long = 13

' Byte Access writes for a True ShowBarOnly flag (VBA True widened to a byte).
Private Const DATABAR_BAR_ONLY_TRUE As Long = &HFF&

' Decoded rule model. Each item is a Dictionary describing one format rule.
Private m_colRules As Collection
Private m_lngVersion As Long
Private m_blnDecodeFailed As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    m_lngVersion = 1
    Set m_colRules = New Collection
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasRules
' Purpose   : Returns true if any rules are loaded in the model.
'---------------------------------------------------------------------------------------
'
Public Function HasRules() As Boolean
    HasRules = (m_colRules.Count > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DecodeFailed
' Purpose   : True when the most recent LoadFromCF14Hex call could not decode the block.
'---------------------------------------------------------------------------------------
'
Public Property Get DecodeFailed() As Boolean
    DecodeFailed = m_blnDecodeFailed
End Property


'---------------------------------------------------------------------------------------
' Procedure : LoadFromCF14Hex
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse a ConditionalFormat14 hex string (concatenated, no "0x" prefix or
'           : separators) into the decoded rule model. CF14 is the authoritative source.
'           : Returns True when the full block was decoded; False on structural failure.
'---------------------------------------------------------------------------------------
'
Public Function LoadFromCF14Hex(strHex As String) As Boolean

    Dim bteData() As Byte
    Dim lngPos As Long
    Dim lngRuleCount As Long
    Dim lngFirstType As Long
    Dim lngType As Long
    Dim lngOperator As Long
    Dim lngRule As Long
    Dim dRule As Dictionary

    m_blnDecodeFailed = False
    Set m_colRules = New Collection
    LoadFromCF14Hex = False
    If Len(strHex) = 0 Then Exit Function

    ' Reject anything that is not a whole number of hex bytes before converting. A block
    ' truncated mid-byte, or one carrying stray characters, must be reported as a decode
    ' failure so the caller keeps the inline binary block - never raise, since an
    ' unhandled error here loses the user's conditional formatting (issue #730).
    If Not IsHexBytes(strHex) Then
        m_blnDecodeFailed = True
        Exit Function
    End If

    bteData = HexToBytes(strHex)
    If UBound(bteData) < 13 Then
        m_blnDecodeFailed = True
        Exit Function
    End If

    ' Parse 14-byte header
    m_lngVersion = ReadInt(bteData, 0)
    lngRuleCount = ReadInt(bteData, 2)
    lngFirstType = ReadInt(bteData, 6)
    If lngRuleCount < 1 Then
        m_blnDecodeFailed = True
        Exit Function
    End If
    lngPos = 14

    For lngRule = 0 To lngRuleCount - 1

        ' Rules after the first carry an 8-byte type prefix (type dword + operator
        ' dword). Data bars use the same prefix for rules after the first; the first
        ' rule's type comes from the header (offset 6) with no per-rule prefix.
        ' For the first rule the field-value operator comes from offset 10 (docs §4.1).
        ' For later rules the operator is the second dword of the prefix (lngPos + 4).
        If lngRule = 0 Then
            lngType = lngFirstType
            If Not CanReadBytes(bteData, 10, 2) Then
                m_blnDecodeFailed = True
                Exit Function
            End If
            lngOperator = ReadInt(bteData, 10)
        Else
            If Not CanReadBytes(bteData, lngPos, RULE_PREFIX_LEN) Then
                m_blnDecodeFailed = True
                Exit Function
            End If
            lngType = ReadLong(bteData, lngPos)
            lngOperator = ReadLong(bteData, lngPos + 4)
        End If

        Select Case lngType
            Case CF_TYPE_DATABAR
                Set dRule = ParseDataBarRule(bteData, lngPos, (lngRule > 0))
            Case Else
                ' Skip the prefix for subsequent non-data-bar rules
                If lngRule > 0 Then lngPos = lngPos + RULE_PREFIX_LEN
                Set dRule = ParseStandardRule(bteData, lngPos, lngType, lngOperator)
        End Select

        If dRule Is Nothing Then
            m_blnDecodeFailed = True
            Exit Function
        End If
        m_colRules.Add dRule
    Next lngRule

    If m_colRules.Count <> lngRuleCount Then
        m_blnDecodeFailed = True
        Exit Function
    End If
    If lngPos <> UBound(bteData) + 1 Then
        m_blnDecodeFailed = True
        Exit Function
    End If

    LoadFromCF14Hex = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseStandardRule
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse an expression / field-value / field-has-focus rule body beginning
'           : at lngPos. Advances lngPos past the full rule body (including trailing
'           : padding) so the caller can read the next rule.
'---------------------------------------------------------------------------------------
'
Private Function ParseStandardRule(bteData() As Byte, ByRef lngPos As Long, lngType As Long, _
    Optional lngOperator As Long = 0) As Dictionary

    Dim dRule As Dictionary
    Dim lngLen1 As Long
    Dim lngLen2 As Long
    Dim lngUnits As Long
    Dim lngTrailPos As Long
    Dim lngTrailEcho As Long
    Dim lngEcho As Long
    Dim lngBodyEnd As Long

    If Not CanReadBytes(bteData, lngPos, BASE_BODY) Then Exit Function

    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare

    ' Format flags occupy four independent bytes, not a bitfield. Read them byte-wise:
    ' a set font flag holds VBA True (&HFF), so the underline byte alone would overflow
    ' a signed Long if the record were read as a single dword (docs section 6).
    dRule.Add "Type", TypeToName(lngType)
    dRule.Add "Enabled", CBool(bteData(lngPos) <> 0)
    dRule.Add "FontBold", CBool(bteData(lngPos + 1) <> 0)
    dRule.Add "FontItalic", CBool(bteData(lngPos + 2) <> 0)
    dRule.Add "FontUnderline", CBool(bteData(lngPos + 3) <> 0)
    dRule.Add "ForeColor", LongToRGB(ReadLong(bteData, lngPos + 4))
    dRule.Add "BackColor", LongToRGB(ReadLong(bteData, lngPos + 8))

    ' Expression length and text (focus rules have length 0)
    If Not CanReadBytes(bteData, lngPos + 12, 4) Then Exit Function
    lngLen1 = ReadLong(bteData, lngPos + 12)
    lngUnits = lngLen1
    If lngType = CF_TYPE_FIELDVALUE Then
        ' Both bounds stored inline, each length-prefixed. Expression2 is only
        ' meaningful for Between / NotBetween; single-value operators leave it empty.
        dRule.Add "Operator", OperatorToName(lngOperator)
        If Not CanReadUtf16(bteData, lngPos + 16, lngLen1) Then Exit Function
        dRule.Add "Expression1", ReadUtf16(bteData, lngPos + 16, lngLen1)
        If Not CanReadBytes(bteData, lngPos + 16 + (lngLen1 * 2), 4) Then Exit Function
        lngLen2 = ReadLong(bteData, lngPos + 16 + (lngLen1 * 2))
        If Not CanReadUtf16(bteData, lngPos + 20 + (lngLen1 * 2), lngLen2) Then Exit Function
        dRule.Add "Expression2", ReadUtf16(bteData, lngPos + 20 + (lngLen1 * 2), lngLen2)
        lngUnits = lngLen1 + lngLen2
    ElseIf lngType = CF_TYPE_EXPRESSION Then
        If Not CanReadUtf16(bteData, lngPos + 16, lngLen1) Then Exit Function
        dRule.Add "Expression1", ReadUtf16(bteData, lngPos + 16, lngLen1)
    End If

    lngBodyEnd = lngPos + BASE_BODY + (lngUnits * 2)
    If Not CanReadBytes(bteData, lngPos, lngBodyEnd - lngPos) Then Exit Function

    ' Read the CF14 trailer BackColor echo (3 bytes). The echo sits (trailerLen - 12)
    ' bytes into the trailing region, i.e. +9 for the 21-byte expression/focus trailer
    ' and +5 for the 17-byte field-value trailer (verified empirically against colored
    ' field-value rules; docs section 4.2).
    If lngType = CF_TYPE_FIELDVALUE Then
        lngTrailPos = lngPos + 20 + (lngUnits * 2)
        lngTrailEcho = 5
    Else
        lngTrailPos = lngPos + 16 + (lngUnits * 2)
        lngTrailEcho = 9
    End If
    If lngTrailPos + lngTrailEcho + 2 <= UBound(bteData) Then
        lngEcho = bteData(lngTrailPos + lngTrailEcho) + _
            (bteData(lngTrailPos + lngTrailEcho + 1) * &H100&) + _
            (bteData(lngTrailPos + lngTrailEcho + 2) * &H10000)
        If lngEcho <> 0 Then dRule.Add "TrailerColor", LongToRGB(lngEcho)
    End If

    ' Advance past the full body (base size plus 2 bytes per expression code unit)
    lngPos = lngBodyEnd

    Set ParseStandardRule = dRule

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseDataBarRule
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse a data bar rule record (CF14 only). When blnHasPrefix is True,
'           : lngPos points at the type dword; otherwise it points at unk1 (rule 0).
'           : Advances lngPos past the full record. Offsets follow docs §4.4.
'---------------------------------------------------------------------------------------
'
Private Function ParseDataBarRule(bteData() As Byte, ByRef lngPos As Long, _
    blnHasPrefix As Boolean) As Dictionary

    Dim dRule As Dictionary
    Dim lngShortLen As Long
    Dim lngLongLen As Long
    Dim lngBody As Long
    Dim lngP As Long
    Dim lngRecordEnd As Long

    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare
    dRule.Add "Type", TypeToName(CF_TYPE_DATABAR)

    If blnHasPrefix Then
        If Not CanReadBytes(bteData, lngPos, RULE_PREFIX_LEN) Then Exit Function
        lngBody = lngPos + RULE_PREFIX_LEN
    Else
        lngBody = lngPos
    End If

    If Not CanReadBytes(bteData, lngBody, 20) Then Exit Function
    dRule.Add "FillColor", LongToRGB(ReadLong(bteData, lngBody + 8))
    lngShortLen = ReadLong(bteData, lngBody + 12)
    If Not CanReadUtf16(bteData, lngBody + 16, lngShortLen) Then Exit Function
    dRule.Add "ShortestValue", ReadUtf16(bteData, lngBody + 16, lngShortLen)
    If Not CanReadBytes(bteData, lngBody + 16 + (lngShortLen * 2), 4) Then Exit Function
    lngLongLen = ReadLong(bteData, lngBody + 16 + (lngShortLen * 2))
    If Not CanReadUtf16(bteData, lngBody + 20 + (lngShortLen * 2), lngLongLen) Then Exit Function
    dRule.Add "LongestValue", ReadUtf16(bteData, lngBody + 20 + (lngShortLen * 2), lngLongLen)

    ' Trailer position relative to the body start (after value fields and unk2)
    lngP = lngBody + 20 + (2 * (lngShortLen + lngLongLen)) + 4
    lngRecordEnd = lngP + DATABAR_TRAILER_LEN
    If Not CanReadBytes(bteData, lngP, DATABAR_TRAILER_LEN) Then Exit Function
    dRule.Add "ShowBarOnly", CBool(bteData(lngP) <> 0)
    dRule.Add "BarColor", LongToRGB(bteData(lngP + 1) + (bteData(lngP + 2) * &H100&) + (bteData(lngP + 3) * &H10000))
    dRule.Add "ShortestLimit", LimitToName(bteData(lngP + 5))
    dRule.Add "LongestLimit", LimitToName(bteData(lngP + 9))

    lngPos = lngRecordEnd

    Set ParseDataBarRule = dRule

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Return the decoded model for storage in the companion .json file.
'---------------------------------------------------------------------------------------
'
Public Function GetDictionary() As Dictionary

    Dim dResult As Dictionary
    Dim colRules As Collection
    Dim varRule As Variant

    Set dResult = New Dictionary
    dResult.CompareMode = TextCompare
    Set colRules = New Collection
    For Each varRule In m_colRules
        colRules.Add varRule
    Next varRule
    dResult.Add "Rules", colRules
    Set GetDictionary = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFromDictionary
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Load the decoded model from a companion .json dictionary (one control).
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromDictionary(dControl As Dictionary)

    Dim varRule As Variant

    Set m_colRules = New Collection
    m_lngVersion = 1
    If dControl Is Nothing Then Exit Sub
    If Not dControl.Exists("Rules") Then Exit Sub
    For Each varRule In dControl("Rules")
        m_colRules.Add varRule
    Next varRule

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildCF14Hex
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Rebuild the ConditionalFormat14 hex byte string from the decoded model.
'           : This block carries every rule (the authoritative copy).
'---------------------------------------------------------------------------------------
'
Public Function BuildCF14Hex() As String

    Dim cData As clsConcat
    Dim lngIndex As Long
    Dim dRule As Dictionary
    Dim lngFirstType As Long
    Dim lngFirstOperator As Long

    If m_colRules.Count = 0 Then Exit Function

    ' First rule type populates the header discriminant. Its operator (field-value
    ' rules only) is stored at header offset 10; Access ignores it for other types.
    lngFirstType = NameToType(m_colRules(1)("Type"))
    If lngFirstType = CF_TYPE_FIELDVALUE Then
        lngFirstOperator = NameToOperator(dNZ(m_colRules(1), "Operator", "Between"))
    End If

    Set cData = New clsConcat
    ' Header: version(2) ruleCount(2) reserved(2) firstType(2) reserved(2) operator(2) reserved(2)
    cData.Add WriteInt(m_lngVersion)
    cData.Add WriteInt(m_colRules.Count)
    cData.Add WriteInt(0)
    cData.Add WriteInt(lngFirstType)
    cData.Add WriteInt(0)
    cData.Add WriteInt(lngFirstOperator)
    cData.Add WriteInt(0)

    For lngIndex = 1 To m_colRules.Count
        Set dRule = m_colRules(lngIndex)
        cData.Add EmitRuleCF14(dRule, (lngIndex = 1))
    Next lngIndex

    BuildCF14Hex = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : EmitRuleCF14
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Emit the hex byte string for a single CF14 rule. Rules after the first
'           : are preceded by an 8-byte type prefix (data bars embed it themselves).
'---------------------------------------------------------------------------------------
'
Private Function EmitRuleCF14(dRule As Dictionary, blnFirst As Boolean) As String

    Dim cData As clsConcat
    Dim lngType As Long
    Dim lngLen1 As Long
    Dim lngLen2 As Long
    Dim lngUnits As Long
    Dim lngTrailLen As Long

    Set cData = New clsConcat
    lngType = NameToType(dRule("Type"))

    If lngType = CF_TYPE_DATABAR Then
        ' Rules after the first carry the type prefix; rule 0 uses the header type only.
        If Not blnFirst Then
            cData.Add WriteLong(CF_TYPE_DATABAR)
            cData.Add WriteLong(0)
        End If
        lngLen1 = Utf16Units(dNZ(dRule, "ShortestValue"))
        lngLen2 = Utf16Units(dNZ(dRule, "LongestValue"))
        cData.Add WriteLong(1)                      ' unk1
        cData.Add WriteLong(0)                      ' reserved
        cData.Add WriteLong(RGBToLong(dNZ(dRule, "FillColor", "0")))
        cData.Add WriteLong(lngLen1)
        cData.Add WriteUtf16(dNZ(dRule, "ShortestValue"))
        cData.Add WriteLong(lngLen2)
        cData.Add WriteUtf16(dNZ(dRule, "LongestValue"))
        cData.Add WriteLong(1)                      ' unk2
        ' Trailer
        cData.Add WriteByte(IIf(CBool(dNZ(dRule, "ShowBarOnly", "0")), DATABAR_BAR_ONLY_TRUE, 0))
        cData.Add WriteColor3(RGBToLong(dNZ(dRule, "BarColor", "0")))
        cData.Add WriteByte(0)                      ' high byte of bar color
        cData.Add WriteByte(NameToLimit(dNZ(dRule, "ShortestLimit", "automatic")))
        cData.Add WriteZeros(3)
        cData.Add WriteByte(NameToLimit(dNZ(dRule, "LongestLimit", "automatic")))
        cData.Add WriteZeros(3)
        EmitRuleCF14 = cData.GetStr
        Exit Function
    End If

    ' Standard rule (expression / field-value / focus)
    If Not blnFirst Then
        cData.Add WriteLong(lngType)                ' type prefix
        If lngType = CF_TYPE_FIELDVALUE Then
            cData.Add WriteLong(NameToOperator(dNZ(dRule, "Operator", "Between")))
        Else
            cData.Add WriteLong(0)                  ' operator (ignored for non field-value)
        End If
    End If

    cData.Add WriteFlags(dRule)
    cData.Add WriteLong(RGBToLong(dNZ(dRule, "ForeColor", "0")))
    cData.Add WriteLong(RGBToLong(dNZ(dRule, "BackColor", "0")))

    Select Case lngType
        Case CF_TYPE_FIELDVALUE
            lngLen1 = Utf16Units(dNZ(dRule, "Expression1"))
            lngLen2 = Utf16Units(dNZ(dRule, "Expression2"))
            cData.Add WriteLong(lngLen1)
            cData.Add WriteUtf16(dNZ(dRule, "Expression1"))
            cData.Add WriteLong(lngLen2)
            cData.Add WriteUtf16(dNZ(dRule, "Expression2"))
            lngUnits = lngLen1 + lngLen2
        Case CF_TYPE_EXPRESSION
            lngLen1 = Utf16Units(dNZ(dRule, "Expression1"))
            cData.Add WriteLong(lngLen1)
            cData.Add WriteUtf16(dNZ(dRule, "Expression1"))
            lngUnits = lngLen1
        Case Else   ' focus
            cData.Add WriteLong(0)
            lngUnits = 0
    End Select

    ' Emit the explicit trailing region: (trailerLen - 12) lead zeros, a 3-byte
    ' BackColor echo, then 9 trailing zero bytes. This places the echo at +9 for the
    ' 21-byte expression/focus trailer and at +5 for the 17-byte field-value trailer
    ' (verified empirically against colored field-value rules; docs section 4.2).
    If lngType = CF_TYPE_FIELDVALUE Then lngTrailLen = 17 Else lngTrailLen = 21
    cData.Add WriteZeros(lngTrailLen - 12)
    If dRule.Exists("TrailerColor") Then
        cData.Add WriteColor3(RGBToLong(CStr(dRule("TrailerColor"))))
    Else
        cData.Add WriteZeros(3)
    End If
    cData.Add WriteZeros(9)

    EmitRuleCF14 = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildLegacyHex
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Rebuild the legacy ConditionalFormat hex byte string. The legacy block is
'           : single-type: it carries only the rules whose type matches the first rule's
'           : type (data bars and other off-type rules live only in CF14). See docs §5.
'           :
'           : Every rule owns two expression slots (Expression1, Expression2) sized by
'           : SlotUnits, regardless of type - an expression rule simply leaves the second
'           : slot empty. Rules before the last carry four extra descriptor dwords naming
'           : the following rule's type, operator, and slot window. Getting any of this
'           : wrong is not benign: Access silently discards the legacy blocks of *every*
'           : control on the form when one block is malformed.
'---------------------------------------------------------------------------------------
'
Public Function BuildLegacyHex() As String

    Dim cBody As clsConcat
    Dim cExpr As clsConcat
    Dim cData As clsConcat
    Dim lngIndex As Long
    Dim dRule As Dictionary
    Dim dNext As Dictionary
    Dim lngType As Long
    Dim lngHeaderType As Long
    Dim lngCount As Long
    Dim lngBlockSize As Long
    Dim lngHeaderBytes As Long
    Dim lngNextStart As Long
    Dim colMatch As Collection

    If m_colRules.Count = 0 Then Exit Function

    ' The legacy block is single-type; use the first rule's type as the header type.
    lngHeaderType = NameToType(m_colRules(1)("Type"))

    ' Collect the rules that the legacy block can represent (matching type)
    Set colMatch = New Collection
    For lngIndex = 1 To m_colRules.Count
        Set dRule = m_colRules(lngIndex)
        lngType = NameToType(dRule("Type"))
        If lngType = lngHeaderType And lngType <> CF_TYPE_DATABAR Then colMatch.Add dRule
    Next lngIndex

    ' Cap at the Access 2007 legacy limit
    Do While colMatch.Count > LEGACY_RULE_CAP
        colMatch.Remove colMatch.Count
    Loop
    lngCount = colMatch.Count

    ' A data-bar-only control has nothing the legacy block can represent. Access emits a
    ' rule-less 96-byte block in that case, but discards it again on the next design-view
    ' save, so its presence carries no information and the companion JSON does not record
    ' it. Returning an empty string omits the block entirely (see FormatCFBlock).
    If lngCount = 0 Then Exit Function

    ' Build per-rule body records and the expression slot buffer.
    Set cBody = New clsConcat
    Set cExpr = New clsConcat

    ' Running start unit of the next rule's slot window
    lngNextStart = RuleSlotUnits(colMatch(1))

    For lngIndex = 1 To lngCount
        Set dRule = colMatch(lngIndex)
            cBody.Add WriteFlags(dRule)
        cBody.Add WriteLong(RGBToLong(dNZ(dRule, "ForeColor", "0")))
        cBody.Add WriteLong(RGBToLong(dNZ(dRule, "BackColor", "0")))

        If lngIndex < lngCount Then
            ' Non-last rule: append descriptor dwords describing the next rule (28-byte record)
            Set dNext = colMatch(lngIndex + 1)
            cBody.Add WriteLong(NameToType(dNext("Type")))
            cBody.Add WriteLong(LegacyOperator(NameToType(dNext("Type")), dNext))
            cBody.Add WriteLong(lngNextStart)
            cBody.Add WriteLong(lngNextStart + SlotUnits(dNZ(dNext, "Expression1")))
            lngNextStart = lngNextStart + RuleSlotUnits(dNext)
        End If

        cExpr.Add WriteUtf16Padded(dNZ(dRule, "Expression1"), SlotUnits(dNZ(dRule, "Expression1")))
        cExpr.Add WriteUtf16Padded(dNZ(dRule, "Expression2"), SlotUnits(dNZ(dRule, "Expression2")))
    Next lngIndex

    ' Assemble: header (28 bytes) + per-rule body + padding to offset 96 + slot buffer.
    Set cData = New clsConcat
    cData.Add WriteLong(1)                              ' version
    cData.Add WriteLong(0)                              ' blockSize placeholder
    cData.Add WriteLong(lngCount)                       ' ruleCount
    cData.Add WriteLong(lngHeaderType)                  ' conditionType
    cData.Add WriteLong(LegacyOperator(lngHeaderType, colMatch(1)))  ' operator (offset 16)
    cData.Add WriteLong(0)                              ' reserved
    cData.Add WriteLong(SlotUnits(dNZ(colMatch(1), "Expression1")))  ' first slot end unit
    cData.Add cBody.GetStr
    lngHeaderBytes = Len(cData.GetStr) \ 2
    If LEGACY_EXPR_OFFSET > lngHeaderBytes Then cData.Add WriteZeros(LEGACY_EXPR_OFFSET - lngHeaderBytes)
    cData.Add cExpr.GetStr

    ' Patch in the real block size at offset 4
    lngBlockSize = Len(cData.GetStr) \ 2
    BuildLegacyHex = PatchLong(cData.GetStr, 4, lngBlockSize)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SlotUnits / RuleSlotUnits
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Size of one legacy expression slot in UTF-16 code units, and the combined
'           : size of a rule's two slots. A slot holds the text plus a null terminator,
'           : but Access never writes a slot shorter than two units - an unused slot (an
'           : expression rule's Expression2, or either slot on a field-has-focus rule)
'           : still occupies two null units.
'---------------------------------------------------------------------------------------
'
Private Function SlotUnits(strText As String) As Long
    SlotUnits = MaxLong(Utf16Units(strText) + 1, LEGACY_MIN_SLOT_UNITS)
End Function

Private Function RuleSlotUnits(dRule As Dictionary) As Long
    RuleSlotUnits = SlotUnits(dNZ(dRule, "Expression1")) + SlotUnits(dNZ(dRule, "Expression2"))
End Function


' ===================================================================================
' Byte/value helpers
' ===================================================================================

'---------------------------------------------------------------------------------------
' Procedure : IsHexBytes
' Purpose   : True when the string is a whole number of hex bytes and nothing else, so
'           : HexToBytes can convert it without raising. Callers must screen input with
'           : this rather than relying on an error handler.
'---------------------------------------------------------------------------------------
'
Private Function IsHexBytes(strHex As String) As Boolean

    Dim lngChar As Long
    Dim intCode As Integer

    If Len(strHex) = 0 Then Exit Function
    If (Len(strHex) Mod 2) <> 0 Then Exit Function
    For lngChar = 1 To Len(strHex)
        intCode = AscW(Mid$(strHex, lngChar, 1))
        Select Case intCode
            Case 48 To 57, 65 To 70, 97 To 102     ' 0-9, A-F, a-f
            Case Else: Exit Function
        End Select
    Next lngChar
    IsHexBytes = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : HexToBytes
' Purpose   : Convert a hex string (2 chars per byte) into a zero-based byte array.
'           : Screen input with IsHexBytes first.
'---------------------------------------------------------------------------------------
'
Private Function HexToBytes(strHex As String) As Byte()

    Dim bteData() As Byte
    Dim lngByte As Long
    Dim lngCount As Long
    Dim strByte As String

    lngCount = Len(strHex) \ 2
    If lngCount = 0 Then Exit Function

    strByte = "&h00"
    ReDim bteData(0 To lngCount - 1)
    ' Drive the loop from the byte count rather than the string length; an odd-length
    ' string would otherwise run one iteration past the end of the array.
    For lngByte = 0 To lngCount - 1
        Mid$(strByte, 3, 2) = Mid$(strHex, (lngByte * 2) + 1, 2)
        bteData(lngByte) = CByte(CLng(strByte))
    Next lngByte
    HexToBytes = bteData

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadInt / ReadLong
' Purpose   : Read a little-endian 16-bit / 32-bit value from a byte array. ReadInt is
'           : unsigned; ReadLong returns the 32-bit pattern as a signed Long, since a
'           : VBA Long cannot hold values above &H7FFFFFFF. Multiplying the high byte
'           : into place overflows once it reaches &H80, which real conditional-format
'           : blocks reach routinely, so the sign bit is folded in separately.
'---------------------------------------------------------------------------------------
'
Private Function ReadInt(bteData() As Byte, lngPos As Long) As Long
    ReadInt = bteData(lngPos) + (bteData(lngPos + 1) * &H100&)
End Function

Private Function ReadLong(bteData() As Byte, lngPos As Long) As Long
    Dim lngHigh As Long
    lngHigh = bteData(lngPos + 3)
    ReadLong = bteData(lngPos) + (bteData(lngPos + 1) * &H100&) + _
        (bteData(lngPos + 2) * &H10000) + ((lngHigh And &H7F&) * &H1000000)
    If (lngHigh And &H80&) <> 0 Then ReadLong = ReadLong Or &H80000000
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadUtf16
' Purpose   : Read a UTF-16LE string of lngUnits code units from a byte array.
'---------------------------------------------------------------------------------------
'
Private Function ReadUtf16(bteData() As Byte, lngPos As Long, lngUnits As Long) As String
    Dim bteStr() As Byte
    Dim lngByte As Long
    If lngUnits <= 0 Then Exit Function
    If Not CanReadUtf16(bteData, lngPos, lngUnits) Then Exit Function
    ReDim bteStr(0 To (lngUnits * 2) - 1)
    For lngByte = 0 To (lngUnits * 2) - 1
        bteStr(lngByte) = bteData(lngPos + lngByte)
    Next lngByte
    ReadUtf16 = bteStr
End Function


'---------------------------------------------------------------------------------------
' Procedure : CanReadBytes / CanReadUtf16
' Purpose   : Bounds checks before reading from a CF14 byte array.
'---------------------------------------------------------------------------------------
'
Private Function CanReadBytes(bteData() As Byte, lngPos As Long, lngCount As Long) As Boolean
    If lngCount <= 0 Then
        CanReadBytes = True
    Else
        CanReadBytes = (lngPos >= LBound(bteData) And lngPos + lngCount - 1 <= UBound(bteData))
    End If
End Function

Private Function CanReadUtf16(bteData() As Byte, lngPos As Long, lngUnits As Long) As Boolean
    If lngUnits <= 0 Then
        CanReadUtf16 = True
    Else
        CanReadUtf16 = CanReadBytes(bteData, lngPos, lngUnits * 2)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteInt / WriteLong / WriteByte / WriteZeros
' Purpose   : Emit little-endian hex byte strings for integer values.
'---------------------------------------------------------------------------------------
'
Private Function WriteInt(ByVal lngValue As Long) As String
    WriteInt = ByteHex(lngValue And &HFF) & ByteHex((lngValue \ &H100) And &HFF)
End Function

Private Function WriteLong(ByVal lngValue As Long) As String
    ' Mask each byte before dividing. Dividing first truncates toward zero, which yields
    ' the wrong high bytes for any negative value - i.e. any dword whose high bit is set,
    ' which is what ReadLong hands back for such a record.
    WriteLong = ByteHex(lngValue And &HFF&) & ByteHex((lngValue And &HFF00&) \ &H100&) & _
        ByteHex((lngValue And &HFF0000) \ &H10000) & _
        ByteHex((lngValue And &HFF000000) \ &H1000000)
End Function

Private Function WriteByte(ByVal lngValue As Long) As String
    WriteByte = ByteHex(lngValue And &HFF)
End Function

Private Function WriteZeros(ByVal lngCount As Long) As String
    If lngCount > 0 Then WriteZeros = String$(lngCount * 2, "0")
End Function

Private Function WriteColor3(ByVal lngColor As Long) As String
    ' Low three bytes of an RGB long (BGR byte order in the stream)
    WriteColor3 = ByteHex(lngColor And &HFF&) & ByteHex((lngColor \ &H100&) And &HFF) & _
        ByteHex((lngColor \ &H10000) And &HFF)
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteUtf16 / WriteUtf16Padded
' Purpose   : Emit a UTF-16LE string (optionally padded to a code-unit capacity).
'---------------------------------------------------------------------------------------
'
Private Function WriteUtf16(strText As String) As String
    Dim bteStr() As Byte
    Dim lngByte As Long
    Dim cData As clsConcat
    If Len(strText) = 0 Then Exit Function
    bteStr = strText
    Set cData = New clsConcat
    For lngByte = 0 To UBound(bteStr)
        cData.Add ByteHex(bteStr(lngByte))
    Next lngByte
    WriteUtf16 = cData.GetStr
End Function

Private Function WriteUtf16Padded(strText As String, lngCapacityUnits As Long) As String
    Dim strHex As String
    Dim lngUnits As Long
    strHex = WriteUtf16(strText)
    lngUnits = Utf16Units(strText)
    ' Pad remaining capacity (including the null terminator) with zeros
    If lngCapacityUnits > lngUnits Then strHex = strHex & WriteZeros((lngCapacityUnits - lngUnits) * 2)
    WriteUtf16Padded = strHex
End Function


'---------------------------------------------------------------------------------------
' Procedure : ByteHex
' Purpose   : Two-character lowercase hex for a single byte value.
'---------------------------------------------------------------------------------------
'
Private Function ByteHex(ByVal lngValue As Long) As String
    Dim strHex As String
    strHex = LCase$(Hex$(lngValue And &HFF))
    If Len(strHex) = 1 Then strHex = "0" & strHex
    ByteHex = strHex
End Function


'---------------------------------------------------------------------------------------
' Procedure : PatchLong
' Purpose   : Overwrite a 4-byte little-endian value at a byte offset in a hex string.
'---------------------------------------------------------------------------------------
'
Private Function PatchLong(strHex As String, lngByteOffset As Long, lngValue As Long) As String
    Dim lngCharPos As Long
    lngCharPos = (lngByteOffset * 2) + 1
    Mid$(strHex, lngCharPos, 8) = WriteLong(lngValue)
    PatchLong = strHex
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteFlags
' Purpose   : Emit the 4-byte format-flags record from the rule's font/enabled booleans.
'           : Byte 0 = enabled, 1 = bold, 2 = italic, 3 = underline (docs §6). These are
'           : four independent bytes rather than a bitfield, and the two groups do not
'           : share a convention: Access writes 1 for enabled but VBA True (&HFF) for a
'           : set font flag. Emitting them as bytes also keeps the underline byte from
'           : forcing a negative Long through the dword writer.
'---------------------------------------------------------------------------------------
'
Private Function WriteFlags(dRule As Dictionary) As String
    WriteFlags = WriteByte(IIf(CBool(dNZ(dRule, "Enabled", "True")), 1, 0)) & _
        WriteFontFlag(dNZ(dRule, "FontBold", "False")) & _
        WriteFontFlag(dNZ(dRule, "FontItalic", "False")) & _
        WriteFontFlag(dNZ(dRule, "FontUnderline", "False"))
End Function

Private Function WriteFontFlag(varValue As Variant) As String
    WriteFontFlag = WriteByte(IIf(CBool(varValue), &HFF, 0))
End Function


'---------------------------------------------------------------------------------------
' Procedure : Type / operator / limit name mapping
'---------------------------------------------------------------------------------------
'
Private Function TypeToName(lngType As Long) As String
    Select Case lngType
        Case CF_TYPE_FIELDVALUE:    TypeToName = "FieldValue"
        Case CF_TYPE_EXPRESSION:    TypeToName = "Expression"
        Case CF_TYPE_FIELDHASFOCUS: TypeToName = "FieldHasFocus"
        Case CF_TYPE_DATABAR:       TypeToName = "DataBar"
        Case Else:                  TypeToName = "Expression"
    End Select
End Function

Private Function NameToType(varName As Variant) As Long
    Select Case CStr(varName)
        Case "FieldValue":      NameToType = CF_TYPE_FIELDVALUE
        Case "Expression":      NameToType = CF_TYPE_EXPRESSION
        Case "FieldHasFocus":   NameToType = CF_TYPE_FIELDHASFOCUS
        Case "DataBar":         NameToType = CF_TYPE_DATABAR
        Case Else:              NameToType = CF_TYPE_EXPRESSION
    End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : OperatorToName / NameToOperator
' Purpose   : Map the AcFormatConditionOperator enum (stored in the binary blocks for
'           : field-value rules) to and from a human-readable name for the companion
'           : .json. Values verified empirically against SaveAsText output for all eight
'           : operators (docs/access-conditional-format.md section 4.3).
'---------------------------------------------------------------------------------------
'
Private Function OperatorToName(lngOperator As Long) As String
    Select Case lngOperator
        Case 0:     OperatorToName = "Between"
        Case 1:     OperatorToName = "NotBetween"
        Case 2:     OperatorToName = "Equal"
        Case 3:     OperatorToName = "NotEqual"
        Case 4:     OperatorToName = "GreaterThan"
        Case 5:     OperatorToName = "LessThan"
        Case 6:     OperatorToName = "GreaterThanOrEqual"
        Case 7:     OperatorToName = "LessThanOrEqual"
        Case Else:  OperatorToName = "Between"
    End Select
End Function

Private Function NameToOperator(varName As Variant) As Long
    Select Case LCase$(CStr(varName))
        Case "between":             NameToOperator = 0
        Case "notbetween":          NameToOperator = 1
        Case "equal":               NameToOperator = 2
        Case "notequal":            NameToOperator = 3
        Case "greaterthan":         NameToOperator = 4
        Case "lessthan":            NameToOperator = 5
        Case "greaterthanorequal":  NameToOperator = 6
        Case "lessthanorequal":     NameToOperator = 7
        Case Else:                  NameToOperator = 0
    End Select
End Function

Private Function LimitToName(bteValue As Byte) As String
    Select Case bteValue
        Case 1:     LimitToName = "number"
        Case 2:     LimitToName = "percent"
        Case Else:  LimitToName = "automatic"
    End Select
End Function

Private Function NameToLimit(varName As Variant) As Long
    Select Case LCase$(CStr(varName))
        Case "number":  NameToLimit = 1
        Case "percent": NameToLimit = 2
        Case Else:      NameToLimit = 0
    End Select
End Function

Private Function LegacyOperator(lngType As Long, dFirstRule As Dictionary) As Long
    ' The legacy header operator (offset 16) carries the first rule's operator for
    ' field-value blocks; expression/focus blocks store 0 (operator is ignored).
    If lngType = CF_TYPE_FIELDVALUE Then
        LegacyOperator = NameToOperator(dNZ(dFirstRule, "Operator", "Between"))
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Small numeric helpers
'---------------------------------------------------------------------------------------
'
Private Function Utf16Units(strText As String) As Long
    Utf16Units = Len(strText)
End Function

Private Function MaxLong(lngA As Long, lngB As Long) As Long
    MaxLong = IIf(lngA >= lngB, lngA, lngB)
End Function

Private Function CLngVal(varValue As Variant) As Long
    If IsNumeric(varValue) Then CLngVal = CLng(varValue)
End Function


'---------------------------------------------------------------------------------------
' Procedure : LongToRGB / RGBToLong
' Purpose   : Convert between Access Long color values and human-readable RGB(R,G,B)
'           : strings for the companion JSON. CF binary colors are always 0..16777215.
'---------------------------------------------------------------------------------------
'
Private Function LongToRGB(lngColor As Long) As String

    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long

    lngR = lngColor And &HFF&
    lngG = (lngColor \ &H100&) And &HFF&
    lngB = (lngColor \ &H10000) And &HFF&
    LongToRGB = "RGB(" & lngR & "," & lngG & "," & lngB & ")"

End Function

Private Function RGBToLong(varValue As Variant) As Long

    Dim strText As String
    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long
    Dim lngComma1 As Long
    Dim lngComma2 As Long

    If IsNumeric(varValue) Then
        RGBToLong = CLng(varValue)
        Exit Function
    End If

    strText = Trim$(CStr(varValue))
    If Left$(strText, 4) = "RGB(" And Right$(strText, 1) = ")" Then
        strText = Mid$(strText, 5, Len(strText) - 5)
        lngComma1 = InStr(1, strText, ",")
        lngComma2 = InStr(lngComma1 + 1, strText, ",")
        If lngComma1 > 0 And lngComma2 > lngComma1 Then
            lngR = CLng(Left$(strText, lngComma1 - 1))
            lngG = CLng(Mid$(strText, lngComma1 + 1, lngComma2 - lngComma1 - 1))
            lngB = CLng(Mid$(strText, lngComma2 + 1))
            RGBToLong = lngR + (lngG * &H100&) + (lngB * &H10000)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : dNz
' Purpose   : Return a dictionary value, or a default when the key is missing/null.
'---------------------------------------------------------------------------------------
'
Private Function dNZ(dItems As Dictionary, strKey As String, Optional varDefault As Variant = "") As Variant
    If dItems Is Nothing Then
        dNZ = varDefault
    ElseIf dItems.Exists(strKey) Then
        dNZ = dItems(strKey)
    Else
        dNZ = varDefault
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_colRules = Nothing
End Sub
