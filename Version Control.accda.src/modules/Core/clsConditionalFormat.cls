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

' Decoded rule model. Each item is a Dictionary describing one format rule.
Private m_colRules As Collection
Private m_lngVersion As Long


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
' Procedure : LoadFromCF14Hex
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse a ConditionalFormat14 hex string (concatenated, no "0x" prefix or
'           : separators) into the decoded rule model. CF14 is the authoritative source.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromCF14Hex(strHex As String)

    Dim bteData() As Byte
    Dim lngPos As Long
    Dim lngRuleCount As Long
    Dim lngFirstType As Long
    Dim lngType As Long
    Dim lngOperator As Long
    Dim lngRule As Long
    Dim dRule As Dictionary

    Set m_colRules = New Collection
    If Len(strHex) = 0 Then Exit Sub
    bteData = HexToBytes(strHex)
    If UBound(bteData) < 13 Then Exit Sub

    ' Parse 14-byte header
    m_lngVersion = ReadInt(bteData, 0)
    lngRuleCount = ReadInt(bteData, 2)
    lngFirstType = ReadInt(bteData, 6)
    lngPos = 14

    For lngRule = 0 To lngRuleCount - 1

        ' Rules after the first carry an 8-byte type prefix (type dword + operator
        ' dword). Data bar records embed their own type dword at the same position, so
        ' the prefix and the data bar header coincide; read the type either way.
        ' For the first rule the type comes from the header (offset 6) and the
        ' field-value operator from offset 10 (docs section 4.1). For later rules the
        ' operator is the second dword of the prefix (offset lngPos + 4).
        If lngRule = 0 Then
            lngType = lngFirstType
            lngOperator = ReadInt(bteData, 10)
        Else
            lngType = ReadLong(bteData, lngPos)
            lngOperator = ReadLong(bteData, lngPos + 4)
        End If

        Select Case lngType
            Case CF_TYPE_DATABAR
                Set dRule = ParseDataBarRule(bteData, lngPos)
            Case Else
                ' Skip the prefix for subsequent non-data-bar rules
                If lngRule > 0 Then lngPos = lngPos + RULE_PREFIX_LEN
                Set dRule = ParseStandardRule(bteData, lngPos, lngType, lngOperator)
        End Select

        If dRule Is Nothing Then Exit For
        m_colRules.Add dRule
    Next lngRule

End Sub


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
    Dim lngFlags As Long
    Dim lngLen1 As Long
    Dim lngLen2 As Long
    Dim lngUnits As Long
    Dim lngTrailPos As Long
    Dim lngTrailEcho As Long
    Dim lngEcho As Long

    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare

    lngFlags = ReadLong(bteData, lngPos)
    dRule.Add "Type", TypeToName(lngType)
    dRule.Add "Enabled", CBool((lngFlags And &HFF) <> 0)
    dRule.Add "FontBold", CBool(((lngFlags \ &H100) And &HFF) <> 0)
    dRule.Add "FontItalic", CBool(((lngFlags \ &H10000) And &HFF) <> 0)
    dRule.Add "FontUnderline", CBool(((lngFlags \ &H1000000) And &HFF) <> 0)
    dRule.Add "ForeColor", LongToRGB(ReadLong(bteData, lngPos + 4))
    dRule.Add "BackColor", LongToRGB(ReadLong(bteData, lngPos + 8))

    ' Expression length and text (focus rules have length 0)
    lngLen1 = ReadLong(bteData, lngPos + 12)
    lngUnits = lngLen1
    If lngType = CF_TYPE_FIELDVALUE Then
        ' Both bounds stored inline, each length-prefixed. Expression2 is only
        ' meaningful for Between / NotBetween; single-value operators leave it empty.
        dRule.Add "Operator", OperatorToName(lngOperator)
        dRule.Add "Expression1", ReadUtf16(bteData, lngPos + 16, lngLen1)
        lngLen2 = ReadLong(bteData, lngPos + 16 + (lngLen1 * 2))
        dRule.Add "Expression2", ReadUtf16(bteData, lngPos + 20 + (lngLen1 * 2), lngLen2)
        lngUnits = lngLen1 + lngLen2
    ElseIf lngType = CF_TYPE_EXPRESSION Then
        dRule.Add "Expression1", ReadUtf16(bteData, lngPos + 16, lngLen1)
    End If

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
    lngPos = lngPos + BASE_BODY + (lngUnits * 2)

    Set ParseStandardRule = dRule

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseDataBarRule
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Parse a data bar rule record (CF14 only) beginning at the type marker.
'           : Advances lngPos past the full record. Offsets follow docs §4.4.
'---------------------------------------------------------------------------------------
'
Private Function ParseDataBarRule(bteData() As Byte, ByRef lngPos As Long) As Dictionary

    Dim dRule As Dictionary
    Dim lngShortLen As Long
    Dim lngLongLen As Long
    Dim lngP As Long

    Set dRule = New Dictionary
    dRule.CompareMode = TextCompare
    dRule.Add "Type", TypeToName(CF_TYPE_DATABAR)

    dRule.Add "FillColor", LongToRGB(ReadLong(bteData, lngPos + 16))
    lngShortLen = ReadLong(bteData, lngPos + 20)
    dRule.Add "ShortestValue", ReadUtf16(bteData, lngPos + 24, lngShortLen)
    lngLongLen = ReadLong(bteData, lngPos + 24 + (lngShortLen * 2))
    dRule.Add "LongestValue", ReadUtf16(bteData, lngPos + 28 + (lngShortLen * 2), lngLongLen)

    ' Trailer position relative to the type marker
    lngP = lngPos + 32 + (2 * (lngShortLen + lngLongLen))
    dRule.Add "ShowBarOnly", CBool(bteData(lngP) <> 0)
    dRule.Add "BarColor", LongToRGB(bteData(lngP + 1) + (bteData(lngP + 2) * &H100&) + (bteData(lngP + 3) * &H10000))
    dRule.Add "ShortestLimit", LimitToName(bteData(lngP + 5))
    dRule.Add "LongestLimit", LimitToName(bteData(lngP + 9))

    ' Record length is P + 13 (relative to the type marker)
    lngPos = lngP + 13

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
        ' Type marker + reserved are part of the record
        lngLen1 = Utf16Units(dNZ(dRule, "ShortestValue"))
        lngLen2 = Utf16Units(dNZ(dRule, "LongestValue"))
        cData.Add WriteLong(CF_TYPE_DATABAR)       ' type
        cData.Add WriteLong(0)                      ' reserved
        cData.Add WriteLong(1)                      ' unk1
        cData.Add WriteLong(0)                      ' reserved
        cData.Add WriteLong(RGBToLong(dNZ(dRule, "FillColor", "0")))
        cData.Add WriteLong(lngLen1)
        cData.Add WriteUtf16(dNZ(dRule, "ShortestValue"))
        cData.Add WriteLong(lngLen2)
        cData.Add WriteUtf16(dNZ(dRule, "LongestValue"))
        cData.Add WriteLong(1)                      ' unk2
        ' Trailer
        cData.Add WriteByte(IIf(CBool(dNZ(dRule, "ShowBarOnly", "0")), 1, 0))
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

    cData.Add WriteLong(FlagsToLong(dRule))
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
'---------------------------------------------------------------------------------------
'
Public Function BuildLegacyHex() As String

    Dim cBody As clsConcat
    Dim cExpr As clsConcat
    Dim cData As clsConcat
    Dim lngIndex As Long
    Dim dRule As Dictionary
    Dim lngType As Long
    Dim lngHeaderType As Long
    Dim lngCount As Long
    Dim lngCapacity As Long
    Dim lngBlockSize As Long
    Dim lngHeaderBytes As Long
    Dim lngOffset24 As Long
    Dim lngNextStart As Long
    Dim lngExprLen As Long
    Dim colMatch As Collection

    If m_colRules.Count = 0 Then Exit Function

    ' The legacy block is single-type; use the first rule's type as the header type.
    lngHeaderType = NameToType(m_colRules(1)("Type"))

    ' Collect the rules that the legacy block can represent (matching type)
    Set colMatch = New Collection
    lngCapacity = 0
    For lngIndex = 1 To m_colRules.Count
        Set dRule = m_colRules(lngIndex)
        lngType = NameToType(dRule("Type"))
        If lngType = lngHeaderType And lngType <> CF_TYPE_DATABAR Then
            colMatch.Add dRule
            lngCapacity = MaxLong(lngCapacity, Utf16Units(dNZ(dRule, "Expression1")) + 1)
            If lngType = CF_TYPE_FIELDVALUE Then
                lngCapacity = MaxLong(lngCapacity, Utf16Units(dNZ(dRule, "Expression2")) + 1)
            End If
        End If
    Next lngIndex

    ' Cap at the Access 2007 legacy limit
    Do While colMatch.Count > LEGACY_RULE_CAP
        colMatch.Remove colMatch.Count
    Loop
    lngCount = colMatch.Count
    If lngCount = 0 Then Exit Function

    ' Compute offset 24 (end unit index of the first expression window for
    ' expression/focus rules, or the slot capacity for between rules).
    If lngHeaderType = CF_TYPE_FIELDVALUE Then
        lngOffset24 = lngCapacity
    Else
        lngOffset24 = Utf16Units(dNZ(colMatch(1), "Expression1")) + 1
    End If

    ' Build per-rule body records and expression buffer.
    Set cBody = New clsConcat
    Set cExpr = New clsConcat

    ' Track the running expression-window position for descriptor dwords.
    lngNextStart = Utf16Units(dNZ(colMatch(1), "Expression1")) + 2

    For lngIndex = 1 To lngCount
        Set dRule = colMatch(lngIndex)
        cBody.Add WriteLong(FlagsToLong(dRule))
        cBody.Add WriteLong(RGBToLong(dNZ(dRule, "ForeColor", "0")))
        cBody.Add WriteLong(RGBToLong(dNZ(dRule, "BackColor", "0")))

        If lngHeaderType = CF_TYPE_FIELDVALUE Then
            ' Between: capacity-based padded slots (verified for single rules)
            cExpr.Add WriteUtf16Padded(dNZ(dRule, "Expression1"), lngCapacity)
            cExpr.Add WriteUtf16Padded(dNZ(dRule, "Expression2"), lngCapacity)
        Else
            ' Expression/focus: tight window (text + 2 null UTF-16 units)
            cExpr.Add WriteUtf16(dNZ(dRule, "Expression1"))
            cExpr.Add WriteZeros(4)

            If lngIndex < lngCount Then
                ' Non-last rule: append descriptor dwords (28-byte record)
                lngExprLen = Utf16Units(dNZ(colMatch(lngIndex + 1), "Expression1"))
                cBody.Add WriteLong(1)
                cBody.Add WriteLong(0)
                cBody.Add WriteLong(lngNextStart)
                cBody.Add WriteLong(lngNextStart + lngExprLen + 1)
                lngNextStart = lngNextStart + lngExprLen + 2
            End If
        End If
    Next lngIndex

    ' Assemble: header (28 bytes) + per-rule body + padding to offset 96 + expressions.
    Set cData = New clsConcat
    cData.Add WriteLong(1)                              ' version
    cData.Add WriteLong(0)                              ' blockSize placeholder
    cData.Add WriteLong(lngCount)                       ' ruleCount
    cData.Add WriteLong(lngHeaderType)                  ' conditionType
    cData.Add WriteLong(LegacyOperator(lngHeaderType, colMatch(1)))  ' operator (offset 16)
    cData.Add WriteLong(0)                              ' reserved
    cData.Add WriteLong(lngOffset24)                    ' expression window end / capacity
    cData.Add cBody.GetStr
    lngHeaderBytes = Len(cData.GetStr) \ 2
    If LEGACY_EXPR_OFFSET > lngHeaderBytes Then cData.Add WriteZeros(LEGACY_EXPR_OFFSET - lngHeaderBytes)
    cData.Add cExpr.GetStr

    ' Patch in the real block size at offset 4
    lngBlockSize = Len(cData.GetStr) \ 2
    BuildLegacyHex = PatchLong(cData.GetStr, 4, lngBlockSize)

End Function


' ===================================================================================
' Byte/value helpers
' ===================================================================================

'---------------------------------------------------------------------------------------
' Procedure : HexToBytes
' Purpose   : Convert a hex string (2 chars per byte) into a zero-based byte array.
'---------------------------------------------------------------------------------------
'
Private Function HexToBytes(strHex As String) As Byte()

    Dim bteData() As Byte
    Dim lngChar As Long
    Dim lngPos As Long
    Dim strByte As String

    strByte = "&h00"
    ReDim bteData(0 To (Len(strHex) \ 2) - 1)
    For lngChar = 1 To Len(strHex) Step 2
        Mid$(strByte, 3, 2) = Mid$(strHex, lngChar, 2)
        bteData(lngPos) = CByte(CLng(strByte))
        lngPos = lngPos + 1
    Next lngChar
    HexToBytes = bteData

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadInt / ReadLong
' Purpose   : Read a little-endian unsigned 16-bit / 32-bit value from a byte array.
'---------------------------------------------------------------------------------------
'
Private Function ReadInt(bteData() As Byte, lngPos As Long) As Long
    ReadInt = bteData(lngPos) + (bteData(lngPos + 1) * &H100&)
End Function

Private Function ReadLong(bteData() As Byte, lngPos As Long) As Long
    ReadLong = bteData(lngPos) + (bteData(lngPos + 1) * &H100&) + _
        (bteData(lngPos + 2) * &H10000) + (CLng(bteData(lngPos + 3)) * &H1000000)
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
    ReDim bteStr(0 To (lngUnits * 2) - 1)
    For lngByte = 0 To (lngUnits * 2) - 1
        bteStr(lngByte) = bteData(lngPos + lngByte)
    Next lngByte
    ReadUtf16 = bteStr
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
    WriteLong = ByteHex(lngValue And &HFF&) & ByteHex((lngValue \ &H100&) And &HFF) & _
        ByteHex((lngValue \ &H10000) And &HFF) & ByteHex((lngValue \ &H1000000) And &HFF)
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
' Procedure : FlagsToLong
' Purpose   : Build the format-flags dword from the rule's font/enabled booleans.
'           : Byte 0 = enabled, 1 = bold, 2 = italic, 3 = underline (docs §6).
'---------------------------------------------------------------------------------------
'
Private Function FlagsToLong(dRule As Dictionary) As Long
    Dim lngFlags As Long
    If CBool(dNZ(dRule, "Enabled", "True")) Then lngFlags = lngFlags Or &H1&
    If CBool(dNZ(dRule, "FontBold", "False")) Then lngFlags = lngFlags Or &H100&
    If CBool(dNZ(dRule, "FontItalic", "False")) Then lngFlags = lngFlags Or &H10000
    If CBool(dNZ(dRule, "FontUnderline", "False")) Then lngFlags = lngFlags Or &H1000000
    FlagsToLong = lngFlags
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
