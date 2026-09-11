VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSourceParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSourceParser
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsSourceParser content loading, output, and sanitization.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Core")
'@Tag("unit")

Private m_parser As clsSourceParser

' Two payload lines for an inline binary block. The canonicalizer skips such a block
' wholesale, so the content is arbitrary; these are kept over 60 characters so they also
' exercise the binary fast path in clsSourceParser.
Private Const BIN_LINE_1 As String = "0x0e00000000000000010000006c0000000000000000000000000000000000ff01"
Private Const BIN_LINE_2 As String = "0x0000000000000000bb4400008541000020454d46000001000000000000000002"


Private Sub Class_Initialize()
    Set m_parser = New clsSourceParser
End Sub


Private Sub Class_Terminate()
    Set m_parser = Nothing
End Sub


Public Sub TestLoadString_ReturnsOutput()
    Dim strContent As String
    strContent = "Begin Form" & vbCrLf & "  Caption =""Test""" & vbCrLf & "End"
    m_parser.LoadString strContent, edbForm
    m_parser.Sanitize ectObjectDefinition
    TestAssert Len(m_parser.GetOutput) > 0, "GetOutput returns content"
End Sub


Public Sub TestHash_Deterministic()
    Dim strContent As String
    strContent = "Some content for hashing"
    m_parser.LoadString strContent, edbModule
    Dim strHash1 As String
    strHash1 = m_parser.Hash()
    m_parser.LoadString strContent, edbModule
    TestAssert m_parser.Hash() = strHash1, "same content produces same hash"
End Sub


Public Sub TestStripClassHeader()
    Dim strContent As String
    Dim strResult As String
    strContent = "VERSION 1.0 CLASS" & vbCrLf & _
        "BEGIN" & vbCrLf & _
        "  MultiUse = -1" & vbCrLf & _
        "END" & vbCrLf & _
        "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database"
    strResult = m_parser.StripClassHeader(strContent, False)
    TestAssert InStr(strResult, "VERSION 1.0 CLASS") = 0, "class header stripped"
    TestAssert InStr(strResult, "Option Compare Database") > 0, "code content preserved"
End Sub


Public Sub TestStripClassHeader_NameOnly()
    Dim strContent As String
    Dim strResult As String
    strContent = "Attribute VB_Name = ""MyClass""" & vbCrLf & _
        "Option Compare Database" & vbCrLf & _
        "Option Explicit"
    strResult = m_parser.StripClassHeader(strContent, True)
    TestAssert InStr(strResult, "VB_Name") = 0, "VB_Name attribute stripped"
    TestAssert InStr(strResult, "Option Compare") > 0, "code preserved"
End Sub


Public Sub TestFormGeometry_SnapsTabularCells()
    Dim strOut As String
    strOut = SanitizeForm(TabularPairForm(105, 1995, 1800))
    TestAssert InStr(strOut, "Left =120") > 0, "origin snaps to 60-twip lattice"
    TestAssert InStr(strOut, "Name =""lblA""") > 0, "first control kept"
    TestAssert InStr(strOut, "Name =""lblB""") > 0, "second control kept"
    TestAssert InStr(strOut, "LayoutCached") = 0, "LayoutCached* stripped at 5.1.0"
End Sub


Public Sub TestFormGeometry_DerivesColumnSpan()
    Dim strOut As String
    strOut = SanitizeForm(ColumnSpanForm)
    TestAssert InStr(strOut, "Name =""lblSpan""") > 0, "spanning control kept"
    TestAssert InStr(strOut, "Width =1800") > 0, "span width derived from tracks"
End Sub


Public Sub TestFormGeometry_PreservesZeroAndCustomGaps()
    Dim strZero As String
    Dim strCustom As String
    strZero = SanitizeForm(TabularPairForm(60, 1260, 1200))
    TestAssert InStr(strZero, "Left =60") > 0, "zero-gap origin stays on lattice"
    TestAssert InStr(strZero, "Left =1260") > 0, "zero-gap second column stays adjacent"
    strCustom = SanitizeForm(TabularPairForm(120, 2100, 1800))
    TestAssert InStr(strCustom, "Left =2100") > 0, "custom 180-twip gap class preserved"
End Sub


Public Sub TestFormGeometry_KeepsEmptyCellAndFreeControl()
    Dim strOut As String
    strOut = SanitizeForm(EmptyAndFreeForm)
    TestAssert InStr(strOut, "Begin EmptyCell") > 0, "EmptyCell spacer survives"
    TestAssert InStr(strOut, "Name =""cellGap""") > 0, "spacer name survives"
    TestAssert InStr(strOut, "Left =123") > 0, "free-positioned Left is not snapped"
End Sub


Public Sub TestFormGeometry_NormalizesEnvelopesWithoutClipping()
    Dim strContent As String
    Dim strOut As String

    strContent = Replace(TabularPairForm(105, 1995, 1800), "Width =6000", "Width =6006", 1, 1)
    strContent = Replace(strContent, "Height =600", "Height =606", 1, 1)
    strOut = SanitizeForm(strContent)
    TestAssert InStr(strOut, "Width =6000") > 0, "form width snaps to nearest lattice point"
    TestAssert InStr(strOut, "Height =600") > 0, "section height snaps to nearest lattice point"

    strContent = Replace(TabularPairForm(105, 1995, 1800), "Width =6000", "Width =3701", 1, 1)
    strContent = Replace(strContent, "Height =600", "Height =401", 1, 1)
    strOut = SanitizeForm(strContent)
    TestAssert InStr(strOut, "Width =3840") > 0, "form grows to canonical child right edge"
    TestAssert InStr(strOut, "Height =480") > 0, "section grows to canonical child bottom edge"
End Sub


Public Sub TestFormGeometry_FailClosedOnUnderdeterminedSpan()
    Dim strOut As String
    strOut = SanitizeForm(SubdividedSpanForm)
    TestAssert InStr(strOut, "Left =100") > 0, "uncertain Left left unchanged"
    TestAssert InStr(strOut, "Width =5000") > 0, "uncertain Width left unchanged"
    TestAssert InStr(strOut, "Left =2100") > 0, "the cell that subdivides it is also left alone"
End Sub


' The logged warning is the only thing most people ever see about a skipped group, so
' its content is pinned here: which group, which tracks, why, and where to read more.
' Warnings do not fail a test because Log.ErrorCount counts only error and critical.
Public Sub TestFormGeometry_WarnsWithActionableDetail()
    Dim lngStart As Long
    Dim colLogged As Collection
    Dim dEntry As Dictionary
    Dim strMsg As String

    lngStart = Log.ErrorJournalCount
    SanitizeForm SubdividedSpanForm
    Set colLogged = Log.GetErrorJournalSince(lngStart)
    TestAssert colLogged.Count = 1, "one warning per unsolvable group"
    If colLogged.Count < 1 Then Exit Sub

    Set dEntry = colLogged(1)
    strMsg = CStr(dEntry("message"))
    TestAssert CStr(dEntry("level")) = "warning", "logged at warning level"
    TestAssert InStr(strMsg, "layout table 1, group 1") > 0, _
        "group named in prose rather than as ""1/1"""
    TestAssert InStr(strMsg, "columns 0 to 2") > 0, "failing track range reported"
    TestAssert InStr(strMsg, "display scaling") > 0, "reason given in plain language"
    TestAssert InStr(strMsg, "wiki/Form-Layout-Geometry") > 0, "links to the wiki page"
End Sub


' A span no other cell subdivides acts as one merged track, so it resolves even though
' none of the tracks it covers has an individual witness. Before this, one such cell
' cost the whole group its canonicalization.
Public Sub TestFormGeometry_DerivesMergedRangeSpan()
    Dim strOut As String
    strOut = SanitizeForm(MergedRangeForm)
    TestAssert InStr(strOut, "Left =17400") > 0, "merged-range origin snaps"
    TestAssert InStr(strOut, "Width =1380") > 0, "merged-range width derived from its own size"
    TestAssert InStr(strOut, "Left =18780") > 0, "neighbouring single track still canonicalized"
    TestAssert InStr(strOut, "Height =360") > 0, "vertical axis runs instead of being abandoned"
End Sub


' An axis with no observed position is a no-op for positions, not a failure for the
' group. Sizes on that axis are still canonical.
Public Sub TestFormGeometry_SnapsSizesWithoutPositions()
    Dim strOut As String
    strOut = SanitizeForm(NoVerticalPositionForm)
    TestAssert InStr(strOut, "Height =240") > 0, "height snaps with no Top to anchor it"
    TestAssert InStr(strOut, "Left =660") > 0, "horizontal axis unaffected"
    TestAssert InStr(strOut, "Top =") = 0, "no Top line is invented"
End Sub


Public Sub TestFormGeometry_IdempotentAndCountStable()
    Dim strOnce As String
    Dim strTwice As String
    strOnce = SanitizeForm(TabularPairForm(105, 1995, 1800))
    strTwice = SanitizeForm(strOnce)
    TestAssert strOnce = strTwice, "second sanitize is byte-identical"
    TestAssert CountToken(strOnce, "Name =") = CountToken(TabularPairForm(105, 1995, 1800), "Name ="), _
        "control names/count unchanged"
End Sub


Public Sub TestFormGeometry_PriorFormatLeavesCaches()
    Dim strOut As String
    strOut = SanitizeForm(TabularPairForm(105, 1995, 1800), EFV_5_0_0)
    TestAssert InStr(strOut, "LayoutCachedLeft") > 0, "5.0.0 keeps LayoutCached*"
    TestAssert InStr(strOut, "Left =105") > 0, "5.0.0 does not snap layout geometry"
End Sub


Public Sub TestFormGeometry_MinimalSanitizeLeavesGeometry()
    Dim strOut As String
    strOut = SanitizeForm(TabularPairForm(105, 1995, 1800), EFV_5_1_0, eslMinimal)
    TestAssert InStr(strOut, "LayoutCachedLeft") > 0, "minimal sanitize keeps LayoutCached*"
    TestAssert InStr(strOut, "Left =105") > 0, "minimal sanitize does not snap geometry"
End Sub


Public Sub TestFormGeometry_ResumesAfterInlineBinary()
    Dim strOut As String
    strOut = SanitizeForm(InlineBinaryForm)
    TestAssert InStr(strOut, "Left =120") > 0, "cells after a binary block still snap"
    TestAssert InStr(strOut, "Name =""lblA""") > 0, "first cell after payload kept"
    TestAssert InStr(strOut, "Name =""lblB""") > 0, "second cell after payload kept"
    TestAssert InStr(strOut, "LayoutCached") = 0, "LayoutCached* still stripped"
End Sub


Public Sub TestFormGeometry_PreservesInlineBinaryPayload()
    Dim strOut As String
    strOut = SanitizeForm(InlineBinaryForm)
    TestAssert InStr(strOut, BIN_LINE_1) > 0, "first payload line unaltered"
    TestAssert InStr(strOut, BIN_LINE_2) > 0, "second payload line unaltered"
    TestAssert CountToken(strOut, "0x") = 2, "no payload line added or dropped"
    TestAssert InStr(strOut, "Name =""imgLogo""") > 0, "control owning the payload kept"
End Sub


Public Sub TestFormGeometry_SnapsTabControl()
    Dim strOut As String
    strOut = SanitizeForm(TabControlForm)
    TestAssert InStr(strOut, "Width =4800") > 0, "tab width snaps to lattice"
    TestAssert InStr(strOut, "Name =""pgOne""") > 0, "tab page survives"
End Sub


' What the VBA section is replaced with is the sanitizer's business and depends on
' SplitLayoutFromVBA, so this only pins down that the boundary still terminates the
' geometry parse and survives into the output.
Public Sub TestFormGeometry_StopsAtCodeBehindForm()
    Dim strOut As String
    strOut = SanitizeForm(TabularPairForm(105, 1995, 1800) & vbCrLf & _
        "CodeBehindForm" & vbCrLf & "Option Compare Database" & vbCrLf & "Option Explicit")
    TestAssert InStr(strOut, "Left =120") > 0, "geometry before the terminator canonicalized"
    TestAssert InStr(strOut, "CodeBehindForm") > 0, "VBA boundary marker retained"
End Sub


' A report's own Right/Bottom pair is dropped because Access recalculates it from the
' printer margins, but the lines between "Begin Report" and that pair still have to run
' through every other sanitizer. The flag that marks the pair is only meaningful inside
' that window, so it is tested ahead of the chain rather than within it; testing it in
' place once cost every report in a real project its checksum and theme color lines.
Public Sub TestSanitizeReport_RemovesEdgesAndKeepsSanitizing()
    Dim strOut As String
    strOut = SanitizeReport(EdgeAndColorReport)
    TestAssert InStr(strOut, "Right =15000") = 0, "report Right removed"
    TestAssert InStr(strOut, "Bottom =8000") = 0, "report Bottom removed"
    TestAssert InStr(strOut, "Checksum =") = 0, "Checksum still removed before the edge pair"
    TestAssert InStr(strOut, "WebImagePadding") = 0, "WebImagePadding still removed before the edge pair"
    TestAssert InStr(strOut, "DatasheetGridlinesColor =") = 0, "themed color still removed before the edge pair"
End Sub


' Only the report's own edges and the color value backed by a theme index come out.
Public Sub TestSanitizeReport_KeepsRemainingProperties()
    Dim strOut As String
    strOut = SanitizeReport(EdgeAndColorReport)
    TestAssert InStr(strOut, "DatasheetGridlinesColor12 =") > 0, "property supplying the theme index retained"
    TestAssert InStr(strOut, "RecordSource =""qryTest""") > 0, "property before the edge pair retained"
    TestAssert InStr(strOut, "ItemSuffix =88") > 0, "property after the edge pair retained"
    TestAssert InStr(strOut, "Name =""lblTitle""") > 0, "control retained"
End Sub


Private Function SanitizeForm(strContent As String, _
    Optional lngFormat As Long = EFV_5_1_0, _
    Optional intSanitize As eSanitizeLevel = eslStandard) As String

    Dim lngOldFormat As Long
    Dim intOldSanitize As eSanitizeLevel
    lngOldFormat = Options.ExportFormatVersion
    intOldSanitize = Options.SanitizeLevel
    Options.ExportFormatVersion = lngFormat
    Options.SanitizeLevel = intSanitize
    m_parser.LoadString strContent, edbForm
    m_parser.Sanitize ectObjectDefinition
    SanitizeForm = m_parser.GetOutput
    Options.ExportFormatVersion = lngOldFormat
    Options.SanitizeLevel = intOldSanitize
End Function


Private Function SanitizeReport(strContent As String, _
    Optional lngFormat As Long = EFV_5_1_0, _
    Optional intSanitize As eSanitizeLevel = eslStandard, _
    Optional intColors As eSanitizeLevel = eslMinimal) As String

    Dim lngOldFormat As Long
    Dim intOldSanitize As eSanitizeLevel
    Dim intOldColors As eSanitizeLevel
    lngOldFormat = Options.ExportFormatVersion
    intOldSanitize = Options.SanitizeLevel
    intOldColors = Options.SanitizeColors
    Options.ExportFormatVersion = lngFormat
    Options.SanitizeLevel = intSanitize
    Options.SanitizeColors = intColors
    m_parser.LoadString strContent, edbReport
    m_parser.Sanitize ectObjectDefinition
    SanitizeReport = m_parser.GetOutput
    Options.ExportFormatVersion = lngOldFormat
    Options.SanitizeLevel = intOldSanitize
    Options.SanitizeColors = intOldColors
End Function


Private Function CountToken(strText As String, strToken As String) As Long
    Dim lngPos As Long
    lngPos = 1
    Do
        lngPos = InStr(lngPos, strText, strToken, vbBinaryCompare)
        If lngPos = 0 Then Exit Do
        CountToken = CountToken + 1
        lngPos = lngPos + Len(strToken)
    Loop
End Function


' Mirrors the property order Access writes for a report: the edge pair sits in the
' middle of the report's own properties, with a themed color value above it and the
' property that carries the theme index below it.
Private Function EdgeAndColorReport() As String
    EdgeAndColorReport = _
        "Version =20" & vbCrLf & _
        "Begin Report" & vbCrLf & _
        "    RecordSource =""qryTest""" & vbCrLf & _
        "    Checksum =1276036339" & vbCrLf & _
        "    DatasheetGridlinesColor =12632256" & vbCrLf & _
        "    WebImagePadding =1" & vbCrLf & _
        "    Right =15000" & vbCrLf & _
        "    Bottom =8000" & vbCrLf & _
        "    DatasheetGridlinesColor12 =12632256" & vbCrLf & _
        "    ItemSuffix =88" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin PageHeaderSection" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""PageHeaderSection""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        "                Begin Label" & vbCrLf & _
        "                    Left =120" & vbCrLf & _
        "                    Name =""lblTitle""" & vbCrLf & _
        "                End" & vbCrLf & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


Private Function TabularPairForm(lngLeftA As Long, lngLeftB As Long, lngWidth As Long) As String
    TabularPairForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =6000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblA", lngLeftA, 120, lngWidth, 360, 0, 0) & _
        CellBlock("Label", "lblB", lngLeftB, 120, lngWidth, 360, 1, 1) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


Private Function ColumnSpanForm() As String
    ColumnSpanForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =6000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =1200" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblSpan", 120, 60, 1860, 300, 0, 1) & _
        CellBlock("Label", "lblLeft", 120, 420, 840, 300, 0, 0, 1, 1) & _
        CellBlock("Label", "lblRight", 1080, 420, 840, 300, 1, 1, 1, 1) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


Private Function EmptyAndFreeForm() As String
    EmptyAndFreeForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =6000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblA", 120, 120, 1200, 300, 0, 0) & _
        CellBlock("EmptyCell", "cellGap", 1380, 120, 60, 300, 1, 1) & _
        "                Begin Label" & vbCrLf & _
        "                    Left =123" & vbCrLf & _
        "                    Top =120" & vbCrLf & _
        "                    Width =900" & vbCrLf & _
        "                    Height =300" & vbCrLf & _
        "                    Name =""lblFree""" & vbCrLf & _
        "                End" & vbCrLf & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


' Two spans that overlap with different bounds, and no single-track cell anywhere. Each
' one subdivides the other, so neither range can be treated as a merged track and no
' boundary is observed for the tracks they end on. This is the shape that stays
' underdetermined; a span nothing else subdivides is solvable (see MergedRangeForm).
Private Function SubdividedSpanForm() As String
    SubdividedSpanForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =8000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblOne", 100, 120, 5000, 300, 0, 2) & _
        CellBlock("Label", "lblTwo", 2100, 480, 5000, 300, 1, 3) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


' A cell spanning tracks 0-2 that nothing else touches, beside single-track cells in 3
' and 4. Tracks 0-2 have no individual witness, but nothing subdivides the range, so its
' total size is a free variable and the span is canonical. Real-world shape: a "Totals"
' label whose attached-label cell spans the empty leading columns of a totals row.
Private Function MergedRangeForm() As String
    MergedRangeForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =22000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblSpan", 17385, 300, 1371, 335, 0, 2) & _
        CellBlock("TextBox", "txtA", 18750, 300, 900, 335, 3, 3) & _
        CellBlock("TextBox", "txtB", 19650, 300, 1200, 335, 4, 4) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


' Continuous-form detail cells: Access omits Top and LayoutCachedTop on every one, so the
' vertical axis has no observed boundary at all. Nothing needs a Top rewrite either, and
' the heights are still canonical, so the group must not be abandoned.
Private Function NoVerticalPositionForm() As String
    NoVerticalPositionForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =8000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        RowlessCellBlock("txtA", 645, 1725, 255, 1) & _
        RowlessCellBlock("txtB", 2370, 3780, 255, 2) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
End Function


' A layout cell with a width and no vertical position of any kind. Kept separate from
' CellBlock because the absence of both Top and LayoutCachedTop is the whole point.
Private Function RowlessCellBlock(strName As String, lngLeft As Long, _
    lngWidth As Long, lngHeight As Long, lngColumn As Long) As String

    RowlessCellBlock = _
        "                Begin TextBox" & vbCrLf & _
        "                    Left =" & lngLeft & vbCrLf & _
        "                    Width =" & lngWidth & vbCrLf & _
        "                    Height =" & lngHeight & vbCrLf & _
        "                    Name =""" & strName & """" & vbCrLf & _
        "                    ColumnStart =" & lngColumn & vbCrLf & _
        "                    ColumnEnd =" & lngColumn & vbCrLf & _
        "                    RowStart =1" & vbCrLf & _
        "                    RowEnd =1" & vbCrLf & _
        "                    LayoutGroup =1" & vbCrLf & _
        "                    GroupTable =1" & vbCrLf & _
        "                End" & vbCrLf
End Function


' The payload sits ahead of the layout cells on purpose. If the parser mis-detected the
' terminator of a binary block it would resume on the wrong line, and the cells that
' follow would never be recognized as a group to snap.
Private Function InlineBinaryForm() As String

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "Version =20"
        .Add "Begin Form"
        .Add "    Width =6000"
        .Add "    Begin"
        .Add "        Begin FormHeader"
        .Add "            Height =600"
        .Add "            Name =""FormHeader"""
        .Add "            Begin"
        .Add "                Begin Image"
        .Add "                    Left =60"
        .Add "                    Top =60"
        .Add "                    Width =480"
        .Add "                    Height =480"
        .Add "                    Name =""imgLogo"""
        .Add "                    PictureData = Begin"
        .Add "                        ", BIN_LINE_1
        .Add "                        ", BIN_LINE_2
        .Add "                    End"
        .Add "                End"

        ' CellBlock brings its own line breaks
        .AppendOnAdd = vbNullString
        .Add CellBlock("Label", "lblA", 105, 120, 1800, 360, 0, 0)
        .Add CellBlock("Label", "lblB", 1995, 120, 1800, 360, 1, 1)

        .AppendOnAdd = vbCrLf
        .Add "            End"
        .Add "        End"
        .Add "    End"
        .AppendOnAdd = vbNullString
        .Add "End"
        InlineBinaryForm = .GetStr
    End With

End Function


Private Function TabControlForm() As String

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "Version =20"
        .Add "Begin Form"
        .Add "    Width =6000"
        .Add "    Begin"
        .Add "        Begin FormHeader"
        .Add "            Height =2100"
        .Add "            Name =""FormHeader"""
        .Add "            Begin"
        .Add "                Begin Tab"
        .Add "                    Left =120"
        .Add "                    Top =120"
        .Add "                    Width =4805"
        .Add "                    Height =1800"
        .Add "                    Name =""tabMain"""
        .Add "                    Begin"
        .Add "                        Begin Page"
        .Add "                            Name =""pgOne"""
        .Add "                        End"
        .Add "                    End"
        .Add "                End"
        .Add "            End"
        .Add "        End"
        .Add "    End"
        .AppendOnAdd = vbNullString
        .Add "End"
        TabControlForm = .GetStr
    End With

End Function


Private Function CellBlock(strType As String, strName As String, _
    lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long, _
    lngColStart As Long, lngColEnd As Long, _
    Optional lngRowStart As Long = 0, Optional lngRowEnd As Long = 0) As String

    CellBlock = _
        "                Begin " & strType & vbCrLf & _
        "                    Left =" & lngLeft & vbCrLf & _
        "                    Top =" & lngTop & vbCrLf & _
        "                    Width =" & lngWidth & vbCrLf & _
        "                    Height =" & lngHeight & vbCrLf & _
        "                    Name =""" & strName & """" & vbCrLf & _
        "                    GroupTable =1" & vbCrLf & _
        "                    LayoutCachedLeft =" & lngLeft & vbCrLf & _
        "                    LayoutCachedTop =" & lngTop & vbCrLf & _
        "                    LayoutCachedWidth =" & (lngLeft + lngWidth) & vbCrLf & _
        "                    LayoutCachedHeight =" & (lngTop + lngHeight) & vbCrLf & _
        IIf(lngColStart > 0, "                    ColumnStart =" & lngColStart & vbCrLf, vbNullString) & _
        IIf(lngColEnd > 0, "                    ColumnEnd =" & lngColEnd & vbCrLf, vbNullString) & _
        IIf(lngRowStart > 0, "                    RowStart =" & lngRowStart & vbCrLf, vbNullString) & _
        IIf(lngRowEnd > 0, "                    RowEnd =" & lngRowEnd & vbCrLf, vbNullString) & _
        "                    LayoutGroup =1" & vbCrLf & _
        "                    GroupTable =1" & vbCrLf & _
        "                End" & vbCrLf
End Function
