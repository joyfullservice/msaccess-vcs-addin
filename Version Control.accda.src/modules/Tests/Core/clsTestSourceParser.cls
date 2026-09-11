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
    strOut = SanitizeForm(UnderdeterminedForm)
    TestAssert InStr(strOut, "Left =100") > 0, "uncertain Left left unchanged"
    TestAssert InStr(strOut, "Width =5000") > 0, "uncertain Width left unchanged"
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


Private Function UnderdeterminedForm() As String
    UnderdeterminedForm = _
        "Version =20" & vbCrLf & _
        "Begin Form" & vbCrLf & _
        "    Width =8000" & vbCrLf & _
        "    Begin" & vbCrLf & _
        "        Begin FormHeader" & vbCrLf & _
        "            Height =600" & vbCrLf & _
        "            Name =""FormHeader""" & vbCrLf & _
        "            Begin" & vbCrLf & _
        CellBlock("Label", "lblWide", 100, 120, 5000, 300, 0, 20) & _
        "            End" & vbCrLf & _
        "        End" & vbCrLf & _
        "    End" & vbCrLf & _
        "End"
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
