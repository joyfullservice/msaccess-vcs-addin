VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestQueryLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestQueryLayout
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Unit tests for the query design layout pipeline, which runs:
'           :
'           :   LvExtra blob -> clsLvExtraParser -> layout Dictionary
'           :   layout Dictionary -> clsQueryComposer.GenerateQdef -> .qdef text
'           :
'           : The grid origin is the one field in this pipeline whose byte order
'           : and keyword order disagree: the blob stores it Top-then-Left while
'           : the qdef must emit it Left-then-Top. A regression that reversed the
'           : reader was invisible to every existing unit test, because nothing
'           : below the round-trip harness exercised the binary blob.
'           :
'           : These tests are database-free, so they run in the normal suite and
'           : fail immediately if either end of that asymmetry is flipped.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

' Grid origin probe values. Deliberately asymmetric, and deliberately distinct
' from every other coordinate in the synthetic blob, so a swap cannot hide
' behind a matching value somewhere else in the layout.
Private Const GRID_TOP As Long = 720
Private Const GRID_LEFT As Long = 1440

' Byte offsets within the synthetic LvExtra blob (see clsLvExtraParser header).
Private Const OFS_WINDOW As Long = 16
Private Const OFS_STATE As Long = 32
Private Const OFS_PANE As Long = 36
Private Const OFS_GRID As Long = 52
Private Const OFS_COLUMNS_SHOWN As Long = 60
Private Const OFS_TABLE_COUNT As Long = 64
Private Const OFS_TABLE_ENTRY As Long = 68
Private Const BLOB_SIZE As Long = 352

Private m_parser As clsLvExtraParser
Private m_qc As clsQueryComposer


Private Sub Class_Initialize()
    Set m_parser = New clsLvExtraParser
End Sub


Private Sub Class_Terminate()
    Set m_parser = Nothing
    Set m_qc = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Reader half: the blob stores the grid origin Top first, then Left. Every other
' RECT in the format is Left first, so this is the trap. The surrounding fields
' are asserted too -- if the parser's cursor drifts by a single Long, the grid
' values could still coincidentally match while ColumnsShown and the table entry
' silently shift.
'---------------------------------------------------------------------------------------
'
Public Sub TestLvExtraGridOriginIsTopThenLeft()

    Dim bteBlob() As Byte
    Dim dLayout As Dictionary
    Dim dTable As Dictionary

    bteBlob = BuildSyntheticBlob()
    Set dLayout = m_parser.ParseLvExtra(bteBlob)

    TestAssert dLayout.Exists("gridTop"), "parser produced a gridTop"
    TestAssert dLayout.Exists("gridLeft"), "parser produced a gridLeft"
    TestAssert dLayout("gridTop") = GRID_TOP, "first grid Long in blob reads as gridTop"
    TestAssert dLayout("gridLeft") = GRID_LEFT, "second grid Long in blob reads as gridLeft"

    ' Cursor integrity around the grid fields.
    TestAssert dLayout("window")("left") = 100, "window left"
    TestAssert dLayout("window")("top") = 200, "window top"
    TestAssert dLayout("window")("right") = 800, "window right"
    TestAssert dLayout("window")("bottom") = 600, "window bottom"
    TestAssert dLayout("state") = 0, "state"
    TestAssert dLayout("designerPane")("left") = 10, "pane left"
    TestAssert dLayout("designerPane")("top") = 20, "pane top"
    TestAssert dLayout("columnsShown") = 539, "columnsShown follows the grid origin"
    TestAssert dLayout("tables").Count = 1, "one table entry parsed"

    Set dTable = dLayout("tables")(1)
    TestAssert dTable("name") = "tblCars", "table name survives the grid offset"
    TestAssert dTable("left") = 60, "table left"
    TestAssert dTable("top") = 15, "table top"

End Sub


'---------------------------------------------------------------------------------------
' Writer half: LoadFromText parses the layout block positionally, so the grid
' origin must be emitted Left then Top. Emitting Top first fails the entire
' Design View import with "Expected: 'Left'", which silently degrades the query
' to SQL View and drops all layout metadata.
'---------------------------------------------------------------------------------------
'
Public Sub TestQdefGridOriginIsLeftThenTop()

    Dim strQdef As String
    Dim astrLines() As String
    Dim lngGrid As Long

    strQdef = GenerateLayoutQdef(BuildLayoutDict())
    astrLines = Split(strQdef, vbCrLf)

    ' The grid origin is the pair immediately above ColumnsShown. Anchoring on
    ' ColumnsShown avoids matching the window, pane, or table Left/Top pairs.
    lngGrid = LineIndex(astrLines, "    ColumnsShown =539") - 2
    TestAssert lngGrid > 0, "found the ColumnsShown anchor in the layout block"
    If lngGrid < 1 Then Exit Sub

    TestAssert astrLines(lngGrid) = "    Left =" & GRID_LEFT, "grid Left emitted first"
    TestAssert astrLines(lngGrid + 1) = "    Top =" & GRID_TOP, "grid Top emitted second"

End Sub


'---------------------------------------------------------------------------------------
' Both halves composed. This is the invariant that actually broke: a blob whose
' grid origin is (Top=720, Left=1440) must emit a qdef reading Left =1440 then
' Top =720. Note this test alone would still pass if both halves were flipped
' together -- the two tests above are what pin each end to ground truth.
'---------------------------------------------------------------------------------------
'
Public Sub TestGridOriginSurvivesBlobToQdef()

    Dim bteBlob() As Byte
    Dim dLayout As Dictionary
    Dim strQdef As String
    Dim astrLines() As String
    Dim lngGrid As Long

    bteBlob = BuildSyntheticBlob()
    Set dLayout = m_parser.ParseLvExtra(bteBlob)
    strQdef = GenerateLayoutQdef(dLayout)
    astrLines = Split(strQdef, vbCrLf)

    lngGrid = LineIndex(astrLines, "    ColumnsShown =539") - 2
    TestAssert lngGrid > 0, "found the ColumnsShown anchor after blob parse"
    If lngGrid < 1 Then Exit Sub

    TestAssert astrLines(lngGrid) = "    Left =" & GRID_LEFT, _
        "blob grid Left reaches the qdef Left line"
    TestAssert astrLines(lngGrid + 1) = "    Top =" & GRID_TOP, _
        "blob grid Top reaches the qdef Top line"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GenerateLayoutQdef
' Purpose   : Run a minimal designer-compatible query through GenerateQdef with
'           : the supplied layout, returning the .qdef text.
'---------------------------------------------------------------------------------------
'
Private Function GenerateLayoutQdef(dLayout As Dictionary) As String
    Set m_qc = New clsQueryComposer
    m_qc.DecomposeSQL "SELECT tblCars.Manufacturer FROM tblCars"
    GenerateLayoutQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, dLayout)
    TestAssert m_qc.WarningCount = 0, "layout qdef generated without warnings"
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildLayoutDict
' Purpose   : Hand-built layout Dictionary in the shape clsLvExtraParser returns
'           : and EmitDesignLayout consumes.
'---------------------------------------------------------------------------------------
'
Private Function BuildLayoutDict() As Dictionary

    Dim dLayout As New Dictionary
    Dim colTables As New Collection

    dLayout.Add "state", 0
    dLayout.Add "window", RectDict(100, 200, 800, 600)
    dLayout.Add "designerPane", RectDict(10, 20, 780, 300)
    dLayout.Add "gridLeft", GRID_LEFT
    dLayout.Add "gridTop", GRID_TOP
    dLayout.Add "columnsShown", 539

    Dim dTable As Dictionary
    Set dTable = RectDict(60, 15, 240, 195)
    dTable.Add "scrollTop", 0
    dTable.Add "name", "tblCars"
    dTable.Add "alias", vbNullString
    colTables.Add dTable
    dLayout.Add "tables", colTables

    Set BuildLayoutDict = dLayout

End Function


Private Function RectDict(ByVal lngLeft As Long, ByVal lngTop As Long, _
                          ByVal lngRight As Long, ByVal lngBottom As Long) As Dictionary
    Dim dRect As New Dictionary
    dRect.Add "left", lngLeft
    dRect.Add "top", lngTop
    dRect.Add "right", lngRight
    dRect.Add "bottom", lngBottom
    Set RectDict = dRect
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildSyntheticBlob
' Purpose   : Construct a minimal but structurally valid LvExtra blob: header,
'           : window RECT, state, designer pane RECT, grid origin (Top then
'           : Left, as Access stores it), ColumnsShown, and a single table entry.
'---------------------------------------------------------------------------------------
'
Private Function BuildSyntheticBlob() As Byte()

    Dim bteBlob() As Byte
    Dim lngPos As Long

    ReDim bteBlob(0 To BLOB_SIZE - 1)

    ' Magic header, then 12 bytes of 0xAA padding.
    bteBlob(0) = &H99
    bteBlob(1) = &H99
    bteBlob(2) = &HCE
    bteBlob(3) = &HAC
    For lngPos = 4 To 15
        bteBlob(lngPos) = &HAA
    Next lngPos

    PutRect bteBlob, OFS_WINDOW, 100, 200, 800, 600
    PutLong bteBlob, OFS_STATE, 0
    PutRect bteBlob, OFS_PANE, 10, 20, 780, 300

    ' Grid origin, in the blob's own reversed order.
    PutLong bteBlob, OFS_GRID, GRID_TOP
    PutLong bteBlob, OFS_GRID + 4, GRID_LEFT

    PutLong bteBlob, OFS_COLUMNS_SHOWN, 539
    PutLong bteBlob, OFS_TABLE_COUNT, 1

    ' Single table entry: 5 Longs, then two fixed 132-byte UTF-16LE fields.
    PutRect bteBlob, OFS_TABLE_ENTRY, 60, 15, 240, 195
    PutLong bteBlob, OFS_TABLE_ENTRY + 16, 0
    PutUtf16 bteBlob, OFS_TABLE_ENTRY + 20, "tblCars"

    BuildSyntheticBlob = bteBlob

End Function


Private Sub PutRect(bteData() As Byte, ByVal lngPos As Long, ByVal lngLeft As Long, _
                    ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long)
    PutLong bteData, lngPos, lngLeft
    PutLong bteData, lngPos + 4, lngTop
    PutLong bteData, lngPos + 8, lngRight
    PutLong bteData, lngPos + 12, lngBottom
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PutLong
' Purpose   : Write a 4-byte little-endian signed Long. Masking before dividing
'           : keeps negative values correct, since VBA integer division rounds
'           : toward zero rather than flooring.
'---------------------------------------------------------------------------------------
'
Private Sub PutLong(bteData() As Byte, ByVal lngPos As Long, ByVal lngValue As Long)
    bteData(lngPos) = lngValue And &HFF&
    bteData(lngPos + 1) = (lngValue And &HFF00&) \ &H100&
    bteData(lngPos + 2) = (lngValue And &HFF0000) \ &H10000
    bteData(lngPos + 3) = ((lngValue And &HFF000000) \ &H1000000) And &HFF&
End Sub


Private Sub PutUtf16(bteData() As Byte, ByVal lngPos As Long, ByVal strText As String)
    Dim lngChar As Long
    Dim lngCode As Long
    For lngChar = 1 To Len(strText)
        lngCode = AscW(Mid$(strText, lngChar, 1)) And &HFFFF&
        bteData(lngPos + (lngChar - 1) * 2) = lngCode And &HFF&
        bteData(lngPos + (lngChar - 1) * 2 + 1) = (lngCode And &HFF00&) \ &H100&
    Next lngChar
End Sub


Private Function LineIndex(astrLines() As String, ByVal strMatch As String) As Long
    Dim lngLine As Long
    LineIndex = -1
    For lngLine = LBound(astrLines) To UBound(astrLines)
        If astrLines(lngLine) = strMatch Then
            LineIndex = lngLine
            Exit Function
        End If
    Next lngLine
End Function
