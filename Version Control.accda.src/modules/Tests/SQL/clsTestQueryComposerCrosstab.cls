VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestQueryComposerCrosstab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestQueryComposerCrosstab
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Design View (Operation =6) block shape tests for clsQueryComposer.
'           : The expectations here were captured from Access rather than inferred:
'           : crosstabs of each shape were saved through the query designer and
'           : re-exported with Application.SaveAsText, and the resulting blocks were
'           : fed back through LoadFromText to confirm Access accepts them. See
'           : DECISIONS.md and docs/access-query-storage.md.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

' Two row headings, an ORDER BY, and no parameters. The second row heading is what
' distinguishes a role marker from a nesting depth.
Private Const SQL_TWO_ROW_HEADINGS As String = _
    "TRANSFORM Count(t1.Colour) AS CountOfColour " & _
    "SELECT t1.Manufacturer, t1.Yr FROM t1 " & _
    "GROUP BY t1.Manufacturer, t1.Yr ORDER BY t1.Manufacturer PIVOT t1.Model"

' Parameterized, joined, and naming its column headings explicitly.
Private Const SQL_FIXED_PIVOT As String = _
    "PARAMETERS [Select Category] Text ( 255 ); " & _
    "TRANSFORM Count(t2.ID) AS Cnt SELECT t1.Category " & _
    "FROM t1 INNER JOIN t2 ON t1.ID = t2.SampleID " & _
    "WHERE (((t1.Category)=[Select Category])) " & _
    "GROUP BY t1.Category PIVOT t2.StatusID In (1, 2, 3)"

Private m_qc As clsQueryComposer


Private Sub Class_Initialize()
    Set m_qc = New clsQueryComposer
End Sub


Private Sub Class_Terminate()
    Set m_qc = Nothing
End Sub


Public Sub TestCrosstabIsDesignerCompatible()

    Dim strQdef As String

    strQdef = DesignQdef(SQL_TWO_ROW_HEADINGS)

    TestAssert m_qc.IsDesignerCompatible, "crosstab decomposes as designer compatible"
    TestAssert InStr(strQdef, "Operation =6" & vbCrLf) > 0, "emitted as Operation =6"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestOutputColumnsCarryRoleMarkers
' Purpose   : Pin the whole block. GroupLevel is a role, not a nesting depth, so both
'           : row headings are 2 rather than 2 and 3; the column heading is 1; and the
'           : aggregate carries no marker at all. Order is row headings in SELECT order,
'           : then the column heading, then the aggregate.
'---------------------------------------------------------------------------------------
'
Public Sub TestOutputColumnsCarryRoleMarkers()

    Dim strExpected As String

    strExpected = JoinLines(Array( _
        "Begin OutputColumns", _
        "    Expression =""t1.Manufacturer""", _
        "    GroupLevel =2", _
        "    Expression =""t1.Yr""", _
        "    GroupLevel =2", _
        "    Expression =""t1.Model""", _
        "    GroupLevel =1", _
        "    Alias =""CountOfColour""", _
        "    Expression =""Count(t1.Colour)""", _
        "End"))

    TestAssert InStr(DesignQdef(SQL_TWO_ROW_HEADINGS), strExpected) > 0, _
        "output columns match the shape captured from Access"

End Sub


Public Sub TestGroupsRepeatsOutputColumnLevels()

    Dim strExpected As String

    strExpected = JoinLines(Array( _
        "Begin Groups", _
        "    Expression =""t1.Manufacturer""", _
        "    GroupLevel =2", _
        "    Expression =""t1.Yr""", _
        "    GroupLevel =2", _
        "    Expression =""t1.Model""", _
        "    GroupLevel =1", _
        "End"))

    TestAssert InStr(DesignQdef(SQL_TWO_ROW_HEADINGS), strExpected) > 0, _
        "groups block repeats the same levels, including the pivot field"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestFixedPivotListOnlyInOutputColumns
' Purpose   : A named heading list belongs to the column-heading expression, but the
'           : matching Groups row groups on the bare field. Emitting the list in both
'           : places is the mistake this guards against.
'---------------------------------------------------------------------------------------
'
Public Sub TestFixedPivotListOnlyInOutputColumns()

    Dim strQdef As String

    strQdef = DesignQdef(SQL_FIXED_PIVOT)

    TestAssert InStr(Block(strQdef, "Begin OutputColumns"), _
        "Expression =""t2.StatusID In (1, 2, 3)""") > 0, _
        "output columns keep the fixed heading list"
    TestAssert InStr(Block(strQdef, "Begin Groups"), _
        "Expression =""t2.StatusID""" & vbCrLf) > 0, _
        "groups row drops the heading list"
    TestAssert InStr(Block(strQdef, "Begin Groups"), " In (") = 0, _
        "no heading list anywhere in the groups block"

End Sub


Public Sub TestCrosstabBlockOrder()

    Dim strQdef As String

    strQdef = DesignQdef(SQL_FIXED_PIVOT)

    TestAssert BlockOrder(strQdef, "Begin OutputColumns", "Begin Parameters"), _
        "parameters block follows output columns"
    TestAssert BlockOrder(strQdef, "Begin Parameters", "Begin Joins"), _
        "parameters block precedes joins"
    TestAssert BlockOrder(strQdef, "Begin Joins", "Begin Groups"), _
        "groups block comes last"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestParameterizedCrosstabWithoutHeadingListStaysDesignView
' Purpose   : This shape prompts in the Access designer, which has to run the query to
'           : discover its column headings. That is a designer constraint, not an import
'           : one -- LoadFromText accepts the qdef and never executes the query -- so the
'           : composer must not decline it. Measured, not assumed.
'---------------------------------------------------------------------------------------
'
Public Sub TestParameterizedCrosstabWithoutHeadingListStaysDesignView()

    Dim strQdef As String

    strQdef = DesignQdef(Replace$(SQL_FIXED_PIVOT, " In (1, 2, 3)", vbNullString))

    TestAssert m_qc.IsDesignerCompatible, _
        "crosstab without a fixed heading list is still designer compatible"
    TestAssert InStr(Block(strQdef, "Begin OutputColumns"), _
        "Expression =""t2.StatusID""" & vbCrLf & "    GroupLevel =1") > 0, _
        "bare pivot field still carries the column-heading marker"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPlainGroupByKeepsItsOwnLevels
' Purpose   : Regression guard for the branch added alongside the crosstab emitter: an
'           : ordinary aggregate query still reports the GroupLevel it parsed, not 2.
'---------------------------------------------------------------------------------------
'
Public Sub TestPlainGroupByKeepsItsOwnLevels()

    Dim strQdef As String

    strQdef = DesignQdef("SELECT t1.Category, Count(t1.ID) AS Total FROM t1 " & _
        "GROUP BY t1.Category")

    TestAssert InStr(strQdef, "Operation =1" & vbCrLf) > 0, "still an ordinary select"
    TestAssert InStr(Block(strQdef, "Begin Groups"), "GroupLevel =0") > 0, _
        "non-crosstab group keeps GroupLevel 0"
    TestAssert InStr(strQdef, "GroupLevel =2") = 0, _
        "no crosstab row-heading marker on a plain group by"

End Sub


'=======================================================================================
' PRIVATE: Helpers
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : DesignQdef
' Purpose   : Decompose SQL and render it as a Design View .qdef, leaving m_qc pointing at
'           : the composer used so the caller can inspect it.
'---------------------------------------------------------------------------------------
'
Private Function DesignQdef(ByVal strSql As String) As String
    Set m_qc = New clsQueryComposer
    m_qc.ContextLabel = "clsTestQueryComposerCrosstab"
    m_qc.DecomposeSQL strSql
    DesignQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary)
End Function


Private Function JoinLines(vLines As Variant) As String
    JoinLines = Join(vLines, vbCrLf) & vbCrLf
End Function


'---------------------------------------------------------------------------------------
' Procedure : Block
' Purpose   : The named block including its trailing End line. Crosstab blocks never
'           : nest, so the first End after the header is the right one.
'---------------------------------------------------------------------------------------
'
Private Function Block(ByVal strQdef As String, ByVal strHeader As String) As String

    Dim lngStart As Long
    Dim lngEnd As Long

    lngStart = InStr(1, strQdef, strHeader)
    If lngStart = 0 Then Exit Function
    lngEnd = InStr(lngStart, strQdef, vbCrLf & "End" & vbCrLf)
    If lngEnd = 0 Then Exit Function
    Block = Mid$(strQdef, lngStart, lngEnd - lngStart + Len(vbCrLf & "End" & vbCrLf))

End Function


Private Function BlockOrder(ByVal strQdef As String, ByVal strFirst As String, _
    ByVal strSecond As String) As Boolean

    Dim lngFirst As Long
    Dim lngSecond As Long

    lngFirst = InStr(1, strQdef, strFirst)
    lngSecond = InStr(1, strQdef, strSecond)
    BlockOrder = (lngFirst > 0) And (lngSecond > 0) And (lngFirst < lngSecond)

End Function
