VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestQueryComposerJoins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestQueryComposerJoins
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Join-graph invariant tests for clsQueryComposer import/export paths.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

Private m_qc As clsQueryComposer


Private Sub Class_Initialize()
    Set m_qc = New clsQueryComposer
End Sub


Private Sub Class_Terminate()
    Set m_qc = Nothing
End Sub


Public Sub TestSpacedTableNameJoin_ImportPath()
    Dim strSql As String
    strSql = "SELECT a.fld FROM [Car Models] " & _
             "INNER JOIN [Car Colours] ON [Car Models].fld = [Car Colours].fld"
    AssertJoinRefsResolve strSql, "spaced table names in join"
End Sub


Public Sub TestFunctionOperandInMultiCondOn_JoinRefsInInputTables()
    ' Pins: DateAdd (or any function) on one side of a split ON condition must
    ' never become LeftTable/RightTable. Refs must resolve to InputTables aliases.
    Dim strSql As String
    Dim colJoins As Collection

    strSql = "SELECT cur.OrderID, cur.OrderDate FROM tblOrders AS cur " & _
             "LEFT JOIN tblOrders AS prior ON (cur.CustomerID = prior.CustomerID) " & _
             "AND (prior.OrderDate = DateAdd('yyyy', -1, cur.OrderDate))"

    AssertJoinRefsResolve strSql, "function operand multi-cond ON"

    Set colJoins = QdefJoinRows(strSql, True)
    TestAssert colJoins.Count = 2, "two split join rows for compound ON"
    AssertAllJoinRowsArePair colJoins, "cur", "prior", "function operand"
    AssertNoJoinRefContains colJoins, "DateAdd", "function operand"
End Sub


Public Sub TestSingleTablePredicateInCompoundOn_KeepsParentPair()
    ' Pins: a condition naming only one table (tblB.flag > 0) must keep the
    ' parent join's pair. Resolving it from the condition alone would emit
    ' LeftTable = RightTable, which Access collapses and which breaks
    ' BuildJoinChain on the next export.
    Dim strSql As String
    Dim colJoins As Collection

    strSql = "SELECT t1.a FROM t1 INNER JOIN t2 " & _
             "ON (t1.x = t2.x) AND (t2.flag > 0)"

    AssertJoinRefsResolve strSql, "single-table predicate in compound ON"

    Set colJoins = QdefJoinRows(strSql, True)
    TestAssert colJoins.Count = 2, "two split join rows for predicate compound ON"
    AssertAllJoinRowsArePair colJoins, "t1", "t2", "single-table predicate"
End Sub


Public Sub TestNonEquiConditionInCompoundOn_UsesBothTables()
    ' Pins: a non-equi condition has no "=" for ExtractTableFromOnSide to split
    ' on, so the pair must come from the parent join rather than collapsing.
    Dim strSql As String
    Dim colJoins As Collection

    strSql = "SELECT A.k FROM tblReadings AS A INNER JOIN tblReadings AS B " & _
             "ON (A.k = B.k) AND (A.v > B.v)"

    AssertJoinRefsResolve strSql, "non-equi condition in compound ON"

    Set colJoins = QdefJoinRows(strSql, True)
    TestAssert colJoins.Count = 2, "two split join rows for non-equi compound ON"
    AssertAllJoinRowsArePair colJoins, "A", "B", "non-equi condition"
End Sub


Public Sub TestJoinFlavorMatrix_Select()
    Dim vCases As Variant
    Dim i As Long
    vCases = Array( _
        "SELECT a FROM t1 INNER JOIN t2 ON t1.id = t2.id", _
        "SELECT a FROM t1 LEFT JOIN t2 ON t1.id = t2.id", _
        "SELECT a FROM t2 RIGHT JOIN t1 ON t2.id = t1.id", _
        "SELECT a FROM (t1 INNER JOIN t2 ON t1.id = t2.id) INNER JOIN t3 ON t2.id = t3.id", _
        "SELECT a FROM t1 INNER JOIN (t2 INNER JOIN t3 ON t2.id = t3.id) ON t1.id = t2.id", _
        "SELECT a FROM t1 INNER JOIN t2 ON (t1.x = t2.x) AND (t1.y = t2.y)", _
        "SELECT a FROM t1, t2 INNER JOIN t3 ON t2.id = t3.id", _
        "SELECT a FROM [Order Details] od1 INNER JOIN [Order Details] od2 ON od1.id = od2.id", _
        "SELECT a FROM ((t1 INNER JOIN t2 ON t1.id = t2.id) LEFT JOIN t3 ON t2.id = t3.id) " & _
            "RIGHT JOIN t4 ON t3.id = t4.id", _
        "SELECT a FROM (t1 LEFT JOIN t2 ON t1.id = t2.id) INNER JOIN t3 ON (t1.x = t3.x) AND (t2.y = t3.y)", _
        "SELECT a FROM t1 AS x INNER JOIN t1 AS y ON x.id = y.id", _
        "SELECT a FROM t1 INNER JOIN t2 AS y ON t1.id = y.id", _
        "SELECT a FROM t1 AS x INNER JOIN t2 ON x.id = t2.id", _
        "SELECT a FROM (t1 INNER JOIN t2 ON t1.id = t2.id), t3 INNER JOIN t4 ON t3.id = t4.id", _
        "SELECT a FROM [Sales.2024] INNER JOIN t2 ON [Sales.2024].id = t2.id", _
        "SELECT a FROM [Table AS Thing] INNER JOIN t2 ON [Table AS Thing].id = t2.id" _
    )
    For i = LBound(vCases) To UBound(vCases)
        AssertJoinRefsResolve CStr(vCases(i)), "select join case " & CStr(i)
    Next i
End Sub


Public Sub TestSpacedNamePositions()
    Dim vCases As Variant
    Dim i As Long
    vCases = Array( _
        "SELECT a FROM [Car Models] INNER JOIN t2 ON [Car Models].id = t2.id", _
        "SELECT a FROM t1 INNER JOIN [Car Colours] ON t1.id = [Car Colours].id", _
        "SELECT a FROM [Car Models] INNER JOIN [Car Colours] " & _
            "ON [Car Models].id = [Car Colours].id", _
        "SELECT a FROM [Car Models] AS cm INNER JOIN [Car Colours] AS cc ON cm.id = cc.id", _
        "SELECT a FROM [Car Models] AS cm1 INNER JOIN [Car Models] AS cm2 ON cm1.id = cm2.id", _
        "SELECT a FROM ([Car Models] INNER JOIN [Car Colours] " & _
            "ON [Car Models].id = [Car Colours].id) " & _
            "LEFT JOIN [Order Details] ON [Car Colours].id = [Order Details].id", _
        "SELECT a FROM [Car Models] INNER JOIN [Car Colours] " & _
            "ON ([Car Models].x = [Car Colours].x) " & _
            "AND ([Car Models].y = [Car Colours].y)", _
        "SELECT a FROM [Car Models], t2 INNER JOIN [Car Colours] " & _
            "ON t2.id = [Car Colours].id" _
    )
    For i = LBound(vCases) To UBound(vCases)
        AssertJoinRefsResolve CStr(vCases(i)), "spaced name position case " & CStr(i)
    Next i
End Sub


Public Sub TestStatementTypes_SpacedNames()
    AssertJoinRefsResolve "UPDATE [Car Models] SET x = 1 FROM [Car Models] " & _
        "INNER JOIN [Car Colours] ON [Car Models].fld = [Car Colours].fld", "update spaced"
    AssertJoinRefsResolve "DELETE [Car Models].fld FROM [Car Models] " & _
        "INNER JOIN [Car Colours] ON [Car Models].fld = [Car Colours].fld", "delete spaced"
    AssertJoinRefsResolve "INSERT INTO tgt SELECT a FROM [Car Models] " & _
        "INNER JOIN [Car Colours] ON [Car Models].fld = [Car Colours].fld", "insert select spaced"
    AssertJoinRefsResolve "SELECT a INTO tgt FROM [Car Models] " & _
        "INNER JOIN [Car Colours] ON [Car Models].fld = [Car Colours].fld", "select into spaced"
End Sub


Public Sub TestReconstructSQL_SpacedJoinOperands()
    Dim colRows As New Collection
    Dim dRow As Dictionary
    Dim strSql As String

    Set dRow = New Dictionary
    dRow.Add "Attribute", 1: dRow.Add "Flag", 1
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 5: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "Car Models": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 5: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "Car Colours": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 7: dRow.Add "Flag", 1
    dRow.Add "Expression", "[Car Models].fld=[Car Colours].fld"
    dRow.Add "Name1", "Car Models": dRow.Add "Name2", "Car Colours"
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 6: dRow.Add "Flag", 0
    dRow.Add "Expression", "a": dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 255: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    m_qc.ContextLabel = "TestReconstructSQL_SpacedJoinOperands"
    strSql = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(strSql, "[Car Models]") > 0, "export path preserves spaced left table"
    TestAssert InStr(strSql, "[Car Colours]") > 0, "export path preserves spaced right table"
    TestAssert m_qc.WarningCount = 0, "export path has no join-ref warnings"
End Sub


Public Sub TestReconstructSQL_CompoundOnOuterParens()
    ' Pins: 5.2.0 wraps every compound ON in outer parentheses, whether Access
    ' stored it as one row per condition or as one compound row, because SQL
    ' View rejects `ON (a = b) AND (c > 0)` (error 3296). 5.1.0 is unchanged.
    Dim lngSaved As Long
    Dim strMerged51 As String
    Dim strMerged52 As String
    Dim strSingle51 As String
    Dim strSingle52 As String
    Dim strSimple52 As String

    lngSaved = Options.ExportFormatVersion
    Options.ExportFormatVersion = EFV_5_1_0
    strMerged51 = NoSpace(m_qc.ReconstructSQL(TwoTableJoinRows("t1.x=t2.x", "t2.flag>0")))
    strSingle51 = NoSpace(m_qc.ReconstructSQL(TwoTableJoinRows("(t1.x=t2.x) AND (t2.flag>0)")))
    Options.ExportFormatVersion = EFV_5_2_0
    strMerged52 = NoSpace(m_qc.ReconstructSQL(TwoTableJoinRows("t1.x=t2.x", "t2.flag>0")))
    strSingle52 = NoSpace(m_qc.ReconstructSQL(TwoTableJoinRows("(t1.x=t2.x) AND (t2.flag>0)")))
    strSimple52 = NoSpace(m_qc.ReconstructSQL(TwoTableJoinRows("t1.x=t2.x")))
    Options.ExportFormatVersion = lngSaved

    TestAssert InStr(1, strMerged51, "ON(t1.x=t2.x)AND(t2.flag>0);", vbTextCompare) > 0, _
        "5.1.0 leaves a merged compound ON bare: " & strMerged51
    TestAssert InStr(1, strSingle51, "ON(t1.x=t2.x)AND(t2.flag>0);", vbTextCompare) > 0, _
        "5.1.0 leaves a single-row compound ON bare: " & strSingle51
    TestAssert InStr(1, strMerged52, "ON((t1.x=t2.x)AND(t2.flag>0));", vbTextCompare) > 0, _
        "5.2.0 wraps a merged compound ON: " & strMerged52
    TestAssert InStr(1, strSingle52, "ON((t1.x=t2.x)AND(t2.flag>0));", vbTextCompare) > 0, _
        "5.2.0 wraps a single-row compound ON: " & strSingle52
    TestAssert InStr(1, strSimple52, "ONt1.x=t2.x;", vbTextCompare) > 0, _
        "5.2.0 leaves a single condition bare: " & strSimple52
End Sub


Public Sub TestReconstructSQL_CrossPairOnKeepsSeparateTerms()
    ' Pins: SQL View needs each top-level AND term of an ON to name one table
    ' from each side. A condition on another table pair stays its own term at
    ' 5.2.0; only conditions on the same pair are enclosed together.
    Dim lngSaved As Long
    Dim strCross51 As String
    Dim strCross52 As String
    Dim strMixed51 As String
    Dim strMixed52 As String

    lngSaved = Options.ExportFormatVersion
    Options.ExportFormatVersion = EFV_5_1_0
    strCross51 = NoSpace(m_qc.ReconstructSQL(ThreeTableJoinRows(False)))
    strMixed51 = NoSpace(m_qc.ReconstructSQL(ThreeTableJoinRows(True)))
    Options.ExportFormatVersion = EFV_5_2_0
    strCross52 = NoSpace(m_qc.ReconstructSQL(ThreeTableJoinRows(False)))
    strMixed52 = NoSpace(m_qc.ReconstructSQL(ThreeTableJoinRows(True)))
    Options.ExportFormatVersion = lngSaved

    TestAssert InStr(1, strCross51, "ON(t2.b=t3.b)AND(t1.c=t3.c);", vbTextCompare) > 0, _
        "5.1.0 cross-pair ON: " & strCross51
    TestAssert InStr(1, strCross52, "ON(t2.b=t3.b)AND(t1.c=t3.c);", vbTextCompare) > 0, _
        "5.2.0 keeps cross-pair conditions as separate terms: " & strCross52
    TestAssert InStr(1, strMixed51, "ON(t2.b=t3.b)AND(t3.flag>0)AND(t1.c=t3.c);", vbTextCompare) > 0, _
        "5.1.0 mixed ON: " & strMixed51
    TestAssert InStr(1, strMixed52, "ON((t2.b=t3.b)AND(t3.flag>0))AND(t1.c=t3.c);", vbTextCompare) > 0, _
        "5.2.0 encloses the same-pair conditions and keeps the other pair separate: " & strMixed52
End Sub


Public Sub TestNestedAndGroupInOn_SplitsIntoConditionRows()
    ' Pins: import descends into an enclosed same-pair group, so the 5.2.0
    ' spelling stores one Attribute 7 row per condition, like Access does.
    Dim colJoins As Collection

    Set colJoins = QdefJoinRows("SELECT t1.a FROM (t1 INNER JOIN t2 ON t1.a = t2.a) " & _
        "INNER JOIN t3 ON ((t2.b = t3.b) AND (t3.flag > 0)) AND (t1.c = t3.c)", True)
    TestAssert colJoins.Count = 4, "one join row per condition, got " & colJoins.Count
End Sub


Private Function ThreeTableJoinRows(ByVal blnSamePairPredicate As Boolean) As Collection
    Dim colRows As New Collection

    colRows.Add QueryRow(1, 1)
    colRows.Add QueryRow(5, 0, , "t1")
    colRows.Add QueryRow(5, 0, , "t2")
    colRows.Add QueryRow(5, 0, , "t3")
    colRows.Add QueryRow(7, 1, "t1.a=t2.a", "t1", "t2")
    colRows.Add QueryRow(7, 1, "t2.b=t3.b", "t2", "t3")
    colRows.Add QueryRow(7, 1, "t1.c=t3.c", "t1", "t3")
    If blnSamePairPredicate Then colRows.Add QueryRow(7, 1, "t3.flag>0", "t2", "t3")
    colRows.Add QueryRow(6, 0, "t1.a")
    colRows.Add QueryRow(255, 0)
    Set ThreeTableJoinRows = colRows
End Function


Private Function TwoTableJoinRows(ParamArray avarConditions() As Variant) As Collection
    Dim colRows As New Collection
    Dim varCondition As Variant

    colRows.Add QueryRow(1, 1)
    colRows.Add QueryRow(5, 0, , "t1")
    colRows.Add QueryRow(5, 0, , "t2")
    For Each varCondition In avarConditions
        colRows.Add QueryRow(7, 1, CStr(varCondition), "t1", "t2")
    Next varCondition
    colRows.Add QueryRow(6, 0, "t1.a")
    colRows.Add QueryRow(255, 0)
    Set TwoTableJoinRows = colRows
End Function


Private Function QueryRow(ByVal lngAttribute As Long, ByVal lngFlag As Long, _
    Optional ByVal strExpression As String, Optional ByVal strName1 As String, _
    Optional ByVal strName2 As String) As Dictionary

    Set QueryRow = New Dictionary
    QueryRow.Add "Attribute", lngAttribute
    QueryRow.Add "Flag", lngFlag
    QueryRow.Add "Expression", strExpression
    QueryRow.Add "Name1", strName1
    QueryRow.Add "Name2", strName2
End Function


Private Function NoSpace(ByVal strText As String) As String
    NoSpace = Replace(Replace(Replace(strText, vbCrLf, vbNullString), vbTab, vbNullString), " ", vbNullString)
End Function


Public Sub TestClosedLoop_SpacedTableNameJoin()
    Dim colRows As Collection
    Dim strSql1 As String
    Dim strSql2 As String
    Dim strQdef As String

    Set colRows = BuildSpacedJoinRows()
    m_qc.ContextLabel = "TestClosedLoop_SpacedTableNameJoin"
    strSql1 = m_qc.ReconstructSQL(colRows)
    TestAssert m_qc.WarningCount = 0, "closed loop reconstruct warnings"

    Set m_qc = New clsQueryComposer
    m_qc.DecomposeSQL strSql1
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary)
    TestAssert m_qc.WarningCount = 0, "closed loop qdef warnings"

    Set m_qc = New clsQueryComposer
    Set colRows = RebuildRowsFromQdef(colRows, strQdef)
    strSql2 = m_qc.ReconstructSQL(colRows)
    TestAssert strSql1 = strSql2, "closed loop SQL stable"
End Sub


Private Function QdefJoinRows(ByVal strSql As String, _
    ByVal blnRequiresDesignView As Boolean) As Collection
    Dim strQdef As String
    Set m_qc = New clsQueryComposer
    m_qc.DecomposeSQL strSql, False
    If blnRequiresDesignView Then
        TestAssert m_qc.RequiresDesignView, "multi-condition ON requires Design View"
    End If
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary)
    TestAssert m_qc.WarningCount = 0, "qdef generated without warnings"
    Set QdefJoinRows = ParseJoinTables(strQdef)
End Function


Private Sub AssertAllJoinRowsArePair(colJoins As Collection, _
    ByVal strExpLeft As String, ByVal strExpRight As String, _
    ByVal strContext As String)
    Dim varRow As Variant
    Dim dJoin As Dictionary
    For Each varRow In colJoins
        Set dJoin = varRow
        TestAssert StrComp(CStr(dJoin("left")), strExpLeft, vbTextCompare) = 0, _
            strContext & " LeftTable is " & strExpLeft & ", got " & CStr(dJoin("left"))
        TestAssert StrComp(CStr(dJoin("right")), strExpRight, vbTextCompare) = 0, _
            strContext & " RightTable is " & strExpRight & ", got " & CStr(dJoin("right"))
        TestAssert StrComp(CStr(dJoin("left")), CStr(dJoin("right")), vbTextCompare) <> 0, _
            strContext & " LeftTable and RightTable must differ"
    Next varRow
End Sub


Private Sub AssertNoJoinRefContains(colJoins As Collection, _
    ByVal strNeedle As String, ByVal strContext As String)
    Dim varRow As Variant
    Dim dJoin As Dictionary
    For Each varRow In colJoins
        Set dJoin = varRow
        TestAssert InStr(1, CStr(dJoin("left")), strNeedle, vbTextCompare) = 0, _
            strContext & " LeftTable must not contain " & strNeedle
        TestAssert InStr(1, CStr(dJoin("right")), strNeedle, vbTextCompare) = 0, _
            strContext & " RightTable must not contain " & strNeedle
    Next varRow
End Sub


Private Sub AssertJoinRefsResolve(ByVal strSql As String, ByVal strContext As String)
    Dim strQdef As String
    Set m_qc = New clsQueryComposer
    m_qc.DecomposeSQL strSql
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary)
    TestAssert m_qc.WarningCount = 0, strContext & " warnings"
    AssertQdefJoinRefsInInputTables strQdef, strContext
End Sub


Private Sub AssertQdefJoinRefsInInputTables(ByVal strQdef As String, ByVal strContext As String)
    Dim dInput As Dictionary
    Dim dJoin As Dictionary
    Dim colJoins As Collection
    Set dInput = ParseInputTableNames(strQdef)
    Set colJoins = ParseJoinTables(strQdef)
    Dim dJn As Variant
    For Each dJn In colJoins
        Set dJoin = dJn
        TestAssert dInput.Exists(CStr(dJoin("left"))), strContext & " left ref " & CStr(dJoin("left"))
        TestAssert dInput.Exists(CStr(dJoin("right"))), strContext & " right ref " & CStr(dJoin("right"))
    Next dJn
End Sub


Private Function ParseInputTableNames(ByVal strQdef As String) As Dictionary
    Dim dNames As New Dictionary
    Dim vLines() As String
    Dim lngI As Long
    Dim strLine As String
    Dim strName As String
    Dim blnInBlock As Boolean

    vLines = Split(strQdef, vbCrLf)
    For lngI = LBound(vLines) To UBound(vLines)
        strLine = Trim$(vLines(lngI))
        If strLine = "Begin InputTables" Then
            blnInBlock = True
        ElseIf strLine = "End" And blnInBlock Then
            Exit For
        ElseIf blnInBlock And Left$(strLine, 6) = "Name =" Then
            strName = UnquoteQdefValue(Mid$(strLine, 7))
            If Len(strName) > 0 And Not dNames.Exists(strName) Then dNames.Add strName, True
        ElseIf blnInBlock And Left$(strLine, 7) = "Alias =" Then
            strName = UnquoteQdefValue(Mid$(strLine, 8))
            If Len(strName) > 0 And Not dNames.Exists(strName) Then dNames.Add strName, True
        End If
    Next lngI
    Set ParseInputTableNames = dNames
End Function


Private Function ParseJoinTables(ByVal strQdef As String) As Collection
    Dim colJoins As New Collection
    Dim vLines() As String
    Dim lngI As Long
    Dim strLine As String
    Dim blnInBlock As Boolean
    Dim dJn As Dictionary

    vLines = Split(strQdef, vbCrLf)
    For lngI = LBound(vLines) To UBound(vLines)
        strLine = Trim$(vLines(lngI))
        If strLine = "Begin Joins" Then
            blnInBlock = True
            Set dJn = Nothing
        ElseIf strLine = "End" And blnInBlock Then
            Exit For
        ElseIf blnInBlock Then
            If Left$(strLine, 11) = "LeftTable =" Then
                If Not dJn Is Nothing Then colJoins.Add dJn
                Set dJn = New Dictionary
                dJn.Add "left", UnquoteQdefValue(Mid$(strLine, 12))
            ElseIf Not dJn Is Nothing And Left$(strLine, 12) = "RightTable =" Then
                dJn.Add "right", UnquoteQdefValue(Mid$(strLine, 13))
            End If
        End If
    Next lngI
    If Not dJn Is Nothing Then colJoins.Add dJn
    Set ParseJoinTables = colJoins
End Function


'---------------------------------------------------------------------------------------
' Procedure : RebuildRowsFromQdef
' Purpose   : Model what Access does on LoadFromText for the two blocks that carry the
'           : join graph: InputTables becomes Attribute 5 rows and Joins becomes
'           : Attribute 7 rows. All other attributes are carried through from the
'           : original row set so the loop stays focused on the corruption path.
'           : ReconstructSQL buckets rows by Attribute, so only order within each
'           : attribute group matters here.
'---------------------------------------------------------------------------------------
'
Private Function RebuildRowsFromQdef(colOriginal As Collection, ByVal strQdef As String) As Collection
    Dim colRows As New Collection
    Dim dRow As Dictionary
    Dim varRow As Variant

    For Each varRow In colOriginal
        Set dRow = varRow
        Select Case CLng(dRow("Attribute"))
            Case 5, 7
                ' Replaced below from the emitted qdef
            Case Else
                colRows.Add dRow
        End Select
    Next varRow

    AppendInputTableRows colRows, strQdef
    AppendJoinRows colRows, strQdef

    Set RebuildRowsFromQdef = colRows
End Function


Private Sub AppendInputTableRows(colRows As Collection, ByVal strQdef As String)
    Dim vLines() As String
    Dim lngI As Long
    Dim strLine As String
    Dim dRow As Dictionary
    Dim blnInBlock As Boolean

    vLines = Split(strQdef, vbCrLf)
    For lngI = LBound(vLines) To UBound(vLines)
        strLine = Trim$(vLines(lngI))
        If strLine = "Begin InputTables" Then
            blnInBlock = True
        ElseIf blnInBlock And strLine = "End" Then
            Exit For
        ElseIf blnInBlock Then
            If Left$(strLine, 6) = "Name =" Then
                Set dRow = New Dictionary
                dRow.Add "Attribute", 5
                dRow.Add "Flag", 0
                dRow.Add "Expression", vbNullString
                dRow.Add "Name1", UnquoteQdefValue(Mid$(strLine, 7))
                dRow.Add "Name2", vbNullString
                colRows.Add dRow
            ElseIf Left$(strLine, 7) = "Alias =" Then
                If Not dRow Is Nothing Then dRow("Name2") = UnquoteQdefValue(Mid$(strLine, 8))
            End If
        End If
    Next lngI
End Sub


Private Sub AppendJoinRows(colRows As Collection, ByVal strQdef As String)
    Dim vLines() As String
    Dim lngI As Long
    Dim strLine As String
    Dim dRow As Dictionary
    Dim blnInBlock As Boolean
    Dim strLeft As String
    Dim strRight As String
    Dim strExpr As String

    vLines = Split(strQdef, vbCrLf)
    For lngI = LBound(vLines) To UBound(vLines)
        strLine = Trim$(vLines(lngI))
        If strLine = "Begin Joins" Then
            blnInBlock = True
        ElseIf blnInBlock And strLine = "End" Then
            Exit For
        ElseIf blnInBlock Then
            If Left$(strLine, 11) = "LeftTable =" Then
                strLeft = UnquoteQdefValue(Mid$(strLine, 12))
            ElseIf Left$(strLine, 12) = "RightTable =" Then
                strRight = UnquoteQdefValue(Mid$(strLine, 13))
            ElseIf Left$(strLine, 12) = "Expression =" Then
                strExpr = UnquoteQdefValue(Mid$(strLine, 13))
            ElseIf Left$(strLine, 6) = "Flag =" Then
                Set dRow = New Dictionary
                dRow.Add "Attribute", 7
                dRow.Add "Flag", CLng(Trim$(Mid$(strLine, 7)))
                dRow.Add "Expression", strExpr
                dRow.Add "Name1", strLeft
                dRow.Add "Name2", strRight
                colRows.Add dRow
            End If
        End If
    Next lngI
End Sub


Private Function BuildSpacedJoinRows() As Collection
    Dim colRows As New Collection
    Dim dRow As Dictionary

    Set dRow = New Dictionary
    dRow.Add "Attribute", 1: dRow.Add "Flag", 1
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 5: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "Car Models": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 5: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "Car Colours": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 7: dRow.Add "Flag", 1
    dRow.Add "Expression", "[Car Models].fld=[Car Colours].fld"
    dRow.Add "Name1", "Car Models": dRow.Add "Name2", "Car Colours"
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 6: dRow.Add "Flag", 0
    dRow.Add "Expression", "fld": dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 255: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set BuildSpacedJoinRows = colRows
End Function


Private Function UnquoteQdefValue(ByVal strValue As String) As String
    strValue = Trim$(strValue)
    If Left$(strValue, 1) = "=" Then strValue = Trim$(Mid$(strValue, 2))
    If Len(strValue) >= 2 Then
        If Left$(strValue, 1) = """" And Right$(strValue, 1) = """" Then
            strValue = Mid$(strValue, 2, Len(strValue) - 2)
        End If
    End If
    UnquoteQdefValue = strValue
End Function
