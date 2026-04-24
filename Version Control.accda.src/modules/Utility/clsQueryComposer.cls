VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueryComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsQueryComposer
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Bidirectional translation class for Access query definitions.
'           :
'           : EXPORT direction:
'           :   MSysQueries rows -> ReconstructSQL() -> Access SQL string
'           :
'           : IMPORT direction:
'           :   Access SQL string -> DecomposeSQL() -> internal structure
'           :   Internal structure + .json metadata -> GenerateQdef() -> .qdef text
'           :
'           : The internal structure is never serialized; it exists only transiently
'           : during the composition/decomposition process. The .sql file is the sole
'           : source of truth for SQL text.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

' Internal structure fields populated by DecomposeSQL / LoadFromMSysQueries
Private m_lngQueryType As Long          ' 1=SELECT, 3=INSERT, 4=UPDATE, 5=DELETE
Private m_lngOption As Long             ' Bitmask: 1=AllFields, 2=DISTINCT, 4=OWNERACCESS, 8=DISTINCTROW, 16=TOP, 48=TOP PERCENT
Private m_strTargetTable As String      ' For action queries
Private m_strWhere As String            ' WHERE clause (Access double-paren format)
Private m_strHaving As String           ' HAVING clause
Private m_colInputTables As Collection  ' Of Dictionary: name, alias
Private m_colOutputColumns As Collection ' Of Dictionary: expression, alias, targetName
Private m_colJoins As Collection        ' Of Dictionary: leftTable, rightTable, expression, flag
Private m_colOrderBy As Collection      ' Of Dictionary: expression, descending
Private m_colGroups As Collection       ' Of Dictionary: expression, groupLevel
Private m_strConnectString As String    ' For pass-through queries
Private m_blnDistinct As Boolean
Private m_lngTopN As Long
Private m_lngTopPercent As Long
Private m_strTransformAgg As String     ' TRANSFORM aggregate expression
Private m_strPivotField As String       ' PIVOT field
Private m_strParameters As String       ' PARAMETERS clause
Private m_strRawSql As String           ' Raw SQL (for pass-through/DDL/UNION)
Private m_blnDesignerCompatible As Boolean
Private m_blnIsPassThrough As Boolean
Private m_blnIsDDL As Boolean


'---------------------------------------------------------------------------------------
' Procedure : ReconstructSQL
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Builds an Access SQL string from MSysQueries attribute rows.
'           : Rows should be sorted by Attribute, then by ordinal position within
'           : each attribute group.
'           :
'           : For SQL View queries (identified by Attribute=0, which contains the raw
'           : SQL in Expression), returns the raw SQL directly.
'---------------------------------------------------------------------------------------
'
Public Function ReconstructSQL(colRows As Collection) As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim dRow As Dictionary
    Dim dTbl As Dictionary
    Dim dFld As Dictionary
    Dim dJn As Dictionary
    Dim dGrp As Dictionary
    Dim dOrd As Dictionary
    Dim lngAttr As Long
    Dim lngFlag As Long
    Dim strExpr As String
    Dim strN1 As String
    Dim strN2 As String
    Dim cSql As New clsConcat
    Dim colFields As Collection
    Dim colTables As Collection
    Dim colJoinParts As Collection
    Dim colWhere As Collection
    Dim colHaving As Collection
    Dim colOrder As Collection
    Dim colGroup As Collection
    Dim lngOperation As Long
    Dim lngOption As Long
    Dim strTarget As String
    Dim strConnect As String
    Dim strRawSql As String
    Dim blnHasRawSql As Boolean
    Dim blnHasAttr1 As Boolean
    Dim strParams As String
    Dim blnDistinct As Boolean
    Dim blnDistinctRow As Boolean
    Dim lngTopN As Long
    Dim blnTopPercent As Boolean

    ' Initialize collections for each section
    Set colFields = New Collection
    Set colTables = New Collection
    Set colJoinParts = New Collection
    Set colWhere = New Collection
    Set colHaving = New Collection
    Set colOrder = New Collection
    Set colGroup = New Collection

    ' First pass: categorize rows by attribute
    For Each dRow In colRows
        lngAttr = CLng(dRow("Attribute"))
        lngFlag = CLng(Nz(dRow("Flag"), 0))
        strExpr = Nz(dRow("Expression"), vbNullString)
        strN1 = Nz(dRow("Name1"), vbNullString)
        strN2 = Nz(dRow("Name2"), vbNullString)

        Select Case lngAttr
            Case 0
                ' Raw SQL (SQL View) -- only when Expression is non-empty.
                ' Designed queries also have an Attribute 0 row with null Expression.
                If Len(strExpr) > 0 Then
                    strRawSql = strExpr
                    blnHasRawSql = True
                End If

            Case 1
                ' Operation type
                lngOperation = lngFlag
                strTarget = strN1
                blnHasAttr1 = True
                ' DDL (Flag=7) and Pass-through (Flag=8) store raw SQL in Expression
                If lngFlag = 7 Or lngFlag = 8 Then
                    strRawSql = strExpr
                    blnHasRawSql = True
                End If

            Case 2
                ' PARAMETERS
                strParams = strExpr

            Case 3
                ' Options bitmask: 2=DISTINCT, 8=DISTINCTROW, 16=TOP, 32=PERCENT
                ' When TOP is set, Name1 holds the TOP n value.
                lngOption = lngFlag
                blnDistinct = (lngFlag And 2) <> 0
                blnDistinctRow = (lngFlag And 8) <> 0
                If (lngFlag And 16) <> 0 And Len(strN1) > 0 Then
                    lngTopN = CLng(strN1)
                    blnTopPercent = (lngFlag And 32) <> 0
                End If

            Case 4
                ' Connect string (pass-through)
                strConnect = strExpr

            Case 5
                ' Input table (also UNION segment when Expression contains SQL)
                Set dTbl = New Dictionary
                dTbl.Add "name", strN1
                dTbl.Add "alias", strN2
                dTbl.Add "expression", strExpr
                colTables.Add dTbl

            Case 6
                ' Output column (Flag differentiates crosstab roles:
                ' 0=aggregate/normal, 1=PIVOT column heading, 2=row heading)
                Set dFld = New Dictionary
                dFld.Add "expression", strExpr
                dFld.Add "alias", strN1
                dFld.Add "targetName", strN2
                dFld.Add "flag", lngFlag
                colFields.Add dFld

            Case 7
                ' Join
                Set dJn = New Dictionary
                dJn.Add "leftTable", strN1
                dJn.Add "rightTable", strN2
                dJn.Add "expression", strExpr
                dJn.Add "flag", lngFlag
                colJoinParts.Add dJn

            Case 8
                ' WHERE
                colWhere.Add strExpr

            Case 9
                ' GROUP BY
                Set dGrp = New Dictionary
                dGrp.Add "expression", strExpr
                dGrp.Add "groupLevel", lngFlag
                colGroup.Add dGrp

            Case 10
                ' HAVING
                colHaving.Add strExpr

            Case 11
                ' ORDER BY
                Set dOrd = New Dictionary
                dOrd.Add "expression", strExpr
                dOrd.Add "descending", (LCase$(strN1) = "d")
                colOrder.Add dOrd

            Case 255
                ' End of record marker - no action needed
        End Select
    Next dRow

    ' If raw SQL was found (SQL View or DDL/pass-through), return it directly
    If blnHasRawSql Then
        ReconstructSQL = strRawSql
        GoTo CleanUp
    End If

    ' Default to SELECT when no Attribute 1 present (SQL View SELECT queries
    ' omit Attribute 1 Flag=1 and Attribute 3 Flag=0 as they are defaults)
    If Not blnHasAttr1 Then lngOperation = 1

    ' Check if this is a crosstab query (operation 6 = TRANSFORM)
    Dim blnIsCrosstab As Boolean
    blnIsCrosstab = (lngOperation = 6)

    ' Build SQL from decomposed structure
    cSql.AppendOnAdd = vbNullString

    ' PARAMETERS
    If Len(strParams) > 0 Then
        cSql.Add "PARAMETERS " & strParams & ";" & vbCrLf
    End If

    ' TRANSFORM (crosstab queries) -- aggregate field is Attribute 6 Flag=0
    If blnIsCrosstab Then
        Dim dAggFld As Dictionary
        For Each dFld In colFields
            If CLng(dFld("flag")) = 0 Then
                Set dAggFld = dFld
                Exit For
            End If
        Next dFld
        If Not dAggFld Is Nothing Then
            cSql.Add "TRANSFORM " & CStr(dAggFld("expression"))
            If Len(dAggFld("alias")) > 0 Then
                cSql.Add " AS " & BracketIfNeeded(CStr(dAggFld("alias")))
            End If
            cSql.Add vbCrLf
        End If
    End If

    ' Build the SELECT keyword + DISTINCT/DISTINCTROW + TOP n [PERCENT] modifiers
    ' shared by SELECT, crosstab, Make-Table, and INSERT INTO ... SELECT.
    Dim strSelectModifiers As String
    strSelectModifiers = vbNullString
    If blnDistinctRow Then
        strSelectModifiers = strSelectModifiers & "DISTINCTROW "
    ElseIf blnDistinct Then
        strSelectModifiers = strSelectModifiers & "DISTINCT "
    End If
    If lngTopN > 0 Then
        strSelectModifiers = strSelectModifiers & "TOP " & CStr(lngTopN) & " "
        If blnTopPercent Then strSelectModifiers = strSelectModifiers & "PERCENT "
    End If

    ' Main query body based on operation type
    Select Case lngOperation
        Case 1, 6  ' SELECT / Crosstab row headings
            cSql.Add "SELECT " & strSelectModifiers
            ' For crosstab, only include row heading fields (Flag=2)
            If blnIsCrosstab Then
                Dim colRowFields As New Collection
                For Each dFld In colFields
                    If CLng(dFld("flag")) = 2 Then colRowFields.Add dFld
                Next dFld
                BuildFieldList cSql, colRowFields
            Else
                BuildFieldList cSql, colFields
            End If

        Case 2  ' Make-Table (SELECT...INTO)
            cSql.Add "SELECT " & strSelectModifiers
            BuildFieldList cSql, colFields
            cSql.Add vbCrLf & "INTO " & BracketIfNeeded(strTarget)

        Case 3  ' INSERT INTO
            cSql.Add "INSERT INTO " & BracketIfNeeded(strTarget)
            ' Check if there are target field names
            Dim colTargetFields As New Collection
            Dim dF As Dictionary
            For Each dF In colFields
                If Len(dF("targetName")) > 0 Then
                    colTargetFields.Add dF("targetName")
                End If
            Next dF
            If colTargetFields.Count > 0 Then
                cSql.Add " ("
                Dim lngTF As Long
                For lngTF = 1 To colTargetFields.Count
                    If lngTF > 1 Then cSql.Add ", "
                    cSql.Add BracketIfNeeded(CStr(colTargetFields(lngTF)))
                Next lngTF
                cSql.Add ")"
            End If
            cSql.Add vbCrLf
            cSql.Add "SELECT " & strSelectModifiers
            BuildFieldList cSql, colFields

        Case 4  ' UPDATE
            cSql.Add "UPDATE "
            BuildFromClause cSql, colTables, colJoinParts
            cSql.Add vbCrLf & "SET "
            BuildUpdateSetClause cSql, colFields

        Case 5  ' DELETE
            cSql.Add "DELETE "
            BuildFieldList cSql, colFields

        Case 9  ' UNION / UNION ALL
            Dim blnUnionAll As Boolean
            blnUnionAll = ((lngOption And 2) = 0)
            Dim lngSeg As Long
            For lngSeg = 1 To colTables.Count
                If lngSeg > 1 Then
                    If blnUnionAll Then
                        cSql.Add vbCrLf & vbCrLf & "UNION ALL "
                    Else
                        cSql.Add vbCrLf & vbCrLf & "UNION "
                    End If
                End If
                cSql.Add RTrim$(CStr(colTables(lngSeg)("expression")))
            Next lngSeg
            cSql.Add ";" & vbCrLf
            ReconstructSQL = cSql.GetStr
            Exit Function
    End Select

    ' FROM clause (for SELECT, INSERT, DELETE, Make-Table -- UPDATE already has it)
    If lngOperation <> 4 Then
        cSql.Add vbCrLf & "FROM "
        BuildFromClause cSql, colTables, colJoinParts
    End If

    ' WHERE
    If colWhere.Count > 0 Then
        cSql.Add vbCrLf & "WHERE " & BuildConditionClause(colWhere)
    End If

    ' GROUP BY (for crosstab, skip PIVOT GROUP BY entries with Flag=1)
    If colGroup.Count > 0 Then
        cSql.Add vbCrLf & "GROUP BY "
        Dim lngG As Long
        Dim blnFirstGroup As Boolean
        blnFirstGroup = True
        For lngG = 1 To colGroup.Count
            If blnIsCrosstab And CLng(colGroup(lngG)("groupLevel")) = 1 Then
                ' Skip PIVOT column heading GROUP BY for crosstab
            Else
                If Not blnFirstGroup Then cSql.Add ", "
                cSql.Add CStr(colGroup(lngG)("expression"))
                blnFirstGroup = False
            End If
        Next lngG
    End If

    ' HAVING
    If colHaving.Count > 0 Then
        cSql.Add vbCrLf & "HAVING " & BuildConditionClause(colHaving)
    End If

    ' ORDER BY
    If colOrder.Count > 0 Then
        cSql.Add vbCrLf & "ORDER BY "
        Dim lngO As Long
        For lngO = 1 To colOrder.Count
            If lngO > 1 Then cSql.Add ", "
            cSql.Add CStr(colOrder(lngO)("expression"))
            If colOrder(lngO)("descending") Then cSql.Add " DESC"
        Next lngO
    End If

    ' PIVOT (crosstab) -- column heading field is Attribute 6 Flag=1
    If blnIsCrosstab Then
        For Each dFld In colFields
            If CLng(dFld("flag")) = 1 Then
                cSql.Add vbCrLf & "PIVOT " & CStr(dFld("expression"))
                Exit For
            End If
        Next dFld
    End If

    ' Terminate with semicolon and newline
    cSql.Add ";" & vbCrLf

    ReconstructSQL = cSql.GetStr

CleanUp:
    CatchAny eelError, "Error reconstructing SQL (" & colRows.Count & " rows)", _
        ModuleName(Me) & ".ReconstructSQL"

End Function


'---------------------------------------------------------------------------------------
' Procedure : DecomposeSQL
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Parses an Access SQL string into the internal structure used for .qdef
'           : generation. Sets m_blnDesignerCompatible to indicate whether the query
'           : can be represented in Design View format.
'           :
'           : Handles: SELECT, INSERT INTO, UPDATE, DELETE, TRANSFORM/PIVOT.
'           : UNION, pass-through, DDL, and complex subqueries set
'           : m_blnDesignerCompatible = False and store raw SQL only.
'---------------------------------------------------------------------------------------
'
Public Sub DecomposeSQL(strSql As String, Optional blnIsPassThrough As Boolean = False)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strNorm As String
    Dim strUpper As String

    ' Reset internal state
    ResetState

    m_strRawSql = strSql
    m_blnIsPassThrough = blnIsPassThrough

    ' Pass-through queries are never designer-compatible
    If blnIsPassThrough Then
        m_blnDesignerCompatible = False
        GoTo CleanUp
    End If

    ' Normalize whitespace for parsing
    strNorm = NormalizeWhitespace(strSql)
    strUpper = UCase$(strNorm)

    ' Check for query types that cannot be represented in Design View
    If IsUnionQuery(strUpper) Then
        m_blnDesignerCompatible = False
        GoTo CleanUp
    End If

    ' DDL statements
    If IsDDLQuery(strUpper) Then
        m_blnIsDDL = True
        m_blnDesignerCompatible = False
        GoTo CleanUp
    End If

    ' Strip trailing semicolon and whitespace
    strNorm = TrimTrailingSemicolon(strNorm)
    strUpper = UCase$(strNorm)

    ' Check for WITH OWNERACCESS OPTION at the end
    Dim lngOwnerAccess As Long
    lngOwnerAccess = FindTopLevelKeyword(strNorm, "WITH OWNERACCESS OPTION")
    If lngOwnerAccess > 0 Then
        ' Verify it's at the very end
        If Len(Trim$(Mid$(strNorm, lngOwnerAccess + 23))) = 0 Then
            m_lngOption = m_lngOption Or 4
            strNorm = RTrim$(Left$(strNorm, lngOwnerAccess - 1))
            strUpper = UCase$(strNorm)
        End If
    End If

    ' Parse PARAMETERS clause if present
    If Left$(strUpper, 11) = "PARAMETERS " Then
        ParseParametersClause strNorm
        strNorm = TrimLeadingWhitespace(strNorm)
        strUpper = UCase$(strNorm)
    End If

    ' Parse TRANSFORM/PIVOT (crosstab)
    If Left$(strUpper, 10) = "TRANSFORM " Then
        ParseTransformClause strNorm
        strNorm = TrimLeadingWhitespace(strNorm)
        strUpper = UCase$(strNorm)
    End If

    ' Determine query type and parse accordingly
    If Left$(strUpper, 7) = "SELECT " Then
        m_lngQueryType = 1
        ParseSelectQuery strNorm
    ElseIf Left$(strUpper, 12) = "INSERT INTO " Then
        m_lngQueryType = 3
        ParseInsertQuery strNorm
    ElseIf Left$(strUpper, 7) = "UPDATE " Then
        m_lngQueryType = 4
        ParseUpdateQuery strNorm
    ElseIf Left$(strUpper, 7) = "DELETE " Then
        m_lngQueryType = 5
        ParseDeleteQuery strNorm
    Else
        ' Unknown query type -- treat as raw SQL
        m_blnDesignerCompatible = False
        GoTo CleanUp
    End If

    ' Check for subqueries in the FROM clause which make Design View impossible
    If HasSubqueries() Then
        m_blnDesignerCompatible = False
    Else
        m_blnDesignerCompatible = True
    End If

CleanUp:
    CatchAny eelError, "Error decomposing SQL. " & SqlSnippet(), _
        ModuleName(Me) & ".DecomposeSQL"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsDesignerCompatible
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Returns True if the SQL parsed by DecomposeSQL can be represented in
'           : Design View format (Operation/InputTables/OutputColumns/Joins/etc.)
'---------------------------------------------------------------------------------------
'
Public Function IsDesignerCompatible() As Boolean
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    IsDesignerCompatible = m_blnDesignerCompatible
    CatchAny eelError, "Error reading IsDesignerCompatible", _
        ModuleName(Me) & ".IsDesignerCompatible"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GenerateQdef
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Generates .qdef text for LoadFromText.
'           :
'           : If blnDesignView is True and the query is designer-compatible, emits
'           : Design View format (Operation, InputTables, OutputColumns, Joins, etc.)
'           : plus the layout from dLayout. Otherwise emits SQL View format
'           : (dbMemo "SQL" = "...").
'           :
'           : dProperties: query properties (ReturnsRecords, ODBCTimeout, etc.)
'           : dColumns: column metadata (Name, AggregateType, ColumnWidth, etc.)
'           : dLayout: design layout from LvExtra parsing (or Nothing for SQL View)
'---------------------------------------------------------------------------------------
'
Public Function GenerateQdef(blnDesignView As Boolean, _
                             dProperties As Dictionary, _
                             dColumns As Dictionary, _
                             Optional dLayout As Dictionary = Nothing, _
                             Optional lngOptionFlag As Long = -1) As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim cOut As New clsConcat

    ' Merge JSON OptionFlag with SQL-derived flags. The SQL is the source of
    ' truth -- if the JSON OptionFlag claims TOP/PERCENT but the parsed SQL did
    ' not produce a TOP n value, drop those bits to avoid generating an Option
    ' line that points to a non-existent TopN. Otherwise Access fails the import
    ' with "Resource failure".
    If lngOptionFlag > 0 Then
        Dim lngMerged As Long
        lngMerged = lngOptionFlag
        If m_lngTopN = 0 And m_lngTopPercent = 0 Then
            ' Clear TOP (16) and PERCENT (32) bits when no TOP value is present
            lngMerged = lngMerged And Not 16
            lngMerged = lngMerged And Not 32
            If lngMerged <> lngOptionFlag Then
                Log.Add "    Dropping TOP/PERCENT bits from OptionFlag (no TOP n in SQL)", Options.ShowDebug
            End If
        End If
        m_lngOption = m_lngOption Or lngMerged
    End If

    If blnDesignView And m_blnDesignerCompatible Then
        EmitDesignViewQdef cOut, dProperties, dColumns, dLayout
    Else
        EmitSqlViewQdef cOut, dProperties, dColumns
    End If

    GenerateQdef = cOut.GetStr

CleanUp:
    CatchAny eelError, "Error generating qdef (DesignView=" & blnDesignView & "). " & SqlSnippet(), _
        ModuleName(Me) & ".GenerateQdef"

End Function


'=======================================================================================
' PRIVATE: SQL Reconstruction helpers
'=======================================================================================

Private Sub BuildFieldList(cSql As clsConcat, colFields As Collection)
    Dim lngF As Long
    Dim dFld As Dictionary
    For lngF = 1 To colFields.Count
        If lngF > 1 Then cSql.Add ", "
        Set dFld = colFields(lngF)
        cSql.Add CStr(dFld("expression"))
        If Len(dFld("alias")) > 0 Then
            cSql.Add " AS " & BracketIfNeeded(CStr(dFld("alias")))
        End If
    Next lngF
End Sub


Private Sub BuildUpdateSetClause(cSql As clsConcat, colFields As Collection)
    Dim lngF As Long
    Dim dFld As Dictionary
    For lngF = 1 To colFields.Count
        If lngF > 1 Then cSql.Add ", "
        Set dFld = colFields(lngF)
        cSql.Add BracketIfNeeded(CStr(dFld("targetName")))
        cSql.Add " = " & CStr(dFld("expression"))
    Next lngF
End Sub


Private Sub BuildFromClause(cSql As clsConcat, colTables As Collection, colJoins As Collection)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim lngJ As Long
    Dim dJn As Dictionary
    Dim strJoinType As String
    Dim dTbl As Dictionary

    If colJoins.Count = 0 Then
        ' No joins -- list tables comma-separated (cartesian product)
        Dim lngT As Long
        For lngT = 1 To colTables.Count
            If lngT > 1 Then cSql.Add ", "
            Set dTbl = colTables(lngT)
            cSql.Add BracketIfNeeded(CStr(dTbl("name")))
            If Len(dTbl("alias")) > 0 Then
                cSql.Add " AS " & BracketIfNeeded(CStr(dTbl("alias")))
            End If
        Next lngT
    Else
        ' Consolidate multi-condition joins: Access stores each ON condition
        ' as a separate Attribute 7 row with potentially swapped Name1/Name2.
        Dim colConsolidated As Collection
        Set colConsolidated = ConsolidateJoins(colJoins)

        ' Normalize RIGHT JOINs to LEFT JOINs so the hub table becomes the
        ' graph root in BuildJoinChain. A RIGHT JOIN B ON expr is semantically
        ' identical to B LEFT JOIN A ON expr. We flip the direction and mark
        ' the join with wasRight so the emission step restores RIGHT JOIN syntax.
        For lngJ = 1 To colConsolidated.Count
            Set dJn = colConsolidated(lngJ)
            If CLng(dJn("flag")) = 3 Then
                Dim strOrigLeft As String
                strOrigLeft = CStr(dJn("leftTable"))
                dJn("leftTable") = CStr(dJn("rightTable"))
                dJn("rightTable") = strOrigLeft
                dJn("flag") = CLng(2)
                dJn.Add "wasRight", True
            End If
        Next lngJ

        ' Build the join chain by following the leftTable->rightTable graph.
        ' DFS with deterministic sorting (flag-first, then alphabetical by
        ' rightTable) ensures stable output regardless of MSysQueries row order.
        Dim colOrdered As Collection
        Set colOrdered = BuildJoinChain(colConsolidated)

        ' Build table alias lookup for proper FROM clause formatting
        Dim dTableLookup As Dictionary
        Set dTableLookup = New Dictionary
        For Each dTbl In colTables
            Dim strRef As String
            If Len(dTbl("alias")) > 0 Then
                strRef = CStr(dTbl("alias"))
                If Not dTableLookup.Exists(strRef) Then
                    dTableLookup.Add strRef, CStr(dTbl("name")) & " AS " & BracketIfNeeded(strRef)
                End If
            Else
                strRef = CStr(dTbl("name"))
                If Not dTableLookup.Exists(strRef) Then
                    dTableLookup.Add strRef, BracketIfNeeded(strRef)
                End If
            End If
        Next dTbl

        ' Sequential accumulation: process the ordered joins one at a time,
        ' building the FROM clause string. Left-nests all join types.
        ' For wasRight joins, restores RIGHT JOIN syntax by placing the new
        ' table on the left and the accumulated subtree as the right operand.
        Dim strAccum As String
        Dim strRootRef As String
        strRootRef = CStr(colOrdered(1)("leftTable"))
        If dTableLookup.Exists(strRootRef) Then
            strAccum = CStr(dTableLookup(strRootRef))
        Else
            strAccum = BracketIfNeeded(strRootRef)
        End If

        Dim blnHasJoins As Boolean
        blnHasJoins = False

        For lngJ = 1 To colOrdered.Count
            Set dJn = colOrdered(lngJ)

            Dim strRightRef As String
            Dim strTableName As String

            If dJn.Exists("wasRight") Then
                ' Restore RIGHT JOIN syntax: the rightTable (after flip) is the
                ' original leftTable, which becomes the left operand of RIGHT JOIN.
                strRightRef = CStr(dJn("rightTable"))
                If dTableLookup.Exists(strRightRef) Then
                    strTableName = CStr(dTableLookup(strRightRef))
                Else
                    strTableName = BracketIfNeeded(strRightRef)
                End If
                If blnHasJoins Then
                    strAccum = strTableName & " RIGHT JOIN (" & strAccum & _
                        ") ON " & CStr(dJn("expression"))
                Else
                    strAccum = strTableName & " RIGHT JOIN " & strAccum & _
                        " ON " & CStr(dJn("expression"))
                End If
            Else
                ' LEFT/INNER JOIN: accumulated on left, new table on right
                Select Case CLng(dJn("flag"))
                    Case 1: strJoinType = " INNER JOIN "
                    Case 2: strJoinType = " LEFT JOIN "
                    Case Else: strJoinType = " INNER JOIN "
                End Select

                strRightRef = CStr(dJn("rightTable"))
                If dTableLookup.Exists(strRightRef) Then
                    strTableName = CStr(dTableLookup(strRightRef))
                Else
                    strTableName = BracketIfNeeded(strRightRef)
                End If
                If blnHasJoins Then
                    strAccum = "(" & strAccum & ")" & strJoinType & strTableName & _
                        " ON " & CStr(dJn("expression"))
                Else
                    strAccum = strAccum & strJoinType & strTableName & _
                        " ON " & CStr(dJn("expression"))
                End If
            End If

            blnHasJoins = True
        Next lngJ

        cSql.Add strAccum
    End If

    CatchAny eelError, "Error building FROM clause. " & SqlSnippet(), _
        ModuleName(Me) & ".BuildFromClause"
End Sub


Private Function ConsolidateJoins(colRawJoins As Collection) As Collection

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Access stores each ON condition as a separate Attribute 7 row.
    ' Two issues can create invalid join graphs:
    '
    ' 1) Swapped Name1/Name2: ON (A.x = B.y) AND (B.z = A.w) produces
    '    rows with Name1/Name2 flipped based on expression column order.
    '    These share the same table pair and are grouped in Phase 1.
    '
    ' 2) Cross-table ON conditions: a single JOIN can have ON conditions
    '    referencing previously-joined tables, producing rows with DIFFERENT
    '    table pairs that create graph cycles. These are detected via
    '    Union-Find in Phase 2 and merged into the appropriate primary join.
    '
    ' NOTE: Phase 1 uses the first row's direction as the canonical left/right
    ' for each table pair. If MSysQueries row order is non-deterministic for
    ' multi-condition joins, the direction could vary between runs. If this
    ' causes issues, add a deterministic normalization rule (e.g. alphabetically
    ' lower table is always leftTable) after Phase 1.

    ' --- Phase 1: Group rows with the same (normalized) table pair ---
    Dim colGrouped As Collection
    Set colGrouped = New Collection

    Dim dGroupIndex As Dictionary
    Set dGroupIndex = New Dictionary

    Dim dJn As Dictionary
    Dim strLeft As String
    Dim strRight As String
    Dim strKey As String

    For Each dJn In colRawJoins
        strLeft = CStr(dJn("leftTable"))
        strRight = CStr(dJn("rightTable"))

        ' Normalize key: alphabetically lower name first
        If StrComp(strLeft, strRight, vbTextCompare) <= 0 Then
            strKey = strLeft & vbNullChar & strRight
        Else
            strKey = strRight & vbNullChar & strLeft
        End If

        If dGroupIndex.Exists(strKey) Then
            Dim dExisting As Dictionary
            Set dExisting = colGrouped(CLng(dGroupIndex(strKey)))
            dExisting("expression") = CStr(dExisting("expression")) & _
                ") AND (" & CStr(dJn("expression"))
            dExisting("merged") = True
        Else
            Dim dNew As Dictionary
            Set dNew = New Dictionary
            dNew.Add "leftTable", strLeft
            dNew.Add "rightTable", strRight
            dNew.Add "expression", CStr(dJn("expression"))
            dNew.Add "flag", CLng(dJn("flag"))
            colGrouped.Add dNew
            dGroupIndex.Add strKey, colGrouped.Count
        End If
    Next dJn

    ' --- Phase 2: Detect secondary conditions via Union-Find ---
    ' An edge where both endpoints are already connected is a secondary
    ' ON condition (it would create a cycle). Merge it into the primary
    ' join that introduced one of its tables.
    Dim colPrimary As Collection
    Set colPrimary = New Collection
    Dim colSecondary As Collection
    Set colSecondary = New Collection
    Dim dUFParent As Dictionary
    Set dUFParent = New Dictionary

    Dim dItem As Dictionary
    For Each dItem In colGrouped
        strLeft = CStr(dItem("leftTable"))
        strRight = CStr(dItem("rightTable"))

        Dim strRootL As String
        Dim strRootR As String
        strRootL = UFFind(dUFParent, strLeft)
        strRootR = UFFind(dUFParent, strRight)

        If strRootL = strRootR Then
            colSecondary.Add dItem
        Else
            dUFParent(strRootL) = strRootR
            colPrimary.Add dItem
        End If
    Next dItem

    ' Merge each secondary condition into the deepest primary join
    ' that introduced one of the secondary's tables (match on rightTable).
    Dim dSec As Dictionary
    For Each dSec In colSecondary
        Dim lngP As Long
        Dim blnMerged As Boolean
        blnMerged = False
        For lngP = colPrimary.Count To 1 Step -1
            Dim dPri As Dictionary
            Set dPri = colPrimary(lngP)
            If StrComp(CStr(dPri("rightTable")), CStr(dSec("leftTable")), vbTextCompare) = 0 Or _
               StrComp(CStr(dPri("rightTable")), CStr(dSec("rightTable")), vbTextCompare) = 0 Then
                dPri("expression") = CStr(dPri("expression")) & _
                    ") AND (" & CStr(dSec("expression"))
                dPri("merged") = True
                blnMerged = True
                Exit For
            End If
        Next lngP
        If Not blnMerged Then
            colPrimary.Add dSec
        End If
    Next dSec

    ' Wrap expressions that were assembled by merging multiple rows.
    ' Expressions that arrived already compound (single MSysQueries row)
    ' are left unchanged to avoid double-wrapping.
    For Each dItem In colPrimary
        If dItem.Exists("merged") Then
            dItem("expression") = "(" & CStr(dItem("expression")) & ")"
        End If
    Next dItem

    Set ConsolidateJoins = colPrimary

    CatchAny eelError, "Error consolidating joins. " & SqlSnippet(), _
        ModuleName(Me) & ".ConsolidateJoins"
End Function


Private Function UFFind(dParent As Dictionary, strNode As String) As String
    ' Union-Find with path compression
    If Not dParent.Exists(strNode) Then
        dParent.Add strNode, strNode
    End If
    Do While CStr(dParent(strNode)) <> strNode
        dParent(strNode) = CStr(dParent(CStr(dParent(strNode))))
        strNode = CStr(dParent(strNode))
    Loop
    UFFind = strNode
End Function


Private Function BuildJoinChain(colJoins As Collection) As Collection
    ' Orders joins into correct SQL nesting order by following the
    ' leftTable -> rightTable graph. Finds the root table (appears
    ' only as leftTable, never as rightTable) and traverses outward.
    ' Uses DFS so that child joins (e.g. tblListAction from
    ' qrySecFirstAction) are emitted immediately after their parent,
    ' matching Access's native join ordering.
    ' For star joins (multiple joins from the same leftTable), sorts
    ' by flag first (INNER before LEFT/RIGHT) then by rightTable name.

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim dJn As Dictionary
    Dim strLeft As String
    Dim strRight As String

    ' Collect all rightTable names
    Dim dRightTables As Dictionary
    Set dRightTables = New Dictionary
    For Each dJn In colJoins
        strRight = CStr(dJn("rightTable"))
        If Not dRightTables.Exists(strRight) Then
            dRightTables.Add strRight, True
        End If
    Next dJn

    ' Find root: a leftTable that never appears as a rightTable
    Dim strRoot As String
    For Each dJn In colJoins
        strLeft = CStr(dJn("leftTable"))
        If Not dRightTables.Exists(strLeft) Then
            strRoot = strLeft
            Exit For
        End If
    Next dJn

    ' Build multimap: leftTable -> Collection of joins from that table
    ' Sort each group by flag first, then rightTable name
    Dim dJoinsByLeft As Dictionary
    Set dJoinsByLeft = New Dictionary
    For Each dJn In colJoins
        strLeft = CStr(dJn("leftTable"))
        If Not dJoinsByLeft.Exists(strLeft) Then
            dJoinsByLeft.Add strLeft, New Collection
        End If
        InsertJoinSorted dJoinsByLeft(strLeft), dJn
    Next dJn

    ' DFS from root to build ordered chain
    Dim colResult As Collection
    Set colResult = New Collection
    Dim dVisited As Dictionary
    Set dVisited = New Dictionary
    dVisited.Add strRoot, True

    DFSTraverse strRoot, dJoinsByLeft, dVisited, colResult

    Set BuildJoinChain = colResult

    CatchAny eelError, "Error building join chain. " & SqlSnippet(), _
        ModuleName(Me) & ".BuildJoinChain"
End Function


Private Sub DFSTraverse(strNode As String, dJoinsByLeft As Dictionary, _
                        dVisited As Dictionary, colResult As Collection)
    If Not dJoinsByLeft.Exists(strNode) Then Exit Sub
    Dim colFromNode As Collection
    Set colFromNode = dJoinsByLeft(strNode)
    Dim dJnItem As Dictionary
    For Each dJnItem In colFromNode
        colResult.Add dJnItem
        Dim strChild As String
        strChild = CStr(dJnItem("rightTable"))
        If Not dVisited.Exists(strChild) Then
            dVisited.Add strChild, True
            DFSTraverse strChild, dJoinsByLeft, dVisited, colResult
        End If
    Next dJnItem
End Sub


Private Sub InsertJoinSorted(colTarget As Collection, dJoin As Dictionary)
    ' Insert a join dictionary sorted by: flag first (INNER=1 before LEFT=2
    ' before RIGHT=3), then by rightTable name within the same flag.
    ' Flag-first ensures INNER JOINs are processed before LEFT/RIGHT JOINs
    ' in star-join scenarios, matching Access's nesting convention.
    Dim lngNewFlag As Long
    Dim strNewRight As String
    Dim lngI As Long
    lngNewFlag = CLng(dJoin("flag"))
    strNewRight = CStr(dJoin("rightTable"))
    For lngI = 1 To colTarget.Count
        Dim lngExistFlag As Long
        lngExistFlag = CLng(colTarget(lngI)("flag"))
        If lngNewFlag < lngExistFlag Then
            colTarget.Add dJoin, Before:=lngI
            Exit Sub
        ElseIf lngNewFlag = lngExistFlag Then
            If StrComp(strNewRight, CStr(colTarget(lngI)("rightTable")), vbTextCompare) < 0 Then
                colTarget.Add dJoin, Before:=lngI
                Exit Sub
            End If
        End If
    Next lngI
    colTarget.Add dJoin
End Sub


Private Function BuildConditionClause(colConditions As Collection) As String
    Dim lngC As Long
    Dim strResult As String
    For lngC = 1 To colConditions.Count
        If lngC > 1 Then strResult = strResult & " AND "
        strResult = strResult & CStr(colConditions(lngC))
    Next lngC
    BuildConditionClause = strResult
End Function


'---------------------------------------------------------------------------------------
' Procedure : SqlSnippet
' Author    : Adam Waller
' Date      : 4/20/2026
' Purpose   : Returns a truncated snippet of m_strRawSql for use in error messages.
'           : Avoids logging multi-thousand-character SQL strings while still giving
'           : enough context to identify which query failed.
'---------------------------------------------------------------------------------------
'
Private Function SqlSnippet() As String
    Const MAX_LEN As Long = 200
    If Len(m_strRawSql) = 0 Then
        SqlSnippet = "(no SQL captured)"
    ElseIf Len(m_strRawSql) > MAX_LEN Then
        SqlSnippet = "SQL: " & Left$(m_strRawSql, MAX_LEN) & "..."
    Else
        SqlSnippet = "SQL: " & m_strRawSql
    End If
End Function


Private Function BracketIfNeeded(strName As String) As String
    ' Return name as-is if already bracketed or contains a dot (qualified)
    If Left$(strName, 1) = "[" Then
        BracketIfNeeded = strName
    ElseIf InStr(1, strName, ".") > 0 Then
        BracketIfNeeded = strName
    ElseIf InStr(1, strName, " ") > 0 Or IsAccessReservedWord(strName) Then
        BracketIfNeeded = "[" & strName & "]"
    Else
        BracketIfNeeded = strName
    End If
End Function


Private Function IsAccessReservedWord(strWord As String) As Boolean
    Select Case UCase$(strWord)
        Case "SELECT", "FROM", "WHERE", "ORDER", "BY", "GROUP", "HAVING", _
             "INSERT", "INTO", "UPDATE", "DELETE", "SET", "VALUES", _
             "JOIN", "INNER", "LEFT", "RIGHT", "OUTER", "ON", "AS", _
             "AND", "OR", "NOT", "IN", "BETWEEN", "LIKE", "IS", "NULL", _
             "EXISTS", "ALL", "ANY", "SOME", "UNION", "EXCEPT", "INTERSECT", _
             "DISTINCT", "TOP", "TRANSFORM", "PIVOT", "TABLE", "CREATE", _
             "ALTER", "DROP", "INDEX", "VIEW", "PROCEDURE", "PARAMETER", _
             "PARAMETERS", "TRUE", "FALSE", "YES", "NO", "DATE", "TIME", _
             "TEXT", "MEMO", "LONG", "SHORT", "SINGLE", "DOUBLE", "CURRENCY", _
             "COUNTER", "YESNO", "BINARY", "OLEOBJECT", "GUID"
            IsAccessReservedWord = True
    End Select
End Function


'=======================================================================================
' PRIVATE: SQL Decomposition (parsing)
'=======================================================================================

Private Sub ResetState()
    m_lngQueryType = 0
    m_lngOption = 0
    m_strTargetTable = vbNullString
    m_strWhere = vbNullString
    m_strHaving = vbNullString
    Set m_colInputTables = New Collection
    Set m_colOutputColumns = New Collection
    Set m_colJoins = New Collection
    Set m_colOrderBy = New Collection
    Set m_colGroups = New Collection
    m_strConnectString = vbNullString
    m_blnDistinct = False
    m_lngTopN = 0
    m_lngTopPercent = 0
    m_strTransformAgg = vbNullString
    m_strPivotField = vbNullString
    m_strParameters = vbNullString
    m_strRawSql = vbNullString
    m_blnDesignerCompatible = False
    m_blnIsPassThrough = False
    m_blnIsDDL = False
End Sub


Private Function NormalizeWhitespace(strSql As String) As String
    Dim strResult As String
    Dim lngPos As Long
    Dim strChar As String
    Dim blnInString As Boolean
    Dim blnLastWasSpace As Boolean

    strResult = Replace(strSql, vbCrLf, " ")
    strResult = Replace(strResult, vbCr, " ")
    strResult = Replace(strResult, vbLf, " ")
    strResult = Replace(strResult, vbTab, " ")

    ' Collapse multiple spaces (outside quotes)
    Dim strOut As String
    strOut = vbNullString
    blnInString = False
    blnLastWasSpace = False
    For lngPos = 1 To Len(strResult)
        strChar = Mid$(strResult, lngPos, 1)
        If strChar = """" Then blnInString = Not blnInString
        If blnInString Then
            strOut = strOut & strChar
            blnLastWasSpace = False
        Else
            If strChar = " " Then
                If Not blnLastWasSpace Then
                    strOut = strOut & " "
                    blnLastWasSpace = True
                End If
            Else
                strOut = strOut & strChar
                blnLastWasSpace = False
            End If
        End If
    Next lngPos

    NormalizeWhitespace = Trim$(strOut)
End Function


Private Function TrimTrailingSemicolon(strSql As String) As String
    TrimTrailingSemicolon = RTrim$(strSql)
    If Right$(TrimTrailingSemicolon, 1) = ";" Then
        TrimTrailingSemicolon = RTrim$(Left$(TrimTrailingSemicolon, Len(TrimTrailingSemicolon) - 1))
    End If
End Function


Private Function TrimLeadingWhitespace(strSql As String) As String
    TrimLeadingWhitespace = LTrim$(strSql)
End Function


Private Function IsUnionQuery(strUpper As String) As Boolean
    ' Check if the query contains UNION outside of quotes
    Dim lngPos As Long
    Dim blnInString As Boolean
    Dim lngDepth As Long

    For lngPos = 1 To Len(strUpper) - 5
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                blnInString = Not blnInString
            Case "("
                If Not blnInString Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString Then lngDepth = lngDepth - 1
            Case "U"
                If Not blnInString And lngDepth = 0 Then
                    If Mid$(strUpper, lngPos, 6) = "UNION " Then
                        IsUnionQuery = True
                        Exit Function
                    End If
                End If
        End Select
    Next lngPos
End Function


Private Function IsDDLQuery(strUpper As String) As Boolean
    If Left$(strUpper, 13) = "CREATE TABLE " Or _
       Left$(strUpper, 12) = "ALTER TABLE " Or _
       Left$(strUpper, 11) = "DROP TABLE " Or _
       Left$(strUpper, 13) = "CREATE INDEX " Or _
       Left$(strUpper, 11) = "DROP INDEX " Then
        IsDDLQuery = True
    End If
End Function


Private Function HasSubqueries() As Boolean
    Dim dTbl As Dictionary
    For Each dTbl In m_colInputTables
        If InStr(1, CStr(dTbl("name")), "(") > 0 Or _
           InStr(1, CStr(dTbl("name")), "SELECT ", vbTextCompare) > 0 Then
            HasSubqueries = True
            Exit Function
        End If
    Next dTbl
End Function


Private Sub ParseParametersClause(ByRef strSql As String)
    Dim lngEnd As Long
    lngEnd = InStr(1, strSql, ";")
    If lngEnd > 0 Then
        m_strParameters = Mid$(strSql, 12, lngEnd - 12)
        strSql = LTrim$(Mid$(strSql, lngEnd + 1))
    End If
End Sub


Private Sub ParseTransformClause(ByRef strSql As String)
    Dim strUpper As String
    Dim lngSelectPos As Long
    strUpper = UCase$(strSql)
    lngSelectPos = FindKeyword(strUpper, "SELECT ", 11)
    If lngSelectPos > 0 Then
        m_strTransformAgg = Trim$(Mid$(strSql, 11, lngSelectPos - 11))
        strSql = Mid$(strSql, lngSelectPos)
    End If
End Sub


Private Sub ParseSelectQuery(strSql As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strRemainder As String
    Dim strUpper As String
    Dim lngPos As Long

    strRemainder = Mid$(strSql, 8) ' Skip "SELECT "
    strUpper = UCase$(strRemainder)

    ' Check DISTINCTROW (must check before DISTINCT since it's a longer match)
    If Left$(strUpper, 12) = "DISTINCTROW " Then
        m_blnDistinct = True
        m_lngOption = m_lngOption Or 8
        strRemainder = Mid$(strRemainder, 13)
        strUpper = Mid$(strUpper, 13)
    ElseIf Left$(strUpper, 9) = "DISTINCT " Then
        m_blnDistinct = True
        m_lngOption = m_lngOption Or 2
        strRemainder = Mid$(strRemainder, 10)
        strUpper = Mid$(strUpper, 10)
    End If

    ' Check TOP N
    If Left$(strUpper, 4) = "TOP " Then
        ParseTopClause strRemainder
        strUpper = UCase$(strRemainder)
    End If

    ' Find FROM keyword to split fields from source
    lngPos = FindTopLevelKeyword(strRemainder, "FROM ")
    If lngPos > 0 Then
        Dim strFieldSection As String
        strFieldSection = Left$(strRemainder, lngPos - 1)

        ' Check for INTO clause (Make-Table query: SELECT ... INTO table FROM ...)
        Dim lngInto As Long
        lngInto = FindTopLevelKeyword(strFieldSection, " INTO ")
        If lngInto > 0 Then
            m_lngQueryType = 2
            m_strTargetTable = UnBracket(Trim$(Mid$(strFieldSection, lngInto + 6)))
            strFieldSection = Left$(strFieldSection, lngInto - 1)
        End If

        ParseFieldList strFieldSection
        strRemainder = Mid$(strRemainder, lngPos + 5)
        ParseFromAndClauses strRemainder
    End If

    ' Parse PIVOT if present (for crosstab queries already stripped of TRANSFORM)
    If Len(m_strTransformAgg) > 0 Then
        ParsePivotFromOrderBy
    End If

    CatchAny eelError, "Error parsing SELECT query. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseSelectQuery"
End Sub


Private Sub ParseInsertQuery(strSql As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strRemainder As String
    Dim strUpper As String
    Dim lngSelectPos As Long
    Dim lngParenOpen As Long
    Dim lngParenClose As Long
    Dim asTargetFields() As String
    Dim blnHasTargetFields As Boolean

    strRemainder = Mid$(strSql, 13) ' Skip "INSERT INTO "

    ' Extract target table
    strUpper = UCase$(strRemainder)
    lngSelectPos = FindTopLevelKeyword(strRemainder, "SELECT ")
    lngParenOpen = InStr(1, strRemainder, "(")

    ' Check if there are explicit target field names before SELECT
    If lngParenOpen > 0 And (lngSelectPos = 0 Or lngParenOpen < lngSelectPos) Then
        m_strTargetTable = Trim$(Left$(strRemainder, lngParenOpen - 1))
        lngParenClose = FindMatchingParen(strRemainder, lngParenOpen)
        ' Extract the target field names between the parens
        Dim strTargetFieldList As String
        strTargetFieldList = Mid$(strRemainder, lngParenOpen + 1, lngParenClose - lngParenOpen - 1)
        asTargetFields = SplitTopLevel(Trim$(strTargetFieldList), ",")
        blnHasTargetFields = True
        strRemainder = Trim$(Mid$(strRemainder, lngParenClose + 1))
    Else
        If lngSelectPos > 0 Then
            m_strTargetTable = Trim$(Left$(strRemainder, lngSelectPos - 1))
            strRemainder = Mid$(strRemainder, lngSelectPos)
        Else
            m_strTargetTable = Trim$(strRemainder)
            GoTo CleanUp
        End If
    End If

    ' Parse the SELECT portion
    strUpper = UCase$(strRemainder)
    If Left$(strUpper, 7) = "SELECT " Then
        strRemainder = Mid$(strRemainder, 8)
        strUpper = UCase$(strRemainder)
        ' Check DISTINCTROW before DISTINCT
        If Left$(strUpper, 12) = "DISTINCTROW " Then
            m_blnDistinct = True
            m_lngOption = m_lngOption Or 8
            strRemainder = Mid$(strRemainder, 13)
        ElseIf Left$(strUpper, 9) = "DISTINCT " Then
            m_blnDistinct = True
            m_lngOption = m_lngOption Or 2
            strRemainder = Mid$(strRemainder, 10)
        End If
        Dim lngFrom As Long
        lngFrom = FindTopLevelKeyword(strRemainder, "FROM ")
        If lngFrom > 0 Then
            ParseFieldList Left$(strRemainder, lngFrom - 1)
            strRemainder = Mid$(strRemainder, lngFrom + 5)
            ParseFromAndClauses strRemainder
        End If
    End If

    ' Map INSERT target field names to output columns by position
    If blnHasTargetFields Then
        Dim lngCol As Long
        For lngCol = 1 To m_colOutputColumns.Count
            If lngCol - 1 <= UBound(asTargetFields) Then
                m_colOutputColumns(lngCol)("targetName") = UnBracket(Trim$(asTargetFields(lngCol - 1)))
            End If
        Next lngCol
    End If

CleanUp:
    CatchAny eelError, "Error parsing INSERT query. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseInsertQuery"
End Sub


Private Sub ParseUpdateQuery(strSql As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strRemainder As String
    Dim lngSet As Long

    strRemainder = Mid$(strSql, 8) ' Skip "UPDATE "

    ' Find SET keyword
    lngSet = FindTopLevelKeyword(strRemainder, "SET ")
    If lngSet > 0 Then
        ' FROM clause is before SET
        ParseFromClause Left$(strRemainder, lngSet - 1)
        strRemainder = Mid$(strRemainder, lngSet + 4)

        ' Find WHERE
        Dim lngWhere As Long
        lngWhere = FindTopLevelKeyword(strRemainder, "WHERE ")
        If lngWhere > 0 Then
            ParseUpdateSetFields Left$(strRemainder, lngWhere - 1)
            m_strWhere = WrapAccessWhere(Trim$(Mid$(strRemainder, lngWhere + 6)))
        Else
            ParseUpdateSetFields strRemainder
        End If
    End If

    CatchAny eelError, "Error parsing UPDATE query. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseUpdateQuery"
End Sub


Private Sub ParseDeleteQuery(strSql As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strRemainder As String
    Dim lngFrom As Long

    strRemainder = Mid$(strSql, 8) ' Skip "DELETE "

    lngFrom = FindTopLevelKeyword(strRemainder, "FROM ")
    If lngFrom > 0 Then
        ' Delete field list (typically "table.*")
        ParseFieldList Left$(strRemainder, lngFrom - 1)
        strRemainder = Mid$(strRemainder, lngFrom + 5)
        ParseFromAndClauses strRemainder
    End If

    CatchAny eelError, "Error parsing DELETE query. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseDeleteQuery"
End Sub


Private Sub ParseFieldList(strFields As String)
    Dim asParts() As String
    Dim lngF As Long
    Dim strField As String
    Dim strAlias As String
    Dim lngAs As Long
    Dim dCol As Dictionary

    asParts = SplitTopLevel(Trim$(strFields), ",")
    For lngF = 0 To UBound(asParts)
        strField = Trim$(asParts(lngF))
        If Len(strField) = 0 Then GoTo NextField

        Set dCol = New Dictionary

        ' Check for AS alias
        lngAs = FindTopLevelKeyword(strField, " AS ")
        If lngAs > 0 Then
            strAlias = Trim$(Mid$(strField, lngAs + 4))
            strField = Trim$(Left$(strField, lngAs - 1))
            ' Remove brackets from alias if present
            If Left$(strAlias, 1) = "[" And Right$(strAlias, 1) = "]" Then
                strAlias = Mid$(strAlias, 2, Len(strAlias) - 2)
            End If
            dCol.Add "alias", strAlias
        Else
            dCol.Add "alias", vbNullString
        End If
        dCol.Add "expression", strField
        dCol.Add "targetName", vbNullString
        m_colOutputColumns.Add dCol
NextField:
    Next lngF
End Sub


Private Sub ParseUpdateSetFields(strSetClause As String)
    Dim asParts() As String
    Dim lngF As Long
    Dim strPart As String
    Dim lngEq As Long
    Dim dCol As Dictionary

    asParts = SplitTopLevel(Trim$(strSetClause), ",")
    For lngF = 0 To UBound(asParts)
        strPart = Trim$(asParts(lngF))
        lngEq = InStr(1, strPart, "=")
        If lngEq > 0 Then
            Set dCol = New Dictionary
            Dim strTargetField As String
            strTargetField = Trim$(Left$(strPart, lngEq - 1))
            ' Retain fully-qualified table.field form, just unbracket each part
            Dim lngDot As Long
            lngDot = InStrRev(strTargetField, ".")
            If lngDot > 0 Then
                dCol.Add "targetName", UnBracket(Left$(strTargetField, lngDot - 1)) & _
                    "." & UnBracket(Mid$(strTargetField, lngDot + 1))
            Else
                dCol.Add "targetName", UnBracket(strTargetField)
            End If
            dCol.Add "expression", Trim$(Mid$(strPart, lngEq + 1))
            dCol.Add "alias", vbNullString
            m_colOutputColumns.Add dCol
        End If
    Next lngF
End Sub


Private Sub ParseFromClause(strFrom As String)
    ' Parse FROM clause into tables and joins
    Dim strClean As String
    strClean = Trim$(strFrom)
    ParseFromExpression strClean
End Sub


Private Sub ParseFromAndClauses(strAfterFrom As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strUpper As String
    Dim lngWhere As Long
    Dim lngGroup As Long
    Dim lngHaving As Long
    Dim lngOrder As Long
    Dim lngPivot As Long
    Dim strFrom As String
    Dim strRemainder As String

    strUpper = UCase$(strAfterFrom)

    ' Find each clause boundary
    lngWhere = FindTopLevelKeyword(strAfterFrom, "WHERE ")
    lngGroup = FindTopLevelKeyword(strAfterFrom, "GROUP BY ")
    lngHaving = FindTopLevelKeyword(strAfterFrom, "HAVING ")
    lngOrder = FindTopLevelKeyword(strAfterFrom, "ORDER BY ")
    lngPivot = FindTopLevelKeyword(strAfterFrom, "PIVOT ")

    ' Find the end of the FROM clause
    Dim lngFromEnd As Long
    lngFromEnd = Len(strAfterFrom) + 1
    If lngWhere > 0 And lngWhere < lngFromEnd Then lngFromEnd = lngWhere
    If lngGroup > 0 And lngGroup < lngFromEnd Then lngFromEnd = lngGroup
    If lngHaving > 0 And lngHaving < lngFromEnd Then lngFromEnd = lngHaving
    If lngOrder > 0 And lngOrder < lngFromEnd Then lngFromEnd = lngOrder
    If lngPivot > 0 And lngPivot < lngFromEnd Then lngFromEnd = lngPivot

    strFrom = Trim$(Left$(strAfterFrom, lngFromEnd - 1))
    ParseFromExpression strFrom

    ' WHERE
    If lngWhere > 0 Then
        Dim lngWhereEnd As Long
        lngWhereEnd = Len(strAfterFrom) + 1
        If lngGroup > lngWhere And lngGroup < lngWhereEnd Then lngWhereEnd = lngGroup
        If lngHaving > lngWhere And lngHaving < lngWhereEnd Then lngWhereEnd = lngHaving
        If lngOrder > lngWhere And lngOrder < lngWhereEnd Then lngWhereEnd = lngOrder
        If lngPivot > lngWhere And lngPivot < lngWhereEnd Then lngWhereEnd = lngPivot
        m_strWhere = WrapAccessWhere(Trim$(Mid$(strAfterFrom, lngWhere + 6, lngWhereEnd - lngWhere - 6)))
    End If

    ' GROUP BY
    If lngGroup > 0 Then
        Dim lngGroupEnd As Long
        lngGroupEnd = Len(strAfterFrom) + 1
        If lngHaving > lngGroup And lngHaving < lngGroupEnd Then lngGroupEnd = lngHaving
        If lngOrder > lngGroup And lngOrder < lngGroupEnd Then lngGroupEnd = lngOrder
        If lngPivot > lngGroup And lngPivot < lngGroupEnd Then lngGroupEnd = lngPivot
        ParseGroupByClause Trim$(Mid$(strAfterFrom, lngGroup + 9, lngGroupEnd - lngGroup - 9))
    End If

    ' HAVING
    If lngHaving > 0 Then
        Dim lngHavingEnd As Long
        lngHavingEnd = Len(strAfterFrom) + 1
        If lngOrder > lngHaving And lngOrder < lngHavingEnd Then lngHavingEnd = lngOrder
        If lngPivot > lngHaving And lngPivot < lngHavingEnd Then lngHavingEnd = lngPivot
        m_strHaving = WrapAccessWhere(Trim$(Mid$(strAfterFrom, lngHaving + 7, lngHavingEnd - lngHaving - 7)))
    End If

    ' ORDER BY
    If lngOrder > 0 Then
        Dim lngOrderEnd As Long
        lngOrderEnd = Len(strAfterFrom) + 1
        If lngPivot > lngOrder And lngPivot < lngOrderEnd Then lngOrderEnd = lngPivot
        ParseOrderByClause Trim$(Mid$(strAfterFrom, lngOrder + 9, lngOrderEnd - lngOrder - 9))
    End If

    ' PIVOT
    If lngPivot > 0 Then
        m_strPivotField = Trim$(Mid$(strAfterFrom, lngPivot + 6))
    End If

    CatchAny eelError, "Error parsing FROM/WHERE/GROUP/HAVING/ORDER/PIVOT clauses. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseFromAndClauses"
End Sub


Private Sub ParseFromExpression(strFrom As String)
    ' Parse the FROM expression into tables and joins
    ' Handle Access-style nested parenthesized JOINs:
    '   FROM ((A INNER JOIN B ON ...) LEFT JOIN C ON ...)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strWork As String
    Dim strUpper As String
    Dim lngJoinPos As Long
    Dim strJoinType As String
    Dim lngFlag As Long

    strWork = Trim$(strFrom)

    ' Strip outer parentheses (Access wraps joins in multiple layers)
    Do While Left$(strWork, 1) = "(" And Right$(strWork, 1) = ")"
        If FindMatchingParen(strWork, 1) = Len(strWork) Then
            strWork = Mid$(strWork, 2, Len(strWork) - 2)
            strWork = Trim$(strWork)
        Else
            Exit Do
        End If
    Loop

    strUpper = UCase$(strWork)

    ' Check for JOIN keywords
    lngJoinPos = FindJoinKeyword(strWork, strJoinType, lngFlag)

    If lngJoinPos = 0 Then
        ' No joins -- parse as comma-separated table list
        Dim asTables() As String
        Dim lngT As Long
        asTables = SplitTopLevel(strWork, ",")
        For lngT = 0 To UBound(asTables)
            Dim strTblExpr As String
            strTblExpr = Trim$(asTables(lngT))
            If Len(strTblExpr) > 0 Then
                AddInputTable strTblExpr
            End If
        Next lngT
    Else
        ' Has joins -- parse recursively
        ParseJoinExpression strWork
    End If

    CatchAny eelError, "Error parsing FROM expression. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseFromExpression"
End Sub


Private Sub ParseJoinExpression(strExpr As String)
    ' Recursively parse join expressions

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strWork As String
    Dim strLeft As String
    Dim strRight As String
    Dim strOnExpr As String
    Dim strJoinType As String
    Dim lngFlag As Long
    Dim lngJoinPos As Long
    Dim lngOnPos As Long

    strWork = Trim$(strExpr)

    ' Strip outer parens
    Do While Left$(strWork, 1) = "(" And Right$(strWork, 1) = ")"
        If FindMatchingParen(strWork, 1) = Len(strWork) Then
            strWork = Mid$(strWork, 2, Len(strWork) - 2)
            strWork = Trim$(strWork)
        Else
            Exit Do
        End If
    Loop

    ' Find the LAST join at the top level (rightmost)
    lngJoinPos = FindLastJoinKeyword(strWork, strJoinType, lngFlag)

    If lngJoinPos = 0 Then
        ' Base case: single table
        AddInputTable strWork
        GoTo CleanUp
    End If

    ' Split at this join
    strLeft = Trim$(Left$(strWork, lngJoinPos - 1))
    strRight = Trim$(Mid$(strWork, lngJoinPos + Len(strJoinType)))

    ' Find ON clause in the right part
    lngOnPos = FindTopLevelKeyword(strRight, " ON ")
    If lngOnPos > 0 Then
        strOnExpr = Trim$(Mid$(strRight, lngOnPos + 4))
        strRight = Trim$(Left$(strRight, lngOnPos - 1))
    End If

    ' Recursively parse the left side (may have nested joins)
    ParseJoinExpression strLeft

    ' Recursively parse the right side (may also have nested joins
    ' in right-nested INNER JOIN chains)
    ParseJoinExpression strRight

    ' Record the join
    Dim dJoin As Dictionary
    Dim strLeftTable As String
    Dim strRightTable As String

    ' Extract table names from the ON expression
    strLeftTable = ExtractJoinLeftTable(strOnExpr)
    strRightTable = ExtractJoinRightTable(strOnExpr, strRight, strLeftTable)

    Set dJoin = New Dictionary
    dJoin.Add "leftTable", strLeftTable
    dJoin.Add "rightTable", strRightTable
    dJoin.Add "expression", strOnExpr
    dJoin.Add "flag", lngFlag
    m_colJoins.Add dJoin

CleanUp:
    CatchAny eelError, "Error parsing JOIN expression. " & SqlSnippet(), _
        ModuleName(Me) & ".ParseJoinExpression"
End Sub


Private Sub AddInputTable(strTableExpr As String)
    Dim dTbl As Dictionary
    Dim strName As String
    Dim strAlias As String
    Dim lngAs As Long

    strName = Trim$(strTableExpr)

    ' Check for AS alias
    lngAs = FindTopLevelKeyword(strName, " AS ")
    If lngAs > 0 Then
        strAlias = Trim$(Mid$(strName, lngAs + 4))
        strName = Trim$(Left$(strName, lngAs - 1))
    Else
        strAlias = vbNullString
    End If

    ' Remove brackets
    strName = UnBracket(strName)
    If Len(strAlias) > 0 Then strAlias = UnBracket(strAlias)

    Set dTbl = New Dictionary
    dTbl.Add "name", strName
    dTbl.Add "alias", strAlias
    m_colInputTables.Add dTbl
End Sub


Private Sub ParseGroupByClause(strGroupBy As String)
    Dim asParts() As String
    Dim lngG As Long
    Dim dGrp As Dictionary

    asParts = SplitTopLevel(Trim$(strGroupBy), ",")
    For lngG = 0 To UBound(asParts)
        If Len(Trim$(asParts(lngG))) > 0 Then
            Set dGrp = New Dictionary
            dGrp.Add "expression", Trim$(asParts(lngG))
            dGrp.Add "groupLevel", 0
            m_colGroups.Add dGrp
        End If
    Next lngG
End Sub


Private Sub ParseOrderByClause(strOrderBy As String)
    Dim asParts() As String
    Dim lngO As Long
    Dim dOrd As Dictionary
    Dim strPart As String
    Dim blnDesc As Boolean

    asParts = SplitTopLevel(Trim$(strOrderBy), ",")
    For lngO = 0 To UBound(asParts)
        strPart = Trim$(asParts(lngO))
        If Len(strPart) > 0 Then
            blnDesc = False
            If Right$(UCase$(strPart), 5) = " DESC" Then
                blnDesc = True
                strPart = Trim$(Left$(strPart, Len(strPart) - 5))
            ElseIf Right$(UCase$(strPart), 4) = " ASC" Then
                strPart = Trim$(Left$(strPart, Len(strPart) - 4))
            End If
            Set dOrd = New Dictionary
            dOrd.Add "expression", strPart
            dOrd.Add "descending", blnDesc
            m_colOrderBy.Add dOrd
        End If
    Next lngO
End Sub


Private Sub ParsePivotFromOrderBy()
    ' For TRANSFORM queries, PIVOT comes after ORDER BY
    ' It may have already been parsed in ParseFromAndClauses
    ' No additional action needed here
End Sub


'=======================================================================================
' PRIVATE: Parsing utility functions
'=======================================================================================

Private Function FindTopLevelKeyword(strSql As String, strKeyword As String) As Long
    ' Find keyword position at the top level
    ' (not inside parentheses, double-quoted strings, or bracketed identifiers)
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = 1 To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                If Not blnInBracket Then blnInString = Not blnInString
            Case "["
                If Not blnInString Then blnInBracket = True
            Case "]"
                If Not blnInString Then blnInBracket = False
            Case "("
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth - 1
        End Select

        If Not blnInString And Not blnInBracket And lngDepth = 0 Then
            If Mid$(strUpper, lngPos, Len(strKeyUpper)) = strKeyUpper Then
                ' Verify word boundary before keyword. Keywords that start
                ' with a space (like " LEFT JOIN ") already carry their own
                ' boundary, so no extra check is needed.
                If Left$(strKeyUpper, 1) = " " Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                ElseIf lngPos = 1 Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                ElseIf Mid$(strUpper, lngPos - 1, 1) = " " Or _
                   Mid$(strUpper, lngPos - 1, 1) = ")" Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                End If
            End If
        End If
    Next lngPos
End Function


Private Function FindKeyword(strUpper As String, strKeyword As String, lngStart As Long) As Long
    ' Simple keyword find from a starting position
    FindKeyword = InStr(lngStart, strUpper, strKeyword, vbTextCompare)
End Function


Private Function FindJoinKeyword(strExpr As String, ByRef strJoinType As String, ByRef lngFlag As Long) As Long
    ' Find the first JOIN keyword at the top level
    Dim lngPos As Long

    lngPos = FindTopLevelKeyword(strExpr, " INNER JOIN ")
    If lngPos > 0 Then
        strJoinType = " INNER JOIN "
        lngFlag = 1
        FindJoinKeyword = lngPos
        Exit Function
    End If

    lngPos = FindTopLevelKeyword(strExpr, " LEFT JOIN ")
    If lngPos > 0 Then
        strJoinType = " LEFT JOIN "
        lngFlag = 2
        FindJoinKeyword = lngPos
        Exit Function
    End If

    lngPos = FindTopLevelKeyword(strExpr, " RIGHT JOIN ")
    If lngPos > 0 Then
        strJoinType = " RIGHT JOIN "
        lngFlag = 3
        FindJoinKeyword = lngPos
        Exit Function
    End If
End Function


Private Function FindLastJoinKeyword(strExpr As String, ByRef strJoinType As String, ByRef lngFlag As Long) As Long
    ' Find the last top-level JOIN keyword
    Dim lngPos As Long
    Dim lngLast As Long
    Dim strLastType As String
    Dim lngLastFlag As Long
    Dim lngSearch As Long
    Dim strTmpType As String
    Dim lngTmpFlag As Long
    Dim strUpper As String

    strUpper = UCase$(strExpr)
    lngSearch = 1
    lngLast = 0

    Do
        lngPos = 0
        ' Check each join type from current search position
        Dim lngInner As Long, lngLeft As Long, lngRight As Long
        lngInner = FindTopLevelKeywordFrom(strExpr, " INNER JOIN ", lngSearch)
        lngLeft = FindTopLevelKeywordFrom(strExpr, " LEFT JOIN ", lngSearch)
        lngRight = FindTopLevelKeywordFrom(strExpr, " RIGHT JOIN ", lngSearch)

        ' Find the nearest one
        If lngInner > 0 Then
            lngPos = lngInner: strTmpType = " INNER JOIN ": lngTmpFlag = 1
        End If
        If lngLeft > 0 And (lngPos = 0 Or lngLeft < lngPos) Then
            lngPos = lngLeft: strTmpType = " LEFT JOIN ": lngTmpFlag = 2
        End If
        If lngRight > 0 And (lngPos = 0 Or lngRight < lngPos) Then
            lngPos = lngRight: strTmpType = " RIGHT JOIN ": lngTmpFlag = 3
        End If

        If lngPos > 0 Then
            lngLast = lngPos
            strLastType = strTmpType
            lngLastFlag = lngTmpFlag
            lngSearch = lngPos + Len(strTmpType)
        End If
    Loop While lngPos > 0

    If lngLast > 0 Then
        FindLastJoinKeyword = lngLast
        strJoinType = strLastType
        lngFlag = lngLastFlag
    End If
End Function


Private Function FindTopLevelKeywordFrom(strSql As String, strKeyword As String, lngStart As Long) As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = lngStart To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                blnInString = Not blnInString
            Case "("
                If Not blnInString Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString Then lngDepth = lngDepth - 1
        End Select

        If Not blnInString And lngDepth = 0 Then
            If Mid$(strUpper, lngPos, Len(strKeyUpper)) = strKeyUpper Then
                FindTopLevelKeywordFrom = lngPos
                Exit Function
            End If
        End If
    Next lngPos
End Function


Private Function FindMatchingParen(strExpr As String, lngStart As Long) As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean

    lngDepth = 0
    For lngPos = lngStart To Len(strExpr)
        Select Case Mid$(strExpr, lngPos, 1)
            Case """"
                blnInString = Not blnInString
            Case "("
                If Not blnInString Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString Then
                    lngDepth = lngDepth - 1
                    If lngDepth = 0 Then
                        FindMatchingParen = lngPos
                        Exit Function
                    End If
                End If
        End Select
    Next lngPos
End Function


Private Function SplitTopLevel(strExpr As String, strDelim As String) As String()
    ' Split a string by delimiter, but only at the top level
    ' (not inside parentheses, double-quoted strings, or bracketed identifiers)
    Dim colParts As New Collection
    Dim lngStart As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean
    Dim asParts() As String

    lngStart = 1
    For lngPos = 1 To Len(strExpr)
        Select Case Mid$(strExpr, lngPos, 1)
            Case """"
                If Not blnInBracket Then blnInString = Not blnInString
            Case "["
                If Not blnInString Then blnInBracket = True
            Case "]"
                If Not blnInString Then blnInBracket = False
            Case "("
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth - 1
            Case Else
                If Mid$(strExpr, lngPos, Len(strDelim)) = strDelim Then
                    If Not blnInString And Not blnInBracket And lngDepth = 0 Then
                        colParts.Add Mid$(strExpr, lngStart, lngPos - lngStart)
                        lngStart = lngPos + Len(strDelim)
                    End If
                End If
        End Select
    Next lngPos

    ' Add the last part
    colParts.Add Mid$(strExpr, lngStart)

    ' Convert to array
    ReDim asParts(0 To colParts.Count - 1)
    Dim lngI As Long
    For lngI = 1 To colParts.Count
        asParts(lngI - 1) = CStr(colParts(lngI))
    Next lngI

    SplitTopLevel = asParts
End Function


Private Function UnBracket(strName As String) As String
    strName = Trim$(strName)
    If Left$(strName, 1) = "[" And Right$(strName, 1) = "]" Then
        UnBracket = Mid$(strName, 2, Len(strName) - 2)
    Else
        UnBracket = strName
    End If
End Function


Private Function ExtractJoinLeftTable(strOnExpr As String) As String
    ' Extract the left table name from an ON expression like "A.col = B.col"
    Dim lngDot As Long
    Dim strLeft As String
    strLeft = Trim$(strOnExpr)

    ' Remove outer parens (and any whitespace exposed after them)
    Do While Left$(strLeft, 1) = "("
        strLeft = Mid$(strLeft, 2)
    Loop
    strLeft = Trim$(strLeft)

    lngDot = InStr(1, strLeft, ".")
    If lngDot > 0 Then
        ExtractJoinLeftTable = UnBracket(Left$(strLeft, lngDot - 1))
    Else
        ExtractJoinLeftTable = strLeft
    End If
End Function


Private Function ExtractJoinRightTable(strOnExpr As String, strTableExpr As String, _
                                       strLeftTable As String) As String
    ' Extract the right table name from the join context.
    ' For simple table references, extract the name directly from strTableExpr.
    ' For nested join expressions, extract both table refs from the ON clause
    ' and return whichever is NOT the left table.
    Dim strClean As String
    strClean = Trim$(strTableExpr)

    ' Strip leading/trailing parens that wrap the entire expression
    Do While Left$(strClean, 1) = "(" And Right$(strClean, 1) = ")"
        If FindMatchingParen(strClean, 1) = Len(strClean) Then
            strClean = Trim$(Mid$(strClean, 2, Len(strClean) - 2))
        Else
            Exit Do
        End If
    Loop

    strClean = UnBracket(strClean)

    If InStr(1, strClean, " JOIN ", vbTextCompare) = 0 Then
        ' Simple table reference (possibly with alias)
        Dim lngSpace As Long
        lngSpace = InStr(1, strClean, " ", vbTextCompare)
        If lngSpace > 0 Then
            ExtractJoinRightTable = Trim$(Left$(strClean, lngSpace - 1))
        Else
            ExtractJoinRightTable = strClean
        End If
        Exit Function
    End If

    ' Nested join expression: the right table is one of the two tables
    ' referenced in the ON clause. Extract both and return the one that
    ' is NOT the left table.
    Dim strTblA As String
    Dim strTblB As String
    strTblA = ExtractTableFromOnSide(strOnExpr, True)
    strTblB = ExtractTableFromOnSide(strOnExpr, False)

    If StrComp(strTblA, strLeftTable, vbTextCompare) = 0 Then
        ExtractJoinRightTable = strTblB
    Else
        ExtractJoinRightTable = strTblA
    End If
End Function


Private Function ExtractTableFromOnSide(strOnExpr As String, blnLeftSide As Boolean) As String
    ' Extract a table name from one side of the first '=' in an ON expression.
    ' blnLeftSide=True returns the table before '=', False returns after '='.
    Dim lngEq As Long
    Dim strSide As String

    lngEq = InStr(1, strOnExpr, "=")
    If lngEq = 0 Then Exit Function

    If blnLeftSide Then
        strSide = Trim$(Left$(strOnExpr, lngEq - 1))
    Else
        strSide = Trim$(Mid$(strOnExpr, lngEq + 1))
    End If

    ' Strip outer parens
    Do While Left$(strSide, 1) = "("
        strSide = Mid$(strSide, 2)
    Loop
    Do While Right$(strSide, 1) = ")"
        strSide = Left$(strSide, Len(strSide) - 1)
    Loop
    strSide = Trim$(strSide)

    Dim lngDot As Long
    lngDot = InStr(1, strSide, ".")
    If lngDot > 0 Then
        ExtractTableFromOnSide = UnBracket(Trim$(Left$(strSide, lngDot - 1)))
    Else
        ExtractTableFromOnSide = UnBracket(strSide)
    End If
End Function


Private Function WrapAccessWhere(strCondition As String) As String
    ' Access wraps WHERE/HAVING in double parentheses: (((expr)))
    ' Compact the expression and wrap if not already wrapped.
    strCondition = CompactSql(Trim$(strCondition))
    If Left$(strCondition, 1) <> "(" Then
        WrapAccessWhere = "(((" & strCondition & ")))"
    Else
        WrapAccessWhere = strCondition
    End If
End Function


Private Function CompactSql(strExpr As String) As String
    ' Strip formatter-added spaces to match native .qdef compact formatting.
    ' Removes spaces after "(", before ")", and around operators
    ' (=, <>, <, >, <=, >=, *, +, -, /). Preserves spaces inside string
    ' literals and bracketed identifiers.
    Dim lngI As Long
    Dim lngLen As Long
    Dim strChar As String
    Dim cOut As New clsConcat
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean

    lngLen = Len(strExpr)
    If lngLen = 0 Then Exit Function

    For lngI = 1 To lngLen
        strChar = Mid$(strExpr, lngI, 1)

        If strChar = """" And Not blnInBracket Then blnInString = Not blnInString
        If strChar = "[" And Not blnInString Then blnInBracket = True
        If strChar = "]" And Not blnInString Then blnInBracket = False

        If blnInString Or blnInBracket Then
            cOut.Add strChar
        ElseIf strChar <> " " Then
            cOut.Add strChar
        Else
            ' This is a space outside a string literal.
            ' Check the surrounding characters to decide whether to keep it.
            Dim strPrev As String
            Dim strNext As String
            If lngI > 1 Then strPrev = Mid$(strExpr, lngI - 1, 1) Else strPrev = vbNullString
            If lngI < lngLen Then strNext = Mid$(strExpr, lngI + 1, 1) Else strNext = vbNullString

            If strPrev = "(" Or strNext = ")" Then
                ' Drop space after ( or before )
            ElseIf IsOperatorChar(strPrev) Or IsOperatorChar(strNext) Then
                ' Drop space adjacent to operators
            Else
                cOut.Add strChar
            End If
        End If
    Next lngI

    CompactSql = cOut.GetStr
End Function


Private Function IsOperatorChar(strCh As String) As Boolean
    IsOperatorChar = (strCh = "=" Or strCh = "<" Or strCh = ">" Or _
                      strCh = "*" Or strCh = "+" Or strCh = "-" Or strCh = "/")
End Function


Private Sub ParseTopClause(ByRef strRemainder As String)
    Dim lngSpace As Long
    Dim strTop As String

    strRemainder = Mid$(strRemainder, 5) ' Skip "TOP "
    lngSpace = InStr(1, strRemainder, " ")
    If lngSpace > 0 Then
        strTop = Left$(strRemainder, lngSpace - 1)
        strRemainder = Mid$(strRemainder, lngSpace + 1)
        If UCase$(Left$(strRemainder, 8)) = "PERCENT " Then
            m_lngTopPercent = CLng(strTop)
            m_lngOption = m_lngOption Or 48  ' TOP + PERCENT
            strRemainder = Mid$(strRemainder, 9)
        Else
            m_lngTopN = CLng(strTop)
            m_lngOption = m_lngOption Or 16  ' TOP
        End If
    End If
End Sub


'=======================================================================================
' PRIVATE: .qdef generation (Design View and SQL View formats)
'=======================================================================================

Private Sub EmitDesignViewQdef(cOut As clsConcat, dProperties As Dictionary, _
                               dColumns As Dictionary, dLayout As Dictionary)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Operation line
    cOut.Add "Operation =" & CStr(m_lngQueryType), vbCrLf

    ' Target table for action queries
    If Len(m_strTargetTable) > 0 Then
        cOut.Add "Name =""" & m_strTargetTable & """", vbCrLf
    End If

    ' Option
    cOut.Add "Option =" & CStr(m_lngOption), vbCrLf

    ' WHERE
    If Len(m_strWhere) > 0 Then
        EmitWrappedString cOut, "Where", m_strWhere
    End If

    ' HAVING
    If Len(m_strHaving) > 0 Then
        EmitWrappedString cOut, "Having", m_strHaving
    End If

    ' Input Tables (order preserved from SQL, which reflects MSysQueries.[Order])
    ' Each table emits a Name = line, optionally followed by an Alias = line
    ' when the SQL aliased it (e.g. "tbl AS tbl_1" in a self-join). Without
    ' the Alias line, Access stores both Attribute 5 rows with null Name2 and
    ' the Attribute 7 join rows collapse to leftTable == rightTable, which on
    ' the next export crashes BuildJoinChain (no graph root). This bit
    ' affected ~723 queries in db-sec alone, including qryBucketCombineRandUCos
    ' and qdelTempInvByQtrSubMergerAcquisition3.
    If m_colInputTables.Count > 0 Then
        cOut.Add "Begin InputTables", vbCrLf
        Dim dTbl As Dictionary
        For Each dTbl In m_colInputTables
            cOut.Add "    Name =""" & CStr(dTbl("name")) & """", vbCrLf
            If Len(Nz(dTbl("alias"), vbNullString)) > 0 Then
                cOut.Add "    Alias =""" & CStr(dTbl("alias")) & """", vbCrLf
            End If
        Next dTbl
        cOut.Add "End", vbCrLf
    End If

    ' Output Columns
    If m_colOutputColumns.Count > 0 Then
        cOut.Add "Begin OutputColumns", vbCrLf
        Dim dCol As Dictionary
        For Each dCol In m_colOutputColumns
            If Len(dCol("alias")) > 0 Then
                cOut.Add "    Alias =""" & CStr(dCol("alias")) & """", vbCrLf
            End If
            If Len(dCol("targetName")) > 0 Then
                cOut.Add "    Name =""" & CStr(dCol("targetName")) & """", vbCrLf
            End If
            EmitWrappedString cOut, "    Expression", CompactSql(CStr(dCol("expression")))
        Next dCol
        cOut.Add "End", vbCrLf
    End If

    ' Joins
    If m_colJoins.Count > 0 Then
        cOut.Add "Begin Joins", vbCrLf
        Dim dJn As Dictionary
        For Each dJn In m_colJoins
            cOut.Add "    LeftTable =""" & CStr(dJn("leftTable")) & """", vbCrLf
            cOut.Add "    RightTable =""" & CStr(dJn("rightTable")) & """", vbCrLf
            EmitWrappedString cOut, "    Expression", CStr(dJn("expression"))
            cOut.Add "    Flag =" & CStr(dJn("flag")), vbCrLf
        Next dJn
        cOut.Add "End", vbCrLf
    End If

    ' ORDER BY
    If m_colOrderBy.Count > 0 Then
        cOut.Add "Begin OrderBy", vbCrLf
        Dim dOrd As Dictionary
        For Each dOrd In m_colOrderBy
            EmitWrappedString cOut, "    Expression", CStr(dOrd("expression"))
            If dOrd("descending") Then
                cOut.Add "    Flag =1", vbCrLf
            Else
                cOut.Add "    Flag =0", vbCrLf
            End If
        Next dOrd
        cOut.Add "End", vbCrLf
    End If

    ' GROUP BY
    If m_colGroups.Count > 0 Then
        cOut.Add "Begin Groups", vbCrLf
        Dim dGrp As Dictionary
        For Each dGrp In m_colGroups
            EmitWrappedString cOut, "    Expression", CompactSql(CStr(dGrp("expression")))
            cOut.Add "    GroupLevel =" & CStr(dGrp("groupLevel")), vbCrLf
        Next dGrp
        cOut.Add "End", vbCrLf
    End If

    ' Properties -- these act as the delimiter between the query structure
    ' blocks and the Begin...End blocks (column metadata / layout).
    ' Always emit the full set of known properties with JSON overrides.
    EmitAllProperties cOut, dProperties

    ' Column metadata (synthesize from SQL if JSON doesn't have it)
    EmitColumnMetadata cOut, dColumns, True

    ' Layout
    If Not dLayout Is Nothing Then
        EmitDesignLayout cOut, dLayout
    End If

    CatchAny eelError, "Error emitting Design View qdef. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitDesignViewQdef"

End Sub


Private Sub EmitSqlViewQdef(cOut As clsConcat, dProperties As Dictionary, _
                            dColumns As Dictionary)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' dbMemo "SQL" = "..." with line wrapping
    EmitDbMemoSql cOut, m_strRawSql

    ' Connect string (always present in SQL View, even if empty)
    cOut.Add "dbMemo ""Connect"" =""""", vbCrLf

    ' Properties
    EmitAllProperties cOut, dProperties

    ' Column metadata
    EmitColumnMetadata cOut, dColumns

    CatchAny eelError, "Error emitting SQL View qdef. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitSqlViewQdef"

End Sub


Private Sub EmitDbMemoSql(cOut As clsConcat, strSql As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strEscaped As String
    Dim lngPos As Long
    Dim lngLineLen As Long
    Dim strLine As String
    Dim blnFirst As Boolean

    ' Escape the SQL for .qdef format:
    ' - Escape backslashes as \\ and double quotes as \" (in that order)
    ' - Replace CR+LF with \015\012 (added after escaping so the leading \ is literal)
    ' - Wrap in dbMemo "SQL" ="..." with continuation lines
    strEscaped = EscapeQdefString(strSql)
    strEscaped = Replace(strEscaped, vbCrLf, "\015\012")
    strEscaped = Replace(strEscaped, vbCr, "\015")
    strEscaped = Replace(strEscaped, vbLf, "\012")

    Const MAX_LINE As Long = 70
    blnFirst = True
    lngPos = 1

    Do While lngPos <= Len(strEscaped)
        If blnFirst Then
            lngLineLen = SafeBreak(strEscaped, lngPos, MAX_LINE)
            If lngPos + lngLineLen - 1 >= Len(strEscaped) Then
                strLine = Mid$(strEscaped, lngPos)
                cOut.Add "dbMemo ""SQL"" =""" & strLine & """", vbCrLf
                Exit Do
            Else
                strLine = Mid$(strEscaped, lngPos, lngLineLen)
                cOut.Add "dbMemo ""SQL"" =""" & strLine & """", vbCrLf
            End If
            blnFirst = False
        Else
            lngLineLen = SafeBreak(strEscaped, lngPos, MAX_LINE)
            If lngPos + lngLineLen - 1 >= Len(strEscaped) Then
                strLine = Mid$(strEscaped, lngPos)
                cOut.Add "    """ & strLine & """", vbCrLf
                Exit Do
            Else
                strLine = Mid$(strEscaped, lngPos, lngLineLen)
                cOut.Add "    """ & strLine & """", vbCrLf
            End If
        End If
        lngPos = lngPos + lngLineLen
    Loop

    CatchAny eelError, "Error emitting dbMemo SQL block. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitDbMemoSql"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SafeBreak
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Returns a safe line length that does not split an escape sequence
'           : (\\, \015, \012, \") across a line boundary. Access LoadFromText
'           : will misparse the file if a backslash escape spans two quoted
'           : segments.
'---------------------------------------------------------------------------------------
'
Private Function SafeBreak(strText As String, lngStart As Long, lngMaxLen As Long) As Long
    Dim lngEnd As Long
    Dim lngCheckPos As Long
    Dim lngI As Long
    Dim lngEscLen As Long

    lngEnd = lngStart + lngMaxLen - 1
    If lngEnd >= Len(strText) Then
        SafeBreak = Len(strText) - lngStart + 1
        Exit Function
    End If

    ' Scan backward from the break point looking for a backslash within
    ' the last 3 characters that could start an escape extending past
    ' the break. Escapes: \\ (2 chars), \" (2 chars), \015 (4 chars), \012 (4 chars).
    For lngI = 0 To 3
        lngCheckPos = lngEnd - lngI
        If lngCheckPos < lngStart Then Exit For

        If Mid$(strText, lngCheckPos, 1) = "\" Then
            lngEscLen = 1
            If lngCheckPos + 1 <= Len(strText) Then
                Dim strNext As String
                strNext = Mid$(strText, lngCheckPos + 1, 1)
                If strNext = """" Or strNext = "\" Then
                    lngEscLen = 2
                ElseIf lngCheckPos + 3 <= Len(strText) Then
                    Dim strSeq As String
                    strSeq = Mid$(strText, lngCheckPos + 1, 3)
                    If strSeq = "015" Or strSeq = "012" Then
                        lngEscLen = 4
                    End If
                End If
            End If

            If lngCheckPos + lngEscLen - 1 > lngEnd Then
                ' Escape extends past the break; move break before the backslash
                SafeBreak = lngCheckPos - lngStart
                If SafeBreak < 1 Then SafeBreak = lngMaxLen
                Exit Function
            Else
                Exit For
            End If
        End If
    Next lngI

    SafeBreak = lngMaxLen
End Function


Private Sub EmitAllProperties(cOut As clsConcat, dProperties As Dictionary)
    ' Emit the complete set of known query properties in native order.
    ' JSON values override defaults. Any extra JSON keys not in the
    ' known set are emitted afterward for future-proofing.

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim dOverrides As Dictionary
    Set dOverrides = ResolvePropertyOverrides(dProperties)

    ' Emit known properties in native order with defaults
    EmitBoolProp cOut, "ReturnsRecords", GetPropOrDefault(dOverrides, "ReturnsRecords", "-1")
    cOut.Add "dbInteger ""ODBCTimeout"" =""" & GetPropOrDefault(dOverrides, "ODBCTimeout", "60") & """", vbCrLf

    ' Action queries get UseTransaction + FailOnError instead of the datasheet properties.
    ' Make-Table (type 2) is a hybrid: full datasheet set plus UseTransaction.
    If m_lngQueryType = 3 Or m_lngQueryType = 4 Or m_lngQueryType = 5 Then
        EmitBoolProp cOut, "UseTransaction", GetPropOrDefault(dOverrides, "UseTransaction", "-1")
        ' FailOnError only appears in native .qdef for UPDATE and DELETE, not INSERT
        If m_lngQueryType = 4 Or m_lngQueryType = 5 Then
            EmitBoolProp cOut, "FailOnError", GetPropOrDefault(dOverrides, "FailOnError", "0")
        Else
            GetPropOrDefault dOverrides, "FailOnError", vbNullString
        End If
        cOut.Add "dbByte ""Orientation"" =""" & GetPropOrDefault(dOverrides, "Orientation", "0") & """", vbCrLf
        ' These are optional for action queries; emit only if provided
        Dim strVal As String
        strVal = GetPropOrDefault(dOverrides, "OrderByOn", vbNullString)
        If Len(strVal) > 0 Then EmitBoolProp cOut, "OrderByOn", strVal
        strVal = GetPropOrDefault(dOverrides, "DefaultView", vbNullString)
        If Len(strVal) > 0 Then cOut.Add "dbByte ""DefaultView"" =""" & strVal & """", vbCrLf
        strVal = GetPropOrDefault(dOverrides, "FilterOnLoad", vbNullString)
        If Len(strVal) > 0 Then EmitBoolProp cOut, "FilterOnLoad", strVal
        strVal = GetPropOrDefault(dOverrides, "OrderByOnLoad", vbNullString)
        If Len(strVal) > 0 Then EmitBoolProp cOut, "OrderByOnLoad", strVal
        ' Consume datasheet-only properties that don't apply to action queries
        GetPropOrDefault dOverrides, "RecordsetType", vbNullString
        GetPropOrDefault dOverrides, "TotalsRow", vbNullString
    Else
        ' SELECT and Make-Table queries get the full datasheet property set
        cOut.Add "dbByte ""RecordsetType"" =""" & GetPropOrDefault(dOverrides, "RecordsetType", "0") & """", vbCrLf
        EmitBoolProp cOut, "OrderByOn", GetPropOrDefault(dOverrides, "OrderByOn", "0")
        cOut.Add "dbByte ""Orientation"" =""" & GetPropOrDefault(dOverrides, "Orientation", "0") & """", vbCrLf
        cOut.Add "dbByte ""DefaultView"" =""" & GetPropOrDefault(dOverrides, "DefaultView", "2") & """", vbCrLf
        EmitBoolProp cOut, "FilterOnLoad", GetPropOrDefault(dOverrides, "FilterOnLoad", "0")
        EmitBoolProp cOut, "OrderByOnLoad", GetPropOrDefault(dOverrides, "OrderByOnLoad", "-1")
        EmitBoolProp cOut, "TotalsRow", GetPropOrDefault(dOverrides, "TotalsRow", "0")
        ' Make-Table queries also get UseTransaction after the datasheet properties
        If m_lngQueryType = 2 Then
            EmitBoolProp cOut, "UseTransaction", GetPropOrDefault(dOverrides, "UseTransaction", "-1")
        End If
        ' Consume action-query-only properties
        GetPropOrDefault dOverrides, "UseTransaction", vbNullString
        GetPropOrDefault dOverrides, "FailOnError", vbNullString
    End If

    ' Emit non-default extras that aren't in the known set
    Dim varKey As Variant
    For Each varKey In dOverrides.Keys
        Select Case CStr(varKey)
            Case "ReturnsRecords", "ODBCTimeout", "RecordsetType", "OrderByOn", _
                 "Orientation", "DefaultView", "FilterOnLoad", "OrderByOnLoad", _
                 "TotalsRow", "UseTransaction", "FailOnError"
                ' Already emitted above
            Case "MaxRecords"
                cOut.Add "dbLong ""MaxRecords"" =""" & dOverrides(varKey) & """", vbCrLf
            Case "RecordLocks"
                ' LvProp-only property; not part of the .qdef text format
            Case "RowHeight"
                cOut.Add "dbLong ""RowHeight"" =""" & dOverrides(varKey) & """", vbCrLf
            Case Else
                cOut.Add "dbText """ & CStr(varKey) & """ =""" & _
                    EscapeQdefString(CStr(dOverrides(varKey))) & """", vbCrLf
        End Select
    Next varKey

    CatchAny eelError, "Error emitting query properties. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitAllProperties"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EscapeQdefString
' Author    : Adam Waller
' Date      : 4/20/2026
' Purpose   : Escape a string value for inclusion inside a "..." segment of a
'           : .qdef text file. Doubles backslashes first (so `\` becomes `\\`),
'           : then escapes embedded double quotes (so `"` becomes `\"`).
'           : The order matters -- escaping quotes first would re-double the
'           : backslash that was just inserted as part of `\"`.
'---------------------------------------------------------------------------------------
'
Private Function EscapeQdefString(strValue As String) As String
    Dim s As String
    s = Replace(strValue, "\", "\\")
    s = Replace(s, """", "\""")
    EscapeQdefString = s
End Function


Private Function ResolvePropertyOverrides(dProperties As Dictionary) As Dictionary
    ' Flatten JSON property values: unwrap {Type, Value} dictionaries
    ' from LvProp parser into plain string values.
    Dim dResult As New Dictionary
    If dProperties Is Nothing Then
        Set ResolvePropertyOverrides = dResult
        Exit Function
    End If
    Dim varKey As Variant
    Dim varVal As Variant
    For Each varKey In dProperties.Keys
        If TypeOf dProperties(varKey) Is Dictionary Then
            varVal = dProperties(varKey)("Value")
        Else
            varVal = dProperties(varKey)
        End If
        ' Normalize boolean values to Access format
        Select Case CStr(varKey)
            Case "ReturnsRecords", "OrderByOn", "FilterOnLoad", "OrderByOnLoad", _
                 "TotalsRow", "UseTransaction", "FailOnError"
                dResult.Add CStr(varKey), BoolToAccessStr(varVal)
            Case Else
                dResult.Add CStr(varKey), CStr(varVal)
        End Select
    Next varKey
    Set ResolvePropertyOverrides = dResult
End Function


Private Function BoolToAccessStr(varValue As Variant) As String
    If IsNumeric(varValue) Then
        If CLng(varValue) <> 0 Then BoolToAccessStr = "-1" Else BoolToAccessStr = "0"
    ElseIf VarType(varValue) = vbBoolean Then
        If CBool(varValue) Then BoolToAccessStr = "-1" Else BoolToAccessStr = "0"
    ElseIf VarType(varValue) = vbString Then
        If varValue = "-1" Or LCase$(varValue) = "true" Then
            BoolToAccessStr = "-1"
        Else
            BoolToAccessStr = "0"
        End If
    Else
        BoolToAccessStr = CStr(varValue)
    End If
End Function


Private Function GetPropOrDefault(dOverrides As Dictionary, strKey As String, strDefault As String) As String
    If dOverrides.Exists(strKey) Then
        GetPropOrDefault = dOverrides(strKey)
        dOverrides.Remove strKey
    Else
        GetPropOrDefault = strDefault
    End If
End Function


Private Sub EmitBoolProp(cOut As clsConcat, strName As String, varValue As Variant)
    Dim strVal As String
    If IsNumeric(varValue) Then
        If CLng(varValue) <> 0 Then strVal = "-1" Else strVal = "0"
    ElseIf VarType(varValue) = vbBoolean Then
        If CBool(varValue) Then strVal = "-1" Else strVal = "0"
    Else
        strVal = CStr(varValue)
    End If
    cOut.Add "dbBoolean """ & strName & """ =""" & strVal & """", vbCrLf
End Sub


Private Sub EmitColumnMetadata(cOut As clsConcat, dColumns As Dictionary, _
                               Optional blnSynthesize As Boolean = False)
    ' Column metadata block: required structural delimiter between properties
    ' and the layout block. If JSON Columns are available (from LvProp), use
    ' those. Otherwise, synthesize from the decomposed SQL structure when
    ' blnSynthesize is True (Design View only).

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' When we have JSON column data, use it as-is
    If Not dColumns Is Nothing Then
        If dColumns.Count > 0 Then
            EmitColumnBlock cOut, dColumns
            GoTo CleanUp
        End If
    End If

    ' For SQL View, only emit if we have JSON data (already handled above)
    If Not blnSynthesize Then GoTo CleanUp

    ' Synthesize column metadata from SQL structure (Design View)
    Dim dSynthesized As New Dictionary
    Dim dOutCol As Dictionary
    Dim strColName As String

    ' Add output columns
    For Each dOutCol In m_colOutputColumns
        ' UPDATE queries: use the qualified targetName (e.g., tblCarsPrice.Price)
        ' INSERT queries: use expression/alias + add targetTable.targetField
        ' SELECT/DELETE: use expression or alias
        If m_lngQueryType = 4 And Len(dOutCol("targetName")) > 0 Then
            strColName = CStr(dOutCol("targetName"))
        ElseIf Len(dOutCol("alias")) > 0 Then
            strColName = CStr(dOutCol("alias"))
        Else
            strColName = CStr(dOutCol("expression"))
        End If
        If Not dSynthesized.Exists(strColName) Then
            Dim dEntry As Dictionary
            Set dEntry = New Dictionary
            dEntry.Add "AggregateType", -1
            dSynthesized.Add strColName, dEntry
        End If
        ' For INSERT queries, also add qualified targetTable.targetField entry
        If m_lngQueryType = 3 And Len(m_strTargetTable) > 0 And Len(dOutCol("targetName")) > 0 Then
            Dim strTargetQual As String
            strTargetQual = m_strTargetTable & "." & CStr(dOutCol("targetName"))
            If Not dSynthesized.Exists(strTargetQual) Then
                Dim dTargetEntry As Dictionary
                Set dTargetEntry = New Dictionary
                dTargetEntry.Add "AggregateType", -1
                dSynthesized.Add strTargetQual, dTargetEntry
            End If
        End If
    Next dOutCol

    ' Add WHERE-referenced fields that aren't already in the output list
    If Len(m_strWhere) > 0 Then
        AddWhereFieldsToColumns dSynthesized
    End If

    ' Always emit the block (even if empty, it acts as a structural delimiter)
    EmitColumnBlock cOut, dSynthesized

CleanUp:
    CatchAny eelError, "Error emitting column metadata. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitColumnMetadata"
End Sub


Private Sub EmitColumnBlock(cOut As clsConcat, dColumns As Dictionary)
    Dim varCol As Variant
    Dim varProp As Variant

    cOut.Add "Begin", vbCrLf
    For Each varCol In dColumns.Keys
        Dim dCol As Dictionary
        Set dCol = dColumns(varCol)
        cOut.Add "    Begin", vbCrLf
        cOut.Add "        dbText ""Name"" =""" & CStr(varCol) & """", vbCrLf

        ' Emit properties in stored order, holding AggregateType for last
        For Each varProp In dCol.Keys
            If varProp <> "AggregateType" Then
                EmitColumnProperty cOut, CStr(varProp), dCol(varProp)
            End If
        Next varProp

        cOut.Add "        dbLong ""AggregateType"" =""" & _
            CStr(Nz(dCol("AggregateType"), -1)) & """", vbCrLf
        cOut.Add "    End", vbCrLf
    Next varCol
    cOut.Add "End", vbCrLf
End Sub


Private Sub EmitColumnProperty(cOut As clsConcat, strName As String, varValue As Variant)
    Dim strType As String
    Dim strVal As String

    ' Determine the QDEF type prefix
    If IsObject(varValue) Then
        ' Custom property stored as {"Type": "dbMemo", "Value": ...}
        Dim dTyped As Dictionary
        Set dTyped = varValue
        strType = CStr(dTyped("Type"))
        strVal = CStr(Nz(dTyped("Value")))
    Else
        ' Known property — look up the type from the translation table
        strType = KnownColumnQdefType(strName)
        If Len(strType) = 0 Then strType = "dbMemo"
        strVal = CStr(Nz(varValue))
    End If

    ' Format and emit the value based on type
    Select Case strType
        Case "dbBoolean"
            cOut.Add "        dbBoolean """ & strName & """ =""" & _
                IIf(Val(strVal) <> 0, "-1", "0") & """", vbCrLf
        Case "dbMemo", "dbText"
            EmitWrappedString cOut, "        " & strType & " """ & strName & """", strVal
        Case Else
            cOut.Add "        " & strType & " """ & strName & """ =""" & strVal & """", vbCrLf
    End Select
End Sub


Private Function KnownColumnQdefType(strName As String) As String
    Select Case strName
        Case "AggregateType":    KnownColumnQdefType = "dbLong"
        Case "ColumnWidth":      KnownColumnQdefType = "dbInteger"
        Case "ColumnHidden":     KnownColumnQdefType = "dbBoolean"
        Case "ColumnOrder":      KnownColumnQdefType = "dbInteger"
        Case "Caption":          KnownColumnQdefType = "dbMemo"
        Case "Description":      KnownColumnQdefType = "dbMemo"
        Case "TextAlign":        KnownColumnQdefType = "dbByte"
        Case "DisplayControl":   KnownColumnQdefType = "dbInteger"
        Case "ResultType":       KnownColumnQdefType = "dbInteger"
        Case "CurrencyLCID":     KnownColumnQdefType = "dbLong"
        Case "ShowDatePicker":   KnownColumnQdefType = "dbInteger"
        Case "IMEMode":          KnownColumnQdefType = "dbInteger"
        Case "IMESentenceMode":  KnownColumnQdefType = "dbInteger"
    End Select
End Function


Private Sub AddWhereFieldsToColumns(dColumns As Dictionary)
    ' Extract table.field references from the WHERE clause and add them
    ' as columns if not already present.
    Dim strWhere As String
    Dim lngI As Long
    Dim lngLen As Long
    Dim lngStart As Long
    Dim strToken As String
    Dim blnInString As Boolean

    strWhere = m_strWhere
    lngLen = Len(strWhere)
    lngStart = 0

    For lngI = 1 To lngLen
        Dim strCh As String
        strCh = Mid$(strWhere, lngI, 1)
        If strCh = """" Then blnInString = Not blnInString
        If blnInString Then GoTo NextChar

        ' Look for table.field patterns: word characters and dots
        If IsIdentChar(strCh) Or strCh = "." Then
            If lngStart = 0 Then lngStart = lngI
        Else
            If lngStart > 0 Then
                strToken = Mid$(strWhere, lngStart, lngI - lngStart)
                ' Only add if it looks like table.field (contains a dot)
                If InStr(1, strToken, ".") > 0 Then
                    If Not dColumns.Exists(strToken) Then
                        Dim dEntry As Dictionary
                        Set dEntry = New Dictionary
                        dEntry.Add "AggregateType", -1
                        dColumns.Add strToken, dEntry
                    End If
                End If
                lngStart = 0
            End If
        End If
NextChar:
    Next lngI

    ' Handle trailing token
    If lngStart > 0 Then
        strToken = Mid$(strWhere, lngStart)
        If InStr(1, strToken, ".") > 0 Then
            If Not dColumns.Exists(strToken) Then
                Dim dEntryFinal As Dictionary
                Set dEntryFinal = New Dictionary
                dEntryFinal.Add "AggregateType", -1
                dColumns.Add strToken, dEntryFinal
            End If
        End If
    End If
End Sub


Private Function IsIdentChar(strCh As String) As Boolean
    Dim lngCode As Long
    lngCode = AscW(strCh)
    IsIdentChar = (lngCode >= 65 And lngCode <= 90) Or _
                  (lngCode >= 97 And lngCode <= 122) Or _
                  (lngCode >= 48 And lngCode <= 57) Or _
                  lngCode = 95 Or lngCode = 91 Or lngCode = 93
End Function


Private Sub EmitDesignLayout(cOut As clsConcat, dLayout As Dictionary)
    ' Design View layout block

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    If dLayout Is Nothing Then GoTo CleanUp

    cOut.Add "Begin", vbCrLf

    ' Window rect
    If dLayout.Exists("state") Then
        cOut.Add "    State =" & CStr(dLayout("state")), vbCrLf
    Else
        cOut.Add "    State =0", vbCrLf
    End If

    ' Window coordinates
    If dLayout.Exists("window") Then
        Dim dWin As Dictionary
        Set dWin = dLayout("window")
        cOut.Add "    Left =" & CStr(dWin("left")), vbCrLf
        cOut.Add "    Top =" & CStr(dWin("top")), vbCrLf
        cOut.Add "    Right =" & CStr(dWin("right")), vbCrLf
        cOut.Add "    Bottom =" & CStr(dWin("bottom")), vbCrLf
    End If

    ' Designer pane coordinates
    If dLayout.Exists("designerPane") Then
        Dim dPane As Dictionary
        Set dPane = dLayout("designerPane")
        cOut.Add "    Left =" & CStr(dPane("left")), vbCrLf
        cOut.Add "    Top =" & CStr(dPane("top")), vbCrLf
        cOut.Add "    Right =" & CStr(dPane("right")), vbCrLf
        cOut.Add "    Bottom =" & CStr(dPane("bottom")), vbCrLf
    End If

    ' Grid origin
    If dLayout.Exists("gridLeft") Then
        cOut.Add "    Left =" & CStr(dLayout("gridLeft")), vbCrLf
        cOut.Add "    Top =" & CStr(dLayout("gridTop")), vbCrLf
    End If

    ' Columns shown
    If dLayout.Exists("columnsShown") Then
        cOut.Add "    ColumnsShown =" & CStr(dLayout("columnsShown")), vbCrLf
    End If

    ' Table entries
    If dLayout.Exists("tables") Then
        Dim colTables As Collection
        Dim dTbl As Dictionary
        Set colTables = dLayout("tables")
        For Each dTbl In colTables
            cOut.Add "    Begin", vbCrLf
            cOut.Add "        Left =" & CStr(dTbl("left")), vbCrLf
            cOut.Add "        Top =" & CStr(dTbl("top")), vbCrLf
            cOut.Add "        Right =" & CStr(dTbl("right")), vbCrLf
            cOut.Add "        Bottom =" & CStr(dTbl("bottom")), vbCrLf
            cOut.Add "        Top =" & CStr(dTbl("scrollTop")), vbCrLf
            cOut.Add "        Name =""" & CStr(dTbl("name")) & """", vbCrLf
            cOut.Add "        Name =""" & Nz(dTbl("alias"), vbNullString) & """", vbCrLf
            cOut.Add "    End", vbCrLf
        Next dTbl
    End If

    cOut.Add "End", vbCrLf

CleanUp:
    CatchAny eelError, "Error emitting design layout. " & SqlSnippet(), _
        ModuleName(Me) & ".EmitDesignLayout"
End Sub


Private Sub EmitWrappedString(cOut As clsConcat, strTag As String, strValue As String)
    ' Emit a potentially long string value with .qdef line wrapping.
    ' Embedded backslashes are escaped as \\ and double quotes as \" so that
    ' Access LoadFromText reads back the original string. Backslashes must be
    ' escaped before quotes (otherwise the backslash from \" would be re-doubled).
    Const MAX_LINE As Long = 70
    Dim lngPos As Long
    Dim lngLineLen As Long
    Dim strLine As String
    Dim strEscaped As String

    strEscaped = EscapeQdefString(strValue)

    If Len(strEscaped) <= MAX_LINE Then
        cOut.Add strTag & " =""" & strEscaped & """", vbCrLf
    Else
        lngPos = 1
        lngLineLen = SafeBreak(strEscaped, lngPos, MAX_LINE)
        cOut.Add strTag & " =""" & Mid$(strEscaped, 1, lngLineLen) & """", vbCrLf
        lngPos = lngLineLen + 1
        Do While lngPos <= Len(strEscaped)
            lngLineLen = SafeBreak(strEscaped, lngPos, MAX_LINE)
            If lngPos + lngLineLen - 1 >= Len(strEscaped) Then
                strLine = Mid$(strEscaped, lngPos)
            Else
                strLine = Mid$(strEscaped, lngPos, lngLineLen)
            End If
            cOut.Add "    """ & strLine & """", vbCrLf
            lngPos = lngPos + lngLineLen
        Loop
    End If
End Sub


'=======================================================================================
' PUBLIC: Properties for accessing internal state after decomposition
'=======================================================================================

Public Property Get RawSQL() As String
    RawSQL = m_strRawSql
End Property

Public Property Let RawSQL(strValue As String)
    m_strRawSql = strValue
End Property

Public Property Get ConnectString() As String
    ConnectString = m_strConnectString
End Property

Public Property Let ConnectString(strValue As String)
    m_strConnectString = strValue
End Property

Public Property Get QueryType() As Long
    QueryType = m_lngQueryType
End Property

Public Property Get IsPassThrough() As Boolean
    IsPassThrough = m_blnIsPassThrough
End Property
