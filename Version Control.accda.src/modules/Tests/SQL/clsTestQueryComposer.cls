VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestQueryComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestQueryComposer
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsQueryComposer decompose/reconstruct and classification.
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


Public Sub TestDecomposeAndReconstruct_SimpleSelect()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a, b FROM tbl1 WHERE x = 1"
    m_qc.DecomposeSQL strSql
    strResult = m_qc.GenerateQdef(False, New Dictionary, New Dictionary)
    TestAssert Len(strResult) > 0, "non-empty qdef output"
    TestAssert InStr(strResult, "SELECT") > 0, "contains SELECT"
    TestAssert InStr(strResult, "tbl1") > 0, "table name preserved"
End Sub


Public Sub TestDecomposeAndReconstruct_WithJoin()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM t1 INNER JOIN t2 ON t1.id = t2.id"
    m_qc.DecomposeSQL strSql
    strResult = m_qc.GenerateQdef(False, New Dictionary, New Dictionary)
    TestAssert InStr(strResult, "INNER JOIN") > 0, "JOIN preserved"
End Sub


Public Sub TestIsDesignerCompatible_SimpleSelect()
    m_qc.DecomposeSQL "SELECT a, b FROM tbl1"
    TestAssert m_qc.IsDesignerCompatible, "simple SELECT is designer compatible"
End Sub


Public Sub TestIsDesignerCompatible_Union()
    m_qc.DecomposeSQL "SELECT a FROM t1 UNION SELECT b FROM t2"
    TestAssert Not m_qc.IsDesignerCompatible, "UNION is not designer compatible"
End Sub


Public Sub TestIsPassThrough()
    m_qc.DecomposeSQL "SELECT 1", True
    TestAssert m_qc.IsPassThrough, "pass-through flag set"
End Sub


Public Sub TestReconstructSQL_PassThroughNoRecords()
    Dim colRows As New Collection
    Dim dRow As Dictionary
    Dim strResult As String

    Set dRow = New Dictionary
    dRow.Add "Attribute", 1
    dRow.Add "Flag", 10
    dRow.Add "Expression", "UPDATE tblX SET Col1 = 'x' WHERE 1 = 0"
    dRow.Add "Name1", "ODBC;"
    dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 255
    dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString
    dRow.Add "Name1", vbNullString
    dRow.Add "Name2", vbNullString
    colRows.Add dRow

    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert strResult = "UPDATE tblX SET Col1 = 'x' WHERE 1 = 0", _
        "Attribute 1 Flag 10 preserves raw pass-through SQL"
End Sub


Public Sub TestRawSQL_Property()
    m_qc.RawSQL = "SELECT 42"
    TestAssert m_qc.RawSQL = "SELECT 42", "RawSQL round trip"
End Sub


Public Sub TestStaleDistinctJson_SqlViewIgnoresOptionFlag()
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT a FROM t1"
    strQdef = m_qc.GenerateQdef(False, New Dictionary, New Dictionary, , 2)
    TestAssert InStr(1, strQdef, "DISTINCT", vbTextCompare) = 0, _
        "SQL View qdef must not gain DISTINCT from stale OptionFlag 2"
    TestAssert ((m_qc.ParsedOptionFlag And 2) = 0), _
        "SQL-derived option bits stay 0 when SQL has no DISTINCT"
End Sub


Public Sub TestStaleDistinctJson_DesignViewIgnoresOptionFlag()
    Dim strQdef As String
    Dim lngOption As Long
    m_qc.DecomposeSQL "SELECT a FROM t1"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 2)
    lngOption = QdefOptionValue(strQdef)
    TestAssert ((lngOption And 2) = 0), _
        "Design View Option must not gain DISTINCT from stale OptionFlag 2"
    TestAssert InStr(1, strQdef, "DISTINCT", vbTextCompare) = 0, _
        "Design View qdef text must not contain DISTINCT when SQL has none"
End Sub


Public Sub TestSqlDistinct_BothViewsKeepDistinct()
    Dim strSqlView As String
    Dim strDesignView As String
    m_qc.DecomposeSQL "SELECT DISTINCT a FROM t1"
    strSqlView = m_qc.GenerateQdef(False, New Dictionary, New Dictionary, , 2)
    strDesignView = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 2)
    TestAssert InStr(1, strSqlView, "DISTINCT", vbTextCompare) > 0, _
        "SQL View keeps DISTINCT from SQL"
    TestAssert ((QdefOptionValue(strDesignView) And 2) = 2), _
        "Design View Option keeps DISTINCT from SQL"
    TestAssert m_qc.WarningCount = 0, _
        "an OptionFlag that matches the SQL raises no warning"
End Sub


Public Sub TestStaleDistinctJson_WarnsAndDoesNotCallLogError()
    Dim lngErrBefore As Long
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT a FROM t1"
    lngErrBefore = Log.ErrorCount
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 2)
    TestAssert Len(strQdef) > 0, "qdef generated for warning check"
    TestAssert m_qc.WarningCount > 0, _
        "stale OptionFlag 2 produces a composer warning"
    TestAssert Log.ErrorCount = lngErrBefore, _
        "OptionFlag disagreement must not call Log.Error"
End Sub


Public Sub TestStaleTopJson_NoTopN_NotApplied()
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT a FROM t1"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 16)
    TestAssert ((QdefOptionValue(strQdef) And 16) = 0), _
        "stale TOP bit is not applied when SQL has no TOP n"
    TestAssert InStr(1, strQdef, "RowCount", vbTextCompare) = 0, _
        "stale TOP bit must not emit RowCount"
End Sub


Public Sub TestStaleDistinctJson_UpdateIgnored()
    Dim strQdef As String
    m_qc.DecomposeSQL "UPDATE t1 SET a = 1"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 2)
    TestAssert ((QdefOptionValue(strQdef) And 2) = 0), _
        "UPDATE must not pick up DISTINCT from OptionFlag 2"
    TestAssert ((m_qc.ParsedOptionFlag And 2) = 0), _
        "parsed UPDATE option bits exclude DISTINCT"
End Sub


Public Sub TestStaleDistinctJson_CompoundOnDesignView()
    Dim strSql As String
    Dim strSqlView As String
    Dim strDesignView As String
    strSql = "SELECT a FROM t1 INNER JOIN t2 ON (t1.id = t2.id) AND (t1.x = t2.x)"
    m_qc.DecomposeSQL strSql
    TestAssert m_qc.RequiresDesignView, _
        "compound ON forces Design View"
    strSqlView = m_qc.GenerateQdef(False, New Dictionary, New Dictionary, , 2)
    strDesignView = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 2)
    TestAssert InStr(1, strSqlView, "DISTINCT", vbTextCompare) = 0, _
        "SQL View compound-ON qdef ignores stale DISTINCT"
    TestAssert ((QdefOptionValue(strDesignView) And 2) = 0), _
        "Design View compound-ON Option ignores stale DISTINCT"
End Sub


Public Sub TestOutputAllFieldsJson_IgnoredWhenSqlHasNoStar()
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT a FROM t1"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 1)
    TestAssert ((QdefOptionValue(strQdef) And 1) = 0), _
        "Output All Fields is not restored from JSON when SQL has no `*`"
    TestAssert m_qc.WarningCount > 0, _
        "a JSON-only Output All Fields bit is reported as a stale companion"
End Sub


Public Sub TestOutputAllFields_ComesFromExplicitStar()
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT t1.*, * FROM t1"
    TestAssert ((m_qc.ParsedOptionFlag And 1) = 1), _
        "an explicit `*` beside named columns sets Output All Fields"
    TestAssert Not m_qc.ImplicitOutputAllFields, _
        "an explicit `*` with output columns is not the implicit shape"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 1)
    TestAssert ((QdefOptionValue(strQdef) And 1) = 1), _
        "Design View Option carries Output All Fields from the SQL"
    TestAssert m_qc.WarningCount = 0, _
        "OptionFlag 1 agrees with an explicit `*`"
End Sub


Public Sub TestOutputAllFields_BareStarIsImplicit()
    m_qc.DecomposeSQL "SELECT * FROM t1"
    TestAssert ((m_qc.ParsedOptionFlag And 1) = 1), _
        "a bare `*` sets Output All Fields while storing no output columns"
    TestAssert m_qc.ImplicitOutputAllFields, _
        "a bare `SELECT *` is the implicit shape exports omit"
End Sub


Public Sub TestReconcileOptionFlag_OmittedFlagIsSilent()
    m_qc.DecomposeSQL "SELECT DISTINCT a FROM t1"
    m_qc.GenerateQdef True, New Dictionary, New Dictionary
    TestAssert m_qc.WarningCount = 0, _
        "a companion with no OptionFlag makes no claim, so nothing is stale"
End Sub


Public Sub TestReconcileOptionFlag_SqlOnlyBitWarns()
    Dim strQdef As String
    m_qc.DecomposeSQL "SELECT DISTINCT a FROM t1"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 0)
    TestAssert m_qc.WarningCount > 0, _
        "OptionFlag 0 against SELECT DISTINCT is a disagreement in the other direction"
    TestAssert ((QdefOptionValue(strQdef) And 2) = 2), _
        "the SQL still wins when the companion dropped the bit"
End Sub


Public Sub TestReconcileOptionFlag_UnionAllVsUnionWarns()
    m_qc.DecomposeSQL "SELECT a FROM t1 UNION SELECT b FROM t2"
    TestAssert m_qc.ParsedOptionFlag = 3, "deduplicating UNION is flag 3"
    m_qc.GenerateQdef False, New Dictionary, New Dictionary, , 1
    TestAssert m_qc.WarningCount > 0, _
        "OptionFlag 1 (UNION ALL) against a plain UNION is reported"
End Sub


Public Sub TestReconcileOptionFlag_BareSelectStarIsSilent()
    m_qc.DecomposeSQL "SELECT * FROM t1"
    m_qc.GenerateQdef False, New Dictionary, New Dictionary, , 0
    TestAssert m_qc.WarningCount = 0, _
        "an omitted bit 1 matches a bare `SELECT *`, which exports never write"

    Set m_qc = New clsQueryComposer
    m_qc.DecomposeSQL "SELECT * FROM t1"
    m_qc.GenerateQdef False, New Dictionary, New Dictionary, , 1
    TestAssert m_qc.WarningCount = 0, _
        "Access's own Attribute 3 = 1 for a bare `SELECT *` is not stale either"
End Sub


Public Sub TestParseInsertQuery_TopN()
    Dim strQdef As String
    m_qc.DecomposeSQL "INSERT INTO t2 ( a ) SELECT TOP 3 t1.a FROM t1"
    TestAssert m_qc.ParsedOptionFlag = 16, _
        "INSERT ... SELECT TOP n sets the TOP bit"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 16)
    TestAssert InStr(1, strQdef, "RowCount =""3""", vbTextCompare) > 0, _
        "the append qdef carries the row count"
    TestAssert InStr(1, strQdef, "Expression =""TOP", vbTextCompare) = 0, _
        "TOP is a modifier, not the first output column's expression"
End Sub


Public Sub TestParseInsertQuery_TopPercent()
    Dim strQdef As String
    m_qc.DecomposeSQL "INSERT INTO t2 ( a ) SELECT DISTINCT TOP 5 PERCENT t1.a FROM t1"
    TestAssert m_qc.ParsedOptionFlag = 50, _
        "INSERT ... SELECT DISTINCT TOP n PERCENT sets bits 2, 16 and 32"
    strQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary, , 50)
    TestAssert InStr(1, strQdef, "RowCount =""5""", vbTextCompare) > 0, _
        "the append qdef carries the percentage as its row count"
    TestAssert m_qc.WarningCount = 0, _
        "a companion matching the parsed INSERT modifiers raises no warning"
End Sub


Public Sub TestApplyOptionFlagToSql_InjectsDistinct()
    Dim strResult As String
    strResult = m_qc.ApplyOptionFlagToSql("SELECT a FROM t1", 2)
    TestAssert InStr(1, strResult, "DISTINCT", vbTextCompare) > 0, _
        "ApplyOptionFlagToSql injects DISTINCT into SELECT"
    strResult = m_qc.ApplyOptionFlagToSql("UPDATE t1 SET a = 1", 2)
    TestAssert InStr(1, strResult, "DISTINCT", vbTextCompare) = 0, _
        "ApplyOptionFlagToSql does not inject DISTINCT into UPDATE"
End Sub


Public Sub TestApplyOptionFlagToSql_DoesNotInjectStar()
    Dim strResult As String
    strResult = m_qc.ApplyOptionFlagToSql("SELECT a FROM t1", 1)
    TestAssert InStr(strResult, "*") = 0, _
        "ApplyOptionFlagToSql does not inject * for Output All Fields"
End Sub


Public Sub TestApplyOptionFlagToSql_UnionSkipsDistinctBit()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM t1 UNION SELECT b FROM t2"
    strResult = m_qc.ApplyOptionFlagToSql(strSql, 3)
    TestAssert InStr(1, strResult, "DISTINCT", vbTextCompare) = 0, _
        "UNION Flag 3 must not inject DISTINCT into the first SELECT"
End Sub


Public Sub TestApplyOptionFlagToSql_NoTopForActionQueries()
    Dim strResult As String

    strResult = m_qc.ApplyOptionFlagToSql("UPDATE t1 SET a = 1", 16, 5)
    TestAssert InStr(1, strResult, "TOP", vbTextCompare) = 0, _
        "Jet cannot parse TOP on an UPDATE, so the bit is not injected"

    strResult = m_qc.ApplyOptionFlagToSql("DELETE * FROM t1", 48, 10, True)
    TestAssert InStr(1, strResult, "TOP", vbTextCompare) = 0, _
        "TOP PERCENT is not injected into a DELETE either"

    ' The guard is specific to TOP: DISTINCTROW is legal on an action query.
    strResult = m_qc.ApplyOptionFlagToSql("DELETE * FROM t1", 8)
    TestAssert InStr(1, strResult, "DISTINCTROW", vbTextCompare) > 0, _
        "DELETE still gains DISTINCTROW from Attribute 3"
End Sub


Public Sub TestApplyOptionFlagToSql_ActionLiteralContainingSelect()
    Dim strSql As String
    Dim strResult As String

    strSql = "UPDATE t1 SET note = 'prefix SELECT value'"
    strResult = m_qc.ApplyOptionFlagToSql(strSql, 18, 5)
    TestAssert strResult = strSql, _
        "SELECT inside a single-quoted UPDATE value cannot bypass the action-query guard"
End Sub


Public Sub TestApplyOptionFlagToSql_InjectsTopAfterDistinct()
    Dim strResult As String
    strResult = m_qc.ApplyOptionFlagToSql("SELECT DISTINCT a FROM t1", 18, 5)
    TestAssert InStr(1, strResult, "SELECT DISTINCT TOP 5 ", vbTextCompare) > 0, _
        "TOP is injected after the existing DISTINCT, where Jet expects it"
    TestAssert m_qc.OptionBitsFromSql(strResult) = 18, _
        "the injected SQL parses back to the same option mask"
End Sub


Public Sub TestOptionBitsFromSql_UnionKeywordInLiteral()
    TestAssert m_qc.OptionBitsFromSql("SELECT a FROM t1 WHERE b = ""UNION ALL""") = 0, _
        "UNION inside a double-quoted literal is data, not a set operator"
    TestAssert m_qc.OptionBitsFromSql("SELECT DISTINCT a FROM t1 WHERE b = 'x UNION y'") = 2, _
        "a single-quoted literal does not hide the real DISTINCT modifier"
End Sub


Public Sub TestOptionBitsFromSql_UnionKeywordInIdentifier()
    m_qc.DecomposeSQL "SELECT t1.[UNION ALL] FROM t1"
    TestAssert m_qc.QueryType <> 9, _
        "a bracketed field named UNION ALL does not make a union query"
    TestAssert m_qc.ParsedOptionFlag = 0, _
        "a bracketed identifier contributes no option bits"
End Sub


Public Sub TestApplyOwnerAccess_LiteralIsNotTheClause()
    Dim strResult As String
    strResult = m_qc.ApplyOptionFlagToSql( _
        "SELECT a FROM t1 WHERE b = ""WITH OWNERACCESS OPTION""", 4)
    TestAssert InStr(1, strResult, "WITH OWNERACCESS OPTION;", vbTextCompare) > 0, _
        "the phrase inside a literal does not satisfy the OWNERACCESS bit"
    TestAssert ((m_qc.OptionBitsFromSql(strResult) And 4) = 4), _
        "the appended clause parses back as OWNERACCESS"
End Sub


Public Sub TestApplyOwnerAccess_NotDuplicatedAcrossLineBreaks()
    Dim strSql As String
    strSql = "SELECT a" & vbCrLf & "FROM t1" & vbCrLf & "WITH" & vbCrLf & _
        "  OWNERACCESS OPTION;"
    TestAssert m_qc.ApplyOptionFlagToSql(strSql, 4) = strSql, _
        "an OWNERACCESS clause Access wrapped across lines is already present"
End Sub


Public Sub TestOptionBitsFromSql_OwnerAccessAfterSingleQuotedLiteral()
    Dim strSql As String

    strSql = "SELECT a FROM t1 WHERE b = 'prefix WITH OWNERACCESS OPTION' " & _
        "WITH OWNERACCESS OPTION;"
    TestAssert ((m_qc.OptionBitsFromSql(strSql) And 4) = 4), _
        "a phrase inside a single-quoted literal cannot hide the real OWNERACCESS clause"
End Sub


Public Sub TestOptionBitsFromSql_MixedUnionUsesDeduplicatingFlag()
    Dim strSql As String

    strSql = "SELECT a FROM t1 UNION ALL SELECT b FROM t2 UNION SELECT c FROM t3"
    TestAssert m_qc.OptionBitsFromSql(strSql) = 3, _
        "any deduplicating operator in a mixed UNION query selects flag 3"
End Sub


Public Sub TestImplicitOutputAllFields_RequiresStarBit()
    m_qc.DecomposeSQL "SELECT FROM ;"
    TestAssert Not m_qc.ImplicitOutputAllFields, _
        "zero parsed output columns alone do not imply a bare star"
End Sub


Public Sub TestReconstructSQL_InjectsDistinctAt510()
    Dim colRows As Collection
    Dim lngSaved As Long
    Dim strResult As String

    Set colRows = RawSelectWithOption(2)
    lngSaved = Options.ExportFormatVersion

    Options.ExportFormatVersion = EFV_5_1_0
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, "DISTINCT", vbTextCompare) > 0, _
        "5.1.0 ReconstructSQL injects DISTINCT from Attribute 3"

    Options.ExportFormatVersion = EFV_5_0_0
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, "DISTINCT", vbTextCompare) = 0, _
        "5.0.0 ReconstructSQL leaves raw Attribute 0 SQL unchanged"

    Options.ExportFormatVersion = lngSaved
End Sub


Public Sub TestReconstructSQL_InjectsOtherModifiersAt510()
    Dim colRows As Collection
    Dim lngSaved As Long
    Dim strResult As String

    lngSaved = Options.ExportFormatVersion
    Options.ExportFormatVersion = EFV_5_1_0

    Set colRows = RawSqlWithOption("SELECT a FROM t1", 48, "5")
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, "SELECT TOP 5 PERCENT ", vbTextCompare) > 0, _
        "5.1.0 injects TOP n PERCENT from Attribute 3"

    Set colRows = RawSqlWithOption("DELETE * FROM t1", 8)
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, "DELETE DISTINCTROW *", vbTextCompare) > 0, _
        "5.1.0 injects DISTINCTROW into DELETE"

    Set colRows = RawSqlWithOption( _
        "INSERT INTO t2 (a) SELECT a FROM t1", 50, "5")
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, "SELECT DISTINCT TOP 5 PERCENT ", vbTextCompare) > 0, _
        "5.1.0 injects modifiers into the SELECT portion of an append query"

    Set colRows = RawSqlWithOption( _
        "SELECT a FROM t1 WHERE b = 'prefix WITH OWNERACCESS OPTION'", 4)
    strResult = m_qc.ReconstructSQL(colRows)
    TestAssert InStr(1, strResult, vbCrLf & "WITH OWNERACCESS OPTION;", vbTextCompare) > 0, _
        "5.1.0 appends OWNERACCESS when the phrase exists only inside a literal"

    Options.ExportFormatVersion = lngSaved
End Sub


Public Sub TestOptionBitsFromSql_Modifiers()
    TestAssert m_qc.OptionBitsFromSql("SELECT DISTINCT a FROM t1") = 2, _
        "SELECT DISTINCT -> bit 2"
    TestAssert ((m_qc.OptionBitsFromSql("SELECT a FROM t1 WITH OWNERACCESS OPTION") And 4) = 4), _
        "OWNERACCESS -> bit 4"
    TestAssert m_qc.OptionBitsFromSql("SELECT a FROM t1 UNION ALL SELECT b FROM t2") = 1, _
        "UNION ALL -> flag 1"
    TestAssert m_qc.OptionBitsFromSql("SELECT a FROM t1 UNION SELECT b FROM t2") = 3, _
        "UNION -> flag 3"
    TestAssert m_qc.OptionBitsFromSql("UPDATE t1 SET a = 1") = 0, _
        "plain UPDATE has no option bits"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : QdefOptionValue
' Author    : Adam Waller
' Date      : 9/8/2026
' Purpose   : Parse the Design View `Option =n` line. Returns -1 when absent
'           : (SQL View qdefs have no Option line).
'---------------------------------------------------------------------------------------
'
Private Function QdefOptionValue(ByVal strQdef As String) As Long
    Dim lngPos As Long
    Dim lngEnd As Long
    Dim strVal As String

    QdefOptionValue = -1
    lngPos = InStr(1, strQdef, "Option =", vbTextCompare)
    If lngPos = 0 Then Exit Function
    lngPos = lngPos + Len("Option =")
    lngEnd = InStr(lngPos, strQdef, vbCrLf)
    If lngEnd = 0 Then lngEnd = Len(strQdef) + 1
    strVal = Trim$(Mid$(strQdef, lngPos, lngEnd - lngPos))
    If IsNumeric(strVal) Then QdefOptionValue = CLng(strVal)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RawSelectWithOption
' Author    : Adam Waller
' Date      : 9/8/2026
' Purpose   : Attribute 0 raw SQL plus Attribute 3 option bits, the SQL View
'           : shape that previously dropped DISTINCT on export.
'---------------------------------------------------------------------------------------
'
Private Function RawSelectWithOption(ByVal lngOption As Long) As Collection
    Set RawSelectWithOption = RawSqlWithOption("SELECT a FROM t1", lngOption)
End Function


Private Function RawSqlWithOption(ByVal strSql As String, ByVal lngOption As Long, _
    Optional ByVal strTopN As String = vbNullString) As Collection

    Dim colRows As Collection
    Dim dRow As Dictionary

    Set colRows = New Collection

    Set dRow = New Dictionary
    dRow.Add "Attribute", 0
    dRow.Add "Flag", 0
    dRow.Add "Expression", strSql
    dRow.Add "Name1", vbNullString
    dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 3
    dRow.Add "Flag", lngOption
    dRow.Add "Expression", vbNullString
    dRow.Add "Name1", strTopN
    dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set RawSqlWithOption = colRows
End Function
