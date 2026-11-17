VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSqlCanonicalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSqlCanonicalizer
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Tests for clsSqlCanonicalizer: each removal rule, each construct whose
'           : parentheses must survive, the fail-safe, and idempotence.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

Private m_canon As clsSqlCanonicalizer


Private Sub Class_Initialize()
    Set m_canon = New clsSqlCanonicalizer
End Sub


Private Sub Class_Terminate()
    Set m_canon = Nothing
End Sub


' Canonicalize strSql, expect strExpected, and expect a second pass to change nothing.
Private Sub AssertCanonical(strSql As String, strExpected As String, strContext As String)
    Dim strResult As String
    strResult = m_canon.Canonicalize(strSql)
    TestAssert strResult = strExpected, strContext
    TestAssert m_canon.Changed = (strExpected <> strSql), strContext & ": Changed flag"
    TestAssert Len(m_canon.DeclineReason) = 0, strContext & ": no decline"
    TestAssert m_canon.Canonicalize(strResult) = strResult, strContext & ": idempotent"
    TestAssert Not m_canon.Changed, strContext & ": second pass changes nothing"
End Sub


Private Sub AssertUnchanged(strSql As String, strContext As String)
    AssertCanonical strSql, strSql, strContext
End Sub


Public Sub TestRule_ClauseWrapper()
    AssertCanonical "SELECT a FROM t WHERE (t.a = 1 AND t.b = 2);", _
        "SELECT a FROM t WHERE t.a = 1 AND t.b = 2;", "pair around the whole WHERE condition"
End Sub


Public Sub TestRule_DesignerWhere()
    AssertCanonical "SELECT a FROM t WHERE (((t.a)=1) AND ((t.b)=2));", _
        "SELECT a FROM t WHERE (t.a=1) AND (t.b=2);", "designer WHERE keeps one pair per condition"
End Sub


Public Sub TestRule_DesignerHaving()
    AssertCanonical "SELECT a FROM t GROUP BY t.a HAVING (((Count(t.b))>1));", _
        "SELECT a FROM t GROUP BY t.a HAVING (Count(t.b))>1;", _
        "HAVING loses its wrappers; the pair around a function call is not a single value"
End Sub


Public Sub TestRule_SingleValues()
    AssertCanonical _
        "SELECT a FROM t WHERE ((t.d)=(#1/1/2000#)) AND ((t.s)=(""x"")) AND ((t.p)=([Enter ID])) AND (([Forms]![frmA]![txtB])=(5));", _
        "SELECT a FROM t WHERE (t.d=#1/1/2000#) AND (t.s=""x"") AND (t.p=[Enter ID]) AND ([Forms]![frmA]![txtB]=5);", _
        "dates, strings, parameters, form references, and numbers are unwrapped"
End Sub


Public Sub TestRule_SeparatorKeptWhenTokensWouldFuse()
    AssertCanonical "SELECT (t.a) AS x FROM t WHERE (t.a)Like""x*"";", _
        "SELECT (t.a) AS x FROM t WHERE t.a Like""x*"";", _
        "a space replaces the parenthesis between a name and a keyword; SELECT is untouched"
End Sub


Public Sub TestRule_EveryWhereInAUnion()
    AssertCanonical "SELECT a FROM t WHERE ((t.a)=1) UNION SELECT b FROM u WHERE ((u.b)=2);", _
        "SELECT a FROM t WHERE t.a=1 UNION SELECT b FROM u WHERE u.b=2;", "each SELECT of a union"
End Sub


Public Sub TestRule_ClauseEndsAtOrderBy()
    AssertCanonical "SELECT a FROM t WHERE ((t.a)=1) ORDER BY a;", _
        "SELECT a FROM t WHERE t.a=1 ORDER BY a;", "WHERE ends at ORDER BY"
End Sub


Public Sub TestProtected_FunctionArgument()
    AssertUnchanged "SELECT a FROM t WHERE Nz((t.x), 0) = 1;", "Nz((x), 0)"
End Sub


Public Sub TestProtected_InList()
    AssertUnchanged "SELECT a FROM t WHERE t.x In (5);", "x In (5)"
End Sub


Public Sub TestProtected_SignedNumber()
    AssertUnchanged "SELECT a FROM t WHERE t.x - (-5) > 0;", "x - (-5)"
End Sub


Public Sub TestProtected_Null()
    AssertUnchanged "SELECT a FROM t WHERE t.a = (Null);", "a = (Null)"
End Sub


Public Sub TestProtected_OnClause()
    AssertCanonical "SELECT a FROM tblA INNER JOIN tblB ON ((tblA.ID) = (tblB.AID)) AND (tblA.X = tblB.X) WHERE ((tblA.ID)=1);", _
        "SELECT a FROM tblA INNER JOIN tblB ON ((tblA.ID) = (tblB.AID)) AND (tblA.X = tblB.X) WHERE tblA.ID=1;", _
        "ON keeps every parenthesis while WHERE is normalized"
End Sub


Public Sub TestProtected_Not()
    AssertUnchanged "SELECT a FROM t WHERE Not (t.a = 1 Or t.b = 2);", "Not (a = 1 Or b = 2)"
End Sub


Public Sub TestProtected_ConditionGroup()
    AssertUnchanged "SELECT a FROM t WHERE (t.a Between 1 And 5) Or t.b = 2;", "(a Between 1 And 5) Or ..."
End Sub


Public Sub TestProtected_Subquery()
    AssertCanonical _
        "SELECT a FROM t WHERE t.a = (SELECT Max(u.a) FROM u WHERE ((u.b)=1)) AND Exists (SELECT (u.c) FROM u);", _
        "SELECT a FROM t WHERE t.a = (SELECT Max(u.a) FROM u WHERE u.b=1) AND Exists (SELECT (u.c) FROM u);", _
        "subquery pairs stay; only the subquery's own WHERE is normalized"
End Sub


Public Sub TestDecline_Comment()
    AssertUnchanged "SELECT a FROM t WHERE ((t.a)=1) # note", "a comment leaves the statement as written"
End Sub


Public Sub TestDecline_UnbalancedInput()
    Dim strSql As String
    strSql = "SELECT a FROM t WHERE ((t.a)=1;"
    TestAssert m_canon.Canonicalize(strSql) = strSql, "unbalanced input is returned unchanged"
    TestAssert Not m_canon.Changed, "unbalanced input: Changed is False"
    TestAssert m_canon.DeclineReason = "unbalanced parentheses in input", "unbalanced input: reason"
End Sub


Public Sub TestFailSafe_TokenChange()
    TestAssert m_canon.CheckRewrite("WHERE (a)", "WHERE b") = "non-parenthesis tokens changed", _
        "a rewrite that changes any other token is refused"
End Sub


Public Sub TestFailSafe_UnbalancedResult()
    TestAssert m_canon.CheckRewrite("WHERE (a)", "WHERE (a") = "unbalanced parentheses in result", _
        "a rewrite that unbalances the parentheses is refused"
End Sub


Public Sub TestFailSafe_ParenthesisRemovalAccepted()
    TestAssert Len(m_canon.CheckRewrite("WHERE ((a))", "WHERE a")) = 0, _
        "removing balanced pairs is accepted"
End Sub


Public Sub TestEmptyAndNoClause()
    AssertUnchanged vbNullString, "empty input"
    AssertUnchanged "SELECT (t.a) FROM t;", "no WHERE or HAVING"
End Sub
