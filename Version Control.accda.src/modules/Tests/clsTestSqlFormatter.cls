VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSqlFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSqlFormatter
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsSqlFormatter.FormatSQL. Exercises formatting of various SQL
'           : shapes and verifies idempotency (formatting an already-formatted string
'           : produces the same result).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")

Private m_fmt As clsSqlFormatter


Private Sub Class_Initialize()
    Set m_fmt = New clsSqlFormatter
End Sub


Private Sub Class_Terminate()
    Set m_fmt = Nothing
End Sub


Public Sub TestFormatSQL_SimpleSelect()
    Dim strResult As String
    strResult = m_fmt.FormatSQL("SELECT a, b FROM tbl WHERE x=1", esdAccess)
    TestAssert Len(strResult) > 0, "returns non-empty result"
    TestAssert InStr(strResult, "SELECT") > 0, "contains SELECT keyword"
End Sub


Public Sub TestFormatSQL_Idempotent()
    Dim strSql As String
    Dim strFirst As String
    Dim strSecond As String
    strSql = "SELECT a, b, c FROM tbl1 INNER JOIN tbl2 ON tbl1.id = tbl2.id WHERE a > 5 ORDER BY b"
    strFirst = m_fmt.FormatSQL(strSql, esdAccess)
    strSecond = m_fmt.FormatSQL(strFirst, esdAccess)
    TestAssert strFirst = strSecond, "formatting is idempotent"
End Sub


Public Sub TestFormatSQL_JoinQuery()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM t1 INNER JOIN t2 ON t1.id=t2.id LEFT JOIN t3 ON t2.id=t3.id"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "INNER JOIN") > 0, "preserves INNER JOIN"
    TestAssert InStr(strResult, "LEFT JOIN") > 0, "preserves LEFT JOIN"
End Sub


Public Sub TestFormatSQL_SubQuery()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM (SELECT b FROM t1) AS sub1"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "SELECT") > 0, "outer SELECT present"
    TestAssert Len(strResult) > Len("SELECT"), "produced multi-line output"
End Sub


Public Sub TestFormatSQL_UnionQuery()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM t1 UNION SELECT b FROM t2"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "UNION") > 0, "UNION preserved"
End Sub


Public Sub TestFormatSQL_EmptyString()
    TestAssert m_fmt.FormatSQL("", esdAccess) = "", "empty input returns empty"
End Sub


Public Sub TestFormatSQL_Parameters()
    Dim strSql As String
    Dim strResult As String
    strSql = "PARAMETERS [pDate] DateTime; SELECT a FROM t1 WHERE dt=[pDate]"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "PARAMETERS") > 0, "PARAMETERS preserved"
End Sub
