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


Private Sub AssertIdempotent(strFormatted As String, intDialect As eSqlDialect, strContext As String)
    TestAssert strFormatted = m_fmt.FormatSQL(strFormatted, intDialect), _
        strContext & " is idempotent"
End Sub


Public Sub TestFormatSQL_SimpleSelect()
    Dim strResult As String
    strResult = m_fmt.FormatSQL("SELECT a, b FROM tbl WHERE x=1", esdAccess)
    TestAssert Len(strResult) > 0, "returns non-empty result"
    TestAssert InStr(strResult, "SELECT") > 0, "contains SELECT keyword"
    AssertIdempotent strResult, esdAccess, "simple select"
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
    AssertIdempotent strResult, esdAccess, "join query"
End Sub


Public Sub TestFormatSQL_SubQuery()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM (SELECT b FROM t1) AS sub1"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "SELECT") > 0, "outer SELECT present"
    TestAssert Len(strResult) > Len("SELECT"), "produced multi-line output"
    AssertIdempotent strResult, esdAccess, "subquery"
End Sub


Public Sub TestFormatSQL_UnionQuery()
    Dim strSql As String
    Dim strResult As String
    strSql = "SELECT a FROM t1 UNION SELECT b FROM t2"
    strResult = m_fmt.FormatSQL(strSql, esdAccess)
    TestAssert InStr(strResult, "UNION") > 0, "UNION preserved"
    AssertIdempotent strResult, esdAccess, "union query"
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
    AssertIdempotent strResult, esdAccess, "parameters"
End Sub


Public Sub TestFormatSQL_LastCharPreserved()
    Dim strResult As String
    strResult = m_fmt.FormatSQL("SELECT a FROM t1 WHERE x = 5", esdAccess)
    TestAssert InStr(strResult, "= 5") > 0, "trailing single-char value preserved"
    AssertIdempotent strResult, esdAccess, "trailing value"
    strResult = m_fmt.FormatSQL("SELECT a FROM t1 ORDER BY b", esdAccess)
    TestAssert InStr(strResult, "b") > 0, "trailing single-char column preserved"
    AssertIdempotent strResult, esdAccess, "trailing column"
End Sub


Public Sub TestFormatSQL_SemicolonPreserved()
    Dim strResult As String
    strResult = m_fmt.FormatSQL("SELECT a FROM t1;", esdAccess)
    TestAssert Right$(strResult, 1) = ";", "trailing semicolon preserved"
    AssertIdempotent strResult, esdAccess, "semicolon"
End Sub


Public Sub TestFormatSQL_MySqlSimple()
    Dim strExpected As String
    Dim strResult As String
    strExpected = "SELECT" & vbCrLf & "  5 AS `TEST`"
    strResult = m_fmt.FormatSQL("SELECT  5 AS `TEST`", esdMySQL)
    TestAssert strResult = strExpected, "MySQL simple select formatted"
    AssertIdempotent strResult, esdMySQL, "MySQL simple"
End Sub


Public Sub TestFormatSQL_AccessDateLiteral()
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT (#1/1/2000#) AS SampleDate # MySQL inline ## comment", esdAccess)
    TestAssert InStr(strResult, "#1/1/2000#") > 0, "date literal preserved"
    TestAssert InStr(strResult, "# MySQL inline ## comment") > 0, "comment preserved"
    AssertIdempotent strResult, esdAccess, "date literal"
End Sub


Public Sub TestFormatSQL_MultiPartNames()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  [dbo].[field]"
        .Add "FROM"
        .Add "  [server].[schema].[table];"
        .Remove 2
        strExpected = .GetStr
    End With
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT [dbo].[field] FROM [server].[schema].[table];", esdMSSQL)
    TestAssert strResult = strExpected, "multi-part names with semicolon"
    AssertIdempotent strResult, esdMSSQL, "multi-part names"
End Sub


Public Sub TestFormatSQL_BangNotation()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  [Forms]![frmColors]![Text18];"
        .Remove 2
        strExpected = .GetStr
    End With
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT [Forms]![frmColors]![Text18];", esdAccess)
    TestAssert strResult = strExpected, "Access bang notation with semicolon"
    AssertIdempotent strResult, esdAccess, "bang notation"
End Sub


Public Sub TestFormatSQL_UnquotedPath()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  foo.* INTO (C:\path\to\bar.accdb) fizz"
        .Add "FROM"
        .Add "  bazz;"
        .Remove 2
        strExpected = .GetStr
    End With
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT foo.* INTO (C:\path\to\bar.accdb) fizz FROM bazz;", esdAccess)
    TestAssert strResult = strExpected, "unquoted path query"
    AssertIdempotent strResult, esdAccess, "unquoted path"
End Sub


Public Sub TestFormatSQL_NestedWhere()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  MSysQueries.Attribute,"
        .Add "  MSysQueries.Flag"
        .Add "FROM"
        .Add "  MSysQueries"
        .Add "WHERE"
        .Add "  ("
        .Add "    ("
        .Add "      (MSysQueries.Flag) In ("
        .Add "        SELECT"
        .Add "          Flag"
        .Add "        from"
        .Add "          [MSysQueries]"
        .Add "        where"
        .Add "          flag = 0"
        .Add "      )"
        .Add "    )"
        .Add "  );"
        .Remove 2
        strExpected = .GetStr
    End With
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT MSysQueries.Attribute, MSysQueries.Flag FROM MSysQueries" & _
        " WHERE (((MSysQueries.Flag) In (SELECT Flag from [MSysQueries] where flag=0)));", _
        esdAccess)
    TestAssert strResult = strExpected, "nested where clause"
    AssertIdempotent strResult, esdAccess, "nested where"
End Sub


Public Sub TestFormatSQL_DoctrineExample()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  count(*),"
        .Add "  `Column1`,"
        .Add "  `Testing`,"
        .Add "  `Testing Three`"
        .Add "FROM"
        .Add "  `Table1`"
        .Add "WHERE"
        .Add "  Column1 = 'testing'"
        .Add "  AND ("
        .Add "    ("
        .Add "      `Column2` = `Column3`"
        .Add "      OR Column4 >= NOW()"
        .Add "    )"
        .Add "  )"
        .Add "GROUP BY"
        .Add "  Column1"
        .Add "ORDER BY"
        .Add "  Column3 DESC"
        .Add "LIMIT"
        .Add "  5, 10"
        .Remove 2
        strExpected = .GetStr
    End With
    Dim strResult As String
    strResult = m_fmt.FormatSQL( _
        "SELECT count(*),`Column1`,`Testing`, `Testing Three` FROM `Table1`" & _
        " WHERE Column1 = 'testing' AND ( (`Column2` = `Column3` OR Column4 >= NOW()) )" & _
        " GROUP BY Column1 ORDER BY Column3 DESC LIMIT 5,10", esdMySQL)
    TestAssert strResult = strExpected, "doctrine example query"
    AssertIdempotent strResult, esdMySQL, "doctrine example"
End Sub
