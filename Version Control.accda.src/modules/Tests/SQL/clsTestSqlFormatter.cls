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
'           : produces the same result). The TestFormatSQL_ tests pin export format
'           : 5.1.0, whose output must never change; TestFormat52_ tests cover the
'           : layout introduced in 5.2.0.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

Private m_fmt As clsSqlFormatter


Private Sub Class_Initialize()
    Set m_fmt = New clsSqlFormatter
    m_fmt.FormatVersion = EFV_5_1_0
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


Public Sub TestFormatSQL_ComparisonOperatorSpacing()
    Dim strResult As String

    ' Closing paren followed by comparison operator should have a space between
    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE ((t1.ID)>0)", esdAccess)
    TestAssert InStr(strResult, ") > 0") > 0, ") > 0 has space before >"
    AssertIdempotent strResult, esdAccess, "paren-gt spacing"

    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE ((t1.ID)<100)", esdAccess)
    TestAssert InStr(strResult, ") < 100") > 0, ") < 100 has space before <"
    AssertIdempotent strResult, esdAccess, "paren-lt spacing"

    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE ((t1.ID)<>5)", esdAccess)
    TestAssert InStr(strResult, ") <> 5") > 0, ") <> 5 has space before <>"
    AssertIdempotent strResult, esdAccess, "paren-neq spacing"

    ' Compound operators (>=, <=) must stay glued
    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE a >= 5", esdAccess)
    TestAssert InStr(strResult, ">= 5") > 0, ">= stays together"
    AssertIdempotent strResult, esdAccess, "gte compound"

    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE a <= 10", esdAccess)
    TestAssert InStr(strResult, "<= 10") > 0, "<= stays together"
    AssertIdempotent strResult, esdAccess, "lte compound"
End Sub


Public Sub TestFormatSQL_ArithmeticOperatorSpacing()
    Dim strResult As String

    ' Binary minus after closing paren should keep spaces on both sides
    strResult = m_fmt.FormatSQL( _
        "SELECT Year([Date])-1 AS PriorYear FROM t1", esdAccess)
    TestAssert InStr(strResult, "Year([Date]) - 1") > 0, _
        "binary minus has spaces: ) - 1"
    AssertIdempotent strResult, esdAccess, "binary minus"

    ' Unary minus (negative number) should still glue to the number
    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE x = -5", esdAccess)
    TestAssert InStr(strResult, "= -5") > 0, _
        "unary minus glued to number: = -5"
    AssertIdempotent strResult, esdAccess, "unary minus"

    ' Negative number after opening paren
    strResult = m_fmt.FormatSQL( _
        "SELECT a FROM t1 WHERE x IN (-1, -2)", esdAccess)
    TestAssert InStr(strResult, "(-1") > 0, "negative after open paren: (-1"
    AssertIdempotent strResult, esdAccess, "neg after open paren"
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


Private Function Format52(strSql As String) As String
    With New clsSqlFormatter
        .FormatVersion = EFV_5_2_0
        Format52 = .FormatSQL(strSql, esdAccess)
    End With
End Function


Private Sub AssertFormat52(strSql As String, strExpected As String, strContext As String)
    Dim strResult As String
    strResult = Format52(strSql)
    TestAssert strResult = strExpected, strContext
    TestAssert Format52(strResult) = strResult, strContext & " is idempotent"
End Sub


' The type and text of every token that is not whitespace, one per line.
Private Function SignificantTokens(strSql As String) As String
    Dim varToken As Variant
    Dim cTokens As clsConcat
    Set cTokens = New clsConcat
    With New clsSqlFormatter
        .Tokenize strSql, esdAccess
        For Each varToken In .Tokens
            If varToken(0) <> ttWhitespace Then cTokens.Add CStr(varToken(0)), "|", CStr(varToken(1)), vbLf
        Next varToken
    End With
    SignificantTokens = cTokens.GetStr
End Function


Public Sub TestFormat52_GroupsThatFitStayInline()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  Nz(t.a, 0) AS x,"
        .Add "  IIf(t.b > 1, 'y', 'n') AS y"
        .Add "FROM"
        .Add "  t"
        .Add "WHERE"
        .Add "  t.c In (1, 2, 3)"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT Nz(t.a, 0) AS x, IIf(t.b > 1, 'y', 'n') AS y FROM t WHERE t.c In (1, 2, 3)", _
        strExpected, "function calls and IN lists that fit stay on one line"
End Sub


Public Sub TestFormat52_LongCallOneArgumentPerLine()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  IIf("
        .Add "    tblCars.Manufacturer = 'Ford',"
        .Add "    tblCarsPrice.Price * 1.05,"
        .Add "    IIf(tblCars.Manufacturer = 'Audi', tblCarsPrice.Price * 1.1, tblCarsPrice.Price)"
        .Add "  ) AS NewPrice"
        .Add "FROM"
        .Add "  tblCars"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT IIf(tblCars.Manufacturer = 'Ford', tblCarsPrice.Price * 1.05, " & _
        "IIf(tblCars.Manufacturer = 'Audi', tblCarsPrice.Price * 1.1, tblCarsPrice.Price)) AS NewPrice FROM tblCars", _
        strExpected, "a call too wide for the line puts one argument per line; a nested call that fits stays inline"
End Sub


Public Sub TestFormat52_FlatJoinChain()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  ((tblA"
        .Add "  INNER JOIN tblB"
        .Add "    ON tblA.ID = tblB.AID)"
        .Add "  INNER JOIN tblC"
        .Add "    ON tblB.ID = tblC.BID)"
        .Add "  LEFT JOIN tblD"
        .Add "    ON tblC.ID = tblD.CID"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM ((tblA INNER JOIN tblB ON tblA.ID = tblB.AID) " & _
        "INNER JOIN tblC ON tblB.ID = tblC.BID) LEFT JOIN tblD ON tblC.ID = tblD.CID", _
        strExpected, "left-nested join groups read as one join per line"
End Sub


Public Sub TestFormat52_NestedJoinOnTheRightIsIndented()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  tblC,"
        .Add "  tblA"
        .Add "  INNER JOIN ("
        .Add "    tblB"
        .Add "    INNER JOIN tblD"
        .Add "      ON tblB.ID = tblD.BID"
        .Add "  )"
        .Add "    ON tblA.ID = tblB.AID"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM tblC, tblA INNER JOIN (tblB INNER JOIN tblD ON tblB.ID = tblD.BID) ON tblA.ID = tblB.AID", _
        strExpected, "a join group on the right of a join is an indented block"
End Sub


Public Sub TestFormat52_OnConditionsIndented()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  tblA"
        .Add "  INNER JOIN tblB"
        .Add "    ON (tblA.ID = tblB.AID)"
        .Add "    AND (tblA.X = tblB.X)"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM tblA INNER JOIN tblB ON (tblA.ID = tblB.AID) AND (tblA.X = tblB.X)", _
        strExpected, "ON and each further condition sit one level under the join"
End Sub


Public Sub TestFormat52_MultilineAndOuterParensInOn()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  tblOrders"
        .Add "  LEFT JOIN tblRates"
        .Add "    ON tblOrders.RateCode = IIf("
        .Add "      tblOrders.IsInternational = True,"
        .Add "      tblOrders.IntlRateCode,"
        .Add "      tblOrders.DomesticRateCode"
        .Add "    )"
        .Add "WHERE"
        .Add "  tblOrders.Active = True"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM tblOrders LEFT JOIN tblRates " & _
        "ON tblOrders.RateCode = IIf(tblOrders.IsInternational = True, tblOrders.IntlRateCode, tblOrders.DomesticRateCode) " & _
        "WHERE tblOrders.Active = True", _
        strExpected, "a broken call inside ON indents under ON and closes at the ON indent"

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  tblA"
        .Add "  INNER JOIN tblB"
        .Add "    ON ("
        .Add "      (tblA.ID = tblB.AID)"
        .Add "      AND (tblB.Active = True)"
        .Add "    )"
        .Add "  INNER JOIN tblC"
        .Add "    ON tblB.ID = tblC.BID,"
        .Add "  tblD"
        .Add "WITH OWNERACCESS OPTION;"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM tblA INNER JOIN tblB ON ((tblA.ID = tblB.AID) AND (tblB.Active = True)) " & _
        "INNER JOIN tblC ON tblB.ID = tblC.BID, tblD WITH OWNERACCESS OPTION;", _
        strExpected, "an outer-parenthesized compound ON puts one condition per line and ends cleanly before the next join, comma, or clause"
End Sub


Public Sub TestFormat52_SemicolonResetsStatementIndent()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "PARAMETERS"
        .Add "  [Enter Year] Short;"
        .Add "INSERT INTO tblA ("
        .Add "  a,"
        .Add "  [Year]"
        .Add ")"
        .Add "SELECT"
        .Add "  u.a,"
        .Add "  [Enter Year]"
        .Add "FROM"
        .Add "  u;"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "PARAMETERS [Enter Year] Short; INSERT INTO tblA (a, [Year]) SELECT u.a, [Enter Year] FROM u;", _
        strExpected, "PARAMETERS semicolon resets indent and starts INSERT INTO on a new line"

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "CREATE TABLE tblTemp ("
        .Add "  ID LONG,"
        .Add "  Name TEXT(50)"
        .Add ");"
        .Add "CREATE UNIQUE INDEX idxAB ON tblTemp ("
        .Add "  ID,"
        .Add "  Name"
        .Add ");"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "CREATE TABLE tblTemp (ID LONG, Name TEXT(50)); CREATE UNIQUE INDEX idxAB ON tblTemp (ID, Name);", _
        strExpected, "multi-statement DDL starts each statement flush left on a new line"
End Sub


Public Sub TestFormat52_ColumnListsOnePerLine()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "INSERT INTO tblA ("
        .Add "  a,"
        .Add "  [Year]"
        .Add ")"
        .Add "SELECT"
        .Add "  u.a,"
        .Add "  2019 AS [Year]"
        .Add "FROM"
        .Add "  u"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "INSERT INTO tblA (a, [Year]) SELECT u.a, 2019 AS [Year] FROM u", _
        strExpected, "an INSERT INTO column list puts one column per line even when it fits"

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "CREATE TABLE tblTemp ("
        .Add "  ID LONG,"
        .Add "  Name TEXT(50)"
        .Add ");"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "CREATE TABLE tblTemp (ID LONG, Name TEXT(50));", _
        strExpected, "a CREATE TABLE column list puts one column per line; a type size stays inline"
End Sub


Public Sub TestFormat52_CreateIndexOnIsNotAJoin()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "CREATE UNIQUE INDEX idxAB ON tblA ("
        .Add "  a,"
        .Add "  b"
        .Add ");"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "CREATE UNIQUE INDEX idxAB ON tblA (a, b);", _
        strExpected, "the ON of CREATE INDEX stays on its line and the index columns go one per line"
End Sub


Public Sub TestFormat52_ParametersAndDeleteAreClauses()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "PARAMETERS"
        .Add "  [Enter ID] Long,"
        .Add "  [Enter Category] Text (255);"
        .Add "DELETE DISTINCTROW"
        .Add "  t.*,"
        .Add "  u.Code"
        .Add "FROM"
        .Add "  t"
        .Add "WHERE"
        .Add "  t.ID = [Enter ID];"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "PARAMETERS [Enter ID] Long, [Enter Category] Text (255); " & _
        "DELETE DISTINCTROW t.*, u.Code FROM t WHERE t.ID = [Enter ID];", _
        strExpected, "PARAMETERS and DELETE put one indented item per line"
End Sub


Public Sub TestFormat52_MakeTableIntoIsAClause()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  t.ID,"
        .Add "  t.Name"
        .Add "INTO"
        .Add "  tblCopy IN 'C:\path\to\other.accdb'"
        .Add "FROM"
        .Add "  t"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT t.ID, t.Name INTO tblCopy IN 'C:\path\to\other.accdb' FROM t", _
        strExpected, "a make-table INTO starts its own clause"
End Sub


Public Sub TestFormat52_OwnerAccessOptionOneLine()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  t.ID"
        .Add "FROM"
        .Add "  t"
        .Add "WITH OWNERACCESS OPTION;"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT t.ID FROM t WITH OWNERACCESS OPTION;", _
        strExpected, "WITH OWNERACCESS OPTION stays on one line"
End Sub


Public Sub TestFormat52_BetweenAndStaysOnLine()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  t"
        .Add "WHERE"
        .Add "  t.d Between #1/1/2020# And #12/31/2020#"
        .Add "  AND t.x Between DateAdd('d', -1, t.y) And t.y"
        .Add "  AND t.z = 1"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM t WHERE t.d Between #1/1/2020# And #12/31/2020# " & _
        "AND t.x Between DateAdd('d', -1, t.y) And t.y AND t.z = 1", _
        strExpected, "the And of Between ... And never starts a line"
End Sub


Public Sub TestFormat52_OperatorSpacingAndUnaryMinus()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  [a] & [b] AS c,"
        .Add "  x ^ 2 AS e,"
        .Add "  -Nz([z], 0) AS f,"
        .Add "  -[w] AS g"
        .Add "FROM"
        .Add "  t"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT [a]&[b] AS c, x^2 AS e, -Nz([z], 0) AS f, -[w] AS g FROM t", _
        strExpected, "& and ^ are spaced; a unary minus stays on its operand"
End Sub


Public Sub TestFormat52_SubqueryBreaks()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  t"
        .Add "WHERE"
        .Add "  t.x In ("
        .Add "    SELECT"
        .Add "      y"
        .Add "    FROM"
        .Add "      u"
        .Add "    WHERE"
        .Add "      u.z = 1"
        .Add "  )"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM t WHERE t.x In (SELECT y FROM u WHERE u.z = 1)", _
        strExpected, "a subquery never stays inline"
End Sub


Public Sub TestFormat52_ConditionGroupInline()
    Dim strExpected As String
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  a"
        .Add "FROM"
        .Add "  t"
        .Add "WHERE"
        .Add "  (t.a = 1 OR t.b = 2)"
        .Add "  AND t.c = 3"
        .Remove 2
        strExpected = .GetStr
    End With
    AssertFormat52 "SELECT a FROM t WHERE (t.a = 1 OR t.b = 2) AND t.c = 3", _
        strExpected, "a parenthesized condition that fits stays on one line"
End Sub


Public Sub TestFormat52_TokensPreservedForTestInputs()
    Dim varSql As Variant
    For Each varSql In Array( _
        "SELECT Nz(t.a, 0) AS x, IIf(t.b > 1, 'y', 'n') AS y FROM t WHERE t.c In (1, 2, 3)", _
        "SELECT a FROM ((tblA INNER JOIN tblB ON tblA.ID = tblB.AID) INNER JOIN tblC ON tblB.ID = tblC.BID) LEFT JOIN tblD ON tblC.ID = tblD.CID", _
        "SELECT [a]&[b] AS c, x\y AS d, x^2 AS e, -Nz([z], 0) AS f, -[w] AS g, a - -1 AS h FROM t", _
        "SELECT a FROM t WHERE t.x In (SELECT y FROM u WHERE u.z = 1) ORDER BY a DESC , b", _
        "PARAMETERS [pDate] DateTime; SELECT a FROM t1 WHERE dt=[pDate] AND x Like ""a*"" AND d = #1/1/2000#;", _
        "INSERT INTO t ( a, b ) SELECT u.a, u.b FROM u WHERE (((u.a)=1));", _
        "TRANSFORM Count(t.a) AS c SELECT t.b FROM t GROUP BY t.b PIVOT t.d In (1, 2);", _
        "PARAMETERS [p1] Long, [p2] Text (255); DELETE DISTINCTROW t.*, u.Code FROM t WHERE t.ID = [p1];", _
        "CREATE TABLE tblTemp (ID LONG, Name TEXT(50)); CREATE UNIQUE INDEX idxAB ON tblA (a, b);", _
        "PARAMETERS [Enter Year] Short; INSERT INTO tblA (a, [Year]) SELECT u.a, [Enter Year] FROM u;", _
        "SELECT a FROM tblA INNER JOIN tblB ON ((tblA.ID = tblB.AID) AND (tblB.Active = True)) INNER JOIN tblC ON tblB.ID = tblC.BID;", _
        "SELECT t.ID, t.Name INTO tblCopy IN 'C:\path\to\other.accdb' FROM t WITH OWNERACCESS OPTION;", _
        "INSERT INTO tblA (a, b) IN 'C:\path\to\other.accdb' SELECT u.a, u.b FROM u IN '' [ODBC;DSN=Example;];", _
        "SELECT a FROM t WHERE t.d Between #1/1/2020# And #12/31/2020# AND t.x = 1")
        TestAssert SignificantTokens(Format52(CStr(varSql))) = SignificantTokens(CStr(varSql)), _
            "tokens preserved: " & Left$(varSql, 40)
        TestAssert Format52(Format52(CStr(varSql))) = Format52(CStr(varSql)), _
            "idempotent: " & Left$(varSql, 40)
    Next varSql
End Sub


Public Sub TestFormat52_TokensPreservedForCorpus()
    Dim strRoot As String
    Dim lngFiles As Long
    strRoot = FSO.BuildPath(CurrentProject.Path, "Testing\Fixtures\queries")
    TestAssert FSO.FolderExists(strRoot), "fixture corpus found at " & strRoot
    If Not FSO.FolderExists(strRoot) Then Exit Sub
    CheckCorpusFolder FSO.GetFolder(strRoot), lngFiles
    TestAssert lngFiles > 0, "corpus contains .sql files"
End Sub


' Pass-through fixtures are server SQL, which the export never formats.
Private Sub CheckCorpusFolder(oFolder As Scripting.Folder, ByRef lngFiles As Long)
    Dim oFile As Scripting.File
    Dim oSub As Scripting.Folder
    Dim strSql As String
    Dim strFormatted As String
    If StrComp(oFolder.Name, "passthrough", vbTextCompare) = 0 Then Exit Sub
    For Each oFile In oFolder.Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "sql", vbTextCompare) = 0 Then
            lngFiles = lngFiles + 1
            strSql = ReadFile(oFile.Path)
            strFormatted = Format52(strSql)
            TestAssert SignificantTokens(strFormatted) = SignificantTokens(strSql), _
                "tokens preserved: " & oFile.Name
            TestAssert Format52(strFormatted) = strFormatted, "idempotent: " & oFile.Name
        End If
    Next oFile
    For Each oSub In oFolder.SubFolders
        CheckCorpusFolder oSub, lngFiles
    Next oSub
End Sub
