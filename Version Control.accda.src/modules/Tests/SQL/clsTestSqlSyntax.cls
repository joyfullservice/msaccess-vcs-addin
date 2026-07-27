VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSqlSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestSqlSyntax
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Unit tests for clsSqlSyntax table-ref and ON-side parsing helpers.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

Private m_syntax As clsSqlSyntax


Private Sub Class_Initialize()
    Set m_syntax = New clsSqlSyntax
End Sub


Private Sub Class_Terminate()
    Set m_syntax = Nothing
End Sub


Public Sub TestSplitTableRef_NameShapes()
    AssertSplitTableRef "tblPlain", "tblPlain", vbNullString, "tblPlain"
    AssertSplitTableRef "[Car Models]", "Car Models", vbNullString, "Car Models"
    AssertSplitTableRef "[Car Models] AS cm", "Car Models", "cm", "Car Models"
    AssertSplitTableRef "[Table AS Thing]", "Table AS Thing", vbNullString, "Table AS Thing"
    AssertSplitTableRef "[Order Details]", "Order Details", vbNullString, "Order Details"
    AssertSplitTableRef "[Sales.2024]", "Sales.2024", vbNullString, "Sales.2024"
End Sub


Public Sub TestTryExtractSimpleTable_NameShapes()
    AssertTryExtract "tblPlain", "tblPlain", "plain"
    AssertTryExtract "[Car Models]", "Car Models", "spaced bracketed name"
    AssertTryExtract "[Car Models] AS cm", "cm", "spaced name with alias"
    AssertTryExtract "[Table AS Thing]", "Table AS Thing", "embedded AS in brackets"
    AssertTryExtract "[Order Details]", "Order Details", "order details"
    AssertTryExtract "[Sales.2024]", "Sales.2024", "embedded dot in brackets"
    AssertTryExtract "tblA INNER JOIN tblB ON a=b", vbNullString, "nested join returns empty"
    AssertTryExtract "[My JOIN Table]", "My JOIN Table", "embedded JOIN in brackets"
End Sub


Public Sub TestExtractTableFromOnSide_DotOutsideBrackets()
    TestAssert m_syntax.ExtractTableFromOnSide("[Sales.2024].[Field] = tblB.x", True) = "Sales.2024", _
        "bracketed table with dot"
    TestAssert m_syntax.ExtractTableFromOnSide("tblA.x = [Sales.2024].[Field]", False) = "Sales.2024", _
        "bracketed table on right side"
End Sub


Private Sub AssertSplitTableRef(ByVal strExpr As String, _
                                ByVal strExpName As String, _
                                ByVal strExpAlias As String, _
                                ByVal strContext As String)
    Dim strName As String
    Dim strAlias As String
    m_syntax.SplitTableRef strExpr, strName, strAlias
    TestAssert strName = strExpName, strContext & " name"
    TestAssert strAlias = strExpAlias, strContext & " alias"
    TestAssert m_syntax.TableRefKey(strName, strAlias) = IIf(Len(strExpAlias) > 0, strExpAlias, strExpName), _
        strContext & " key"
End Sub


Private Sub AssertTryExtract(ByVal strExpr As String, ByVal strExpected As String, ByVal strContext As String)
    TestAssert m_syntax.TryExtractSimpleTable(strExpr) = strExpected, strContext
End Sub
