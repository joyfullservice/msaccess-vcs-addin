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


Public Sub TestRawSQL_Property()
    m_qc.RawSQL = "SELECT 42"
    TestAssert m_qc.RawSQL = "SELECT 42", "RawSQL round trip"
End Sub
