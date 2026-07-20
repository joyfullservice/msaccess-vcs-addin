Attribute VB_Name = "modTestScopedSync"
'---------------------------------------------------------------------------------------
' Module    : modTestScopedSync
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Regression tests for category-scoped ExportByType / ImportByType API
'           : validation and JSON result shape.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestExportByTypeRejectsUnknownType()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ExportByType("not_a_real_type")
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "unknown type fails"
    TestAssert InStr(CStr(dResult("error")), "Unknown object type") > 0, "error mentions unknown"
End Sub


Public Sub TestImportByTypeRejectsTableData()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ImportByType("table_data")
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "table_data import rejected"
    TestAssert InStr(CStr(dResult("error")), "Import not supported") > 0, "error mentions unsupported"
End Sub


Public Sub TestImportByTypeRejectsUnknownType()
    Dim strJson As String
    Dim dResult As Dictionary

    strJson = VCS.ImportByType(Null)
    Set dResult = ParseJson(strJson)
    TestAssert Not CBool(dResult("success")), "Null type fails"
End Sub
