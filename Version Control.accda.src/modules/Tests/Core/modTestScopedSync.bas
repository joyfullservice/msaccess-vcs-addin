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


Public Sub TestExportToFolderRefusesAnotherDatabasesExportFolder()

    Dim strFolder As String
    Dim dResult As Dictionary
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    ' A folder with a vcs-index.idx belongs to some database; an export to another
    ' folder runs without an index and would overwrite it unchecked.
    strFolder = GetTempFolder("vcs_export_to")
    FSO.CreateTextFile(strFolder & PathSep & "vcs-index.idx").Close

    Set dResult = ParseJson(VCS.ExportObject("module", "modExample", strFolder))
    TestAssert Not CBool(dResult("success")), "a folder holding an index is refused"
    TestAssert dResult("error_pattern") = "export_folder_has_index", "the refusal names the reason"
    TestAssert Len(Operation.ExportFolderOverride) = 0, "the refusal stages no export folder"

CleanUp:
    On Error Resume Next
    If Len(strFolder) Then FSO.DeleteFolder strFolder, True
    If lngErr <> 0 Then TestAssert False, "export folder refusal " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

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
