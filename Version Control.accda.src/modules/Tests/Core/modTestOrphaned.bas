Attribute VB_Name = "modTestOrphaned"
'---------------------------------------------------------------------------------------
' Module    : modTestOrphaned
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Regression tests for scoped FileExtensions artifact cleanup and folder
'           : artifact removal.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestFormAllScopeIncludesSidecars()
    Dim cForm As IDbComponent
    Dim colAll As Collection
    Dim colIndexed As Collection

    Set cForm = New clsDbForm
    Set colAll = cForm.FileExtensions(efesAll)
    Set colIndexed = cForm.FileExtensions(efesIndexed)
    TestAssert colAll.Count = colIndexed.Count + 1, "form efesAll adds svg only"
    TestAssert ExtensionInCollection("svg", colAll), "form efesAll includes svg"
    TestAssert Not ExtensionInCollection("svg", colIndexed), "form efesIndexed excludes svg"
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        TestAssert ExtensionInCollection("json", colIndexed), "form efesIndexed includes json"
        TestAssert ExtensionInCollection("json", colAll), "form efesAll includes json"
    End If
End Sub


Public Sub TestClearOrphanedComponentArtifactsRemovesSidecars()
    Dim cForm As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strFolder As String
    Dim strFake As String
    Dim varExt As Variant

    Set cForm = New clsDbForm
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keep", vbNullString

    strFolder = cForm.BaseFolder
    strFake = "vcs_test_orphan_fake"
    VerifyPath strFolder

    For Each varExt In cForm.FileExtensions(efesAll)
        WriteFile "", strFolder & strFake & "." & varExt
    Next varExt

    ClearOrphanedComponentArtifacts cForm, dBaseNames

    TestAssert Not FSO.FileExists(strFolder & strFake & ".svg"), "deleted fake svg"
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        TestAssert FSO.FileExists(strFolder & strFake & ".json"), "indexed json untouched by artifact cleanup"
        TestAssert FSO.FileExists(strFolder & strFake & ".form"), "indexed form untouched"
    Else
        TestAssert FSO.FileExists(strFolder & strFake & ".bas"), "indexed bas untouched"
    End If
    TestAssert FSO.FileExists(strFolder & strFake & ".cls"), "indexed cls untouched"

    CleanupOrphanTestFiles cForm, strFake
End Sub


Public Sub TestClearOrphanedComponentArtifactsPreservesLiveObject()
    Dim cForm As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strFolder As String
    Dim strKeep As String
    Dim varExt As Variant

    Set cForm = New clsDbForm
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    strKeep = "vcs_test_orphan_keep"
    dBaseNames.Add strKeep, vbNullString

    strFolder = cForm.BaseFolder
    VerifyPath strFolder

    For Each varExt In cForm.FileExtensions(efesAll)
        WriteFile "", strFolder & strKeep & "." & varExt
    Next varExt

    ClearOrphanedComponentArtifacts cForm, dBaseNames

    For Each varExt In cForm.FileExtensions(efesAll)
        TestAssert FSO.FileExists(strFolder & strKeep & "." & varExt), "kept live " & varExt
    Next varExt

    CleanupOrphanTestFiles cForm, strKeep
End Sub


Public Sub TestClearOrphanedArtifactFoldersRemovesOrphanCommandBarImages()
    Dim cBar As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strFolder As String
    Dim strFake As String

    Set cBar = New clsDbCommandBar
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keepbar", vbNullString

    strFolder = cBar.BaseFolder
    strFake = "vcs_test_orphan_fakebar"
    VerifyPath strFolder
    VerifyPath strFolder & strFake & "_Images"
    WriteFile "", strFolder & strFake & "_Images\orphan.bmp"

    ClearOrphanedArtifactFolders cBar, dBaseNames, "_Images"

    TestAssert Not FSO.FolderExists(strFolder & strFake & "_Images"), "orphan images folder removed"
End Sub


Public Sub TestClearOrphanedArtifactFoldersPreservesLiveCommandBarImages()
    Dim cBar As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strFolder As String
    Dim strKeep As String

    Set cBar = New clsDbCommandBar
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    strKeep = "vcs_test_orphan_keepbar"
    dBaseNames.Add strKeep, vbNullString

    strFolder = cBar.BaseFolder
    VerifyPath strFolder
    VerifyPath strFolder & strKeep & "_Images"
    WriteFile "", strFolder & strKeep & "_Images\keep.bmp"

    ClearOrphanedArtifactFolders cBar, dBaseNames, "_Images"

    TestAssert FSO.FolderExists(strFolder & strKeep & "_Images"), "live images folder preserved"

    LogUnhandledErrors
    On Error Resume Next
    FSO.DeleteFolder strFolder & strKeep & "_Images", True
    On Error GoTo 0
End Sub


Public Sub TestClearOrphanedArtifactFoldersRemovesOrphanThemeFolder()
    Dim cTheme As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strFolder As String
    Dim strFake As String

    Set cTheme = New clsDbTheme
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keeptheme", vbNullString

    strFolder = cTheme.BaseFolder
    strFake = "vcs_test_orphan_faketheme"
    VerifyPath strFolder
    VerifyPath strFolder & strFake
    WriteFile "", strFolder & strFake & "\theme.xml"

    ClearOrphanedArtifactFolders cTheme, dBaseNames

    TestAssert Not FSO.FolderExists(strFolder & strFake), "orphan theme folder removed"
End Sub


Private Sub CleanupOrphanTestFiles(cmp As IDbComponent, strBase As String)
    Dim varExt As Variant
    For Each varExt In cmp.FileExtensions(efesAll)
        DeleteFile cmp.BaseFolder & strBase & "." & varExt, True
    Next varExt
End Sub


Private Function ExtensionInCollection(strExt As String, colExts As Collection) As Boolean
    Dim varItem As Variant
    For Each varItem In colExts
        If StrComp(CStr(varItem), strExt, vbTextCompare) = 0 Then
            ExtensionInCollection = True
            Exit Function
        End If
    Next varItem
End Function
