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

' WriteFile deletes the target when handed an empty string, so fixture files must
' carry content to exist on disk.
Private Const FixtureContent As String = "vcs orphan test fixture"


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
    Dim strSavedFolder As String
    Dim strRoot As String
    Dim strFolder As String
    Dim strFake As String
    Dim varExt As Variant

    strSavedFolder = BeginOrphanSandbox(strRoot)
    On Error GoTo CleanUp

    Set cForm = New clsDbForm
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keep", vbNullString

    strFolder = cForm.BaseFolder
    strFake = "vcs_test_orphan_fake"
    VerifyPath strFolder

    For Each varExt In cForm.FileExtensions(efesAll)
        WriteFile FixtureContent, strFolder & strFake & "." & varExt
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

CleanUp:
    EndOrphanSandbox strSavedFolder, strRoot
End Sub


Public Sub TestClearOrphanedComponentArtifactsPreservesLiveObject()
    Dim cForm As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strSavedFolder As String
    Dim strRoot As String
    Dim strFolder As String
    Dim strKeep As String
    Dim varExt As Variant

    strSavedFolder = BeginOrphanSandbox(strRoot)
    On Error GoTo CleanUp

    Set cForm = New clsDbForm
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    strKeep = "vcs_test_orphan_keep"
    dBaseNames.Add strKeep, vbNullString

    strFolder = cForm.BaseFolder
    VerifyPath strFolder

    For Each varExt In cForm.FileExtensions(efesAll)
        WriteFile FixtureContent, strFolder & strKeep & "." & varExt
    Next varExt

    ClearOrphanedComponentArtifacts cForm, dBaseNames

    For Each varExt In cForm.FileExtensions(efesAll)
        TestAssert FSO.FileExists(strFolder & strKeep & "." & varExt), "kept live " & varExt
    Next varExt

CleanUp:
    EndOrphanSandbox strSavedFolder, strRoot
End Sub


Public Sub TestClearOrphanedArtifactFoldersRemovesOrphanCommandBarImages()
    Dim cBar As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strSavedFolder As String
    Dim strRoot As String
    Dim strFolder As String
    Dim strFake As String

    strSavedFolder = BeginOrphanSandbox(strRoot)
    On Error GoTo CleanUp

    Set cBar = New clsDbCommandBar
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keepbar", vbNullString

    strFolder = cBar.BaseFolder
    strFake = "vcs_test_orphan_fakebar"
    VerifyPath strFolder
    VerifyPath strFolder & strFake & "_Images" & PathSep
    WriteFile FixtureContent, strFolder & strFake & "_Images" & PathSep & "orphan.bmp"

    ClearOrphanedArtifactFolders cBar, dBaseNames, "_Images"

    TestAssert Not FSO.FolderExists(strFolder & strFake & "_Images"), "orphan images folder removed"

CleanUp:
    EndOrphanSandbox strSavedFolder, strRoot
End Sub


Public Sub TestClearOrphanedArtifactFoldersPreservesLiveCommandBarImages()
    Dim cBar As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strSavedFolder As String
    Dim strRoot As String
    Dim strFolder As String
    Dim strKeep As String

    strSavedFolder = BeginOrphanSandbox(strRoot)
    On Error GoTo CleanUp

    Set cBar = New clsDbCommandBar
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    strKeep = "vcs_test_orphan_keepbar"
    dBaseNames.Add strKeep, vbNullString

    strFolder = cBar.BaseFolder
    VerifyPath strFolder
    VerifyPath strFolder & strKeep & "_Images" & PathSep
    WriteFile FixtureContent, strFolder & strKeep & "_Images" & PathSep & "keep.bmp"

    ClearOrphanedArtifactFolders cBar, dBaseNames, "_Images"

    TestAssert FSO.FolderExists(strFolder & strKeep & "_Images"), "live images folder preserved"

CleanUp:
    EndOrphanSandbox strSavedFolder, strRoot
End Sub


Public Sub TestClearOrphanedArtifactFoldersRemovesOrphanThemeFolder()
    Dim cTheme As IDbComponent
    Dim dBaseNames As Dictionary
    Dim strSavedFolder As String
    Dim strRoot As String
    Dim strFolder As String
    Dim strFake As String

    strSavedFolder = BeginOrphanSandbox(strRoot)
    On Error GoTo CleanUp

    Set cTheme = New clsDbTheme
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    dBaseNames.Add "vcs_test_orphan_keeptheme", vbNullString

    strFolder = cTheme.BaseFolder
    strFake = "vcs_test_orphan_faketheme"
    VerifyPath strFolder
    VerifyPath strFolder & strFake & PathSep
    WriteFile FixtureContent, strFolder & strFake & PathSep & "theme.xml"

    ClearOrphanedArtifactFolders cTheme, dBaseNames

    TestAssert Not FSO.FolderExists(strFolder & strFake), "orphan theme folder removed"

CleanUp:
    EndOrphanSandbox strSavedFolder, strRoot
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BeginOrphanSandbox
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Redirect the export folder to a scratch folder so orphan-cleanup fixtures
'           : never write into (or delete from) the live source tree. Returns the prior
'           : ExportFolder value and reports the scratch root through strRoot.
'---------------------------------------------------------------------------------------
'
Private Function BeginOrphanSandbox(ByRef strRoot As String) As String
    BeginOrphanSandbox = Options.ExportFolder
    strRoot = GetTempFolder("vcs_orphan") & PathSep
    Options.ExportFolder = strRoot
End Function


'---------------------------------------------------------------------------------------
' Procedure : EndOrphanSandbox
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Restore the export folder and discard the scratch folder. Restoring the
'           : option matters more than the cleanup, so it happens first.
'---------------------------------------------------------------------------------------
'
Private Sub EndOrphanSandbox(strSavedFolder As String, strRoot As String)
    Options.ExportFolder = strSavedFolder
    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strRoot) Then FSO.DeleteFolder StripSlash(strRoot), True
    Err.Clear
    On Error GoTo 0
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
