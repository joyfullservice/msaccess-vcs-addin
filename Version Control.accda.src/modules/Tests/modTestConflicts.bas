Attribute VB_Name = "modTestConflicts"
'---------------------------------------------------------------------------------------
' Module    : modTestConflicts
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Integration tests for export conflict detection. Verifies that modifying
'           : a source file externally triggers a conflict when CheckExportConflicts runs.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


'---------------------------------------------------------------------------------------
' Procedure : TestExportConflict_DetectsModifiedSource
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Integration test for the exact bug scenario: after a build, external changes
'           : to source files should be detected as conflicts during export.
'           : 1. Pick a module via GetAllFromDB
'           : 2. Append a comment to its source file (changes content + timestamp)
'           : 3. Run CheckExportConflicts with that single item
'           : 4. Assert a conflict was detected
'           : 5. Restore the original file content
'---------------------------------------------------------------------------------------
'
Public Sub TestExportConflict_DetectsModifiedSource()

    Dim cCategory As IDbComponent
    Dim dAllModules As Dictionary
    Dim dOneItem As Dictionary
    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim strFile As String
    Dim strOriginal As String
    Dim cItem As IDbComponent

    ' Access GetAllFromDB through the IDbComponent interface (same pattern as modExport)
    Set cCategory = New clsDbModule
    Set dAllModules = cCategory.GetAllFromDB(False)
    If dAllModules.Count = 0 Then Exit Sub

    ' Pick the first module
    Set cItem = dAllModules.Items()(0)
    strFile = cItem.SourceFile

    ' Must have a source file on disk and an index entry to detect conflicts
    If Not FSO.FileExists(strFile) Then Exit Sub
    If Not VCSIndex.Exists(cItem, strFile) Then Exit Sub

    ' Save original file content
    strOriginal = ReadFile(strFile)

    On Error GoTo Cleanup

    ' Append a comment line to simulate external modification
    WriteFile strOriginal & vbCrLf & "' Test conflict marker " & Now, strFile

    ' Build a single-item dictionary (mimics what modExport does)
    Set dOneItem = New Dictionary
    dOneItem.Add strFile, cItem

    ' Initialize conflicts and run the check
    Set dCategories = New Dictionary
    Set dCategory = New Dictionary
    dCategory.Add "Class", cCategory
    dCategory.Add "Objects", dOneItem
    dCategories.Add cCategory.Category, dCategory
    VCSIndex.Conflicts.Initialize dCategories, eatExport
    VCSIndex.CheckExportConflicts dOneItem

    ' The modified source file should have been detected as a conflict
    TestAssert VCSIndex.Conflicts.Count > 0, "conflict detected for modified source file"

Cleanup:
    ' Restore original file content unconditionally
    If Len(strOriginal) > 0 Then WriteFile strOriginal, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExportConflict_UnmodifiedSourceNoConflict
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Same setup as the conflict test, but without modifying the file. Verifies
'           : that CheckExportConflicts does NOT flag a false positive.
'---------------------------------------------------------------------------------------
'
Public Sub TestExportConflict_UnmodifiedSourceNoConflict()

    Dim cCategory As IDbComponent
    Dim dAllModules As Dictionary
    Dim dOneItem As Dictionary
    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim strFile As String
    Dim cItem As IDbComponent

    ' Access GetAllFromDB through the IDbComponent interface
    Set cCategory = New clsDbModule
    Set dAllModules = cCategory.GetAllFromDB(False)
    If dAllModules.Count = 0 Then Exit Sub

    ' Pick the first module
    Set cItem = dAllModules.Items()(0)
    strFile = cItem.SourceFile

    If Not FSO.FileExists(strFile) Then Exit Sub
    If Not VCSIndex.Exists(cItem, strFile) Then Exit Sub

    ' Build a single-item dictionary (same setup, but no file modification)
    Set dOneItem = New Dictionary
    dOneItem.Add strFile, cItem

    ' Initialize conflicts and run the check
    Set dCategories = New Dictionary
    Set dCategory = New Dictionary
    dCategory.Add "Class", cCategory
    dCategory.Add "Objects", dOneItem
    dCategories.Add cCategory.Category, dCategory
    VCSIndex.Conflicts.Initialize dCategories, eatExport
    VCSIndex.CheckExportConflicts dOneItem

    ' No conflict expected for an unmodified source file
    TestAssert VCSIndex.Conflicts.Count = 0, "no conflict for unmodified source"

End Sub
