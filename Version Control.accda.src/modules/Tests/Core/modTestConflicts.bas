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
'@Tag("integration")


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

    On Error GoTo CleanUp

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

CleanUp:
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


'---------------------------------------------------------------------------------------
' Procedure : TestTableDefSourceFile_ResetsLinkTypeCache
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Full build reuses one clsDbTableDef instance. After binding a local table,
'           : binding a linked table on the same instance must still resolve .json paths.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDefSourceFile_ResetsLinkTypeCache()

    Dim cTable As IDbComponent
    Dim tdf As AccessObject
    Dim strLocalFile As String
    Dim strLinkedFile As String
    Dim strLocalName As String
    Dim strLinkedName As String

    For Each tdf In CurrentData.AllTables
        If tdf.Name Like "MSys*" Or tdf.Name Like "~*" Then
            ' Skip system tables
        Else
            Set cTable = New clsDbTableDef
            Set cTable.DbObject = tdf
            If cTable.SourceFile Like "*.xml" Then
                strLocalFile = cTable.SourceFile
                strLocalName = tdf.Name
            ElseIf cTable.SourceFile Like "*.json" Then
                strLinkedFile = cTable.SourceFile
                strLinkedName = tdf.Name
            End If
            If Len(strLocalFile) > 0 And Len(strLinkedFile) > 0 Then Exit For
        End If
    Next tdf

    If Len(strLocalFile) = 0 Or Len(strLinkedFile) = 0 Then Exit Sub

    Set cTable = New clsDbTableDef
    Set cTable.DbObject = CurrentData.AllTables(strLocalName)
    TestAssert cTable.SourceFile Like "*.xml", "local table uses .xml source path"

    Set cTable.DbObject = CurrentData.AllTables(strLinkedName)
    TestAssert cTable.SourceFile Like "*.json", _
        "linked table uses .json after local table on same instance"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExportConflict_LegacyTableDefXmlIndexKey
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Pre-fix full builds could index linked tables under .xml keys. Export must
'           : not false-positive when only a legacy .xml index entry exists.
'---------------------------------------------------------------------------------------
'
Public Sub TestExportConflict_LegacyTableDefXmlIndexKey()
'@Tag("integration")

    Dim cCategory As IDbComponent
    Dim dAll As Dictionary
    Dim varKey As Variant
    Dim cItem As IDbComponent
    Dim strJsonFile As String
    Dim strXmlFile As String
    Dim dOneItem As Dictionary
    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim blnHadJson As Boolean
    Dim strSavedHash As String
    Dim strSavedOther As String
    Dim strSavedMeta As String
    Dim dteSavedImport As Date
    Dim dteSavedExport As Date
    Dim dteSavedSourceMod As Date

    Set cCategory = New clsDbTableDef
    Set dAll = cCategory.GetAllFromDB(False)

    For Each varKey In dAll.Keys
        strJsonFile = CStr(varKey)
        If Not (strJsonFile Like "*.json") Then GoTo NextTable
        If Not FSO.FileExists(strJsonFile) Then GoTo NextTable

        Set cItem = dAll(strJsonFile)
        strXmlFile = cCategory.BaseFolder & FSO.GetBaseName(strJsonFile) & ".xml"

        blnHadJson = VCSIndex.Exists(cItem, strJsonFile)
        If blnHadJson Then
            With VCSIndex.Item(cItem, strJsonFile)
                strSavedHash = .FileHash
                strSavedOther = .OtherHash
                strSavedMeta = .MetaHash
                dteSavedImport = .ImportDate
                dteSavedExport = .ExportDate
                dteSavedSourceMod = .SourceModified
            End With
            VCSIndex.Remove cItem, strJsonFile
        End If

        With VCSIndex.Item(cItem, strXmlFile)
            .FilePropertiesHash = GetSourceFilesPropertyHash(cItem)
            .ImportDate = Now
            .SourceModified = GetSourceModifiedDate(cItem)
        End With

        Set dOneItem = New Dictionary
        dOneItem.Add strJsonFile, cItem
        Set dCategories = New Dictionary
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        dCategory.Add "Objects", dOneItem
        dCategories.Add cCategory.Category, dCategory
        VCSIndex.Conflicts.Initialize dCategories, eatExport
        VCSIndex.CheckExportConflicts dOneItem

        TestAssert VCSIndex.Conflicts.Count = 0, _
            "legacy .xml index key does not false-positive export conflict"

        VCSIndex.Remove cItem, strXmlFile
        If blnHadJson Then
            With VCSIndex.Item(cItem, strJsonFile)
                .FileHash = strSavedHash
                .OtherHash = strSavedOther
                .MetaHash = strSavedMeta
                .ImportDate = dteSavedImport
                .ExportDate = dteSavedExport
                .SourceModified = dteSavedSourceMod
                .FilePropertiesHash = GetSourceFilesPropertyHash(cItem)
            End With
        End If
        Exit Sub

NextTable:
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestModuleImport_IndexesEachFileOnSharedInstance
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Full build reuses one clsDbModule instance. Each import must index under
'           : its own file name, not a stale @Folder path from the prior import.
'---------------------------------------------------------------------------------------
'
Public Sub TestModuleImport_IndexesEachFileOnSharedInstance()
'@Tag("integration")

    Dim cMod As IDbComponent
    Dim strFile1 As String
    Dim strFile2 As String
    Dim strBase As String

    ' Use non-test utility modules so the import under test does not remove/reload
    ' modules that are actively hosting the test runner or other test subs.
    strBase = Options.GetExportFolder & "modules\"
    strFile1 = strBase & "Utility\modTimer.bas"
    strFile2 = strBase & "Utility\modZip.bas"
    If Not FSO.FileExists(strFile1) Then Exit Sub
    If Not FSO.FileExists(strFile2) Then Exit Sub

    Set cMod = New clsDbModule
    cMod.Import strFile1
    cMod.Import strFile2

    TestAssert VCSIndex.Exists(cMod, strFile1), "first imported module indexed"
    TestAssert VCSIndex.Exists(cMod, strFile2), "second imported module indexed under its own name"

End Sub
