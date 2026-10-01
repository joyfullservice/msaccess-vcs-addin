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
    Dim dSysTables As Dictionary

    Set dSysTables = GetSystemTableNames
    For Each tdf In CurrentData.AllTables
        If dSysTables.Exists(tdf.Name) Or tdf.Name Like "~*" Then
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
    Dim strRepoRoot As String
    Dim eelSavedLevel As eErrorLevel

    ' Use fixture modules that are not already loaded in the add-in project.
    ' Re-importing live modules (e.g. modTimer) fails to remove the in-use
    ' component and leaves duplicates (modTimer1) plus a VBE project-reset prompt.
    strRepoRoot = Git.GetRepositoryRoot
    If Len(strRepoRoot) = 0 Then Exit Sub
    strBase = strRepoRoot & "Testing\Fixtures\modules\"
    strFile1 = strBase & "vcs_test_import_alpha.bas"
    strFile2 = strBase & "vcs_test_import_beta.bas"
    If Not FSO.FileExists(strFile1) Then Exit Sub
    If Not FSO.FileExists(strFile2) Then Exit Sub

    RemoveTestImportFixtureModule "vcs_test_import_alpha"
    RemoveTestImportFixtureModule "vcs_test_import_beta"

    ' Import skips indexing at eelError or above, and no operation begins in this
    ' project during a test run to clear a level left by an earlier test.
    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError

    Set cMod = New clsDbModule
    cMod.Import strFile1
    cMod.Import strFile2

    TestAssert VCSIndex.Exists(cMod, strFile1), "first imported module indexed"
    TestAssert VCSIndex.Exists(cMod, strFile2), "second imported module indexed under its own name"

    Operation.ErrorLevel = eelSavedLevel
    RemoveTestImportFixtureModule "vcs_test_import_alpha"
    RemoveTestImportFixtureModule "vcs_test_import_beta"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestModuleImportFast_IndexesEachFileOnSharedInstance
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Full-build two-pass import must index each file under its own path on the
'           : shared clsDbModule instance, matching the per-file Import contract.
'---------------------------------------------------------------------------------------
'
Public Sub TestModuleImportFast_IndexesEachFileOnSharedInstance()
'@Tag("integration")

    Dim cMod As clsDbModule
    Dim strFile1 As String
    Dim strFile2 As String
    Dim strBase As String
    Dim strRepoRoot As String
    Dim eelSavedLevel As eErrorLevel

    strRepoRoot = Git.GetRepositoryRoot
    If Len(strRepoRoot) = 0 Then Exit Sub
    strBase = strRepoRoot & "Testing\Fixtures\modules\"
    strFile1 = strBase & "vcs_test_import_alpha.bas"
    strFile2 = strBase & "vcs_test_import_beta.bas"
    If Not FSO.FileExists(strFile1) Then Exit Sub
    If Not FSO.FileExists(strFile2) Then Exit Sub

    RemoveTestImportFixtureModule "vcs_test_import_alpha"
    RemoveTestImportFixtureModule "vcs_test_import_beta"

    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError

    Set cMod = New clsDbModule
    cMod.ImportFast strFile1
    cMod.ImportFast strFile2
    cMod.FinalizeImports

    TestAssert VCSIndex.Exists(cMod.Parent, strFile1), "first batch-imported module indexed"
    TestAssert VCSIndex.Exists(cMod.Parent, strFile2), "second batch-imported module indexed under its own name"

    Operation.ErrorLevel = eelSavedLevel
    RemoveTestImportFixtureModule "vcs_test_import_alpha"
    RemoveTestImportFixtureModule "vcs_test_import_beta"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDefGetFileList_ExcludesMetadataSidecar
' Author    : Adam Waller
' Date      : 7/14/2026
' Purpose   : Local tables export schema as .xml and metadata as a sibling .json sidecar.
'           : GetFileList must not treat the sidecar as a linked-table definition.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDefGetFileList_ExcludesMetadataSidecar()

    Dim strSavedExport As String
    Dim strRoot As String
    Dim strTblDefs As String
    Dim cTable As IDbComponent
    Dim dFiles As Dictionary
    Dim strLocalXml As String
    Dim strLocalJson As String
    Dim strLinkedJson As String

    strSavedExport = Options.ExportFolder
    strRoot = GetTempFolder("vcs_tbldef_sidecar") & PathSep
    strTblDefs = strRoot & "tbldefs" & PathSep
    VerifyPath strTblDefs

    strLocalXml = strTblDefs & "tblLocalSidecar.xml"
    strLocalJson = strTblDefs & "tblLocalSidecar.json"
    strLinkedJson = strTblDefs & "tblLinkedOnly.json"

    WriteFile "<?xml version=""1.0""?><root/>", strLocalXml
    WriteFile "{""Info"":{""Description"":""metadata""},""Items"":{""Properties"":{}}}", strLocalJson
    WriteFile "{""Info"":{""Class"":""clsDbTableDef""},""Items"":{" & _
        """Connect"":""ODBC;DSN=test"",""Name"":""tblLinkedOnly""," & _
        """SourceTableName"":""tblLinkedOnly"",""Attributes"":0}}", strLinkedJson

    Options.ExportFolder = strRoot
    Set cTable = New clsDbTableDef
    Set dFiles = cTable.GetFileList

    TestAssert dFiles.Exists(strLocalXml), "local table .xml included in file list"
    TestAssert Not dFiles.Exists(strLocalJson), "metadata sidecar .json excluded from file list"
    TestAssert dFiles.Exists(strLinkedJson), "linked table .json included in file list"

    Options.ExportFolder = strSavedExport
    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strRoot) Then FSO.DeleteFolder StripSlash(strRoot), True
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportCasingFixtures
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Import the two casing fixture modules, the holder always and the declarer
'           : if requested. The declarer declares VCSCASINGPROBE, and when it is imported
'           : after the holder the VBE re-cases the identifier in the holder. Returns
'           : false if the fixtures are not available. The identifier is unique to these
'           : fixtures so that it cannot re-case any code of the add-in itself.
'           : Import skips indexing at eelError or above, and no operation begins in this
'           : project during a test run to clear a level left by an earlier test, so the
'           : level is cleared here and restored by RemoveCasingFixtures.
'---------------------------------------------------------------------------------------
'
Private Function ImportCasingFixtures(cMod As IDbComponent, strHolderFile As String, _
    strDeclarerFile As String, blnDeclarer As Boolean, eelSavedLevel As eErrorLevel) As Boolean

    Dim strRepoRoot As String
    Dim strBase As String

    strRepoRoot = Git.GetRepositoryRoot
    If Len(strRepoRoot) = 0 Then Exit Function
    strBase = strRepoRoot & "Testing\Fixtures\modules\"
    strHolderFile = strBase & "vcs_test_casing_holder.bas"
    strDeclarerFile = strBase & "vcs_test_casing_declarer.bas"
    If Not FSO.FileExists(strHolderFile) Then Exit Function
    If Not FSO.FileExists(strDeclarerFile) Then Exit Function

    RemoveTestImportFixtureModule "vcs_test_casing_holder"
    RemoveTestImportFixtureModule "vcs_test_casing_declarer"

    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError

    Set cMod = New clsDbModule
    cMod.Import strHolderFile
    If blnDeclarer Then cMod.Import strDeclarerFile
    ImportCasingFixtures = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveCasingFixtures
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Remove the casing fixture modules and their index entries, and restore the
'           : error level that ImportCasingFixtures cleared.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveCasingFixtures(cMod As IDbComponent, strHolderFile As String, strDeclarerFile As String, _
    eelSavedLevel As eErrorLevel)

    Operation.ErrorLevel = eelSavedLevel
    RemoveTestImportFixtureModule "vcs_test_casing_holder"
    RemoveTestImportFixtureModule "vcs_test_casing_declarer"
    If Not cMod Is Nothing Then
        If VCSIndex.Exists(cMod, strHolderFile) Then VCSIndex.Remove cMod, strHolderFile
        If VCSIndex.Exists(cMod, strDeclarerFile) Then VCSIndex.Remove cMod, strDeclarerFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCodeHash_IgnoresVbeRecasing
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : Import a module that uses an identifier, then another module that declares
'           : it with a different case. The VBE re-cases the identifier in the first
'           : module, which nobody edited, and its stored hash must still match.
'           : (Calls the comparison directly, since IsModified can return before hashing.)
'---------------------------------------------------------------------------------------
'
Public Sub TestCodeHash_IgnoresVbeRecasing()
'@Tag("integration")

    Dim cMod As IDbComponent
    Dim strHolderFile As String
    Dim strDeclarerFile As String
    Dim eelSavedLevel As eErrorLevel
    Dim strStored As String
    Dim strLegacyBefore As String
    Dim strLegacyAfter As String
    Dim strCode As String
    Dim blnRecased As Boolean

    If Not ImportCasingFixtures(cMod, strHolderFile, strDeclarerFile, False, eelSavedLevel) Then Exit Sub
    TestAssert VCSIndex.Exists(cMod, strHolderFile), "precondition: the holder was indexed on import"
    If Not VCSIndex.Exists(cMod, strHolderFile) Then GoTo CleanUp

    strStored = VCSIndex.Item(cMod, strHolderFile).OtherHash
    strLegacyBefore = GetCodeModuleHash(edbModule, "vcs_test_casing_holder", True)
    TestAssert Left$(strStored, Len(cstrCodeHashPrefix)) = cstrCodeHashPrefix, "import stores a prefixed hash"

    ' Now declare the identifier with a different case in another module
    cMod.Import strDeclarerFile

    ' Precondition: the VBE really re-cased the identifier in the holder
    strCode = CurrentVBProject.VBComponents("vcs_test_casing_holder").CodeModule.Lines(1, 999999)
    blnRecased = (InStr(1, strCode, "VCSCASINGPROBE", vbBinaryCompare) > 0)
    TestAssert blnRecased, "precondition: the holder now contains VCSCASINGPROBE"
    If Not blnRecased Then GoTo CleanUp

    ' Control: the old hash no longer matches, so the old comparison would have failed
    strLegacyAfter = GetCodeModuleHash(edbModule, "vcs_test_casing_holder", True)
    TestAssert StrComp(strLegacyBefore, strLegacyAfter, vbBinaryCompare) <> 0, _
        "control: the case-sensitive hash changed after the re-casing"

    TestAssert CodeModuleHashMatches(strStored, edbModule, "vcs_test_casing_holder"), _
        "stored hash still matches after the VBE re-cased an identifier"

CleanUp:
    RemoveCasingFixtures cMod, strHolderFile, strDeclarerFile, eelSavedLevel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCodeHash_DetectsStringCaseChange
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : The negative control: a change of case inside a string is a real change and
'           : must not match, and neither must an added line of code.
'---------------------------------------------------------------------------------------
'
Public Sub TestCodeHash_DetectsStringCaseChange()
'@Tag("integration")

    Dim cMod As IDbComponent
    Dim strHolderFile As String
    Dim strDeclarerFile As String
    Dim eelSavedLevel As eErrorLevel
    Dim strStored As String
    Dim cmpHolder As VBComponent
    Dim lngLine As Long
    Dim strLine As String
    Dim lngString As Long

    If Not ImportCasingFixtures(cMod, strHolderFile, strDeclarerFile, False, eelSavedLevel) Then Exit Sub
    TestAssert VCSIndex.Exists(cMod, strHolderFile), "precondition: the holder was indexed on import"
    If Not VCSIndex.Exists(cMod, strHolderFile) Then GoTo CleanUp

    strStored = VCSIndex.Item(cMod, strHolderFile).OtherHash
    Set cmpHolder = CurrentVBProject.VBComponents("vcs_test_casing_holder")

    TestAssert CodeModuleHashMatches(strStored, edbModule, "vcs_test_casing_holder"), _
        "control: unchanged module matches"

    ' Find the line with the string
    With cmpHolder.CodeModule
        For lngLine = 1 To .CountOfLines
            If InStr(1, .Lines(lngLine, 1), """Probe", vbBinaryCompare) > 0 Then
                lngString = lngLine
                Exit For
            End If
        Next lngLine
    End With
    TestAssert lngString > 0, "precondition: found the line with the string"
    If lngString = 0 Then GoTo CleanUp

    ' Change the case inside the string
    strLine = cmpHolder.CodeModule.Lines(lngString, 1)
    cmpHolder.CodeModule.ReplaceLine lngString, Replace(strLine, "Probe", "PROBE", , , vbBinaryCompare)
    TestAssert Not CodeModuleHashMatches(strStored, edbModule, "vcs_test_casing_holder"), _
        "case change inside a string is detected"

    ' Restore it, and add a line of code instead
    cmpHolder.CodeModule.ReplaceLine lngString, strLine
    TestAssert CodeModuleHashMatches(strStored, edbModule, "vcs_test_casing_holder"), _
        "control: restored module matches again"
    cmpHolder.CodeModule.InsertLines cmpHolder.CodeModule.CountOfLines + 1, _
        "Public Sub vcsTestCasingExtra()" & vbCrLf & "End Sub"
    TestAssert Not CodeModuleHashMatches(strStored, edbModule, "vcs_test_casing_holder"), _
        "an added line of code is detected"

CleanUp:
    RemoveCasingFixtures cMod, strHolderFile, strDeclarerFile, eelSavedLevel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUpgradeLegacyCodeHashes
' Author    : Adam Waller
' Date      : 9/29/2026
' Purpose   : An index entry with an older hash that still matches its module is
'           : upgraded without touching the dates. One that does not match is left as
'           : is, and no entry is created for an object that has none.
'---------------------------------------------------------------------------------------
'
Public Sub TestUpgradeLegacyCodeHashes()
'@Tag("integration")

    Const cstrFakeLegacy As String = "0123456789abcdef0123456789abcdef"

    Dim cMod As IDbComponent
    Dim strHolderFile As String
    Dim strDeclarerFile As String
    Dim eelSavedLevel As eErrorLevel
    Dim dteExport As Date
    Dim dteImport As Date
    Dim strLegacy As String

    If Not ImportCasingFixtures(cMod, strHolderFile, strDeclarerFile, True, eelSavedLevel) Then Exit Sub
    TestAssert VCSIndex.Exists(cMod, strHolderFile), "precondition: the holder was indexed on import"
    If Not VCSIndex.Exists(cMod, strHolderFile) Then GoTo CleanUp
    TestAssert VCSIndex.Exists(cMod, strDeclarerFile), "precondition: the declarer was indexed on import"
    If Not VCSIndex.Exists(cMod, strDeclarerFile) Then GoTo CleanUp

    ' A legacy entry that still matches its module is upgraded
    With VCSIndex.Item(cMod, strHolderFile)
        dteExport = .ExportDate
        dteImport = .ImportDate
        strLegacy = GetCodeModuleHash(edbModule, "vcs_test_casing_holder", True)
        .OtherHash = strLegacy
    End With
    VCSIndex.UpgradeLegacyCodeHashes
    With VCSIndex.Item(cMod, strHolderFile)
        TestAssert Left$(.OtherHash, Len(cstrCodeHashPrefix)) = cstrCodeHashPrefix, "matching legacy hash was upgraded"
        TestAssert CodeModuleHashMatches(.OtherHash, edbModule, "vcs_test_casing_holder"), "upgraded hash matches the module"
        TestAssert .ExportDate = dteExport, "export date was not touched"
        TestAssert .ImportDate = dteImport, "import date was not touched"
    End With

    ' A legacy entry that does not match is left alone
    VCSIndex.Item(cMod, strHolderFile).OtherHash = cstrFakeLegacy
    VCSIndex.Item(cMod, strDeclarerFile).OtherHash = cstrFakeLegacy
    VCSIndex.UpgradeLegacyCodeHashes
    TestAssert StrComp(VCSIndex.Item(cMod, strHolderFile).OtherHash, cstrFakeLegacy, vbBinaryCompare) = 0, _
        "non-matching legacy hash was left as it was"

    ' No entry is created for an object that has none
    VCSIndex.Remove cMod, strHolderFile
    VCSIndex.Item(cMod, strDeclarerFile).OtherHash = cstrFakeLegacy
    VCSIndex.UpgradeLegacyCodeHashes
    TestAssert Not VCSIndex.Exists(cMod, strHolderFile), "no index entry was created for an object without one"

CleanUp:
    RemoveCasingFixtures cMod, strHolderFile, strDeclarerFile, eelSavedLevel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveTestImportFixtureModule
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Remove a sandbox module imported by TestModuleImport_* if present.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveTestImportFixtureModule(strName As String)

    LogUnhandledErrors
    On Error Resume Next
    CurrentVBProject.VBComponents.Remove CurrentVBProject.VBComponents(strName)
    Err.Clear

End Sub
