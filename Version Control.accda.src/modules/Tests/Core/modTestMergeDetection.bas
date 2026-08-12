Attribute VB_Name = "modTestMergeDetection"
'---------------------------------------------------------------------------------------
' Module    : modTestMergeDetection
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Regression tests for multi-file merge change-detection (AllFilesHash,
'           : companion .json / .cls indexing, and the legacy-entry upgrade path).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestMergeDetectsFormMetadataOnlyJsonChange()
    RunMetadataOnlyMergeTest GetTestFormComponent, "frmVCSMain"
End Sub


Public Sub TestMergeDetectsModuleMetadataOnlyJsonChange()
    RunMetadataOnlyMergeTest GetTestModuleComponent, "modTestIndex"
End Sub


Public Sub TestMergeDetectsMacroMetadataOnlyJsonChange()
    RunMetadataOnlyMergeTest GetTestMacroComponent, "autoexec"
End Sub


Public Sub TestMergeIgnoresTimestampOnlyWhenAllFilesHashMatches()
    Dim cForm As IDbComponent
    Dim strFile As String
    Dim dModified As Dictionary
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    blnCreatedJson = EnsureCompanionJson(strFile)
    SeedMergeIndexBaseline cForm, strFile
    VCSIndex.Item(cForm, strFile).FilePropertiesHash = "stale_property_hash"

    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    TestAssert Not dModified.Exists(strFile), "timestamp-only drift ignored when AllFilesHash matches"
    RestoreMergeIndexBaseline cForm, strFile, blnCreatedJson
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeSkipsContentHashWhenPropertyHashMatches
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Pin the fast path in GetModifiedSourceFiles. When the property hash (date
'           : and size of every indexed file) matches the index, no file content is read
'           : at all. Poisoning AllFilesHash proves that: if the content hash were still
'           : being consulted, the bogus value would report the file as modified.
'           :
'           : This is what keeps a no-change merge from reading every source file. See
'           : the 2026-07-29 DECISIONS.md entry.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeSkipsContentHashWhenPropertyHashMatches()

    Dim cForm As IDbComponent
    Dim cIdx As clsVCSIndexItem
    Dim strFile As String
    Dim dModified As Dictionary
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    blnCreatedJson = EnsureCompanionJson(strFile)
    SeedMergeIndexBaseline cForm, strFile

    ' Poison the content hash while leaving the property hash valid
    Set cIdx = VCSIndex.Item(cForm, strFile)
    cIdx.AllFilesHash = "bogus_content_hash"

    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    TestAssert Not dModified.Exists(strFile), _
        "matching property hash short-circuits before the content hash is consulted"

    ' Restore a valid baseline so later tests are not affected
    RestoreMergeIndexBaseline cForm, strFile, blnCreatedJson

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeSkipsTableDataAfterExport
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Export must seed the index for table data so a subsequent merge scan skips
'           : unchanged source files. Before the fix, only Merge updated the index.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeSkipsTableDataAfterExport()
    '@Tag("integration")

    Dim cCategory As IDbComponent
    Dim cTable As clsDbTableData
    Dim dbs As DAO.Database
    Dim strTable As String
    Dim strFile As String
    Dim dModified As Dictionary

    strTable = "vcs_test_export_idx"
    strFile = vbNullString
    Set dbs = CurrentDb
    Set cCategory = New clsDbTableData

    LogUnhandledErrors
    On Error Resume Next
    dbs.Execute "DROP TABLE [" & strTable & "]"
    On Error GoTo 0
    dbs.Execute "CREATE TABLE [" & strTable & "] (ID LONG, Name TEXT(10))"
    dbs.Execute "INSERT INTO [" & strTable & "] (ID, Name) VALUES (1, 'a')"

    Set cTable = New clsDbTableData
    cTable.Format = etdTabDelimited
    On Error Resume Next
    Set cTable.Parent.DbObject = CurrentData.AllTables(strTable)
    On Error GoTo 0
    If cTable.Parent.DbObject Is Nothing Then
        TestAssert True, "SKIP: could not bind test table"
        GoTo CleanUp
    End If

    cTable.Parent.Export
    strFile = Options.GetExportFolder & "tables\" & strTable & ".txt"
    If Not FSO.FileExists(strFile) Then
        TestAssert True, "SKIP: table data export file missing"
        GoTo CleanUp
    End If

    Set dModified = VCSIndex.GetModifiedSourceFiles(cCategory)
    TestAssert Not dModified.Exists(strFile), _
        "freshly exported table data skipped on merge scan"

CleanUp:
    On Error Resume Next
    dbs.Execute "DROP TABLE [" & strTable & "]"
    If Len(strFile) > 0 Then
        If FSO.FileExists(strFile) Then DeleteFile strFile
        VCSIndex.Remove cCategory, strFile
    End If
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeUsesSharedScanMetadata
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : A merge build passes one folder metadata map covering every category rather
'           : than letting each category scan its own folder. Verify the shared map
'           : produces the same verdict as the per-category fallback.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeUsesSharedScanMetadata()

    Dim cForm As IDbComponent
    Dim strFile As String
    Dim dShared As Dictionary
    Dim dOwnScan As Dictionary
    Dim dRootMeta As Dictionary
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    blnCreatedJson = EnsureCompanionJson(strFile)
    SeedMergeIndexBaseline cForm, strFile

    ' A recursive scan of the export root covers every category folder
    Set dRootMeta = ScanFolderMetadata(Options.GetExportFolder)
    Set dShared = VCSIndex.GetModifiedSourceFiles(cForm, dRootMeta)
    Set dOwnScan = VCSIndex.GetModifiedSourceFiles(cForm)

    TestAssert dShared.Exists(strFile) = dOwnScan.Exists(strFile), _
        "shared folder scan agrees with the per-category scan"
    TestAssert dShared.Count = dOwnScan.Count, "same number of modified files reported"

    RestoreMergeIndexBaseline cForm, strFile, blnCreatedJson

End Sub


Public Sub TestGetSourceFilesContentHashIncludesJson()
    Dim cForm As IDbComponent
    Dim strFile As String
    Dim strHashBefore As String
    Dim strHashAfter As String
    Dim strJson As String
    Dim strOriginal As String
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    strJson = SwapExtension(strFile, "json")
    blnCreatedJson = EnsureCompanionJson(strFile)
    strOriginal = ReadFile(strJson)
    strHashBefore = GetSourceFilesContentHash(cForm, strFile)

    WriteFile strOriginal & vbCrLf, strJson
    strHashAfter = GetSourceFilesContentHash(cForm, strFile)
    TestAssert strHashBefore <> strHashAfter, "content hash reflects json change"

    WriteFile strOriginal, strJson
    CleanupCreatedCompanionJson strFile, blnCreatedJson
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeDetectsClsOnlyChangeWithLegacyIndexEntry
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : Issue #748 regression. An index entry written before AllFilesHash existed
'           : must still report a form as modified when only its companion .cls changes.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeDetectsClsOnlyChangeWithLegacyIndexEntry()

    Dim cForm As IDbComponent
    Dim strFile As String
    Dim strCls As String
    Dim strOriginal As String
    Dim dModified As Dictionary

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    strCls = SwapExtension(strFile, "cls")
    If Not FSO.FileExists(strCls) Then
        TestAssert True, "SKIP: frmVCSMain.cls fixture missing"
        Exit Sub
    End If

    strOriginal = ReadFile(strCls)
    SeedLegacyMergeIndexBaseline cForm, strFile

    WriteFile strOriginal & " ", strCls
    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    TestAssert dModified.Exists(strFile), _
        "legacy index entry detects .cls-only change (issue #748)"

    WriteFile strOriginal, strCls
    SeedMergeIndexBaseline cForm, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeDoesNotMarkLegacyEntrySyncedAfterMissedEdit
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : Pin the silent-data-loss half of issue #748. The legacy multi-file branch
'           : must not refresh FilePropertiesHash to the edited tree's dates/sizes, which
'           : would record a companion-only edit as synced and hide it from later merges.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeDoesNotMarkLegacyEntrySyncedAfterMissedEdit()

    Dim cForm As IDbComponent
    Dim cIdx As clsVCSIndexItem
    Dim strFile As String
    Dim strCls As String
    Dim strOriginal As String
    Dim strSeededPropHash As String
    Dim dModified As Dictionary

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    strCls = SwapExtension(strFile, "cls")
    If Not FSO.FileExists(strCls) Then
        TestAssert True, "SKIP: frmVCSMain.cls fixture missing"
        Exit Sub
    End If

    strOriginal = ReadFile(strCls)
    SeedLegacyMergeIndexBaseline cForm, strFile
    strSeededPropHash = VCSIndex.Item(cForm, strFile).FilePropertiesHash

    WriteFile strOriginal & " ", strCls
    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    Set cIdx = VCSIndex.Item(cForm, strFile)

    TestAssert dModified.Exists(strFile), "edited form reported modified"
    TestAssert cIdx.FilePropertiesHash = strSeededPropHash, _
        "legacy multi-file branch does not refresh FilePropertiesHash to the edited state"

    WriteFile strOriginal, strCls
    SeedMergeIndexBaseline cForm, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeBackfillsAllFilesHashForLegacyEntry
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : When a legacy multi-file entry is still in its synced state (property hash
'           : matches), the fast path records AllFilesHash so later scans can arbitrate
'           : companion edits precisely without a full re-export.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeBackfillsAllFilesHashForLegacyEntry()

    Dim cForm As IDbComponent
    Dim cIdx As clsVCSIndexItem
    Dim strFile As String
    Dim dModified As Dictionary
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    ' Ensure the form is multi-file so the backfill path runs
    blnCreatedJson = EnsureCompanionJson(strFile)
    SeedLegacyMergeIndexBaseline cForm, strFile

    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    Set cIdx = VCSIndex.Item(cForm, strFile)

    TestAssert Not dModified.Exists(strFile), "clean legacy entry not reported modified"
    TestAssert Len(cIdx.AllFilesHash) > 0, "AllFilesHash backfilled on clean fast path"
    TestAssert cIdx.AllFilesHash = GetSourceFilesContentHash(cForm, strFile), _
        "backfilled AllFilesHash matches current content"

    RestoreMergeIndexBaseline cForm, strFile, blnCreatedJson

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestMergeLegacySingleFileUsesPrimaryHash
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : A legacy entry for a single-file component (no companion on disk) must
'           : still resolve via the primary content hash and not be reported modified
'           : on timestamp-only drift. The multi-file conservative branch must not
'           : fire here -- FileExtensions.Count is not the criterion; existence is.
'---------------------------------------------------------------------------------------
'
Public Sub TestMergeLegacySingleFileUsesPrimaryHash()

    Dim cModule As IDbComponent
    Dim strFile As String
    Dim strJson As String
    Dim dModified As Dictionary

    Set cModule = GetTestModuleComponent
    If cModule Is Nothing Then
        TestAssert True, "SKIP: modTestIndex not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cModule, "modTestIndex")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: modTestIndex.bas fixture missing"
        Exit Sub
    End If

    ' Guard: a companion .json would make this multi-file
    strJson = SwapExtension(strFile, "json")
    If FSO.FileExists(strJson) Then
        TestAssert True, "SKIP: modTestIndex has a companion .json"
        Exit Sub
    End If

    SeedLegacyMergeIndexBaseline cModule, strFile
    ' Force the property-hash mismatch path while content is unchanged
    VCSIndex.Item(cModule, strFile).FilePropertiesHash = "stale_property_hash"

    Set dModified = VCSIndex.GetModifiedSourceFiles(cModule)
    TestAssert Not dModified.Exists(strFile), _
        "legacy single-file entry dismisses timestamp-only drift via primary hash"

    SeedMergeIndexBaseline cModule, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAltExportPromotionRecordsSourceFileState
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : When an export resolves a conflict by moving the temp copy into the export
'           : folder, UpdateFromAltExport must measure the source files where they now
'           : live. Update leaves those values blank for an alternate export, so copying
'           : them across left a multi-file component with no recorded content hash --
'           : reported modified by every later merge, which for an add-in form fails with
'           : "Merging not supported for add-in forms" even though nothing changed.
'---------------------------------------------------------------------------------------
'
Public Sub TestAltExportPromotionRecordsSourceFileState()

    Dim cForm As IDbComponent
    Dim cIdx As clsVCSIndexItem
    Dim strFile As String
    Dim dModified As Dictionary
    Dim blnCreatedJson As Boolean

    Set cForm = GetTestFormComponent
    If cForm Is Nothing Then
        TestAssert True, "SKIP: frmVCSMain not available"
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cForm, "frmVCSMain")
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: frmVCSMain.form fixture missing"
        Exit Sub
    End If

    ' A companion .json makes this multi-file, where a missing content hash cannot
    ' be ruled out as a companion-only edit.
    blnCreatedJson = EnsureCompanionJson(strFile)

    ' Record the entry a conflict temp export leaves behind, then promote it the way
    ' the export loop does once the files have been moved into the export folder.
    VCSIndex.Update cForm, eatAltExport, GetFileHash(strFile)
    VCSIndex.UpdateFromAltExport cForm

    Set cIdx = VCSIndex.Item(cForm, strFile)
    TestAssert cIdx.FilePropertiesHash = GetSourceFilesPropertyHash(cForm, strFile), _
        "promotion records the property hash of the promoted files"
    TestAssert cIdx.AllFilesHash = GetSourceFilesContentHash(cForm, strFile), _
        "promotion records the combined content hash of the promoted files"

    Set dModified = VCSIndex.GetModifiedSourceFiles(cForm)
    TestAssert Not dModified.Exists(strFile), _
        "a promoted component is not reported modified on the next merge scan"

    ' The AlternateExport entry seeded above is dropped when the index is saved.
    RestoreMergeIndexBaseline cForm, strFile, blnCreatedJson

End Sub


Private Sub RunMetadataOnlyMergeTest(cCategory As IDbComponent, strBaseName As String)
    Dim strFile As String
    Dim strJson As String
    Dim strOriginal As String
    Dim dModified As Dictionary
    Dim blnCreatedJson As Boolean

    If cCategory Is Nothing Then
        TestAssert True, "SKIP: component not available for " & strBaseName
        Exit Sub
    End If

    strFile = ResolveSourceFilePath(cCategory, strBaseName)
    If Len(strFile) = 0 Then
        TestAssert True, "SKIP: fixture missing: " & strBaseName
        Exit Sub
    End If

    strJson = SwapExtension(strFile, "json")
    blnCreatedJson = EnsureCompanionJson(strFile)
    If Not FSO.FileExists(strJson) Then
        CleanupCreatedCompanionJson strFile, blnCreatedJson
        TestAssert True, "SKIP: companion json missing for " & strBaseName
        Exit Sub
    End If

    strOriginal = ReadFile(strJson)
    SeedMergeIndexBaseline cCategory, strFile

    WriteFile strOriginal & " ", strJson
    Set dModified = VCSIndex.GetModifiedSourceFiles(cCategory)
    TestAssert dModified.Exists(strFile), "metadata-only json change detected: " & strBaseName

    WriteFile strOriginal, strJson
    RestoreMergeIndexBaseline cCategory, strFile, blnCreatedJson
End Sub


Private Sub SeedMergeIndexBaseline(cCategory As IDbComponent, strFile As String)
    Dim cIdx As clsVCSIndexItem

    Set cIdx = VCSIndex.Item(cCategory, strFile)
    cIdx.FileHash = GetFileHash(strFile)
    cIdx.FilePropertiesHash = GetSourceFilesPropertyHash(cCategory, strFile)
    cIdx.AllFilesHash = GetSourceFilesContentHash(cCategory, strFile)
End Sub


' Leave the index describing the files that actually remain on disk. Seeding a
' baseline before removing a companion created for the test records hashes for a
' file that no longer exists, which reports the component modified on every merge
' from then on -- in this project, against the developer's own live index.
Private Sub RestoreMergeIndexBaseline(cCategory As IDbComponent, strFile As String, _
    blnCreatedJson As Boolean)

    CleanupCreatedCompanionJson strFile, blnCreatedJson
    SeedMergeIndexBaseline cCategory, strFile

End Sub


' Same as SeedMergeIndexBaseline but leaves AllFilesHash empty, simulating an index
' entry written by a pre-AllFilesHash build (5.0.1 and earlier).
Private Sub SeedLegacyMergeIndexBaseline(cCategory As IDbComponent, strFile As String)
    Dim cIdx As clsVCSIndexItem

    Set cIdx = VCSIndex.Item(cCategory, strFile)
    cIdx.FileHash = GetFileHash(strFile)
    cIdx.FilePropertiesHash = GetSourceFilesPropertyHash(cCategory, strFile)
    cIdx.AllFilesHash = vbNullString
End Sub


' Returns True when this call created the companion file (caller must delete it).
Private Function EnsureCompanionJson(strFile As String) As Boolean
    Dim strJson As String

    strJson = SwapExtension(strFile, "json")
    If Not FSO.FileExists(strJson) Then
        WriteFile "{}", strJson
        EnsureCompanionJson = True
    End If
End Function


Private Sub CleanupCreatedCompanionJson(strFile As String, blnCreated As Boolean)
    Dim strJson As String

    If Not blnCreated Then Exit Sub
    strJson = SwapExtension(strFile, "json")
    If FSO.FileExists(strJson) Then DeleteFile strJson
End Sub


Private Function ResolveSourceFilePath(cCategory As IDbComponent, strBaseName As String) As String
    Dim dFiles As Dictionary
    Dim varFile As Variant

    Set dFiles = cCategory.GetFileList
    For Each varFile In dFiles
        If StrComp(FSO.GetBaseName(CStr(varFile)), strBaseName, vbTextCompare) = 0 Then
            ResolveSourceFilePath = CStr(varFile)
            Exit Function
        End If
    Next varFile
End Function


Private Function GetTestFormComponent() As IDbComponent
    Dim cForm As IDbComponent

    LogUnhandledErrors
    On Error Resume Next
    Set cForm = New clsDbForm
    Set cForm.DbObject = CurrentProject.AllForms("frmVCSMain")
    On Error GoTo 0
    If cForm.DbObject Is Nothing Then Exit Function
    Set GetTestFormComponent = cForm
End Function


Private Function GetTestModuleComponent() As IDbComponent
    Dim cModule As IDbComponent

    LogUnhandledErrors
    On Error Resume Next
    Set cModule = New clsDbModule
    Set cModule.DbObject = CurrentProject.AllModules("modTestIndex")
    On Error GoTo 0
    If cModule.DbObject Is Nothing Then Exit Function
    Set GetTestModuleComponent = cModule
End Function


Private Function GetTestMacroComponent() As IDbComponent
    Dim cMacro As IDbComponent

    LogUnhandledErrors
    On Error Resume Next
    Set cMacro = New clsDbMacro
    Set cMacro.DbObject = CurrentProject.AllMacros("autoexec")
    On Error GoTo 0
    If cMacro.DbObject Is Nothing Then Exit Function
    Set GetTestMacroComponent = cMacro
End Function
