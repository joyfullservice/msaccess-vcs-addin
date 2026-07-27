Attribute VB_Name = "modTestMergeDetection"
'---------------------------------------------------------------------------------------
' Module    : modTestMergeDetection
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Regression tests for multi-file merge change-detection (AllFilesHash and
'           : companion .json indexing).
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
    CleanupCreatedCompanionJson strFile, blnCreatedJson
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
        TestAssert True, "SKIP: companion json missing for " & strBaseName
        Exit Sub
    End If

    strOriginal = ReadFile(strJson)
    SeedMergeIndexBaseline cCategory, strFile

    WriteFile strOriginal & " ", strJson
    Set dModified = VCSIndex.GetModifiedSourceFiles(cCategory)
    TestAssert dModified.Exists(strFile), "metadata-only json change detected: " & strBaseName

    WriteFile strOriginal, strJson
    SeedMergeIndexBaseline cCategory, strFile
    CleanupCreatedCompanionJson strFile, blnCreatedJson
End Sub


Private Sub SeedMergeIndexBaseline(cCategory As IDbComponent, strFile As String)
    Dim cIdx As clsVCSIndexItem

    Set cIdx = VCSIndex.Item(cCategory, strFile)
    cIdx.FileHash = GetFileHash(strFile)
    cIdx.FilePropertiesHash = GetSourceFilesPropertyHash(cCategory, strFile)
    cIdx.AllFilesHash = GetSourceFilesContentHash(cCategory, strFile)
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
