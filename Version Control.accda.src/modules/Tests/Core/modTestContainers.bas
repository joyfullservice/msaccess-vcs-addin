Attribute VB_Name = "modTestContainers"
'---------------------------------------------------------------------------------------
' Module    : modTestContainers
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for source file metadata functions in modContainers:
'           : GetSourceModifiedDate, GetSourceFilesPropertyHash,
'           : GetSourceFilesContentHash, GetSourceBasePath,
'           : GetLastModifiedSourceFile.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("integration")


'---------------------------------------------------------------------------------------
' Procedure : TestGetSourceModifiedDate
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Verify GetSourceModifiedDate returns a real date for a component whose
'           : source file exists on disk.
'---------------------------------------------------------------------------------------
'
Public Sub TestGetSourceModifiedDate()

    Dim cModule As IDbComponent
    Dim dteResult As Date

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub

    dteResult = GetSourceModifiedDate(cModule)
    TestAssert dteResult > 0, "returns non-zero date"

    ' Compare against FSO directly
    Dim dteFSO As Date
    dteFSO = FSO.GetFile(cModule.SourceFile).DateLastModified
    TestAssert Abs(dteResult - dteFSO) < 1 / 86400, "matches FSO DateLastModified within 1 second"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGetSourceFilesPropertyHash
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Verify GetSourceFilesPropertyHash returns a non-empty hash and is
'           : deterministic (same input = same output).
'---------------------------------------------------------------------------------------
'
Public Sub TestGetSourceFilesPropertyHash()

    Dim cModule As IDbComponent
    Dim strHash1 As String
    Dim strHash2 As String

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub

    strHash1 = GetSourceFilesPropertyHash(cModule)
    TestAssert Len(strHash1) > 0, "returns non-empty hash"

    strHash2 = GetSourceFilesPropertyHash(cModule)
    TestAssert strHash1 = strHash2, "deterministic (same result on second call)"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGetLastModifiedSourceFile
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Verify GetLastModifiedSourceFile returns a full path (not a bare filename).
'---------------------------------------------------------------------------------------
'
Public Sub TestGetLastModifiedSourceFile()

    Dim cModule As IDbComponent
    Dim strResult As String

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub

    strResult = GetLastModifiedSourceFile(cModule)
    TestAssert Len(strResult) > 0, "returns non-empty path"
    TestAssert InStr(strResult, "\") > 0 Or InStr(strResult, "/") > 0, _
        "returns full path with folder separator"
    TestAssert FSO.FileExists(strResult), "returned path is a real file"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSourceDateAndHashConsistency
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : When GetSourceFilesPropertyHash finds a file (non-empty hash),
'           : GetSourceModifiedDate should also find it (non-zero date).
'---------------------------------------------------------------------------------------
'
Public Sub TestSourceDateAndHashConsistency()

    Dim cModule As IDbComponent
    Dim strHash As String
    Dim dteDate As Date

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub

    strHash = GetSourceFilesPropertyHash(cModule)
    dteDate = GetSourceModifiedDate(cModule)

    ' Both should agree on whether the file exists
    If Len(strHash) > 0 Then
        TestAssert dteDate > 0, "hash found file, date should be non-zero"
    End If
    If dteDate > 0 Then
        TestAssert Len(strHash) > 0, "date found file, hash should be non-empty"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPropertyHashIdenticalWithAndWithoutMetaScan
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : GetSourceFilesPropertyHash reads file dates and sizes either from a
'           : per-file FSO.GetFile or from a batch Win32 folder scan. Merge change
'           : detection uses the scan map while export records the hash through the FSO
'           : path, so the two must agree byte for byte.
'           :
'           : If they ever diverge (a date conversion or precision difference, for
'           : instance) nothing breaks visibly -- the merge short-circuit simply stops
'           : firing and every source file gets read again. This test turns that silent
'           : performance regression into a visible failure.
'---------------------------------------------------------------------------------------
'
Public Sub TestPropertyHashIdenticalWithAndWithoutMetaScan()

    Dim cModule As IDbComponent
    Dim dMeta As Dictionary
    Dim strSource As String
    Dim strFsoHash As String
    Dim strScanHash As String

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub
    strSource = cModule.SourceFile

    Set dMeta = ScanFolderMetadata(cModule.BaseFolder)

    strFsoHash = GetSourceFilesPropertyHash(cModule, strSource)
    strScanHash = GetSourceFilesPropertyHash(cModule, strSource, dMeta)

    TestAssert Len(strFsoHash) > 0, "FSO path returns a hash"
    TestAssert strFsoHash = strScanHash, "folder scan produces the same property hash as FSO"

    ' The same must hold for a recursive scan rooted above the category folder, which is
    ' what a merge build shares across every category.
    Set dMeta = ScanFolderMetadata(Options.GetExportFolder)
    strScanHash = GetSourceFilesPropertyHash(cModule, strSource, dMeta)
    TestAssert strFsoHash = strScanHash, "shared root scan produces the same property hash"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestContentHashIdenticalWithAndWithoutMetaScan
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : The folder scan map only replaces the per-extension existence check in
'           : GetSourceFilesContentHash, so the resulting hash must be unchanged.
'---------------------------------------------------------------------------------------
'
Public Sub TestContentHashIdenticalWithAndWithoutMetaScan()

    Dim cModule As IDbComponent
    Dim dMeta As Dictionary
    Dim strSource As String
    Dim strFsoHash As String
    Dim strScanHash As String

    Set cModule = GetTestComponentWithSourceFile
    If cModule Is Nothing Then Exit Sub
    strSource = cModule.SourceFile

    Set dMeta = ScanFolderMetadata(cModule.BaseFolder)

    strFsoHash = GetSourceFilesContentHash(cModule, strSource)
    strScanHash = GetSourceFilesContentHash(cModule, strSource, dMeta)

    TestAssert Len(strFsoHash) > 0, "FSO path returns a hash"
    TestAssert strFsoHash = strScanHash, "folder scan produces the same content hash as FSO"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGetSourceBasePathMatchesFso
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : GetSourceBasePath replaces three FSO calls in the per-file scan loop, so it
'           : must return exactly what the FSO expression it replaced returned. Checked
'           : against every real source file in a category plus the awkward shapes:
'           : a dot in a folder name, multiple dots in a file name, and no extension.
'---------------------------------------------------------------------------------------
'
Public Sub TestGetSourceBasePathMatchesFso()

    Dim cModule As IDbComponent
    Dim varFile As Variant
    Dim strPath As String
    Dim lngChecked As Long

    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub

    For Each varFile In cModule.GetFileList
        strPath = CStr(varFile)
        TestAssert GetSourceBasePath(cModule, strPath) = FsoBasePath(strPath), _
            "matches FSO for " & FSO.GetFileName(strPath)
        lngChecked = lngChecked + 1
        If lngChecked >= 25 Then Exit For
    Next varFile

    ' Awkward shapes, independent of what happens to be on disk
    CheckBasePath cModule, "C:\folder\name.sql"
    CheckBasePath cModule, "C:\fold.er\name.sql"
    CheckBasePath cModule, "C:\fold.er\na.me.sql"
    CheckBasePath cModule, "C:\folder\name"
    CheckBasePath cModule, "C:\fold.er\name"

End Sub


Private Sub CheckBasePath(cmp As IDbComponent, ByVal strPath As String)
    TestAssert GetSourceBasePath(cmp, strPath) = FsoBasePath(strPath), _
        "matches FSO for " & strPath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FsoBasePath
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : The original FSO expression GetSourceBasePath replaced.
'---------------------------------------------------------------------------------------
'
Private Function FsoBasePath(ByVal strPath As String) As String
    FsoBasePath = FSO.BuildPath(FSO.GetParentFolderName(strPath), FSO.GetBaseName(strPath))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTestComponent
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Helper to get a real IDbComponent for testing. Returns Nothing if no
'           : modules are available (test should exit gracefully).
'---------------------------------------------------------------------------------------
'
Private Function GetTestComponent() As IDbComponent
    If CurrentProject.AllModules.Count = 0 Then Exit Function
    Dim cModule As IDbComponent
    Set cModule = New clsDbModule
    Set cModule.DbObject = CurrentProject.AllModules(0)
    Set GetTestComponent = cModule
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTestComponentWithSourceFile
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : A component whose source file is on disk, or Nothing once the reason is
'           : recorded as a failure. These tests read real exported files -- dates,
'           : sizes, contents -- and the suite runs with the development copy as the
'           : current database, which has its source tree beside it. A missing file
'           : there is a real problem rather than a configuration to skip past.
'           :
'           : Returning Nothing is what stops the caller before FSO.GetFile raises.
'           : These procedures carry no error handler, so asserting existence and then
'           : reading anyway reported a bare "File not found" from whichever line got
'           : there first, which says far less than the assertion does.
'---------------------------------------------------------------------------------------
'
Private Function GetTestComponentWithSourceFile() As IDbComponent

    Dim cModule As IDbComponent

    Set cModule = GetTestComponent
    If cModule Is Nothing Then
        TestAssert False, "a module exists to test with"
        Exit Function
    End If

    If Not FSO.FileExists(cModule.SourceFile) Then
        TestAssert False, "source file exists on disk: " & cModule.SourceFile
        Exit Function
    End If

    Set GetTestComponentWithSourceFile = cModule

End Function


'---------------------------------------------------------------------------------------
' Procedure : TestQuerySourceFileMemoization
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : clsDbQuery.SourceFile is memoized because a change scan reads it repeatedly
'           : per component. Verify the cached value matches the first computed value, and
'           : that rebinding DbObject invalidates it -- a stale path would silently point
'           : a component at another query's source file.
'---------------------------------------------------------------------------------------
'
Public Sub TestQuerySourceFileMemoization()

    Dim cQuery As IDbComponent
    Dim strFirst As String
    Dim strCached As String
    Dim strSecond As String
    Dim lngIdx As Long

    If CurrentData.AllQueries.Count = 0 Then
        TestAssert True, "SKIP: no queries in this database"
        Exit Sub
    End If

    Set cQuery = New clsDbQuery
    Set cQuery.DbObject = CurrentData.AllQueries(0)

    strFirst = cQuery.SourceFile
    TestAssert Len(strFirst) > 0, "returns a source file path"

    ' Repeated reads must return the same value from the cache
    For lngIdx = 1 To 10
        strCached = cQuery.SourceFile
        TestAssert strCached = strFirst, "stable across repeated reads"
    Next lngIdx

    ' The cached path has to reflect the bound object, not the first one ever seen
    If CurrentData.AllQueries.Count > 1 Then
        Set cQuery.DbObject = CurrentData.AllQueries(1)
        strSecond = cQuery.SourceFile
        TestAssert strSecond <> strFirst, "rebinding DbObject invalidates the cached path"

        ' And rebinding back returns the original
        Set cQuery.DbObject = CurrentData.AllQueries(0)
        TestAssert cQuery.SourceFile = strFirst, "rebinding to the original restores its path"
    End If

    ' A freshly bound instance must agree with the memoized one
    Dim cFresh As IDbComponent
    Set cFresh = New clsDbQuery
    Set cFresh.DbObject = CurrentData.AllQueries(0)
    TestAssert cFresh.SourceFile = strFirst, "memoized path matches a freshly computed one"

End Sub


Public Sub TestResolveComponentTypeMenusAlias()
    TestAssert ResolveComponentType("menus") = edbCommandBar, "menus alias"
    TestAssert ResolveComponentType("menu") = edbCommandBar, "menu alias"
End Sub


Public Sub TestResolveComponentTypeArgNull()
    TestAssert ResolveComponentTypeArg(Null) = -1, "Null rejected"
    TestAssert ResolveComponentTypeArg(Empty) = -1, "Empty rejected"
End Sub


Public Sub TestResolveComponentTypeArgEnum()
    TestAssert ResolveComponentTypeArg(edbQuery) = edbQuery, "enum passthrough"
End Sub


Public Sub TestGetContainersForTypesDedupes()
    Dim col As Collection
    Dim strError As String

    Set col = GetContainersForTypes(Array("forms", "forms"), strError)
    TestAssert Len(strError) = 0, "no error"
    TestAssert col.Count = 1, "duplicate collapsed"
End Sub


Public Sub TestGetContainersForTypesUnknown()
    Dim col As Collection
    Dim strError As String

    Set col = GetContainersForTypes("not_a_real_type", strError)
    TestAssert Len(strError) > 0, "unknown type rejected"
End Sub


Public Sub TestGetContainersForTypesImportRejectsTableData()
    Dim col As Collection
    Dim strError As String

    Set col = GetContainersForTypes("table_data", strError, True)
    TestAssert Len(strError) > 0, "table_data import rejected"
End Sub


Public Sub TestComponentTypeSupportsScopedImport()
    TestAssert ComponentTypeSupportsScopedImport(edbQuery), "queries supported"
    TestAssert Not ComponentTypeSupportsScopedImport(edbTableData), "table data unsupported"
End Sub


Public Sub TestGetContainersForTypesCanonicalOrder()
    Dim col As Collection
    Dim cCategory As IDbComponent
    Dim strError As String
    Dim lngPrev As Long
    Dim lngCurrent As Long

    Set col = GetContainersForTypes(Array("reports", "modules", "forms"), strError)
    TestAssert Len(strError) = 0, "no error"
    TestAssert col.Count = 3, "three categories"

    lngPrev = -1
    For Each cCategory In col
        lngCurrent = GetCanonicalContainerOrder(cCategory.ComponentType)
        TestAssert lngCurrent > lngPrev, "canonical order: " & cCategory.Category
        lngPrev = lngCurrent
    Next cCategory
End Sub


Private Function GetCanonicalContainerOrder(intType As eDatabaseComponentType) As Long

    Dim lngIdx As Long
    Dim cCont As IDbComponent

    lngIdx = 0
    For Each cCont In GetContainers()
        If cCont.ComponentType = intType Then
            GetCanonicalContainerOrder = lngIdx
            Exit Function
        End If
        lngIdx = lngIdx + 1
    Next cCont

    GetCanonicalContainerOrder = -1

End Function
