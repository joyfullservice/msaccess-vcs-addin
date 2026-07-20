Attribute VB_Name = "modTestContainers"
'---------------------------------------------------------------------------------------
' Module    : modTestContainers
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for source file metadata functions in modContainers:
'           : GetSourceModifiedDate, GetSourceFilesPropertyHash,
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

    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub

    ' Source file must exist for the test to be meaningful
    TestAssert FSO.FileExists(cModule.SourceFile), "source file exists on disk"

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

    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub

    TestAssert FSO.FileExists(cModule.SourceFile), "source file exists"

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

    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub

    TestAssert FSO.FileExists(cModule.SourceFile), "source file exists"

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

    Set cModule = GetTestComponent
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
