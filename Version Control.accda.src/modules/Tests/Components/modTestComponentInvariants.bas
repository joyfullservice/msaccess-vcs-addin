Attribute VB_Name = "modTestComponentInvariants"
'---------------------------------------------------------------------------------------
' Module    : modTestComponentInvariants
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : IDbComponent contract checks. Every component class must satisfy basic
'           : invariants: non-empty Category, valid ComponentType, unique BaseFolder, etc.
'           : Migrated from Private tests in modTestSuite.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("integration")


Public Sub TestComponentPropertyAccess()
    Dim colContainers As Collection
    Dim cnt As IDbComponent
    Dim varTest As Variant
    Set colContainers = GetContainers
    For Each cnt In colContainers
        varTest = cnt.Name
        varTest = cnt.DateModified
        varTest = cnt.SourceFile
        TestAssert cnt.DbObject Is Nothing, cnt.Category & " DbObject is Nothing when unset"
    Next
End Sub


Public Sub TestUniqueComponentCategory()
    Dim colContainers As Collection
    Dim dList As Dictionary
    Dim cnt As IDbComponent
    Set colContainers = GetContainers
    Set dList = New Dictionary
    For Each cnt In colContainers
        TestAssert Not dList.Exists(cnt.Category), cnt.Category & " category is unique"
        dList.Add cnt.Category, vbNullString
    Next
End Sub


Public Sub TestUniqueComponentType()
    Dim colContainers As Collection
    Dim dList As Dictionary
    Dim cnt As IDbComponent
    Set colContainers = GetContainers
    Set dList = New Dictionary
    For Each cnt In colContainers
        TestAssert Not dList.Exists(cnt.ComponentType), _
            "component type " & cnt.ComponentType & " is unique"
        dList.Add cnt.ComponentType, vbNullString
    Next
End Sub


Public Sub TestUniqueBaseSubfolder()
    Dim colContainers As Collection
    Dim dList As Dictionary
    Dim cnt As IDbComponent
    Set colContainers = GetContainers
    Set dList = New Dictionary
    For Each cnt In colContainers
        If Not cnt.SingleFile Then
            TestAssert Not dList.Exists(cnt.BaseFolder), _
                cnt.Category & " BaseFolder is unique"
            dList.Add cnt.BaseFolder, vbNullString
        End If
    Next
End Sub


Public Sub TestNonEmptyCategory()
    Dim colContainers As Collection
    Dim cnt As IDbComponent
    Set colContainers = GetContainers
    For Each cnt In colContainers
        TestAssert Len(cnt.Category) > 0, _
            "component type " & cnt.ComponentType & " has non-empty Category"
    Next
End Sub


Public Sub TestFileExtensionsNonEmpty()
    Dim colContainers As Collection
    Dim cnt As IDbComponent
    Dim colExts As Collection
    Set colContainers = GetContainers
    For Each cnt In colContainers
        Set colExts = cnt.FileExtensions
        TestAssert colExts.Count > 0, cnt.Category & " FileExtensions is non-empty"
    Next
End Sub


Public Sub TestFileExtensionScopeInvariant()
    Dim colContainers As Collection
    Dim cnt As IDbComponent
    Dim colIndexed As Collection
    Dim colAll As Collection
    Dim varExt As Variant

    Set colContainers = GetContainers
    For Each cnt In colContainers
        Set colIndexed = cnt.FileExtensions(efesIndexed)
        Set colAll = cnt.FileExtensions(efesAll)
        TestAssert colAll.Count >= colIndexed.Count, cnt.Category & " efesAll count >= efesIndexed"
        For Each varExt In colIndexed
            TestAssert ExtensionInIndexedScope(CStr(varExt), colAll), _
                cnt.Category & " indexed ext in efesAll: " & varExt
        Next varExt
    Next cnt
End Sub


Private Function ExtensionInIndexedScope(strExt As String, colAll As Collection) As Boolean
    Dim varItem As Variant
    For Each varItem In colAll
        If StrComp(CStr(varItem), strExt, vbTextCompare) = 0 Then
            ExtensionInIndexedScope = True
            Exit Function
        End If
    Next varItem
End Function


Public Sub TestFrmVCSTestRunnerSourceRequiresEdgeControl()
    Dim strFile As String
    strFile = Options.GetExportFolder & "forms\frmVCSTestRunner.form"
    TestAssert FSO.FileExists(strFile), "frmVCSTestRunner.form fixture exists"
    TestAssert FormSourceRequiresEdgeControl(strFile), "frmVCSTestRunner source requires Edge control"
End Sub


Public Sub TestFrmVCSMainSourceDoesNotRequireEdgeControl()
    Dim strFile As String
    strFile = Options.GetExportFolder & "forms\frmVCSMain.form"
    TestAssert FSO.FileExists(strFile), "frmVCSMain.form fixture exists"
    TestAssert Not FormSourceRequiresEdgeControl(strFile), "frmVCSMain source does not require Edge control"
End Sub


Public Sub TestEdgeTestRunnerSupportedOnModernAccess()
    If modTestRunnerUI.EdgeTestRunnerSupported() Then
        TestAssert True, "EdgeTestRunnerSupported on this Access build"
    End If
End Sub


Public Sub TestExporterRevisionsCategoryKeys()
    Dim colContainers As Collection
    Dim dCategories As Dictionary
    Dim dRevisions As Dictionary
    Dim varRevCat As Variant
    Dim cnt As IDbComponent

    Set dRevisions = GetExporterRevisions
    Set dCategories = New Dictionary
    Set colContainers = GetContainers
    For Each cnt In colContainers
        If Not dCategories.Exists(cnt.Category) Then dCategories.Add cnt.Category, vbNullString
    Next cnt
    For Each varRevCat In dRevisions.Keys
        TestAssert dCategories.Exists(CStr(varRevCat)), _
            "GetExporterRevisions key matches a component category: " & varRevCat
    Next varRevCat
End Sub


Public Sub TestExporterRevisionsInCategoryHashes()
    Dim dHashes As Dictionary
    Dim dRevisions As Dictionary
    Dim varRevCat As Variant

    Set dRevisions = GetExporterRevisions
    Set dHashes = Options.GetCategoryHashes
    For Each varRevCat In dRevisions.Keys
        TestAssert dHashes.Exists(CStr(varRevCat)), _
            "GetCategoryHashes includes revisioned category: " & varRevCat
    Next varRevCat
End Sub
