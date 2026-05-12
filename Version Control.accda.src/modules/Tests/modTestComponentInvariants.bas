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
