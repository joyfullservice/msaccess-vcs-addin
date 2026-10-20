Attribute VB_Name = "modTestVbaProjectDate"
'---------------------------------------------------------------------------------------
' Module    : modTestVbaProjectDate
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/6/2026
' Purpose   : VBAProjectDate tells the module, form, and report change detection that
'           : the whole VBA project matches the index, so it can skip hashing code.
'           : Recording a single object must not move it: a merge conflict scan that
'           : temp-exported one modified module used to vouch for every module checked
'           : after it, and their changes in the database were overwritten without a
'           : conflict. These tests use a private index instance, never the live one.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("integration")

' A date no real project carries, so any refresh from the project is visible.
Private Const cdteStale As Date = #1/1/2000#


Public Sub TestAltExportOfModuleKeepsVBAProjectDate()
    AssertUpdateKeepsDate GetTestModuleComponent, eatAltExport, "module, alternate export"
End Sub


Public Sub TestExportOfModuleKeepsVBAProjectDate()
    AssertUpdateKeepsDate GetTestModuleComponent, eatExport, "module, export"
End Sub


Public Sub TestImportOfModuleKeepsVBAProjectDate()
    AssertUpdateKeepsDate GetTestModuleComponent, eatImport, "module, import"
End Sub


Public Sub TestAltExportOfFormKeepsVBAProjectDate()
    AssertUpdateKeepsDate GetTestFormComponent, eatAltExport, "form, alternate export"
End Sub


Public Sub TestExportOfFormKeepsVBAProjectDate()
    AssertUpdateKeepsDate GetTestFormComponent, eatExport, "form, export"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCertifyRecordsDateOfSavedProject
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/6/2026
' Purpose   : An operation that checked the whole project certifies the project date.
'           : Only meaningful while the project is saved, so it stays empty otherwise.
'---------------------------------------------------------------------------------------
'
Public Sub TestCertifyRecordsDateOfSavedProject()

    Dim cIndex As clsVCSIndex
    Dim dteProject As Date
    Dim blnRecorded As Boolean

    If Not CurrentVBProject.Saved Then Exit Sub
    dteProject = GetVbaProjectDate
    If dteProject = 0 Then Exit Sub

    Set cIndex = New clsVCSIndex
    cIndex.VBAProjectDate = cdteStale
    cIndex.CertifyVBAProjectDate dteProject

    blnRecorded = (cIndex.VBAProjectDate = dteProject)
    TestAssert blnRecorded, "certifying the current project date records it"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCertifyIgnoresProjectSavedSinceCheck
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/6/2026
' Purpose   : If the project was saved after the date the operation captured (a hook or
'           : a casing correction changed code), the certificate must not be issued.
'---------------------------------------------------------------------------------------
'
Public Sub TestCertifyIgnoresProjectSavedSinceCheck()

    Dim cIndex As clsVCSIndex
    Dim dteProject As Date
    Dim blnKept As Boolean

    dteProject = GetVbaProjectDate
    If dteProject = 0 Then Exit Sub

    Set cIndex = New clsVCSIndex
    cIndex.VBAProjectDate = cdteStale
    cIndex.CertifyVBAProjectDate DateAdd("s", -1, dteProject)

    blnKept = (cIndex.VBAProjectDate = cdteStale)
    TestAssert blnKept, "a date older than the project is not certified"

End Sub


Public Sub TestCertifyIgnoresMissingDate()

    Dim cIndex As clsVCSIndex
    Dim blnKept As Boolean

    Set cIndex = New clsVCSIndex
    cIndex.VBAProjectDate = cdteStale
    cIndex.CertifyVBAProjectDate 0

    blnKept = (cIndex.VBAProjectDate = cdteStale)
    TestAssert blnKept, "an operation that captured no date certifies nothing"

End Sub


Public Sub TestCertifyIgnoresDisabledIndex()

    Dim cIndex As clsVCSIndex
    Dim blnKept As Boolean

    Set cIndex = New clsVCSIndex
    cIndex.Disabled = True
    cIndex.VBAProjectDate = cdteStale
    cIndex.CertifyVBAProjectDate GetVbaProjectDate

    blnKept = (cIndex.VBAProjectDate = cdteStale)
    TestAssert blnKept, "a disabled index certifies nothing"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertUpdateKeepsDate
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/6/2026
' Purpose   : Record one component in a private index and check that VBAProjectDate
'           : is left where it was.
'---------------------------------------------------------------------------------------
'
Private Sub AssertUpdateKeepsDate(cItem As IDbComponent, intAction As eIndexOperationType, _
    strContext As String)

    Dim cIndex As clsVCSIndex
    Dim blnKept As Boolean

    If cItem Is Nothing Then Exit Sub

    Set cIndex = New clsVCSIndex
    cIndex.VBAProjectDate = cdteStale
    cIndex.Update cItem, intAction, "hash"

    blnKept = (cIndex.VBAProjectDate = cdteStale)
    TestAssert blnKept, "recording one object leaves VBAProjectDate unchanged (" & strContext & ")"

End Sub


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
