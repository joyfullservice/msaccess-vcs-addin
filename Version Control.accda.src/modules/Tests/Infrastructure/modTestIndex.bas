Attribute VB_Name = "modTestIndex"
'---------------------------------------------------------------------------------------
' Module    : modTestIndex
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Integration tests for clsVCSIndex: update, item retrieval, exists check.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("integration")


Public Sub TestIndexExists()
    Dim cModule As IDbComponent
    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub
    If Not FSO.FileExists(cModule.SourceFile) Then Exit Sub

    ' The add-in's own modules should be in the index (they were exported)
    TestAssert VCSIndex.Exists(cModule, cModule.SourceFile), _
        "exported module exists in index"
End Sub


Public Sub TestIndexItem_ReturnsObject()
    Dim cModule As IDbComponent
    Dim cIdx As Object
    Set cModule = GetTestComponent
    If cModule Is Nothing Then Exit Sub
    If Not FSO.FileExists(cModule.SourceFile) Then Exit Sub
    If Not VCSIndex.Exists(cModule, cModule.SourceFile) Then Exit Sub

    Set cIdx = VCSIndex.Item(cModule)
    TestAssert Not cIdx Is Nothing, "Item returns non-Nothing"
End Sub


Public Sub TestIndexExists_NonexistentItem()
    Dim cModule As IDbComponent
    Set cModule = New clsDbModule

    ' A component with no DbObject set shouldn't be in the index
    TestAssert Not VCSIndex.Exists(cModule, "C:\nonexistent\fake_module.bas"), _
        "nonexistent item not in index"
End Sub


Public Sub TestIndexConflictsObject()
    TestAssert Not VCSIndex.Conflicts Is Nothing, "Conflicts returns object"
    TestAssert VCSIndex.Conflicts.Count >= 0, "Conflicts count is non-negative"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestIndexSchemaStateRoundTrip
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : The SchemaState section drives re-export of external database schemas, so
'           : it has to survive a save/load cycle of the binary index format.
'---------------------------------------------------------------------------------------
'
Public Sub TestIndexSchemaStateRoundTrip()

    Dim cSave As clsVCSIndex
    Dim cLoad As clsVCSIndex
    Dim strFolder As String

    strFolder = GetTempFolder("VCSIdx")

    Set cSave = New clsVCSIndex
    cSave.SchemaState("Sample Connection") = "rev=1;ddl=sp_GetDDL"
    cSave.Save strFolder

    Set cLoad = New clsVCSIndex
    cLoad.LoadFromFile StripSlash(strFolder) & PathSep & "vcs-index.idx"

    TestAssert cLoad.SchemaState("Sample Connection") = "rev=1;ddl=sp_GetDDL", _
        "schema state survives index round trip"

    ' Schema names are matched without regard to case
    TestAssert cLoad.SchemaState("sample connection") = "rev=1;ddl=sp_GetDDL", _
        "schema state lookup is case insensitive"

    ' An unrecorded schema returns an empty string, which forces a baseline re-export
    TestAssert cLoad.SchemaState("Never Exported") = vbNullString, _
        "unrecorded schema returns empty state"

    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder StripSlash(strFolder), True
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTestComponent
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Helper to get a real module IDbComponent.
'---------------------------------------------------------------------------------------
'
Private Function GetTestComponent() As IDbComponent
    If CurrentProject.AllModules.Count = 0 Then Exit Function
    Dim cModule As IDbComponent
    Set cModule = New clsDbModule
    Set cModule.DbObject = CurrentProject.AllModules(0)
    Set GetTestComponent = cModule
End Function
