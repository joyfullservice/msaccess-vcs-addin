Attribute VB_Name = "modTestCommandBar"
'---------------------------------------------------------------------------------------
' Module    : modTestCommandBar
' Author    : Adam Waller
' Date      : 7/16/2026
' Purpose   : Command bar import/export round-trip and classification tests for
'           : customized built-in controls (split buttons, combos, issue #584).
'           : Run: ?VCS.RunTests("modTestCommandBar")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("access")


Private Const ModuleName As String = "modTestCommandBar"
Private Const FIXTURE_RELATIVE_PATH As String = "Testing\Fixtures\menus\regression\"


Public Sub TestFaceIdReconstructableDiscriminator()
    TestAssert Len(ControlIdToName(293)) > 0, "FaceId 293 is reconstructable"
    TestAssert Len(ControlIdToName(50000)) = 0, "FaceId 50000 is not reconstructable"
End Sub


Public Sub TestIdAddableOnlyControlTypes()
    '@Tag("slow")
    AssertAddableByIdOnly msoControlSplitButtonPopup, 212, "View split (13)"
    AssertAddableByIdOnly msoControlSplitButtonMRUPopup, 2598, "Office Links (14)"
    AssertAddableByIdOnly msoControlExpandingGrid, 177, "Multiple Pages (16)"
    AssertAddableByIdOnly msoControlGrid, 3079, "Font/Fore Color (18)"
    AssertNotBlankAddable msoControlSplitButtonPopup, "split button popup (13)"
    AssertNotBlankAddable msoControlSplitButtonMRUPopup, "split MRU popup (14)"
End Sub


Public Sub TestMenuFixtureRoundtripCustomizedSplitButton()
    '@Tag("slow")
    RunMenuFixtureRoundtrip "customized_split_button.json", 13, 212
End Sub


Public Sub TestMenuFixtureRoundtripOfficeLinksZoom()
    '@Tag("slow")
    RunMenuFixtureRoundtrip "office_links_zoom.json", 14, 2598
    RunMenuFixtureRoundtrip "office_links_zoom.json", 4, 1733
End Sub


Public Sub TestMenuFixtureRoundtripCustomComboList()
    '@Tag("slow")
    ' A custom (BuiltIn:false) combo box with a user-defined list must repopulate its
    ' items on import (regression: the combo-list import assignment was reversed) and
    ' round-trip the List array idempotently. ListCount is derived, not serialized.
    Const strFixtureFile As String = "custom_combo_list.json"

    Dim cBar As IDbComponent
    Dim strFixture As String
    Dim strBarName As String
    Dim strPass1 As String
    Dim strPass2 As String
    strPass1 = vbNullString
    strPass2 = vbNullString
    Dim lngErrBefore As Long
    Dim bar As CommandBar
    Dim cbo As CommandBarComboBox
    Dim dRoot As Dictionary
    Dim dItems As Dictionary
    Dim colControls As Collection
    Dim dCtl As Dictionary
    Dim colList As Collection

    strFixture = GetMenuFixturePath(strFixtureFile)
    If Not FSO.FileExists(strFixture) Then
        TestAssert True, "SKIP: fixture not found: " & strFixture
        Exit Sub
    End If

    strBarName = FSO.GetBaseName(strFixture)
    DeleteCommandBarIfExists strBarName

    lngErrBefore = Log.ErrorCount
    Set cBar = New clsDbCommandBar
    cBar.Import strFixture
    TestAssert Log.ErrorCount = lngErrBefore, "import logged no errors"

    Set bar = Nothing
    LogUnhandledErrors
    On Error Resume Next
    Set bar = CommandBars(strBarName)
    If CatchAny(eelNoError, vbNullString) Or bar Is Nothing Then
        TestAssert False, "command bar not created: " & strBarName
        GoTo CleanUp
    End If
    On Error GoTo 0

    Set cbo = FindComboControl(bar)
    TestAssert Not cbo Is Nothing, "custom combo control exists after import"
    If Not cbo Is Nothing Then
        TestAssert Not cbo.BuiltIn, "combo is custom (not built-in)"
        TestAssert cbo.ListCount = 3, "combo list repopulated with 3 items (import fix)"
        If cbo.ListCount >= 3 Then
            TestAssert cbo.List(1) = "Alpha", "list item 1 preserved"
            TestAssert cbo.List(2) = "Beta", "list item 2 preserved"
            TestAssert cbo.List(3) = "Gamma", "list item 3 preserved"
        End If
    End If

    strPass1 = BuildPath2(Environ("TEMP"), strBarName & "_pass1.json")
    strPass2 = BuildPath2(Environ("TEMP"), strBarName & "_pass2.json")
    cBar.Export strPass1
    cBar.Export strPass2
    TestAssert FSO.FileExists(strPass1), "pass 1 export created"
    TestAssert GetStringHash(ReadFile(strPass1), True) = GetStringHash(ReadFile(strPass2), True), _
        "pass 1 and pass 2 exports are identical (idempotent)"

    Set dRoot = ReadJsonFile(strPass1)
    Set dItems = dRoot("Items")
    Set colControls = dItems("Controls")
    Set dCtl = FindControlWithList(colControls)
    If dCtl Is Nothing Then
        TestAssert False, "combo List serialized in export"
    Else
        Set colList = dCtl("List")
        TestAssert colList.Count = 3, "exported List has 3 items"
    End If

CleanUp:
    DeleteCommandBarIfExists strBarName
    On Error Resume Next
    FSO.DeleteFile strPass1, True
    FSO.DeleteFile strPass2, True
    Err.Clear
End Sub


Public Sub TestIssue584CustomizedBuiltInButtonRoundtrip()
    '@Tag("slow")
    ' Assigning a custom OnAction to a built-in button makes Access flip BuiltIn to
    ' False; the control keeps its (reconstructable) FaceId. #584 requires the icon
    ' and action to survive export, and the export to be idempotent.
    Const strBarName As String = "vcs_test_issue584"
    Const IdCompactRepair As Long = 2071
    Const strOnAction As String = "=Issue584TestHandler()"

    Dim cBar As IDbComponent
    Dim strPass1 As String
    Dim strPass2 As String
    strPass1 = vbNullString
    strPass2 = vbNullString
    Dim dRoot As Dictionary
    Dim dItems As Dictionary
    Dim colControls As Collection
    Dim dCtl As Dictionary
    Dim lngErrBefore As Long
    Dim bar As CommandBar
    Dim ctl As CommandBarControl

    DeleteCommandBarIfExists strBarName
    Set bar = CommandBars.Add(strBarName, msoBarPopup)
    Set ctl = bar.Controls.Add(msoControlButton, IdCompactRepair)
    ctl.OnAction = strOnAction

    lngErrBefore = Log.ErrorCount
    Set cBar = New clsDbCommandBar
    Set cBar.DbObject = bar

    strPass1 = BuildPath2(Environ("TEMP"), "vcs_test_issue584_pass1.json")
    strPass2 = BuildPath2(Environ("TEMP"), "vcs_test_issue584_pass2.json")
    cBar.Export strPass1
    cBar.Export strPass2

    TestAssert Log.ErrorCount = lngErrBefore, "export logged no errors"
    TestAssert FSO.FileExists(strPass1), "export file created"
    TestAssert GetStringHash(ReadFile(strPass1), True) = GetStringHash(ReadFile(strPass2), True), _
        "export is idempotent"

    Set dRoot = ReadJsonFile(strPass1)
    Set dItems = dRoot("Items")
    Set colControls = dItems("Controls")
    Set dCtl = FindControlByOnAction(colControls, strOnAction)
    If dCtl Is Nothing Then
        TestAssert False, "customized control preserved in export (matched by OnAction)"
    Else
        TestAssert dCtl.Exists("FaceId"), "reconstructable FaceId preserved"
        If dCtl.Exists("FaceId") Then TestAssert dCtl("FaceId") = IdCompactRepair, "FaceId matches built-in icon"
    End If

    DeleteCommandBarIfExists strBarName
    On Error Resume Next
    FSO.DeleteFile strPass1, True
    FSO.DeleteFile strPass2, True
    Err.Clear
End Sub


Private Sub RunMenuFixtureRoundtrip(strFixtureFile As String, lngType As Long, lngId As Long)

    Dim cBar As IDbComponent
    Dim strFixture As String
    Dim strBarName As String
    Dim strPass1 As String
    Dim strPass2 As String
    strPass1 = vbNullString
    strPass2 = vbNullString
    Dim lngErrBefore As Long
    Dim bar As CommandBar
    Dim ctl As CommandBarControl

    strFixture = GetMenuFixturePath(strFixtureFile)
    If Not FSO.FileExists(strFixture) Then
        TestAssert True, "SKIP: fixture not found: " & strFixture
        Exit Sub
    End If

    strBarName = FSO.GetBaseName(strFixture)
    DeleteCommandBarIfExists strBarName

    lngErrBefore = Log.ErrorCount
    Set cBar = New clsDbCommandBar
    cBar.Import strFixture

    TestAssert Log.ErrorCount = lngErrBefore, "import logged no errors for " & strFixtureFile

    Set bar = Nothing
    LogUnhandledErrors
    On Error Resume Next
    Set bar = CommandBars(strBarName)
    If CatchAny(eelNoError, vbNullString) Or bar Is Nothing Then
        TestAssert False, "command bar not created: " & strBarName
        GoTo CleanUp
    End If

    Set ctl = FindControlById(bar, lngId)
    TestAssert Not ctl Is Nothing, "control Id " & lngId & " exists after import"
    If Not ctl Is Nothing Then
        TestAssert ctl.Type = lngType, "control type matches"
        TestAssert ctl.BuiltIn, "control is built-in after import"
    End If

    strPass1 = BuildPath2(Environ("TEMP"), strBarName & "_pass1.json")
    strPass2 = BuildPath2(Environ("TEMP"), strBarName & "_pass2.json")
    cBar.Export strPass1
    cBar.Export strPass2
    TestAssert FSO.FileExists(strPass1), "pass 1 export created"
    TestAssert FSO.FileExists(strPass2), "pass 2 export created"
    TestAssert GetStringHash(ReadFile(strPass1), True) = GetStringHash(ReadFile(strPass2), True), _
        "pass 1 and pass 2 exports are identical (idempotent)"

CleanUp:
    DeleteCommandBarIfExists strBarName
    On Error Resume Next
    FSO.DeleteFile strPass1, True
    FSO.DeleteFile strPass2, True
    Err.Clear

End Sub


Private Function GetMenuFixturePath(strFile As String) As String
    GetMenuFixturePath = CodeProject.Path & PathSep & ".." & PathSep & FIXTURE_RELATIVE_PATH & strFile
End Function


Private Sub DeleteCommandBarIfExists(strName As String)
    On Error Resume Next
    CommandBars(strName).Delete
    Err.Clear
End Sub


Private Function FindControlByOnAction(colControls As Collection, strOnAction As String) As Dictionary

    Dim dCtl As Dictionary
    Dim dChild As Dictionary
    Dim colChildren As Collection

    For Each dCtl In colControls
        If dCtl.Exists("OnAction") Then
            If dCtl("OnAction") = strOnAction Then
                Set FindControlByOnAction = dCtl
                Exit Function
            End If
        End If
        If dCtl.Exists("Controls") Then
            Set colChildren = dCtl("Controls")
            Set dChild = FindControlByOnAction(colChildren, strOnAction)
            If Not dChild Is Nothing Then
                Set FindControlByOnAction = dChild
                Exit Function
            End If
        End If
    Next dCtl

End Function


Private Function FindComboControl(bar As CommandBar) As CommandBarComboBox
    Dim ctl As CommandBarControl
    For Each ctl In bar.Controls
        If ctl.Type = msoControlComboBox Then
            Set FindComboControl = ctl
            Exit Function
        End If
    Next ctl
End Function


Private Function FindControlWithList(colControls As Collection) As Dictionary
    Dim dCtl As Dictionary
    For Each dCtl In colControls
        If dCtl.Exists("List") Then
            Set FindControlWithList = dCtl
            Exit Function
        End If
    Next dCtl
End Function


Private Sub AssertAddableByIdOnly(lngType As Long, lngId As Long, strLabel As String)
    TestAssert ProbeBuiltInControlAddable(lngType, lngId), strLabel & " should be addable by Id"
    TestAssert Not ProbeBlankControlAddable(lngType), strLabel & " should not be blank-addable"
End Sub


Private Sub AssertNotBlankAddable(lngType As Long, strLabel As String)
    TestAssert Not ProbeBlankControlAddable(lngType), strLabel & " should not be blank-addable"
End Sub


Private Function FindControlById(bar As CommandBar, lngId As Long) As CommandBarControl
    FindControlByIdInControls bar.Controls, lngId, FindControlById
End Function


Private Sub FindControlByIdInControls(ctls As CommandBarControls, lngId As Long, ByRef ctlResult As CommandBarControl)

    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup

    For Each ctl In ctls
        If ctl.Id = lngId Then
            Set ctlResult = ctl
            Exit Sub
        End If
        If TypeOf ctl Is CommandBarPopup Then
            Set popup = ctl
            FindControlByIdInControls popup.Controls, lngId, ctlResult
            If Not ctlResult Is Nothing Then Exit Sub
        End If
    Next ctl

End Sub


Private Function ProbeBlankControlAddable(lngType As Long) As Boolean

    Const strTempBar As String = "__VCS_BlankProbeBar__"

    Dim cb As CommandBar
    Dim ctl As CommandBarControl

    LogUnhandledErrors
    On Error Resume Next
    Application.CommandBars(strTempBar).Delete
    CatchAny eelNoError, vbNullString
    Set cb = Application.CommandBars.Add(strTempBar, msoBarFloating, False, True)
    If cb Is Nothing Then Exit Function
    Set ctl = cb.Controls.Add(Type:=lngType, Temporary:=True)
    ProbeBlankControlAddable = (Err.Number = 0) And Not ctl Is Nothing
    cb.Delete
    CatchAny eelNoError, vbNullString

End Function
