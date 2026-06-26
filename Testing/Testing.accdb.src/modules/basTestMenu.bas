Attribute VB_Name = "basTestMenu"
'---------------------------------------------------------------------------------------
' Module    : basTestMenu
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Build the "VCS Test Menu" command bar for round-trip and manual testing
'           : of hybrid built-in classification (custom, minimal built-in, customized
'           : replica, template-copy object-opener, nested popup). Run CreateTestMenu
'           : once with the add-in loaded, then export menus to source.
'           : Immediate window: CreateTestMenu
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Const TestMenuName As String = "VCS Test Menu"

Private Const TemplateBarName As String = "MSAccessVCSCustomBuiltinCommandBarTemplate"

Public Const IdPrint As Long = 4
Public Const IdClose As Long = 106
Public Const IdOpenTable As Long = 1835
Public Const IdCompactRepair As Long = 2071


'---------------------------------------------------------------------------------------
' Procedure : CreateTestMenu
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Create or replace the VCS Test Menu popup covering every export branch.
'---------------------------------------------------------------------------------------
'
Public Sub CreateTestMenu()

    Dim bar As CommandBar
    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup
    Dim tmplBar As CommandBar
    Dim tmplCtl As CommandBarControl

    RemoveTestMenu

    Set bar = CommandBars.Add(TestMenuName, msoBarPopup)

    ' Custom control (BuiltIn=false path).
    Set ctl = bar.Controls.Add(msoControlButton)
    With ctl
        .Caption = "Custom Button"
        .OnAction = "=MenuHandler(""VCS Test Menu.Custom Button"")"
    End With

    ' Addable clean built-in -> minimal export (Print, Id 4).
    bar.Controls.Add msoControlButton, IdPrint

    ' Addable built-in with custom OnAction -> replica export (#584).
    Set ctl = bar.Controls.Add(msoControlButton, IdCompactRepair)
    ctl.OnAction = "=MenuHandler(""VCS Test Menu.Customized Compact"")"

    ' Non-addable object-opener copied from the add-in template bar.
    On Error Resume Next
    Set tmplBar = Application.CommandBars(TemplateBarName)
    If tmplBar Is Nothing Then
        Debug.Print "CreateTestMenu: template bar not loaded; skipping Open Table (1835)"
    Else
        Set tmplCtl = tmplBar.FindControl(ID:=IdOpenTable)
        If tmplCtl Is Nothing Then
            Debug.Print "CreateTestMenu: Open Table (1835) not on template bar"
        Else
            Set ctl = tmplCtl.Copy(bar)
            ' Point at a table that exists in the Testing database.
            ctl.Parameter = "tblInternal"
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' Custom popup with a custom child and a built-in child.
    Set popup = bar.Controls.Add(msoControlPopup)
    popup.Caption = "More"
    Set ctl = popup.Controls.Add(msoControlButton)
    With ctl
        .Caption = "Sub Custom"
        .OnAction = "=MenuHandler(""VCS Test Menu.More.Sub Custom"")"
    End With
    popup.Controls.Add msoControlButton, IdClose

    Debug.Print "CreateTestMenu: created """ & TestMenuName & """ with " & bar.Controls.Count & " top-level controls."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveTestMenu
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Delete the test menu if it exists.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveTestMenu()
    On Error Resume Next
    CommandBars(TestMenuName).Delete
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShowTestMenu
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Show the test menu at the cursor (Immediate window helper).
'---------------------------------------------------------------------------------------
'
Public Sub ShowTestMenu()
    On Error Resume Next
    CommandBars(TestMenuName).ShowPopup
    If Err.Number <> 0 Then Debug.Print "ShowTestMenu error: " & Err.Description
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindControlById
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Return the first control with the given Id on a bar (including popups).
'---------------------------------------------------------------------------------------
'
Public Function FindControlById(bar As CommandBar, lngId As Long) As CommandBarControl
    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup

    For Each ctl In bar.Controls
        If ctl.ID = lngId Then
            Set FindControlById = ctl
            Exit Function
        End If
        If TypeOf ctl Is CommandBarPopup Then
            Set popup = ctl
            Set FindControlById = FindControlByIdInControls(popup.Controls, lngId)
            If Not FindControlById Is Nothing Then Exit Function
        End If
    Next ctl
End Function


Private Function FindControlByIdInControls(ctls As CommandBarControls, lngId As Long) As CommandBarControl
    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup

    For Each ctl In ctls
        If ctl.ID = lngId Then
            Set FindControlByIdInControls = ctl
            Exit Function
        End If
        If TypeOf ctl Is CommandBarPopup Then
            Set popup = ctl
            Set FindControlByIdInControls = FindControlByIdInControls(popup.Controls, lngId)
            If Not FindControlByIdInControls Is Nothing Then Exit Function
        End If
    Next ctl
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindControlByCaption
' Author    : Adam Waller
' Date      : 6/26/2026
' Purpose   : Return the first control whose caption matches (ignoring & mnemonics).
'---------------------------------------------------------------------------------------
'
Public Function FindControlByCaption(bar As CommandBar, strCaption As String) As CommandBarControl
    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup

    For Each ctl In bar.Controls
        If StripMnemonic(ctl.Caption) = strCaption Then
            Set FindControlByCaption = ctl
            Exit Function
        End If
        If TypeOf ctl Is CommandBarPopup Then
            Set popup = ctl
            Set FindControlByCaption = FindControlByCaptionInControls(popup.Controls, strCaption)
            If Not FindControlByCaption Is Nothing Then Exit Function
        End If
    Next ctl
End Function


Private Function FindControlByCaptionInControls(ctls As CommandBarControls, strCaption As String) As CommandBarControl
    Dim ctl As CommandBarControl
    Dim popup As CommandBarPopup

    For Each ctl In ctls
        If StripMnemonic(ctl.Caption) = strCaption Then
            Set FindControlByCaptionInControls = ctl
            Exit Function
        End If
        If TypeOf ctl Is CommandBarPopup Then
            Set popup = ctl
            Set FindControlByCaptionInControls = FindControlByCaptionInControls(popup.Controls, strCaption)
            If Not FindControlByCaptionInControls Is Nothing Then Exit Function
        End If
    Next ctl
End Function


Private Function StripMnemonic(strCaption As String) As String
    StripMnemonic = Replace(strCaption, "&", vbNullString)
End Function
