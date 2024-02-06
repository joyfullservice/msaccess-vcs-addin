Attribute VB_Name = "modVbeForm"
'---------------------------------------------------------------------------------------
' Module    : modVbeForm
' Author    : Adam Waller / Adapted from FormSerializer
' Date      : 1/24/2022
' Purpose   : Serialize a MSForms 2.0 form into human-readable JSON output.
'---------------------------------------------------------------------------------------

''
' FormSerializer v1.0.0
' (c) Georges Kuenzli - https://github.com/gkuenzli/vbaDeveloper
'
' `FormSerializer` produces a string JSON description of a MSForm.
'
' @module FormSerializer
' @author gkuenzli
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Compare Database
Option Private Module
Option Explicit


''
' Convert a VBComponent of type MSForm to a JSON descriptive data
'
' @method SerializeMSForm
' @param {VBComponent} FormComponent
' @return {Dictionary} MSForm JSON descriptive dictionary
''
Public Function SerializeMSForm(ByVal FormComponent As VBComponent) As Dictionary
    Set SerializeMSForm = GetMSFormProperties(FormComponent)
End Function

Private Function GetMSFormProperties(ByVal FormComponent As VBComponent) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Name", FormComponent.Name
    dict.Add "Designer", GetDesigner(FormComponent)
    dict.Add "Properties", GetProperties(FormComponent, FormComponent.Properties)
    Set GetMSFormProperties = dict
End Function

Private Function GetDesigner(ByVal FormComponent As VBComponent) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Controls", GetControls(FormComponent.Designer.Controls)
    Set GetDesigner = dict
End Function

Private Function GetProperties(ByVal Context As Object, ByVal Properties As VBIDE.Properties) As Dictionary
    Dim dict As New Dictionary
    Dim p As VBIDE.Property
    Dim i As Long
    For i = 1 To Properties.Count
        Set p = Properties(i)
        If IsSerializableProperty(Context, p) Then
            dict.Add p.Name, GetValue(Context, p)
        End If
    Next i
    Set GetProperties = dict
End Function

Private Function IsSerializableProperty(ByVal Context As Object, ByVal Property As VBIDE.Property) As Boolean
    Dim tp As VbVarType
    On Error Resume Next
    tp = VarType(Property.Value)
    On Error GoTo 0
    IsSerializableProperty = _
        (tp <> vbEmpty) And (tp <> vbError) And _
        Left(Property.Name, 1) <> "_" And _
        InStr("ActiveControls,Controls,Handle,MouseIcon,Picture,Selected,DesignMode,ShowToolbox,ShowGridDots,SnapToGrid,GridX,GridY,DrawBuffer,CanPaste", Property.Name) = 0

    If TypeName(Context) = "VBComponent" Then
        ' We must ignore Top and Height MSForm properties since these seem to be related to the some settings in the Windows user profile.
        IsSerializableProperty = _
            IsSerializableProperty And _
            InStr("Top,Height", Property.Name) = 0
    End If
End Function

Private Function GetProperty(ByVal Context As Object, ByVal Property As VBIDE.Property) As Dictionary
    Dim dict As New Dictionary
    dict.Add "Name", Property.Name
    If Property.Name = "Controls" Then
    Else
        dict.Add "Value", GetValue(Context, Property)
    End If
    Set GetProperty = dict
End Function

Private Function GetControls(ByVal Controls As MSForms.Controls) As Collection
    Dim coll As New Collection
    Dim ctrl As MSForms.Control
    For Each ctrl In Controls
        If Not ControlExistsInSubElements(Controls, ctrl.Name, 0) Then
            coll.Add GetControl(ctrl)
        End If
    Next ctrl
    Set GetControls = coll
End Function

Private Function ControlExistsInSubElements(ByVal Controls As MSForms.Controls, ByVal Name As String, ByVal Depth As Long) As Boolean
    Dim ctrl As MSForms.Control
    Dim o As Object
    For Each ctrl In Controls
        Set o = ctrl
        If Depth > 0 Then
            If Name = ctrl.Name Then
                ControlExistsInSubElements = True
                Exit Function
            End If
        End If
        On Error Resume Next
        ControlExistsInSubElements = ControlExistsInSubElements(o.Controls, Name, Depth + 1)
        On Error GoTo 0
        If ControlExistsInSubElements Then
            Exit Function
        End If
    Next ctrl
End Function

Private Function GetControl(ByVal ctl As Object) As Dictionary ' MSForms.Control
    Dim dic As Dictionary
    Dim varName As Variant
    Set dic = New Dictionary
    ' Loop through properties, adding each value to dictionary
    For Each varName In GetPropertyList(TypeName(ctl))
        AddProperty dic, ctl, varName
    Next varName
    ' Return dictionary of control properties
    Set GetControl = dic
End Function

Private Sub AddProperty(dic As Dictionary, o As Object, strName As Variant)
    Select Case strName
        Case "Class":       dic.Add strName, TypeName(o)
        Case "Font":        dic.Add strName, GetFont(o.Font)
        Case "Picture":     dic.Add strName, GetPicture(o.Picture)
        Case "MouseIcon":   dic.Add strName, GetPicture(o.MouseIcon)
        Case "Controls":    dic.Add strName, GetControls(o.Controls)
        Case "Pages":       dic.Add strName, GetPages(o.Pages)
        Case "Tabs":        dic.Add strName, GetTabs(o.Tabs)
        Case Else
            ' Standard property.
            ' Use CallByName on object to get value if the property exists
            On Error Resume Next
            dic.Add strName, CallByName(o, strName, VbGet)
            If Err Then Err.Clear
    End Select
End Sub

Private Function GetPages(ByVal Pages As MSForms.Pages) As Collection
    Dim coll As New Collection
    Dim i As Long
    Dim p As MSForms.Page
    For i = 0 To Pages.Count - 1
        Set p = Pages(i)
        coll.Add GetControl(p)
    Next i
    Set GetPages = coll
End Function

Private Function GetTabs(ByVal Tabs As Tabs) As Collection
    Dim coll As New Collection
    Dim i As Long
    Dim t As MSForms.Tab
    For i = 0 To Tabs.Count - 1
        Set t = Tabs(i)
        coll.Add GetControl(t)
    Next i
    Set GetTabs = coll
End Function

Private Function GetFont(ByVal fnt As NewFont) As Dictionary
    Set GetFont = New Dictionary
    With GetFont
        .Add "Bold", fnt.Bold
        .Add "Charset", fnt.Charset
        .Add "Italic", fnt.Italic
        .Add "Name", fnt.Name
        .Add "Size", fnt.Size
        .Add "Strikethrough", fnt.Strikethrough
        .Add "Underline", fnt.Underline
        .Add "Weight", fnt.Weight
    End With
End Function

Private Function GetPicture(ByVal Picture As IPictureDisp) As String

    ' TODO: implement a Base64-encoding of the picture
    'StdFunctions.SavePicture Picture, strFileName

End Function

Private Function GetValue(ByVal Context As Object, ByVal Property As VBIDE.Property) As Variant
    If VarType(Property.Value) = vbObject Then
        Select Case TypeName(Property.Value)
            Case "Properties"
                Set GetValue = GetProperties(Context, Property.Value)
            Case Else
                Set GetValue = Nothing
        End Select
    Else
        GetValue = Property.Value
    End If
End Function

Private Function GetPropertyList(strType As String) As Collection

    Set GetPropertyList = New Collection
    With GetPropertyList

        ' Generic control level properties
        .Add "Class"
        .Add "Name"
        .Add "Cancel"
        .Add "ControlSource"
        .Add "ControlTipText"
        .Add "Default"
        .Add "Height"
        .Add "HelpContextID"
        .Add "LayoutEffect"
        .Add "Left"
        .Add "RowSource"
        .Add "RowSourceType"
        .Add "TabIndex"
        .Add "TabStop"
        .Add "Tag"
        .Add "Top"
        .Add "Visible"
        .Add "Width"

        ' Specific properties based on control type
        Select Case strType
            Case "CheckBox"
                .Add "Accelerator"
                .Add "Alignment"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "Caption"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "GroupName"
                .Add "Locked"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PicturePosition"
                .Add "SpecialEffect"
                .Add "TextAlign"
                .Add "TripleState"
                .Add "Value"
                .Add "WordWrap"

            Case "ComboBox", "RefEdit"  ' (Also used for Excel Reference control)
                .Add "AutoSize"
                .Add "AutoTab"
                .Add "AutoWordSelect"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "BorderColor"
                .Add "BorderStyle"
                .Add "BoundColumn"
            '    .Add "CanPaste"
                .Add "ColumnCount"
                .Add "ColumnHeads"
                .Add "ColumnWidths"
                .Add "DragBehavior"
                .Add "DropButtonStyle"
                .Add "Enabled"
                .Add "EnterFieldBehavior"
                .Add "Font"
                .Add "ForeColor"
                .Add "HideSelection"
                .Add "IMEMode"
                .Add "ListRows"
                .Add "ListStyle"
                .Add "ListWidth"
                .Add "Locked"
                .Add "MatchEntry"
                .Add "MatchRequired"
                .Add "MaxLength"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "SelectionMargin"
                .Add "ShowDropButtonWhen"
                .Add "SpecialEffect"
                .Add "Style"
                .Add "Text"
                .Add "TextAlign"
                .Add "TextColumn"
                .Add "TopIndex"
                .Add "Value"

            Case "CommandButton"
                .Add "Accelerator"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "Caption"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "Locked"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PicturePosition"
                .Add "TakeFocusOnClick"
                .Add "WordWrap"

            Case "Frame"
                .Add "BackColor"
                .Add "BorderColor"
                .Add "BorderStyle"
                '.Add "CanPaste"
                .Add "CanRedo"
                .Add "CanUndo"
                .Add "Caption"
                .Add "Controls"
                .Add "Cycle"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "InsideHeight"
                .Add "InsideWidth"
                .Add "KeepScrollBarsVisible"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PictureAlignment"
                .Add "PictureSizeMode"
                .Add "PictureTiling"
                .Add "ScrollBars"
                .Add "ScrollHeight"
                .Add "ScrollLeft"
                .Add "ScrollTop"
                .Add "ScrollWidth"
                .Add "SpecialEffect"
                .Add "VerticalScrollBarSide"
                .Add "Zoom"

            Case "Image"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "BorderColor"
                .Add "BorderStyle"
                .Add "Enabled"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PictureAlignment"
                .Add "PictureSizeMode"
                .Add "PictureTiling"
                .Add "SpecialEffect"

            Case "Label"
                .Add "Accelerator"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "BorderColor"
                .Add "BorderStyle"
                .Add "Caption"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PicturePosition"
                .Add "SpecialEffect"
                .Add "TextAlign"
                .Add "WordWrap"

            Case "ListBox"
                .Add "BackColor"
                .Add "BorderColor"
                .Add "BorderStyle"
                .Add "BoundColumn"
                .Add "ColumnHeads"
                .Add "ColumnWidths"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "IMEMode"
                .Add "IntegralHeight"
                .Add "ListIndex"
                .Add "ListStyle"
                .Add "Locked"
                .Add "MatchEntry"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "MultiSelect"
                .Add "Selected"
                .Add "SpecialEffect"
                .Add "Text"
                .Add "TextAlign"
                .Add "TextColumn"
                .Add "TopIndex"
                .Add "Value"

            Case "MultiPage"
                .Add "BackColor"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "MultiRow"
                .Add "Pages"
                .Add "Style"
                .Add "TabFixedHeight"
                .Add "TabFixedWidth"
                .Add "TabOrientation"
                .Add "Value"

            Case "OptionButton"
                .Add "Accelerator"
                .Add "Alignment"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "Caption"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "GroupName"
                .Add "Locked"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PicturePosition"
                .Add "SpecialEffect"
                .Add "TextAlign"
                .Add "TripleState"
                .Add "Value"
                .Add "WordWrap"

            Case "Page"
                .Add "Accelerator"
                '.Add "CanPaste"
                .Add "CanRedo"
                .Add "CanUndo"
                .Add "Caption"
                .Add "Controls"
                .Add "ControlTipText"
                .Add "Cycle"
                .Add "Enabled"
                .Add "Index"
                .Add "InsideHeight"
                .Add "InsideWidth"
                .Add "KeepScrollBarsVisible"
                .Add "Name"
                .Add "Picture"
                .Add "PictureAlignment"
                .Add "PictureSizeMode"
                .Add "PictureTiling"
                .Add "ScrollBars"
                .Add "ScrollHeight"
                .Add "ScrollLeft"
                .Add "ScrollTop"
                .Add "ScrollWidth"
                .Add "Tag"
                .Add "TransitionEffect"
                .Add "TransitionPeriod"
                .Add "VerticalScrollBarSide"
                .Add "Visible"
                .Add "Zoom"

            Case "ScrollBar"
                .Add "BackColor"
                .Add "Delay"
                .Add "Enabled"
                .Add "ForeColor"
                .Add "LargeChange"
                .Add "Max"
                .Add "Min"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Orientation"
                .Add "ProportionalThumb"
                .Add "SmallChange"
                .Add "Value"

            Case "SpinButton"
                .Add "BackColor"
                .Add "Delay"
                .Add "Enabled"
                .Add "ForeColor"
                .Add "Max"
                .Add "Min"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Orientation"
                .Add "SmallChange"
                .Add "Value"

            Case "Tab"
                .Add "Accelerator"
                .Add "Caption"
                .Add "ControlTipText"
                .Add "Enabled"
                .Add "Index"
                .Add "Name"
                .Add "Tag"
                .Add "Visible"

            Case "TabStrip"
                .Add "BackColor"
                .Add "ClientHeight"
                .Add "ClientLeft"
                .Add "ClientTop"
                .Add "ClientWidth"
                .Add "Enabled"
                .Add "Font"
                .Add "ForeColor"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "MultiRow"
                .Add "Style"
                .Add "TabFixedHeight"
                .Add "TabFixedWidth"
                .Add "TabOrientation"
                .Add "Tabs"
                .Add "Value"

            Case "TextBox"
                .Add "AutoSize"
                .Add "AutoTab"
                .Add "AutoWordSelect"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "BorderColor"
                .Add "BorderStyle"
                '.Add "CanPaste"
                .Add "CurLine"
                .Add "DragBehavior"
                .Add "Enabled"
                .Add "EnterFieldBehavior"
                .Add "EnterKeyBehavior"
                .Add "Font"
                .Add "ForeColor"
                .Add "HideSelection"
                .Add "IMEMode"
                .Add "IntegralHeight"
                .Add "Locked"
                .Add "MaxLength"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "MultiLine"
                .Add "PasswordChar"
                .Add "ScrollBars"
                .Add "SelectionMargin"
                .Add "SpecialEffect"
                .Add "TabKeyBehavior"
                .Add "Text"
                .Add "TextAlign"
                .Add "Value"
                .Add "WordWrap"

            Case "ToggleButton"
                .Add "Accelerator"
                .Add "Alignment"
                .Add "AutoSize"
                .Add "BackColor"
                .Add "BackStyle"
                .Add "Caption"
                .Add "Enabled"
                .Add "ForeColor"
                .Add "GroupName"
                .Add "Locked"
                .Add "MouseIcon"
                .Add "MousePointer"
                .Add "Picture"
                .Add "PicturePosition"
                .Add "SpecialEffect"
                .Add "TextAlign"
                .Add "TripleState"
                .Add "Value"
                .Add "WordWrap"

            Case Else
                Debug.Print "Warning: Unknown ActiveX Control Type Name : " & strType

        End Select
    End With

End Function
