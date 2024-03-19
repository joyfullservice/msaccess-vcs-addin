Attribute VB_Name = "modRepair"
'---------------------------------------------------------------------------------------
' Module    : modRepair
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : This module is for functions used to repair conditions in the host
'           : database that may impair the functionality of the add-in.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : RepairColorDefinitionBlocks
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Go through all the form objects and set the color properties so that
'           : the theme index lines are correctly stored in the source files.
'           : See issue #183 for more details.
'---------------------------------------------------------------------------------------
'
Public Function RepairColorDefinitionBlocks()

    Dim obj As AccessObject
    Dim frm As Form
    Dim ctl As Control
    Dim sec As Section
    Dim intSec As Integer
    ' Loop through all forms
    For Each obj In CurrentProject.AllForms
        ' Open in design view so we can make changes.
        DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
        Set frm = Forms(obj.Name)
        ' Form properties
        SetColorProperties frm.Properties
        ' Control properties
        For Each ctl In frm.Controls
            SetColorProperties ctl.Properties
        Next ctl
        ' Section properties (header, detail, footer, etc...)
        For intSec = acDetail To 20 ' Max sections?
            On Error Resume Next
            Set sec = frm.Section(intSec)
            If Err Then
                ' Invalid section
                Err.Clear
            Else
                SetColorProperties sec.Properties
            End If
        Next intSec
        ' Save and close form
        DoCmd.Close acForm, obj.Name, acSaveYes
    Next obj

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetColorProperties
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapplies the existing color properties to update the internal color
'           : definitions.
'---------------------------------------------------------------------------------------
'
Private Sub SetColorProperties(prpCollection As Properties)

    Dim prp As Property
    Dim dItems As Dictionary
    Dim dProp As Dictionary
    Dim strBase As String
    Dim varKey As Variant
    Dim strProp As String
    Dim lngGradient As Long

    Set dItems = New Dictionary
    lngGradient = -1

    ' Loop through properties, collecting the color-related properties
    For Each prp In prpCollection
        With prp
            If InStr(1, .Name, "Color") > 0 _
                Or EndsWith(.Name, "Shade") _
                Or EndsWith(.Name, "Tint") Then
                ' Save this property value
                ' Build base name of property
                strBase = MultiReplace(.Name, _
                    "ThemeColorIndex", vbNullString, _
                    "Color", vbNullString, _
                    "Shade", vbNullString, _
                    "Tint", vbNullString)
                ' Save in dictionary using base name as the key
                '     Fore
                '       |---- ForeColor = 12345
                '       |---- ForeThemeColorIndex = 3
                '       |---- ForeShade = 4
                '       |---- ForeTint = 100
                If Not dItems.Exists(strBase) Then
                    Set dProp = New Dictionary
                    dItems.Add strBase, dProp
                End If
                dItems(strBase)(.Name) = .Value

            ElseIf .Name = "Gradient" Then
                ' Save gradient value
                lngGradient = .Value
            End If
        End With
    Next prp
    ' Now, with all the properties collected, we can check
    ' for the presence of the required items to represent the color
    For Each varKey In dItems.Keys
        Set dProp = dItems(varKey)
        strProp = varKey & "ThemeColorIndex"
        If dProp.Exists(strProp) Then
            ' Has index. Check value
            If dProp(strProp) = -1 Then
                ' Using absolute color, not theme
                ReApplyValue prpCollection, dProp, varKey & "Color"
            Else
                ' Using theme color
                ReApplyValue prpCollection, dProp, varKey & "ThemeColorIndex"
                ReApplyValue prpCollection, dProp, varKey & "Shade"
                ReApplyValue prpCollection, dProp, varKey & "Tint"
            End If
        Else
            ' No theme index. Use color value
            ReApplyValue prpCollection, dProp, varKey & "Color"
        End If
    Next varKey

    ' Restore any gradient property
    If lngGradient >= 0 Then prpCollection("Gradient") = lngGradient

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReApplyValue
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapply the property value to ensure it has been saved.
'---------------------------------------------------------------------------------------
'
Private Sub ReApplyValue(colProps As Properties, dProps As Dictionary, strName As String)
    If dProps.Exists(strName) Then
        colProps(strName).Value = colProps(strName).Value
    End If
End Sub
