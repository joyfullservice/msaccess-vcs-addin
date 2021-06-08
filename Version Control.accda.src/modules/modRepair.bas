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
    For Each prp In prpCollection
        With prp
            If InStr(1, .Name, "Color", vbTextCompare) > 0 Then
                .Value = .Value
            End If
        End With
    Next prp
End Sub

