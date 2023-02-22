Attribute VB_Name = "modUIAutomation"
'---------------------------------------------------------------------------------------
' Module    : modUIAutomation
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Use UI Automation to access elements not available through the VBA
'           : object model.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ModuleName = "modUIAutomation"


'---------------------------------------------------------------------------------------
' Procedure : GetSelectedNavPaneObject
' Author    : Adam Waller
' Date      : 2/14/2023
' Purpose   : Return the item currently selected in the navigation pane.
'           : Tip: The Accessibility Insights for Windows utility is a great way to
'           : explore the UI elements in an application.
'---------------------------------------------------------------------------------------
'
Public Function GetSelectedNavPaneObject() As AccessObject

    Dim oClient As UIAutomationClient.CUIAutomation
    Dim oSelected As UIAutomationClient.IUIAutomationElement
    Dim oElement As UIAutomationClient.IUIAutomationElement
    Dim oCondition As UIAutomationClient.IUIAutomationCondition

    ' Create new automation client
    Set oClient = New UIAutomationClient.CUIAutomation

    ' Get currently selected element
    Set oSelected = oClient.GetFocusedElement

    ' Drill down to selected item name
    If oSelected.CurrentControlType = UIA_PaneControlTypeId Then

        ' Build condition for navigation pane item with keyboard focus
        Set oCondition = oClient.CreateAndCondition( _
            oClient.CreatePropertyCondition(UIA_HasKeyboardFocusPropertyId, True), _
            oClient.CreatePropertyCondition(UIA_ClassNamePropertyId, "NetUINavPaneItem"))

        ' Attempt to find the selected item (looking for keyboard focus)
        Set oElement = oSelected.FindFirst(TreeScope_Descendants, oCondition)

        ' If an item was found, the continue to drill down to get the name and type
        If Not oElement Is Nothing Then
            Set GetSelectedNavPaneObject = GetUnderlyingDbObjectFromButton(oClient, oElement)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUnderlyingDbObjectFromButton
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Return the database object from the UI button
'---------------------------------------------------------------------------------------
'
Private Function GetUnderlyingDbObjectFromButton(oClient As CUIAutomation, oElement As IUIAutomationElement) As AccessObject

    Dim strName As String
    Dim strImage As String
    Dim objItem As AccessObject
    
    ' Read name from button name
    strName = oElement.CurrentName
    
    ' Get the object type from the image name
    strImage = GetImageName(oClient, oElement)
    
    ' Just in case something doesn't work right...
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    
    ' There are multiple icons for some objects
    If strImage Like "Table*" Then
        Set objItem = CurrentData.AllTables(strName)
    ElseIf strImage Like "*Query" Then
        Set objItem = CurrentData.AllQueries(strName)
    Else
        ' These objects have a single representative icon
        Select Case strImage
            Case "Form"
                Set objItem = CurrentProject.AllForms(strName)
            Case "Report"
                Set objItem = CurrentProject.AllReports(strName)
            Case "Macro"
                Set objItem = CurrentProject.AllMacros(strName)
            Case "Class Module", "Module"
                Set objItem = CurrentProject.AllModules(strName)
                
            ' Some ADP specific project items
            Case "Function"
                Set objItem = CurrentData.AllFunctions(strName)
            Case "StoredProcedure"
                Set objItem = CurrentData.AllStoredProcedures(strName)
            Case "Diagram"
                Set objItem = CurrentData.AllDatabaseDiagrams(strName)
                
            ' Check for no image name returned
            Case vbNullString
                ' No image name found
            
            ' Unrecognized name
            Case Else
                Debug.Print "Navigation pane item image name not recognized: " _
                    & strImage & " (for " & strName & ")"
        End Select
    End If
    
    ' Report any errors retrieving underlying object
    CatchAny eelError, "Error getting underlying object for " & strName, _
        ModuleName & ".GetUnderlyingDbObjectFromButton"
    
    ' Return database object if we found a matching one
    If Not objItem Is Nothing Then Set GetUnderlyingDbObjectFromButton = objItem
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetImageName
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Get the image name from the icon on the button
'---------------------------------------------------------------------------------------
'
Private Function GetImageName(oClient As CUIAutomation, oElement As IUIAutomationElement) As String

    Dim oImage As UIAutomationClient.IUIAutomationElement
    Dim oCondition As UIAutomationClient.IUIAutomationCondition

    ' Build condition for navigation pane item with keyboard focus
    Set oCondition = oClient.CreateAndCondition( _
        oClient.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ImageControlTypeId), _
        oClient.CreatePropertyCondition(UIA_ClassNamePropertyId, "NetUIImage"))

    ' Attempt to find the selected item (looking for keyboard focus)
    Set oImage = oElement.FindFirst(TreeScope_Descendants, oCondition)
    
    ' Return name of image, if found
    If Not oImage Is Nothing Then GetImageName = oImage.CurrentName
    
End Function
