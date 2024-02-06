VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' API calls for getting screen cordinates in relation to controls
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getsystemmetrics
Private Enum SystemMetricsConstants
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXSIZEFRAME = SM_CXFRAME
    SM_CYSIZEFRAME = SM_CYFRAME
End Enum

' Valid control types for popup menu controls
Public Enum eControlType
    ectControlButton = msoControlButton
    ectControlEdit = msoControlEdit
    ectControlDropdown = msoControlDropdown
    ectControlComboBox = msoControlComboBox
    ectControlPopup = msoControlPopup
End Enum

' A reference to the menu bar we are working with
Private m_barMenu As CommandBar
Private dFaceIds As Dictionary
Private lngTotal As Long
Private lngNames As Long


'---------------------------------------------------------------------------------------
' Procedure : SysMetric
' Author    : Adapted from work by Paul J. Champion in 2004
' Date      : 1/22/2024
' Purpose   : Return a system metric in pixels or twips.
'---------------------------------------------------------------------------------------
'
Private Function SysMetric(SystemMetricRequired As SystemMetricsConstants, Optional ConvertToTwips As Boolean = True) As Variant
    If ConvertToTwips Then
        Select Case SystemMetricRequired
            'Case SM_CYSCREEN, SM_CYHSCROLL, SM_CYCAPTION, _
                    SM_CYBORDER, SM_CXDLGFRAME, SM_CYDLGFRAME, _
                    SM_CYVTHUMB, SM_CYICON, SM_CYCURSOR, _
                    SM_CYMENU, SM_CYFULLSCREEN, SM_CYKANJIWINDOW, _
                    SM_CYVSCROLL, SM_CYMIN, SM_CYSIZE, _
                    SM_CYFRAME, SM_CYMINTRACK, SM_CYDOUBLECLK, _
                    SM_CYICONSPACING, SM_CYSIZEFRAME, SM_CYFIXEDFRAME
            ' The ones we actually use
            Case SM_CYCAPTION, SM_CYBORDER, SM_CXDLGFRAME, SM_CYDLGFRAME, _
                    SM_CYFRAME, SM_CYSIZEFRAME
                SysMetric = GetTwipsFromPixels(GetSystemMetrics(SystemMetricRequired), False)
            Case Else: SysMetric = GetTwipsFromPixels(GetSystemMetrics(SystemMetricRequired), True)
        End Select
    Else
        SysMetric = GetSystemMetrics(SystemMetricRequired)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPixelsFromTwips
' Author    : Adam Waller
' Date      : 1/19/2024
' Purpose   : Return pixels (i.e. cursor position, menu position) from twips.
'           : Caches the ratio value after first call.
'---------------------------------------------------------------------------------------
'
Private Function GetPixelsFromTwips(lngTwips As Long, blnX As Boolean) As Long

    Static lngPixels(True To False) As Long
    Dim hDC As Long

    If lngPixels(blnX) = 0 Then
        hDC = GetDC(Application.hWndAccessApp)
        lngPixels(blnX) = 1440 / GetDeviceCaps(hDC, IIf(blnX, LOGPIXELSX, LOGPIXELSY))
        ReleaseDC Application.hWndAccessApp, hDC
    End If
    GetPixelsFromTwips = lngTwips / lngPixels(blnX)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTwipsFromPixels
' Author    : Adam Waller
' Date      : 1/22/2024
' Purpose   : Return a twips size from a pixel size.
'---------------------------------------------------------------------------------------
'
Private Function GetTwipsFromPixels(lngPixels As Long, blnX As Boolean) As Long

    Static lngTwips(True To False) As Long
    Dim hDC As Long

    If lngTwips(blnX) = 0 Then
        hDC = GetDC(Application.hWndAccessApp)
        lngTwips(blnX) = 1440 / GetDeviceCaps(hDC, IIf(blnX, LOGPIXELSX, LOGPIXELSY))
        ReleaseDC Application.hWndAccessApp, hDC
    End If
    GetTwipsFromPixels = lngPixels * lngTwips(blnX)

End Function


'---------------------------------------------------------------------------------------
' Procedure : PopupBelow
' Author    : Adam Waller
' Date      : 1/19/2024
' Purpose   : Displays a popup menu below a specified control. (From bottom left corner)
'---------------------------------------------------------------------------------------
'
Public Sub PopupBelow(ctl As Access.Control, strMenu As String)

    Dim objParent As Object
    Dim lngFocus As Long
    Dim lngOffsetX As Long
    Dim lngOffsetY As Long
    Dim hwdMdi As LongPtr
    Dim recMdi As RECT

    ' Work our way back up to the parent form, calculating the offset as we go.
    Set objParent = ctl
    Do
        lngOffsetX = lngOffsetX + objParent.Left
        lngOffsetY = lngOffsetY + objParent.Top
        Set objParent = objParent.Parent
        If objParent Is Screen.ActiveForm Then
            ' We have a reference to the parent form
            Exit Do
        End If
    Loop

    ' Add height of control to set popup just below
    lngOffsetY = lngOffsetY + ctl.Height

    ' Add any title bar height for popup forms
    If objParent.PopUp Then
        Select Case objParent.BorderStyle
            Case 0      ' None
            Case 1, 3   ' Thin, Dialog
                lngOffsetX = lngOffsetX + SysMetric(SM_CXDLGFRAME)
                lngOffsetY = lngOffsetY + SysMetric(SM_CYCAPTION) + SysMetric(SM_CYDLGFRAME)
            Case 2      ' Sizeable
                lngOffsetX = lngOffsetX + SysMetric(SM_CXSIZEFRAME) + SysMetric(SM_CXBORDER)
                lngOffsetY = lngOffsetY + SysMetric(SM_CYCAPTION) + SysMetric(SM_CYSIZEFRAME) + SysMetric(SM_CYBORDER)
        End Select
    Else
        ' Subtract border size from x,y to reflect the actual edge metric
        lngOffsetX = lngOffsetX - SysMetric(SM_CXBORDER) * 2
        lngOffsetY = lngOffsetY - SysMetric(SM_CYBORDER) * 2
    End If

    ' Get the top left coordinates of the Access application multi document interface area,
    ' Or the open form, in the case of a popup form.
    hwdMdi = GetFocus
    GetWindowRect hwdMdi, recMdi

    ' Show the popup at the specified location (Left aligned)
    CommandBars(strMenu).ShowPopup _
        GetPixelsFromTwips(lngOffsetX, True) + recMdi.Left + CommandBars(strMenu).Width, _
        GetPixelsFromTwips(lngOffsetY, False) + recMdi.Top

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddCustomControl
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Add a custom control to the popup menu
'---------------------------------------------------------------------------------------
'
Public Function AddCustomControl(intType As eControlType, strCaption As String, Optional intFaceID As Integer, _
    Optional strMsoImageId As String, Optional barMenu As CommandBar) As CommandBarControl

    ' Make sure we have a menu bar to work with
    If Not BarDefined(barMenu) Then Exit Function

    ' Add control
    Set AddCustomControl = m_barMenu.Controls.Add(intType, , , , False)
    With AddCustomControl
        .Caption = strCaption
        If intFaceID <> 0 Then .FaceId = intFaceID
        If Len(strMsoImageId) Then .Picture = CommandBars.GetImageMso(strMsoImageId, 16, 16)
        CheckControlActions m_barMenu.Controls, m_barMenu.Name
    End With

    ' Popup menu to show new control
    m_barMenu.ShowPopup

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddExistingControl
' Author    : Adam Waller
' Date      : 1/22/2024
' Purpose   : Add an existing (built-in) control to the menu.
'---------------------------------------------------------------------------------------
'
Public Function AddExistingControl(lngId As Long, Optional barMenu As CommandBar)

    'Dim ctlExisting As Object

    ' Make sure we have a menu bar to work with
    If Not BarDefined(barMenu) Then Exit Function

    ' Add control
    Set AddExistingControl = m_barMenu.Controls.Add(, lngId, , , False)
'    With AddExistingControl
'        .Caption = strCaption
'        If intFaceID <> 0 Then .FaceId = intFaceID
'        If Len(strMsoImageId) Then .Picture = CommandBars.GetImageMso(strMsoImageId, 16, 16)
'        CheckControlActions m_barMenu.Controls, m_barMenu.Name
'    End With

    ' Popup menu to show new control
    m_barMenu.ShowPopup

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindID
' Author    : Dale Fye
' Date      : 2014-02-08
' Purpose   : Find control IDs with matching captions. (Existing items)
'---------------------------------------------------------------------------------------
'
Public Sub FindID(CaptionContains As String)
    Dim cbr As CommandBar
    Dim ctrl As CommandBarControl

    For Each cbr In Application.CommandBars
        For Each ctrl In cbr.Controls

            'Because many of the controls contain keyboard shortcuts, they
            'have ampersands embedded within the caption that must be
            'removed before doing the comparison to the search string.
            If InStr(Replace(ctrl.Caption, "&", ""), CaptionContains) > 0 Then
                Debug.Print cbr.Name, ctrl.Caption, ctrl.Id
            End If
        Next
    Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LastItem
' Author    : Adam Waller
' Date      : 1/22/2024
' Purpose   : Get the last control added to the command bar.
'---------------------------------------------------------------------------------------
'
Public Function LastItem(Optional barMenu As CommandBar) As CommandBarControl

    ' Make sure we have a menu bar to work with
    If Not BarDefined(barMenu) Then Exit Function

    ' Return last control
    Set LastItem = m_barMenu.Controls(m_barMenu.Controls.Count)

End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyOnAction
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Update the command bar actions to reflect the command bar name, and the
'           : caption of the control. (Without the ampersand control characters)
'---------------------------------------------------------------------------------------
'
Public Sub VerifyOnAction(Optional cbrMenu As CommandBar)

    Dim strTemp As String
    Dim cbr As CommandBar
    Dim ctl As CommandBarControl

    ' Make sure we have a menu bar to work with
    If Not BarDefined(cbrMenu) Then Exit Sub

    ' Loop through custom command bars
    For Each cbr In CommandBars
        If Not cbr.BuiltIn Then
            ' Recursively process controls
            CheckControlActions m_barMenu.Controls, cbr.Name
        End If
    Next cbr

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckControlActions
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Recursive function to check command bar control actions
'---------------------------------------------------------------------------------------
'
Private Sub CheckControlActions(ctls As CommandBarControls, strPath As String)

    Dim strName As String
    Dim strSource As String
    Dim ctl As CommandBarControl
    Dim strTemp As String
    Dim strOnAction As String

    ' Loop through controls
    For Each ctl In ctls
        If TypeOf ctl Is CommandBarPopup Then
            ' Recursively call this function on the nested controls.
            CheckControlActions ctl.Controls, strPath & "." & ctl.Name
        Else
            ' Use caption or index to make name
            strName = NZ2(ctl.Caption, ctl.Index)

            ' Build source path
            strSource = strPath & "." & strName

            ' Add a Unicode placeholder for double ampersand, and remove
            ' control ampersand characters to make captions more readable.
            strTemp = ChrW(55357) & ChrW(56999)
            strSource = MultiReplace(strSource, _
                "&&", strTemp, _
                "&", vbNullString, _
                strTemp, "&", _
                """", vbNullString)

            ' Build out the desired OnAction value
            strOnAction = "=MenuHandler(""" & strSource & """)"
            If ctl.OnAction <> strOnAction Then ctl.OnAction = strOnAction
        End If
    Next ctl

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddMenu
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Add a new command bar (popup menu)
'---------------------------------------------------------------------------------------
'
Public Function AddMenu(strName As String) As CommandBar
    Set m_barMenu = CommandBars.Add(strName, msoBarPopup)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ListCustomMenus
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : List out any custom command bars in the current database.
'---------------------------------------------------------------------------------------
'
Public Sub ListCustomMenus()
    Dim bar As CommandBar
    Debug.Print "Custom Command Bars:"
    Debug.Print "=========================="
    For Each bar In CommandBars
        If Not bar.BuiltIn Then Debug.Print "CommandBars(""" & bar.Name & """)"
    Next bar
    Debug.Print "=========================="
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetPictureFromMsoImage
' Author    : Adam Waller
' Date      : 1/17/2024
' Purpose   : Apply a MSO image to a command bar control (popup menu item)
'---------------------------------------------------------------------------------------
'
Public Sub SetPictureFromMsoImage(ctl As CommandBarControl, strImageCaseSensitive As String)
    ctl.Picture = CommandBars.GetImageMso(strImageCaseSensitive, 16, 16)
    ctl.Parent.ShowPopup
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BarDefined
' Author    : Adam Waller
' Date      : 1/16/2024
' Purpose   : Returns true if a command bar object has been specified, either in this
'           : call, or a previous one.
'---------------------------------------------------------------------------------------
'
Private Function BarDefined(cbrMenu As CommandBar) As Boolean
    If Not cbrMenu Is Nothing Then Set m_barMenu = cbrMenu
    If m_barMenu Is Nothing Then
        MsgBox "You must specify a menu bar object on first call", vbExclamation
    Else
        BarDefined = True
    End If
End Function
