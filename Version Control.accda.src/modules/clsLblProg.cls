VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLblProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLblProg
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Display a progress bar using three labels on a form
'---------------------------------------------------------------------------------------
Option Explicit
Option Compare Database


' Set the application name here to utilize the appropriate
' object types in early binding.
#Const APPLICATION_NAME = "Microsoft Access"

#If APPLICATION_NAME = "Microsoft Access" Then
    ' Use Access specific controls/sizing
    Private Const sngOffset As Single = 15
    Private mlblBack As Access.Label    ' existing label for back
    Private mlblFront As Access.Label   ' label created for front
    Private mlblCaption As Access.Label ' progress bar caption
#Else
    ' Generic VBA objects
    Private Const sngOffset As Single = 1.5
    Private mlblBack As Label           ' existing label for back
    Private mlblFront As Label          ' label created for front
    Private mlblCaption As Label        ' progress bar caption
#End If


' Public properties
Public Max As Double                    ' Maximum value of progress bar at 100%
Public Smooth As Boolean                ' Set to true for smooth updates < 1%

' Private properties
Private mdblVal As Double               ' Current value of progress bar
Private mdblFullWidth As Double         ' Width of front label at 100%
Private mdblIncSize As Double           ' Icrement size
Private mblnHideCap As Boolean          ' Show percent complete caption
Private mobjParent As Object            ' Parent object of back label
Private mdteLastUpdate As Date          ' Time last updated
Private mblnNotSmooth As Boolean        ' Display smoothly by doevents after every update.


'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Initialize the progress bar before using it
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(BackLabel As Label, Optional FrontLabel As Label, Optional CaptionLabel As Label)

    Dim objParent As Object ' could be a form or tab control
    Dim frm As Form
    
    Set mobjParent = BackLabel.Parent
    Set mlblBack = BackLabel
    
    #If APPLICATION_NAME = "Microsoft Access" Then
        ' Use existing controls
        Set mlblFront = FrontLabel
        Set mlblCaption = CaptionLabel
    #Else
        ' Create front controls dynamically
        Set mlblFront = mobjParent.Controls.Add("forms.label.1", "", False)
        Set mlblFront = CreateControl(GetParentFormName(BackLabel), acLabel, , BackLabel.Parent.Name)
        Set mlblCaption = mobjParent.Controls.Add("forms.label.1", "", False)
        Set mlblCaption = CreateControl(GetParentFormName(BackLabel), acLabel, , BackLabel.Parent)
        ' Refresh display of parent form
        mobjParent.Repaint
    #End If
        
    ' Set properties for back label
    With mlblBack
        .Visible = True
        .SpecialEffect = 2  ' sunken. Seems to lose when not visible.
    End With
    
    ' Set properties for front label
    With mlblFront
        mdblFullWidth = mlblBack.Width - (sngOffset * 2)
        .Left = mlblBack.Left + sngOffset
        .Top = mlblBack.Top + sngOffset
        .Width = 0
        .Height = mlblBack.Height - (sngOffset * 2)
        .Caption = vbNullString
        .BackColor = 8388608
        .BackStyle = 1
        .Visible = True
    End With
    
    ' set properties for caption label
    With mlblCaption
        .Left = mlblBack.Left + 2
        .Top = mlblBack.Top + 2
        .Width = mlblBack.Width - 4
        .Height = mlblBack.Height - 4
        .TextAlign = 2 'fmTextAlignCenter
        .BackStyle = 0 'fmBackStyleTransparent
        .Caption = "0%"
        .Visible = Not Me.HideCaption
        .ForeColor = 16777215   ' white
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Remove temporary controls, if applicable
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()

    ' Ignore any errors if object references are no longer valid
    On Error Resume Next
    
    #If APPLICATION_NAME = "Microsoft Access" Then
        ' Hide front controls
        mlblFront.Visible = False
        mlblCaption.Visible = False
    #Else
        ' Remove front controls
        mobjParent.Controls.Remove (mlblFront.Name)
        mobjParent.Controls.Remove (mlblCaption.Name)
    #End If
    Err.Clear ' Clear the Errors if thrown; they're not useful
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Value
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Get the current value of the progress bar.
'---------------------------------------------------------------------------------------
'
Public Property Get Value() As Double
    Value = mdblVal
End Property


'---------------------------------------------------------------------------------------
' Procedure : Value
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Set the value of the progress bar, checking visibility of controls, and
'           : updating screen as needed.
'---------------------------------------------------------------------------------------
'
Public Property Let Value(ByVal dblVal As Double)

    Dim dblChange As Double
    
    ' Check visibility
    If Me.Max = 0 Or dblVal = 0 Then
        mlblFront.Visible = False
        mlblCaption.Visible = False
    Else
        ' Ensure controls are visible
        If Not mlblBack.Visible Then mlblBack.Visible = True
        If Not mlblFront.Visible Then mlblFront.Visible = True
        If Not mblnHideCap And Not mlblCaption.Visible Then mlblCaption.Visible = True
    End If
    
    ' Don't allow value to exceed maximum value
    If dblVal > Me.Max Then dblVal = Me.Max
    
    ' Measure change
    dblChange = Abs(dblVal - mdblVal)
    
    ' Set value and update display if needed.
    If dblChange > 0 And Me.Max > 0 Then
        ' See if we need to update the display
        ' (Normally updated every 1% or each increment if Smooth = True
        If Me.Smooth Or ((dblChange / Me.Max) > 0.01) Then
            mdblVal = dblVal
            Update
        End If
    End If

End Property


'---------------------------------------------------------------------------------------
' Procedure : IncrementSize
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Can use a custom increment size if other than 1.
'---------------------------------------------------------------------------------------
'
Public Property Get IncrementSize() As Double
    IncrementSize = mdblIncSize
End Property
Public Property Let IncrementSize(ByVal dblSize As Double)
    mdblIncSize = dblSize
End Property


'---------------------------------------------------------------------------------------
' Procedure : HideCaption
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Optionally hide the caption display
'---------------------------------------------------------------------------------------
'
Public Property Get HideCaption() As Boolean
    HideCaption = mblnHideCap
End Property

Public Property Let HideCaption(ByVal blnHide As Boolean)
    mblnHideCap = blnHide
End Property


'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Update the display
'---------------------------------------------------------------------------------------
'
Private Sub Update()

    Dim intPercent As Integer
    Dim dblWidth As Double
    
    ' Set size and caption
    intPercent = mdblVal * (100 / Me.Max)
    dblWidth = mdblVal * (mdblFullWidth / Me.Max)
    If dblWidth <= mdblFullWidth Then
        mlblFront.Width = dblWidth
    Else
        mlblFront.Width = mdblFullWidth
    End If
    
    mlblCaption.Caption = intPercent & "%"
    
    ' Use white or black, depending on progress
    If Me.Value > (Me.Max / 2) Then
        mlblCaption.ForeColor = 16777215   ' white
    Else
        mlblCaption.ForeColor = 0  ' black
    End If

    ' Use DoEvents to repaint display
    If mblnNotSmooth Then
        If mdteLastUpdate <> Now Then
            ' Update every second
            DoEvents
            mdteLastUpdate = Now
        End If
    Else
        DoEvents
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Increment
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Increment the progress bar, defaulting to custom increment size, then
'           : passed value, then 1 as default. (Normally used without any argument.)
'---------------------------------------------------------------------------------------
'
Public Sub Increment(Optional dblAmount As Double = 1)

    Dim dblAdd As Double
    Dim dblNew As Double
    
    ' Calculate how much to add
    If dblAmount = 1 And Me.IncrementSize > 0 Then dblAdd = Me.IncrementSize
    If dblAdd = 0 Then dblAdd = dblAmount
    
    ' Check boundaries
    dblNew = Me.Value + dblAdd
    If dblNew > Me.Max Then dblNew = Me.Max
    If dblNew < 0 Then dblNew = 0

    ' Set new value (and any needed display updates)
    Me.Value = dblNew

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Remove or hide front labels. (Will need to initialize again after this.)
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Class_Terminate
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Reset
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Reset back to zero.
'---------------------------------------------------------------------------------------
'
Public Sub Reset()
    Me.Value = 0
    Update
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Hide
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Hide the visible elements of the progress bar (but retain values)
'---------------------------------------------------------------------------------------
'
Public Sub Hide()
    On Error Resume Next
    mlblCaption.Visible = False
    mlblFront.Visible = False
    mlblBack.Visible = False
    CatchAny eelNoError, vbNullString, vbNullString, False, True
End Sub


