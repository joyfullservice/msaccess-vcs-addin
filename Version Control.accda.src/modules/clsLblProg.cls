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
' Date      : 4/21/2022
' Purpose   : Display a progress bar using three labels on a form
'---------------------------------------------------------------------------------------
Option Explicit
Option Compare Database


' Set the application name here to utilize the appropriate
' object types in early binding.
#Const APPLICATION_NAME = "Microsoft Access"

Const CompletionFormat As String = "0.00" ' Defines the number of digits to display.

' Public properties
Private m_Max As Double                 ' Maximum value of progress bar at 100%
Public Smooth As Boolean                ' Set to true for smooth updates < 1%

' Private properties
Private Type udtProg

    ' Application specific properties
    #If APPLICATION_NAME = "Microsoft Access" Then
        ' Use Access specific controls/sizing
        lblBack As Access.Label         ' Existing label for back
        lblFront As Access.Label        ' Label for moving bar
        lblCaption As Access.Label      ' Progress bar caption
    #Else
        ' Generic VBA objects
        lblBack As Label                ' Existing label for back
        lblFront As Label               ' Label created for front
        lblCaption As Label             ' Label created for caption
    #End If

    ' Shared private properties
    dblValue As Double                  ' Current value of progress bar
    dblDisplayed As Double              ' Displayed value
    dblFullWidth As Double              ' Width of front label at 100%
    sngOffset As Single                 ' Offset for front label
    dblIncrementSize As Double          ' Increment size
    blnHideCaption As Boolean           ' Show/hide percent complete caption
    objParent As Object                 ' Parent object of back label
    dteStartTime As Date                ' Time of first value change
    blnShowTimeRemain As Boolean        ' Show time on progress bar in addition to Percentage
    blnShowTimeComplete As Boolean      ' Show time of completion (clock time) in addition to percentage.
                                        ' Only Time remaining OR clock time can be shown; if you turn on one the other will turn off.
End Type
Private this As udtProg


'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Initialize the progress bar before using it
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(ByRef BackLabel As Label, ByRef FrontLabel As Label, ByRef CaptionLabel As Label)
    
    Set this.objParent = BackLabel.Parent
    Set this.lblBack = BackLabel
    
    #If APPLICATION_NAME = "Microsoft Access" Then
        this.sngOffset = 15
        ' Use existing controls
        Set this.lblFront = FrontLabel
        Set this.lblCaption = CaptionLabel
    #Else
        this.sngOffset = 1.5
        ' Create front controls dynamically
        Set this.lblFront = mobjParent.Controls.Add("forms.label.1", "", False)
        Set this.lblFront = CreateControl(GetParentFormName(BackLabel), acLabel, , BackLabel.Parent.Name)
        Set this.lblCaption = mobjParent.Controls.Add("forms.label.1", "", False)
        Set this.lblCaption = CreateControl(GetParentFormName(BackLabel), acLabel, , BackLabel.Parent)
        ' Refresh display of parent form
        this.objParent.Repaint
    #End If
        
    ' Set properties for back label
    With this.lblBack
        .Visible = True
        .SpecialEffect = 2  ' sunken. Seems to lose when not visible.
    End With
    
    ' Set properties for front label
    With this.lblFront
        this.dblFullWidth = this.lblBack.Width - (this.sngOffset * 2)
        .Left = this.lblBack.Left + this.sngOffset
        .Top = this.lblBack.Top + this.sngOffset
        .Width = 0
        .Height = this.lblBack.Height - (this.sngOffset * 2)
        .Caption = vbNullString
        .BackColor = 8388608
        .BackStyle = 1
        .Visible = True
    End With
    
    ' set properties for caption label
    With this.lblCaption
        .Left = this.lblBack.Left + 2
        .Top = this.lblBack.Top + 2
        .Width = this.lblBack.Width - 4
        .Height = this.lblBack.Height - 4
        .TextAlign = 2 'fmTextAlignCenter
        .TopMargin = this.sngOffset * 2
        .BackStyle = 0 'fmBackStyleTransparent
        .Caption = "0%"
        .Visible = Not this.blnHideCaption
        .ForeColor = 16777215   ' white
    End With
    Update

End Sub

Public Property Get HideCaption() As Boolean
    HideCaption = this.blnHideCaption
End Property
Public Property Let HideCaption(NewVal As Boolean)
    this.blnHideCaption = NewVal
End Property

Public Property Get DisplayTimeRemaining() As Boolean
    DisplayTimeRemaining = this.blnShowTimeRemain
End Property
Public Property Let DisplayTimeRemaining(NewVal As Boolean)
    this.blnShowTimeRemain = NewVal
    If NewVal Then this.blnShowTimeComplete = False
End Property
Public Property Get DisplayTimeComplete() As Boolean
    DisplayTimeComplete = this.blnShowTimeComplete
End Property
Public Property Let DisplayTimeComplete(NewVal As Boolean)
    this.blnShowTimeComplete = NewVal
    If NewVal Then this.blnShowTimeRemain = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Max Property
' Author    : hecon5
' Date      : 5/13/2022
' Purpose   : Set maximum value for value.
'---------------------------------------------------------------------------------------
'
Public Property Get Max() As Double
    Max = m_Max
End Property
Public Property Let Max(NewVal As Double)
    m_Max = NewVal
    Update True
End Property
'---------------------------------------------------------------------------------------
' Procedure : Value
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Get the current value of the progress bar.
'---------------------------------------------------------------------------------------
'
Public Property Get Value() As Double
    Value = this.dblValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Value
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Set the value of the progress bar, checking visibility of controls, and
'           : updating screen as needed.
'---------------------------------------------------------------------------------------
'
Public Property Let Value(ByVal dblValue As Double)

    Dim dblChange As Double
    
    ' Check visibility
    If Me.Max = 0 Or dblValue = 0 Then
        this.lblFront.Visible = False
        this.lblCaption.Visible = False
    Else
        ' Ensure controls are visible
        If Not this.lblBack.Visible Then this.lblBack.Visible = True
        If Not this.lblFront.Visible Then this.lblFront.Visible = True
        If Not this.blnHideCaption And Not this.lblCaption.Visible Then this.lblCaption.Visible = True
    End If
    
    ' Don't allow value to exceed maximum value
    If dblValue > Me.Max Then
        this.dblValue = Me.Max
'    ElseIf Me.Max = 0 Then
'        this.dblValue = 0
    Else
        ' Update value
        this.dblValue = dblValue
    End If
    
    ' Record start time on first change
    If this.dteStartTime = 0 Then this.dteStartTime = Now
    
    ' Measure change from currently displayed value
    dblChange = Abs(this.dblValue - this.dblDisplayed)
    
    ' Set value and update display if needed.
    If dblChange > 0 And Me.Max > 0 Then
        ' See if we need to update the display
        ' (Normally updated every 1% or each increment if Smooth = True
        If Me.Smooth Or ((dblChange / Me.Max) > 0.01) Then Update
    End If

End Property

Public Property Get IncrementSize() As Double
    If this.dblIncrementSize <= 0 Then this.dblIncrementSize = 1
    IncrementSize = this.dblIncrementSize
End Property
Public Property Let IncrementSize(NewVal As Double)
    this.dblIncrementSize = NewVal
End Property

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
    If dblAmount = 1 Then
        dblAdd = Me.IncrementSize
    Else
        dblAdd = dblAmount
    End If
    
    ' Check boundaries
    dblNew = Me.Value + dblAdd
    If dblNew > Me.Max Then dblNew = Me.Max
    If dblNew < 0 Then dblNew = 0

    ' Set new value (and any needed display updates)
    Me.Value = dblNew

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EstCompletionTime
' Author    : Adam Waller
' Date      : 4/21/2022
' Purpose   : Return an estimated completion time based on the start date and remaining
'           : segments. (Simple calculation assuming even increment intervals.)
'---------------------------------------------------------------------------------------
'
Public Property Get EstCompletionTime() As Date

    Dim dblSeconds As Double
    Dim dblRemaining As Double
    
    If this.dteStartTime = 0 Then
        EstCompletionTime = Now
    Else
        With this
            ' Get elapsed seconds
            dblSeconds = DateDiff("s", .dteStartTime, Now)
            ' Calculate remaining seconds
            If .dblValue > 0 Then dblRemaining = (dblSeconds / .dblValue) * (Me.Max - .dblValue)
            ' Convert to completion time
            EstCompletionTime = DateAdd("s", dblRemaining, Now)
        End With
    End If
    
End Property


'---------------------------------------------------------------------------------------
' Procedure : EstRemainingTime
' Author    : Adam Waller
' Date      : 4/21/2022
' Purpose   : Return a string value with the estimated hours/minutes/seconds remaining.
'           : Example return values: "73 seconds", "41 minutes", "80 hours"
'---------------------------------------------------------------------------------------
'
Public Property Get EstRemainingTime() As String

    Dim dblSeconds As Double
    
    dblSeconds = DateDiff("s", Now, Me.EstCompletionTime)
    If dblSeconds < 120 Then
        EstRemainingTime = Format(dblSeconds, CompletionFormat) & " seconds"
    ElseIf dblSeconds < (60 * 120) Then
        EstRemainingTime = Format(dblSeconds / 60, CompletionFormat) & " minutes"
    Else
        EstRemainingTime = Format(dblSeconds / 3600, CompletionFormat) & " hours"
    End If
    
End Property


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
    this.dteStartTime = 0
    Update True
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
    With this
        .lblCaption.Visible = False
        .lblFront.Visible = False
        .lblBack.Visible = False
    End With
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Update the display
'---------------------------------------------------------------------------------------
'
Private Sub Update(Optional ByRef ForceUpdate As Boolean = False)

    Dim intPercent As Integer
    Dim dblWidth As Double
    Static dteLastUpdate As Date

    ' Set size and caption
    With this
        If Me.Max = 0 Then
            .lblFront.Visible = False
            .lblCaption.Visible = False
            intPercent = 0
            dblWidth = 0
        Else
            intPercent = .dblValue * (100 / Me.Max)
            dblWidth = .dblValue * (.dblFullWidth / Me.Max)
        End If
        .lblFront.Width = dblWidth
        
        If .blnShowTimeComplete Then
            .lblCaption.Caption = intPercent & "% Est. Completion Time: " & Me.EstCompletionTime
        ElseIf .blnShowTimeRemain Then
            .lblCaption.Caption = intPercent & "% Est. Remaining Time: " & Me.EstRemainingTime
        Else
            .lblCaption.Caption = intPercent & "%"
        End If

        .dblDisplayed = .dblValue
    End With
    
    ' Use white or black, depending on progress level
    If Me.Value > (Me.Max / 2) Then
        ' Show white on blue background after 50%
        this.lblCaption.ForeColor = 16777215
    Else
        ' Show black on grey background
        this.lblCaption.ForeColor = 0
    End If

    ' Use DoEvents to repaint display
    If Me.Smooth Then
        ' Always update display
        DoEvents
    ElseIf (dteLastUpdate <> Now) Or ForceUpdate Then
        ' Update every second
        DoEvents
        dteLastUpdate = Now
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetParentFormName
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Return the name of the parent form
'---------------------------------------------------------------------------------------
'
Private Function GetParentFormName(ctlControl As Control) As String

    ' returns the name of the parent form
    Dim objParent As Object
    
    Set objParent = ctlControl
    
    Do While Not TypeOf objParent Is Form
       Set objParent = objParent.Parent
    Loop
    
    ' Now we should have the parent form
    GetParentFormName = objParent.Name
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Remove temporary controls, if applicable
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Me.Reset
    Me.Hide
    ' Ignore any errors if object references are no longer valid
    On Error Resume Next
    #If APPLICATION_NAME = "Microsoft Access" Then
        ' Hide front controls
        this.lblFront.Visible = False
        this.lblCaption.Visible = False
    #Else
        ' Remove front controls
        this.objParent.Controls.Remove (mlblFront.Name)
        this.objParent.Controls.Remove (mlblCaption.Name)
    #End If
    Err.Clear ' Clear any errors, they mean the form closed before the class terminated.

End Sub
