Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database


Private Const cstrCmdBarName As String = "Version Control"

' Model object used for menu commands (supports multiple versioning systems)
Private m_Model As IVersionControl

' Menu command bar
Private m_CommandBar As Office.CommandBar

' Menu button events
Private WithEvents m_evtSaveAll As VBIDE.CommandBarEvents
Attribute m_evtSaveAll.VB_VarHelpID = -1
Private WithEvents m_evtSave As VBIDE.CommandBarEvents
Attribute m_evtSave.VB_VarHelpID = -1
Private WithEvents m_evtCommit As VBIDE.CommandBarEvents
Attribute m_evtCommit.VB_VarHelpID = -1
Private WithEvents m_evtDiff As VBIDE.CommandBarEvents
Attribute m_evtDiff.VB_VarHelpID = -1
'------
Private WithEvents m_evtExpSel As VBIDE.CommandBarEvents
Attribute m_evtExpSel.VB_VarHelpID = -1
Private WithEvents m_evtExpAll As VBIDE.CommandBarEvents
Attribute m_evtExpAll.VB_VarHelpID = -1
Private WithEvents m_evtSimImp As VBIDE.CommandBarEvents
Attribute m_evtSimImp.VB_VarHelpID = -1
Private WithEvents m_evtImpObj As VBIDE.CommandBarEvents
Attribute m_evtImpObj.VB_VarHelpID = -1
Private WithEvents m_evtImpAll As VBIDE.CommandBarEvents
Attribute m_evtImpAll.VB_VarHelpID = -1
Private WithEvents m_evtClrPrj As VBIDE.CommandBarEvents
Attribute m_evtClrPrj.VB_VarHelpID = -1
Private WithEvents m_evtRstPth As VBIDE.CommandBarEvents
Attribute m_evtRstPth.VB_VarHelpID = -1
Private WithEvents m_evtDscVCS As VBIDE.CommandBarEvents
Attribute m_evtDscVCS.VB_VarHelpID = -1



'---------------------------------------------------------------------------------------
' Procedure : Construct
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Construct an instance of this class using a specific model
'---------------------------------------------------------------------------------------
'
Public Sub Construct(cModel As IVersionControl)
    
    ' Save reference to model
    If Not m_Model Is Nothing Then m_Model.Terminate
    Set m_Model = cModel

    ' Verify that the required software is installed
    If m_Model.HasRequiredSoftware(True) Then
    
        ' Set up toolbar
        If CommandBarExists(cstrCmdBarName) Then
            Set m_CommandBar = Application.VBE.CommandBars(cstrCmdBarName)
            ' Reassign buttons so we can capture events
            RemoveAllButtons
        Else
            ' Add toolbar
            Set m_CommandBar = Application.VBE.CommandBars.Add
            With m_CommandBar
                .Name = cstrCmdBarName
                .Position = msoBarTop
                .Visible = True
            End With
        End If
        
        ' Assign/reassign buttons so we can capture events
        AddAllButtons
    
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CommandBarExists
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the command bar exists. (Is visible)
'---------------------------------------------------------------------------------------
'
Private Function CommandBarExists(strName As String) As Boolean
    Dim cmdBar As CommandBar
    For Each cmdBar In Application.VBE.CommandBars
        If cmdBar.Name = strName Then
            CommandBarExists = True
            Exit For
        End If
    Next cmdBar
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddAllButtons
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Add the buttons to the command bar
'---------------------------------------------------------------------------------------
'
Private Sub AddAllButtons()

    If m_CommandBar Is Nothing Then Exit Sub

    ' Add buttons with event handlers
    With Application.VBE.Events
'        Set m_evtCommit = .CommandBarEvents(AddButton("Commit Module/Project", 270))
'        Set m_evtDiff = .CommandBarEvents(AddButton("Diff Module/Project", 2042, , True))
'        Set m_evtSave = .CommandBarEvents(AddButton("Export Selected", 3))
'        Set m_evtSaveAll = .CommandBarEvents(AddButton("Export All", 749, , , msoButtonIconAndCaption))
        
        Set m_evtDscVCS = .CommandBarEvents(AddButton("Disconnect VCS Lib", 2309))
        Set m_evtClrPrj = .CommandBarEvents(AddButton("Clear Project", 7674))
        Set m_evtRstPth = .CommandBarEvents(AddButton("Reset Export Path", 19699, , True))
        
        Set m_evtImpAll = .CommandBarEvents(AddButton("Import All", 2116, , , msoButtonIconAndCaption))
        Set m_evtImpObj = .CommandBarEvents(AddButton("Import Objects", 14428))
        Set m_evtSimImp = .CommandBarEvents(AddButton("Simulate Objects Import", 17267, , True))
        
        Set m_evtExpSel = .CommandBarEvents(AddButton("Export Selected", 1679))
        Set m_evtExpAll = .CommandBarEvents(AddButton("Export All", 2109, , , msoButtonIconAndCaption))
        
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddButton
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Add a button to the command bar, and connects to event handler
'---------------------------------------------------------------------------------------
'
Private Function AddButton(strCaption As String, intFaceID As Integer, _
    Optional intPositionBefore As Integer = 1, Optional blnBeginGroup As Boolean = False, Optional intStyle As MsoButtonStyle) As CommandBarButton
    
    Dim btn As CommandBarButton
    
    Set btn = m_CommandBar.Controls.Add(msoControlButton, , , intPositionBefore)
    btn.Caption = strCaption
    btn.FaceId = intFaceID
    btn.Style = intStyle
    If blnBeginGroup Then btn.BeginGroup = True
    Set AddButton = btn

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveAllButtons
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Removes all the buttons from the command bar
'---------------------------------------------------------------------------------------
'
Private Sub RemoveAllButtons()
    Dim btn As CommandBarButton
    If Not m_CommandBar Is Nothing Then
        For Each btn In m_CommandBar.Controls
            btn.Delete
        Next btn
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Release all references
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()

    ' Clear event handlers
    Set m_evtCommit = Nothing
    Set m_evtDiff = Nothing
    Set m_evtSave = Nothing
    
    ' Finish cleaning up
    RemoveAllButtons
    If Not m_CommandBar Is Nothing Then
        m_CommandBar.Delete
        Set m_CommandBar = Nothing
    End If
    ' Don't terminate a circular reference
    ' since menu is a child of the model
    Set m_Model = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : (multiple)
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Event handlers for button clicks
'---------------------------------------------------------------------------------------
'
Private Sub m_evtCommit_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_Model.Commit
    handled = True
End Sub
Private Sub m_evtDiff_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    m_Model.Diff
    handled = True
End Sub
Private Sub m_evtSave_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If CloseAllFormsReports Then ExportSelected
    handled = True
End Sub
Private Sub m_evtSaveAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If CloseAllFormsReports Then m_Model.ExportAll
    handled = True
End Sub

'--------------------------------
Private Sub m_evtExpAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    If CloseAllFormsReports Then m_Model.ExportAll
    handled = True
End Sub
Private Sub m_evtExpSel_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    If CloseAllFormsReports Then ExportSelected
    handled = True
End Sub

Private Sub m_evtSimImp_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    pub_LIBVCS_ImportObjects True
    handled = True
End Sub
Private Sub m_evtImpObj_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    pub_LIBVCS_ImportObjects
    handled = True
End Sub
Private Sub m_evtImpAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    MsgBox "pub_LIBVCS_ImportAll not complete yet (i.e.:tables import doesn't work)"
'    pub_LIBVCS_ImportAll
    handled = True
End Sub
Private Sub m_evtRstPth_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    pub_LIBVCS_ChangeExportPath
    handled = True
End Sub
Private Sub m_evtClrPrj_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    MsgBox "Temporarily disabled because pub_LIBVCS_ImportAll is not complete yet (i.e.:tables import doesn't work)"
'    pub_LIBVCS_ResetProject
    handled = True
End Sub
Private Sub m_evtDscVCS_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Debug.Print CommandBarControl.Caption
    If pub_LIBVCS_RemoveReferenceByName("MSAccessVCS") Then
        MsgBox "- Reference to 'MSAccessVCS' has been removed", , "Libray is disconnected!"
    End If
    handled = True
End Sub





'---------------------------------------------------------------------------------------
' Procedure : ExportSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Export the selected component or project
'---------------------------------------------------------------------------------------
'
Private Sub ExportSelected()
    If SelectionInActiveProject Then
        m_Model.Export
    Else
        MsgBox "Please select a component in " & CurrentProject.Name & " and try again.", vbExclamation, CodeProject.Name
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Manually fire the terminate event
'---------------------------------------------------------------------------------------
'
Public Sub Terminate()
    Call Class_Terminate
End Sub