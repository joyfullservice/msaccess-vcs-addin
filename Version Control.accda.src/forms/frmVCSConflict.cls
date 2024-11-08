Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 2/4/2021
' Purpose   : Close form if user presses the escape key. (Cancel property = True)
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdContinue_Click
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : User selected to continue. Close and process conflicts.
'---------------------------------------------------------------------------------------
'
Private Sub cmdContinue_Click()

    Dim lngRemaining As Long

    lngRemaining = ActionDecisionsNeeded
    If lngRemaining > 0 Then
        MsgBox2 "Please Resolve Conflicts", _
            lngRemaining & " item(s) need to be resolved to continue.", _
            "You can cancel this operation if you need to do further review.", vbInformation
    Else
        VCSIndex.Conflicts.ApproveResolutions = True
        DoCmd.Close acForm, Me.Name
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Setting the control source causes delayed display. This way the display
'           : is instant when the form is opened.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    ' Load translations
    Translation.ApplyTo Me

    ' Display version (better performance than bound control)
    lblVersion.Caption = Replace(lblVersion.Caption, "${version}", GetVCSVersion())

    Dim frmList As Form_frmVCSConflictList
    Set frmList = Me.sfrmConflictList.Form

    ' Update heading caption
    If Log.OperationType = eotExport Then
        lblHeading.Caption = "These source files have changed since the last export"
        With frmList.cboResolution
            .AddItem eResolveConflict.ercSkip & ";" & "Skip"
            .AddItem eResolveConflict.ercOverwrite & ";" & "Overwrite source file"
        End With
    Else
        lblHeading.Caption = "These database objects have changed since the last export"
        With frmList.cboResolution
            .AddItem eResolveConflict.ercSkip & ";" & "Skip"
            .AddItem eResolveConflict.ercOverwrite & ";" & "Overwrite database object"
        End With
    End If

    ' Change to resizable form
    MakeDialogResizable Me

    ' Set initial column size
    DoEvents
    Form_Resize

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Adjust column widths on subform.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
    ScaleColumns Me.sfrmConflictList.Form, , _
        Array("txtObjectDate", "txtFileDate", "txtDiff")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdOverwriteAll_Click
' Author    : Adam Waller
' Date      : 11/1/2021
' Purpose   : Overwrite all items in the list
'---------------------------------------------------------------------------------------
'
Private Sub cmdOverwriteAll_Click()
    sfrmConflictList.SetFocus
    CodeDb.Execute "update tblConflicts set Resolution=" & eResolveConflict.ercOverwrite, dbFailOnError
    sfrmConflictList.Requery
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSkipAll_Click
' Author    : Adam Waller
' Date      : 11/1/2021
' Purpose   : Skip all items in the list
'---------------------------------------------------------------------------------------
'
Private Sub cmdSkipAll_Click()
    sfrmConflictList.SetFocus
    CodeDb.Execute "update tblConflicts set Resolution=" & eResolveConflict.ercSkip, dbFailOnError
    sfrmConflictList.Requery
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NeedsActionDecision
' Author    : Adam Waller
' Date      : 2/22/2023
' Purpose   : Count how many records need a decision before we can continue.
'---------------------------------------------------------------------------------------
'
Private Function ActionDecisionsNeeded() As Long

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset

    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset( _
        "select count(*) as Remaining from (select id from tblConflicts where NZ(Resolution)=0)", _
        dbOpenDynaset, dbOpenForwardOnly, dbReadOnly)

    ActionDecisionsNeeded = Nz(rst!Remaining)
    rst.Close
    Set rst = Nothing
    Set dbs = Nothing

End Function
