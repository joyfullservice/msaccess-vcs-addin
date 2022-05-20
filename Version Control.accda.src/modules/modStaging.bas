Attribute VB_Name = "modStaging"
'---------------------------------------------------------------------------------------
' Module    : modStaging
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Handle the staging and restoring of the main form.
'           : (It may be automatically closed when closing the current database, or it
'           :  may be closed intentionally before importing a form with the same name.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


Private Type udtThis
    strHeading As String
    strSubHeading As String
    strLogHtml As String
    blnProgVisible As Boolean
    dblProgMax As Double
    dblProgValue As Double
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : StageForm
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Stage the content of the form so we can close it then restore it later.
'---------------------------------------------------------------------------------------
'
Public Sub StageForm(ByRef frm As Form_frmVCSMain)

    Dim blnLogActive As Boolean
    
    With this
        ' Get headings and content
        .strHeading = frm.lblHeading.Caption
        .strSubHeading = frm.lblSubheading.Caption
        .strLogHtml = frm.txtLog.Value
    
        ' Disconnect from logging class
        If Not Log.ProgressBar Is Nothing Then
            .blnProgVisible = frm.lblProgFront.Visible
            .dblProgMax = Log.ProgressBar.Max
            .dblProgValue = Log.ProgressBar.Value
        End If
        Log.ReleaseConsole
    End With
    
    ' Temporarily deactivate the log so we don't trigger warnings when closing the form.
    blnLogActive = Log.Active
    Log.Active = False
    
    ' Close the form, if it is open
    DoCmd.Close acForm, frm.Name
    Set frm = Nothing
    
    ' Restore active property to original value
    Log.Active = blnLogActive
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreForm
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Restore the form to the staged values.
'---------------------------------------------------------------------------------------
'
Public Sub RestoreForm(frm As Form_frmVCSMain)

    With this
        ' Get headings and content
        frm.lblHeading.Caption = .strHeading
        frm.lblSubheading.Caption = .strSubHeading
        frm.txtLog.Value = .strLogHtml
        frm.txtLog.Visible = True
    
        ' Reconnect to logging class
        Log.SetConsole frm.txtLog, frm.GetProgressBar
        If .blnProgVisible Then
            Log.ProgressBar.Max = .dblProgMax
            Log.ProgressBar.Value = .dblProgValue
        Else
            Log.ProgressBar.Hide
        End If
        
        ' Assume that the action buttons should be hidden
        frm.cmdClose.SetFocus
        frm.HideActionButtons
        frm.Visible = True
        DoEvents
    End With
    
End Sub
