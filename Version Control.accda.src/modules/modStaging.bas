﻿Attribute VB_Name = "modStaging"
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
Public Sub StageMainForm()

    Dim blnLogActive As Boolean
    Dim frm As Form_frmVCSMain

    ' Make sure the form is actually open, just in case.
    ' (We want to capture the current text on the form
    '  to display it again when the form is restored.)
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden

    ' Get reference to form instance
    Set frm = Form_frmVCSMain

    With this
        ' Get headings and content
        .strHeading = frm.lblHeading.Caption
        .strSubHeading = frm.lblSubheading.Caption
        .strLogHtml = Nz(frm.txtLog.Value)

        ' Disconnect from logging class
        If Not Log.ProgressBar Is Nothing Then
            .blnProgVisible = frm.lblProgFront.Visible
            .dblProgMax = Log.ProgressBar.Max
            .dblProgValue = Log.ProgressBar.Value
        End If
        Log.ReleaseConsole
    End With

    ' Make sure we stage any current operation before closing the main
    ' form to avoid a warning to the user about canceling the current operation.
    Operation.Stage
    DoCmd.Close acForm, frm.Name
    Operation.Restore
    Set frm = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreForm
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Restore the form to the staged values.
'---------------------------------------------------------------------------------------
'
Public Sub RestoreMainForm()

    Dim frm As Form_frmVCSMain

    ' Open form in hidden mode, and get reference to instance
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    Set frm = Form_frmVCSMain

    ' Restore settings on form
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
