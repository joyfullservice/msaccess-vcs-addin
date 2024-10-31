Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSMain
' Author    : Adam Waller
' Date      : 3/30/2022
' Purpose   : Main form for performing an export or build operation.
'           : This color scheme can be changed, I just wanted something more aesthetically
'           : pleasing than the default wizards and forms.
'           : Color scheme: https://coolors.co/383f51-e0e0e6-ffffff-ef8354-d3d7ef
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' This property can be set to export or merge a specific subset of containers
Public intContainerFilter As eContainerFilter

' Used for exporting or loading a single object
Public objSingleObject As AccessObject

' Path to the last log file, in case the user wants to view the log after the operation.
' (The Log object has already been reset at this point, so we can't use Log.LogFilePath.)
Public strLastLogFilePath As String

' Use this property to set the path to the source files (such as a build triggered from the API)
Public strSourcePath As String


'---------------------------------------------------------------------------------------
' Procedure : cmdBuild_Click
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Initiate the process to build from source
'---------------------------------------------------------------------------------------
'
Public Sub cmdBuild_Click()

    Dim strFolder As String

    ' Make sure we use the add-in to build the add-in.
    If CodeProject.FullName = CurrentProject.FullName Then
        MsgBox2 "Build must be run from Add-In", "Instead of opening this form to build the add-in," & vbCrLf & _
            "please install and use the Version Control add-in from the Add-in menu", , vbExclamation
        DoCmd.Hourglass False
        Exit Sub
    End If

    ' Get source files folder
    If Len(Me.strSourcePath) Then
        ' Use specified build folder
        strFolder = Me.strSourcePath
    Else
        ' Attempt to get the source folder from the current database, or from
        ' a folder picker dialog.
        strFolder = GetSourceFolder
        ' Exit out of build if the user cancelled any of the confirmations.
        If strFolder = vbNullString Then Exit Sub
    End If

    ' Build project using the selected source folder
    ' (Use a timer so we can release the reference to this form before beginning the
    '  build process, just in case we need to import a form with the same name.)
    If strFolder <> vbNullString Then SetTimer "Build", strFolder, chkFullBuild

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSourceFolder
' Author    : Adam Waller
' Date      : 10/19/2023
' Purpose   : Return the source files folder from either the currently open database
'           : or from a folder picker dialog. (Returns an empty string if the user
'           : cancels the selection.)
'---------------------------------------------------------------------------------------
'
Private Function GetSourceFolder() As String

    Dim strMsg(0 To 2) As String
    Dim intChoice As VbMsgBoxResult

    DoCmd.Hourglass True
    DoEvents

    ' Close the current database if it is currently open.
    If DatabaseFileOpen Then
        If FolderHasVcsOptionsFile(Options.GetExportFolder) Then
            strMsg(0) = "Build " & CurrentVBProject.Name & " (" & CurrentProject.Name & ") from source?"
            If chkFullBuild Then
                strMsg(1) = "Click 'Yes' to rebuild* this database from source files in this folder:" & vbCrLf & Options.GetExportFolder & vbCrLf & _
                    "* (This database will be renamed as a backup before building " & CurrentProject.Name & " from source.)"
            Else
                strMsg(1) = "Click 'Yes' to merge* any changed source files into this database." & vbCrLf & _
                    "* (A backup of this database will be created before importing any source files.)"
            End If
            strMsg(2) = "Click 'No' to select another project, or 'Cancel' to go back to the previous screen."
            If Not chkFullBuild And Not Me.Visible Then
                ' Skip confirmation for merge build initiated from Ribbon
                intChoice = vbYes
            Else
                ' Require user confirmation for full builds, or if main form is visible.
                intChoice = MsgBox2(strMsg(0), strMsg(1), strMsg(2), vbYesNoCancel + vbQuestion + vbDefaultButton3, , vbYes)
            End If
            If intChoice = vbYes Then
                ' Rebuild the open project
                GetSourceFolder = Options.GetExportFolder
            ElseIf intChoice = vbCancel Then
                ' Canceled out of build option.
                DoCmd.Hourglass False
                Exit Function
            End If
        End If
    End If

    ' If we aren't doing the current database, then prompt user to find a folder
    ' with source files to use for the build.
    If GetSourceFolder = vbNullString Then

        ' Show a folder picker to select the file with source code.
        DoCmd.Hourglass False
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            .ButtonName = "Select Source Files Folder"
            '.InitialFileName = Options.GetExportFolder
            .Title = "Select Source Folder"
            .Show
            If .SelectedItems.Count > 0 Then
                ' Selected a folder
                If FolderHasVcsOptionsFile(.SelectedItems(1)) Then
                    ' Has source files
                    GetSourceFolder = .SelectedItems(1) & PathSep
                    DoCmd.Hourglass True
                Else
                    MsgBox2 "Source files not found", "Required source files were not found in this folder.", _
                        "You selected: " & .SelectedItems(1), vbExclamation
                    Exit Function
                End If
            Else
                ' Canceled dialog
                Exit Function
            End If
        End With
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildFromSource
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Show the GUI for building the database from source.
'---------------------------------------------------------------------------------------
'
Public Sub StartBuild(blnFullBuild As Boolean)

    Dim strType As String

    cmdClose.SetFocus
    HideActionButtons
    DoEvents
    With txtLog
        .ScrollBars = 0
        .Visible = True
        .SetFocus
    End With
    Log.SetConsole Me.txtLog, GetProgressBar
    Me.Visible = True

    ' Show the status
    strType = IIf(blnFullBuild, "Building", "Merging")
    SetStatusText "Running...", strType & " From Source", _
        "A summary of the build progress can be seen on this screen, and additional details are included in the log file."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FinishBuild
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Finish the build process
'---------------------------------------------------------------------------------------
'
Public Sub FinishBuild(blnFullBuild As Boolean) 'Optional strType As String = "Build")

    Dim strType As String

    ' Turn on scroll bars in case the user wants to scroll back through the log.
    txtLog.ScrollBars = 2

    ' Display final UI messages.
    Log.Flush
    strType = IIf(blnFullBuild, "Build", "Merge")
    SetStatusText "Finished", strType & " Complete", _
        "Additional details can be found in the project " & LCase(strType) & " log file.<br><br>You may now close this window."
    cmdOpenLogFile.Visible = (Log.LogFilePath <> vbNullString)
    Me.strLastLogFilePath = Log.LogFilePath

    ' Close form if running in silent mode (such as when running automated builds)
    If InteractionMode = eimSilent Then DoCmd.Close acForm, Me.Name

End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkFullBuild_Click
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkFullBuild_Click()
    SetBuildCaption
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetBuildCaption
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Set the appropriate caption for the build/merge button.
'---------------------------------------------------------------------------------------
'
Private Sub SetBuildCaption()
    cmdBuild.Caption = IIf(chkFullBuild, _
        " Build from Source", _
        " Merge from Source")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdClose_Click
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Close the form
'---------------------------------------------------------------------------------------
'
Private Sub cmdClose_Click()
    ' Ignore the error if the user resumes (cancels the close operation)
    LogUnhandledErrors
    On Error Resume Next
    DoCmd.Close acForm, Me.Name
    Catch 2501  ' Close form was canceled.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConfirmCancel
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Confirm that the user really wants to cancel the current operation.
'---------------------------------------------------------------------------------------
'
Private Function ConfirmCancel() As Boolean
    ConfirmCancel = MsgBox2("Cancel Current Operation?", _
        "You are in the midst of a running process. Are you sure you want to cancel?", _
        "Click [Yes] to cancel the process, or [No] to resume.", _
        vbYesNo + vbDefaultButton2 + vbExclamation, , vbYes) = vbYes
End Function


'---------------------------------------------------------------------------------------
' Procedure : cmdExport_Click
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Export source code from current database
'---------------------------------------------------------------------------------------
'
Public Sub cmdExport_Click()

    cmdClose.SetFocus
    HideActionButtons
    DoEvents
    With txtLog
        .ScrollBars = 0
        .Visible = True
        .SetFocus
    End With
    Log.SetConsole Me.txtLog, GetProgressBar

    ' Show the status
    SetStatusText "Running...", "Exporting source code", _
        "A summary of the export progress can be seen on this screen, and additional details are included in the log file."

    ' See if we are exporting a single object, or everything.
    If Me.objSingleObject Is Nothing Then
        ' Export the source code using the specified filter.
        modImportExport.ExportSource chkFullExport, Me.intContainerFilter, Me
    Else
        modImportExport.ExportSingleObject Me.objSingleObject, Me
    End If

    ' Turn on scroll bars in case the user wants to scroll back through the log.
    txtLog.ScrollBars = 2
    Log.Flush

    ' Don't attempt to access controls if we are in the process of closing the form.
    If FormLoaded(Me) Then
        SetStatusText "Finished", "Export Complete", _
            "Additional details can be found in the project export log file.<br><br>You may now close this window."
        cmdOpenLogFile.Visible = (Me.strLastLogFilePath <> vbNullString)
        Me.strLastLogFilePath = Me.strLastLogFilePath
        DoEvents
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HideActionButtons
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Hide the action buttons when running a command.
'---------------------------------------------------------------------------------------
'
Public Sub HideActionButtons()
    cmdExport.Visible = False
    chkFullExport.Visible = False
    cmdBuild.Visible = False
    chkFullBuild.Visible = False
    cmdOptions.Visible = False
    cmdHelp.Visible = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdHelp_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Go to the GitHub documentation page.
'---------------------------------------------------------------------------------------
'
Private Sub cmdHelp_Click()
    Application.FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Documentation"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdOptions_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Open options dialog (for this project)
'---------------------------------------------------------------------------------------
'
Private Sub cmdOptions_Click()
    If DatabaseFileOpen Then
        ' Force reload of options from current project before opening the form.
        Set Options = Nothing
        Form_frmVCSOptions.Visible = True
    Else
        ' The table data subform may crash Access when the options form is closed with no database open.
        MsgBox2 "No Database Open", "Please open a database file before configuring VCS options.", , vbInformation
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
Public Sub Form_Load()

    ' Display version (better performance than bound control)
    lblVersion.Caption = "Version " & GetVCSVersion()

    SetStatusText "Choose Action", "What would you like to do?", _
        "<strong><em>Export</em></strong> source to generate source files from the current database." & _
        "<br><br><strong><em>Import</em></strong> source files to rebuild this database from source."

    ' Set defaults based on current options.
    chkFullBuild = Not Options.UseMergeBuild
    chkFullExport = Not Options.UseFastSave

    ' You can only export if you have a database open.
    cmdExport.Enabled = DatabaseFileOpen
    chkFullExport.Enabled = DatabaseFileOpen

    If DatabaseFileOpen Then

        ' Require full export after options change
        If VCSIndex.OptionsHash <> Options.GetHash Then
            chkFullExport = True
            chkFullExport.Enabled = False
        End If

        ' Merge build only available after full build.
        ' (Attempting a merge build of the entire database may
        '  not work correctly due to objects that depend upon
        '  each other.)
        If VCSIndex.FullBuildDate = 0 Then
            chkFullBuild = True
            chkFullBuild.Enabled = False
        End If
    End If

    ' Set caption on build button
    SetBuildCaption

    ' Turn off the timer, just in case it was left on.
    Me.TimerInterval = 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AutoClose
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Use the timer to automatically close the form in 2 seconds.
'           : (This keeps the application from "hanging" during the pause between
'           :  completion and close.)
'---------------------------------------------------------------------------------------
'
Public Sub AutoClose()
    'The procedure may be called when the form has been closed.
    'In this case, a VBA error may occur, so we check if the
    'form is loaded before setting the property. We do not use
    'the Me.Name because that would be also an error.
    If IsLoaded(acForm, "frmVCSMain", False) Then
        Me.TimerInterval = 2000
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Timer
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Automatically close form.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Timer()
    Me.TimerInterval = 0
    cmdClose_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetStatusText
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Update the status text to direct the user to the next task, or show the
'           : status of the current operation.
'---------------------------------------------------------------------------------------
'
Public Sub SetStatusText(strHeading As String, strSubHeading As String, strDescriptionHtml As String)
    If Not FormLoaded(Me) Then Exit Sub
    lblHeading.Caption = strHeading
    lblSubheading.Caption = strSubHeading
    txtDescription.Value = strDescriptionHtml
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetProgressBar
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Set up the progress bar.
'---------------------------------------------------------------------------------------
'
Public Function GetProgressBar() As clsLblProg

    Dim cProg As clsLblProg

    ' Create a new progress bar class, and initialize with form controls
    Set cProg = New clsLblProg
    cProg.Initialize lblProgBack, lblProgFront, lblProgCaption
    cProg.SetRepaintInterval 0.5

    ' Move caption up slightly for better alignment on this very small progress bar
    lblProgCaption.Top = lblProgBack.Top + 1

    ' Return reference to caller
    Set GetProgressBar = cProg

End Function


'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Verify that the user wants to cancel the current operation
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)

    Static intAttempt As Integer

    ' Allow the form to close on the third attempt, just in case the log
    ' is stuck in active status for some reason.
    If intAttempt > 2 Then Exit Sub

    ' Check to see if we have an active job running.
    If Log.Active Then
        If ConfirmCancel Then Log.Error eelCritical, "Canceled Operation", Me.Name & ".Form_Unload"
        ' Either way, we want the log to complete first.
        Cancel = True
        intAttempt = intAttempt + 1
    End If

    ' Release the log console if we are closing the form
    If Not Cancel Then Log.ReleaseConsole

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdOpenLogFile_Click
' Author    : Adam Waller
' Date      : 11/6/2020
' Purpose   : Open the log file
'---------------------------------------------------------------------------------------
'
Private Sub cmdOpenLogFile_Click()
    cmdClose.SetFocus
    If FSO.FileExists(strLastLogFilePath) Then
        ' (Note, parentheses are required for the path argument)
        CreateObject("Shell.Application").Open (strLastLogFilePath)
    End If
End Sub
