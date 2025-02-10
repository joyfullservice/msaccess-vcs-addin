VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVersionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsVersionControl
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This class is publicly exposed through the VCS object in modAPI and can
'           : be used to automate add-in functionality from other systems.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Show
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Show the Version Control System main form
'---------------------------------------------------------------------------------------
'
Public Sub Show()
    If HasFormOpen Then Exit Sub
    DoCmd.OpenForm "frmVCSMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShowOptions
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Show the options form
'---------------------------------------------------------------------------------------
'
Public Sub ShowOptions()
    If HasFormOpen Then Exit Sub
    If DatabaseFileOpen Then
        ' Force reload of options from current project before opening the form.
        Set modObjects.Options = Nothing
        DoCmd.OpenForm "frmVCSOptions"
    Else
        ' The table data subform may crash Access when the options form is closed with no database open.
        MsgBox2 "No Database Open", "Please open a database file before configuring VCS options.", , vbInformation
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Export the source code for the current database
'---------------------------------------------------------------------------------------
'
Public Sub Export()
    If HasFormOpen Then Exit Sub
    RunExport ecfAllObjects
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVBA
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Export just the VBA related components
'---------------------------------------------------------------------------------------
'
Public Sub ExportVBA()
    If HasFormOpen Then Exit Sub
    RunExport ecfVBAItems
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportSelected
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Export the selected object
'---------------------------------------------------------------------------------------
'
Public Sub ExportSelected()

    Dim objSelected As AccessObject

    If HasFormOpen Then Exit Sub

    Set objSelected = GetSelectedNavPaneObject
    If objSelected Is Nothing Then
        MsgBox2 "Please Select an Object First", _
            "Select a single object in the Navigation Pane to export.", _
            "(This item must have the keyboard focus.)", vbInformation
    Else
        ' Export the item
        RunExport , objSelected
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunExport
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Handle different kinds of exports based on filter
'---------------------------------------------------------------------------------------
'
Private Sub RunExport(Optional intFilter As eContainerFilter = ecfAllObjects, Optional objItem As AccessObject)
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    With Form_frmVCSMain
        If objItem Is Nothing Then
            .intContainerFilter = intFilter
        Else
            Set .objSingleObject = objItem
        End If
        .Visible = True
        .cmdExport_Click
        If Log.ErrorLevel < eelError Then .AutoClose
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Build
' Author    : Adam Waller
' Date      : 3/30/2022
' Purpose   : Initiate a full build from source
'---------------------------------------------------------------------------------------
'
Public Sub Build(Optional strSourceFolder As String)
    If HasFormOpen Then Exit Sub
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    With Form_frmVCSMain
        ' Make sure we are doing a full build.
        If Not .chkFullBuild Then .chkFullBuild = True
        .strSourcePath = strSourceFolder
        .cmdBuild_Click
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeBuild
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Merge modified source files into the existing database.
'---------------------------------------------------------------------------------------
'
Public Sub MergeBuild()
    If HasFormOpen Then Exit Sub
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    With Form_frmVCSMain
        ' See if merge build is available
        If .chkFullBuild.Enabled Then
            ' Initiate the merge build
            .chkFullBuild = False
            .cmdBuild_Click
        Else
            MsgBox2 "Merge Build Not Available", _
                "A full build must be performed before you can merge from source.", _
                "Please perform a full build of this database first.", vbInformation
        End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildAs
' Author    : Adam Waller
' Date      : 7/8/2023
' Purpose   : Build the source files to a new file name or location.
'           : If source files are identified for the currently open database, it will
'           : build from those files. Otherwise it will ask the use to select a folder
'           : with source files.
'---------------------------------------------------------------------------------------
'
Public Sub BuildAs()

    Dim strNewDbPath As String
    Dim strSourceFolder As String

    If HasFormOpen Then Exit Sub

    ' See if we can find source files for the currently open database.
    If DatabaseFileOpen Then
        If FolderHasVcsOptionsFile(Options.GetExportFolder) Then
            ' Get the source folder location.
            strSourceFolder = Options.GetExportFolder
        End If
    End If

    ' If we aren't doing the current database, then prompt user to find a folder
    ' with source files to use for the build.
    If strSourceFolder = vbNullString Then

        ' Show a folder picker to select the file with source code.
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            .ButtonName = "Select Source Files Folder"
            .Title = "Select Source Folder"
            .Show
            If .SelectedItems.Count > 0 Then
                ' Selected a folder
                If FolderHasVcsOptionsFile(.SelectedItems(1)) Then
                    ' Has source files
                    strSourceFolder = .SelectedItems(1) & PathSep
                Else
                    MsgBox2 "Source files not found", "Required source files were not found in this folder.", _
                        "You selected: " & .SelectedItems(1), vbExclamation
                    Exit Sub
                End If
            Else
                ' Canceled dialog
                Exit Sub
            End If
        End With
    End If

    ' At this point, we should have identified the source files folder.
    ' Now, we need to get the desired new database file name.
    With Application.FileDialog(msoFileDialogSaveAs)
        .AllowMultiSelect = False
        .ButtonName = "Build Here"
        .Title = "Build New Database File"
        If DatabaseFileOpen Then
            .InitialFileName = CurrentProject.FullName
        Else
            .InitialFileName = GetOriginalDbFullPathFromSource(strSourceFolder)
        End If
        .Show
        ' Return path to new file
        If .SelectedItems.Count > 0 Then strNewDbPath = .SelectedItems(1)
    End With

    ' Proceed with build if we have the new path.
    If Len(strNewDbPath) Then
        modImportExport.Build strSourceFolder, True, , strNewDbPath
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeAllSource
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Merge ALL source files into the current database, regardless of modified
'           : status. (Used for testing and development purposes.)
'---------------------------------------------------------------------------------------
'
Public Sub MergeAllSource()
    If HasFormOpen Then Exit Sub
    modImportExport.MergeAllSource
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadSelected
' Author    : Adam Waller
' Date      : 2/21/2023
' Purpose   : Load the selected item from source files
'---------------------------------------------------------------------------------------
'
Public Sub LoadSelected()

    Dim objSelected As AccessObject
    Dim cComponentClass As IDbComponent
    Dim strObjectName As String
    Dim strSourceFilePath As String

    If HasFormOpen Then Exit Sub

    Set objSelected = GetSelectedNavPaneObject
    If objSelected Is Nothing Then
        MsgBox2 "Please Select an Object First", _
            "Select a single object in the Navigation Pane to reload from source.", _
            "(This item must have the keyboard focus.)", vbInformation
    Else
        ' Remove all object references to selected item
        Set cComponentClass = GetClassFromObject(objSelected)
        With cComponentClass
            strObjectName = .Name
            strSourceFilePath = .SourceFile
            Set .DbObject = Nothing
        End With
        Set objSelected = Nothing

        ' Import the object from source files
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
        Form_frmVCSMain.StartBuild False
        LoadSingleObject cComponentClass, strObjectName, strSourceFilePath
        Form_frmVCSMain.FinishBuild False
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetOptions
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Access the options
'---------------------------------------------------------------------------------------
'
Public Function Options() As clsOptions
    Set Options = modObjects.Options
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenRepository
' Author    : Adam Waller
' Date      : 2/8/2025
' Purpose   : Open the configured tool for interacting with the repository.
'           : (I.e GitHub Desktop or SourceTree)
'---------------------------------------------------------------------------------------
'
Public Sub OpenRepository()

    Dim strFolder As String
    Dim strPath As String

    ' Build quoted path to source folder
    strFolder = """" & StripSlash(Options.GetExportFolder) & """"

    ' Launch user-specified tool
    Select Case Nz2(GetSetting(PROJECT_NAME, "Options", "Open Repository Tool"), 0)
        Case eraUndefined
            MsgBox2 "No Repository Tool Selected", _
                "Please choose a program to launch in the add-in Options", _
                "GitHub Desktop is a great choice for most users", vbInformation

        Case eraGitHubDesktop
            ' Open GitHub Desktop, and switch to the relevant repository.
            ' Utilizes a built-in batch file installed as a part of GitHub Desktop
            ShellEx "GitHub.bat", strFolder, , False

        Case eraVSCode  ' Visual Studio Code
            ShellEx "code", "--reuse-window " & strFolder, , False

        Case eraSourceTree
            ' Look up install path for SourceTree
            strPath = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\SourceTree\InstallLocation")
            ' Check install for all users
            If strPath = vbNullString Then strPath = RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\SourceTree\InstallLocation")
            If Len(strPath) Then
                Shell FSO.BuildPath(strPath, "SourceTree.exe") & " -f " & strFolder
                AppActivate "Sourcetree", True  ' Try to switch to this window
            Else
                MsgBox2 "SourceTree Not Found", "Path to SourceTree.exe could not be found in the Windows Registry", , vbExclamation
            End If

        Case eraTortoiseGit
            ShellEx "TortoiseGitProc.exe", "/path:" & strFolder
            ' Could also go directly to commit:
            'ShellEx "TortoiseGitProc.exe", "/command:commit /path:" & strFolder

    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenSourceFolder
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open the source files folder for the current project
'---------------------------------------------------------------------------------------
'
Public Sub OpenSourceFolder()
    Shell "explorer """ & Options.GetExportFolder & """", vbNormalFocus
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenExportLog
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open the export log file
'---------------------------------------------------------------------------------------
'
Public Sub OpenExportLog()
    OpenLog "Export.log"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenBuildLog
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open the build log file
'---------------------------------------------------------------------------------------
'
Public Sub OpenBuildLog()
    OpenLog "Build.log"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenLog
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open a log file by name, or show a message if the file is not found.
'---------------------------------------------------------------------------------------
'
Private Sub OpenLog(strName As String)

    Dim strPath As String

    strPath = Options.GetExportFolder & strName
    If FSO.FileExists(strPath) Then
        ' Note that the parentheses around the path are required for this to work.
        CreateObject("Shell.Application").Open (strPath)
    Else
        MsgBox2 "Log File Not Found", "Could not find the following file in the export location:", strPath, vbExclamation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReloadRibbon
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Reload the add-in ribbon to reflect changes in XML source
'---------------------------------------------------------------------------------------
'
Public Sub ReloadRibbon()
    modCOMAddIn.ReloadRibbon
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RepairColors
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapply the color properties in the current database
'---------------------------------------------------------------------------------------
'
Public Sub RepairColors()
    If HasFormOpen Then Exit Sub
    RepairColorDefinitionBlocks
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SplitFiles
' Author    : Adam Waller
' Date      : 11/14/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SplitFiles()
    If HasFormOpen Then Exit Sub
    DoCmd.OpenForm "frmVCSSplitFiles"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LocalizeLibraryReferences
' Author    : Adam Waller
' Date      : 5/4/2022
' Purpose   : Localize Access database libraries to files in same folder as the
'           : current database.
'---------------------------------------------------------------------------------------
'
Public Sub LocalizeLibraryReferences()
    If HasFormOpen Then Exit Sub
    modLibReference.LocalizeLibraryReferences True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Version
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Return the version of this (CodeDB) instance of VCS.
'---------------------------------------------------------------------------------------
'
Public Property Get Version() As String
    Version = GetVCSVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : GoToLink...
' Author    : Adam Waller
' Date      : 3/4/2022
' Purpose   : Wrapper functions to visit online links (called from Ribbon menu)
'---------------------------------------------------------------------------------------
'
Public Sub GoToLinkAdvancedTools()
    ' TODO: Create page describing advanced tools.
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki"
End Sub
Public Sub GoToLinkHome()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin"
End Sub
Public Sub GoToLinkDocumentation()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki"
End Sub
Public Sub GoToLinkSupport()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/issues"
End Sub
Public Sub GoToLinkDownload()
    FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/releases"
End Sub
Public Sub ActivateHook()
    If HasFormOpen Then Exit Sub
    If modExportOnSaveHook.ActivateHook = False Then
        MsgBox "Error activating the hook", vbCritical, "Unable to activate hook"
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Save the current state so we can restore it after the current operation.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    SaveState
    ' When the class is initialized, make sure the ribbon is active (if installed).
    ' This way if the COM add-in is not active, it will be automatically activated
    ' when the add-in is opened from the [Database Tools\Add-ins] menu. (See #451)
    If GetInstallSettings.blnUseRibbonAddIn Then modCOMAddIn.VerifyRibbon
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Release any outstanding objects and restore the session state
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    modObjects.ReleaseObjects
    RestoreState
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveState
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Save any user settings before running our code.
'---------------------------------------------------------------------------------------
'
Private Sub SaveState()

    Dim strValue As String

    ' Just in case the user terminated the code before the settings could be restored,
    ' check for any existing user settings that might need to be restored.
    If ReadSetting("Save Time") <> vbNullString Then RestoreState

    ' Save a value so we can delete this section later
    PreserveSetting "Save Time", Now()

    ' Error trapping setting. (We need this to "Break in Class Modules" for this add-in)
    strValue = Application.GetOption("Error Trapping")
    If strValue <> "1" Then
        PreserveSetting "Error Trapping", strValue
        Application.SetOption "Error Trapping", 1
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestoreState
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Restore user settings.
'---------------------------------------------------------------------------------------
'
Private Sub RestoreState()

    Dim strValue As String

    strValue = ReadSetting("Error Trapping")
    If strValue <> vbNullString Then
        If Application.GetOption("Error Trapping") <> strValue Then
            Application.SetOption "Error Trapping", CLng(strValue)
        End If
    End If

    ' Remove temporary registry section
    DeleteSetting PROJECT_NAME, "State"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreserveSetting
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Wrapper functions for saving and reading session state settings.
'---------------------------------------------------------------------------------------
'
Private Sub PreserveSetting(strName As String, strSetting As String)
    SaveSetting PROJECT_NAME, "State", strName, strSetting
End Sub
Private Function ReadSetting(strName As String, Optional strDefault As String) As String
    ReadSetting = GetSetting(PROJECT_NAME, "State", strName, strDefault)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasFormOpen
' Author    : Adam Waller
' Date      : 3/29/2024
' Purpose   : Make sure no other forms are currently loaded before initiating a new
'           : operation. (Force linear usage of one operation at a time.)
'           : Also warns user that they must close the form before they can proceed.
'
'---------------------------------------------------------------------------------------
'
Private Function HasFormOpen(Optional blnWarnUser As Boolean = True) As Boolean

    Dim objForm As AccessObject
    Dim frm As Access.Form
    Dim strCaption As String

    ' Cache collection of form names used by this add-in
    Static dFormNames As Dictionary

    ' Build cached list of form names
    If dFormNames Is Nothing Then
        Set dFormNames = New Dictionary
        For Each objForm In CodeProject.AllForms
            dFormNames.Add objForm.Name, vbNullString
        Next objForm
    End If

    ' Check for open forms (from current database, add-in, or library database)
    For Each frm In Application.Forms
        If dFormNames.Exists(frm.Name) Then
            ' Found a matching form name.
            If frm.CurrentView = acCurViewDesign Then
                ' Design view is okay, since the add-in won't have a form in design
                ' view when running as an add-in. (No conflict)
            Else
                ' A form with the same name. This could be a problem.
                If blnWarnUser Then
                    If frm.Caption = vbNullString Then
                        strCaption = frm.Name
                    Else
                        ' Users may recognize caption better than name
                        strCaption = frm.Caption & " (" & frm.Name & ")"
                    End If
                    MsgBox2 "Add-in Form Already Open", _
                        "Please close '" & strCaption & "' before running this action.", , vbInformation
                    HasFormOpen = True
                    Exit Function
                End If
            End If
        End If
    Next

End Function
