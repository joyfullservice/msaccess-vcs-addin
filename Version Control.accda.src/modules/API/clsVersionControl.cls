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
'@Folder("API")


'---------------------------------------------------------------------------------------
' Procedure : Show
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Show the Version Control System main form
'---------------------------------------------------------------------------------------
'
Public Sub Show()
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
    If Operation.Begin(eotOther) Then
        If DatabaseFileOpen Then
            ' Force reload of options from current project before opening the form.
            Set modObjects.Options = Nothing
            DoCmd.OpenForm "frmVCSOptions"
        Else
            ' The table data subform may crash Access when the options form is closed with no database open.
            MsgBox2 "No Database Open", "Please open a database file before configuring VCS options.", , vbInformation
        End If
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
    RunExport ecfAllObjects
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FullExport
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Force a full export of all objects (ignoring Fast Save)
'---------------------------------------------------------------------------------------
'
Public Sub FullExport()
    RunExport ecfAllObjects, , True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVBA
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Export just the VBA related components
'---------------------------------------------------------------------------------------
'
Public Sub ExportVBA()
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

    Set objSelected = GetSelectedNavPaneObject
    If objSelected Is Nothing Then
        MsgBox2 "Please Select an Object First", _
            "Select a single object in the Navigation Pane to export (this item must have the keyboard focus).", _
            "If you are using a non-English version of Access and it is still not working, please open an issue on GitHub to add support for your language.", vbInformation
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
Private Sub RunExport(Optional intFilter As eContainerFilter = ecfAllObjects, Optional objItem As AccessObject, Optional blnFullExport As Boolean = False)
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    With Form_frmVCSMain
        If objItem Is Nothing Then
            .intContainerFilter = intFilter
        Else
            Set .objSingleObject = objItem
        End If
        ' Set full export checkbox if requested
        If blnFullExport Then .chkFullExport = True
        .Visible = True
        .cmdExport_Click
        If Operation.ErrorLevel < eelError Then .AutoClose
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Build
' Author    : Adam Waller
' Date      : 3/30/2022
' Purpose   : Initiate a full build from source
'---------------------------------------------------------------------------------------
'
Public Sub Build(Optional ByVal strSourceFolder As String)
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
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    With Form_frmVCSMain
        ' See if merge build is available
        If .chkFullBuild.Enabled Then
            ' Initiate the merge build
            .chkFullBuild = False
            .cmdBuild_Click
        Else
            MsgBox2 T("Merge Build Not Available"), _
                T("A full build or full export must be performed before you can merge from source."), _
                T("Please perform a full build or full export of this database first."), vbInformation
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

    If Not Operation.Begin(eotBuild) Then Exit Sub

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
                    Operation.Finish eorCanceled
                    Exit Sub
                End If
            Else
                ' Canceled dialog
                Operation.Finish eorCanceled
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
        modBuild.Build strSourceFolder, True, , strNewDbPath
    Else
        Operation.Finish eorCanceled
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
    If Operation.Begin(eotMerge) Then modBuild.MergeAllSource
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

    If Not Operation.Begin(eotMerge) Then Exit Sub

    Set objSelected = GetSelectedNavPaneObject
    If objSelected Is Nothing Then
        MsgBox2 "Please Select an Object First", _
            "Select a single object in the Navigation Pane to reload from source (this item must have the keyboard focus).", _
            "If you are using a non-English version of Access and it is still not working, please open an issue on GitHub to add support for your language.", vbInformation
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
            ' See discussion in #580 regarding command to use here
            ShellEx "TortoiseGitProc.exe", "/command:repostatus /closeonend:0 /path:" & strFolder

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
    OpenLog "Export"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenBuildLog
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open the build log file
'---------------------------------------------------------------------------------------
'
Public Sub OpenBuildLog()
    OpenLog "Build"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenLog
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Open the most recent log file matching the base name (e.g., "Export")
'---------------------------------------------------------------------------------------
'
Private Sub OpenLog(strBaseName As String)

    Dim strFolder As String
    Dim strPath As String
    Dim oFolder As Object
    Dim oFile As Object
    Dim strNewest As String

    ' Look in the logs subfolder
    strFolder = FSO.BuildPath(Options.GetExportFolder, "logs")

    If FSO.FolderExists(strFolder) Then
        Set oFolder = FSO.GetFolder(strFolder)
        ' Find the most recent file matching pattern (alphabetical = chronological with timestamp naming)
        For Each oFile In oFolder.Files
            If oFile.Name Like strBaseName & "_*.log" Then
                If oFile.Path > strNewest Then strNewest = oFile.Path
            End If
        Next oFile
    End If

    If Len(strNewest) > 0 Then
        ' Note that the parentheses around the path are required for this to work.
        CreateObject("Shell.Application").Open (strNewest)
    Else
        strPath = FSO.BuildPath(strFolder, strBaseName & "_*.log")
        MsgBox2 "Log File Not Found", "Could not find any log files matching:", strPath, vbExclamation
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
' Procedure : RebuildAddIn
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Rebuild the VCS add-in from source files and install it.
'           : Access will close during this process.
'           : Source path is resolved from (1) current database, (2) registry, or
'           : (3) a folder picker dialog.
'---------------------------------------------------------------------------------------
'
Public Sub RebuildAddIn()

    Dim strPath As String

    ' Try to derive source folder from the current project
    If StrComp(FSO.GetBaseName(CurrentProject.Name), ADDIN_BASENAME, vbTextCompare) = 0 Then
        strPath = CurrentProject.Path & PathSep & CurrentProject.Name & ".src" & PathSep
    End If

    ' Fall back to the registry (saved during previous builds)
    If Len(strPath) = 0 Then
        strPath = GetSetting(PROJECT_NAME, "Install", "Source Path", vbNullString)
    End If

    ' Last resort: ask the user to pick the source folder
    If Len(strPath) = 0 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = T("Select Add-In Source Folder")
            .ButtonName = T("Select")
            If .Show = True Then
                strPath = StripSlash(.SelectedItems(1)) & PathSep
            Else
                Exit Sub
            End If
        End With
    End If

    ' Validate source folder exists
    If Not FSO.FolderExists(strPath) Then
        MsgBox2 T("Source Folder Not Found"), _
            T("The source folder was not found:") & vbCrLf & strPath, _
            , vbExclamation
        Exit Sub
    End If
    If Not FSO.FileExists(strPath & "vcs-options.json") Then
        MsgBox2 T("Invalid Source Folder"), _
            T("The source folder does not contain the expected VCS source files.") & vbCrLf & strPath, _
            , vbExclamation
        Exit Sub
    End If

    ' Verify resource files are available
    VerifyResources

    ' Launch the worker to handle the build and install disconnected
    Worker.Run_BuildAndInstall strPath

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RepairColors
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapply the color properties in the current database
'---------------------------------------------------------------------------------------
'
Public Sub RepairColors()
    If Operation.Begin(eotOther) Then
        RepairColorDefinitionBlocks
        Operation.Finish
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SplitFiles
' Author    : Adam Waller
' Date      : 11/14/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SplitFiles()
    If Operation.Begin(eotOther) Then DoCmd.OpenForm "frmVCSSplitFiles"
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
    If Operation.Begin(eotOther) Then
        modLibReference.LocalizeLibraryReferences True
        Operation.Finish
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ActivateHook
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Activates the hook for automatic export on save.
'---------------------------------------------------------------------------------------
'
Public Sub ActivateHook()
    If Operation.Begin(eotOther) Then
        If modExportOnSaveHook.ActivateHook = False Then
            MsgBox "Error activating the hook", vbCritical, "Unable to activate hook"
        End If
        Operation.Finish
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StandardizeLetterCasing
' Author    : Adam Waller
' Date      : 3/18/2026
' Purpose   : Apply standard letter casing rules to VBA identifiers.
'           : Requires a clsStandardLetterCasing module in the current project.
'           : If the module is not found, offers to create a starter template.
'---------------------------------------------------------------------------------------
'
Public Sub StandardizeLetterCasing()

    Dim colResult As Collection
    Dim colIssues As New Collection
    Dim varChange As Variant
    Dim strDetails As String
    Dim strIssueDetails As String
    Dim lngShown As Long
    Const MaxDetailLines As Long = 10

    If Not Operation.Begin(eotOther) Then Exit Sub

    Set colResult = modLetterCasing.StandardizeLetterCasing(colIssues)

    If colResult Is Nothing Then
        If vbYes = MsgBox2(T("Module Not Found"), _
            T("The clsStandardLetterCasing class module was not found in the current project."), _
            T("Would you like to create a starter template?"), _
            vbYesNo + vbQuestion) Then
            CreateLetterCasingTemplate
        End If
    Else
        If colResult.Count > 0 Then
            lngShown = 0
            For Each varChange In colResult
                lngShown = lngShown + 1
                If lngShown <= MaxDetailLines Then
                    If Len(strDetails) > 0 Then strDetails = strDetails & vbCrLf
                    strDetails = strDetails & varChange
                End If
            Next varChange
            If colResult.Count > MaxDetailLines Then
                strDetails = strDetails & vbCrLf & T("... and {0} more (see export log for full list)", _
                    var0:=colResult.Count - MaxDetailLines)
            End If
            MsgBox2 T("Letter Casing Standardized"), _
                T("{0} correction(s) applied to VBA identifiers.", var0:=colResult.Count), _
                strDetails, vbInformation
        End If
        If colIssues.Count > 0 Then
            lngShown = 0
            For Each varChange In colIssues
                lngShown = lngShown + 1
                If lngShown <= MaxDetailLines Then
                    If Len(strIssueDetails) > 0 Then strIssueDetails = strIssueDetails & vbCrLf
                    strIssueDetails = strIssueDetails & varChange
                End If
            Next varChange
            If colIssues.Count > MaxDetailLines Then
                strIssueDetails = strIssueDetails & vbCrLf & T("... and {0} more (see export log for full list)", _
                    var0:=colIssues.Count - MaxDetailLines)
            End If
            MsgBox2 T("Invalid Letter Casing Entries"), _
                T("{0} invalid entry/entries found in clsStandardLetterCasing. Correct the module and run again.", _
                    var0:=colIssues.Count), _
                strIssueDetails, vbExclamation
        ElseIf colResult.Count = 0 Then
            MsgBox2 T("Letter Casing Consistent"), _
                T("All VBA identifiers already match the standard letter casing rules."), _
                , vbInformation
        End If
    End If

    Operation.Finish

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateLetterCasingTemplate
' Author    : Adam Waller
' Date      : 3/18/2026
' Purpose   : Create a starter clsStandardLetterCasing module in the current project
'           : and open it in the VBE for editing.
'---------------------------------------------------------------------------------------
'
Private Sub CreateLetterCasingTemplate()

    Dim cmp As VBIDE.VBComponent
    Dim strCode As String

    Set cmp = CurrentVBProject.VBComponents.Add(vbext_ct_ClassModule)
    cmp.Name = "clsStandardLetterCasing"

    strCode = "'---------------------------------------------------------------------------------------" & vbCrLf & _
              "' Module    : clsStandardLetterCasing" & vbCrLf & _
              "' Purpose   : Enforce consistent letter casing for VBA identifiers." & vbCrLf & _
              "'           : Add one Dim line per identifier: Dim IdentifierName 'IdentifierName" & vbCrLf & _
              "'           : The comment holds the canonical casing. Run Standardize Letter Casing" & vbCrLf & _
              "'           : from the ribbon to apply, or it runs automatically on export and build." & vbCrLf & _
              "'           : See: https://nolongerset.com/standardizelettercasing/" & vbCrLf & _
              "'---------------------------------------------------------------------------------------" & vbCrLf & _
              "Option Compare Database" & vbCrLf & _
              "Option Explicit" & vbCrLf & _
              vbCrLf & _
              "' Add identifiers below (one per line):" & vbCrLf & _
              "Dim i 'i" & vbCrLf & _
              "Dim strSql 'strSql"

    cmp.CodeModule.InsertLines 1, strCode

    MsgBox2 T("Template Created"), _
        T("The clsStandardLetterCasing class module has been added to your project."), _
        T("Add Dim lines for identifiers whose casing you want to standardize, then run this command again."), _
        vbInformation

    ' Save the module in the Access database components before opening
    DoCmd.Save acModule, cmp.Name
    DoEvents
    DoCmd.OpenModule "clsStandardLetterCasing"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Export a single named object by type and name.
'           : Accepts any type string recognized by ResolveComponentType (singular,
'           : plural, or alias forms). For core Access objects the name is required;
'           : for single-file component types the name is ignored.
'           : Returns JSON: {"success":true/false, "logPath":"...", "error":"..."}
'           : Public method for external API access (e.g., via MCP tools).
'---------------------------------------------------------------------------------------
'
Public Function ExportObject(ByVal strObjectType As String, Optional ByVal strObjectName As String) As String

    Dim intComponentType As eDatabaseComponentType
    Dim objItem As AccessObject
    Dim cComponent As IDbComponent
    Dim cItem As IDbComponent
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim blnFound As Boolean
    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    ' Resolve type string to eDatabaseComponentType via the shared helper
    intComponentType = ResolveComponentType(strObjectType)
    If intComponentType = -1 Then
        dResult.Add "success", False
        dResult.Add "error", "Unknown object type: " & strObjectType & _
            ". See ResolveComponentType in modContainers for accepted values."
        ExportObject = ConvertToJson(dResult)
        Exit Function
    End If

    ' Branch based on whether this is a core AccessObject type
    Select Case intComponentType
        Case edbQuery, edbForm, edbReport, edbModule, edbTableDef, edbMacro
            ' --- Core AccessObject path (uses ExportSingleObject) ---
            If Len(strObjectName) = 0 Then
                dResult.Add "success", False
                dResult.Add "error", "Object name is required for type: " & strObjectType
                ExportObject = ConvertToJson(dResult)
                Exit Function
            End If

            On Error Resume Next
            Select Case intComponentType
                Case edbQuery:      Set objItem = CurrentData.AllQueries(strObjectName)
                Case edbForm:       Set objItem = CurrentProject.AllForms(strObjectName)
                Case edbReport:     Set objItem = CurrentProject.AllReports(strObjectName)
                Case edbModule:     Set objItem = CurrentProject.AllModules(strObjectName)
                Case edbTableDef:   Set objItem = CurrentData.AllTables(strObjectName)
                Case edbMacro:      Set objItem = CurrentProject.AllMacros(strObjectName)
            End Select
            On Error GoTo ErrHandler

            If objItem Is Nothing Then
                dResult.Add "success", False
                dResult.Add "error", "Object not found: " & strObjectType & " '" & strObjectName & "'"
                ExportObject = ConvertToJson(dResult)
                Exit Function
            End If

            ' Begin operation and delegate to ExportSingleObject
            If Not Operation.Begin(eotExport) Then
                dResult.Add "success", False
                dResult.Add "error", "Could not begin export operation (another operation may be running)"
                ExportObject = ConvertToJson(dResult)
                Exit Function
            End If

            modExport.ExportSingleObject objItem, , _
                blnNoIndex:=(Operation.Source = eosMCPTool) Or (Operation.Source = eosExternalAPI)
            Set objItem = Nothing

        Case Else
            ' --- Non-core component type path ---
            If Not Operation.Begin(eotExport) Then
                dResult.Add "success", False
                dResult.Add "error", "Could not begin export operation (another operation may be running)"
                ExportObject = ConvertToJson(dResult)
                Exit Function
            End If

            ' Set up options and logging (mirrors ExportSingleObject setup)
            If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

            Dim blnNoIndex As Boolean
            Dim cDisabledIndex As clsVCSIndex
            blnNoIndex = (Operation.Source = eosMCPTool) Or (Operation.Source = eosExternalAPI)

            If blnNoIndex Then
                ' Skip index persistence for direct agent/API exports, but still
                ' reload options below so session overrides are honored.
                Set cDisabledIndex = New clsVCSIndex
                cDisabledIndex.Disabled = True
                Set VCSIndex = cDisabledIndex
            Else
                Set VCSIndex = Nothing
            End If
            Set Options = Nothing
            Options.LoadProjectOptions
            If Operation.Source = eosMCPTool Or Operation.Source = eosExternalAPI Then
                Options.LoadOptionOverrides
            End If
            Log.Clear
            Log.SourcePath = Options.GetExportFolder
            Log.Active = True
            Perf.StartTiming

            If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

            With Log
                .Spacer
                .Add T("Beginning Export of Component"), False
                .Add CurrentProject.Name
                .Add T("VCS Version {0}", var0:=GetVCSVersion)
                .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
                .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
                If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
                .Add Now
                .Spacer
            End With

            Set cComponent = GetComponentClass(intComponentType)
            If cComponent Is Nothing Then
                dResult.Add "success", False
                dResult.Add "error", "Could not create component class for type: " & strObjectType
                GoTo ExportCleanUp
            End If

            If cComponent.SingleFile Then
                ' Single-file type: export all items (name is irrelevant)
                Log.Add T("Exporting {0}...", var0:=cComponent.Category)
                Set dItems = cComponent.GetAllFromDB()
                For Each varKey In dItems.Keys
                    Set cItem = dItems(varKey)
                    cItem.Export
                Next varKey
            Else
                ' Multi-file type: find the named item
                If Len(strObjectName) = 0 Then
                    dResult.Add "success", False
                    dResult.Add "error", "Object name is required for type: " & strObjectType
                    GoTo ExportCleanUp
                End If

                Log.Add T("Exporting {0}: {1}...", var0:=cComponent.Category, var1:=strObjectName)
                Set dItems = cComponent.GetAllFromDB()
                blnFound = False
                For Each varKey In dItems.Keys
                    Set cItem = dItems(varKey)
                    If StrComp(FSO.GetBaseName(CStr(varKey)), strObjectName, vbTextCompare) = 0 Then
                        cItem.Export
                        blnFound = True
                        Exit For
                    End If
                Next varKey

                If Not blnFound Then
                    dResult.Add "success", False
                    dResult.Add "error", "Object not found: " & strObjectType & " '" & strObjectName & "'"
                    GoTo ExportCleanUp
                End If
            End If

            ' Export AGENTS.md
            modResource.ExtractResource "AGENTS.md", Options.GetExportFolder

            Log.Spacer
            Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

ExportCleanUp:
            Perf.EndTiming
            With Log
                .Add vbNewLine & Perf.GetReports, False
                .SaveFile
                .Active = False
                .Flush
            End With
            If blnNoIndex Then
                Set VCSIndex = Nothing
            Else
                VCSIndex.Save
            End If

    End Select

    ' Build result if not already set by an error path above
    If Not dResult.Exists("success") Then
        dResult.Add "success", (Operation.ErrorLevel < eelError)
        If Operation.ErrorLevel >= eelError Then
            dResult.Add "error", "Export completed with errors. Check the log for details."
        End If
        dResult.Add "logPath", Log.LogFilePath
    End If

    Operation.Finish
    ExportObject = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    ExportObject = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportObject
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Import (load) a single named object from source files.
'           : Accepts any type string recognized by ResolveComponentType (singular,
'           : plural, or alias forms). For single-file component types the name is
'           : ignored and the entire category is imported.
'           : Returns JSON: {"success":true/false, "error":"..."}
'           : Public method for external API access (e.g., via MCP tools).
'---------------------------------------------------------------------------------------
'
Public Function ImportObject(ByVal strObjectType As String, Optional ByVal strObjectName As String) As String

    Dim intComponentType As eDatabaseComponentType
    Dim cComponent As IDbComponent
    Dim cItem As IDbComponent
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim strSourceFile As String
    Dim blnFound As Boolean
    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    ' Resolve type string to eDatabaseComponentType via the shared helper
    intComponentType = ResolveComponentType(strObjectType)
    If intComponentType = -1 Then
        dResult.Add "success", False
        dResult.Add "error", "Unknown object type: " & strObjectType & _
            ". See ResolveComponentType in modContainers for accepted values."
        ImportObject = ConvertToJson(dResult)
        Exit Function
    End If

    ' Create the component class
    Set cComponent = GetComponentClass(intComponentType)
    If cComponent Is Nothing Then
        dResult.Add "success", False
        dResult.Add "error", "Could not create component class for type: " & strObjectType
        ImportObject = ConvertToJson(dResult)
        Exit Function
    End If

    ' Determine the source file path based on component type
    If cComponent.SingleFile Then
        ' Single-file type: use the component's fixed source file path
        strSourceFile = cComponent.SourceFile
        If Len(strSourceFile) = 0 Or Not FSO.FileExists(strSourceFile) Then
            dResult.Add "success", False
            dResult.Add "error", "Source file not found for " & strObjectType
            If Len(strSourceFile) > 0 Then dResult.Add "expectedPath", strSourceFile
            ImportObject = ConvertToJson(dResult)
            Exit Function
        End If
        ' Use the category as the display name for single-file types
        If Len(strObjectName) = 0 Then strObjectName = cComponent.Category
    Else
        ' Multi-file type: need the object name
        If Len(strObjectName) = 0 Then
            dResult.Add "success", False
            dResult.Add "error", "Object name is required for type: " & strObjectType
            ImportObject = ConvertToJson(dResult)
            Exit Function
        End If

        ' For core AccessObject types, try to resolve the database object
        Select Case intComponentType
            Case edbQuery, edbForm, edbReport, edbModule, edbTableDef, edbMacro
                Dim objItem As AccessObject
                On Error Resume Next
                Select Case intComponentType
                    Case edbQuery:      Set objItem = CurrentData.AllQueries(strObjectName)
                    Case edbForm:       Set objItem = CurrentProject.AllForms(strObjectName)
                    Case edbReport:     Set objItem = CurrentProject.AllReports(strObjectName)
                    Case edbModule:     Set objItem = CurrentProject.AllModules(strObjectName)
                    Case edbTableDef:   Set objItem = CurrentData.AllTables(strObjectName)
                    Case edbMacro:      Set objItem = CurrentProject.AllMacros(strObjectName)
                End Select
                On Error GoTo ErrHandler

                If Not objItem Is Nothing Then
                    Set cComponent.DbObject = objItem
                    strSourceFile = cComponent.SourceFile
                    Set objItem = Nothing
                End If

            Case Else
                ' For non-core multi-file types, search GetAllFromDB by name
                Set dItems = cComponent.GetAllFromDB()
                For Each varKey In dItems.Keys
                    Set cItem = dItems(varKey)
                    If StrComp(FSO.GetBaseName(CStr(varKey)), strObjectName, vbTextCompare) = 0 Then
                        strSourceFile = CStr(varKey)
                        Set cComponent = cItem
                        Exit For
                    End If
                Next varKey
        End Select

        ' If we couldn't get the source file, try searching by file name
        If Len(strSourceFile) = 0 Then
            strSourceFile = FindSourceFile(cComponent, strObjectName)
        End If

        If Len(strSourceFile) = 0 Or Not FSO.FileExists(strSourceFile) Then
            dResult.Add "success", False
            dResult.Add "error", "Source file not found for " & strObjectType & " '" & strObjectName & "'"
            If Len(strSourceFile) > 0 Then dResult.Add "expectedPath", strSourceFile
            ImportObject = ConvertToJson(dResult)
            Exit Function
        End If
    End If

    ' Begin operation and delegate to LoadSingleObject
    If Not Operation.Begin(eotMerge) Then
        dResult.Add "success", False
        dResult.Add "error", "Could not begin import operation (another operation may be running)"
        ImportObject = ConvertToJson(dResult)
        Exit Function
    End If

    LoadSingleObject cComponent, strObjectName, strSourceFile, _
        blnNoIndex:=(Operation.Source = eosMCPTool) Or (Operation.Source = eosExternalAPI)

    dResult.Add "success", (Operation.ErrorLevel < eelError)
    If Operation.ErrorLevel >= eelError Then
        dResult.Add "error", "Import completed with errors. Check the log for details."
    End If
    dResult.Add "logPath", Log.LogFilePath
    ImportObject = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    ImportObject = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportByType
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Category-scoped export: sync one or more whole component categories
'           : from the database to source files. Reconciles deletions (orphaned source
'           : files) within the named categories. Conflicts are auto-resolved under
'           : MCP/API or prompted interactively.
'           :
'           : This is a middle tier between ExportObject (one named object, no orphan
'           : cleanup) and Export/FullExport (entire project with global front-matter).
'           :
'           : varTypes accepts a single eDatabaseComponentType enum, a string alias
'           : (anything ResolveComponentType accepts), or an Array of either.
'           : blnFullExport=False (default) re-exports only changed objects;
'           : True re-exports all objects. Both modes reconcile deletions.
'           :
'           : Returns JSON: {"success":true/false, "logPath":"...", "error":"..."}
'---------------------------------------------------------------------------------------
'
Public Function ExportByType(ByVal varTypes As Variant, _
    Optional ByVal blnFullExport As Boolean = False) As String

    Dim dResult As Dictionary
    Dim colContainers As Collection
    Dim strError As String

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    If Not DatabaseFileOpen Then
        dResult.Add "success", False
        dResult.Add "error", "No database is currently open."
        ExportByType = ConvertToJson(dResult)
        Exit Function
    End If

    Set colContainers = GetContainersForTypes(varTypes, strError)
    If Len(strError) > 0 Then
        dResult.Add "success", False
        dResult.Add "error", strError
        ExportByType = ConvertToJson(dResult)
        Exit Function
    End If

    If Not ScopedExportPreflight(strError) Then
        dResult.Add "success", False
        dResult.Add "error", strError
        ExportByType = ConvertToJson(dResult)
        Exit Function
    End If

    If Not Operation.Begin(eotExport) Then
        dResult.Add "success", False
        dResult.Add "error", "Could not begin export operation (another operation may be running)"
        ExportByType = ConvertToJson(dResult)
        Exit Function
    End If

    BeginScopedSyncLog T("Beginning Category Export")
    modExport.ExportScoped colContainers, blnFullExport
    ExportByType = FinishScopedSync(dResult, colContainers, True)
    Exit Function

ErrHandler:
    ExportByType = ScopedSyncErrHandler(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportByType
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Category-scoped import (merge): merge source files for one or more
'           : whole component categories into the current database. Reconciles
'           : deletions (DB objects no longer represented in source) within the
'           : named categories. Conflicts are auto-resolved under MCP/API or prompted
'           : interactively.
'           :
'           : This is a middle tier between ImportObject (one named object, no orphan
'           : reconciliation) and MergeBuild (entire project with backup).
'           : No database backup is taken - the caller/pipeline owns backups.
'           :
'           : varTypes accepts a single eDatabaseComponentType enum, a string alias
'           : (anything ResolveComponentType accepts), or an Array of either.
'           : blnFullImport=False (default) merges only changed source files;
'           : True merges all source files. Both modes reconcile deletions.
'           :
'           : Returns JSON: {"success":true/false, "logPath":"...", "error":"..."}
'---------------------------------------------------------------------------------------
'
Public Function ImportByType(ByVal varTypes As Variant, _
    Optional ByVal blnFullImport As Boolean = False) As String

    Dim dResult As Dictionary
    Dim colContainers As Collection
    Dim strError As String

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    If Not DatabaseFileOpen Then
        dResult.Add "success", False
        dResult.Add "error", "No database is currently open."
        ImportByType = ConvertToJson(dResult)
        Exit Function
    End If

    Set colContainers = GetContainersForTypes(varTypes, strError, True)
    If Len(strError) > 0 Then
        dResult.Add "success", False
        dResult.Add "error", strError
        ImportByType = ConvertToJson(dResult)
        Exit Function
    End If

    If Not ScopedImportPreflight(strError) Then
        dResult.Add "success", False
        dResult.Add "error", strError
        ImportByType = ConvertToJson(dResult)
        Exit Function
    End If

    If Not Operation.Begin(eotMerge) Then
        dResult.Add "success", False
        dResult.Add "error", "Could not begin import operation (another operation may be running)"
        ImportByType = ConvertToJson(dResult)
        Exit Function
    End If

    BeginScopedSyncLog T("Beginning Category Import")
    modBuild.MergeScoped colContainers, blnFullImport
    ImportByType = FinishScopedSync(dResult, colContainers, False)
    Exit Function

ErrHandler:
    ImportByType = ScopedSyncErrHandler(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : BeginScopedSyncLog
' Purpose   : Shared setup for ExportByType / ImportByType.
'---------------------------------------------------------------------------------------
'
Private Sub BeginScopedSyncLog(strHeading As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set VCSIndex = Nothing
    Set Options = Nothing
    Options.LoadProjectOptions
    If Operation.Source = eosMCPTool Or Operation.Source = eosExternalAPI Then
        Options.LoadOptionOverrides
    End If
    Log.Clear
    Log.SourcePath = Options.GetExportFolder
    Log.Active = True
    Perf.StartTiming

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    With Log
        .Spacer
        .Add strHeading, False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FinishScopedSync
' Purpose   : Shared wrap-up for ExportByType / ImportByType.
'---------------------------------------------------------------------------------------
'
Private Function FinishScopedSync(dResult As Dictionary, colContainers As Collection, _
    blnExport As Boolean) As String

    Dim blnCanceled As Boolean
    Dim blnCritical As Boolean
    Dim blnUserCanceled As Boolean
    Dim intResult As eOperationResult
    Dim strColor As String

    blnUserCanceled = VCSIndex.Conflicts.UserCanceled
    blnCritical = (Operation.ErrorLevel = eelCritical)
    blnCanceled = blnCritical Or blnUserCanceled

    If blnCanceled Then
        strColor = "Red"
    ElseIf Operation.ErrorLevel >= eelError Then
        strColor = "red"
    Else
        strColor = "green"
    End If

    Log.Spacer
    If Not blnCanceled Then
        Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, strColor, True
    End If

    VCSIndex.ClearTempExportFolder
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
        .Flush
    End With

    If Not blnCanceled Then
        If blnExport Then modExport.UpdateScopedExportIndex colContainers
        VCSIndex.Save
    End If

    If blnUserCanceled Then
        intResult = eorCanceled
    ElseIf blnCritical Or Operation.ErrorLevel >= eelError Then
        intResult = eorFailed
    Else
        intResult = eorSuccess
    End If

    If Not dResult.Exists("success") Then
        dResult.Add "success", (intResult = eorSuccess)
        If intResult <> eorSuccess Then
            dResult.Add "error", IIf(blnUserCanceled, _
                "Operation was canceled.", _
                "Operation completed with errors. Check the log for details.")
        End If
        dResult.Add "logPath", Log.LogFilePath
    End If

    Operation.Finish intResult
    FinishScopedSync = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ScopedSyncErrHandler
' Purpose   : Ensure scoped sync operations always finish the Operation on error.
'---------------------------------------------------------------------------------------
'
Private Function ScopedSyncErrHandler(dResult As Dictionary) As String

    If Operation.Status = eosRunning Then
        Perf.EndTiming
        Log.Active = False
        Operation.Finish eorFailed
    End If

    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    ScopedSyncErrHandler = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindSourceFile
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Find the source file for a named object when the component doesn't
'           : already have a database object reference set.
'---------------------------------------------------------------------------------------
'
Private Function FindSourceFile(cComponent As IDbComponent, strObjectName As String) As String

    Dim strFolder As String
    Dim strFile As String

    ' Get the base folder for this component type
    strFolder = cComponent.BaseFolder

    ' Try common extensions used for this type of component
    Dim varExt As Variant
    Dim aExtensions As Variant

    Select Case TypeName(cComponent)
        Case "clsDbQuery":      aExtensions = Array(".sql", ".qdef", ".bas")
        Case "clsDbForm":       aExtensions = Array(".form", ".cls", ".bas")
        Case "clsDbReport":     aExtensions = Array(".report", ".cls", ".bas")
        Case "clsDbModule":     aExtensions = Array(".bas", ".cls")
        Case "clsDbTableDef":   aExtensions = Array(".xml", ".sql")
        Case "clsDbMacro":      aExtensions = Array(".macro", ".bas")
        Case Else:              aExtensions = Array(".bas", ".cls", ".xml", ".sql")
    End Select

    For Each varExt In aExtensions
        strFile = strFolder & strObjectName & CStr(varExt)
        If FSO.FileExists(strFile) Then
            FindSourceFile = strFile
            Exit Function
        End If
    Next varExt

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunVBA
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Execute agent-generated VBA code in a temporary module.
'           : Creates a temp module, wraps code in a function with error handling,
'           : compiles to validate, executes, captures result, cleans up.
'           : Returns JSON: {"success":true/false, "result":"...", "error":"...",
'           :   "compileError":"..."}
'           : Gated by McpAllowRunVBA option.
'---------------------------------------------------------------------------------------
'
Public Function RunVBA(ByVal strCode As String) As String

    Dim dResult As Dictionary
    Dim cmp As VBIDE.VBComponent
    Dim strModuleName As String
    Dim strWrapper As String
    Dim varReturn As Variant
    Dim blnModuleCreated As Boolean

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary
    strModuleName = "MCP_Temp_" & Format$(Timer * 100, "0")

    ' Check permission
    If Not Options.McpAllowRunVBA Then
        dResult.Add "success", False
        dResult.Add "error", "RunVBA is disabled. Enable McpAllowRunVBA in VCS Options > MCP to allow agent-generated code execution."
        RunVBA = ConvertToJson(dResult)
        Exit Function
    End If

    ' Build the wrapper function with error number/description stored in
    ' module-level vars. Function names must NOT start with an underscore --
    Dim strNumberedCode As String
    ' Inject sequential 1-based line numbers into the agent's code so Err.Erl
    ' can pinpoint the line that raised the runtime error. The number written
    ' on each line equals the 1-based ordinal of that line within strCode,
    ' which means an "errorLine" value of 7 maps directly to line 7 of what
    ' the caller submitted (the counter advances on every physical input
    ' line, including blanks/comments/continuations, even though only
    ' executable lines actually carry a number).
    strNumberedCode = AddVbaLineNumbers(strCode, "    ")
    ' Wrapper structure: capture user code's runtime error into module-level
    ' state, then Err.Clear before returning. Without the Err.Clear the
    ' raised error leaks out through Application.Run to the caller, which
    ' makes error 2517 (procedure not found) indistinguishable from a real
    ' user-raised runtime error in the diagnostics.
    strWrapper = "Private m_ErrNum As Long" & vbCrLf & _
                 "Private m_ErrDesc As String" & vbCrLf & _
                 "Private m_ErrLine As Long" & vbCrLf & vbCrLf & _
                 "Public Function MCP_TempFunction() As Variant" & vbCrLf & _
                 "    On Error Resume Next" & vbCrLf & _
                 strNumberedCode & vbCrLf & _
                 "    m_ErrNum = Err.Number" & vbCrLf & _
                 "    m_ErrDesc = Err.Description" & vbCrLf & _
                 "    m_ErrLine = Erl" & vbCrLf & _
                 "    Err.Clear" & vbCrLf & _
                 "End Function" & vbCrLf & vbCrLf & _
                 "Public Function MCP_GetErrNum() As Long" & vbCrLf & _
                 "    MCP_GetErrNum = m_ErrNum" & vbCrLf & _
                 "End Function" & vbCrLf & vbCrLf & _
                 "Public Function MCP_GetErrDesc() As String" & vbCrLf & _
                 "    MCP_GetErrDesc = m_ErrDesc" & vbCrLf & _
                 "End Function" & vbCrLf & vbCrLf & _
                 "Public Function MCP_GetErrLine() As Long" & vbCrLf & _
                 "    MCP_GetErrLine = m_ErrLine" & vbCrLf & _
                 "End Function"

    ' Create temp module in the current database. New modules in Access are
    ' usually pre-populated with "Option Compare Database" (and sometimes
    ' "Option Explicit" depending on VBE settings). InsertLines 1 prepends,
    ' which would leave duplicate Option statements at the bottom and trigger
    ' "Multiple Option Compare statements are not allowed". Clear the module
    ' first so our wrapper is the entire content.
    Set cmp = CurrentVBProject.VBComponents.Add(vbext_ct_StdModule)
    cmp.Name = strModuleName
    blnModuleCreated = True
    If cmp.CodeModule.CountOfLines > 0 Then
        cmp.CodeModule.DeleteLines 1, cmp.CodeModule.CountOfLines
    End If
    cmp.CodeModule.InsertLines 1, "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf & strWrapper

    ' Compile the project to validate the code. We must explicitly activate
    ' the host CurrentVBProject in the VBE first -- otherwise acCmdCompileAllModules
    ' compiles whichever project is currently active in the VBE (typically the
    ' add-in itself, which already compiles), leaving our just-inserted temp
    ' module uncompiled. Application.IsCompiled would then return False because
    ' the host has an uncompiled module, and we'd report a misleading
    ' "compile failed" with no details. Same pattern as CompileVBA below.
    On Error Resume Next
    Set VBE.ActiveVBProject = CurrentVBProject
    DoCmd.RunCommand acCmdCompileAllModules
    DoEvents

    If Not Application.IsCompiled Then
        ' Distinguish between pre-existing host compile issues and a wrapper
        ' that itself failed to compile by removing the temp module and
        ' re-checking the host's compile state.
        Dim strHostState As String
        strHostState = "unknown"
        If blnModuleCreated Then
            CurrentVBProject.VBComponents.Remove cmp
            blnModuleCreated = False
            DoCmd.RunCommand acCmdCompileAllModules
            DoEvents
            If Application.IsCompiled Then
                strHostState = "host compiles cleanly without the wrapper -- the failure is in the agent code"
            Else
                strHostState = "host project does not compile on its own -- fix host compile errors before retrying RunVBA"
            End If
        End If
        dResult.Add "success", False
        dResult.Add "compileError", "VBA compilation failed. " & strHostState & "."
        dResult.Add "generatedSource", "Option Compare Database" & vbCrLf & _
            "Option Explicit" & vbCrLf & vbCrLf & strWrapper
        GoTo CleanUp
    End If
    On Error GoTo ErrHandler

    ' Execute the wrapper function. File-path qualifiers in Application.Run
    ' only work for referenced library projects (.accda add-ins), NOT for
    ' the current database's own project. When the temp module lives in the
    ' current database (the normal agent-calling-database case), we must use
    ' an unqualified call. When developing the add-in itself (the add-in IS
    ' the current database), the module is also in the current database, so
    ' unqualified still works. The file-path prefix is only needed when
    ' CurrentVBProject is a library reference (not the current database).
    Dim strProjPrefix As String
    Dim strCurrentDbPath As String
    strCurrentDbPath = GetUncPath(CurrentProject.FullName)
    If StrComp(CurrentVBProject.FileName, strCurrentDbPath, vbTextCompare) = 0 Then
        strProjPrefix = vbNullString
    Else
        strProjPrefix = FSO.GetParentFolderName(CurrentVBProject.FileName) & "\" & _
            FSO.GetBaseName(CurrentVBProject.FileName) & "."
    End If
    ResetCurrentVBProjectState
    varReturn = Application.Run(strProjPrefix & "MCP_TempFunction")

    ' Check for runtime errors captured by the wrapper (Err is cleared
    ' inside MCP_TempFunction before returning, so any error here is from
    ' Application.Run itself, not the user code).
    Dim lngErrNum As Long
    Dim strErrDesc As String
    Dim lngErrLine As Long
    lngErrNum = Application.Run(strProjPrefix & "MCP_GetErrNum")
    If lngErrNum <> 0 Then
        strErrDesc = Application.Run(strProjPrefix & "MCP_GetErrDesc")
        lngErrLine = Application.Run(strProjPrefix & "MCP_GetErrLine")
        dResult.Add "success", False
        dResult.Add "error", strErrDesc
        dResult.Add "errorNumber", lngErrNum
        ' Erl resolves to the most recent labeled line at the time the error
        ' was raised. Because every executable user line is auto-numbered
        ' above, this is normally the exact failing line. Omit when 0 (means
        ' the error fired before any numbered line ran, e.g. a setup-only
        ' script with no numbered statements).
        If lngErrLine > 0 Then dResult.Add "errorLine", lngErrLine
    Else
        dResult.Add "success", True
        If Not IsEmpty(varReturn) And Not IsNull(varReturn) Then
            dResult.Add "result", CStr(varReturn)
        Else
            dResult.Add "result", vbNullString
        End If
    End If

CleanUp:
    If blnModuleCreated Then
        On Error Resume Next
        ' Release our stored reference BEFORE removing the component.
        ' Holding a stored COM reference through Remove leaves a dangling
        ' pointer when the underlying VBComponent is destroyed; the
        ' implicit Release at function exit then dereferences freed memory.
        Set cmp = Nothing
        CurrentVBProject.VBComponents.Remove CurrentVBProject.VBComponents(strModuleName)
        On Error GoTo 0
    End If
    ResetCurrentVBProjectState
    RunVBA = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    GoTo CleanUp

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddVbaLineNumbers
' Author    : Adam Waller
' Date      : 4/20/2026
' Purpose   : Prepend sequential 1-based line numbers to each executable VBA
'           : statement in strCode so Err.Erl can identify the failing line
'           : when the code is executed inside the RunVBA wrapper.
'           : The number written on each line equals the 1-based ordinal of
'           : that line within strCode (the counter advances on every input
'           : line, blanks/comments/continuations included), so an Erl value
'           : of N maps directly to "line N of the original submission".
'           : Lines that cannot legally carry a line number are passed
'           : through untouched:
'           :   - blank lines
'           :   - pure-comment lines (first non-whitespace char is ')
'           :   - continuations of the prior physical line (previous line
'           :     ended with " _" or vbTab & "_")
'           :   - lines whose first non-whitespace char is already a digit
'           :     (caller already supplied a number; theirs wins)
'           : Input is split on vbLf (handles both LF and CRLF input);
'           : output uses vbCrLf and prefixes each line with strIndent so the
'           : result can be embedded directly inside an indented wrapper.
'---------------------------------------------------------------------------------------
'
Private Function AddVbaLineNumbers(ByVal strCode As String, Optional ByVal strIndent As String = "    ") As String

    Dim aLines() As String
    Dim lngIdx As Long
    Dim strLine As String
    Dim strTrimmed As String
    Dim strFirstChar As String
    Dim strCharBefore As String
    Dim blnPrevContinuation As Boolean
    Dim blnEndsWithContinuation As Boolean
    Dim strResult As String

    aLines = Split(strCode, vbLf)

    For lngIdx = LBound(aLines) To UBound(aLines)
        ' Strip a trailing CR if present so we cleanly handle CRLF input.
        strLine = aLines(lngIdx)
        If Len(strLine) > 0 Then
            If Right$(strLine, 1) = vbCr Then strLine = Left$(strLine, Len(strLine) - 1)
        End If
        strTrimmed = Trim$(strLine)

        ' Decide whether this physical line can carry a line number.
        If Len(strTrimmed) > 0 And Not blnPrevContinuation Then
            strFirstChar = Left$(strTrimmed, 1)
            If strFirstChar <> "'" And Not (strFirstChar Like "[0-9]") Then
                ' Inject the 1-based ordinal between the leading whitespace
                ' (added below via strIndent) and the trimmed code body.
                strLine = (lngIdx + 1) & " " & LTrim$(strLine)
            End If
        End If

        ' Determine whether THIS (trimmed) line ends with a continuation
        ' marker so the NEXT physical line is treated as a continuation.
        ' VBA's continuation marker is " _" (or tab + "_") at end of line;
        ' a bare trailing "_" without preceding whitespace is just an
        ' identifier character, not a line continuation.
        blnEndsWithContinuation = False
        If Len(strTrimmed) >= 2 Then
            If Right$(strTrimmed, 1) = "_" Then
                strCharBefore = Mid$(strTrimmed, Len(strTrimmed) - 1, 1)
                If strCharBefore = " " Or strCharBefore = vbTab Then
                    blnEndsWithContinuation = True
                End If
            End If
        End If
        blnPrevContinuation = blnEndsWithContinuation

        If lngIdx > LBound(aLines) Then strResult = strResult & vbCrLf
        strResult = strResult & strIndent & strLine
    Next lngIdx

    AddVbaLineNumbers = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExecuteSQL
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Execute a read-only SELECT query via CurrentDb.OpenRecordset.
'           : Returns JSON rows (field names + values), capped at lngMaxRows.
'           : Rejects non-SELECT statements.
'           : Public method for external API access (e.g., via MCP tools).
'---------------------------------------------------------------------------------------
'
Public Function ExecuteSQL(ByVal strSql As String, Optional ByVal lngMaxRows As Long = 100) As String

    Dim dResult As Dictionary
    Dim dRows As Collection
    Dim dRow As Dictionary
    Dim rst As DAO.Recordset
    Dim dbs As DAO.Database
    Dim lngRow As Long
    Dim lngField As Long

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    ' Check permission
    If Not Options.McpAllowExecuteSQL Then
        dResult.Add "success", False
        dResult.Add "error", "ExecuteSQL is disabled. Enable McpAllowExecuteSQL in VCS Options > MCP."
        ExecuteSQL = ConvertToJson(dResult)
        Exit Function
    End If

    ' Validate: only SELECT statements allowed
    Dim strTrimmed As String
    strTrimmed = Trim$(strSql)
    If LCase$(Left$(strTrimmed, 7)) <> "select " And LCase$(Left$(strTrimmed, 7)) <> "select" & vbTab Then
        dResult.Add "success", False
        dResult.Add "error", "Only SELECT statements are allowed. Received: " & Left$(strTrimmed, 50)
        ExecuteSQL = ConvertToJson(dResult)
        Exit Function
    End If

    ' Execute query
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)

    Set dRows = New Collection
    Do While Not rst.EOF And lngRow < lngMaxRows
        Set dRow = New Dictionary
        For lngField = 0 To rst.Fields.Count - 1
            If IsNull(rst.Fields(lngField).Value) Then
                dRow.Add rst.Fields(lngField).Name, Null
            Else
                dRow.Add rst.Fields(lngField).Name, rst.Fields(lngField).Value
            End If
        Next lngField
        dRows.Add dRow
        lngRow = lngRow + 1
        rst.MoveNext
    Loop

    dResult.Add "success", True
    dResult.Add "rows", dRows
    dResult.Add "rowCount", lngRow
    dResult.Add "truncated", (Not rst.EOF)

    rst.Close
    Set rst = Nothing
    Set dbs = Nothing

    ExecuteSQL = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    dResult.Add "errorNumber", CLng(Err.Number)
    ExecuteSQL = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOption
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Return the current value of any clsOptions property by name.
'           : Uses CallByName for dynamic property access.
'           : Public method for external API access (e.g., via MCP tools).
'---------------------------------------------------------------------------------------
'
Public Function GetOption(ByVal strName As String) As Variant

    LogUnhandledErrors
    On Error GoTo ErrHandler

    GetOption = CallByName(Options, strName, VbGet)
    Exit Function

ErrHandler:
    ' Return error info as JSON string when property doesn't exist
    Dim dResult As Dictionary
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", "Could not read option '" & strName & "': " & Err.Description
    GetOption = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetOption
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Set a clsOptions property by name. The change is applied immediately
'           : in memory AND persisted to a session-scoped override file in the mcp/
'           : subfolder so it survives Options reloads across operations.
'           : The user's vcs-options.json is never modified.
'---------------------------------------------------------------------------------------
'
Public Function SetOption(ByVal strName As String, ByVal varValue As Variant) As String

    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    ' Apply to current in-memory options
    CallByName Options, strName, VbLet, varValue

    ' Persist to session override file so the change survives Options reload
    SaveOptionOverride strName, varValue

    dResult.Add "success", True
    dResult.Add "option", strName
    dResult.Add "value", varValue
    SetOption = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", "Could not set option '" & strName & "': " & Err.Description
    SetOption = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionOverride
' Author    : Adam Waller
' Date      : 4/15/2026
' Purpose   : Write an option override to the session-specific file in the mcp/
'           : subfolder. Creates the subfolder and file if they don't exist.
'           : Uses the SessionId from modObjects (set via RegisterSession) or
'           : falls back to "default" for non-session API callers.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptionOverride(strName As String, varValue As Variant)

    Dim strFolder As String
    Dim strFile As String
    Dim strSessionId As String
    Dim dFile As Dictionary
    Dim dInfo As Dictionary
    Dim dOptions As Dictionary

    LogUnhandledErrors
    On Error Resume Next

    ' Determine session ID
    strSessionId = modObjects.SessionId
    If Len(strSessionId) = 0 Then strSessionId = "default"

    ' Build file path
    strFolder = Options.GetExportFolder & "mcp" & PathSep
    strFile = strFolder & "options-" & strSessionId & ".json"

    ' Create mcp subfolder if needed
    If Not FSO.FolderExists(strFolder) Then
        VerifyPath strFolder & "placeholder"
    End If

    ' Read existing file or create new dictionary
    If FSO.FileExists(strFile) Then
        Set dFile = ReadJsonFile(strFile)
    End If
    If dFile Is Nothing Then Set dFile = New Dictionary

    ' Update Info section
    If Not dFile.Exists("Info") Then
        Set dInfo = New Dictionary
        dInfo.Add "SessionId", strSessionId
        dInfo.Add "Source", IIf(Operation.Source = eosMCPTool, "MCP", "API")
        dInfo.Add "Created", Format$(Now, "yyyy-mm-dd hh:nn:ss")
        dInfo.Add "Modified", Format$(Now, "yyyy-mm-dd hh:nn:ss")
        dFile.Add "Info", dInfo
    Else
        dFile("Info")("Modified") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End If

    ' Update Options section
    If Not dFile.Exists("Options") Then
        Set dOptions = New Dictionary
        dFile.Add "Options", dOptions
    End If
    If dFile("Options").Exists(strName) Then
        dFile("Options")(strName) = varValue
    Else
        dFile("Options").Add strName, varValue
    End If

    ' Write file
    WriteFile ConvertToJson(dFile, 4), strFile

    CatchAny eelWarning, "Writing option override for " & strName, ModuleName(Me) & ".SaveOptionOverride"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RegisterSession
' Author    : Adam Waller
' Date      : 4/15/2026
' Purpose   : Register an MCP/API session identifier used to scope option overrides.
'           : The session ID determines the override filename (options-{id}.json in mcp/).
'           : Called by the MCP server at startup; idempotent.
'---------------------------------------------------------------------------------------
'
Public Function RegisterSession(ByVal strSessionId As String) As String

    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    modObjects.SessionId = strSessionId

    Set dResult = New Dictionary
    dResult.Add "success", True
    dResult.Add "session_id", strSessionId
    RegisterSession = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", "Could not register session: " & Err.Description
    RegisterSession = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : EndSession
' Author    : Adam Waller
' Date      : 4/15/2026
' Purpose   : End an MCP/API session by deleting its override file and reloading
'           : options without those overrides. If strSessionId is empty, uses the
'           : currently registered session.
'---------------------------------------------------------------------------------------
'
Public Function EndSession(Optional ByVal strSessionId As String) As String

    Dim dResult As Dictionary
    Dim strFile As String

    LogUnhandledErrors
    On Error GoTo ErrHandler

    If Len(strSessionId) = 0 Then strSessionId = modObjects.SessionId
    If Len(strSessionId) = 0 Then strSessionId = "default"

    ' Delete the session's override file
    strFile = Options.GetExportFolder & "mcp" & PathSep & "options-" & strSessionId & ".json"
    If FSO.FileExists(strFile) Then DeleteFile strFile

    ' Clear session ID if it matches
    If StrComp(strSessionId, modObjects.SessionId, vbTextCompare) = 0 Then
        modObjects.SessionId = vbNullString
    End If

    ' Reload options without this session's overrides
    Set Options = Nothing
    Options.LoadProjectOptions

    Set dResult = New Dictionary
    dResult.Add "success", True
    dResult.Add "session_id", strSessionId
    dResult.Add "overrides_cleared", True
    EndSession = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", "Could not end session: " & Err.Description
    EndSession = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLogContent
' Author    : Adam Waller
' Date      : 4/14/2026
' Purpose   : Read the most recent log file matching the base name (e.g., "Export").
'           : Returns the log file content as a string.
'           : Public method for external API access (e.g., via MCP tools).
'---------------------------------------------------------------------------------------
'
Public Function GetLogContent(Optional ByVal strBaseName As String = "Export") As String

    Dim strFolder As String
    Dim strPath As String
    Dim oFolder As Object
    Dim oFile As Object
    Dim strNewest As String
    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Set dResult = New Dictionary

    ' Look in the logs subfolder
    strFolder = FSO.BuildPath(Options.GetExportFolder, "logs")

    If FSO.FolderExists(strFolder) Then
        Set oFolder = FSO.GetFolder(strFolder)
        For Each oFile In oFolder.Files
            If oFile.Name Like strBaseName & "_*.log" Then
                If oFile.Path > strNewest Then strNewest = oFile.Path
            End If
        Next oFile
    End If

    If Len(strNewest) > 0 Then
        dResult.Add "success", True
        dResult.Add "logPath", strNewest
        dResult.Add "content", ReadFile(strNewest)
    Else
        dResult.Add "success", False
        dResult.Add "error", "No log files found matching '" & strBaseName & "_*.log' in " & strFolder
    End If

    GetLogContent = ConvertToJson(dResult)
    Exit Function

ErrHandler:
    Set dResult = New Dictionary
    dResult.Add "success", False
    dResult.Add "error", Err.Description
    GetLogContent = ConvertToJson(dResult)

End Function


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
' Procedure : GetVCSVersion
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Return the version of this (CodeDB) instance of VCS.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function GetVCSVersion() As String
    GetVCSVersion = modVCSUtility.GetVCSVersion
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetExportFolder
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Return the current export folder path for the project.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function GetExportFolder() As String
    GetExportFolder = Options.GetExportFolder
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDatabaseOpen
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Return whether a database file is currently open.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function IsDatabaseOpen() As Boolean
    IsDatabaseOpen = modDatabase.DatabaseFileOpen
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetProjectName
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Return the name of the currently open project/database.
'           : Returns empty string if no database is open.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function GetProjectName() As String
    If modDatabase.DatabaseFileOpen Then
        GetProjectName = CurrentProject.Name
    Else
        GetProjectName = vbNullString
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsVBACompiled
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Check if the VBA project is in a compiled state.
'           : Returns True if compiled, False otherwise.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function IsVBACompiled() As Boolean
    LogUnhandledErrors
    On Error Resume Next
    ' Application.IsCompiled returns True if project is compiled, False if not
    IsVBACompiled = Application.IsCompiled
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : CompileVBA
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Compile all VBA modules in the current database and return success status.
'           : Returns True if compilation succeeded (project is compiled), False if failed.
'           : Public method for external API access (e.g., via API function).
'---------------------------------------------------------------------------------------
'
Public Function CompileVBA(Optional ByVal blnSuppressWarnings As Boolean = False) As Boolean
    LogUnhandledErrors
    On Error GoTo ErrHandler

    Dim blnCompiledAfter As Boolean

    ' Suppress warnings if requested (we'll restore them in cleanup)
    If blnSuppressWarnings Then
        DoCmd.SetWarnings False
    End If

    ' Make sure we are running this in the CurrentDB, not the CodeDB
    Set VBE.ActiveVBProject = CurrentVBProject

    ' Compile all modules
    DoCmd.RunCommand acCmdCompileAllModules
    DoEvents

    ' Check if compilation succeeded by checking IsCompiled
    LogUnhandledErrors
    On Error Resume Next
    blnCompiledAfter = Application.IsCompiled
    On Error GoTo ErrHandler

    ' Restore warnings if we suppressed them
    If blnSuppressWarnings Then
        DoCmd.SetWarnings True
    End If

    ' Return True if compiled, False if not
    CompileVBA = blnCompiledAfter
    Exit Function

ErrHandler:
    ' Always restore warnings if we suppressed them, even on error
    If blnSuppressWarnings Then
        LogUnhandledErrors
        On Error Resume Next
        DoCmd.SetWarnings True
        On Error GoTo 0
    End If

    ' Return False on any error
    CompileVBA = False
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunRoundtripTests
' Author    : Adam Waller
' Date      : 4/24/2026
' Purpose   : Run the object round-trip regression harness against a fixture
'           : folder. Each fixture is imported into a sandbox object, exported,
'           : re-imported, and re-exported, with comparisons at every step.
'           :
'           : strFixtureFolder
'           :   - When empty (default), uses the corpus shipped with the add-in
'           :     at <addin>.src/../Testing/Fixtures/.
'           :   - Otherwise, the absolute path of a folder containing fixture
'           :     subfolders by component type (e.g. queries/select/, queries/
'           :     regression/) following the layout documented in
'           :     Wiki/Regression-Testing.md.
'           :
'           : blnRebaseline
'           :   - False (default): comparisons that mismatch are reported as
'           :     failures.
'           :   - True: mismatched fixture files are overwritten by the actual
'           :     export output (with name-bearing fields normalized back to
'           :     the canonical name). Use sparingly and review the diff in
'           :     git before committing.
'           :
'           : Returns a JSON string with this shape:
'           :   { "success": Boolean,
'           :     "fixtureFolder": String,
'           :     "scratchFolder": String,
'           :     "rebaseline": Boolean,
'           :     "logPath": String,
'           :     "stats": { "total","passed","failed","skipped","errors","elapsedSeconds" },
'           :     "results": [ { "fixture", "name", "type", "status", "checks":[...] }, ... ] }
'           :
'           : Live progress is also emitted through the standard Log singleton
'           : (visible in the main console form) and the per-session log file
'           : at <fixtureFolder>/logs/ObjectRoundtrip_<opId>.log.
'           :
'           : This is a developer/CI tool, not a user feature. It will refuse
'           : to run while another VCS operation (export/build/merge) is in
'           : progress, and it disables the change-tracking index for the
'           : duration of the run so test imports/exports do not pollute
'           : vcs-index.idx.
'---------------------------------------------------------------------------------------
'
Public Function RunRoundtripTests(Optional ByVal strFixtureFolder As String = vbNullString, _
    Optional ByVal blnRebaseline As Boolean = False) As String
    RunRoundtripTests = modTestRoundtrip.RunObjectRoundtripTests(strFixtureFolder, blnRebaseline)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidateQuerySqlBuilder
' Author    : Adam Waller
' Date      : 4/29/2026
' Purpose   : Run the query SQL builder validation tool against the current database.
'           : The validator compares fast MSysQueries-based SQL reconstruction against
'           : Access's QueryDefs.SQL text and writes diagnostic artifacts for review.
'           : Pass no arguments to validate all queries, or one/more query names to
'           : validate a targeted batch in a single log/artifact folder.
'---------------------------------------------------------------------------------------
'
Public Function ValidateQuerySqlBuilder(ParamArray varQueryNames() As Variant) As String
    ValidateQuerySqlBuilder = modTestQuerySqlBuilder.RunQuerySqlBuilderValidation(varQueryNames)
End Function


'---------------------------------------------------------------------------------------
' Procedure : InspectQuerySqlStorage
' Author    : Adam Waller
' Date      : 4/30/2026
' Purpose   : Developer diagnostic helper. Creates a temporary sandbox query from
'           : SQL, optionally asks Access to save it from Design View, reads the
'           : resulting MSysQueries rows, and deletes the sandbox query.
'---------------------------------------------------------------------------------------
'
Public Function InspectQuerySqlStorage(ByVal strSql As String, _
    Optional ByVal blnForceDesignView As Boolean = False) As String
    InspectQuerySqlStorage = modTestQuerySqlBuilder.InspectQuerySqlStorage(strSql, blnForceDesignView)
End Function


'---------------------------------------------------------------------------------------
' Procedure : CompareSqlBuilderArtifacts
' Author    : Adam Waller
' Date      : 4/30/2026
' Purpose   : Developer diagnostic helper. Compares validation artifact
'           : access.sql/generated.sql pairs through DAO row counts and schemas,
'           : writing a ledger beside the artifact folders.
'---------------------------------------------------------------------------------------
'
Public Function CompareSqlBuilderArtifacts(ByVal strArtifactRoot As String) As String
    CompareSqlBuilderArtifacts = modTestQuerySqlBuilder.CompareSqlBuilderArtifacts(strArtifactRoot)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HandleTestAssertion
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Assertion dispatcher called from modTestAssert.TestAssert via
'           : Application.Run. When a test run is active, records the assertion
'           : result and returns True. When no run is active, returns False so
'           : the caller falls back to Debug.Assert.
'---------------------------------------------------------------------------------------
'
Public Function HandleTestAssertion(ByVal blnCondition As Boolean, _
                                     ByVal varContext As Variant) As Boolean
    If TestRunner.State <> etrsRunning Then
        HandleTestAssertion = False
        Exit Function
    End If
    HandleTestAssertion = True
    TestRunner.RecordAssertion blnCondition, varContext
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunTests
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Scan for tests in CurrentVBProject and run them. Streams progress to
'           : the frmVCSMain log console. Returns a JSON summary of results.
'           : Pass one or more filter arguments to restrict which tests run:
'           :   ?VCS.RunTests("clsTestParser", "modTestEncoding")
'           :   ?VCS.RunTests("-slow")
'           : Omit arguments to run all tests. The DefaultTestFilter option is NOT
'           : applied here; use RunFilteredTests for that (ribbon entry point).
'---------------------------------------------------------------------------------------
'
Public Function RunTests(ParamArray varFilters() As Variant) As String

    Dim blnFiltered As Boolean

    ' Determine whether a filter was supplied
    blnFiltered = (UBound(varFilters) >= LBound(varFilters))

    If blnFiltered Then
        RunTests = ExecuteTests(CVar(varFilters), True)
    Else
        RunTests = ExecuteTests(Empty, False)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunTestsHeadless
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Automation/CI entry point. Runs tests with no forms and no prompts
'           : (missing modTestAssert is installed silently), always exports JUnit
'           : XML, and returns a machine-readable JSON summary including allPassed,
'           : cancelled, junitPath, and statePath. Same filter syntax as RunTests.
'           : Example (COM automation):
'           :   ?VCS.RunTestsHeadless("SQL", "-slow")
'---------------------------------------------------------------------------------------
'
Public Function RunTestsHeadless(ParamArray varFilters() As Variant) As String

    Dim blnFiltered As Boolean

    blnFiltered = (UBound(varFilters) >= LBound(varFilters))

    If blnFiltered Then
        RunTestsHeadless = ExecuteTests(CVar(varFilters), True, blnHeadless:=True)
    Else
        RunTestsHeadless = ExecuteTests(Empty, False, blnHeadless:=True)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunFilteredTests
' Author    : Adam Waller
' Date      : 5/21/2026
' Purpose   : Ribbon entry point for running tests. Reads the DefaultTestFilter option
'           : and applies it automatically. If the option is empty, runs all tests.
'---------------------------------------------------------------------------------------
'
Public Function RunFilteredTests() As String

    Dim strFilter As String
    Dim arrParts() As String
    Dim lngIdx As Long

    strFilter = Trim$(Nz(Options.DefaultTestFilter, vbNullString))

    If Len(strFilter) = 0 Then
        RunFilteredTests = ExecuteTests(Empty, False)
    Else
        arrParts = Split(strFilter, ",")
        For lngIdx = LBound(arrParts) To UBound(arrParts)
            arrParts(lngIdx) = Trim$(arrParts(lngIdx))
        Next lngIdx
        RunFilteredTests = ExecuteTests(CVar(arrParts), True)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenTestRunner
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Open the web test runner to VIEW the last run's results without running.
'           : The UI rehydrates from the in-memory results (failed set ready to re-run).
'           : Requires the web runner to be enabled and supported on this Access build.
'---------------------------------------------------------------------------------------
'
Public Sub OpenTestRunner()
    If Not Options.UseWebTestRunner Or Not modTestRunnerUI.EdgeTestRunnerSupported() Then
        MsgBox2 T("Web Test Runner Unavailable"), _
            T("The HTML test runner is disabled or not supported on this Access build."), _
            T("Enable it in Options, or use Run Tests for the console runner."), vbInformation
        Exit Sub
    End If
    modTestRunnerUI.OpenTestRunnerForResults
    modTestRunnerUI.WaitForWebRunnerReady
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTestResultsJUnit
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Generate JUnit XML from the durable test-state.json without running tests.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTestResultsJUnit(Optional ByVal strPath As String = vbNullString)

    Dim strResult As String

    strResult = modTestJUnit.ExportFromState(strPath)
    If Len(strResult) = 0 Then
        MsgBox2 T("No Test State"), _
            T("No durable test state file was found to export."), _
            T("Run tests first, or check the test-results folder."), vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTestResultsHtml
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Generate the HTML test-results report from durable test-state.json
'           : without running tests.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTestResultsHtml(Optional ByVal strPath As String = vbNullString)

    Dim strResult As String

    strResult = modTestReport.ExportResultsHtml(strPath)
    If Len(strResult) = 0 Then
        MsgBox2 T("No Test State"), _
            T("No durable test state file was found to export."), _
            T("Run tests first, or check the test-results folder."), vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExecuteTests
' Author    : Adam Waller
' Date      : 5/21/2026
' Purpose   : Shared implementation for RunTests, RunFilteredTests, and
'           : RunTestsHeadless. Scans for tests, optionally applies filters, runs the
'           : suite, and returns a JSON summary. Headless mode is for automation/CI:
'           : no forms, no prompts, console path forced, JUnit always exported.
'---------------------------------------------------------------------------------------
'
Private Function ExecuteTests(varFilters As Variant, blnFiltered As Boolean, _
    Optional ByVal blnHeadless As Boolean = False) As String

    Dim blnAllPassed As Boolean
    Dim lngPassed As Long
    Dim lngFailed As Long
    Dim lngErrored As Long
    Dim lngEmpty As Long
    Dim colKeys As Collection
    Dim lngIdx As Long
    Dim strFilter As String
    Dim eimPriorMode As eInteractionMode
    Dim blnLogWasActive As Boolean
    Dim blnUseWebRunner As Boolean
    Dim strJUnitPath As String

    If Not DatabaseFileOpen Then
        If blnHeadless Then
            ExecuteTests = "{""error"":""No database open""}"
        Else
            MsgBox2 T("No Database Open"), _
                T("Please open a database file before running tests."), _
                , vbInformation
        End If
        Exit Function
    End If

    ' Ensure modTestAssert is present (installed silently in headless mode; an
    ' automation session has no one to answer a prompt)
    If Not TestAssertModuleExists Then
        If blnHeadless Then
            InstallTestAssertModule
        ElseIf vbYes = MsgBox2(T("Test Helper Not Found"), _
            T("The modTestAssert module was not found in the current project."), _
            T("Would you like to install it now?"), _
            vbYesNo + vbQuestion) Then
            InstallTestAssertModule
        Else
            Exit Function
        End If
    End If

    ' Offer to migrate if no module (other than modTestAssert) uses TestAssert yet
    ' (skipped in headless mode)
    If Not TestAssertInUse And Not blnHeadless Then
        If vbYes = MsgBox2(T("TestAssert Not In Use"), _
            T("No test modules are using TestAssert yet. Would you like to migrate " & _
              "Debug.Assert calls to TestAssert in your test modules?"), _
            T("You can always do this manually later."), _
            vbYesNo + vbQuestion) Then
            MigrateDebugAssert
        End If
    End If

    ' Match the round-trip harness: log to file and suppress blocking MsgBox dialogs
    ' so import/build errors surface in the test console and TestRun_*.log instead.
    eimPriorMode = Operation.InteractionMode
    blnLogWasActive = Log.Active

    blnUseWebRunner = Options.UseWebTestRunner And modTestRunnerUI.EdgeTestRunnerSupported() _
        And Not blnHeadless

    ' Web runner path: open the form only. Do NOT begin an Operation here — bridge
    ' callbacks own the operation lifecycle when the user clicks Run. Beginning here
    ' left eosRunning stuck after hide/reopen and blocked the next Run Tests click.
    If blnUseWebRunner Then
        If blnFiltered Then
            For lngIdx = LBound(varFilters) To UBound(varFilters)
                If Len(strFilter) > 0 Then strFilter = strFilter & ", "
                strFilter = strFilter & Trim$(CStr(varFilters(lngIdx)))
            Next lngIdx
            modTestRunnerUI.SetWebRunnerDefaultFilter strFilter
        Else
            modTestRunnerUI.SetWebRunnerDefaultFilter vbNullString
        End If

        modTestRunnerUI.OpenWebTestRunner
        modTestRunnerUI.WaitForWebRunnerReady
        Exit Function
    End If

    If Not Operation.Begin(eotTestRun) Then Exit Function

    Operation.InteractionMode = eimSilent
    Log.Active = True
    Log.ClearErrorJournal

    ' No console form in headless mode (automation sessions have no UI to update)
    If Not blnHeadless Then PrepareTestConsole

    ' Build filter description for log heading
    If blnFiltered Then
        For lngIdx = LBound(varFilters) To UBound(varFilters)
            If Len(strFilter) > 0 Then strFilter = strFilter & ", "
            strFilter = strFilter & CStr(varFilters(lngIdx))
        Next lngIdx
    End If

    ' Display heading
    With Log
        .Spacer
        If blnFiltered Then
            .Add T("Running Tests (filtered: {0})", var0:=strFilter), False
        Else
            .Add T("Running Tests"), False
        End If
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add Now
        .Spacer
    End With

    ' Scan for tests
    Log.Add T("Scanning for tests...")
    Log.Flush
    TestRunner.Scan

    If TestRunner.Tests.Count = 0 Then
        Log.Add T("No test subs found in the current project.")
        If Not blnHeadless Then FinalizeTestConsole True
        Log.Active = blnLogWasActive
        Log.SuppressDebugOutput = False
        Operation.InteractionMode = eimPriorMode
        Operation.Finish eorSuccess
        Exit Function
    End If

    ' Apply filter if specified (resolves module, suite, procedure, and tag matches)
    If blnFiltered Then
        Set colKeys = TestRunner.ResolveFilters(varFilters)
        If colKeys.Count = 0 Then
            Log.Add T("No tests found matching filter: {0}", var0:=strFilter)
            If Not blnHeadless Then FinalizeTestConsole True
            Log.Active = blnLogWasActive
            Log.SuppressDebugOutput = False
            Operation.InteractionMode = eimPriorMode
            Operation.Finish eorSuccess
            Exit Function
        End If
        Log.Add T("Found {0} sub(s) matching filter. Running...", var0:=colKeys.Count)
    Else
        Log.Add T("Found {0} test sub(s). Running...", var0:=TestRunner.Tests.Count)
    End If
    Log.Spacer True, 61

    DoCmd.Hourglass True

    TestRunner.InvokeGlobalTestSetup

    ' Run tests (filtered or all)
    If blnFiltered Then
        TestRunner.RunSelected colKeys
    Else
        TestRunner.RunAll
    End If
    Dim strJsonPath As String
    strJsonPath = TestRunner.SaveResults
    modTestState.PersistAfterRun

    ' CI must always get the JUnit XML, regardless of the export option
    If blnHeadless Then strJUnitPath = modTestJUnit.ExportFromState()

    ' Detect cancellation (before computing summary, so "all passed" is not
    ' incorrectly reported when only a subset ran before the user canceled)
    Dim blnCanceled As Boolean
    blnCanceled = (TestRunner.State = etrsCancelled)

    ' Summary
    lngPassed = TestRunner.PassedCount
    lngFailed = TestRunner.FailedCount
    lngErrored = TestRunner.ErroredCount
    lngEmpty = TestRunner.EmptyCount
    blnAllPassed = (Not blnCanceled) And (lngFailed = 0 And lngErrored = 0 And lngPassed > 0)

    Log.Flush
    DoCmd.Hourglass False

    ' Add performance data to log file and save
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
    End With

    If blnHeadless Then
        ExecuteTests = TestRunner.GetResultsAsJson(strJUnitPath, modTestState.GetStateFilePath)
    Else
        FinalizeTestConsole blnAllPassed, Log.LogFilePath, blnCanceled, strJsonPath
        ExecuteTests = TestRunner.GetResultsAsJson
    End If

    TestRunner.InvokeGlobalTestTeardown

    Log.Active = blnLogWasActive
    Log.SuppressDebugOutput = False
    Operation.InteractionMode = eimPriorMode
    Operation.Finish IIf(blnCanceled, eorCanceled, IIf(blnAllPassed, eorSuccess, eorFailed))

End Function


'---------------------------------------------------------------------------------------
' Procedure : PrepareTestConsole
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Open frmVCSMain and configure it for test runner output.
'---------------------------------------------------------------------------------------
'
Private Sub PrepareTestConsole()
    On Error Resume Next
    If Not IsLoaded(acForm, "frmVCSMain", False) Then
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    End If
    With Form_frmVCSMain
        .ApplyTestConsoleLayout
        .ResetForOperation
        .SetStatusText T("Running Tests"), _
            T("Executing tests in the current project"), _
            T("Results are displayed below as each test completes.")
    End With
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FinalizeTestConsole
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Update the main form after test execution completes.
'---------------------------------------------------------------------------------------
'
Private Sub FinalizeTestConsole(blnAllPassed As Boolean, _
    Optional ByVal strLogPath As String = vbNullString, _
    Optional ByVal blnCanceled As Boolean = False, _
    Optional ByVal strCopyPath As String = vbNullString)

    Dim strReportPath As String

    On Error Resume Next
    With Form_frmVCSMain
        .txtLog.ScrollBars = 2
        .strLastLogFilePath = strLogPath
        .cmdOpenLogFile.Visible = (Len(strLogPath) > 0)
        .cmdCopyPath.Visible = .cmdOpenLogFile.Visible
        .strCopyToClipboardPath = IIf(Len(strCopyPath) > 0, strCopyPath, strLogPath)
        .strLastTestResultsPath = vbNullString
        .cmdOpenTestResults.Visible = False
        If Options.ExportTestResultsHtml Then
            strReportPath = modTestState.GetTestResultsFolder() & "test-results.html"
            If FSO.FileExists(strReportPath) Then
                .strLastTestResultsPath = strReportPath
                .cmdOpenTestResults.Visible = True
            End If
        End If
        If blnCanceled Then
            .SetStatusText T("Canceled"), T("Tests Canceled"), _
                T("The test run was canceled. Partial results are available in the log file.")
        ElseIf blnAllPassed Then
            .SetStatusText T("Finished"), T("All Tests Passed"), _
                T("All tests completed successfully. Click the log button to view the full results.")
        Else
            .SetStatusText T("Finished"), T("Some Tests Failed"), _
                T("Review the results above for details on failed tests. Click the log button to view the full log.")
        End If
    End With
    Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : InstallTestAssertModule
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Install the modTestAssert helper module into the current project.
'           : Follows the same pattern as CreateLetterCasingTemplate.
'---------------------------------------------------------------------------------------
'
Public Sub InstallTestAssertModule()

    Dim cmp As VBIDE.VBComponent
    Dim cCode As New clsConcat

    If TestAssertModuleExists Then
        MsgBox2 T("Already Installed"), _
            T("The modTestAssert module already exists in the current project."), _
            , vbInformation
        Exit Sub
    End If

    Set cmp = CurrentVBProject.VBComponents.Add(vbext_ct_StdModule)
    cmp.Name = "modTestAssert"

    With cCode
        .AppendOnAdd = vbCrLf
        .Add "'---------------------------------------------------------------------------------------"
        .Add "' Module    : modTestAssert"
        .Add "' Purpose   : Lightweight assertion module for VBA test suites. TestAssert is a"
        .Add "'           : near-drop-in replacement for Debug.Assert that integrates with the"
        .Add "'           : VCS add-in test runner when available, and falls back to Debug.Assert"
        .Add "'           : when the add-in is not loaded or no test run is active."
        .Add "'---------------------------------------------------------------------------------------"
        .Add "Option Compare Database"
        .Add "Option Explicit"
        .Add "'@Folder(""Tests"")"
        .Add vbNullString
        .Add "Private m_strAddInPath As String"
        .Add vbNullString
        .Add "Public Sub TestAssert(ByVal condition As Boolean, _"
        .Add "                      Optional ByVal context As Variant)"
        .Add "    On Error Resume Next"
        .Add "    Dim blnHandled As Boolean"
        .Add "    blnHandled = Application.Run(ResolveAddInPath() & "".HandleTestAssertion"", condition, context)"
        .Add "    If Err.Number <> 0 Or Not blnHandled Then"
        .Add "        Err.Clear"
        .Add "        Debug.Assert condition"
        .Add "    End If"
        .Add "End Sub"
        .Add vbNullString
        .Add "Private Function ResolveAddInPath() As String"
        .Add "    If Len(m_strAddInPath) = 0 Then"
        .Add "        m_strAddInPath = GetSetting(""MSAccessVCS"", ""Install"", ""Install Folder"", _"
        .Add "            Environ$(""AppData"") & ""\MSAccessVCS"") & ""\Version Control"""
        .Add "    End If"
        .Add "    ResolveAddInPath = m_strAddInPath"
        .Add "End Function"
        .Add vbNullString
        .Add "'---------------------------------------------------------------------------------------"
        .Add "' Procedure : GlobalTestSetup"
        .Add "' Purpose   : Optional suite-level setup. Called once per VCS.RunTests invocation,"
        .Add "'           : immediately before the first test runs."
        .Add "'---------------------------------------------------------------------------------------"
        .Add "Public Sub GlobalTestSetup()"
        .Add "    ' Add one-time setup here (e.g. reset shared test state, sweep leftover temp data)."
        .Add "    ' Leave empty if your project does not need suite-level setup."
        .Add "End Sub"
        .Add vbNullString
        .Add "'---------------------------------------------------------------------------------------"
        .Add "' Procedure : GlobalTestTeardown"
        .Add "' Purpose   : Optional suite-level cleanup. Called once per VCS.RunTests invocation"
        .Add "'           : after all tests complete and results JSON is built."
        .Add "'---------------------------------------------------------------------------------------"
        .Add "Public Sub GlobalTestTeardown()"
        .Add "    ' Add one-time cleanup here (e.g. delete temp objects created during the run)."
        .Add "    ' Leave empty if your project does not need suite-level teardown."
        .Add "End Sub"
    End With

    If cmp.CodeModule.CountOfLines > 0 Then
        cmp.CodeModule.DeleteLines 1, cmp.CodeModule.CountOfLines
    End If
    cmp.CodeModule.InsertLines 1, cCode.GetStr

    MsgBox2 T("Test Helper Installed"), _
        T("The modTestAssert module has been added to your project."), _
        T("Use TestAssert in your test subs to record assertions. " & _
          "Run tests from the VCS ribbon or call VCS.RunTests."), _
        vbInformation

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MigrateDebugAssert
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Scoped migration tool: replace Debug.Assert with TestAssert in test
'           : modules only. Uses @Folder("Tests") discovery when folder annotations
'           : exist in the project, otherwise falls back to *Test* in the module name.
'           : The module that defines TestAssert itself is skipped (detected by
'           : procedure signature, so a renamed helper is still excluded) -- its
'           : Debug.Assert is the intentional fallback when the add-in is not
'           : driving the run. Shows a preview and requires confirmation.
'---------------------------------------------------------------------------------------
'
Public Sub MigrateDebugAssert()

    Dim proj As VBIDE.VBProject
    Dim cmp As VBIDE.VBComponent
    Dim blnHasFolders As Boolean
    Dim dReplacements As New Dictionary
    Dim lngTotalCount As Long
    Dim blnIsTest As Boolean
    Dim strHeader As String
    Dim strFolder As String
    Dim lngFolderPos As Long
    Dim lngQStart As Long
    Dim lngQEnd As Long
    Dim varParts As Variant
    Dim p As Long
    Dim colLines As Collection
    Dim lngLine As Long
    Dim strLine As String
    Dim strPreview As String
    Dim varModName As Variant
    Dim colModLines As Collection
    Dim lngApplied As Long
    Dim lngIdx As Long
    Dim strOriginal As String
    Dim strNew As String

    If Not DatabaseFileOpen Then
        MsgBox2 T("No Database Open"), _
            T("Please open a database file first."), , vbInformation
        Exit Sub
    End If

    Set proj = CurrentVBProject

    ' Check if project uses @Folder annotations
    blnHasFolders = TestRunner.State >= etrsReady
    If Not blnHasFolders Then
        TestRunner.Scan
    End If

    ' Scan for Debug.Assert in qualifying modules
    For Each cmp In proj.VBComponents
        If cmp.Type <> vbext_ct_StdModule Then GoTo NextModule

        blnIsTest = False

        On Error Resume Next
        If cmp.CodeModule.CountOfLines = 0 Then GoTo NextModule
        strHeader = cmp.CodeModule.Lines(1, IIf(cmp.CodeModule.CountOfLines < 30, cmp.CodeModule.CountOfLines, 30))

        ' Check for @Folder("Tests") annotation
        strFolder = vbNullString
        lngFolderPos = InStr(1, UCase$(strHeader), "'@FOLDER(")
        If lngFolderPos > 0 Then
            lngQStart = InStr(lngFolderPos, strHeader, """")
            If lngQStart > 0 Then
                lngQEnd = InStr(lngQStart + 1, strHeader, """")
                If lngQEnd > lngQStart Then
                    strFolder = Mid$(strHeader, lngQStart + 1, lngQEnd - lngQStart - 1)
                End If
            End If
        End If
        On Error GoTo 0

        If Len(strFolder) > 0 Then
            varParts = Split(strFolder, ".")
            For p = LBound(varParts) To UBound(varParts)
                If StrComp(varParts(p), "Tests", vbTextCompare) = 0 Then
                    blnIsTest = True
                    Exit For
                End If
            Next p
        ElseIf InStr(1, cmp.Name, "Test", vbTextCompare) > 0 Then
            blnIsTest = True
        End If

        If Not blnIsTest Then GoTo NextModule

        ' Scan this module for Debug.Assert. If the module defines TestAssert
        ' itself, skip it entirely -- its Debug.Assert is the intentional
        ' fallback for when the add-in is not driving the run. Detected by
        ' procedure signature (not module name) so a renamed helper is still
        ' excluded.
        Set colLines = New Collection
        On Error Resume Next
        For lngLine = 1 To cmp.CodeModule.CountOfLines
            strLine = Trim$(cmp.CodeModule.Lines(lngLine, 1))
            If InStr(1, strLine, "Public Sub TestAssert(", vbTextCompare) > 0 Then
                On Error GoTo 0
                GoTo NextModule
            End If
            If InStr(1, strLine, "Debug.Assert ", vbTextCompare) > 0 Or _
               InStr(1, strLine, "Debug.Assert(", vbTextCompare) > 0 Then
                colLines.Add lngLine
            End If
        Next lngLine
        On Error GoTo 0

        If colLines.Count > 0 Then
            dReplacements.Add cmp.Name, colLines
            lngTotalCount = lngTotalCount + colLines.Count
        End If

NextModule:
    Next cmp

    If lngTotalCount = 0 Then
        MsgBox2 T("No Replacements Found"), _
            T("No Debug.Assert calls were found in test modules."), _
            , vbInformation
        Exit Sub
    End If

    ' Build preview text
    strPreview = T("{0} occurrence(s) in {1} module(s):", _
        var0:=lngTotalCount, var1:=dReplacements.Count) & vbCrLf & vbCrLf
    For Each varModName In dReplacements.Keys
        Set colModLines = dReplacements(CStr(varModName))
        strPreview = strPreview & CStr(varModName) & " (" & colModLines.Count & " lines)" & vbCrLf
    Next varModName
    strPreview = strPreview & vbCrLf & _
        T("Note: Debug.Assert breaks into the debugger on failure. " & _
          "TestAssert records the failure and continues. " & _
          "This change can be reverted with git.")

    ' Confirm
    If vbYes <> MsgBox2(T("Migrate Debug.Assert to TestAssert"), _
        T("Replace Debug.Assert with TestAssert in the following test modules?"), _
        strPreview, _
        vbYesNo + vbQuestion) Then
        Exit Sub
    End If

    ' Apply replacements (bottom to top so line numbers stay valid)
    For Each varModName In dReplacements.Keys
        Set cmp = proj.VBComponents(CStr(varModName))
        Set colModLines = dReplacements(CStr(varModName))
        For lngIdx = colModLines.Count To 1 Step -1
            lngLine = colModLines(lngIdx)
            strOriginal = cmp.CodeModule.Lines(lngLine, 1)
            strNew = Replace(strOriginal, "Debug.Assert ", "TestAssert ", , , vbTextCompare)
            strNew = Replace(strNew, "Debug.Assert(", "TestAssert(", , , vbTextCompare)
            If strNew <> strOriginal Then
                cmp.CodeModule.ReplaceLine lngLine, strNew
                lngApplied = lngApplied + 1
            End If
        Next lngIdx
    Next varModName

    MsgBox2 T("Migration Complete"), _
        T("{0} Debug.Assert call(s) replaced with TestAssert.", var0:=lngApplied), _
        , vbInformation

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAssertModuleExists
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Check if modTestAssert already exists in CurrentVBProject.
'---------------------------------------------------------------------------------------
'
Private Function TestAssertModuleExists() As Boolean
    Dim cmp As VBIDE.VBComponent
    On Error Resume Next
    Set cmp = CurrentVBProject.VBComponents("modTestAssert")
    TestAssertModuleExists = (Not cmp Is Nothing)
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : TestAssertInUse
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Returns True if any standard module (other than modTestAssert itself)
'           : contains a "TestAssert " call. Used to decide whether to offer the
'           : Debug.Assert migration prompt -- once the user has adopted TestAssert
'           : in at least one place, we stop asking.
'---------------------------------------------------------------------------------------
'
Private Function TestAssertInUse() As Boolean

    Dim cmp As VBIDE.VBComponent
    Dim strCode As String

    On Error Resume Next
    For Each cmp In CurrentVBProject.VBComponents
        If cmp.Type = vbext_ct_StdModule Then
            If StrComp(cmp.Name, "modTestAssert", vbTextCompare) <> 0 Then
                If cmp.CodeModule.CountOfLines > 0 Then
                    strCode = cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
                    If InStr(1, strCode, "TestAssert ", vbTextCompare) > 0 Then
                        TestAssertInUse = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next cmp
    On Error GoTo 0

End Function


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


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Save the current state so we can restore it after the current operation.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ' When the class is initialized, make sure the ribbon is active (if installed).
    ' This way if the COM add-in is not active, it will be automatically activated
    ' when the add-in is opened from the [Database Tools\Add-ins] menu. (See #451)
    If GetInstallSettings.blnUseRibbonAddIn Then modCOMAddIn.VerifyRibbon
End Sub
