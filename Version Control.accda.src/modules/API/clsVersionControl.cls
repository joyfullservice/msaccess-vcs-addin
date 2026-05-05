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

    Dim lngResult As Long

    If Not Operation.Begin(eotOther) Then Exit Sub

    lngResult = modLetterCasing.StandardizeLetterCasing

    Select Case lngResult
        Case Is > 0
            MsgBox2 T("Letter Casing Standardized"), _
                T("{0} correction(s) applied to VBA identifiers.", var0:=lngResult), _
                , vbInformation
        Case 0
            MsgBox2 T("Letter Casing Consistent"), _
                T("All VBA identifiers already match the standard letter casing rules."), _
                , vbInformation
        Case -1
            If vbYes = MsgBox2(T("Module Not Found"), _
                T("The clsStandardLetterCasing class module was not found in the current project."), _
                T("Would you like to create a starter template?"), _
                vbYesNo + vbQuestion) Then
                CreateLetterCasingTemplate
            End If
    End Select

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

    ' Execute the wrapper function. Application.Run's qualified syntax is
    ' "[ProjectName].[FunctionName]" -- module name is NOT supported. From
    ' add-in code calling back into the host, an unqualified call resolves
    ' against the add-in's own project, so we must qualify with the host's
    ' VBProject name. CurrentVBProject points at the host (not the add-in)
    ' here. The empirically verified path-without-extension qualifier used
    ' in the reverse direction (modAPI.GetRunCmdAddInFullLibName) does NOT
    ' work for the host project, only for loaded library references.
    Dim strProjPrefix As String
    strProjPrefix = CurrentVBProject.Name & "."
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
    ' Always remove the temp module
    If blnModuleCreated Then
        On Error Resume Next
        CurrentVBProject.VBComponents.Remove cmp
        On Error GoTo 0
    End If

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
' Author    : VCS contributors
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
' Author    : VCS contributors


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
