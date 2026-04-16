Attribute VB_Name = "modExport"
'---------------------------------------------------------------------------------------
' Module    : modExport
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Export functions for saving database objects to source files.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modDatabase, modFileAccess, modVCSUtility,
'           : modSourceUpgrade, modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName As String = "modExport"


'---------------------------------------------------------------------------------------
' Procedure : ExportSource
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Export source files from the currently open database.
'---------------------------------------------------------------------------------------
'
Public Sub ExportSource(blnFullExport As Boolean, Optional intFilter As eContainerFilter = ecfAllObjects, Optional frmMain As Form_frmVCSMain)

    Dim dCategories As Dictionary
    Dim colCategories As Collection
    Dim dCategory As Dictionary
    Dim dObjects As Dictionary
    Dim varCatKey As Variant
    Dim varKey As Variant
    Dim cCategory As IDbComponent
    Dim cDbObject As IDbComponent
    Dim lngCount As Long
    Dim lngCurrent As Long
    Dim strTempFile As String

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Can't export without an open database
    If Not DatabaseFileOpen Then Exit Sub

    ' If we are running this from the current database, we need to run it a different
    ' way to prevent file corruption issues. (This really shouldn't happen after v4.02)
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        MsgBox2 T("Unabled to Export Running Database", "Please launch the export using the add-in menu or ribbon"), , vbExclamation
        Exit Sub
    End If

    ' Reload the project options and reset the logs
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

    ' Check error handling mode after loading project options
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Display heading early so the user sees output before heavy I/O
    With Log
        .Spacer
        .Add T("Beginning Export of Source Files"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Export Format: {0}", var0:=ExportFormatToVersion(Options.ExportFormatVersion))
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        If Not frmMain Is Nothing Then frmMain.strLastLogFilePath = .LogFilePath
        .Flush
    End With

    ' Determine which categories need full re-export due to options changes
    Dim dCurrentHashes As Dictionary
    Dim dStoredHashes As Dictionary
    Dim dStaleCategories As Dictionary
    Dim blnGlobalChanged As Boolean

    Set dCurrentHashes = Options.GetCategoryHashes
    Set dStoredHashes = VCSIndex.CategoryHashes
    Set dStaleCategories = New Dictionary

    ' Check global options (ExportFormatVersion, AccessVersion)
    If dCurrentHashes.Exists("_Global") Then
        If Not dStoredHashes.Exists("_Global") Then
            blnGlobalChanged = True
        ElseIf dCurrentHashes("_Global") <> dStoredHashes("_Global") Then
            blnGlobalChanged = True
        End If
    End If

    ' Check category-specific options
    Dim varHashKey As Variant
    For Each varHashKey In dCurrentHashes.Keys
        If CStr(varHashKey) <> "_Global" Then
            If Not dStoredHashes.Exists(CStr(varHashKey)) Then
                dStaleCategories.Add CStr(varHashKey), True
            ElseIf dCurrentHashes(CStr(varHashKey)) <> dStoredHashes(CStr(varHashKey)) Then
                dStaleCategories.Add CStr(varHashKey), True
            End If
        End If
    Next varHashKey

    ' Global change affects all categories
    If blnGlobalChanged Then blnFullExport = True

    ' Log export mode after determining category changes
    If blnFullExport Then
        Log.Add T("Performing Full Export")
    ElseIf dStaleCategories.Count > 0 Then
        Log.Add T("Using Fast Save (re-exporting {0} changed categories)", _
            var0:=dStaleCategories.Count)
    Else
        Log.Add T("Using Fast Save")
    End If

    ' Check VBE project access
    If CurrentVBProject.Protection = vbext_pp_locked Then
        If IsMDE Then
            MsgBox2 T("Compiled Database"), _
                T("The current database is a compiled MDE/ACCDE file and does not contain the original VBA source code."), _
                T("Please use the original uncompiled .accdb file instead."), vbExclamation
        Else
            MsgBox2 T("Project Locked"), _
                T("Project is protected with a password."), _
                T("Please unlock the project before using this tool."), vbExclamation
        End If
        Log.Spacer
        Log.Add T("Export Canceled"), , , "Red", True
        Log.Flush
        Operation.ErrorLevel = eelCritical
        Exit Sub
    End If

    ' Check project VCS version
    Select Case Options.CompareLoadedVersion
        Case evcNewerVersion
            Log.Flush
            If MsgBox2(T("Newer VCS Version Detected"), _
                T("This project uses VCS version {0}, but version {1} is currently installed." & _
                    vbNewLine & "Would you like to continue anyway?", _
                    var0:=Options.GetLoadedVersion, var1:=GetVCSVersion), _
                T("Click YES to continue this operation, or NO to cancel."), _
                vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Log.Spacer
                    Log.Add T("Export Canceled"), , , "Red", True
                    Log.Flush
                    Operation.ErrorLevel = eelCritical
                    Exit Sub
            End If
        Case evcOlderVersion
            Log.Add T("Updated VCS ({0} -> {1})", var0:=Options.GetLoadedVersion, var1:=GetVCSVersion), , , "blue"
    End Select

    ' Notify about newer export format version
    If Options.ExportFormatVersion < LATEST_EXPORT_FORMAT Then
        Log.Add T("Note: Export format {0} is available (currently using {1}). Update via Options > Export when ready.", _
            var0:=ExportFormatToVersion(LATEST_EXPORT_FORMAT), _
            var1:=ExportFormatToVersion(Options.ExportFormatVersion)), , , "blue"
    End If

    ' Perform any needed upgrades to source files
    If blnFullExport Then UpgradeSourceFiles

    ' Migrate file extensions between format versions
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        MigrateFileExtensions
    Else
        RevertFileExtensions
    End If

    ' Run any custom sub before export
    If Options.RunBeforeExport <> vbNullString Then
        Log.Add T("Running {0}...", var0:=Options.RunBeforeExport)
        Log.Flush
        Perf.OperationStart "RunBeforeExport"
        RunSubInCurrentProject Options.RunBeforeExport
        Perf.OperationEnd
    End If

    ' Close any open database objects.
    If Not CloseDatabaseObjects Then
        MsgBox2 T("Please close all database objects"), _
            T("All database objects (i.e.forms, reports, tables, queries, etc...) must be closed to export source code."), _
            , vbExclamation
        Exit Sub
    End If

    ' Cache persistent connections to Access back-end databases
    CacheBackEndConnections

    ' Enforce any supplied letter casing rules
    StandardizeLetterCasing

    ' Export any external database schemas
    ExportSchemas blnFullExport
    If Operation.ErrorLevel = eelCritical Then GoTo CleanUp

    ' Finish header section
    Log.Spacer
    If blnFullExport Then
        Log.Add T("Scanning source files...")
    ElseIf dStaleCategories.Count > 0 Then
        Log.Add T("Scanning for changes (some categories require full export)...")
    Else
        Log.Add T("Scanning for changes...")
    End If
    Log.Flush

    ' Set up progress bar to show status on large projects
    Set colCategories = GetContainers(intFilter)
    Log.ProgressBar.Reset
    Log.ProgressBar.Max = GetQuickObjectCount(colCategories) + GetQuickFileCount(colCategories)

    ' Scan database objects for changes
    Set dCategories = New Dictionary
    VCSIndex.Conflicts.Initialize dCategories, eatExport
    Perf.OperationStart "Scan DB Objects"
    For Each cCategory In colCategories
        Perf.CategoryStart cCategory.Category
        Operation.Pulse
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        ' Get collection of database objects (IDbComponent classes)
        Dim blnFullForCategory As Boolean
        blnFullForCategory = blnFullExport Or dStaleCategories.Exists(cCategory.Category)
        Set dObjects = cCategory.GetAllFromDB(Not blnFullForCategory)
        If dObjects.Count = 0 Then
            Log.Add IIf(blnFullForCategory, _
                T("No {0} found in this database.", var0:=T(LCase(cCategory.Category))), _
                T("No modified {0} found in this database.", var0:=T(LCase(cCategory.Category)))), _
                Options.ShowDebug
        End If
        dCategory.Add "Objects", dObjects
        dCategories.Add cCategory.Category, dCategory
        VCSIndex.CheckExportConflicts dObjects
        ' Clear any orphaned files in this category
        ClearOrphanedSourceFiles cCategory
        Perf.CategoryEnd 0
        ' Handle critical error or cancel during scan
        If Operation.ErrorLevel = eelCritical Then
            Log.Add vbNullString
            Perf.OperationEnd   ' Scan DB Objects
            GoTo CleanUp
        End If
    Next cCategory
    Perf.OperationEnd
    Log.ProgressBar.Reset

    ' Check for any conflicts
    With VCSIndex.Conflicts
        If .Count > 0 Then
            ' Resolve conflicts (auto-resolve for agent/API, prompt for user)
            .ResolveOrPrompt
            If .ApproveResolutions Then
                Log.Add T("Resolving source conflicts"), False
                .Resolve
            Else
                ' Cancel export
                Log.Spacer
                Log.Add T("Export Canceled"), , , "Red", True
                Operation.ErrorLevel = eelCritical
                GoTo CleanUp
            End If
        End If
    End With

    ' Loop through all categories
    For Each varCatKey In dCategories.Keys

        ' Get category class and collection of items
        Set dCategory = dCategories(varCatKey)
        Set cCategory = dCategory("Class")
        Set dObjects = dCategory("Objects")

        ' Only show category details when it contains objects
        lngCount = dObjects.Count
        If lngCount > 0 Then

            ' Show category header and clear out any orphaned files.
            Log.Spacer Options.ShowDebug
            Log.PadRight T("Exporting {0}...", var0:=T(LCase(cCategory.Category))), , Options.ShowDebug
            Perf.CategoryStart cCategory.Category
            lngCurrent = 0

            ' Post category start to MCP server
            MCP.Log "Exporting " & lngCount & " " & LCase(cCategory.Category) & "..."
            Log.Flush

            ' Loop through each object in this category.
            For Each varKey In dObjects.Keys

                ' Export object
                lngCurrent = lngCurrent + 1
                Set cDbObject = dObjects(varKey)
                Log.Add "  " & cDbObject.Name, Options.ShowDebug
                Log.Progress lngCurrent, lngCount, cDbObject.Name
                Operation.Pulse

                ' If we have already exported this object while scanning for changes, use that copy.
                strTempFile = Replace(cDbObject.SourceFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
                If FSO.FileExists(strTempFile) Then
                    ' Move the temp file(s) over to the source export folder.
                    cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, _
                        FSO.GetParentFolderName(cDbObject.SourceFile) & PathSep
                    ' Update the index with the values from the alternate export
                    VCSIndex.UpdateFromAltExport cDbObject
                Else
                    ' Export a fresh copy
                    cDbObject.Export
                End If

                ' Bail out if we hit a critical error.
                CatchAny eelError, T("Error exporting {0}", var0:=cDbObject.Name), ModuleName & ".ExportSource", True, True
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp

                ' Some kinds of objects are combined into a single export file, such
                ' as database properties. For these, we just need to run the export once.
                If cCategory.SingleFile Then Exit For

            Next varKey

            ' Show category wrap-up.
            If Options.ShowDebug Then
                Log.Add T("[{0}] {1} processed.", var0:=lngCount, var1:=T(LCase(cCategory.Category)))
            Else
                Log.Add "[" & lngCount & "]"
            End If
            'Log.Flush  ' Gives smoother output, but slows down export.
            Perf.CategoryEnd lngCount

            ' During fast save, log how many unchanged objects were skipped
            If Not blnFullForCategory Then
                Dim lngSkipped As Long
                lngSkipped = cCategory.QuickCount - lngCount
                If lngSkipped > 0 Then
                    Log.Add T("  Skipped {0} unchanged {1}", var0:=lngSkipped, var1:=LCase(cCategory.Category)), Options.ShowDebug
                End If
            End If
        End If

    Next varCatKey

    ' Ensure that we have created the .gitignore and .gitattributes files in Git environments.
    CheckGitFiles

    ' Export AGENTS.md file for AI agent assistance
    modResource.ExtractResource "AGENTS.md", Options.GetExportFolder

    ' Run any custom sub after export
    If Options.RunAfterExport <> vbNullString Then
        Log.Add T("Running {0}...", var0:=Options.RunAfterExport)
        Perf.OperationStart "RunAfterExport"
        RunSubInCurrentProject Options.RunAfterExport
        Perf.OperationEnd
        CatchAny eelError, T("Error running {0}", var0:=Options.RunAfterExport), ModuleName & ".ExportSource", True, True
    End If

    ' Log any unused .env connection entries (full export only)
    If blnFullExport Then LogUnusedEnvEntries
    CheckGitignoreForEnv

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

CleanUp:

    ' Run any cleanup routines
    CloseBackEndConnections
    ClearEnvCache
    VCSIndex.ClearTempExportFolder
    RemoveThemeZipFiles

    ' Add performance data to log file and save file
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
        .Flush
    End With

    ' Check for VCS_ImportExport.bas (Used with other forks)
    CheckForLegacyModules

    ' Restore original fast save option, and save options with project
    Options.SaveOptionsForProject

    ' Save index file
    With VCSIndex
        .ExportDate = Now
        If blnFullExport Then .FullExportDate = Now
        Set .CategoryHashes = dCurrentHashes
        .Save
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportSingleObject
' Author    : Adam Waller
' Date      : 2/22/2023
' Purpose   : Export a single object (such as a selected item)
'           : When blnNoIndex is True, the VCS index is disabled for the duration
'           : of the call, skipping the expensive full-file parse/serialize cycle
'           : and conflict detection. Used by MCP/API callers that treat the export
'           : as a deliberate action (like a user saving directly in the designer).
'---------------------------------------------------------------------------------------
'
Public Sub ExportSingleObject(objItem As AccessObject, Optional frmMain As Form_frmVCSMain, _
    Optional blnNoIndex As Boolean = False)

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim dObjects As Dictionary
    Dim cDbObject As IDbComponent
    Dim strTempFile As String

    ' Guard clause
    If objItem Is Nothing Then Exit Sub

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure the object is currently closed
    With objItem
        Select Case .Type
            Case acForm, acMacro, acModule, acQuery, acReport, acTable
                If SysCmd(acSysCmdGetObjectState, .Type, .Name) <> adStateClosed Then
                    DoCmd.Close .Type, .Name, acSavePrompt
                End If
        End Select
    End With

    If blnNoIndex Then
        ' Skip the expensive index load and options reload. The caller has already
        ' set up Options and is treating this as a direct edit (agent-as-user).
        VCSIndex.Disabled = True
        Log.Clear
        Log.SourcePath = Options.GetExportFolder
        Log.Active = True
        Perf.StartTiming
    Else
        ' Reload the project options and reset the logs
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
    End If

    ' Check error handling mode after loading project options
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Display heading
    With Log
        .Spacer
        .Add T("Beginning Export of Single Object"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        ' Log operation source (file only, not console)
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
        .Add T("Exporting {0}...", var0:=objItem.Name)
        .Flush
        ' Save export log file path
        If Not frmMain Is Nothing Then frmMain.strLastLogFilePath = .LogFilePath
    End With

    ' Cache persistent connections to Access back-end databases
    CacheBackEndConnections

    ' Get a database component class from the item
    Set cDbObject = GetClassFromObject(objItem)

    If Not blnNoIndex Then
        ' Check for conflicts
        Set dObjects = New Dictionary
        Set dCategory = New Dictionary
        Set dCategories = New Dictionary
        dObjects.Add cDbObject.SourceFile, cDbObject
        dCategory.Add "Class", cDbObject
        dCategory.Add "Objects", dObjects
        dCategories.Add cDbObject.Category, dCategory
        VCSIndex.Conflicts.Initialize dCategories, eatExport
        VCSIndex.CheckExportConflicts dObjects

        ' Resolve any outstanding conflict, or allow user to cancel.
        With VCSIndex.Conflicts
            If .Count > 0 Then
                ' Resolve conflicts (auto-resolve for agent/API, prompt for user)
                .ResolveOrPrompt
                If .ApproveResolutions Then
                    Log.Add T("Resolving source conflicts"), False
                    .Resolve
                Else
                    ' Cancel export
                    Log.Spacer
                    Log.Add T("Export Canceled"), , , "Red", True
                    Operation.ErrorLevel = eelCritical
                    GoTo CleanUp
                End If
            End If
        End With

        ' Check to see if we still have an item to export.
        If dCategories.Count = 0 Then
            Log.Add T("Skipped after conflict resolution."), , , "blue", True
            GoTo PostExport
        End If
    End If

    ' If we have already exported this object while scanning for changes, use that copy.
    If Not blnNoIndex Then
        strTempFile = Replace(cDbObject.SourceFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
        If FSO.FileExists(strTempFile) Then
            ' Move the temp file(s) over to the source export folder.
            cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, _
                FSO.GetParentFolderName(cDbObject.SourceFile) & PathSep
            ' Update the index with the values from the alternate export
            VCSIndex.UpdateFromAltExport cDbObject
        Else
            cDbObject.Export
        End If
    Else
        ' Export a fresh copy
        cDbObject.Export
    End If
    ExportDependentObjects cDbObject

    ' Export AGENTS.md file for AI agent assistance
    modResource.ExtractResource "AGENTS.md", Options.GetExportFolder

PostExport:

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

CleanUp:

    ' Run any cleanup routines
    CloseBackEndConnections
    ClearEnvCache
    If Not blnNoIndex Then VCSIndex.ClearTempExportFolder

    ' Add performance data to log file and save file
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
        .Flush
    End With

    If blnNoIndex Then
        VCSIndex.Disabled = False
    Else
        ' Save index file (don't change export date for single item export)
        VCSIndex.Save
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportMultipleObjects
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : Export multiple objects, passing a dictionary containing AccessObject.
'---------------------------------------------------------------------------------------
'
Public Sub ExportMultipleObjects(objItems As Dictionary, Optional bolForceClose As Boolean = True)

    Dim frm As Form_frmVCSMain

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim dObjects As Dictionary
    Dim cDbObject As IDbComponent
    Dim objItem As Access.AccessObject
    Dim strTempFile As String
    Dim varKey As Variant
    Dim varCategory As Variant
    Dim varObject As Variant

    ' Guard clause
    If objItems Is Nothing Then Exit Sub
    If objItems.Count = 0 Then Exit Sub

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Reset the log file
    Log.Clear

    ' Use the main form to display progress
    DoCmd.OpenForm "frmVCSMain", , , , , acHidden
    Set frm = Form_frmVCSMain   ' Connect to hidden instance
    With frm

        ' Prepare the UI screen
        .cmdClose.SetFocus
        .HideActionButtons
        DoEvents
        With .txtLog
            .ScrollBars = 0
            .Visible = True
            .SetFocus
        End With
        Log.SetConsole .txtLog, .GetProgressBar
        .strLastLogFilePath = Log.LogFilePath

        ' Show the status
        .SetStatusText T("Running..."), T("Automatically exporting the saved source code"), _
            T("A summary of the export progress can be seen on this screen, and additional details are included in the log file.")
        .Visible = True
    End With

    ' Make sure the object is currently closed
    If bolForceClose Then
        For Each varKey In objItems.Keys
            With objItems.Item(varKey)
                Select Case .Type
                    Case acForm, acMacro, acModule, acQuery, acReport, acTable
                        If SysCmd(acSysCmdGetObjectState, .Type, .Name) <> adStateClosed Then
                            DoCmd.Close .Type, .Name, acSavePrompt
                        End If
                End Select
            End With
        Next
    End If

    ' Reload the project options and reset the logs
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

    ' Check error handling mode after loading project options
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Display heading
    With Log
        .Spacer
        .Add T("Beginning Export of Multiple Objects"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        ' Log operation source (file only, not console)
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
        .Flush
    End With

    ' Cache persistent connections to Access back-end databases
    CacheBackEndConnections

    Set dCategories = New Dictionary

    For Each varKey In objItems.Keys
        Set objItem = objItems.Item(varKey)
        Log.Add T("Exporting {0}...", var0:=objItem.Name)
        Log.Flush

        ' FIXME: Hackish, need to figure a clean way of communicating types instead of encoding the key
        Dim lngObjectType As Access.AcObjectType
        LogUnhandledErrors
        On Error Resume Next
        lngObjectType = CLng(Split(varKey, "|")(0))
        On Error GoTo 0
        If lngObjectType = acTableDataMacro Then
            Set cDbObject = New clsDbTableDataMacro
        Else
            ' Get a database component class from the item
            Set cDbObject = GetClassFromObject(objItem)
        End If

        ' Check for conflicts
        If Not dCategories.Exists(cDbObject.Category) Then
            Set dObjects = New Dictionary
            Set dCategory = New Dictionary

            dObjects.Add cDbObject.SourceFile, cDbObject
            dCategory.Add "Class", cDbObject
            dCategory.Add "Objects", dObjects
            dCategories.Add cDbObject.Category, dCategory
        Else
            dCategories.Item(cDbObject.Category).Item("Objects").Add cDbObject.SourceFile, cDbObject
        End If

        VCSIndex.Conflicts.Initialize dCategories, eatExport
        VCSIndex.CheckExportConflicts dObjects
    Next

    ' Resolve any outstanding conflict, or allow user to cancel.
    With VCSIndex.Conflicts
        If .Count > 0 Then
            ' Resolve conflicts (auto-resolve for agent/API, prompt for user)
            .ResolveOrPrompt
            If .ApproveResolutions Then
                Log.Add T("Resolving source conflicts"), False
                .Resolve
            Else
                ' Cancel export
                Log.Spacer
                Log.Add T("Export Canceled"), , , "Red", True
                Operation.ErrorLevel = eelCritical
                GoTo CleanUp
            End If
        End If
    End With

    ' Check to see if we still have an item to export.
    If dCategories.Count = 0 Then
        Log.Add T("Skipped after conflict resolution."), , , "blue", True
    Else
        For Each varCategory In dCategories.Keys
            Set dCategory = dCategories.Item(varCategory)
            Set dObjects = dCategory.Item("Objects")
            For Each varObject In dObjects.Keys
                Set cDbObject = dObjects.Item(varObject)
                Operation.Pulse

                ' If we have already exported this object while scanning for changes, use that copy.
                strTempFile = Replace(cDbObject.SourceFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
                If FSO.FileExists(strTempFile) Then
                    ' Move the temp file(s) over to the source export folder.
                    cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, _
                        FSO.GetParentFolderName(cDbObject.SourceFile) & PathSep
                    ' Update the index with the values from the alternate export
                    VCSIndex.UpdateFromAltExport cDbObject
                Else
                    ' Export a fresh copy
                    cDbObject.Export
                End If
                ExportDependentObjects cDbObject
            Next
        Next
    End If

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

CleanUp:

    ' Run any cleanup routines
    CloseBackEndConnections
    ClearEnvCache
    VCSIndex.ClearTempExportFolder

    ' Add performance data to log file and save log
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
        .Flush
    End With

    ' Save index file (don't change export date for multiple items export)
    VCSIndex.Save

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportDependentObjects
' Author    : Adam Waller
' Date      : 6/18/2025
' Purpose   : When exporting a selected object, it may be helpful to also export other
'           : dependent objects. For example, when selecting a table, we may also want
'           : to export table data (if applicable), and table data macros.
'---------------------------------------------------------------------------------------
'
Private Sub ExportDependentObjects(cDbObject As IDbComponent)

    Dim dObjects As Dictionary
    Dim cCategory As IDbComponent
    Dim cItem As IDbComponent
    Dim varKey As Variant

    ' Special cases based on component type
    Select Case cDbObject.ComponentType
        Case edbTableDef    ' Selected table

            ' Table Data
            Set cCategory = New clsDbTableData
            Set dObjects = cCategory.GetAllFromDB(True)
            For Each varKey In dObjects.Keys
                Set cItem = dObjects(varKey)
                If cItem.Name = cDbObject.Name Then
                    ' Found matching name.
                    cItem.Export
                    Exit For
                End If
            Next varKey

            ' Table Data Macro
            Set cCategory = New clsDbTableDataMacro
            Set dObjects = cCategory.GetAllFromDB(True)
            For Each varKey In dObjects.Keys
                Set cItem = dObjects(varKey)
                If cItem.Name = cDbObject.Name Then
                    ' Found matching name.
                    cItem.Export
                    Exit For
                End If
            Next varKey
    End Select

    ' Hidden attribute may apply to any selected object
    Set cCategory = New clsDbHiddenAttribute
    If cCategory.IsModified Then
        ' Since we only store this property if the item is hidden, we should
        ' export the hidden objects source file if any changes are detected.
        cCategory.Export
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportSchemas
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Export any external database schemas configured in options
'---------------------------------------------------------------------------------------
'
Public Sub ExportSchemas(blnFullExport As Boolean)

    Dim varKey As Variant
    Dim strName As String
    Dim strType As String
    Dim cSchema As IDbSchema
    Dim dParams As Dictionary
    Dim strFile As String
    Dim lngCount As Long

    ' Skip this section if there are no connections defined.
    If Options.SchemaExports.Count = 0 Then Exit Sub

    ' Loop through schemas
    Log.Spacer
    Log.Add T("Scanning external databases...")
    Perf.OperationStart "Scan External Databases"
    For Each varKey In Options.SchemaExports.Keys
        strName = varKey

        ' Load parameters for initializing the connection
        Set dParams = GetSchemaInitParams(strName)
        If dParams("Enabled") = False Then
            Log.Add T(" - {0} - Connection disabled", var0:=strName), False
        ElseIf dParams("Connect") = vbNullString Then
            Log.Add " - " & strName, False
            Log.Add T("   No connection string found. (.env)"), , , "Red", , True
            Log.Error eelWarning, T("File not found: {0}", var0:=strFile), ModuleName & ".ExportSchemas"
            Log.Add T("Set the connection string for this external database connection in VCS options to automatically create this file."), False
            Log.Add T("(This file may contain authentication credentials and should be excluded from version control.)"), False
        Else
            ' Show server type along with name
            Select Case Options.SchemaExports(varKey)("DatabaseType")
                Case eDatabaseServerType.estMsSql
                    strType = " (MSSQL)"
                    Set cSchema = New clsSchemaMsSql
                Case eDatabaseServerType.estMySql
                    strType = " (MySQL)"
                    Set cSchema = New clsSchemaMySql
            End Select
            Log.Add " - " & strName & strType
            Perf.CategoryStart strName & strType
            Log.Flush

            ' Export/sync the server objects
            cSchema.Initialize dParams
            cSchema.Export blnFullExport
            lngCount = cSchema.ObjectCount(True)
        End If
        Perf.CategoryEnd lngCount

        ' Check for error
        If Operation.ErrorLevel = eelCritical Then Exit For
    Next varKey
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveThemeZipFiles
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Removes any existing theme zip files. We don't run this inline because
'           : the extraction runs asychrously through the OS and we don't want to slow
'           : down the export waiting for each extraction to complete.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveThemeZipFiles()
    Dim strFolder As String
    If Options.ExtractThemeFiles Then
        strFolder = Options.GetExportFolder & "themes" & PathSep
        If FSO.FolderExists(strFolder) Then ClearFilesByExtension strFolder, "zip"
    End If
End Sub
