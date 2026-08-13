Attribute VB_Name = "modBuild"
'---------------------------------------------------------------------------------------
' Module    : modBuild
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Build and merge functions for importing source files into a database.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modDatabase, modFileAccess, modVCSUtility,
'           : modLoadFromText, modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName As String = "modBuild"

' Set while a merge is preparing the database in place, to limit the crash trace to the
' operation that needs it. (See TraceInPlaceMerge.)
Private m_blnTraceInPlaceMerge As Boolean

' Set when the VBA project reset stage could not reset the project, so that the merge
' falls back to a database reopen when it resumes. These survive the timer stages because
' the add-in's own project is never the one being reset.
Private m_blnInPlaceResetFailed As Boolean

' Set when the in-place preparation confirmed that the database was accessible to other
' clients, so that a merge which then imports nothing can skip the post-merge check.
Private m_blnVerifiedAccessible As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Build (Full build or Merge Build)
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Build the project from source files.
'           : blnResumed is set only when a merge build re-enters this procedure on a
'           : fresh call stack, after PrepareMergeInPlace has prepared the database and
'           : ResetProjectForInPlaceMerge has reset its VBA project. (See the merge
'           : preparation section below.) It preserves the log and performance timers
'           : started by the first call.
'           :
'           : dOutcome is an out-param for unattended callers (see
'           : clsVersionControl.BuildHeadless). Operation.Finish releases the Log and
'           : Perf singletons, so a caller that returns a result to a pipeline cannot
'           : read the log path or the error counts afterwards -- they have to be
'           : captured here, on the way out. Populated on every path that reaches
'           : CleanUp; a merge that hands off to the reset timer returns before that,
'           : which is why headless merges force the reopen path.
'---------------------------------------------------------------------------------------
'
Public Sub Build(strSourceFolder As String, blnFullBuild As Boolean _
                , Optional intFilter As eContainerFilter = ecfAllObjects _
                , Optional strAlternatePath As String _
                , Optional blnResumed As Boolean _
                , Optional ByRef dOutcome As Dictionary)

    Const FunctionName As String = ModuleName & ".Build"

    Dim strPath As String
    Dim strBackup As String
    Dim blnNoChanges As Boolean
    Dim strCurrentDbFilename As String
    Dim cCategory As IDbComponent
    Dim dCategories As Dictionary
    Dim varCategory As Variant
    Dim dCategory As Dictionary
    Dim dFiles As Dictionary
    Dim dScanMeta As Dictionary
    Dim colCategories As Collection
    Dim varFile As Variant
    Dim strType As String
    Dim blnSuccess As Boolean
    Dim blnPrepared As Boolean
    Dim lngCount As Long
    Dim lngCurrent As Long
    Dim cModule As clsDbModule

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Close any previous cached connections
    CloseCachedConnections
    CloseBackEndConnections
    ClearEnvCache
    ClearConnState

    ' The type of build will be used in various messages and log entries.
    strType = IIf(blnFullBuild, T("Build"), T("Merge"))

    ' We need to check the current db name later, so we need to cache it (especially for builds).
    strCurrentDbFilename = CurrentProject.FullName

    ' Make sure we can find the source files
    If Not FolderHasVcsOptionsFile(strSourceFolder) Then
        MsgBox2 T("Source files not found") _
            , T("Required source files were not found in the following folder:"), strSourceFolder, vbExclamation
        GoTo CleanUp
    End If

    ' Verify that the source files are being merged into the correct database.
    strPath = GetOriginalDbFullPathFromSource(strSourceFolder)
    If strPath = vbNullString Then
        MsgBox2 T("Unable to determine database file name.") _
            , T("Required source files were not found or could not be parsed: "), strSourceFolder, vbExclamation
        GoTo CleanUp

    ElseIf strCurrentDbFilename = vbNullString Then
        ' No database currently open. Proceed with build

    ElseIf StrComp(strPath, strCurrentDbFilename, vbTextCompare) <> 0 Then
        If blnFullBuild Then
            ' Full build allows you to use source file name.
            If Not MsgBox2(T("Current Database filename does not match source filename."), _
                    T("Do you want to {0} to the Source Defined Filename?" & vbNewLine & vbNewLine & _
                        "Current: {1}" & vbNewLine & _
                        "Source: {2}", var0:=strType, var1:=strCurrentDbFilename, var2:=strPath), _
                    T("[Ok] = Build with Source Configured Name") & vbNewLine & vbNewLine & _
                        T("Otherwise cancel and select 'Build As...' from the ribbon to change build name. " & _
                        "Performing an export from this file name will also reset the file name, but will " & _
                        "overwrite source. If this file stared as a copy of an existing source controlled " & _
                        "database, select 'Build As...' to avoid overwriting."), _
                    vbQuestion + vbOKCancel + vbDefaultButton1, _
                    T("{0} Name Conflict", var0:=strType), _
                    vbOK) = vbOK Then

                ' Launch the GUI form (it was closed a moment ago)
                DoCmd.OpenForm "frmVCSMain"
                Form_frmVCSMain.StartBuild blnFullBuild
                Log.Error eelCritical, T("{0} aborted. Name mismatch.", var0:=strType), FunctionName
                GoTo CleanUp
            End If
        Else
            MsgBox2 T("Cannot {0} to a different database.", var0:=strType) _
                , T("The database file name for the source files must match the currently open database.") _
                , T("Current: {0}" & vbNewLine & _
                    "Source: {1}", var0:=strCurrentDbFilename, var1:=strPath), vbExclamation _
                , T("{0} Name Conflict", var0:=strType) _
                , vbOK
            GoTo CleanUp
        End If
    End If

    ' Additional checks when a database is currently open.
    If DatabaseFileOpen Then
        ' For full builds, close the current database if it is currently open.
        If blnFullBuild Then
            ' Attempt to close the current database after staging the main form
            If IsLoaded(acForm, "frmVCSMain") Then StageMainForm
            CloseCurrentDatabase2
            ' If the database is still open, then we have a problem that we can't resolve here.
            If DatabaseFileOpen Then
                MsgBox2 T("Unable to Close Database"), _
                    T("The current database must be closed to perform a full build."), , vbExclamation
                Operation.Result = eorFailed
                GoTo CleanUp
            Else
                ' Restore main form as we continue the build
                RestoreMainForm
            End If
        End If
    End If

    ' Load options from project
    Set Options = Nothing
    Options.LoadOptionsFromFile StripSlash(strSourceFolder) & PathSep & "vcs-options.json"
    ' Temporarily override the export folder to always read files from the specified source folder.
    ' (This is needed if the source folder is renamed, or when building to an alternate file.)
    Options.ExportFolder = strSourceFolder
    If Operation.Source = eosMCPTool Or Operation.Source = eosExternalAPI Then
        Options.LoadOptionOverrides
    End If

    ' Update VBA debug mode after loading options
    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Start log and performance timers before merge prep so those messages are preserved.
    ' (A resumed merge keeps the log and timers from the call that prepared the database,
    '  so the preparation entries stay in the same log and performance report.)
    If Not blnResumed Then
        m_blnTraceInPlaceMerge = False
        m_blnVerifiedAccessible = False
        Log.Clear
        Log.SourcePath = strSourceFolder
        Log.Active = True
        Perf.StartTiming
    End If

    ' Build original file name for database
    If blnFullBuild Then
        ' Use alternate path if provided, otherwise extract the original database path from the source files.
        strPath = Nz2(strAlternatePath, GetOriginalDbFullPathFromSource(strSourceFolder))
        If strPath = vbNullString Then
            MsgBox2 T("Unable to determine database file name") _
                , T("Required source files were not found or could not be parsed:"), strSourceFolder, vbExclamation
            GoTo CleanUp
        End If
    Else
        ' All objects must be closed and unloaded before source files are merged in,
        ' since most objects are deleted and reimported. Closing and shift-opening the
        ' database guarantees this, but on a large database it is one of the most
        ' expensive parts of a merge. When the user opts in, prepare the database in
        ' place instead, and resume the merge on a fresh call stack.
        If blnResumed Then
            ' The database was prepared, and its VBA project reset, before this call stack
            ' existed. (See PrepareMergeInPlace and ResetProjectForInPlaceMerge.) Nothing
            ' to do here but honor a reset that did not succeed.
            TraceInPlaceMerge "merge stage: resumed on fresh stack"
            If m_blnInPlaceResetFailed Then
                Log.Add T("Unable to reset the VBA project for the current database.")
                Log.Add T("Falling back to closing and reopening the database.")
                m_blnTraceInPlaceMerge = False
                ReopenBeforeMerge strPath
            ElseIf Not FlushVbaProjectAfterReset Then
                ReopenBeforeMerge strPath
            End If
        ElseIf Options.SkipReopenBeforeMerge And dOutcome Is Nothing Then
            ' The in-place path hands off to a timer and returns on this stack, so it
            ' can never report an outcome to a synchronous caller. A caller that passed
            ' dOutcome is waiting for a result, so it takes the reopen path instead --
            ' slower, but it runs to completion here. (See BuildHeadless.)
            '
            ' Crash tracing is off. The in-place merge and VBA project save are settled
            ' (see DECISIONS.md 2026-07-29), so the trace noise and the log rewrite that
            ' each entry costs are not worth paying on every merge. Uncomment to trace a
            ' fault: this one line re-enables every TraceInPlaceMerge call site, including
            ' the reset and save traces in modVbeUtility, which receive this same flag.
            ' This is also the seam for a future trace-logging option.
            'm_blnTraceInPlaceMerge = True
            m_blnInPlaceResetFailed = False
            Log.Add T("Preparing current database for merge...")
            Log.Flush
            Perf.OperationStart "Prepare Merge In Place"
            blnPrepared = PrepareMergeInPlace
            Perf.OperationEnd
            If blnPrepared Then
                ' Hand off to the reset stage and let this call stack unwind completely.
                ' Nothing is finished or torn down here: the staged operation, log, and
                ' timers are picked up by the merge stage that follows the reset.
                Log.Flush
                Operation.Stage
                TraceInPlaceMerge "prep: staging reset timer"
                SetTimer "MergeReset", strSourceFolder, CStr(CLng(intFilter))
                Exit Sub
            Else
                ' Fall back to the reliable path.
                m_blnTraceInPlaceMerge = False
                ReopenBeforeMerge strPath
            End If
        Else
            ReopenBeforeMerge strPath
        End If

        ' Run any pre-merge instructions after the database has been prepared
        ' with all objects closed/unloaded.
        If Options.RunBeforeMerge <> vbNullString Then
            Log.Add T("Running {0}...", var0:=Options.RunBeforeMerge)
            Log.Flush
            Perf.OperationStart "RunBeforeMerge"
            RunSubInCurrentProject Options.RunBeforeMerge
            Perf.OperationEnd
            CatchAny eelError, T("Error running {0}", var0:=Options.RunBeforeMerge), FunctionName, True, True
        End If
    End If

    ' Launch the GUI form, unless the caller asked for silence. Opening it makes it
    ' visible (StartBuild -> ResetForOperation sets Me.Visible), so an unattended
    ' build would flash a window on a runner's desktop for no benefit. Gated on the
    ' interaction mode rather than on Operation.Source, because an API or MCP build
    ' is still watched by someone unless it explicitly said otherwise, and the form
    ' is where they watch it. Everything below tolerates the form being absent.
    If Operation.InteractionMode <> eimSilent Then
        DoCmd.OpenForm "frmVCSMain"
        Form_frmVCSMain.StartBuild blnFullBuild
    End If

    ' Minimize the VBE window to prevent it from stealing focus
    ' when VBA components are imported during the build.
    TraceInPlaceMerge "phase: minimizing VBE window"
    MinimizeVBEWindow
    TraceInPlaceMerge "phase: VBE window minimized"

    ' Display the build header.
    DoCmd.Hourglass True
    With Log
        .Spacer
        If blnFullBuild Then
            .Add T("Beginning build from Source"), False
        Else
            .Add T("Beginning merge from source"), False
        End If
        .Add FSO.GetFileName(strPath)
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=strPath), False
        .Add T("Export Folder: {0}", var0:=strSourceFolder), False
        ' Log operation source (file only, not console)
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
        .Flush
    End With

    ' Check project VCS version
    If Options.CompareLoadedVersion = evcNewerVersion Then
        If MsgBox2(T("Newer VCS Version Detected"), _
            T("This project uses VCS version {0} but version {1} is currently installed." & _
                    vbNewLine & "Would you like to continue anyway?" _
                , var0:=Options.GetLoadedVersion, var1:=GetVCSVersion), _
            T("Click YES to continue this operation, or NO to cancel."), _
            vbExclamation + vbYesNo + vbDefaultButton2, , vbYes) <> vbYes Then
            Operation.ErrorLevel = eelCritical
            GoTo CleanUp
        End If
    End If

    ' Rename original file as a backup
    strBackup = GetBackupFileName(strPath)
    If blnFullBuild Then
        If FSO.FileExists(strPath) Then
            Log.Add T("Saving backup of original database...")
            Name strPath As strBackup
            If CatchAny(eelCritical, T("Unable to rename original file"), FunctionName) Then GoTo CleanUp
            Log.Add T("Saved as {0}.", var0:=FSO.GetFileName(strBackup))
        End If
    Else
        ' Backups for merge builds performed later,
        ' but only if we have changes we are actually merging.
    End If

    ' Create a new database with the original name
    If blnFullBuild Then
        Perf.OperationStart "Create new database"
        If LCase$(FSO.GetExtensionName(strPath)) = "adp" Then
            ' ADP project
            Application.NewAccessProject strPath
        Else
            ' Regular Access database
            Application.NewCurrentDatabase strPath, GetFileFormat(strSourceFolder)
        End If
        Perf.OperationEnd
        If DatabaseFileOpen Then
            Log.Add T("Created blank database for import. (v{0})", var0:=DbVersion)
        Else
            CatchAny eelCritical, T("Unable to create database file"), FunctionName
            Log.Add T("This may occur when building an older database version if the " & _
                "'New database sort order' (collation) option is not set to 'Legacy'")
            GoTo CleanUp
        End If
    End If

    ' Reset LoadFromText state now that the target database is open and its folder path is known.
    modLoadFromText.Reset

    ' Now that we have a new database file, we can load the index.
    Set VCSIndex = Nothing

    If blnFullBuild Then
        ' Remove any non-built-in references before importing from source.
        Log.Add T("Removing non built-in references..."), False
        RemoveNonBuiltInReferences

        ' Check for any RunBeforeBuild
        If Options.RunBeforeBuild <> vbNullString Then
            ' Run any pre-build bootstrapping code
            PrepareRunBootstrap
        End If
    End If

    ' Warm persistent connections to linked Access back-end files before merge
    ' conflict temp-exports (same as full export operations).
    CacheBackEndConnections

    ' Build collections of files to import/merge
    Log.Add T("Scanning source files...")
    Log.Flush

    ' Remove misplaced duplicate module/form/report copies before scanning (agent/git drift).
    Dim cModuleCategory As IDbComponent
    Dim cFormCategory As IDbComponent
    Dim cReportCategory As IDbComponent
    Set cModuleCategory = New clsDbModule
    Set cFormCategory = New clsDbForm
    Set cReportCategory = New clsDbReport
    RemoveDuplicateModuleFiles cModuleCategory.BaseFolder
    RemoveDuplicateFormFiles cFormCategory.BaseFolder
    RemoveDuplicateReportFiles cReportCategory.BaseFolder

    Set dCategories = New Dictionary
    VCSIndex.Conflicts.Initialize dCategories, eatImport

    ' Walk the source folders once for file dates and sizes, and share the result with
    ' every category below. Letting each category scan for itself re-walks the whole
    ' source tree repeatedly, since several component types report the export root as
    ' their BaseFolder. (A full build imports every file regardless, so it never scans.)
    Set colCategories = GetContainers(intFilter)
    If Not blnFullBuild Then Set dScanMeta = GetSharedScanMetadata(colCategories)

    Perf.OperationStart "Scan Source Files"
    For Each cCategory In colCategories
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        Operation.Pulse
        ' Get collection of source files
        If blnFullBuild Then
            ' Return all the source files
            dCategory.Add "Files", cCategory.GetFileList
        Else
            ' Merge build
            If cCategory.ComponentType = edbTableData And Not Options.MergeTableData Then
                ' Reconciling table data on a merge is optional, since it changes records
                ' rather than object definitions.
                Log.Add T("Not merging {0}. (Merge table data option is turned off)", _
                    var0:=T(LCase(cCategory.Category))), Options.ShowDebug
                dCategory.Add "Files", New Dictionary
            Else
                ' Return just the modified source files for merge, including source file paths
                ' representing orphaned objects that no longer exist in the database.
                dCategory.Add "Files", VCSIndex.GetModifiedSourceFiles(cCategory, dScanMeta)
            End If
        End If
        ' Check count of modified source files.
        If dCategory("Files").Count = 0 Then
            Log.Add T(IIf(blnFullBuild, "No {0} source files found.", "No modified {0} source files found."), _
                var0:=T(LCase(cCategory.Category))), Options.ShowDebug
        Else
            dCategories.Add cCategory.Category, dCategory
            ' For merge builds, check for import conflicts or orphaned database objects
            If Not blnFullBuild Then
                ' Record any conflicts for later review
                VCSIndex.CheckMergeConflicts cCategory, dCategory("Files")
            End If
        End If
        ' Check for critical error or cancel
        If Operation.ErrorLevel = eelCritical Then
            Log.Add vbNullString
            Perf.OperationEnd
            GoTo CleanUp
        End If
    Next cCategory
    Perf.OperationEnd

    ' Check for any conflicts
    With VCSIndex.Conflicts
        If .Count > 0 Then
            ' Resolve conflicts (auto-resolve for agent/API, prompt for user)
            .ResolveOrPrompt
            If .ApproveResolutions Then
                Log.Add T("Resolving source conflicts"), False
                .Resolve
            Else
                ' Cancel build/merge
                Log.Spacer
                Log.Add T("Build Canceled")
                Operation.ErrorLevel = eelCritical
                GoTo CleanUp
            End If
        End If
    End With

    ' A merge may not find any changed files
    If dCategories.Count = 0 And Not blnFullBuild Then
        blnNoChanges = True
        Log.Add T("No changes found.")
    Else
        ' Perform a backup if we have changes to merge
        If Not blnFullBuild Then
            LogUnhandledErrors
            Log.Add T("Saving backup of original database...")
            FSO.CopyFile strPath, strBackup
            If CatchAny(eelCritical, T("Unable to back up current database"), FunctionName) Then GoTo CleanUp
            Log.Add T("Saved as {0}.", var0:=FSO.GetFileName(strBackup))
        End If
        Log.Spacer
    End If

    ' Loop through all categories
    For Each varCategory In dCategories.Keys

        ' Set reference to object category class and file list
        Set cCategory = dCategories(varCategory)("Class")
        Set dFiles = dCategories(varCategory)("Files")

        ' Show category header
        Log.Spacer Options.ShowDebug
        Log.PadRight T(IIf(blnFullBuild, "Importing {0}...", "Merging {0}..."), _
            var0:=T(LCase(cCategory.Category))), , Options.ShowDebug
        Perf.CategoryStart cCategory.Category
        lngCount = dFiles.Count
        lngCurrent = 0
        Log.Flush
        TraceInPlaceMerge "merge: " & cCategory.Category & " (" & lngCount & " files)"

        ' Loop through each file in this category.
        If blnFullBuild And cCategory.ComponentType = edbModule Then

            Set cModule = cCategory

            For Each varFile In dFiles.Keys
                lngCurrent = lngCurrent + 1
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                Log.Progress lngCurrent, lngCount, FSO.GetFileName(varFile)
                Operation.Pulse
                cModule.ImportFast CStr(varFile)
                CatchAny eelError, T("Build error in: {0}", var0:=varFile), FunctionName, True, True
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp
            Next varFile

            cModule.FinalizeImports
            CatchAny eelError, T("Build error finalizing modules"), FunctionName, True, True
            If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp

        Else

            For Each varFile In dFiles.Keys
                ' Import/merge the file
                lngCurrent = lngCurrent + 1
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                Log.Progress lngCurrent, lngCount, FSO.GetFileName(varFile)
                Operation.Pulse
                If blnFullBuild Then
                    cCategory.Import CStr(varFile)
                Else
                    cCategory.Merge CStr(varFile)
                    If Options.ExportAfterMerge Then
                        ' Merging imports the object, which then makes it available
                        ' to export from this category/object class.
                        ' (Forms are exported later after initializing)
                        If cCategory.ComponentType <> edbForm Then cCategory.Export
                    End If
                End If
                CatchAny eelError, T(IIf(blnFullBuild, "Build error in: {0}", "Merge error in: {0}"), _
                    var0:=varFile), FunctionName, True, True

                ' Bail out if we hit a critical error.
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp

            Next varFile

        End If

        ' Show category wrap-up.
        PadTableDataMergeCompleteIfNeeded cCategory
        If Options.ShowDebug Then
            Log.Add T("[{0}] {1} processed.", var0:=dFiles.Count, var1:=T(LCase(cCategory.Category)))
        Else
            Log.Add "[" & dFiles.Count & "]"
        End If
        Perf.CategoryEnd dFiles.Count
        ReleaseDbReferences
        TraceInPlaceMerge "merge: " & cCategory.Category & " complete"

    Next varCategory
    TraceInPlaceMerge "phase: merge loop complete"

    If Operation.ErrorLevel <> eelCritical Then PromptAndSaveConnections

    ' Check for merge items that might affect other components
    If Not blnFullBuild Then
        ' Check for any object visible in the object navigation pane that might have a description property.
        If ContainerHasAnyObject(dCategories, _
            edbAdpFunction, edbAdpServerView, edbAdpStoredProcedure, edbAdpTable, edbAdpTrigger, _
            edbForm, edbMacro, edbModule, edbQuery, edbReport, edbTableData, edbTableDataMacro, edbTableDef) Then
            ' Merge any changes to the document properties (i.e. description)
            Log.Add T("Merging any changed document properties..."), Options.ShowDebug
            MergeIfChanged edbDocument
        End If
    End If

    ' Reopen the database so the themes are loaded
    If ContainerHasObject(dCategories, edbTheme) Then
        Log.Add T("Reopening database...")
        Log.Flush
        TraceInPlaceMerge "phase: reopening for themes"
        StageMainForm
        CloseCurrentDatabase2
        ShiftOpenDatabase strPath
        RestoreMainForm
        TraceInPlaceMerge "phase: reopened for themes"
    End If

    ' Initialize forms to ensure that the colors/themes are rendered properly
    ' (This must be done after all objects are imported, since subforms/subreports
    '  may be involved, and must already exist in the database.)
    If ContainerHasObject(dCategories, edbForm) Then
        Log.Add T("Initializing forms...")
        TraceInPlaceMerge "phase: initializing forms"
        InitializeForms dCategories
        TraceInPlaceMerge "phase: forms initialized"
    End If

    ' Update operation result in case this is queried in the AfterBuild hooks
    ' Assume success if we have not jumped to the cleanup.
    Operation.Result = eorSuccess

    ' Run any post-build/merge instructions
    If blnFullBuild Then
        If Options.RunAfterBuild <> vbNullString Then
            Log.Add T("Running {0}...", var0:=Options.RunAfterBuild)
            Log.Flush
            Perf.OperationStart "RunAfterBuild"
            RunSubInCurrentProject Options.RunAfterBuild
            Perf.OperationEnd
        End If
    Else
        ' Merge build
        If Options.RunAfterMerge <> vbNullString Then
            Log.Add T("Running {0}...", Options.RunAfterMerge)
            Log.Flush
            Perf.OperationStart "RunAfterMerge"
            RunSubInCurrentProject Options.RunAfterMerge
            Perf.OperationEnd
        End If
    End If

    ' Enforce any supplied letter casing rules.
    '
    ' Full builds only. Applying corrections saves the VBA project, and saving locks the
    ' database against other clients, so the accessibility check below then reopens it --
    ' measured at 83 seconds on a merge that imported nothing and corrected one identifier.
    ' A merge does not need it: the code it brings in comes from source files that the
    ' export pass already standardized, and export standardizes again before writing, so
    ' source consistency never depends on the merge. At worst the database carries
    ' non-canonical casing until the next export corrects it.
    If blnFullBuild Then
        Dim colCasingChanges As Collection
        Dim varChange As Variant
        Set colCasingChanges = StandardizeLetterCasing
        If Not colCasingChanges Is Nothing Then
            If colCasingChanges.Count > 0 Then
                Log.Add T("{0} letter casing correction(s) applied:", var0:=colCasingChanges.Count), False
                For Each varChange In colCasingChanges
                    Log.Add "  " & varChange, False
                Next varChange
            End If
        End If
    End If

    ' Log any errors after build/merge
    CatchAny eelError, T("Error running {0}", var0:=CallByName(Options, "RunAfter" & strType, VbGet)), FunctionName, True, True

    ' Validate the build. Unlike the RunAfter* hooks above, this one gates success:
    ' a build that produced every object but cannot actually run is a failed build,
    ' and a deployment pipeline needs to hear about it here rather than from users.
    If Options.ValidateAfterBuild <> vbNullString Then
        If Not RunBuildValidation() Then
            Operation.ErrorLevel = eelCritical
            GoTo CleanUp
        End If
    End If

    ' If the database is not accessible to other clients, reopen it in shared mode.
    ' Uses an out-of-process worker to detect the engine-level lock state that an
    ' in-process check cannot see, and reopens unconditionally when the worker is
    ' unavailable to probe it.
    '
    ' The check itself costs a worker round trip, and there is one case where its answer
    ' is already known: an in-place merge verifies accessibility before it begins, so if
    ' it then imported nothing, nothing has happened since that could have changed it.
    If DatabaseFileOpen And Not (blnNoChanges And m_blnVerifiedAccessible) Then
        If Not DatabaseAccessibleToOtherClients Then
            Log.Add T("Reopening database in shared mode...")
            Log.Flush
            Perf.OperationStart "Reopen DB (shared mode)"
            ReleaseDbReferences
            StageMainForm
            CloseCurrentDatabase2
            ShiftOpenDatabase strPath
            RestoreMainForm
            Perf.OperationEnd
        End If
    End If

    ' Show final output and save log
    Dim dMissing As Dictionary
    Dim varEnvKey As Variant
    Dim strColor As String
    If Log.ErrorCount > 0 Then
        strColor = "red"
    ElseIf Log.WarningCount > 0 Then
        strColor = "#CC7700"
    Else
        strColor = "green"
    End If
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, strColor, True
    blnSuccess = True

    ' Show warning/error summary if any issues occurred
    If Log.WarningCount > 0 Or Log.ErrorCount > 0 Then
        Log.Add vbNullString
        Log.Spacer
        If Log.WarningCount > 0 And Log.ErrorCount > 0 Then
            Log.Add T("{0} warning(s), {1} error(s)", _
                var0:=Log.WarningCount, var1:=Log.ErrorCount), , , strColor, True
        ElseIf Log.WarningCount > 0 Then
            Log.Add T("{0} warning(s)", var0:=Log.WarningCount), , , strColor, True
        Else
            Log.Add T("{0} error(s)", var0:=Log.ErrorCount), , , strColor, True
        End If
        If Log.ErrorCount > 0 Then
            Log.Add T("See log for details."), , , strColor
        End If
        ' List missing .env keys if any
        Set dMissing = GetMissingEnvKeys
        If dMissing.Count > 0 Then
            Log.Add vbNullString
            With New clsConcat
                .AppendOnAdd = ", "
                For Each varEnvKey In dMissing.Keys
                    .Add CStr(varEnvKey)
                Next varEnvKey
                .Remove 2
                Log.Add T("Missing .env keys: {0}", var0:=.GetStr), , , strColor
            End With
            Log.Add T("Ensure these keys are defined in: {0}", _
                var0:=GetEnvFilePath), , , strColor
        End If
        Log.Spacer
    End If

CleanUp:

    ' Close cached connections
    CloseCachedConnections
    CloseBackEndConnections
    ClearEnvCache
    ClearConnState

    ' Add performance data to log file and save file.
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
    End With

    ' Show message if build failed
    If Operation.ErrorLevel = eelCritical Or Not blnSuccess Then
        Log.Spacer
        Log.Add T("Build Failed."), , , "red", True
        Log.Flush
    End If

    ' Capture the outcome for an unattended caller before Operation.Finish releases
    ' the Log and Perf singletons. (Log.SaveFile above is what sets SavedLogFilePath.)
    If Not dOutcome Is Nothing Then
        With dOutcome
            .Item("success") = blnSuccess And (Operation.ErrorLevel <> eelCritical)
            .Item("logPath") = Log.SavedLogFilePath
            .Item("errorCount") = Log.ErrorCount
            .Item("warningCount") = Log.WarningCount
            .Item("durationMs") = CLng(Perf.TotalTime * 1000)
            If DatabaseFileOpen Then .Item("databasePath") = CurrentProject.FullName
        End With
    End If

    ' Wrap up build.
    DoCmd.Hourglass False
    If IsLoaded(acForm, "frmVCSMain") Then
        ' Finish up on GUI
        Form_frmVCSMain.FinishBuild blnFullBuild, blnSuccess
    Else
        ' Allow navigation pane to refresh list of objects.
        DoEvents
    End If

    ' Save index file after build is complete, or discard index for "Build As..."
    ' discard update if build failed.
    If strAlternatePath = vbNullString And blnSuccess Then
        If blnFullBuild Then
            ' NOTE: Add a couple seconds since some items may still be in the process of saving.
            VCSIndex.FullBuildDate = DateAdd("s", 2, Now)
        Else
            VCSIndex.MergeBuildDate = DateAdd("s", 2, Now)
        End If
        VCSIndex.Save strSourceFolder
    End If
    Set VCSIndex = Nothing

    ' A failed build must not finish as "complete". Operation.Result is only set to
    ' eorSuccess on the path that reaches the end of the build, so every failure that
    ' jumped to CleanUp still carries eorUnknown -- which Finish maps to a "complete"
    ' MCP callback. Say what actually happened before finishing.
    If Not blnSuccess Or Operation.ErrorLevel = eelCritical Then Operation.Result = eorFailed

    ' Wait to finish the build till after we have saved the index.
    Operation.Finish

    ' Show MessageBox if not using GUI for build (skip for API/MCP operations).
    If Forms.Count = 0 And blnSuccess _
        And Operation.Source = eosUserInterface Then
        MsgBox2 T("Build Complete for '{0}'", var0:=CurrentProject.Name), _
            T("Note that some settings may not take effect until this database is reopened."), _
            T("A backup of the previous build was saved as '{0}'.", var0:=FSO.GetFileName(strBackup)), vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunBuildValidation
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Run the Options.ValidateAfterBuild function in the freshly built database
'           : and report whether it approved the build. Anything other than an explicit
'           : True is a failure: a missing procedure, a raised error, no return value,
'           : or a return that is not a Boolean. Silence must not read as approval when
'           : the answer gates a deployment.
'---------------------------------------------------------------------------------------
'
Private Function RunBuildValidation() As Boolean

    Const FunctionName As String = ModuleName & ".RunBuildValidation"

    Dim strProc As String
    Dim varResult As Variant
    Dim blnRan As Boolean
    Dim blnPassed As Boolean
    Dim lngErrorsBefore As Long

    strProc = Options.ValidateAfterBuild
    lngErrorsBefore = Log.ErrorCount

    Log.Add T("Validating build with {0}...", var0:=strProc)
    Log.Flush
    Perf.OperationStart "ValidateAfterBuild"
    varResult = RunProcInCurrentProject(strProc, blnRan)
    Perf.OperationEnd

    ' RunProcInCurrentProject logs and clears whatever the hook raised, and logs when the
    ' procedure cannot be found, so the error count is how both come back to us.
    If Not blnRan Or Log.ErrorCount > lngErrorsBefore Then
        Log.Add T("Build validation failed: unable to run {0}.", var0:=strProc), , , "red", True
        Exit Function
    End If

    If IsEmpty(varResult) Then
        Log.Error eelError, T("{0} did not return a value.", var0:=strProc), FunctionName
        Log.Add T("The validation procedure must be a Function that returns True on success."), False
        Exit Function
    End If

    LogUnhandledErrors
    On Error Resume Next
    blnPassed = CBool(varResult)
    If CatchAny(eelError, T("{0} returned a value that is not True or False.", var0:=strProc), _
        FunctionName) Then blnPassed = False
    On Error GoTo 0

    If blnPassed Then
        Log.Add T("Build validation passed.")
    Else
        Log.Add T("Build validation failed: {0} returned False.", var0:=strProc), , , "red", True
    End If

    RunBuildValidation = blnPassed

End Function


'---------------------------------------------------------------------------------------
' Procedure : TraceInPlaceMerge
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Crash-trace breadcrumb, active only while a merge is preparing the database
'           : in place. That path manipulates the VBA project and can fault inside
'           : VBE7.DLL, taking Access down without unwinding, so its progress has to
'           : reach disk step by step. Every other build and merge path skips the writes.
'---------------------------------------------------------------------------------------
'
Public Sub TraceInPlaceMerge(strStep As String)
    If m_blnTraceInPlaceMerge Then LogCrashTrace strStep
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatabaseAccessibleToOtherClients
' Author    : Adam Waller
' Date      : 8/7/2026
' Purpose   : Report whether external clients can open the current database file, and
'           : assume they cannot when there is no way to find out.
'           :
'           : The JET/ACE engine does not expose this lock state to same-process callers,
'           : so the only reliable test is the out-of-process worker probe. When the user
'           : has disabled the helper script (#727) no probe is possible, and every caller
'           : here treats "not accessible" as the safe answer: the post-build check
'           : reopens in shared mode, and the in-place merge falls back to the reopen path
'           : it used before the probe existed. Returning False without launching anything
'           : states that intent, rather than leaving it to depend on a skipped worker job
'           : yielding an empty result.
'---------------------------------------------------------------------------------------
'
Private Function DatabaseAccessibleToOtherClients() As Boolean
    If modInstall.UseWorkerScript Then
        DatabaseAccessibleToOtherClients = Worker.IsDatabaseAccessible
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetProjectForInPlaceMerge
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Clear the target database's VBA project run-state as its own timer stage.
'           :
'           : This exists as a separate stage on purpose. The reset has to happen when as
'           : little as possible is alive: not merely on a stack the target project does
'           : not own, but before the merge has set anything up. Doing it inside Build
'           : meant the whole of Build's prologue — reloaded options, reacquired database
'           : handles, resolved paths — was already live when the project was reset, and
'           : the merge then continued to use references taken before it. The caller
'           : stages the main form (releasing the form instance and the log console) and
'           : releases cached references before arming this stage, so by the time the
'           : reset runs, only this frame and the add-in's own singletons remain.
'           :
'           : Nothing may run on this call stack after the reset — see the caller, which
'           : arms the next stage before calling this. Executing the VBE Reset control
'           : returns immediately, but the teardown it triggers lands later, when the
'           : thread next reaches a message pump. Both observed crashes were in the first
'           : substantial work done after the Execute on the same stack: continuing the
'           : merge in an earlier design, then RestoreMainForm in this one. Opening a form
'           : pumps messages, so it collided with the teardown. Cheap statements such as
'           : recording the result below have proven survivable; anything that pumps has
'           : not, so it belongs on the next stack.
'           :
'           : Records failure rather than reporting it, because the merge that resumes on
'           : the next stage is the thing that has to fall back to a database reopen. The
'           : main form is not restored here either: the merge stage reopens it, and
'           : frmVCSMain.ResetForOperation rebinds the log console and clears the console
'           : text anyway, so there is nothing worth restoring first.
'---------------------------------------------------------------------------------------
'
Public Sub ResetProjectForInPlaceMerge()
    TraceInPlaceMerge "reset stage: resetting project (nothing runs after this)"
    m_blnInPlaceResetFailed = Not ResetCurrentVBProjectState(m_blnTraceInPlaceMerge)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FlushVbaProjectAfterReset
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Save the target database's VBA project once its run-state has been cleared,
'           : so that unsaved VBE edits are not overwritten by the source files about to be
'           : merged in. Returns False when the merge should fall back to a reopen instead.
'           :
'           : This runs *after* the reset rather than during the preparation so the save
'           : cannot encounter run-state in the project it is writing. (An earlier theory
'           : held that run-state was what defeated the save outright; that turned out to
'           : be the caller's own VBA stack instead -- see SaveCurrentVBProject. Saving
'           : after the reset remains the safer order regardless.) Nothing is lost by
'           : waiting: the reset clears run-state, not editor buffers, so unsaved edits are
'           : still present afterwards.
'           :
'           : Accessibility is re-checked only when the save actually wrote something. A
'           : write is the only thing that can lock the database, and the preparation
'           : already confirmed the database was accessible, so a save that changes nothing
'           : cannot have invalidated that. This keeps the common case (a clean project,
'           : where the save is a no-op) free of a second worker round trip.
'---------------------------------------------------------------------------------------
'
Private Function FlushVbaProjectAfterReset() As Boolean

    Dim blnWasDirty As Boolean

    FlushVbaProjectAfterReset = True

    blnWasDirty = Not CurrentVBProject.Saved
    If Not blnWasDirty Then Exit Function

    TraceInPlaceMerge "merge stage: saving VBA project after reset"
    If Not SaveCurrentVBProject(m_blnTraceInPlaceMerge) Then
        ' Nothing was written, so the database cannot have been locked by us. Proceed, but
        ' record it: any unsaved edits in objects the merge replaces will be overwritten.
        Log.Add T("Note: unsaved VBA changes could not be saved before merging."), False
        Exit Function
    End If

    ' The save wrote the project. Confirm the database is still usable by other clients
    ' before the merge reaches its backup, which an exclusive lock would block.
    TraceInPlaceMerge "merge stage: rechecking lock state after save"
    If DatabaseFileOpen Then
        If Not DatabaseAccessibleToOtherClients Then
            Log.Add T("Database is not accessible to other clients after saving VBA changes.")
            Log.Add T("Falling back to closing and reopening the database.")
            m_blnTraceInPlaceMerge = False
            FlushVbaProjectAfterReset = False
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReopenBeforeMerge
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Close and shift-open the current database so that every object is closed
'           : and unloaded before source files are merged into it. This is the default
'           : (and most reliable) way to prepare for a merge.
'---------------------------------------------------------------------------------------
'
Private Sub ReopenBeforeMerge(strPath As String)
    Log.Add T("Closing and reopening current database before merge...")
    Perf.OperationStart "Reopen DB before Merge"
    StageMainForm
    CloseCurrentDatabase2
    ShiftOpenDatabase strPath
    RestoreMainForm
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PrepareMergeInPlace
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Prepare the current database for a merge without closing and reopening it,
'           : reaching the same starting conditions by other means:
'           :
'           :  1. Every open object is closed (merge deletes and reimports objects).
'           :  2. Pending VBE edits are flushed, so unsaved code is not lost.
'           :  3. References that would be invalidated are released.
'           :
'           : Clearing the VBA project's run-state is the fourth requirement, and it is
'           : what makes this safe rather than merely faster: importing a component into
'           : a project that holds run-state resets that project implicitly, part way
'           : through the merge, invalidating references the merge is still using (the
'           : crash in DECISIONS.md 2026-07-06). Doing it deliberately and up front,
'           : with nothing yet cached, replaces an uncontrolled reset with a controlled
'           : one.
'           :
'           : The reset itself deliberately does NOT happen here — see
'           : ResetProjectForInPlaceMerge, which runs as its own timer stage once this
'           : call stack has unwound. Staging the main form is the last step here so that
'           : the add-in's console is released before that stage runs.
'           :
'           : Returns True when the database is ready to merge in place. Returns False
'           : if any step could not be completed, in which case the caller must fall
'           : back to ReopenBeforeMerge rather than merge in an unknown state.
'---------------------------------------------------------------------------------------
'
Private Function PrepareMergeInPlace() As Boolean

    Const FunctionName As String = ModuleName & ".PrepareMergeInPlace"

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' The reset cannot be survived when the project it acts on is the one running this
    ' code, so there is no in-place path to offer while the add-in is open as the current
    ' database. (This is the normal workflow for developing the add-in itself.)
    TraceInPlaceMerge "prep: checking host project"
    If ResetWouldEndOurOwnCode Then
        Log.Add T("Cannot prepare in place while the add-in is open as the current database.")
        GoTo FallBack
    End If

    ' Close all open database objects. A cancelled or failed close means we cannot
    ' guarantee that objects are unloaded.
    TraceInPlaceMerge "prep: closing open objects"
    If Not CloseDatabaseObjects Then
        Log.Add T("Unable to close all open objects.")
        GoTo FallBack
    End If

    ' Release every reference the reset would invalidate.
    TraceInPlaceMerge "prep: releasing cached references"
    ReleaseScanState
    TraceInPlaceMerge "prep: references released"

    If CatchAny(eelWarning, T("Error preparing database for merge"), FunctionName, True, True) Then GoTo FallBack

    ' A database that other clients cannot open has to be reopened whether we do it now or
    ' the post-merge check does it later, so there is no saving left to protect and we take
    ' the proven path instead. Reopening now is strictly better than reopening after: the
    ' shift-open leaves the database accessible, so the post-merge check then finds nothing
    ' to do, and the mid-merge backup (critical on failure, and blocked by an exclusive
    ' lock) is not attempted against a locked file. Checked here rather than on entry
    ' because the steps above can escalate the lock, so this is the only point that
    ' reflects the state the merge would actually run in.
    TraceInPlaceMerge "prep: checking lock state"
    If DatabaseFileOpen Then
        If DatabaseAccessibleToOtherClients Then
            m_blnVerifiedAccessible = True
        Else
            Log.Add T("Database is not accessible to other clients.")
            GoTo FallBack
        End If
    End If

    ' Release the add-in's own console: this drops the form instance and the log's
    ' reference to its controls, the same way it is released around a database reopen.
    ' The staged content is restored after the reset. Deliberately the last step, so that
    ' the fallback path never has to unpick a staged form.
    TraceInPlaceMerge "prep: staging main form"
    StageMainForm

    PrepareMergeInPlace = True
    Exit Function

FallBack:
    Log.Add T("Falling back to closing and reopening the database.")

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseScanState
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Release every reference that a VBA project reset or a database
'           : close/reopen would invalidate, so that work can safely continue on the
'           : other side of that boundary.
'           :
'           : Two earlier attempts to avoid the pre-merge reopen crashed Access because
'           : references survived the boundary: an in-place project reset (DECISIONS.md
'           : 2026-07-06) and a deferred reopen that reused the scan's component classes
'           : (reverted in 0e4b93b0, which named this helper as the prerequisite).
'           :
'           : Component classes cache database objects internally, so a category
'           : dictionary built by a scan cannot be carried across the boundary and
'           : reused. Pass it here to drop the class references; the source file paths
'           : it holds are plain strings and remain valid.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseScanState(Optional dCategories As Dictionary)

    Const FunctionName As String = ModuleName & ".ReleaseScanState"

    Dim varCategory As Variant
    Dim dCategory As Dictionary

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Drop component classes built against the current instance of the database.
    If Not dCategories Is Nothing Then
        For Each varCategory In dCategories.Keys
            Set dCategory = dCategories(varCategory)
            If Not dCategory Is Nothing Then
                If dCategory.Exists("Class") Then Set dCategory("Class") = Nothing
            End If
        Next varCategory
    End If

    ' Release cached database handles and connections.
    ReleaseDbReferences
    CloseCachedConnections
    CloseBackEndConnections
    ClearEnvCache
    ClearConnState

    ' Discard cached state tied to the current database file location.
    modLoadFromText.Reset

    If Err Then Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadSingleObject
' Author    : Adam Waller
' Date      : 2/23/2023
' Purpose   : Reload a single object from source files.
'           : NOTE: Be very careful to release all references to the object you
'           : are attempting to import.
'           : When blnNoIndex is True, the VCS index is disabled for the duration
'           : of the call, skipping the expensive full-file parse/serialize cycle
'           : and conflict detection. Used by MCP/API callers that treat the import
'           : as a deliberate action (like a user saving directly in the designer).
'           : strSavedLogPath returns the log file written by this call. Callers
'           : cannot read Log.SavedLogFilePath afterwards, since Operation.Finish
'           : releases the Log singleton before this returns.
'---------------------------------------------------------------------------------------
'
Public Sub LoadSingleObject(cComponentClass As IDbComponent, strName As String, _
    strSourceFilePath As String, Optional blnNoIndex As Boolean = False, _
    Optional ByRef strSavedLogPath As String)

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim dSourceFiles As Dictionary
    Dim intResult As eOperationResult

    ' Guard clauses
    If cComponentClass Is Nothing Then Exit Sub
    If Not FSO.FileExists(strSourceFilePath) Then Exit Sub

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure the object is currently closed. (This is really important, since we
    ' will be deleting the object before adding it from source.)
    With cComponentClass
        Select Case .ComponentType
            Case acForm, acMacro, acModule, acQuery, acReport, acTable
                If SysCmd(acSysCmdGetObjectState, .ComponentType, strName) <> adStateClosed Then
                    DoCmd.Close .ComponentType, strName, acSavePrompt
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
        .Add T("Beginning Import of Single Object"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        ' Log operation source (file only, not console)
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
        .Add T("Importing {0}...", var0:=strName)
        .Flush
    End With

    If Not blnNoIndex Then
        ' Check for conflicts
        Set dSourceFiles = New Dictionary
        Set dCategory = New Dictionary
        Set dCategories = New Dictionary
        dSourceFiles.Add strSourceFilePath, vbNullString
        dCategory.Add "Class", cComponentClass
        dCategory.Add "Files", dSourceFiles
        dCategories.Add cComponentClass, dCategory
        VCSIndex.Conflicts.Initialize dCategories, eatImport
        VCSIndex.CheckMergeConflicts cComponentClass, dSourceFiles

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
                    Log.Add T("Import Canceled"), , , "Red", True
                    Operation.ErrorLevel = eelCritical
                    intResult = eorCanceled
                    GoTo CleanUp
                End If
            End If
        End With

        ' Check to see if we still have an item to import.
        If dCategories.Count = 0 Then
            Log.Add T("Skipped after conflict resolution."), , , "blue", True
            GoTo PostMerge
        End If
    End If

    ' Replace the existing object with the source file
    cComponentClass.Merge strSourceFilePath
    MergeDependentObjects cComponentClass, strName

PostMerge:

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True
    intResult = eorSuccess

CleanUp:

    ' Run any cleanup routines. This runs even when the index is disabled, since the
    ' table merge compares against a copy exported into this folder either way.
    VCSIndex.ClearTempExportFolder

    ' Save the index before timing stops, otherwise the "Save Index" timer it starts
    ' never reaches the report and the cost of writing the whole file is invisible.
    If blnNoIndex Then
        VCSIndex.Disabled = False
    ElseIf Not VCSIndex.Conflicts.UserCanceled Then
        ' Save index file (don't change export date for single item export).
        ' Skipped if the user canceled a conflict dialog so the same conflicts
        ' will reappear on the next run.
        VCSIndex.Save
    End If

    ' Add performance data to log file and save file
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        strSavedLogPath = .SavedLogFilePath
        .Active = False
        .Flush
    End With

    Operation.Finish intResult

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PadTableDataMergeCompleteIfNeeded
' Author    : Adam Waller
' Date      : 8/13/2026
' Purpose   : When table data printed child lines under the category heading, pad a
'           : completion sentence so the wrap-up [N] aligns with other categories.
'---------------------------------------------------------------------------------------
'
Private Sub PadTableDataMergeCompleteIfNeeded(cCategory As IDbComponent)
    If Options.ShowDebug Then Exit Sub
    If cCategory.ComponentType <> edbTableData Then Exit Sub
    If Not Log.AtNewLine Then Exit Sub
    Log.PadRight T("Table data merge complete.")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeDependentObjects
' Author    : Adam Waller
' Date      : 6/18/2025
' Purpose   : Merge in any dependent objects related to the selected object.
'           : (I.e. table data for a selected table)
'---------------------------------------------------------------------------------------
'
Private Sub MergeDependentObjects(cComponentClass As IDbComponent, strName As String)

    Dim cItem As clsDbTableData
    Dim strFile As String
    Dim intFormat As eTableDataExportFormat

    ' Special cases based on component type
    Select Case cComponentClass.ComponentType

        ' Table object
        Case edbTableDef

            ' Table Data
            Set cItem = New clsDbTableData
            If Options.TablesToExportData.Exists(strName) Then
                ' Convert string format option to enum value
                intFormat = Options.GetTableExportFormat(dNZ(Options.TablesToExportData, strName & "\Format"))
                If intFormat > etdNoData Then
                    ' Set a reference to the table object so the table data class can build the source file name.
                    Set cItem.Parent.DbObject = CurrentData.AllTables(strName)
                    cItem.Format = intFormat
                    strFile = cItem.Parent.SourceFile
                    If FSO.FileExists(strFile) Then
                        Log.Add T("Merging table data for {0}", , , , strName), Options.ShowDebug
                        ' Merge, not Import. Import appends (XML) or deletes every row and
                        ' reloads (tab-delimited), which was safe only while the definition
                        ' merge above always deleted and recreated the table first. It no
                        ' longer does -- a definition that already matches source is left
                        ' alone, rows and all -- so appending would duplicate rows or trip
                        ' the primary key, and the delete-and-reload would fail against any
                        ' table referenced by a relationship. Merge reconciles against the
                        ' key instead, and inserts everything when the table is empty
                        ' because the definition really was rebuilt.
                        cItem.Parent.Merge strFile
                    End If
                End If
            End If

            ' Table Data Macro
            ' (Already loaded with table definition)
    End Select

    ' Could consider merging hidden attribute here if requested.
    ' (We don't need to add the complexity unless there is an actual need for this.)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeAllSource
' Author    : Adam Waller
' Date      : 5/16/2023
' Purpose   : Forcibly merge all source files into the current database. This is used
'           : in testing to confirm that we can successfully merge all types of source
'           : files into the database. (Not something an end user would normally use.)
'---------------------------------------------------------------------------------------
'
Public Sub MergeAllSource()

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim cCategory As IDbComponent
    Dim varCategory As Variant
    Dim dFiles As Dictionary
    Dim varFile As Variant
    Dim lngCount As Long
    Dim lngCurrent As Long

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure all database objects are currently closed (This is really important,
    ' since we will be deleting most objects before importing them from source.)
    CloseDatabaseObjects

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
        .Add T("Beginning Merge of All Source Files"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        ' Log operation source (file only, not console)
        If Len(Operation.SourceName) > 0 Then .Add T("Source: {0}", var0:=Operation.SourceName), False
        .Add Now
        .Spacer
        .Add T("Scanning source files...")
        .Flush
    End With

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
        Log.Add T("Merge Canceled"), , , "Red", True
        Log.Flush
        Operation.ErrorLevel = eelCritical
        Exit Sub
    End If

    ' Build collections of files to import/merge
    Set dCategories = New Dictionary
    Perf.OperationStart "Scan Source Files"
    For Each cCategory In GetContainers
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        dCategory.Add "Files", cCategory.GetFileList
        dCategories.Add cCategory, dCategory
    Next cCategory
    Perf.OperationEnd


    ' Loop through all categories
    Log.Spacer
    For Each varCategory In dCategories.Keys

        ' Set reference to object category class
        Set cCategory = varCategory
        Set dFiles = dCategories(varCategory)("Files")

        ' Only show category details when source files are found
        If dFiles.Count = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add T("No {0} source files found.", var0:=LCase(cCategory.Category)), Options.ShowDebug
        Else
            ' Show category header
            Log.Spacer Options.ShowDebug
            Log.PadRight T("Merging ") & LCase(cCategory.Category) & "...", , Options.ShowDebug
            Perf.CategoryStart cCategory.Category
            lngCount = dFiles.Count
            lngCurrent = 0
            Log.Flush

            ' Loop through each file in this category.
            For Each varFile In dFiles.Keys
                ' Import/merge the file
                lngCurrent = lngCurrent + 1
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                Log.Progress lngCurrent, lngCount, FSO.GetFileName(varFile)
                Operation.Pulse
                cCategory.Merge CStr(varFile)
                CatchAny eelError, T("Merge error in: {0}", var0:=varFile), ModuleName & ".MergeAllSource", True, True

                ' Bail out if we hit a critical error.
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp
            Next varFile

            ' Show category wrap-up.
            PadTableDataMergeCompleteIfNeeded cCategory
            Log.Add "[" & dFiles.Count & "]" & IIf(Options.ShowDebug, " " & LCase(cCategory.Category) & T(" processed."), vbNullString)
            Perf.CategoryEnd dFiles.Count
        End If
    Next varCategory

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

CleanUp:

    ' Run any cleanup routines
    VCSIndex.ClearTempExportFolder

    ' Add performance data to log file and save file
    Perf.EndTiming
    With Log
        .Add vbNewLine & Perf.GetReports, False
        .SaveFile
        .Active = False
        .Flush
    End With

    ' Save index file (don't change export date for single item export)
    VCSIndex.Save

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeScoped
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Category-scoped merge: merge source files for the named container(s)
'           : into the current database, reconciling deletions (DB objects no longer
'           : represented in source) within each category. This is a middle tier
'           : between LoadSingleObject (one object, no orphan reconciliation) and
'           : Build/MergeAllSource (full project with backup and global front-matter).
'           :
'           : blnFullMerge controls the file set: True = all source files (GetFileList)
'           : unioned with orphan-deletion entries; False = changed source files only
'           : (GetModifiedSourceFiles, which already includes orphan-deletion entries).
'           :
'           : Assumes the caller has already begun an Operation, loaded Options/Index/
'           : Log, and started Perf timing. Does NOT take a database backup.
'---------------------------------------------------------------------------------------
'
Public Sub MergeScoped(colContainers As Collection, blnFullMerge As Boolean)

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim dFiles As Dictionary
    Dim dScanMeta As Dictionary
    Dim varCategory As Variant
    Dim varFile As Variant
    Dim cCategory As IDbComponent
    Dim cItem As IDbComponent
    Dim dItems As Dictionary
    Dim lngCount As Long
    Dim lngCurrent As Long
    Dim intSave As AcCloseSave
    Dim strPath As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Remove misplaced duplicate copies for targeted layout categories only
    For Each cCategory In colContainers
        Select Case cCategory.ComponentType
            Case edbModule
                RemoveDuplicateModuleFiles cCategory.BaseFolder
            Case edbForm
                RemoveDuplicateFormFiles cCategory.BaseFolder
            Case edbReport
                RemoveDuplicateReportFiles cCategory.BaseFolder
        End Select
    Next cCategory

    ' Close open objects of the targeted categories before merging
    If Operation.InteractionMode = eimNormal _
        And Operation.Source <> eosMCPTool _
        And Operation.Source <> eosExternalAPI Then
        intSave = acSavePrompt
    Else
        intSave = acSaveYes
    End If

    If Not CloseOpenObjectsForContainers(colContainers, intSave) Then
        Log.Spacer
        Log.Add T("Merge Canceled"), , , "Red", True
        Operation.ErrorLevel = eelCritical
        Exit Sub
    End If

    SaveUnsavedVbaProjectIfNeeded colContainers
    CacheBackEndConnections

    ' Build collections of files to import/merge
    Set dCategories = New Dictionary
    VCSIndex.Conflicts.Initialize dCategories, eatImport

    ' One shared file date/size map for the targeted categories (see Build).
    If Not blnFullMerge Then Set dScanMeta = GetSharedScanMetadata(colContainers)

    Perf.OperationStart "Scan Source Files"
    For Each cCategory In colContainers
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        Operation.Pulse
        If blnFullMerge Then
            Set dFiles = cCategory.GetFileList
            Set dItems = cCategory.GetAllFromDB
            If Not cCategory.SingleFile Then
                For Each varFile In dItems.Items
                    Set cItem = varFile
                    If Not dFiles.Exists(cItem.SourceFile) Then
                        dFiles.Add cItem.SourceFile, vbNullString
                    End If
                Next varFile
            End If
            dCategory.Add "Files", dFiles
        Else
            If cCategory.ComponentType = edbTableData Then
                ' Unreachable in practice: ComponentTypeSupportsScopedImport rejects table
                ' data before a scoped import starts. Kept as a guard because a scoped
                ' import takes no backup, unlike a merge build.
                Log.Add T("Not merging {0}. (Imported only on full build)", _
                    var0:=T(LCase(cCategory.Category))), Options.ShowDebug
                dCategory.Add "Files", New Dictionary
            Else
                dCategory.Add "Files", VCSIndex.GetModifiedSourceFiles(cCategory, dScanMeta)
            End If
        End If
        If dCategory("Files").Count = 0 Then
            Log.Add T(IIf(blnFullMerge, "No {0} source files found.", "No modified {0} source files found."), _
                var0:=T(LCase(cCategory.Category))), Options.ShowDebug
        Else
            dCategories.Add cCategory.Category, dCategory
            If Not blnFullMerge Then
                VCSIndex.CheckMergeConflicts cCategory, dCategory("Files")
            End If
        End If
        If Operation.ErrorLevel = eelCritical Then
            Log.Add vbNullString
            Perf.OperationEnd
            Exit Sub
        End If
    Next cCategory
    Perf.OperationEnd

    ' Check for any conflicts
    With VCSIndex.Conflicts
        If .Count > 0 Then
            .ResolveOrPrompt
            If .ApproveResolutions Then
                Log.Add T("Resolving source conflicts"), False
                .Resolve
            Else
                Log.Spacer
                Log.Add T("Merge Canceled"), , , "Red", True
                Operation.ErrorLevel = eelCritical
                Exit Sub
            End If
        End If
    End With

    ' A merge may not find any changed files
    If dCategories.Count = 0 And Not blnFullMerge Then
        Log.Add T("No changes found.")
    End If

    ' Loop through all categories and merge
    For Each varCategory In dCategories.Keys
        Set cCategory = dCategories(varCategory)("Class")
        Set dFiles = dCategories(varCategory)("Files")

        Log.Spacer Options.ShowDebug
        Log.PadRight T(IIf(blnFullMerge, "Importing {0}...", "Merging {0}..."), _
            var0:=T(LCase(cCategory.Category))), , Options.ShowDebug
        Perf.CategoryStart cCategory.Category
        lngCount = dFiles.Count
        lngCurrent = 0
        Log.Flush

        For Each varFile In dFiles.Keys
            lngCurrent = lngCurrent + 1
            Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
            Log.Progress lngCurrent, lngCount, FSO.GetFileName(varFile)
            Operation.Pulse
            cCategory.Merge CStr(varFile)
            If Not blnFullMerge And Options.ExportAfterMerge Then
                If cCategory.ComponentType <> edbForm Then cCategory.Export
            End If
            CatchAny eelError, T(IIf(blnFullMerge, "Build error in: {0}", "Merge error in: {0}"), _
                var0:=varFile), ModuleName & ".MergeScoped", True, True
            If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: Exit Sub
            If cCategory.SingleFile Then Exit For
        Next varFile

        PadTableDataMergeCompleteIfNeeded cCategory
        If Options.ShowDebug Then
            Log.Add T("[{0}] {1} processed.", var0:=dFiles.Count, var1:=T(LCase(cCategory.Category)))
        Else
            Log.Add "[" & dFiles.Count & "]"
        End If
        Perf.CategoryEnd dFiles.Count
        ReleaseDbReferences
    Next varCategory

    If Operation.ErrorLevel <> eelCritical Then PromptAndSaveConnections

    If Not blnFullMerge Then
        If ContainerHasAnyObject(dCategories, _
            edbAdpFunction, edbAdpServerView, edbAdpStoredProcedure, edbAdpTable, edbAdpTrigger, _
            edbForm, edbMacro, edbModule, edbQuery, edbReport, edbTableData, edbTableDataMacro, edbTableDef) Then
            Log.Add T("Merging any changed document properties..."), Options.ShowDebug
            MergeIfChanged edbDocument
        End If
    End If

    strPath = CurrentProject.FullName
    If ContainerHasObject(dCategories, edbTheme) Then
        Log.Add T("Reopening database...")
        Log.Flush
        StageMainForm
        CloseCurrentDatabase2
        ShiftOpenDatabase strPath
        RestoreMainForm
    End If

    If ContainerHasObject(dCategories, edbForm) Then
        Log.Add T("Initializing forms...")
        InitializeForms dCategories
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSharedScanMetadata
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return a single file date/size map (see ScanFolderMetadata) covering the
'           : base folders of every container about to be scanned for changes, so the
'           : disk is walked once per folder rather than once per category.
'           :
'           : Several component types (project properties, VBE references, connections,
'           : documents, nav pane groups) report the export root as their BaseFolder. A
'           : recursive scan of the root already covers every other category, so as soon
'           : as one of those is in the list this collapses to a single walk. When the
'           : list contains none of them -- a narrowly scoped sync of, say, just menus --
'           : only the folders actually needed are scanned, so a small operation does not
'           : pay for a full-tree walk.
'---------------------------------------------------------------------------------------
'
Private Function GetSharedScanMetadata(colContainers As Collection) As Dictionary

    Dim cCategory As IDbComponent
    Dim dFolders As Dictionary
    Dim dMeta As Dictionary
    Dim dFolderMeta As Dictionary
    Dim varFolder As Variant
    Dim varKey As Variant
    Dim varFolderKeys As Variant
    Dim strRoot As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    strRoot = AddSlash(Options.GetExportFolder)

    ' Collect the distinct base folders, watching for the export root.
    Set dFolders = New Dictionary
    dFolders.CompareMode = TextCompare
    For Each cCategory In colContainers
        If StrComp(AddSlash(cCategory.BaseFolder), strRoot, vbTextCompare) = 0 Then
            ' One recursive scan from the root covers every category.
            Set dFolders = New Dictionary
            dFolders.Add strRoot, vbNullString
            Exit For
        End If
        If Not dFolders.Exists(cCategory.BaseFolder) Then
            dFolders.Add cCategory.BaseFolder, vbNullString
        End If
    Next cCategory

    If dFolders.Count = 1 Then
        ' Single walk (the export root, or a scoped sync of one category). Use the scan
        ' directly rather than copying thousands of entries into a second dictionary.
        varFolderKeys = dFolders.Keys
        Set dMeta = ScanFolderMetadata(CStr(varFolderKeys(0)))
    ElseIf dFolders.Count > 1 Then
        ' Merge the folder scans into one map keyed by full path.
        Set dMeta = New Dictionary
        dMeta.CompareMode = TextCompare
        For Each varFolder In dFolders.Keys
            Set dFolderMeta = ScanFolderMetadata(CStr(varFolder))
            For Each varKey In dFolderMeta.Keys
                If Not dMeta.Exists(varKey) Then dMeta.Add varKey, dFolderMeta(varKey)
            Next varKey
        Next varFolder
    End If

    ' Nothing to scan leaves this Nothing, and callers fall back to their own scan.
    Set GetSharedScanMetadata = dMeta

    CatchAny eelError, T("Error scanning source folder metadata"), _
        ModuleName & ".GetSharedScanMetadata"

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBackupFileName
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return an unused filename for the database backup befor build
'---------------------------------------------------------------------------------------
'
Private Function GetBackupFileName(strPath As String) As String

    Const cstrSuffix As String = "_VCSBackup"

    Dim strFile As String
    Dim intCnt As Integer
    Dim strTest As String
    Dim strBase As String
    Dim strExt As String
    Dim strFolder As String
    Dim strIncrement As String

    strFolder = FSO.GetParentFolderName(strPath) & PathSep
    strFile = FSO.GetFileName(strPath)
    strBase = FSO.GetBaseName(strFile) & cstrSuffix
    strExt = "." & FSO.GetExtensionName(strFile)

    ' Attempt up to 500 versions of the file name. (i.e. Database_VSBackup45.accdb)
    For intCnt = 1 To 500
        strTest = strFolder & strBase & strIncrement & strExt
        If FSO.FileExists(strTest) Then
            ' Try next number
            strIncrement = CStr(intCnt)
        Else
            ' Return file name
            GetBackupFileName = strTest
            Exit Function
        End If
    Next intCnt

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileFormat
' Author    : Adam Waller
' Date      : 5/7/2021
' Purpose   : Return the file format version from the source files, or 0 if not found.
'---------------------------------------------------------------------------------------
'
Private Function GetFileFormat(strSourcePath As String) As Long

    Dim strPath As String

    ' Attempt to read the file format version from the CurrentProject export
    strPath = StripSlash(strSourcePath) & PathSep & "project.json"
    GetFileFormat = dNZ(ReadJsonFile(strPath), "Items\FileFormat")

End Function


'---------------------------------------------------------------------------------------
' Procedure : PrepareRunBootstrap
' Author    : Adam Waller
' Date      : 4/21/2021
' Purpose   : Prepares the database to run the RunBeforeBuild code by loading all
'           : GUID references and importing the module specified in RunBeforeBuild.
'           : The bootstrap module (and any other objects) will get replaced from
'           : source during the main build, but this allows any custom functions to
'           : run before the main build, such as copying missing library files into
'           : the same folder as the database.
'---------------------------------------------------------------------------------------
'
Private Sub PrepareRunBootstrap()

    Dim strModule As String
    Dim strName As String
    Dim varFile As Variant

    ' Update output since there may be some delays
    Log.Add T("Loading bootstrap...")
    Log.Flush
    Perf.OperationStart "Bootstrap"

    ' Load all GUID references to support early binding in bootstrap sub
    With New clsDbVbeReference
        .ImportReferences .Parent.SourceFile, True
    End With

    ' Identify and load module for bootstrap code
    strModule = Split(Options.RunBeforeBuild, ".")(0)
    With New clsDbModule
        With .Parent
            For Each varFile In .GetFileList
                ' Look for matching name
                strName = GetObjectNameFromFileName(CStr(varFile))
                If StrComp(strName, strModule, vbTextCompare) = 0 Then
                    ' This is the module we need to import
                    Log.Add T("Importing bootstrap module '{0}'", var0:=strName), False
                    .Import CStr(varFile)
                    Exit For
                End If
            Next varFile
        End With
    End With

    ' Make sure we actually have a module before we attempt to run the code
    If CurrentProject.AllModules.Count = 0 Then
        ' Could not find source file
        Log.Error eelError, T("Could not find source file for {0}", var0:=strModule), ModuleName & ".PrepareRunBootstrap"
    Else
        ' Important: We need to Run Project.Sub not Project.Module.Sub
        strName = Split(Options.RunBeforeBuild, ".")(1)

        ' Run any pre-build bootstrapping code
        Log.Add T("Running {0}", var0:=Options.RunBeforeBuild)
        Perf.OperationStart "RunBeforeBuild"
        RunSubInCurrentProject strName
        Perf.OperationEnd
    End If

    ' Now go back and remove all the non built-in references so they come
    ' back in the correct order, just in case a library was at a higher level.
    Log.Add T("Removing non built-in references after running bootstrap"), False
    RemoveNonBuiltInReferences

    Perf.OperationEnd   ' Bootstrap

End Sub


'---------------------------------------------------------------------------------------
' Procedure : InitializeForms
' Author    : Adam Waller
' Date      : 7/2/2021
' Purpose   : Opens and closes each form in design view to complete the process of
'           : fully rendering the colors and applying the theme. (This is needed to
'           : provide a consistent output after importing from source.)
'           : Pass this function the dictionary of container of objects being
'           : imported into the database. (All object types)
'---------------------------------------------------------------------------------------
'
Public Sub InitializeForms(dContainers As Dictionary)

    Dim frm As IDbComponent
    Dim dFiles As Dictionary
    Dim dAllForms As Dictionary
    Dim cAllForms As IDbComponent
    Dim varKey As Variant
    Dim blnIsAddin As Boolean
    Dim lngCount As Long
    Dim lngCurrent As Long

    ' Trap any errors that may occur when opening forms
    LogUnhandledErrors
    On Error Resume Next

    ' See if we imported any forms
    Set cAllForms = New clsDbForm
    If dContainers.Exists(cAllForms.Category) Then

        ' Are we working on the add-in project itself?
        blnIsAddin = (CurrentVBProject.Name = PROJECT_NAME)

        ' Get reference to forms container
        Set dFiles = dContainers(cAllForms.Category)("Files")
        lngCount = dFiles.Count
        lngCurrent = 0

        ' Loop through the forms in the current database
        Set dAllForms = cAllForms.GetAllFromDB
        For Each varKey In dAllForms.Keys

            ' See if this form matches one of the files we just imported
            Set frm = dAllForms(varKey)
            If dFiles.Exists(frm.SourceFile) Then
                lngCurrent = lngCurrent + 1

                ' Don't attempt to initialize add-in main form
                ' (Likely not needed, and would require staging)
                If frm.Name <> "frmVCSMain" Then

                    ' Open the form in design view to initialize layout, colors and theme
                    Perf.OperationStart "Initialize Forms"
                    Log.Add "  " & frm.Name, Options.ShowDebug
                    Log.Progress lngCurrent, lngCount, frm.Name
                    If blnIsAddin Then
                        OpenFormInCurrentDb frm.Name, acDesign, , , , acHidden
                    Else
                        DoCmd.OpenForm frm.Name, acDesign, , , , acHidden
                    End If
                    DoEvents
                    ' Set a property value so Access thinks we have something to save.
                    Forms(frm.Name).TAG = Forms(frm.Name).TAG    ' (This doesn't actually change anything)
                    ' Save and close the form with the recomputed geometry
                    DoCmd.Close acForm, frm.Name, acSaveYes
                    Perf.OperationEnd
                End If

                ' Log any errors
                CatchAny eelError, T("Error while initializing form {0}", var0:=frm.Name), ModuleName & ".InitializeForms"

                ' Update the index, since the save date may have changed, but reuse the code hash
                ' since we just calculated it after importing the form.
                With VCSIndex.Item(frm)
                    VCSIndex.Update frm, eatImport, .FileHash, .OtherHash, _
                        strMetaHash:=.MetaHash
                End With

                ' For merge operations, we might be also exporting after initializing
                If Operation.OperationType = eotMerge And Options.ExportAfterMerge Then
                    frm.Export
                End If
            End If
        Next varKey
    End If

    ' Check for any unhandled errors
    CatchAny eelError, "Unhandled error while initializing forms", ModuleName & ".InitializeForms"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenFormInCurrentDb
' Author    : Adam Waller
' Date      : 6/24/2025
' Purpose   : Open a form from the current database instead of the add-in, when forms
'           : with the same names exist in both places.
'           : IMPORTANT: Note that FilterName and WhereCondition have been changed from
'           : Variant to String so that the subsequent arguments are not discarded in
'           : the call to Application.Run. (It appears that once a missing argument is
'           : identified, all subsequent arguments are ignored.)
'---------------------------------------------------------------------------------------
'
Private Sub OpenFormInCurrentDb(FormName, Optional View As AcFormView = acNormal, Optional FilterName As String, _
    Optional WhereCondition As String, Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, Optional OpenArgs)

    Dim strCmd As String

    LogUnhandledErrors
    On Error Resume Next

    ' Build out command targeting the current database's OpenForm2 wrapper
    strCmd = CurrentProject.Path & PathSep & FSO.GetBaseName(CurrentProject.Name) & ".OpenForm2"

    ' Run in current database, passing in all parameters
    Application.Run strCmd, FormName, View, FilterName, WhereCondition, DataMode, WindowMode, OpenArgs

    ' When the target database predates the OpenForm2 wrapper (cross-version build),
    ' Application.Run raises error 2517 (procedure not found). In that case, fall back
    ' to opening the form directly. Any other error is left set for the caller to report.
    If Catch(2517) Then
        DoCmd.OpenForm FormName, View, FilterName, WhereCondition, DataMode, WindowMode, OpenArgs
    End If

End Sub
