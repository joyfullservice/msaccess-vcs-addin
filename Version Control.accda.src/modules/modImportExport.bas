Attribute VB_Name = "modImportExport"
'---------------------------------------------------------------------------------------
' Module    : modImportExport
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Main export/import/merge functions for add-in.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName As String = "modImportExport"


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
    Log.Clear
    Log.SourcePath = Options.GetExportFolder
    Log.Active = True
    Perf.StartTiming

    ' Check error handling mode after loading project options
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' If options (or VCS version) have changed, a full export will be required
    If (VCSIndex.OptionsHash <> Options.GetHash) Then blnFullExport = True

    ' Display heading
    With Log
        .Spacer
        .Add T("Beginning Export of Source Files"), False
        .Add CurrentProject.Name
        .Add T("VCS Version {0}", var0:=GetVCSVersion)
        .Add T("Full Path: {0}", var0:=CurrentProject.FullName), False
        .Add T("Export Folder: {0}", var0:=Options.GetExportFolder), False
        .Add IIf(blnFullExport, T("Performing Full Export"), T("Using Fast Save"))
        .Add Now
        ' Save the log file path
        If Not frmMain Is Nothing Then frmMain.strLastLogFilePath = .LogFilePath
    End With

    ' Check VBE Project protection
    If CurrentVBProject.Protection = vbext_pp_locked Then
        MsgBox2 T("Project Locked"), _
            T("Project is protected with a password."), _
            T("Please unlock the project before using this tool."), vbExclamation
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

    ' Perform any needed upgrades to source files
    If blnFullExport Then UpgradeSourceFiles

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

    ' Export any external database schemas
    ExportSchemas blnFullExport
    If Operation.ErrorLevel = eelCritical Then GoTo CleanUp

    ' Finish header section
    Log.Spacer
    Log.Add IIf(blnFullExport, T("Scanning source files..."), T("Scanning for changes..."))
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
        Set dObjects = cCategory.GetAllFromDB(Not blnFullExport)
        If dObjects.Count = 0 Then
            Log.Add IIf(blnFullExport, _
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
            ' Show the conflicts resolution dialog
            .ShowDialog
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
            Log.ProgMax = lngCount
            Perf.CategoryStart cCategory.Category

            ' Loop through each object in this category.
            For Each varKey In dObjects.Keys

                ' Export object
                Set cDbObject = dObjects(varKey)
                Log.Add "  " & cDbObject.Name, Options.ShowDebug
                Operation.Pulse

                ' If we have already exported this object while scanning for changes, use that copy.
                strTempFile = Replace(cDbObject.SourceFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
                If FSO.FileExists(strTempFile) Then
                    ' Move the temp file(s) over to the source export folder.
                    cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, cDbObject.BaseFolder
                    ' Update the index with the values from the alternate export
                    VCSIndex.UpdateFromAltExport cDbObject
                Else
                    ' Export a fresh copy
                    cDbObject.Export
                End If

                ' Bail out if we hit a critical error.
                CatchAny eelError, T("Error exporting {0}", var0:=cDbObject.Name), ModuleName & ".ExportSource", True, True
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp
                Log.Increment

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
        End If

    Next varCatKey

    ' Ensure that we have created the .gitignore and .gitattributes files in Git environments.
    CheckGitFiles

    ' Run any custom sub after export
    If Options.RunAfterExport <> vbNullString Then
        Log.Add T("Running {0}...", var0:=Options.RunAfterExport)
        Perf.OperationStart "RunAfterExport"
        RunSubInCurrentProject Options.RunAfterExport
        Perf.OperationEnd
        CatchAny eelError, T("Error running {0}", var0:=Options.RunAfterExport), ModuleName & ".ExportSource", True, True
    End If

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True

CleanUp:

    ' Run any cleanup routines
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
        .OptionsHash = Options.GetHash
        .Save
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportSingleObject
' Author    : Adam Waller
' Date      : 2/22/2023
' Purpose   : Export a single object (such as a selected item)
'---------------------------------------------------------------------------------------
'
Public Sub ExportSingleObject(objItem As AccessObject, Optional frmMain As Form_frmVCSMain)

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

    ' Reload the project options and reset the logs
    Set VCSIndex = Nothing
    Set Options = Nothing
    Options.LoadProjectOptions
    Log.Clear
    Log.SourcePath = Options.GetExportFolder
    Log.Active = True
    Perf.StartTiming

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
        .Add Now
        .Spacer
        .Add T("Exporting {0}...", var0:=objItem.Name)
        .Flush
        ' Save export log file path
        If Not frmMain Is Nothing Then frmMain.strLastLogFilePath = .LogFilePath
    End With

    ' Get a database component class from the item
    Set cDbObject = GetClassFromObject(objItem)

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
            ' Show the conflicts resolution dialog
            .ShowDialog
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
        ' If we have already exported this object while scanning for changes, use that copy.
        strTempFile = Replace(cDbObject.SourceFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
        If FSO.FileExists(strTempFile) Then
            ' Move the temp file(s) over to the source export folder.
            cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, cDbObject.BaseFolder
            ' Update the index with the values from the alternate export
            VCSIndex.UpdateFromAltExport cDbObject
        Else
            ' Export a fresh copy
            cDbObject.Export
        End If
    End If

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
        .Add Now
        .Spacer
        .Flush
    End With

    Set dCategories = New Dictionary

    For Each varKey In objItems.Keys
        Set objItem = objItems.Item(varKey)
        Log.Add T("Exporting {0}...", var0:=objItem.Name)
        Log.Flush

        ' FIXME: Hackish, need to figure a clean way of communicating types instead of encoding the key
        Dim lngObjectType As Access.AcObjectType
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
            ' Show the conflicts resolution dialog
            .ShowDialog
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
                    cDbObject.MoveSource FSO.GetParentFolderName(strTempFile) & PathSep, cDbObject.BaseFolder
                    ' Update the index with the values from the alternate export
                    VCSIndex.UpdateFromAltExport cDbObject
                Else
                    ' Export a fresh copy
                    cDbObject.Export
                End If
            Next
        Next
    End If

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

    ' Save index file (don't change export date for multiple items export)
    VCSIndex.Save

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
' Procedure : Build (Full build or Merge Build)
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Build the project from source files.
'---------------------------------------------------------------------------------------
'
Public Sub Build(strSourceFolder As String, blnFullBuild As Boolean _
                , Optional intFilter As eContainerFilter = ecfAllObjects _
                , Optional strAlternatePath As String)

    Const FunctionName As String = ModuleName & ".Build"

    Dim strPath As String
    Dim strBackup As String
    Dim strCurrentDbFilename As String
    Dim cCategory As IDbComponent
    Dim dCategories As Dictionary
    Dim varCategory As Variant
    Dim dCategory As Dictionary
    Dim dFiles As Dictionary
    Dim varFile As Variant
    Dim strType As String
    Dim blnSuccess As Boolean

    Dim strText As String   ' Remove later

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Close the previous cached connections, if any
    CloseCachedConnections

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
            ' If the database file was open in exclusive mode (such as after a build)
            ' we might have to call this function a second time to actually close the file.
            If DatabaseFileOpen Then CloseCurrentDatabase2
            ' If the database is still open, then we have a problem that we can't resolve here.
            If DatabaseFileOpen Then
                MsgBox2 T("Unable to Close Database"), _
                    T("The current database must be closed to perform a full build."), , vbExclamation
                Operation.Result = eorFailed
                GoTo CleanUp
            End If
        End If
    End If

    ' Load options from project
    Set Options = Nothing
    Options.LoadOptionsFromFile StripSlash(strSourceFolder) & PathSep & "vcs-options.json"
    ' Override the export folder when exporting to an alternate path.
    If Len(strAlternatePath) Then Options.ExportFolder = strSourceFolder

    ' Update VBA debug mode after loading options
    LogUnhandledErrors FunctionName
    On Error Resume Next

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
        ' Run any pre-merge instructions
        strText = dNZ(Options.GitSettings, "RunBeforeMerge")
        If strText <> vbNullString Then
            Log.Add T("Running {0}...", var0:=strText)
            Perf.OperationStart "RunBeforeMerge"
            RunSubInCurrentProject strText
            Perf.OperationEnd
        End If

        ' Now, just to make sure all objects are closed and unloaded, we will
        ' close and shift-open the database before merging source files into it.
        Log.Add T("Closing and reopening current database before merge...")
        Perf.OperationStart "Reopen DB before Merge"
        CloseCurrentDatabase2
        ShiftOpenDatabase strPath
        Perf.OperationEnd
    End If

    ' Start log and performance timers
    Log.Clear
    Log.SourcePath = strSourceFolder
    Log.Active = True
    Perf.StartTiming

    ' Launch the GUI form
    DoCmd.OpenForm "frmVCSMain"
    Form_frmVCSMain.StartBuild blnFullBuild

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
            vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
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
            Log.Add T("Created blank database for import. (v{0})", var0:=CurrentProject.FileFormat)
        Else
            CatchAny eelCritical, T("Unable to create database file"), FunctionName
            Log.Add T("This may occur when building an older database version if the " & _
                "'New database sort order' (collation) option is not set to 'Legacy'")
            GoTo CleanUp
        End If
    End If

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

    ' Build collections of files to import/merge
    Log.Add T("Scanning source files...")
    Log.Flush
    Set dCategories = New Dictionary
    VCSIndex.Conflicts.Initialize dCategories, eatImport
    Perf.OperationStart "Scan Source Files"
    For Each cCategory In GetContainers(intFilter)
        Set dCategory = New Dictionary
        dCategory.Add "Class", cCategory
        Operation.Pulse
        ' Get collection of source files
        If blnFullBuild Then
            ' Return all the source files
            dCategory.Add "Files", cCategory.GetFileList
        Else
            ' Merge build
            If cCategory.ComponentType = edbTableData Then
                ' Some component types are only imported on full build
                Log.Add T("Not merging {0}. (Imported only on full build)", _
                    var0:=T(LCase(cCategory.Category))), Options.ShowDebug
                dCategory.Add "Files", New Dictionary
            Else
                ' Return just the modified source files for merge, including source file paths
                ' representing orphaned objects that no longer exist in the database.
                dCategory.Add "Files", VCSIndex.GetModifiedSourceFiles(cCategory)
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
            ' Show the conflicts resolution dialog
            .ShowDialog
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
        Log.ProgMax = dFiles.Count
        Perf.CategoryStart cCategory.Category

        ' Loop through each file in this category.
        For Each varFile In dFiles.Keys
            ' Import/merge the file
            Log.Increment
            Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
            Operation.Pulse
            If blnFullBuild Then
                cCategory.Import CStr(varFile)
            Else
                cCategory.Merge CStr(varFile)
            End If
            CatchAny eelError, T(IIf(blnFullBuild, "Build error in: {0}", "Merge error in: {0}"), _
                var0:=varFile), FunctionName, True, True

            ' Bail out if we hit a critical error.
            If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp

        Next varFile

        ' Show category wrap-up.
        If Options.ShowDebug Then
            Log.Add T("[{0}] {1} processed.", var0:=dFiles.Count, var1:=T(LCase(cCategory.Category)))
        Else
            Log.Add "[" & dFiles.Count & "]"
        End If
        Perf.CategoryEnd dFiles.Count

    Next varCategory

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
        StageMainForm
        CloseCurrentDatabase2
        ShiftOpenDatabase strPath
        RestoreMainForm
    End If

    ' Initialize forms to ensure that the colors/themes are rendered properly
    ' (This must be done after all objects are imported, since subforms/subreports
    '  may be involved, and must already exist in the database.)
    If ContainerHasObject(dCategories, edbForm) Then
        Log.Add T("Initializing forms...")
        InitializeForms dCategories
    End If

    ' Update operation result in case this is queried in the AfterBuild hooks
    ' Assume success if we have not jumped to the cleanup.
    Operation.Result = eorSuccess

    ' Run any post-build/merge instructions
    If blnFullBuild Then
        If Options.RunAfterBuild <> vbNullString Then
            Log.Add T("Running {0}...", var0:=Options.RunAfterBuild)
            Perf.OperationStart "RunAfterBuild"
            RunSubInCurrentProject Options.RunAfterBuild
            Perf.OperationEnd
        End If
    Else
        ' Merge build
        If Options.RunAfterMerge <> vbNullString Then
            Log.Add T("Running {0}...", Options.RunAfterMerge)
            Perf.OperationStart "RunAfterMerge"
            RunSubInCurrentProject Options.RunAfterMerge
            Perf.OperationEnd
        End If
    End If

    ' Log any errors after build/merge
    CatchAny eelError, T("Error running {0}", var0:=CallByName(Options, "RunAfter" & strType, VbGet)), FunctionName, True, True

    ' Show final output and save log
    Log.Spacer
    Log.Add T("Done. ({0} seconds)", var0:=Round(Perf.TotalTime, 2)), , False, "green", True
    blnSuccess = True

CleanUp:

    ' Close the cached connections, if any
    CloseCachedConnections

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

    ' Wrap up build.
    DoCmd.Hourglass False
    If Forms.Count > 0 Then
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

    ' Wait to finish the build till after we have saved the index.
    Operation.Finish

    ' Show MessageBox if not using GUI for build.
    If Forms.Count = 0 And blnSuccess Then
        ' Show message box when build is complete.
        MsgBox2 T("Build Complete for '{0}'", var0:=CurrentProject.Name), _
            T("Note that some settings may not take effect until this database is reopened."), _
            T("A backup of the previous build was saved as '{0}'.", var0:=FSO.GetFileName(strBackup)), vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadSingleObject
' Author    : Adam Waller
' Date      : 2/23/2023
' Purpose   : Reload a single object from source files.
'           : NOTE: Be very careful to release all references to the object you
'           : are attempting to import.
'---------------------------------------------------------------------------------------
'
Public Sub LoadSingleObject(cComponentClass As IDbComponent, strName As String, strSourceFilePath As String)

    Dim dCategories As Dictionary
    Dim dCategory As Dictionary
    Dim dSourceFiles As Dictionary

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

    ' Reload the project options and reset the logs
    Set VCSIndex = Nothing
    Set Options = Nothing
    Options.LoadProjectOptions
    Log.Clear
    Log.SourcePath = Options.GetExportFolder
    Log.Active = True
    Perf.StartTiming

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
        .Add Now
        .Spacer
        .Add T("Importing {0}...", var0:=strName)
        .Flush
    End With

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
            ' Show the conflicts resolution dialog
            .ShowDialog
            If .ApproveResolutions Then
                Log.Add T("Resolving source conflicts"), False
                .Resolve
            Else
                ' Cancel export
                Log.Spacer
                Log.Add T("Import Canceled"), , , "Red", True
                Operation.ErrorLevel = eelCritical
                GoTo CleanUp
            End If
        End If
    End With

    ' Check to see if we still have an item to import.
    If dCategories.Count = 0 Then
        Log.Add T("Skipped after conflict resolution."), , , "blue", True
    Else
        ' TODO: Maybe copy the existing object to the recycle bin, just in case
        ' the user makes a mistake. (Similar to how GitHub Desktop works)

        ' Replace the existing object with the source file
        cComponentClass.Merge strSourceFilePath
    End If

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

    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure all database objects are currently closed (This is really important,
    ' since we will be deleting most objects before importing them from source.)
    CloseDatabaseObjects

    ' Reload the project options and reset the logs
    Set VCSIndex = Nothing
    Set Options = Nothing
    Options.LoadProjectOptions
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
        .Add Now
        .Spacer
        .Add T("Scanning source files...")
        .Flush
    End With

    ' Check VBE Project protection
    If CurrentVBProject.Protection = vbext_pp_locked Then
        MsgBox2 T("Project Locked"), _
            T("Project is protected with a password."), _
            T("Please unlock the project before using this tool."), vbExclamation
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
            Log.ProgMax = dFiles.Count
            Perf.CategoryStart cCategory.Category

            ' Loop through each file in this category.
            For Each varFile In dFiles.Keys
                ' Import/merge the file
                Log.Increment
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                Operation.Pulse
                cCategory.Merge CStr(varFile)
                CatchAny eelError, T("Merge error in: {0}", var0:=varFile), ModuleName & ".MergeAllSource", True, True

                ' Bail out if we hit a critical error.
                If Operation.ErrorLevel = eelCritical Then Log.Add vbNullString: GoTo CleanUp
            Next varFile

            ' Show category wrap-up.
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


'---------------------------------------------------------------------------------------
' Procedure : CheckForLegacyModules
' Author    : Adam Waller
' Date      : 7/16/2020
' Purpose   : Informs the user if the database contains a legacy module from another
'           : fork of this project. (Some users might not realize that these are not
'           : needed anymore.)
'---------------------------------------------------------------------------------------
'
Private Sub CheckForLegacyModules()

    ' Check for legacy file
    If Options.ShowVCSLegacy Then
        If FSO.FileExists(Options.GetExportFolder & FSO.BuildPath("modules", "VCS_ImportExport.bas")) Then
            MsgBox2 T("Legacy Files not Needed"), _
                T("Other forks of the MSAccessVCS project used additional VBA modules to export code.") & vbNewLine & _
                T("This is no longer needed when using the installed Version Control Add-in.") & vbNewLine & vbNewLine & _
                T("Feel free to remove the legacy VCS_* modules from your database project and enjoy" & vbNewLine & _
                "a simpler, cleaner code base for ongoing development.  :-)"), _
                T("NOTE: This message can be disabled in 'Options -> Show Legacy Prompt'."), _
                vbInformation, T("Just a Suggestion...")
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpgradeSourceFiles
' Author    : Adam Waller
' Date      : 8/7/2024
' Purpose   : Removes any legacy file formats used by earlier versions of this add-in.
'---------------------------------------------------------------------------------------
'
Private Sub UpgradeSourceFiles()

    Dim strBase As String

    strBase = Options.GetExportFolder

    ' Remove legacy files by extension
    ClearFilesByExtension strBase & "sqltables", "tdf"
    ClearFilesByExtension strBase & "relations", "txt"      ' Relationships (pre-json)
    ClearFilesByExtension strBase & "report", "pv"          ' Print vars text file (pre-json)
    ClearFilesByExtension strBase & "tbldefs", "LNKD"       ' Formerly used for linked tables
    ClearFilesByExtension strBase & "tbldefs", "bas"        ' Moved to XML format
    ClearFilesByExtension strBase & "tbldefs", "tdf"

    ' Clear any print settings files if not using this option
    If Not Options.SavePrintVars Then
        ClearFilesByExtension "forms", "json"
        ClearFilesByExtension "reports", "json"
    End If

End Sub


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
Public Sub InitializeForms(cContainers As Dictionary)

    Dim cont As IDbComponent
    Dim frm As IDbComponent
    Dim dForms As Dictionary
    Dim cAllForms As IDbComponent
    Dim varKey As Variant

    ' Trap any errors that may occur when opening forms
    LogUnhandledErrors
    On Error Resume Next

    ' See if we imported any forms
    For Each cont In cContainers
        If cont.ComponentType = edbForm Then

            ' Loop through the forms in the current database
            Set cAllForms = New clsDbForm
            Set dForms = cAllForms.GetAllFromDB
            Log.ProgMax = dForms.Count
            For Each varKey In dForms.Keys

                ' See if this form matches one of the files we just imported
                Set frm = dForms(varKey)
                If cContainers(cont)("Files").Exists(frm.SourceFile) Then

                    ' Don't attempt to initialize add-in main form
                    ' (Likely not needed, and would require staging)
                    If frm.Name <> "frmVCSMain" Then

                        ' Open each form in design view
                        Perf.OperationStart "Initialize Forms"
                        DoCmd.OpenForm frm.Name, acDesign, , , , acHidden
                        DoEvents
                        DoCmd.Close acForm, frm.Name, acSaveNo
                        Perf.OperationEnd
                    End If
                    Log.Increment

                    ' Log any errors
                    CatchAny eelError, T("Error while initializing form {0}", var0:=frm.Name), ModuleName & ".InitializeForms"

                    ' Update the index, since the save date has changed, but reuse the code hash
                    ' since we just calculated it after importing the form.
                    With VCSIndex.Item(frm)
                        VCSIndex.Update frm, eatImport, .FileHash, .OtherHash
                    End With

                End If
            Next varKey
        End If
    Next cont

    ' Check for any unhandled errors
    CatchAny eelError, "Unhandled error while initializing forms", ModuleName & ".InitializeForms"

End Sub
