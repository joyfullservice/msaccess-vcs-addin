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
Public Sub ExportSource(blnFullExport As Boolean)

    Dim cCategory As IDbComponent
    Dim cDbObject As IDbComponent
    Dim sngStart As Single
    Dim lngCount As Long
    
    ' Use inline error handling functions to trap and log errors.
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Can't export without an open database
    If CurrentDb Is Nothing And CurrentProject.Connection Is Nothing Then Exit Sub
    
    ' If we are running this from the current database, we need to run it a different
    ' way to prevent file corruption issues.
    If StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0 Then
        RunExportForCurrentDB
        Exit Sub
    Else
        ' Close any open forms or reports.
        If Not CloseAllFormsReports Then
            MsgBox2 "Please close forms and reports", _
                "All forms and reports must be closed to export source code.", _
                , vbExclamation
            Exit Sub
        End If
    End If
    
    ' Reload the project options and reset the logs
    Set VCSIndex = Nothing
    Set Options = Nothing
    Options.LoadProjectOptions
    Log.Clear
    Log.Active = True
    Perf.StartTiming

    ' If options (or VCS version) have changed, a full export will be required
    If (VCSIndex.OptionsHash <> Options.GetHash) Then blnFullExport = True

    ' Begin timer at start of export.
    sngStart = Timer

    ' Display heading
    With Options
        '.ShowDebug = True
        '.UseFastSave = False
        Log.Spacer
        Log.Add "Beginning Export of all Source", False
        Log.Add CurrentProject.Name
        Log.Add "VCS Version " & GetVCSVersion
        Log.Add IIf(blnFullExport, "Performing Full Export", "Using Fast Save")
        Log.Add Now
    End With
    
    ' Run any custom sub before export
    If Options.RunBeforeExport <> vbNullString Then
        Log.Add "Running " & Options.RunBeforeExport & "..."
        Log.Flush
        Perf.OperationStart "RunBeforeExport"
        RunSubInCurrentProject Options.RunBeforeExport
        Perf.OperationEnd
    End If

    ' Finish header section
    Log.Spacer
    Log.Flush
    
    ' Loop through all categories
    For Each cCategory In GetAllContainers
        
        ' Clear any source files for nonexistant objects
        cCategory.ClearOrphanedSourceFiles
            
        ' Only show category details when it contains objects
        lngCount = cCategory.Count(Not blnFullExport)
        If lngCount = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add IIf(blnFullExport, "No ", "No modified ") & LCase(cCategory.Category) & " found in this database.", Options.ShowDebug
        Else
            ' Show category header and clear out any orphaned files.
            Log.Spacer Options.ShowDebug
            Log.PadRight "Exporting " & LCase(cCategory.Category) & "...", , Options.ShowDebug
            Log.ProgMax = lngCount
            Perf.ComponentStart cCategory.Category

            ' Loop through each object in this category.
            For Each cDbObject In cCategory.GetAllFromDB(Not blnFullExport)
                
                ' Export object
                Log.Increment
                Log.Add "  " & cDbObject.Name, Options.ShowDebug
                cDbObject.Export
                CatchAny eelError, "Error exporting " & cDbObject.Name, ModuleName & ".ExportSource", True, True
                    
                ' Some kinds of objects are combined into a single export file, such
                ' as database properties. For these, we just need to run the export once.
                If cCategory.SingleFile Then Exit For
                
            Next cDbObject
            
            ' Show category wrap-up.
            Log.Add "[" & lngCount & "]" & IIf(Options.ShowDebug, " " & LCase(cCategory.Category) & " processed.", vbNullString)
            'Log.Flush  ' Gives smoother output, but slows down export.
            Perf.ComponentEnd lngCount
        End If
        
        ' Bail out if we hit a critical error.
        If Log.ErrorLevel = eelCritical Then GoTo CleanUp
        
    Next cCategory
    
    ' Run any cleanup routines
    RemoveThemeZipFiles
    
    ' Run any custom sub after export
    If Options.RunAfterExport <> vbNullString Then
        Log.Add "Running " & Options.RunAfterExport & "..."
        Perf.OperationStart "RunAfterExport"
        RunSubInCurrentProject Options.RunAfterExport
        Perf.OperationEnd
        CatchAny eelError, "Error running " & Options.RunAfterExport, ModuleName & ".ExportSource", True, True
    End If
    
    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)", , False, "green", True
    
    ' Add performance data to log file and save file
    Perf.EndTiming
    With Log
        .Add vbCrLf & Perf.GetReports, False
        .SaveFile FSO.BuildPath(Options.GetExportFolder, "Export.log")
        .Active = False
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
    
CleanUp:

    ' Clear references to FileSystemObject and other objects
    Set FSO = Nothing
    Set VCSIndex = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Build (Full build or Merge Build)
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Build the project from source files.
'---------------------------------------------------------------------------------------
'
Public Sub Build(strSourceFolder As String, blnFullBuild As Boolean)

    Dim strPath As String
    Dim strBackup As String
    Dim cCategory As IDbComponent
    Dim sngStart As Single
    Dim colFiles As Collection
    Dim varFile As Variant
    Dim strType As String
    
    Dim strText As String   ' Remove later
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' The type of build will be used in various messages and log entries.
    strType = IIf(blnFullBuild, "Build", "Merge")
    
    ' For full builds, close the current database if it is currently open.
    If blnFullBuild Then
        If Not (CurrentDb Is Nothing And CurrentProject.Connection Is Nothing) Then
            ' Need to close the current database before we can replace it.
            RunBuildAfterClose strSourceFolder
            Exit Sub
        End If
    End If
    
    ' Make sure we can find the source files
    If Not FolderHasVcsOptionsFile(strSourceFolder) Then
        MsgBox2 "Source files not found", "Required source files were not found in the following folder:", strSourceFolder, vbExclamation
        Exit Sub
    End If
    
    ' Verify that the source files are being merged into the correct database.
    If Not blnFullBuild Then
        strPath = GetOriginalDbFullPathFromSource(strSourceFolder)
        If strPath = vbNullString Then
            MsgBox2 "Unable to determine database file name", "Required source files were not found or could not be decrypted:", strSourceFolder, vbExclamation
            Exit Sub
        ElseIf StrComp(strPath, CurrentProject.FullName, vbTextCompare) <> 0 Then
            MsgBox2 "Cannot merge to a different database", _
                "The database file name for the source files must match the currently open database.", _
                "Current: " & CurrentProject.FullName & vbCrLf & _
                "Source: " & strPath, vbExclamation
            Exit Sub
        End If
    End If
    
    Set Options = Nothing
    Options.LoadOptionsFromFile StripSlash(strSourceFolder) & PathSep & "vcs-options.json"
    
    ' Build original file name for database
    If blnFullBuild Then
        strPath = GetOriginalDbFullPathFromSource(strSourceFolder)
        If strPath = vbNullString Then
            MsgBox2 "Unable to determine database file name", "Required source files were not found or could not be decrypted:", strSourceFolder, vbExclamation
            Exit Sub
        End If
    Else
        ' Run any pre-merge instructions
        'strText = dNZ(Options.GitSettings, "RunBeforeMerge")
        If strText <> vbNullString Then
            Log.Add "Running " & strText & "..."
            Perf.OperationStart "RunBeforeMerge"
            RunSubInCurrentProject strText
            Perf.OperationEnd
        End If
    End If
    
    ' Start log and performance timers
    Log.Clear
    Log.Active = True
    sngStart = Timer
    Perf.StartTiming
    
    ' Check if we are building the add-in file
    If FSO.GetFileName(strPath) = CodeProject.Name Then
        ' When building this add-in file, we should output to the debug
        ' window instead of the GUI form. (Since we are importing
        ' a form with the same name as the GUI form.)
        ShowIDE
    Else
        ' Launch the GUI form
        Form_frmVCSMain.StartBuild
    End If

    ' Display the build header.
    DoCmd.Hourglass True
    With Log
        .Spacer
        .Add "Beginning " & strType & " from Source", False
        .Add FSO.GetFileName(strPath)
        .Add "VCS Version " & GetVCSVersion
        .Add Now
        .Spacer
        .Flush
    End With
    
    ' Rename original file as a backup
    strBackup = GetBackupFileName(strPath)
    If FSO.FileExists(strPath) Then
        Log.Add "Saving backup of original database..."
        Name strPath As strBackup
        Log.Add "Saved as " & FSO.GetFileName(strBackup) & "."
    End If
    
    ' Create a new database with the original name
    If blnFullBuild Then
        Perf.OperationStart "Create new database"
        If LCase$(FSO.GetExtensionName(strPath)) = "adp" Then
            ' ADP project
            Application.NewAccessProject strPath
        Else
            ' Regular Access database
            Application.NewCurrentDatabase strPath
        End If
        Perf.OperationEnd
        Log.Add "Created blank database for import."
    End If
    
    ' Now that we have a new database file, we can load the index.
    Set VCSIndex = Nothing
    
    If blnFullBuild Then
    
        ' Remove any non-built-in references before importing from source.
        Log.Add "Removing non built-in references...", False
        RemoveNonBuiltInReferences
        
        ' Check for any RunBeforeBuild
        If Options.RunBeforeBuild <> vbNullString Then
            ' Run any pre-build bootstrapping code
            PrepareRunBootstrap
        End If
        
    End If
    
    ' Loop through all categories
    Log.Spacer
    For Each cCategory In GetAllContainers
        
        ' Get collection of source files
        If blnFullBuild Then
            ' Return all the source files
            Set colFiles = cCategory.GetFileList
        Else
            ' Return just the modified source files for merge
            ' (Optionally uses the git integration to determine changes.)
            Set colFiles = VCSIndex.GetModifiedSourceFiles(cCategory)
        End If
        
        ' Only show category details when source files are found
        If colFiles.Count = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add "No " & LCase(cCategory.Category) & " source files found.", Options.ShowDebug
        Else
            ' Show category header
            Log.Spacer Options.ShowDebug
            Log.PadRight IIf(blnFullBuild, "Importing ", "Merging ") & LCase(cCategory.Category) & "...", , Options.ShowDebug
            Log.ProgMax = colFiles.Count
            Perf.ComponentStart cCategory.Category

            ' Loop through each file in this category.
            For Each varFile In colFiles
                ' Import/merge the file
                Log.Increment
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                If blnFullBuild Then
                    cCategory.Import CStr(varFile)
                Else
                    cCategory.Merge CStr(varFile)
                End If
                CatchAny eelError, strType & " error in: " & varFile, ModuleName & ".Build", True, True
            Next varFile
            
            ' Show category wrap-up.
            Log.Add "[" & colFiles.Count & "]" & IIf(Options.ShowDebug, " " & LCase(cCategory.Category) & " processed.", vbNullString)
            Perf.ComponentEnd colFiles.Count
        End If
    Next cCategory

    ' Run any post-build/merge instructions
    If blnFullBuild Then
        If Options.RunAfterBuild <> vbNullString Then
            Log.Add "Running " & Options.RunAfterBuild & "..."
            Perf.OperationStart "RunAfterBuild"
            RunSubInCurrentProject Options.RunAfterBuild
            Perf.OperationEnd
        End If
    Else
        ' Merge build
        'If Options.RunAfterMerge <> vbNullString Then
        '    Log.Add "Running " & Options.RunAfterMerge & "..."
        '    Perf.OperationStart "RunAfterMerge"
        '    RunSubInCurrentProject Options.RunAfterMerge
        '    Perf.OperationEnd
        'End If
    End If
    
    ' Log any errors after build/merge
    CatchAny eelError, "Error running " & CallByName(Options, "RunAfter" & strType, VbGet), ModuleName & ".Build", True, True
    
    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)", , False, "green", True
    
    ' Add performance data to log file and save file.
    Perf.EndTiming
    With Log
        .Add vbCrLf & Perf.GetReports, False
        .SaveFile FSO.BuildPath(Options.GetExportFolder, strType & ".log")
        .Active = False
    End With
    
    ' Wrap up build.
    DoCmd.Hourglass False
    If Forms.Count > 0 Then
        ' Finish up on GUI
        Form_frmVCSMain.FinishBuild blnFullBuild
    Else
        ' Allow navigation pane to refresh list of objects.
        DoEvents
    End If
    
    ' Save index file (After build complete)
    If blnFullBuild Then
        ' NOTE: Add a couple seconds since some items may still be in the process of saving.
        VCSIndex.FullBuildDate = DateAdd("s", 2, Now)
    Else
        VCSIndex.MergeBuildDate = DateAdd("s", 2, Now)
    End If
    VCSIndex.Save
    Set VCSIndex = Nothing
        
    ' Show MessageBox if not using GUI for build.
    If Forms.Count = 0 Then
        ' Show message box when build is complete.
        MsgBox2 strType & " Complete for '" & CurrentProject.Name & "'", _
            "Note that some settings may not take effect until this database is reopened.", _
            "A backup of the previous build was saved as '" & FSO.GetFileName(strBackup) & "'.", vbInformation
    End If
    
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
    
    ' Attempt up to 100 versions of the file name. (i.e. Database_VSBackup45.accdb)
    For intCnt = 1 To 100
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
' Procedure : VerifyHash
' Author    : Adam Waller
' Date      : 7/29/2020
' Purpose   : Verify that we can decrypt the hash value in the options file, if found.
'           : Returns false if a hash is found but cannot be decrypted.
'---------------------------------------------------------------------------------------
'
Private Function VerifyHash(strOptionsFile As String) As Boolean
    
    Dim dFile As Dictionary
    Dim strHash As String
    
    Set dFile = ReadJsonFile(strOptionsFile)
    strHash = dNZ(dFile, "Info\Hash")
    
    ' Check hash value
    If strHash = vbNullString Then
        ' Could not find hash.
        VerifyHash = True
    Else
        ' Return true if we can successfully decrypt the hash.
        VerifyHash = False
    End If
    
End Function


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
            MsgBox2 "Legacy Files not Needed", _
                "Other forks of the MSAccessVCS project used additional VBA modules to export code." & vbCrLf & _
                "This is no longer needed when using the installed Version Control Add-in." & vbCrLf & vbCrLf & _
                "Feel free to remove the legacy VCS_* modules from your database project and enjoy" & vbCrLf & _
                "a simpler, cleaner code base for ongoing development.  :-)", _
                "NOTE: This message can be disabled in 'Options -> Show Legacy Prompt'.", vbInformation, "Just a Suggestion..."
        End If
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
    Dim cMod As clsDbModule
    
    ' Update output since there may be some delays
    Log.Add "Loading bootstrap..."
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
                    Log.Add "Importing bootstrap module '" & strName & "'", False
                    .Import CStr(varFile)
                    Exit For
                End If
            Next varFile
        End With
    End With
    
    ' Make sure we actually have a module before we attempt to run the code
    If CurrentProject.AllModules.Count = 0 Then
        ' Could not find source file
        Log.Error eelError, "Could not find source file for " & strModule, ModuleName & ".PrepareRunBootstrap"
    Else
        ' Important: We need to Run Project.Sub not Project.Module.Sub
        strName = Split(Options.RunBeforeBuild, ".")(1)
        
        ' Run any pre-build bootstrapping code
        Log.Add "Running " & Options.RunBeforeBuild
        Perf.OperationStart "RunBeforeBuild"
        RunSubInCurrentProject strName
        Perf.OperationEnd
    End If
    
    ' Now go back and remove all the non built-in references so they come
    ' back in the correct order, just in case a library was at a higher level.
    Log.Add "Removing non built-in references after running bootstrap", False
    RemoveNonBuiltInReferences
    
    Perf.OperationEnd   ' Bootstrap
    
End Sub