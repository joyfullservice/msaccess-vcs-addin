'---------------------------------------------------------------------------------------
' Module    : modImportExport
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Main export/import/merge functions for add-in.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
Private Const moduleName As String = "modImportExport:"
 
'---------------------------------------------------------------------------------------
' Procedure : ExportSource
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Export source files from the currently open database.
'---------------------------------------------------------------------------------------
'
Public Sub ExportSource(blnFullExport As Boolean)
    On Error Resume Next
    Dim FunctionName as String
    FunctionName = moduleName & "ExportSource:"
    
    Dim cCategory As IDbComponent
    Dim cDbObject As IDbComponent
    Dim sngStart As Single
    Dim lngCount As Long

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
    Set Options = Nothing
    Options.LoadProjectOptions
    Log.Clear
    Set VCSIndex = Nothing
    VCSIndex.LoadFromFile
    Perf.StartTiming

    ' Run any custom sub before export
    If Options.RunBeforeExport <> vbNullString Then
        Log.Add "Running " & Options.RunBeforeExport & "..."
        Perf.OperationStart "RunBeforeExport"
        RunSubInCurrentProject Options.RunBeforeExport
        Perf.OperationEnd
    End If

    ' Save property with the version of Version Control we used for the export.
    If GetDBProperty("Last VCS Version") <> GetVCSVersion Then
        SetDBProperty "Last VCS Version", GetVCSVersion
        blnFullExport = True
    End If
    ' Set this as text to save display in current user's locale rather than Zulu time.
    SetDBProperty "Last VCS Export", Now, dbText ' dbDate

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
        If Not blnFullExport Then Log.Add "Using Fast Save"
        Log.Add Now
        Log.Spacer
        Log.Flush
    End With
    
    ' Loop through all categories
    For Each cCategory In GetAllContainers
        
        ' Clear any source files for nonexistant objects
        cCategory.ClearOrphanedSourceFiles
            
        ' Only show category details when it contains objects
        lngCount = cCategory.Count(Not blnFullExport)
        If lngCount = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add "No " & LCase(cCategory.Category) & " found in this database.", Options.ShowDebug
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
                CatchAny eelError, Err.Number & ":" & Err.Description, FunctionName & ":Exporting:" & _
                    Options.UseFastSave & LCase(cCategory.Category) & " " & cDbObject.Name, True, True
                On Error Goto 0
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
    End If
    
    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    
    ' Add performance data to log file
    Perf.EndTiming
    Log.Add vbCrLf & Perf.GetReports, False
    
    ' Save log file to disk
    Log.SaveFile FSO.BuildPath(Options.GetExportFolder, "Export-" & Format(Now, "YYYY-MM-DD-hh.mm.ss") & ".log")
    
    ' Check for VCS_ImportExport.bas (Used with other forks)
    CheckForLegacyModules
    
    ' Restore original fast save option, and save options with project
    Options.SaveOptionsForProject
    
    ' Save index file
    VCSIndex.ExportDate = Now
    If blnFullExport Then VCSIndex.FullExportDate = Now
    VCSIndex.Save

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

    On Error Resume Next
    Dim FunctionName As String
    FunctionName = moduleName & "ExportSource:"

    Dim strPath As String
    Dim strBackup As String
    Dim cCategory As IDbComponent
    Dim sngStart As Single
    Dim colFiles As Collection
    Dim varFile As Variant
    Dim strType As String
    
    Dim strText As String   ' Remove later
    
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
    
    ' If we are using encryption, make sure we are able to decrypt the values.
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
    
    ' If we are using encryption, make sure we are able to decrypt the values.
    ' NOTE: There is no CurrentProject at this point, so we will have limited
    ' functionality with the options class.
    Set Options = Nothing
    Options.LoadOptionsFromFile strSourceFolder & "vcs-options.json"
    If Options.Security = esEncrypt And Not VerifyHash(strSourceFolder & "vcs-options.json") Then
        MsgBox2 "Encryption Key Mismatch", "The required encryption key is either missing or incorrect.", _
            "Please update the encryption key before building this project from source.", vbExclamation
        Exit Sub
    End If
    
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
    If FSO.FileExists(strPath) Then Name strPath As strBackup
    Log.Add "Saving backup of original database..."
    Log.Add "Saved as " & FSO.GetFileName(strBackup) & "."
    
    ' Create a new database with the original name
    If blnFullBuild Then
        If LCase$(FSO.GetExtensionName(strPath)) = "adp" Then
            ' ADP project
            Application.NewAccessProject strPath
        Else
            ' Regular Access database
            Application.NewCurrentDatabase strPath
        End If
        Log.Add "Created blank database for import."
    End If
    
    ' Now that we have a new database file, we can load the index.
    Set VCSIndex = Nothing
    VCSIndex.LoadFromFile
    
    ' Remove any non-built-in references before importing from source.
    Log.Spacer
    If blnFullBuild Then
        Log.Add "Removing non built-in references...", False
        RemoveNonBuiltInReferences
    End If

    ' Loop through all categories
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
                ' Import the file
                Log.Increment
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                If blnFullBuild Then
                    cCategory.Import CStr(varFile)
                    CatchAny eelError, Err.Number & ":" & Err.Description, FunctionName & ":Importing:" & _
                        Options.UseFastSave & LCase(cCategory.Category) & " " & FSO.GetFileName(varFile), True, True
                Else
                    cCategory.Merge CStr(varFile)
                    CatchAny eelError, Err.Number & ":" & Err.Description, FunctionName & ":Merging:" & _
                        Options.UseFastSave & LCase(cCategory.Category) & " " & FSO.GetFileName(varFile), True, True
                End If
            Next varFile
            
            ' Show category wrap-up.
            Log.Add "[" & colFiles.Count & "]" & IIf(Options.ShowDebug, " " & LCase(cCategory.Category) & " processed.", vbNullString)
            'Log.Flush  ' Gives smoother output, but slows down the import.
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
        'If Options.runaftermerge <> vbNullString Then
            Log.Add "Running " & Options.RunAfterBuild & "..."
            Perf.OperationStart "RunAfterBuild"
            RunSubInCurrentProject Options.RunAfterBuild
            Perf.OperationEnd
        'End If
    End If

    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    
    ' Add performance data to log file
    Perf.EndTiming
    Log.Add vbCrLf & Perf.GetReports, False
    
    ' Write log file to disk
    Log.SaveFile FSO.BuildPath(Options.GetExportFolder, IIf(blnFullBuild, "Import", "Merge") & Format(Now, "-YYYY-MM-DD-hh.mm.ss") & ".log")

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
    
    strFolder = FSO.GetParentFolderName(strPath) & "\"
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
        strFolder = Options.GetExportFolder & "themes\"
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
        VerifyHash = CanDecrypt(strHash)
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
        If FSO.FileExists(Options.GetExportFolder & "modules\VCS_ImportExport.bas") Then
            MsgBox2 "Legacy Files not Needed", _
                "Other forks of the MSAccessVCS project used additional VBA modules to export code." & vbCrLf & _
                "This is no longer needed when using the installed Version Control Add-in." & vbCrLf & vbCrLf & _
                "Feel free to remove the legacy VCS_* modules from your database project and enjoy" & vbCrLf & _
                "a simpler, cleaner code base for ongoing development.  :-)", _
                "NOTE: This message can be disabled in 'Options -> Show Legacy Prompt'.", vbInformation, "Just a Suggestion..."
        End If
    End If
    
End Sub