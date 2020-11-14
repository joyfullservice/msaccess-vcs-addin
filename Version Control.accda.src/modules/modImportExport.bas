Option Compare Database
Option Explicit
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ExportSource
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Export source files from the currently open database.
'---------------------------------------------------------------------------------------
'
Public Sub ExportSource()

    Dim cCategory As IDbComponent
    Dim cDbObject As IDbComponent
    Dim sngStart As Single
    Dim blnFullExport As Boolean

    ' Can't export without an open database
    If CurrentDb Is Nothing And CurrentProject.Connection Is Nothing Then Exit Sub
    
    ' Close any open forms or reports unless we are running from the add-in file.
    If CurrentProject.FullName <> CodeProject.FullName Then
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
        If .UseFastSave Then Log.Add "Using Fast Save"
        Log.Add Now
        Log.Spacer
        Log.Flush
    End With
    
    ' Loop through all categories
    For Each cCategory In GetAllContainers
        
        ' Clear any source files for nonexistant objects
        cCategory.ClearOrphanedSourceFiles
            
        ' Only show category details when it contains objects
        If cCategory.Count = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add "No " & cCategory.Category & " found in this database.", Options.ShowDebug
        Else
            ' Show category header and clear out any orphaned files.
            Log.Spacer Options.ShowDebug
            Log.PadRight "Exporting " & cCategory.Category & "...", , Options.ShowDebug
            Log.ProgMax = cCategory.Count
            Perf.ComponentStart cCategory.Category

            ' Loop through each object in this category.
            For Each cDbObject In cCategory.GetAllFromDB
                
                ' Check for fast save option
                If Options.UseFastSave And Not blnFullExport Then
                    If HasMoreRecentChanges(cDbObject) Then
                        Log.Increment
                        Log.Add "  " & cDbObject.Name, Options.ShowDebug
                        cDbObject.Export
                    Else
                        Log.Add "  (Skipping '" & cDbObject.Name & "')", Options.ShowDebug
                    End If
                Else
                    ' Always export object
                    Log.Increment
                    Log.Add "  " & cDbObject.Name, Options.ShowDebug
                    cDbObject.Export
                End If
                    
                ' Some kinds of objects are combined into a single export file, such
                ' as database properties. For these, we just need to run the export once.
                If cCategory.SingleFile Then Exit For
                
            Next cDbObject
            
            ' Show category wrap-up.
            Log.Add "[" & cCategory.Count & "]" & IIf(Options.ShowDebug, " " & cCategory.Category & " processed.", vbNullString)
            'Log.Flush  ' Gives smoother output, but slows down export.
            Perf.ComponentEnd cCategory.Count
        End If
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
    Log.SaveFile FSO.BuildPath(Options.GetExportFolder, "Export.log")
    
    ' Check for VCS_ImportExport.bas (Used with other forks)
    CheckForLegacyModules
    
    ' Restore original fast save option, and save options with project
    Options.SaveOptionsForProject
    
    ' Clear reference to FileSystemObject
    Set FSO = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Build
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Build the project from source files.
'---------------------------------------------------------------------------------------
'
Public Sub Build(strSourceFolder As String)

    Dim strPath As String
    Dim strText As String
    Dim cCategory As IDbComponent
    Dim sngStart As Single
    Dim colFiles As Collection
    Dim varFile As Variant
    
    ' Close the current database if it is currently open.
    If Not (CurrentDb Is Nothing And CurrentProject.Connection Is Nothing) Then
        ' Need to close the current database before we can replace it.
        RunBuildAfterClose strSourceFolder
        Exit Sub
    End If
    
    ' Make sure we can find the source files
    If Not FolderHasVcsOptionsFile(strSourceFolder) Then
        MsgBox2 "Source files not found", "Required source files were not found in the following folder:", strSourceFolder, vbExclamation
        Exit Sub
    End If
    
    ' Now reset the options and logs
    Set Options = Nothing
    Options.LoadOptionsFromFile strSourceFolder & "vcs-options.json"
    Log.Clear

    ' If we are using encryption, make sure we are able to decrypt the values
    If Options.Security = esEncrypt And Not VerifyHash(strSourceFolder & "vcs-options.json") Then
        MsgBox2 "Encryption Key Mismatch", "The required encryption key is either missing or incorrect.", _
            "Please update the encryption key before building this project from source.", vbExclamation
        Exit Sub
    End If
    
    ' Build original file name for database
    strPath = GetOriginalDbFullPathFromSource(strSourceFolder)
    If strPath = vbNullString Then
        MsgBox2 "Unable to determine database file name", "Required source files were not found or could not be decrypted:", strSourceFolder, vbExclamation
        Exit Sub
    End If
    
    ' Start performance timers
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
        .Add "Beginning Build from Source", False
        .Add FSO.GetFileName(strPath)
        .Add "VCS Version " & GetVCSVersion
        .Add Now
        .Spacer
        .Flush
    End With
    
    ' Rename original file as a backup
    strText = GetBackupFileName(strPath)
    If FSO.FileExists(strPath) Then Name strPath As strText
    Log.Add "Saving backup of original database..."
    Log.Add "Saved as " & FSO.GetFileName(strText) & "."
    
    ' Create a new database with the original name
    If LCase$(FSO.GetExtensionName(strPath)) = "adp" Then
        ' ADP project
        Application.NewAccessProject strPath
    Else
        ' Regular Access database
        Application.NewCurrentDatabase strPath
    End If
    Log.Add "Created blank database for import."
    Log.Spacer
    
    ' Remove any non-built-in references before importing from source.
    Log.Add "Removing non built-in references...", False
    RemoveNonBuiltInReferences

    ' Loop through all categories
    For Each cCategory In GetAllContainers
        
        ' Get collection of source files
        Set colFiles = cCategory.GetFileList
        
        ' Only show category details when source files are found
        If colFiles.Count = 0 Then
            Log.Spacer Options.ShowDebug
            Log.Add "No " & cCategory.Category & " source files found.", Options.ShowDebug
        Else
            ' Show category header
            Log.Spacer Options.ShowDebug
            Log.PadRight "Importing " & cCategory.Category & "...", , Options.ShowDebug
            Log.ProgMax = colFiles.Count
            Perf.ComponentStart cCategory.Category

            ' Loop through each file in this category.
            For Each varFile In colFiles
                ' Import the file
                Log.Increment
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                cCategory.Import CStr(varFile)
            Next varFile
            
            ' Show category wrap-up.
            Log.Add "[" & colFiles.Count & "]" & IIf(Options.ShowDebug, " " & cCategory.Category & " processed.", vbNullString)
            'Log.Flush  ' Gives smoother output, but slows down the import.
            Perf.ComponentEnd colFiles.Count
        End If
    Next cCategory

    ' Run any post-build instructions
    If Options.RunAfterBuild <> vbNullString Then
        Log.Add "Running " & Options.RunAfterBuild & "..."
        Perf.OperationStart "RunAfterBuild"
        RunSubInCurrentProject Options.RunAfterBuild
        Perf.OperationEnd
    End If

    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    
    ' Add performance data to log file
    Perf.EndTiming
    Log.Add vbCrLf & Perf.GetReports, False
    
    ' Write log file to disk
    Log.SaveFile FSO.BuildPath(Options.GetExportFolder, "Import.log")

    DoCmd.Hourglass False
    If Forms.Count > 0 Then
        ' Finish up on GUI
        Form_frmVCSMain.FinishBuild
    Else
        ' Show message box when build is complete.
        MsgBox2 "Build Complete", "Some settings will not take effect until the database is restarted.", , vbInformation
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllContainers
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return a collection of all containers.
'           : NOTE: The order doesn't matter for export, but is VERY important
'           : when building the project from source.
'---------------------------------------------------------------------------------------
'
Private Function GetAllContainers() As Collection
    
    Dim blnADP As Boolean
    Dim blnMDB As Boolean
    
    blnADP = (CurrentProject.ProjectType = acADP)
    blnMDB = (CurrentProject.ProjectType = acMDB)
    
    Set GetAllContainers = New Collection
    With GetAllContainers
        ' Shared objects in both MDB and ADP formats
        If blnMDB Then .Add New clsDbTheme
        .Add New clsDbVbeProject
        .Add New clsDbVbeReference
        .Add New clsDbVbeForm
        .Add New clsDbProjProperty
        .Add New clsDbSavedSpec
        If blnADP Then
            ' Some types of objects only exist in ADP projects
            .Add New clsAdpFunction
            .Add New clsAdpServerView
            .Add New clsAdpProcedure
            .Add New clsAdpTable
            .Add New clsAdpTrigger
        ElseIf blnMDB Then
            ' These objects only exist in DAO databases
            .Add New clsDbSharedImage
            .Add New clsDbImexSpec
            .Add New clsDbProperty
            .Add New clsDbTableDef
            .Add New clsDbQuery
        End If
        ' Additional objects to import after ADP/MDB specific items
        .Add New clsDbForm
        .Add New clsDbMacro
        .Add New clsDbModule
        .Add New clsDbReport
        .Add New clsDbTableData
        If blnMDB Then
            .Add New clsDbTableDataMacro
            .Add New clsDbRelation
            .Add New clsDbDocument
            .Add New clsDbNavPaneGroup
            .Add New clsDbHiddenAttribute
        End If
    End With
    
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
' Procedure : RemoveNonBuiltInReferences
' Author    : Adam Waller
' Date      : 10/20/2020
' Purpose   : Remove any references that are not built-in. (Sometimes additional
'           : references are added when creating a new database, not not really needed
'           : when building the project from source.)
'---------------------------------------------------------------------------------------
'
Private Sub RemoveNonBuiltInReferences()

    Dim intCnt As Integer
    Dim strName As String
    Dim ref As Access.Reference
    
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If Not ref.BuiltIn Then
            strName = ref.Name
            Application.References.Remove ref
            Log.Add "  Removed " & strName, False
        End If
        Set ref = Nothing
    Next intCnt
    
End Sub