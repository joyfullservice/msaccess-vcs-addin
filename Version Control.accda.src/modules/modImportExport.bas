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

    ' Run any custom sub before export
    If Options.RunBeforeExport <> vbNullString Then RunSubInCurrentProject Options.RunBeforeExport

    ' Save property with the version of Version Control we used for the export.
    If GetDBProperty("Last VCS Version") <> GetVCSVersion Then
        SetDBProperty "Last VCS Version", GetVCSVersion
        blnFullExport = True
    End If
    ' Set this as text to save display in current user's locale rather than Zulu time.
    SetDBProperty "Last VCS Export", Now, dbText ' dbDate

    sngStart = Timer
    Set colVerifiedPaths = New Collection   ' Reset cache
    VerifyPath Options.GetExportFolder

    ' Display heading
    With Options
        '.ShowDebug = True
        '.UseFastSave = False
        Log.Spacer
        Log.Add "Beginning Export of all Source", False
        Log.Add CurrentProject.Name
        Log.Add "VCS Version " & GetVCSVersion
        If .UseFastSave Then Log.Add "Using Fast Save"
        Log.Add Now()
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

            ' Loop through each object in this category.
            For Each cDbObject In cCategory.GetAllFromDB()
                
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
            
        End If
    Next cCategory

    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    Log.SaveFile Options.GetExportFolder & "\Export.log"
    
    ' Restore original fast save option, and save options with project
    Options.SaveOptionsForProject
    
    ' Clear reference to FileSystemObject
    Set FSO = Nothing
    
    ' Run any custom sub before export
    If Options.RunAfterExport <> vbNullString Then RunSubInCurrentProject Options.RunAfterExport

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
    
    ' Now reset the options and logs
    Set Options = Nothing
    Options.LoadOptionsFromFile strSourceFolder & "vcs-options.json"
    Log.Clear
    sngStart = Timer
    
    ' Make sure we can find the source files
    If Not FolderHasVcsOptionsFile(strSourceFolder) Then
        MsgBox2 "Source files not found", "Required source files were not found in the following folder:", strSourceFolder, vbExclamation
        Exit Sub
    End If
    
    ' Build original file name for database
    strPath = GetOriginalDbFullPathFromSource(strSourceFolder)
    If strPath = vbNullString Then
        MsgBox2 "Unable to determine database file name", "Required source files were not found or could not be decrypted:", strSourceFolder, vbExclamation
        Exit Sub
    End If
    
    ' Launch the GUI and display the build header.
    DoCmd.Hourglass True
    Form_frmVCSMain.StartBuild
    With Log
        .Spacer
        .Add "Beginning Build from Source", False
        .Add FSO.GetFileName(strPath)
        .Add "VCS Version " & GetVCSVersion
        .Add Now()
        .Spacer
        .Flush
    End With
    
    ' Rename original file as a backup
    strText = GetBackupFileName(strPath)
    Name strPath As strText
    Log.Add "Saving backup of original database..."
    Log.Add "Saved as " & FSO.GetFileName(strText) & "."
    
    ' Create a new database with the original name
    If LCase(FSO.GetExtensionName(strPath)) = "adp" Then
        ' ADP project
        Application.NewAccessProject strPath
    Else
        ' Regular Access database
        Application.NewCurrentDatabase strPath
    End If
    Log.Add "Created blank database for import."
    Log.Spacer
    

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

            ' Loop through each file in this category.
            For Each varFile In colFiles
                ' Import the file
                Log.Increment
                Log.Add "  " & FSO.GetFileName(varFile), Options.ShowDebug
                cCategory.Import CStr(varFile)
            Next varFile
            
            ' Show category wrap-up.
            Log.Add "[" & cCategory.Count & "]" & IIf(Options.ShowDebug, " " & cCategory.Category & " processed.", vbNullString)
            'Log.Flush  ' Gives smoother output, but slows down the import.
        End If
    Next cCategory

    ' Show final output and save log
    Log.Spacer
    Log.Add "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    Log.SaveFile Options.GetExportFolder & "\Import.log"

    ' Finish up on GUI
    Form_frmVCSMain.FinishBuild
    DoCmd.Hourglass False
    
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
        .Add New clsDbVbeProject
        .Add New clsDbVbeReference
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
            .Add New clsDbIMEXSpec
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
Private Function GetBackupFileName(strPath As String)
    
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