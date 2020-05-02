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

    Dim colContainers As Collection
    Dim cCategory As IDbComponent
    Dim cDbObject As IDbComponent
    Dim sngStart As Single
    Dim blnFullExport As Boolean
    
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
    
    
    ' Build containers of object types
    Set colContainers = New Collection
    With colContainers
        ' Shared objects in both MDB and ADP formats
        .Add New clsDbForm
        .Add New clsDbMacro
        .Add New clsDbModule
        .Add New clsDbReport
        .Add New clsDbTableData
        .Add New clsDbProjProperty
        .Add New clsDbVbeReference
        .Add New clsDbSavedSpec
        .Add New clsDbVbeProject
        If CurrentProject.ProjectType = acADP Then
            ' Some types of objects only exist in ADP projects
            .Add New clsAdpFunction
            .Add New clsAdpServerView
            .Add New clsAdpProcedure
            .Add New clsAdpTable
            .Add New clsAdpTrigger
        Else
            ' These objects only exist in DAO databases
            .Add New clsDbTableDef
            .Add New clsDbTableDataMacro
            .Add New clsDbQuery
            .Add New clsDbRelation
            .Add New clsDbProperty
        End If
    End With
    
    ' Loop through all categories
    For Each cCategory In colContainers
        
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