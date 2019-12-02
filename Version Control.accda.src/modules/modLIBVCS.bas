'@Folder("MSAccessVCS")
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : modLIBVCS
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : Extend some functionality like the possibility to choose export path and
'             give name parts list to select tables for what to save also data
'---------------------------------------------------------------------------------------
'

Private LastObjectsImportPath As String
'

Public Function pub_LIBVCS_LoadVCSModel(Optional commaSeparatedListOfNamePartsOfTablesToIncludeForDataSave As String = "tbl_L_;", _
                                        Optional commaSeparatedListOfNamePartsOfTablesToExcludeFromIncludedOnes As String = "tbl_L_NotThisOne;" _
                                        ) As Boolean

    Dim exportPath As String
    exportPath = getExportPath()
    If exportPath <> "" Then
        Debug.Print "VCS Export Path: " & exportPath
        
        Dim TablesToSaveData As Collection
        Set TablesToSaveData = getTablesToSaveData(commaSeparatedListOfNamePartsOfTablesToIncludeForDataSave, commaSeparatedListOfNamePartsOfTablesToExcludeFromIncludedOnes)
        
        Dim colParams As New Collection
        With colParams
            .Add Array("System", "GitHub")  ' IMPORTANT: Set this first, before other settings.
            .Add Array("Export Folder", exportPath)
            ' Optional parameters
            .Add Array("Show Debug", False)
            .Add Array("Include VBE", False)
            .Add Array("Fast Save", True)
            .Add Array("Save Print Vars", False)
            .Add Array("Save Query SQL", True)
            .Add Array("Save Table SQL", True)
            
            Dim varTableName As Variant
            For Each varTableName In TablesToSaveData
                .Add Array("Save Table", varTableName)
            Next
        End With
    
        ' Pass the parameters to the wrapper function
        LoadVersionControlMenu colParams
        pub_LIBVCS_LoadVCSModel = True
    End If
End Function

Public Function pub_LIBVCS_RemoveReferenceByName(LibName As String) As Boolean
    On Error GoTo err_RemoveByName
    Dim ref As Reference
    For Each ref In Application.References
        If ref.Name = LibName Then
            Application.References.Remove ref
            pub_LIBVCS_RemoveReferenceByName = True
            Exit For
        End If
    Next
ext_RemoveByName:
    Exit Function
err_RemoveByName:
    Select Case Err.Number
    Case Else
        MsgBox Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Passed parameter LibName: " & vbCrLf & LibName
        'Err.Raise Err.Number
    End Select
    Resume ext_RemoveByName:
End Function

Public Function pub_LIBVCS_ChangeExportPath()
    ResetExportFolderPath
    pub_LIBVCS_LoadVCSModel UserSettings_Get("VCSParams", "TablesToInclude", ""), UserSettings_Get("VCSParams", "TablesToExclude", "")
End Function

Public Function pub_LIBVCS_ImportAll()
    Dim ImportPath As String
    ImportPath = getImportPath
    If Dir(ImportPath) <> "" Then ImportAllSource True, getImportPath Else MsgBox "Import path not valid", VbMsgBoxStyle.vbCritical, "Import aborted"
End Function

Public Function pub_LIBVCS_ResetProjectAndImportAll()
    Dim ImportPath As String
    ImportPath = getImportPath
    If Dir(ImportPath) <> "" Then ImportProject True, ImportPath Else MsgBox "Import path not valid", VbMsgBoxStyle.vbCritical, "Reset/Import aborted"
End Function

Public Function pub_LIBVCS_ResetProject()
    If Not ResetProject(True) Then MsgBox "Project reset not successful", VbMsgBoxStyle.vbCritical, "Reset aborted"
End Function

'---------------------------------------------------------------------------------------
' Procedure : pub_LIBVCS_ImportObjects
' Author    : Marco Salernitano
' Date      : 14-Nov-2019
' Purpose   : Import multiple objects automatically recognizing their type (public call)
'---------------------------------------------------------------------------------------

Public Function pub_LIBVCS_ImportObjects(Optional simulate As Boolean) ' it's a function in order to be called by macros
    Dim ObjectImportPaths   As Collection
    Dim ObjectImportCounts  As Scripting.Dictionary
    Dim ObjectImportKey     As Variant
    Dim msg                 As String
    Set ObjectImportPaths = getObjectImportPaths
    If ObjectImportPaths.Count > 0 Then
        Set ObjectImportCounts = ImportObjects(ObjectImportPaths, simulate)
        For Each ObjectImportKey In ObjectImportCounts.Keys
            Debug.Print ObjectImportKey, ObjectImportCounts(ObjectImportKey)
            msg = msg & ObjectImportKey & ": " & ObjectImportCounts(ObjectImportKey) & vbCrLf
        Next
        MsgBox msg, , IIf(simulate, "(Simulation) ", "") & "Import results:"
    End If
End Function

Private Function getObjectImportPaths() As Collection
    Dim StartingFolder
    StartingFolder = LastObjectsImportPath
    If StartingFolder = "" Then StartingFolder = GetExportFolderPath
    ' If code reaches here, we don't have a copy of the path
    ' in the cached list of verified paths. Verify and add
    If StartingFolder = "" Or Dir(StartingFolder, vbDirectory) = "" Then StartingFolder = CurrentProject.Path
    Set getObjectImportPaths = SelectImportObjects(StartingFolder)
    If getObjectImportPaths.Count > 0 Then
        'MsgBox "Import path not valid", VbMsgBoxStyle.vbCritical, "Reset/Import aborted"
    'Else
        LastObjectsImportPath = getObjectImportPaths.Item(getObjectImportPaths.Count)
        LastObjectsImportPath = Left(LastObjectsImportPath, InStrRev(LastObjectsImportPath, "\"))
    End If
End Function

Private Function getExportPath() As String
    getExportPath = GetExportFolderPath
    ' If code reaches here, we don't have a copy of the path
    ' in the cached list of verified paths. Verify and add
    If getExportPath = "" Then
        getExportPath = CStr(SelectExportFolder(CurrentProject.Path))
        If getExportPath <> "" Then
            If Right(getExportPath, 1) <> "\" Then getExportPath = getExportPath & "\"
        Else
            Exit Function
        End If
    ElseIf Dir(getExportPath, vbDirectory) = "" Then
    ' Path does not seem to exist.
        Dim Answer
        Answer = MsgBox("Saved Path:" & vbCrLf & vbCrLf & getExportPath & vbCrLf & vbCrLf & "doesn't exist." & vbCrLf & vbCrLf & "Answer Yes to create it, No to choose another one or Cancel to abort.", vbYesNoCancel Or vbExclamation, "Export folder not found!")
        Select Case Answer
            Case vbYes
                ' Create it.
                MkDirIfNotExist getExportPath
            Case vbNo
                ' Ask for alternate path
                getExportPath = CStr(SelectExportFolder(CurrentProject.Path))
                If getExportPath <> "" Then
                    If Right(getExportPath, 1) <> "\" Then getExportPath = getExportPath & "\"
                Else
                    Exit Function
                End If
            Case Else
                Exit Function
        End Select
    Else
        Exit Function
    End If
    If Dir(getExportPath, vbDirectory) <> "" Then SetExportFolderPath getExportPath
End Function

Private Function getImportPath() As String
    Dim StartingFolder
    StartingFolder = GetExportFolderPath
    ' If code reaches here, we don't have a copy of the path
    ' in the cached list of verified paths. Verify and add
    If StartingFolder = "" Or Dir(StartingFolder, vbDirectory) = "" Then StartingFolder = CurrentProject.Path
    getImportPath = CStr(SelectImportFolder(StartingFolder))
    If getImportPath <> "" Then
        If Right(getImportPath, 1) <> "\" Then getImportPath = getImportPath & "\"
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetExportFolderPath, SetExportFolderPath, ResetExportFolderPath
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : Returns the saved or default export path
'---------------------------------------------------------------------------------------
Private Function GetExportFolderPath() As String
    GetExportFolderPath = UserSettings_Get("VCSParams", "ExportFolder", "")
End Function

Private Sub SetExportFolderPath(exportPath As String)
    If exportPath = "" Then Exit Sub
    UserSettings_Set "VCSParams", "ExportFolder", (exportPath)
End Sub

Private Sub ResetExportFolderPath()
    UserSettings_Del "VCSParams", "ExportFolder"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTablesToSaveData
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : retrieve list of table for what to export data
'             (list is set at start time with pub_LIBVCS_LoadVCSModel)
'---------------------------------------------------------------------------------------
Private Function getTablesToSaveData(Optional TablesToInclude_List As String, Optional IncludedTablesToExclude_List As String) As Collection
    Dim TablesToSaveData    As New Collection
    Dim Includes            As Variant
    Dim Excludes            As Variant
    Dim tblDef              As Variant
    Dim iName               As Variant
    Dim eName               As Variant
    
    'UserSettings_Set TablesToInclude_List , IncludedTablesToExclude_List
    UserSettings_Set "VCSParams", "TablesToInclude", (TablesToInclude_List)
    UserSettings_Set "VCSParams", "TablesToExclude", (IncludedTablesToExclude_List)
    Includes = Split(TablesToInclude_List, ";", , vbTextCompare)
    Excludes = Split(IncludedTablesToExclude_List, ";", , vbTextCompare)
    If UBound(Includes) <> -1 Then
        For Each tblDef In CurrentDb.TableDefs
            For Each iName In Includes
                If iName <> "" And InStr(tblDef.Name, iName) <> 0 Then
                    For Each eName In Excludes
                        If eName <> "" And InStr(tblDef.Name, eName) <> 0 Then
                            iName = Empty
                            Exit For
                        End If
                    Next
                    If Not IsEmpty(iName) Then
                        Debug.Print tblDef.Name
                        TablesToSaveData.Add tblDef.Name
                        Exit For
                    End If
                End If
            Next
        Next
        Debug.Print
    End If
    Set getTablesToSaveData = TablesToSaveData
End Function

Private Function SelectExportFolder(Optional StartingFolder) As Variant
'Requires reference to Microsoft Office 12.0 Object Library.
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
   With fDialog
      .InitialFileName = StartingFolder
      .AllowMultiSelect = False
      .Title = "Please select export folder"
      .Filters.Clear
      If .Show = True Then
        SelectExportFolder = .SelectedItems(1)
      Else
         MsgBox "You clicked Cancel in the folder dialog box."
      End If
   End With
End Function

Private Function SelectImportFolder(Optional StartingFolder) As Variant
'Requires reference to Microsoft Office 12.0 Object Library.
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
   With fDialog
      .InitialFileName = StartingFolder
      .AllowMultiSelect = False
      .Title = "Please select Import folder"
      .Filters.Clear
      If .Show = True Then
        SelectImportFolder = .SelectedItems(1)
      Else
         MsgBox "You clicked Cancel in the folder dialog box."
      End If
   End With
End Function

Private Function SelectImportObjects(Optional StartingFolder) As Collection
'Requires reference to Microsoft Office 12.0 Object Library.
    Dim selectedItem As Variant
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Set SelectImportObjects = New Collection
    With fDialog
        .InitialFileName = StartingFolder
        .AllowMultiSelect = True
        .Title = "Please select Objects to import"
        .Filters.Clear
        .Filters.Add "All files", "*.*"
        If .Show = True Then
            For Each selectedItem In .SelectedItems
                SelectImportObjects.Add selectedItem
            Next
        Else
            MsgBox "You clicked Cancel in the folder dialog box."
        End If
   End With
End Function

Private Function getProjectName()
    getProjectName = CurrentProject.FullName ' VBE.ActiveVBProject.Name
End Function

'---------------------------------------------------------------------------------------
' Procedure : UserSettings_Get/Set/Del
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : functions to store settings in user registry
'---------------------------------------------------------------------------------------
Private Function UserSettings_Get(DBSection As String, DBUserSettingKey As String, Optional default) As Variant
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = getProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'pub_Settings_Get("Environment")
    UserSettings_Get = GetSetting(DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey, default)
End Function

Private Sub UserSettings_Set(DBSection As String, DBUserSettingKey As String, DBUserSettingValue As Variant)
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = getProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'pub_Settings_Get("Environment")
    SaveSetting DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey, DBUserSettingValue
End Sub

Private Sub UserSettings_Del(DBSection As String, Optional DBUserSettingKey As String)
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = getProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'DBSettings.Setting("Environment")
    If DBUserSettingKey <> "" Then
        'Delete only the specified Key
        DeleteSetting DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey
    Else
        'Delete all the specified Section
        DeleteSetting DbName & " - " & DBEnvironment, DBSection
    End If
End Sub