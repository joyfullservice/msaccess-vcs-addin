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

Public Function getExportPath() As String
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

Public Function getImportPath() As String
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
' Procedure : GetSourceFolderPath
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

Private Function getProjectName()
    getProjectName = CurrentProject.FullName ' VBE.ActiveVBProject.Name
End Function

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