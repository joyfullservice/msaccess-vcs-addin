'@Folder("MSAccessVCS")
Option Compare Database
Option Explicit

Public Function pub_LIBVCS_LoadVCSModel(Optional commaSeparatedListOfNamePartsOfTablesToIncludeForDataSave As String = "tbl_L_;", _
                                        Optional commaSeparatedListOfNamePartsOfTablesToExcludeFromIncludedOnes As String = "tbl_L_NotThisOne;" _
                                        ) As Boolean

    Dim exportPath As String
    exportPath = getExportPath()
    
    If exportPath <> "" Then
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
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetSourceFolderPath
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : Returns the saved or default export path
'---------------------------------------------------------------------------------------

Private Function GetExportFolderPath() As String
    Dim default
    default = CurrentProject.Path & "\" & VBE.ActiveVBProject.Name & ".src\"
    GetExportFolderPath = UserSettings_Get("VCSParams", "ExportFolder", default)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetSourceFolderPath
' Author    : Marco Salernitano
' Date      : 25-Jun-2019
' Purpose   : save the last used export path
'---------------------------------------------------------------------------------------

Private Sub SetExportFolderPath(exportPath As String)
    If exportPath = "" Then Exit Sub
    UserSettings_Set "VCSParams", "ExportFolder", (exportPath)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getExportPath
' Author    : Marco Salernitano
' Date      : 5/15/2015
' Purpose   : Verifies the export path and give to user the possibility to choose a custom one
'---------------------------------------------------------------------------------------
'
Private Function getExportPath() As String
    
    getExportPath = GetExportFolderPath
    ' If code reaches here, we don't have a copy of the path
    ' in the cached list of verified paths. Verify and add
    If Dir(getExportPath, vbDirectory) = "" Then
    ' Path does not seem to exist.
        Dim Answer
        Answer = MsgBox("Saved or Default Path:" & vbCrLf & vbCrLf & getExportPath & vbCrLf & vbCrLf & "doesn't exist." & vbCrLf & vbCrLf & "Answer Yes to create it, No to choose another one or Cancel to abort.", vbYesNoCancel Or vbExclamation, "Export folder not found!")
        Select Case Answer
            Case vbYes
                ' Create it.
                MkDirIfNotExist getExportPath
            Case vbNo
                ' Ask for alternate path
                getExportPath = CStr(SelectFolder(CurrentProject.Path))
                If getExportPath <> "" Then
                    If Right(getExportPath, 1) <> "\" Then getExportPath = getExportPath & "\"
                End If
            Case Else
                Exit Function
        End Select
        If Dir(getExportPath, vbDirectory) <> "" Then SetExportFolderPath getExportPath
    End If
End Function

Private Function getTablesToSaveData(Optional TablesToInclude_List As String, Optional IncludedTablesToExclude_List As String) As Collection
    Dim TablesToSaveData    As New Collection
    Dim Includes            As Variant
    Dim Excludes            As Variant
    Dim tblDef              As Variant
    Dim iName               As Variant
    Dim eName               As Variant
    
    Includes = Split(TablesToInclude_List, ";", vbTextCompare)
    Excludes = Split(IncludedTablesToExclude_List, ";", vbTextCompare)
    If UBound(Includes) <> -1 Then
        For Each tblDef In CurrentDb.TableDefs
            For Each iName In Includes
                If InStr(tblDef.Name, iName) <> 0 Then
                    For Each eName In Excludes
                        If InStr(tblDef.Name, eName) <> 0 Then
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

Private Function SelectFolder(Optional StartingFolder) As Variant
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
        SelectFolder = .SelectedItems(1)
      Else
         MsgBox "You clicked Cancel in the folder dialog box."
      End If
   End With
End Function

Private Function cProjectName()
    cProjectName = VBE.ActiveVBProject.Name
End Function

Private Function UserSettings_Get(DBSection As String, DBUserSettingKey As String, Optional default) As Variant
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = cProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'pub_Settings_Get("Environment")
    UserSettings_Get = GetSetting(DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey, default)
End Function

Private Sub UserSettings_Set(DBSection As String, DBUserSettingKey As String, DBUserSettingValue As Variant)
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = cProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'pub_Settings_Get("Environment")
    SaveSetting DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey, DBUserSettingValue
End Sub

Private Sub UserSettings_Del(DBSection As String, Optional DBUserSettingKey As String)
    Dim DbName As String
    Dim DBEnvironment As String
    On Error Resume Next
    DbName = cProjectName 'CurrentDb.Properties("AppTitle")
    DBEnvironment = "DEV"  'DBSettings.Setting("Environment")
    If DBUserSettingKey <> "" Then
        'Delete only the specified Key
        DeleteSetting DbName & " - " & DBEnvironment, DBSection, DBUserSettingKey
    Else
        'Delete all the specified Section
        DeleteSetting DbName & " - " & DBEnvironment, DBSection
    End If
End Sub