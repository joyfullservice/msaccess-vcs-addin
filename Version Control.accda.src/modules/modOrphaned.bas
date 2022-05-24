Attribute VB_Name = "modOrphaned"
'---------------------------------------------------------------------------------------
' Module    : modVCSUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions relating to detecting and removing orphaned items in relation
'           : to the source code files and existing database objects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName = "modOrphaned"


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFolders
' Author    : Casper Englund
' Date      : 2020-06-04
' Purpose   : Clears existing source folders that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFolders(cType As IDbComponent)
    
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim colNames As Collection
    Dim cItem As IDbComponent
    Dim oFolder As Scripting.Folder
    Dim oSubFolder As Scripting.Folder
    Dim strSubFolderName As String
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names for actual database objects
    Perf.OperationStart "Clear Orphaned Folders"
    Set colNames = New Collection
    Set dItems = cType.GetAllFromDB(False)
    For Each varKey In dItems.Keys
        Set cItem = dItems(varKey)
        colNames.Add FSO.GetFileName(cItem.SourceFile)
    Next varKey
    
    Set oFolder = FSO.GetFolder(cType.BaseFolder)
    For Each oSubFolder In oFolder.SubFolders
            
        strSubFolderName = oSubFolder.Name
        ' Remove any subfolder that doesn't have a matching name.
        If Not InCollection(colNames, strSubFolderName) Then
            ' Object not found in database. Remove subfolder.
            oSubFolder.Delete True
            Log.Add "  Removing orphaned folder: " & strSubFolderName, Options.ShowDebug
        End If
        
    Next oSubFolder
    
    ' Remove base folder if we don't have any subfolders or files in it
    With oFolder
        If .SubFolders.Count = 0 Then
            If .Files.Count = 0 Then .Delete
        End If
    End With
    Perf.OperationEnd
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 2/23/2021
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'           : Note that this function is integrated with the index, so deleted files
'           : are removed from the index, and potential conflicts are logged as well.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent, ParamArray StrExtensions())
    
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim dBaseNames As Dictionary
    Dim dExtensions As Dictionary
    Dim strBaseName As String
    Dim strFileName As String
    Dim strFile As String
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim varExt As Variant
    Dim strExt As String
    Dim cItem As IDbComponent
    Dim strHash As String
    Dim cXItem As clsVCSIndexItem
    Dim dteModified As Date
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Set up dictionary objects for case-insensitive comparison
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    Set dExtensions = New Dictionary
    dExtensions.CompareMode = TextCompare
    
    ' Cache a list of base source file names for actual database objects
    Perf.OperationStart "Clear Orphaned Files"
    Set dItems = cType.GetAllFromDB(False)
    For Each varKey In dItems.Keys
        Set cItem = dItems(varKey)
        dBaseNames.Add FSO.GetBaseName(cItem.SourceFile), vbNullString
    Next varKey
    
    ' Build dictionary of allowed extensions
    For Each varExt In StrExtensions
        dExtensions.Add varExt, vbNullString
    Next varExt
        
    ' Loop through files in folder
    Set oFolder = FSO.GetFolder(cType.BaseFolder)
    For Each oFile In oFolder.Files
    
        ' Get base name and file extension
        ' (For performance reasons, minimize property access on oFile)
        strFileName = oFile.Name
        strFile = cType.BaseFolder & strFileName
        strBaseName = FSO.GetBaseName(strFileName)
        strExt = Mid$(strFileName, Len(strBaseName) + 2)
        
        ' See if extension exists in cached list
        If dExtensions.Exists(strExt) Then
            
            ' See if base file name exists in list of database objects
            If Not dBaseNames.Exists(strBaseName) Then
                
                ' Object not found in database. Check the index
                If VCSIndex.Exists(cType, strFileName) Then
                    
                    ' If file is unchanged from the index, we can go ahead and delete it.
                    ' (The source file matches the last version imported or exported)
                    strHash = VCSIndex.GetFilePropertyHash(strFile)
                    If VCSIndex.Item(cType, strFileName).FilePropertiesHash = strHash Then
                    
                        ' Remove file and index entry
                        Log.Add "  Removing orphaned file: " & cType.BaseFolder & strFileName, Options.ShowDebug
                        DeleteFile strFile, True
                        VCSIndex.Remove cType, strFile
                    Else
                        ' File properties different from index. Add as a conflict to resolve.
                        ' (This can happen when the last export was during a different daylight savings time
                        ' setting, as the past file modified date returned by FSO is not adjusted for DST.)
                        Log.Add "  Orphaned source file does not match last export: " & strFile, Options.ShowDebug
                        VCSIndex.Conflicts.Add cType, strFile, 0, GetLastModifiedDate(strFile), ercDelete, strFile, ercDelete
                    End If
                    
                Else
                    ' Object does not exist in the index. It might be a new file added
                    ' by another developer. Don't delete it, as it may need to be merged
                    ' into the database. (Defaults to skip deleting the file)
                    Log.Add "  Found new source file: " & strFile, Options.ShowDebug
                    VCSIndex.Conflicts.Add cType, strFile, 0, GetLastModifiedDate(strFile), ercDelete, strFile, ercSkip
                End If
            End If
        End If
    Next oFile
    
    ' Remove base folder if we don't have any files in it
    If oFolder.Files.Count = 0 Then oFolder.Delete True
    Perf.OperationEnd
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetObjectNameListFromFileList
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Return a dictionary of unique object names from the file names.
'           : (Translates the names from the safe file name to the original object name.)
'---------------------------------------------------------------------------------------
'
Public Function GetObjectNameListFromFileList(dFileList As Dictionary) As Dictionary
    
    Dim varKey As Variant
    Dim dNames As Dictionary
    Dim strName As String
    
    Set dNames = New Dictionary
    For Each varKey In dFileList.Keys
        strName = GetObjectNameFromFileName(CStr(varKey))
        If Not dNames.Exists(strName) Then dNames.Add strName, vbNullString
    Next varKey
    
    Set GetObjectNameListFromFileList = dNames
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDbObjectNameList
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Returns a dictionary of the object names. (Using `DBObject.Name` property)
'---------------------------------------------------------------------------------------
'
Public Function GetDbObjectNameList(cComponent As IDbComponent) As Dictionary
    
    Dim dAllItems As Dictionary
    Dim cItem As IDbComponent
    Dim varKey As Variant
    Dim strName As String
    Dim dNames As Dictionary
    
    ' Return dictionary of all items
    Set dAllItems = cComponent.GetAllFromDB(False)
    Set dNames = New Dictionary
    
    ' Get name for each item
    For Each varKey In dAllItems.Keys
        Set cItem = dAllItems(varKey)
        strName = cItem.DbObject.Name
        If Not dNames.Exists(strName) Then dNames.Add strName, cItem
    Next varKey
    
    ' Return list of object names
    Set GetDbObjectNameList = dNames
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveOrphanedDatabaseObjects
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Remove orphaned database objects when the source file no longer exists
'           : for that object. (Works for most standard database objects)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveOrphanedDatabaseObjects(cCategory As IDbComponent)

    Dim varKey As Variant
    Dim dObjects As Dictionary
    Dim dSource As Dictionary
    Dim strName As String
    Dim cItem As IDbComponent
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Get list of source file object names
    Set dSource = GetObjectNameListFromFileList(cCategory.GetFileList)
    
    ' Get list of current database objects
    Set dObjects = GetDbObjectNameList(cCategory)
    
    ' Loop through objects, getting list
    For Each varKey In dObjects.Keys
        strName = CStr(varKey)
        If Not dSource.Exists(strName) Then
            ' No source file found for this object
            Set cItem = dObjects(varKey)
            If cItem.IsModified Then
                ' Item should be removed, but appears to be modified in the database
                With cItem
                    ' Log this item as a conflict so the user can make a decision on whether to
                    ' proceed with removing the orphaned database object.
                    VCSIndex.Conflicts.Add cItem, .SourceFile, VCSIndex.Item(cItem).ExportDate, _
                        0, ercDelete, .SourceFile, ercSkip
                    Log.Add "The " & cCategory.Name & " '" & strName & "' appears to have been modified since the last export, " & _
                        "but does not have a corresponding source file. Normally this would be deleted as an orphaned object during " & _
                        "the merge operation, but has been flagged as a conflict for user resolution. (Could not find " & .SourceFile & ")", False
                End With
            Else
                ' Index is current. Safe to remove
                Log.Add "Removing orphaned " & cCategory.Name & " '" & strName & "'"
                DoCmd.DeleteObject cCategory.ComponentType, strName
                CatchAny eelError, "Error removing orphaned " & cCategory.Name & " '" & strName & "'", ModuleName & ".RemoveOrphanedDatabaseObjects"
            End If
        End If
    Next varKey
    
End Sub
