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
    Dim oFolder As Folder
    Dim oSubFolder As Folder
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
    
    Dim oFolder As Folder
    Dim oFile As File
    Dim dBaseNames As Dictionary
    Dim dExtensions As Dictionary
    Dim strBaseName As String
    Dim strFile As String
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim varExt As Variant
    Dim strExt As String
    Dim cItem As IDbComponent
    Dim strHash As String
    Dim cXItem As clsVCSIndexItem
    
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
        strFile = oFile.Name
        strBaseName = FSO.GetBaseName(strFile)
        strExt = Mid$(strFile, Len(strBaseName) + 2)
        
        ' See if extension exists in cached list
        If dExtensions.Exists(strExt) Then
            
            ' See if base file name exists in list of database objects
            If Not dBaseNames.Exists(strBaseName) Then
                
                ' Object not found in database. Check the index
                If VCSIndex.Exists(cType, strFile) Then
                    
                    ' If file is unchanged from the index, we can go ahead and delete it.
                    ' (The source file matches the last version imported or exported)
                    strHash = VCSIndex.GetFilePropertyHash(strFile)
                    If VCSIndex.Item(cType, strFile).FilePropertiesHash = strHash Then
                    
                        ' Remove file and index entry
                        Log.Add "  Removing orphaned file: " & strFile, Options.ShowDebug
                        DeleteFile strFile, True
                        VCSIndex.Remove cType, strFile
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


