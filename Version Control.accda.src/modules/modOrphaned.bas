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
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 2/23/2021
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'           : Note that this function is integrated with the index, so deleted files
'           : are removed from the index, and potential conflicts are logged as well.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent)

    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim dBaseNames As Dictionary
    Dim dExtensions As Dictionary
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim varExt As Variant
    Dim cItem As IDbComponent

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
        If cType.SingleFile Then Exit For
    Next varKey

    ' Check for single-file exports with no items
    If cType.SingleFile Then
        If dItems.Count = 0 Then
            ' No more of these items exist in the database.
            ' (For example, IMEX specs)
            If FSO.FileExists(cType.SourceFile) Then
                ' Compare to index to check for any source changes.
                CompareToIndex cType, cType.SourceFile, cType.FileExtensions, dBaseNames
            End If
        End If
    Else
        ' Build dictionary of included extensions
        For Each varExt In cType.FileExtensions
            dExtensions.Add varExt, vbNullString
        Next varExt

        ' Loop through files in folder
        Set oFolder = FSO.GetFolder(cType.BaseFolder)
        For Each oFile In oFolder.Files
            ' Compare to index to find orphaned source files.
            CompareToIndex cType, oFile.Path, dExtensions, dBaseNames
            ' Increment the progress bar as we scan through the files
            Log.Increment
        Next oFile

        ' Remove base folder if we don't have any files in it
        If oFolder.Files.Count = 0 Then oFolder.Delete True
    End If

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompareToIndex
' Author    : Adam Waller
' Date      : 12/5/2023
' Purpose   : Compare the potential orphaned file to the index to determine if we need
'           : to list this as a possible conflict item.
'---------------------------------------------------------------------------------------
'
Private Sub CompareToIndex(cType As IDbComponent, strFilePath As String, dExtensions As Dictionary, dBaseNames As Dictionary)

    Dim strFileName As String
    Dim strBaseName As String
    Dim strExt As String
    Dim strHash As String

    ' Get base name and file extension to build primary source file name
    strFileName = FSO.GetFileName(strFilePath)
    strBaseName = FSO.GetBaseName(strFileName)
    strExt = Mid$(strFileName, Len(strBaseName) + 2)

    ' See if extension exists in cached list
    If dExtensions.Exists(strExt) Then

        ' See if base file name exists in list of database objects
        If Not dBaseNames.Exists(strBaseName) Then

            ' See if this is the primary file extension for this component type
            If StrComp(strExt, dExtensions(0), vbTextCompare) = 0 Then

                ' Object not found in database. Check the index
                If VCSIndex.Exists(cType, strFileName) Then

                    ' If file is unchanged from the index, we can go ahead and delete it.
                    ' (The source file matches the last version imported or exported)
                    strHash = GetSourceFilesPropertyHash(cType, strFilePath)
                    If VCSIndex.Item(cType, strFileName).FilePropertiesHash = strHash Then

                        ' Remove file and index entry
                        Log.Add "  Removing orphaned file: " & cType.BaseFolder & strFileName, Options.ShowDebug
                        DeleteFile strFilePath, True
                        VCSIndex.Remove cType, strFileName
                    Else
                        ' File properties different from index. Add as a conflict to resolve.
                        ' (This can happen when the last export was during a different daylight savings time
                        ' setting, as the past file modified date returned by FSO is not adjusted for DST.)
                        Log.Add "  Orphaned source file does not match last export: " & strFilePath, Options.ShowDebug
                        VCSIndex.Conflicts.Add cType, strFilePath, 0, GetSourceModifiedDate(cType, strFilePath), ercDelete, strFilePath, ercDelete
                    End If
                Else
                    ' Object does not exist in the index. It might be a new file added
                    ' by another developer. Don't delete it, as it may need to be merged
                    ' into the database. (Defaults to skip deleting the file)
                    Log.Add "  Found new source file: " & strFilePath, Options.ShowDebug
                    VCSIndex.Conflicts.Add cType, strFilePath, 0, GetSourceModifiedDate(cType, strFilePath), ercDelete, strFilePath, ercSkip
                End If

            Else
                ' Not the primary extension for this component type.
                ' If the primary source file exists, we will let that file handle evaluate any conflicts
                If Not FSO.FileExists(SwapExtension(strFilePath, CStr(dExtensions(0)))) Then
                    ' The primary source file does not exist. Go ahead and delete this orphaned file.
                    Log.Add "  Removing orphaned file: " & cType.BaseFolder & strFileName, Options.ShowDebug
                    DeleteFile strFilePath, True
                End If
            End If
        End If
    End If

End Sub
