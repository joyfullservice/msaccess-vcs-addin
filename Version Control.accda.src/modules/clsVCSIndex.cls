VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVCSIndex
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   : Maintain index of source files and database objects so that changes
'           : can be detected.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' File name for index
Private Const cstrFileName As String = "vcs-index.json"

' General properties
Public MergeBuildDate As Date
Public FullBuildDate As Date
Public ExportDate As Date
Public FullExportDate As Date
Public OptionsHash As String
Public Disabled As Boolean

' Git integration
Public LastMergedCommit As String

' Action types for update function
Public Enum eIndexOperationType
    eatExport
    eatImport
    eatAltExport    ' Alternate export folder, such as a scan for changes
End Enum

' Index of component/file information, based on source files
Private m_dIndex As Dictionary
Private m_dGitIndex As Dictionary
Private m_strFile As String
Private m_Git As clsGitIntegration
Private m_Conflicts As clsConflicts
Private m_strTempExportFolderPath As String


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Load the state for the project.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromFile(Optional strFile As String)

    Dim dFile As Dictionary
    Dim dItem As Dictionary
    Dim varKey As Variant
    Dim blnSavedValue As Boolean

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Reset class to blank values
    Call Class_Initialize

    ' Load properties
    m_strFile = strFile
    If m_strFile = vbNullString Then m_strFile = DefaultFilePath
    If FSO.FileExists(m_strFile) Then
        ' Convert dates back to local time for processing
        blnSavedValue = JsonOptions.ConvertDateToIso
        JsonOptions.ConvertDateToIso = True
        Set dFile = ReadJsonFile(m_strFile)
        JsonOptions.ConvertDateToIso = blnSavedValue
        If Not dFile Is Nothing Then
            If dFile.Exists("Items") Then
                ' Load properties from class
                For Each varKey In dFile("Items").Keys
                    If m_dIndex.Exists(varKey) Then
                        Select Case varKey
                            Case "Components", "AlternateExport"
                                ' Load as dictionary
                                Set dItem = dFile("Items")(varKey)
                                Set m_dIndex(varKey) = dItem
                            Case Else
                                ' Set property by name
                                CallByName Me, CStr(varKey), VbLet, Nz(dFile("Items")(varKey), 0)
                        End Select
                    End If
                Next varKey
            End If
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Save
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Save to a file
'---------------------------------------------------------------------------------------
'
Public Sub Save(Optional strInFolder As String)

    Dim varKey As Variant
    Dim varValue As Variant
    Dim strFile As String
    Dim blnSavedValue As Boolean

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Load dictionary from properties
    For Each varKey In m_dIndex.Keys
        Select Case varKey
            Case "Components", "AlternateExport"
            Case Else
                varValue = CallByName(Me, CStr(varKey), VbGet)
                ' Save blank dates as null
                If Right(varKey, 4) = "Date" Then
                    m_dIndex(varKey) = ZNDate(varValue)
                Else
                    m_dIndex(varKey) = CStr(varValue)
                End If
        End Select
    Next varKey

    ' Sort files and components
    SortComponentSection m_dIndex, "Components"

    ' Remove the AlternateExport section, since this is not
    ' needed after the completion of the export process.
    If m_dIndex.Exists("AlternateExport") Then m_dIndex.Remove "AlternateExport"

    ' Build file path
    If strInFolder = vbNullString Then
        strFile = m_strFile
    Else
        strFile = StripSlash(strInFolder) & PathSep & cstrFileName
    End If

    ' Turn on ISO date conversion to save index dates in UTC
    blnSavedValue = JsonOptions.ConvertDateToIso
    JsonOptions.ConvertDateToIso = True

    ' Save index to file
    If m_strFile <> vbNullString Then
        WriteFile BuildJsonFile(TypeName(Me), m_dIndex, "Version Control System Index"), strFile
    End If

    ' Restore previous setting
    JsonOptions.ConvertDateToIso = blnSavedValue

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SortComponentSection
' Author    : Adam Waller
' Date      : 9/1/2022
' Purpose   : Save a named section of components as
'---------------------------------------------------------------------------------------
'
Private Sub SortComponentSection(ByRef dIndex As Dictionary, strSection As String)

    Dim dComponents As Dictionary
    Dim varCat As Variant

    If dIndex.Exists(strSection) Then
        Set dComponents = dIndex(strSection)
        For Each varCat In dComponents.Keys
            ' Sort list of files in each category
            Set dComponents(varCat) = SortDictionaryByKeys(dComponents(varCat))
        Next varCat

        ' Sort list of categories
        Set dIndex(strSection) = SortDictionaryByKeys(dComponents)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Updates an item in the index.
'---------------------------------------------------------------------------------------
'
Public Function Update(cItem As IDbComponent, intAction As eIndexOperationType, _
    strFileHash As String, Optional strOtherHash As String, Optional dteDateTime As Date) As Dictionary

    Dim dItem As Dictionary
    Dim strSection As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Function

    ' Determine section
    If intAction = eatAltExport Then
        strSection = "AlternateExport"
    Else
        strSection = "Components"
    End If

    ' Look up dictionary item, creating if needed.
    Set dItem = Me.Item(cItem, , strSection).dParent

    ' Update dictionary values
    With dItem

        ' Update hash
        If strFileHash = vbNullString Then
            ' Remove hash if not used.
            If .Exists("FileHash") Then .Remove "FileHash"
        Else
            .Item("FileHash") = strFileHash
        End If

        ' Code module hash (Detects changes to VBA code that are not reflected in modified date)
        If strOtherHash = vbNullString Then
            ' Remove hash if not used.
            If .Exists("OtherHash") Then .Remove "OtherHash"
        Else
            .Item("OtherHash") = strOtherHash
        End If

        ' Add timestamp, defaulting to now
        If dteDateTime = 0 Then dteDateTime = Now
        Select Case intAction
            Case eatExport, eatAltExport
                .Item("ExportDate") = dteDateTime
            Case eatImport
                .Item("ImportDate") = dteDateTime
        End Select

        ' Save timestamp of exported source file.
        dteDateTime = GetSourceModifiedDate(cItem, cItem.SourceFile)
        .Item("SourceModified") = ZNDate(dteDateTime)

        ' Save hash of file properties
        .Item("FilePropertiesHash") = GetSourceFilesPropertyHash(cItem)

    End With

    ' Return dictionary object in case the caller wants to
    ' manipulate additional values.
    Set Update = dItem

End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateFromAltExport
' Author    : Adam Waller
' Date      : 9/1/2022
' Purpose   : Update the index entry from the alternate export section. This is done
'           : after resolving a conflict where the exported file in the temp folder is
'           : moved to the primary export folder. (This way we don't have to compute
'           : the hashes and index values again.
'---------------------------------------------------------------------------------------
'
Public Sub UpdateFromAltExport(cItem As IDbComponent)

    Dim strFile As String
    Dim dAltItem As Dictionary
    Dim varKey As Variant

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Get file name from component
    strFile = FSO.GetFileName(cItem.SourceFile)

    ' Look for entry in AlternateExport section
    With m_dIndex("AlternateExport")
        If .Exists(cItem.Category) Then
            If .Item(cItem.Category).Exists(strFile) Then
                ' Get reference to alternate export entry
                Set dAltItem = .Item(cItem.Category).Item(strFile)
            End If
        End If
    End With

    ' Update entry in main component section with values
    ' from the alternate export section.
    If Not dAltItem Is Nothing Then
        With LoadItem(cItem, strFile, "Components").dParent
            For Each varKey In dAltItem.Keys
                .Item(varKey) = dAltItem(varKey)
            Next varKey
        End With
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Remove an item from the index when the object and file no longer exist.
'---------------------------------------------------------------------------------------
'
Public Sub Remove(cItem As IDbComponent, Optional strSourceFile As String)

    Dim strFile As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(Nz2(strSourceFile, cItem.SourceFile))

    ' Remove dictionary objects.
    With m_dIndex("Components")
        If .Exists(cItem.Category) Then
            If .Item(cItem.Category).Exists(strFile) Then
                .Item(cItem.Category).Remove strFile
                ' Remove category if no more items
                If .Item(cItem.Category).Count = 0 Then
                    .Remove cItem.Category
                End If
            End If
        End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetItem
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Returns a dictionary object with the saved values, creating if needed.
'---------------------------------------------------------------------------------------
'
Public Function Item(cItem As IDbComponent, Optional strSourceFile As String, _
    Optional strSection As String = "Components") As clsVCSIndexItem

    Dim strFile As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Function

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(Nz2(strSourceFile, cItem.SourceFile))
    Set Item = LoadItem(cItem, strFile, strSection)

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadItem
' Author    : Adam Waller
' Date      : 9/23/2021
' Purpose   : Loads the item from the index, creating if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Function LoadItem(cItem As IDbComponent, strSourceFile As String, strSection As String) As clsVCSIndexItem

    Dim dItem As Dictionary

    ' Get or create dictionary objects.
    With m_dIndex(strSection)
        If Not .Exists(cItem.Category) Then Set .Item(cItem.Category) = New Dictionary
        If Not .Item(cItem.Category).Exists(strSourceFile) Then Set .Item(cItem.Category)(strSourceFile) = New Dictionary
        Set dItem = .Item(cItem.Category)(strSourceFile)
        Set LoadItem = New clsVCSIndexItem
        With LoadItem
            ' Load properties to class
            Set .dParent = dItem
            .ExportDate = Nz2(dNZ(dItem, "ExportDate"), 0)
            .ImportDate = Nz2(dNZ(dItem, "ImportDate"), 0)
            .SourceModified = Nz2(dNZ(dItem, "SourceModified"), 0)
            .FilePropertiesHash = dNZ(dItem, "FilePropertiesHash")
            .FileHash = dNZ(dItem, "FileHash")
            .OtherHash = dNZ(dItem, "OtherHash")
        End With
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Returns true if the file exists in the index.
'---------------------------------------------------------------------------------------
'
Public Function Exists(cCategory As IDbComponent, strSourceFilePath As String) As Boolean

    Dim strFile As String
    Dim blnExists

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(strSourceFilePath)

    ' See if the entry exists in the index
    With m_dIndex("Components")
        If .Exists(cCategory.Category) Then
            blnExists = .Item(cCategory.Category).Exists(strFile)
        End If
    End With

    ' Return result
    Exists = blnExists

End Function


'---------------------------------------------------------------------------------------
' Procedure : DefaultDevModeHash
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Return a hash of the printer settings from the default printer. (Used to
'           : determine whether a form or report is using any specific or custom
'           : print settings, thereby requiring us to save the print settings.
'---------------------------------------------------------------------------------------
'
Public Property Get DefaultDevModeHash() As String

    ' Cache the result for future calls
    Static strHash As String

    If strHash = vbNullString Then
        With New clsDevMode
            .LoadFromDefaultPrinter
            strHash = .GetHash
        End With
    End If

    ' Return hash
    DefaultDevModeHash = strHash

End Property


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Return a collection of source files that appear to be different from
'           : the index hash of the previous export. File modified dates will be used
'           ' to determine which files have changed.
'           : NOTE: This will also include source file paths for database objects that
'           : no longer exist. (Orphaned database objects not represented in source.)
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCategory As IDbComponent) As Dictionary

    Dim dAllFiles As Dictionary
    Dim dItems As Dictionary
    Dim varItem As Variant
    Dim cItem As IDbComponent
    Dim strHash As String
    Dim varFile As Variant
    Dim strFile As String
    Dim strPath As String
    Dim blnModified As Boolean
    Dim blnMissing As Boolean

    Set GetModifiedSourceFiles = New Dictionary
    With GetModifiedSourceFiles

        ' Get a full list of source files and database objects
        Set dAllFiles = cCategory.GetFileList
        Set dItems = cCategory.GetAllFromDB

        ' Loop through files and check modified dates.
        ' Get a list of all the files for this component.
        For Each varItem In dAllFiles
            strFile = varItem

            ' Reset flags
            blnModified = True
            blnMissing = False

            ' See if the database object currently exists
            If Not cCategory.SingleFile Then
                blnMissing = Not dItems.Exists(strFile)
            End If

            ' Check the index for any existing entry, so we can compare the saved modified date.
            If (Not blnMissing) And Me.Exists(cCategory, strFile) Then
                ' Get hash of current source files.
                strHash = GetSourceFilesPropertyHash(cCategory, strFile)
                ' Build the path to find the item in the index.
                strPath = Join(Array("Components", cCategory.Category, FSO.GetFileName(strFile), "FilePropertiesHash"), "\")
                ' File is considered not modified if the index date matches the file modification date.
                blnModified = Not (dNZ(m_dIndex, strPath) = strHash)
            End If

            ' Add modified files to collection
            If blnModified Then .Add strFile, vbNullString
        Next varItem

        ' Now loop through database objects, and compare to source file list
        ' to identify orphaned database objects no longer represented in source.
        For Each varItem In dItems.Items
            Set cItem = varItem
            ' We are only concerned about non single-file items.
            If Not cItem.SingleFile Then
                ' See if the path exists in the list of source files
                If Not dAllFiles.Exists(cItem.SourceFile) Then
                    ' This object does not appear to exist in the source files.
                    ' Add to list of files to process. (This will remove the
                    ' existing database object during the merge.)
                    .Add cItem.SourceFile, vbNullString
                End If
            End If
        Next varItem
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Conflicts
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Expose conflicts as an extension of the index
'---------------------------------------------------------------------------------------
'
Public Function Conflicts() As clsConflicts
    If m_Conflicts Is Nothing Then Set m_Conflicts = New clsConflicts
    Set Conflicts = m_Conflicts
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckExportConflicts
' Author    : Adam Waller
' Date      : 9/15/2021
' Purpose   : Check for potential conflicts when exporting these items
'---------------------------------------------------------------------------------------
'
Public Sub CheckExportConflicts(dDbComponents As Dictionary)

    Dim varKey As Variant
    Dim cItem As IDbComponent
    Dim sngStart As Single

    ' Capture start time so we can display some progress if this
    ' scan takes a while.
    sngStart = Timer

    ' Loop through all components
    For Each varKey In dDbComponents.Keys
        Set cItem = dDbComponents(varKey)
        If IsExportConflict(cItem) Then
            ' Add to list of conflicts
            Me.Conflicts.Add cItem, _
                CStr(varKey), _
                Me.Item(cItem).ExportDate, _
                GetSourceModifiedDate(cItem, cItem.SourceFile), _
                ercOverwrite
        End If
        ' Increment the progress bar
        Log.Increment
        ' Exit loop if single source file
        If cItem.SingleFile Then Exit For
        ' Display more detail if export runs more than 1 second.
        If sngStart > 0 And (Timer - sngStart > 1) Then
            Log.Add " - Checking " & dDbComponents.Count & " " & LCase(cItem.Category) & "...", , , , , True
            sngStart = 0
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsExportConflict
' Author    : Adam Waller
' Date      : 9/15/2021
' Purpose   : Returns true if exporting the item would conflict with an existing file
'           : that is different AND newer than the last export of this object.
'---------------------------------------------------------------------------------------
'
Private Function IsExportConflict(cItem As IDbComponent) As Boolean

    Dim blnConflict As Boolean
    Dim strHash As String
    Dim strFile As String
    Dim strTempFile As String
    Dim blnCompareTemp As Boolean

    ' Check for special cases based on component type
    If cItem.ComponentType = edbTableData Then
        ' Table data is not considered an export conflict, and should
        ' always be overwritten.
        IsExportConflict = False
        Exit Function
    End If

    ' See if the export file even exists before we worry about conflicts.
    strFile = cItem.SourceFile
    If FSO.FileExists(strFile) Then

        ' Get a hash of the file properties
        strHash = GetSourceFilesPropertyHash(cItem)

        ' Check to see if this object is in the index
        ' (Has been exported/imported before.)
        If Me.Exists(cItem, strFile) Then

            ' Item found in index.
            ' Check to see if it matches the property hash in the index.
            If Me.Item(cItem).FilePropertiesHash = strHash Then
                ' Source file has not changed since last export/import.
                ' We should be good to overwrite it.
            Else
                ' Doesn't match. This could be a conflict, depending on the dates.
                ' If the modified date of the source file is NEWER than the
                ' last export date...
                If Me.Item(cItem).ExportDate < GetSourceModifiedDate(cItem) Then
                    ' ...and the content is different, then we have a conflict.
                    blnCompareTemp = True
                End If
            End If
        Else
            ' Without an existing index entry, we will need to compare a fresh
            ' export of the item to the source file and check for differing content.
            blnCompareTemp = True
        End If
    End If

    ' Compare the content hash of a fresh export to see if it matches the existing file.
    If blnCompareTemp Then
        strTempFile = Replace(strFile, Options.GetExportFolder, GetTempExportFolder)
        Log.Add "  Exporting " & cItem.Name, False
        cItem.Export strTempFile
        blnConflict = Not SourceMatches(cItem, strFile, strTempFile)
    End If

    ' Return result
    IsExportConflict = blnConflict

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckImportConflicts
' Author    : Adam Waller
' Date      : 2/24/2023
' Purpose   : Check for potential conflicts when merging into this database.
'---------------------------------------------------------------------------------------
'
Public Sub CheckMergeConflicts(cComponent As IDbComponent, dFiles As Dictionary)

    Dim varKey As Variant
    Dim cItem As IDbComponent
    Dim sngStart As Single
    Dim strFile As String
    Dim intConflict As eResolveConflict

    ' Capture start time so we can display some progress if this
    ' scan takes a while.
    sngStart = Timer
    Perf.OperationStart "Check Merge Conflicts"

    ' Loop through the list of source files
    ' (Each key is a full path to a source file)
    For Each varKey In dFiles.Keys
        strFile = varKey
        intConflict = IsMergeConflict(cComponent, strFile)
        If intConflict Then
            ' Get database item
            If cComponent.GetAllFromDB.Exists(strFile) Then
                ' Get individual item (form, report)
                Set cItem = cComponent.GetAllFromDB.Item(strFile)
            Else
                ' Get parent class for single file items
                Set cItem = cComponent
            End If
            ' Add to list of conflicts
            Me.Conflicts.Add cItem, _
                strFile, _
                cItem.DateModified, _
                GetSourceModifiedDate(cComponent, strFile), _
                intConflict, _
                strFile, _
                ercSkip
        End If
        ' Increment the progress bar
        Log.Increment
        ' Exit loop if single source file
        If cComponent.SingleFile Then Exit For
        ' Display more detail if export runs more than 1 second.
        If sngStart > 0 And (Timer - sngStart > 1) Then
            Log.Add " - Checking " & dFiles.Count & " " & LCase(cComponent.Category) & "...", , , , , True
            sngStart = 0
        End If
    Next varKey
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsMergeConflict
' Author    : Adam Waller
' Date      : 12/5/2023
' Purpose   : Returns true if we are about to overwrite changes in the database that
'           : don't match the current source file.
'---------------------------------------------------------------------------------------
'
Private Function IsMergeConflict(cItemClass As IDbComponent, strSourceFile As String) As eResolveConflict

    Dim blnConflict As Boolean
    Dim cItem As IDbComponent
    Dim strTempFile As String

    ' Check for special cases based on component type
    If cItemClass.ComponentType = edbTableData Then
        ' Table data is not considered an import conflict, and should
        ' always be overwritten.
        IsMergeConflict = ercNone
        Exit Function
    End If

    ' Single file objects are treated a little differently
    If cItemClass.SingleFile Then
        If cItemClass.IsModified Then IsMergeConflict = ercOverwrite
    Else
        ' See if the database object even exists before we worry about conflicts.
        If cItemClass.GetAllFromDB.Exists(strSourceFile) Then

            ' Get class instance for database item
            Set cItem = cItemClass.GetAllFromDB.Item(strSourceFile)

            ' See if this object appears to be modified since the last import/export
            If cItem.IsModified Then
                If Not FSO.FileExists(strSourceFile) Then
                    ' Modified object, and source file does not exist. Could be a case
                    ' where we are removing an orphaned database object.
                    IsMergeConflict = ercDelete
                Else
                    ' Just because the modification date is different doesn't always
                    ' mean the object definition has actually changed. By comparing
                    ' a fresh export of the object, we can determine for sure if the
                    ' source files match.
                    strTempFile = Replace(strSourceFile, Options.GetExportFolder, GetTempExportFolder)
                    Log.Add "  Exporting " & cItem.Name, False
                    cItem.Export strTempFile
                    ' Because we are using a newly exported temp file, the file properties
                    ' hash won't match, since the file dates are going to be different.
                    If Not SourceMatches(cItemClass, strSourceFile, strTempFile) Then
                        IsMergeConflict = ercOverwrite
                    End If
                End If
            Else
                ' Based on the index, it does not appear that this item
                ' has changed since the last export or import operation.
                ' We should be good to overwrite (or skip importing) this
                ' object in a merge operation.
            End If
        Else
            ' Database object does not exist. No conflict to merge this file.
        End If

    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceMatches
' Author    : Adam Waller
' Date      : 12/7/2023
' Purpose   : Returns true if the two represented source components match at a source
'           : file level, not counting file modification dates.
'---------------------------------------------------------------------------------------
'
Private Function SourceMatches(cType As IDbComponent, strFile1 As String, strFile2 As String) As Boolean

    Dim dPaths(1 To 2) As Dictionary
    Dim strFiles(1 To 2) As String
    Dim intCnt As Integer
    Dim varExt As Variant
    Dim strPath As String
    Dim varKey As Variant

    strFiles(1) = strFile1
    strFiles(2) = strFile2

    ' Start by retrieving the paths of existing files for this component.
    For intCnt = 1 To 2
        Set dPaths(intCnt) = New Dictionary
        For Each varExt In cType.FileExtensions
            strPath = SwapExtension(strFiles(intCnt), CStr(varExt))
            If FSO.FileExists(strPath) Then
                ' Add path to existing source file
                dPaths(intCnt).Add FSO.GetFileName(strPath), strPath
            End If
        Next varExt
    Next intCnt

    ' Simple comparison of dictionaries
    If dPaths(1).Count <> dPaths(2).Count Then Exit Function
    ' Check for matching keys
    For Each varKey In dPaths(1).Keys
        If Not dPaths(2).Exists(varKey) Then Exit Function
    Next varKey

    ' If count and keys match, check file sizes
    For Each varKey In dPaths(1).Keys
        If FSO.GetFile(dPaths(1).Item(varKey)).Size _
        <> FSO.GetFile(dPaths(2).Item(varKey)).Size Then Exit Function
    Next varKey

    ' Compare hash of content
    For Each varKey In dPaths(1).Keys
        If GetFileHash(CStr(dPaths(1).Item(varKey))) _
        <> GetFileHash(CStr(dPaths(2).Item(varKey))) Then Exit Function
    Next varKey

    ' If we have reached this point without exiting, then the source must match.
    SourceMatches = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTempExportFolder
' Author    : Adam Waller
' Date      : 8/30/2022
' Purpose   : Returns the temp folder used for temporary source exports when the
'           : modified status of the the component cannot be determined without running
'           : an export operation. Persisting this folder during the export operation
'           : allows us to perform a single export of the component while still giving
'           : the user the ability to decide whether to overwrite the existing file.
'           : NOTE: If the folder does not exist, a new one will be created.
'           : This should be cleared after the export operation using the corresponding
'           : ClearTempExportFolder sub.
'---------------------------------------------------------------------------------------
'
Public Property Get GetTempExportFolder() As String
    If m_strTempExportFolderPath = vbNullString Then m_strTempExportFolderPath = GetTempFolder("VCS")
    GetTempExportFolder = m_strTempExportFolderPath & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : ClearTempExportFolder
' Author    : Adam Waller
' Date      : 8/30/2022
' Purpose   : Removes the temporary export folder, including any source files exported
'           : into this folder while scanning for changes.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTempExportFolder()
    If m_strTempExportFolderPath <> vbNullString Then
        If FSO.FolderExists(m_strTempExportFolderPath) Then
            LogUnhandledErrors
            On Error Resume Next
            ' Use FSO to delete the folder and contents
            FSO.DeleteFolder m_strTempExportFolderPath, True
            CatchAny eelWarning, "Unable to delete temp folder: '" & m_strTempExportFolderPath & _
                "' You will need to manually remove this folder.", ModuleName(Me) & ".ClearTempExportFolder"
        End If
        m_strTempExportFolderPath = vbNullString
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Git
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Reference to the Git integration class
'---------------------------------------------------------------------------------------
'
Private Function Git() As clsGitIntegration
    If m_Git Is Nothing Then Set m_Git = New clsGitIntegration
    Set Git = m_Git
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function DefaultFilePath() As String
    If DatabaseFileOpen Then DefaultFilePath = Options.GetExportFolder & cstrFileName
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Set up the dictionary object and keys for reflection.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set m_dIndex = New Dictionary
    With m_dIndex
        .Add "MergeBuildDate", Null
        .Add "FullBuildDate", Null
        .Add "ExportDate", Null
        .Add "FullExportDate", Null
        .Add "OptionsHash", vbNullString
        .Add "LastMergedCommit", vbNullString
        Set .Item("Components") = New Dictionary
        Set .Item("AlternateExport") = New Dictionary
    End With

    ' Load Git integration
    Set m_dGitIndex = Nothing

End Sub
