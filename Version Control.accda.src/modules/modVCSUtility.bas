'---------------------------------------------------------------------------------------
' Module    : modVCSUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Utility functions specific to the VCS project
'---------------------------------------------------------------------------------------

Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : GetAllContainers
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return a collection of all containers.
'           : NOTE: The order doesn't matter for export, but is VERY important
'           : when building the project from source.
'---------------------------------------------------------------------------------------
'
Public Function GetAllContainers() As Collection
    
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
' Procedure : HasMoreRecentChanges
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : Returns true if the database object has been modified more recently
'           : than the exported file or source object.
'---------------------------------------------------------------------------------------
'
Public Function HasMoreRecentChanges(objItem As IDbComponent) As Boolean
    ' File dates could be a second off (between exporting the file and saving the report)
    ' so ignore changes that are less than three seconds apart.
    If objItem.DateModified > 0 And objItem.SourceModified > 0 Then
        HasMoreRecentChanges = (DateDiff("s", objItem.DateModified, objItem.SourceModified) < -3)
    Else
        ' If we can't determine one or both of the dates, return true so the
        ' item is processed as though more recent changes were detected.
        HasMoreRecentChanges = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVCSVersion
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Gets the version of the version control system. (Used to turn off fast
'           : save until a full export has been run with the current version of
'           : the MSAccessVCS addin.)
'---------------------------------------------------------------------------------------
'
Public Function GetVCSVersion() As String

    Dim dbs As Database
    Dim prp As DAO.Property

    Set dbs = CodeDb

    For Each prp In dbs.Properties
        If prp.Name = "AppVersion" Then
            ' Return version
            GetVCSVersion = prp.Value
        End If
    Next prp

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveComponentAsText
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Wrapper for Application.SaveAsText that verifies that the path exists,
'           : and then removes any existing file before saving the object as text.
'---------------------------------------------------------------------------------------
'
Public Sub SaveComponentAsText(intType As AcObjectType, strName As String, strFile As String)
    
    Dim strTempFile As String
    
    On Error GoTo ErrHandler
    
    ' Export to temporary file
    strTempFile = GetTempFile
    Perf.OperationStart "App.SaveAsText()"
    Application.SaveAsText intType, strName, strTempFile
    Perf.OperationEnd
    VerifyPath strFile
    
    ' Sanitize certain object types
    Select Case intType
        Case acForm, acReport, acQuery, acMacro
            ' Sanitizing converts to UTF-8
            If FSO.FileExists(strFile) Then DeleteFile (strFile)
            SanitizeFile strTempFile
            FSO.MoveFile strTempFile, strFile
        Case Else
            ' Handle UCS conversion if needed
            ConvertUcs2Utf8 strTempFile, strFile
    End Select
    
    ' Normal exit
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    If Err.Number = 2950 And intType = acTableDataMacro Then
        ' This table apparently didn't have a Table Data Macro.
        Exit Sub
    Else
        ' Some other error.
        Err.Raise Err.Number
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadComponentFromText
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Load the object into the database from the saved source file.
'---------------------------------------------------------------------------------------
'
Public Sub LoadComponentFromText(intType As AcObjectType, strName As String, strFile As String)

    Dim strTempFile As String
    Dim blnConvert As Boolean
    
    ' Check UCS-2-LE requirement for the current database.
    ' (Cached after first call)
    Select Case intType
        Case acForm, acReport, acQuery, acMacro, acTableDataMacro
            blnConvert = RequiresUcs2
    End Select
    
    ' Only run conversion if needed.
    If blnConvert Then
        ' Perform file conversion, and import from temp file.
        strTempFile = GetTempFile
        ConvertUtf8Ucs2 strFile, strTempFile, False
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strTempFile
        Perf.OperationEnd
        DeleteFile strTempFile, True
    Else
        ' Load UTF-8 file
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strFile
        Perf.OperationEnd
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveNonBuiltInReferences
' Author    : Adam Waller
' Date      : 10/20/2020
' Purpose   : Remove any references that are not built-in. (Sometimes additional
'           : references are added when creating a new database, but not not really
'           : needed in the completed database when building the project from source.)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveNonBuiltInReferences()

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


'---------------------------------------------------------------------------------------
' Procedure : GetOriginalDbFullPathFromSource
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Determine the original full path of the database, based on the files
'           : in the source folder.
'---------------------------------------------------------------------------------------
'
Public Function GetOriginalDbFullPathFromSource(strFolder As String) As String
    
    Dim strPath As String
    Dim dContents As Dictionary
    Dim strFile As String
    
    strPath = StripSlash(strFolder) & "\vbe-project.json"
    If FSO.FileExists(strPath) Then
        Set dContents = ReadJsonFile(strPath)
        strFile = Decrypt(dNZ(dContents, "Items\FileName"))
        If Left$(strFile, 4) = "rel:" Then
            ' Use parent folder of source folder
            GetOriginalDbFullPathFromSource = StripSlash(strFolder) & "\..\" & FSO.GetFileName(Mid$(strFile, 5))
        ElseIf InStr(1, strFile, "@{") > 0 Then
            ' Decryption failed.
            ' We might be able to figure out a relative path from the export path.
            strPath = StripSlash(strFolder) & "\vcs-options.json"
            If FSO.FileExists(strPath) Then
                Set dContents = ReadJsonFile(strPath)
                ' Make sure we can read something, but that the export folder is blank.
                ' (Default, which indicates that it would be in the parent folder of the
                '  source directory.)
                If dNZ(dContents, "Info\AddinVersion") <> vbNullString _
                    And dNZ(dContents, "Options\ExportFolder") = vbNullString Then
                    ' Use parent folder of source directory
                    GetOriginalDbFullPathFromSource = StripSlash(strFolder) & "\..\" & FSO.GetFileName(strFile)
                End If
            End If
        Else
            ' Return full path to file.
            GetOriginalDbFullPathFromSource = strFile
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderHasVcsOptionsFile
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Returns true if the folder as a vcs-options.json file, which is required
'           : to build a project from source files.
'---------------------------------------------------------------------------------------
'
Public Function FolderHasVcsOptionsFile(strFolder As String) As Boolean
    FolderHasVcsOptionsFile = FSO.FileExists(StripSlash(strFolder) & "\vcs-options.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteJsonFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Creates a json file with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'           : Set blnIgnoreHeaderOnlyChanges to true when the file should only be
'           : written when the dItems dictionary value changes. This helps reduce the
'           : number of files marked as changed when the actual content is the same,
'           : but a newer version of VCS was used to export the file.
'---------------------------------------------------------------------------------------
'
Public Sub WriteJsonFile(ClassMe As Object, dItems As Dictionary, strFile As String, strDescription As String, _
    Optional strFileFormat As String = "0.0")
    
    Dim dContents As Dictionary
    Dim dHeader As Dictionary
    Dim dFile As Dictionary
    Dim dExisting As Dictionary
    
    Set dContents = New Dictionary
    Set dHeader = New Dictionary
    
    ' Compare with existing file
    If FSO.FileExists(strFile) Then
        Set dFile = ReadJsonFile(strFile)
        If Not dFile Is Nothing Then
            ' Check file format version
            If strFileFormat <> "0.0" And dNZ(dFile, "Info\Export File Format") <> strFileFormat Then
                ' Rewrite file using new format.
            Else
                If dFile.Exists("Items") Then
                    Set dExisting = dFile("Items")
                    If DictionaryEqual(dItems, dExisting) Then
                        ' No changes to content. Leave existing file.
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    ' Build dictionary structure
    dHeader.Add "Class", TypeName(ClassMe)
    dHeader.Add "Description", strDescription
    If strFileFormat <> "0.0" Then dHeader.Add "Export File Format", strFileFormat
    dContents.Add "Info", dHeader
    dContents.Add "Items", dItems
    
    ' Write to file in Json format
    WriteFile ConvertToJson(dContents, JSON_WHITESPACE), strFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFolders
' Author    : Casper Englund
' Date      : 2020-06-04
' Purpose   : Clears existing source folders that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFolders(cType As IDbComponent)
    
    Dim colNames As Collection
    Dim cItem As IDbComponent
    Dim oFolder As Folder
    Dim oSubFolder As Folder
    Dim strSubFolderName As String
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names for actual database objects
    Set colNames = New Collection
    For Each cItem In cType.GetAllFromDB(False)
        colNames.Add FSO.GetFileName(cItem.SourceFile)
    Next cItem
    
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
    
    ' Remove base folder if we don't have any subfolders in it
    If oFolder.SubFolders.Count = 0 Then oFolder.Delete
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 12/14/2016
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent, ParamArray StrExtensions())
    
    Dim oFolder As Folder
    Dim oFile As File
    Dim colNames As Collection
    Dim strFile As String
    Dim varExt As Variant
    Dim strPrimaryExt As String
    Dim cItem As IDbComponent
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names for actual database objects
    Perf.OperationStart "Clear Orphaned"
    Set colNames = New Collection
    For Each cItem In cType.GetAllFromDB(False)
        colNames.Add FSO.GetFileName(cItem.SourceFile)
    Next cItem
    If colNames.Count > 0 Then strPrimaryExt = "." & FSO.GetExtensionName(colNames(1))
    
    ' Loop through files in folder
    Set oFolder = FSO.GetFolder(cType.BaseFolder)
    For Each oFile In oFolder.Files
    
        ' Check against list of extensions
        For Each varExt In StrExtensions
        
            ' Check for matching extension on wanted list.
            If FSO.GetExtensionName(oFile.Path) = varExt Then
                
                ' Build a file name using the primary extension to
                ' match the list of source files.
                strFile = FSO.GetBaseName(oFile.Name) & strPrimaryExt
                ' Remove any file that doesn't have a matching name.
                If Not InCollection(colNames, strFile) Then
                    ' Object not found in database. Remove file.
                    DeleteFile oFile.ParentFolder.Path & "\" & oFile.Name, True
                    Log.Add "  Removing orphaned file: " & strFile, Options.ShowDebug
                End If
                
                ' No need to check other extensions since we
                ' already had a match and processed the file.
                Exit For
            End If
        Next varExt
    Next oFile
    
    ' Remove base folder if we don't have any files in it
    If oFolder.Files.Count = 0 Then oFolder.Delete True
    Perf.OperationEnd
    
End Sub