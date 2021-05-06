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
Public Sub SaveComponentAsText(intType As AcObjectType, _
                                strName As String, _
                                strFile As String, _
                                Optional cDbObjectClass As IDbComponent = Nothing)
    
    Dim strTempFile As String
    Dim strPrintSettingsFile As String
    
    On Error GoTo ErrHandler
    
    ' Export to temporary file
    strTempFile = GetTempFile
    Perf.OperationStart "App.SaveAsText()"
    Application.SaveAsText intType, strName, strTempFile
    Perf.OperationEnd
    VerifyPath strFile
    
    ' Sanitize certain object types
    Select Case intType
        Case acForm, acReport
            With New clsDevMode
                ' Build print settings file name.
                strPrintSettingsFile = .GetPrintSettingsFileName(cDbObjectClass)
                ' See if we are exporting print vars.
                If Options.SavePrintVars = True Then
                    ' Grab the printer settings before sanitizing the file.
                    .LoadFromExportFile strTempFile
                    ' Only need to save print settings if they are different
                    ' from the default printer settings.
                    If (.GetHash <> VCSIndex.DefaultDevModeHash) And .HasData Then
                        WriteJsonFile TypeName(cDbObjectClass), .GetDictionary, _
                        strPrintSettingsFile, strName & " Print Settings"
                    Else
                        ' No print settings in this object.
                        If FSO.FileExists(strPrintSettingsFile) Then DeleteFile strPrintSettingsFile
                    End If
                Else
                    ' Remove any existing (now orphaned) print settings file.
                    If FSO.FileExists(strPrintSettingsFile) Then DeleteFile strPrintSettingsFile
                End If
            End With
            ' Sanitizing converts to UTF-8
            If FSO.FileExists(strFile) Then DeleteFile strFile
            SanitizeFile strTempFile
            FSO.MoveFile strTempFile, strFile
    
        Case acQuery, acMacro
            ' Sanitizing converts to UTF-8
            If FSO.FileExists(strFile) Then DeleteFile strFile
            SanitizeFile strTempFile
            FSO.MoveFile strTempFile, strFile
            
        Case acModule '(ANSI text file)
            ' Modules may contain extended characters that need UTF-8 conversion
            ' to display correctly in some editors.
            If StringHasExtendedASCII(ReadFile(strTempFile, GetSystemEncoding)) Then
                ' Convert to UTF-8
                ConvertAnsiUtf8 strTempFile, strFile
            Else
                ' Leave as ANSI
                If FSO.FileExists(strFile) Then DeleteFile strFile
                FSO.MoveFile strTempFile, strFile
            End If
        
        Case acTableDataMacro
            ' Table data macros are stored in XML format
            If FSO.FileExists(strFile) Then SanitizeXML strFile
            
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
Public Sub LoadComponentFromText(intType As AcObjectType, _
                                strName As String, _
                                strFile As String, _
                                Optional cDbObjectClass As IDbComponent = Nothing)

    Dim strTempFile As String
    Dim strPrintSettingsFile As String
    Dim strSourceFile As String
    Dim blnConvert As Boolean
    Dim dFile As Dictionary
    
    ' The path to the source file may change if we add print settings.
    strSourceFile = strFile
    
    ' Add DevMode structures back into forms/reports
    Select Case intType
        Case acForm, acReport
            'Insert print settings (if needed)
            If Not (cDbObjectClass Is Nothing) Then
                With New clsDevMode
                    ' Manually build the print settings file path since we don't have
                    ' a database object we can use with the clsDevMode.GetPrintSettingsFileName
                    strPrintSettingsFile = cDbObjectClass.BaseFolder & GetSafeFileName(strName) & ".json"
                    Set dFile = ReadJsonFile(strPrintSettingsFile)
                    ' Check to ensure dictionary was loaded
                    If Not (dFile Is Nothing) Then
                    ' Insert DevMode structures into file before importing.
                        ' Load default printer settings, then overlay
                        ' settings saved with report.
                        .ApplySettings dFile("Items")
                        ' Insert the settings into a combined export file.
                        strSourceFile = .AddToExportFile(strFile)
                    End If
                End With
            End If
    End Select
    
    ' Check UCS-2-LE requirement for the current database.
    ' (Cached after first call)
    Select Case intType
        Case acForm, acReport, acQuery, acMacro, acTableDataMacro
            blnConvert = RequiresUcs2
        Case acModule
            ' Always convert from UTF-8 in case the file contains
            ' UTF-8 encoded characters but does not have a BOM.
            blnConvert = True
    End Select
    
    ' Only run conversion if needed.
    If blnConvert Then
        ' Perform file conversion, and import from temp file.
        strTempFile = GetTempFile
        If intType = acModule Then
            ' Convert back to ANSI for VBA modules
            ConvertUtf8Ansi strSourceFile, strTempFile, False
        Else
            ' Other objects converted to UCS2
            ConvertUtf8Ucs2 strSourceFile, strTempFile, False
        End If
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strTempFile
        Perf.OperationEnd
        DeleteFile strTempFile, True
    Else
        ' Load UTF-8 file
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strSourceFile
        Perf.OperationEnd
    End If
    
    ' Remove any temporary combined source file
    If strSourceFile <> strFile Then DeleteFile strSourceFile
    
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
    
    Perf.OperationStart "Clear References"
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If Not ref.BuiltIn Then
            strName = ref.Name
            Application.References.Remove ref
            Log.Add "  Removed " & strName, False
        End If
        Set ref = Nothing
    Next intCnt
    Perf.OperationEnd
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetOriginalDbFullPathFromSource
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Determine the original full path of the database, based on the files
'           : in the source folder. (Assumes that options have been loaded)
'---------------------------------------------------------------------------------------
'
Public Function GetOriginalDbFullPathFromSource(strFolder As String) As String
    
    Dim strPath As String
    Dim dContents As Dictionary
    Dim strFile As String
    Dim strExportFolder As String
    Dim lngLevel As Long
    
    strPath = FSO.BuildPath(strFolder, "vbe-project.json")
    If Not FSO.FileExists(strPath) Then
        Log.Error eelCritical, "Unable to find source file: " & strPath, "GetOriginalDbFullPathFromSource"
        GetOriginalDbFullPathFromSource = vbNullString
    Else
        ' Look up file name from VBE project file name
        Set dContents = ReadJsonFile(strPath)
        strFile = dNZ(dContents, "Items\FileName")
        
        ' Convert legacy relative path
        If Left$(strFile, 4) = "rel:" Then strFile = Mid$(strFile, 5)
            
        ' Trim off any tailing slash
        strExportFolder = StripSlash(strFolder)
        
        ' Check export folder settings
        If Options.ExportFolder = vbNullString Then
            ' Default setting, using parent folder of source directory
            GetOriginalDbFullPathFromSource = strFolder & PathSep & ".." & PathSep & strFile
        Else
            ' Check to see if we are using an absolute export path  (\\* or *:*)
            If StartsWith(Options.ExportFolder, PathSep & PathSep) _
                Or (InStr(2, Options.ExportFolder, ":") > 0) Then
                ' We don't save the absolute path in source code, so the user
                ' needs to determine the file location.
                Exit Function
            Else
                ' Calculate how many levels deep to create original path
                lngLevel = UBound(Split(StripSlash(Options.ExportFolder), PathSep))
                If lngLevel < 0 Then lngLevel = 0   ' Handle "\" to export in current folder.
                GetOriginalDbFullPathFromSource = strExportFolder & PathSep & _
                    Repeat(".." & PathSep, lngLevel) & strFile
            End If
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
    FolderHasVcsOptionsFile = FSO.FileExists(FSO.BuildPath(strFolder, "vcs-options.json"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteJsonFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Creates a json file with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'           : Set the file format version only when the dictionary structure changes
'           : with potentially breaking changes for prior versions.
'---------------------------------------------------------------------------------------
'
Public Sub WriteJsonFile(strClassName As String, dItems As Dictionary, strFile As String, strDescription As String, _
    Optional dblExportFormatVersion As Double)
    
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
            If dblExportFormatVersion <> 0 And dNZ(dFile, "Info\Export File Format") <> vbNullString Then
                ' Rewrite file using upgraded format.
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
    dHeader.Add "Class", strClassName
    dHeader.Add "Description", strDescription
    If dblExportFormatVersion <> 0 Then dHeader.Add "Export File Format", dblExportFormatVersion
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
' Date      : 2/23/2021
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent, ParamArray StrExtensions())
    
    Dim oFolder As Folder
    Dim oFile As File
    Dim dBaseNames As Dictionary
    Dim dExtensions As Dictionary
    Dim strBaseName As String
    Dim strFile As String
    Dim varExt As Variant
    Dim strExt As String
    Dim cItem As IDbComponent
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Set up dictionary objects for case-insensitive comparison
    Set dBaseNames = New Dictionary
    dBaseNames.CompareMode = TextCompare
    Set dExtensions = New Dictionary
    dExtensions.CompareMode = TextCompare
    
    ' Cache a list of base source file names for actual database objects
    Perf.OperationStart "Clear Orphaned"
    For Each cItem In cType.GetAllFromDB(False)
        dBaseNames.Add FSO.GetBaseName(cItem.SourceFile), vbNullString
    Next cItem
    
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
                ' Object not found in database. Remove file.
                DeleteFile FSO.BuildPath(oFile.ParentFolder.Path, oFile.Name), True
                Log.Add "  Removing orphaned file: " & strFile, Options.ShowDebug
            End If
        End If
    Next oFile
    
    ' Remove base folder if we don't have any files in it
    If oFolder.Files.Count = 0 Then oFolder.Delete True
    Perf.OperationEnd
    
End Sub