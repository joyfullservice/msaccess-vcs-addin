Attribute VB_Name = "modVCSUtility"
'---------------------------------------------------------------------------------------
' Module    : modVCSUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Utility functions specific to the VCS project but not publicly exposed.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName = "modVCSUtility"


'---------------------------------------------------------------------------------------
' Procedure : GetAllContainers
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return a collection of all containers.
'           : NOTE: The order doesn't matter for export, but is VERY important
'           : when building the project from source.
'---------------------------------------------------------------------------------------
'
Public Function GetContainers(Optional intFilter As eContainerFilter = ecfAllObjects) As Collection
    
    Dim blnADP As Boolean
    Dim blnMDB As Boolean
    
    blnADP = (CurrentProject.ProjectType = acADP)
    blnMDB = (CurrentProject.ProjectType = acMDB)
    
    Set GetContainers = New Collection
    With GetContainers
        Select Case intFilter
            
            ' Primary case for processing all objects
            Case ecfAllObjects
            
                ' Shared objects in both MDB and ADP formats
                .Add New clsDbProject
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
                    .Add New clsDbProperty
                    .Add New clsDbSharedImage
                    .Add New clsDbTheme
                    .Add New clsDbImexSpec
                    .Add New clsDbTableDef
                    .Add New clsDbQuery
                End If
                ' Additional objects to import after ADP/MDB specific items
                .Add New clsDbForm
                .Add New clsDbMacro
                .Add New clsDbReport
                .Add New clsDbTableData
                .Add New clsDbModule
                If blnMDB Then
                    .Add New clsDbTableDataMacro
                    .Add New clsDbRelation
                    .Add New clsDbDocument
                    .Add New clsDbNavPaneGroup
                    .Add New clsDbHiddenAttribute
                End If
            
            ' Process only items that may contain VBA code
            Case ecfVBAItems
            
                .Add New clsDbForm
                .Add New clsDbReport
                .Add New clsDbModule
        
        End Select
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
'           : Returns a hash of the file content (if applicable) to track changes.
'---------------------------------------------------------------------------------------
'
Public Function SaveComponentAsText(intType As AcObjectType, _
                                strName As String, _
                                strFile As String, _
                                Optional cDbObjectClass As IDbComponent = Nothing) As String
    
    Dim strTempFile As String
    Dim strPrintSettingsFile As String
    Dim strContent As String
    Dim strHash As String
    
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
                        WriteFile BuildJsonFile(TypeName(cDbObjectClass), .GetDictionary, _
                          strName & " Print Settings"), strPrintSettingsFile
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
            strHash = SanitizeFile(strTempFile, True)
            FSO.MoveFile strTempFile, strFile
    
        Case acQuery, acMacro
            ' Sanitizing converts to UTF-8
            If FSO.FileExists(strFile) Then DeleteFile strFile
            strHash = SanitizeFile(strTempFile, True)
            FSO.MoveFile strTempFile, strFile
            
        ' Case acModule - Use VBE export instead.
        
        Case acTableDataMacro
            ' Table data macros are stored in XML format
            If FSO.FileExists(strFile) Then strHash = SanitizeXML(strFile, True)
            
        Case Else
            ' Handle UCS conversion if needed
            ConvertUcs2Utf8 strTempFile, strFile
        
    End Select
    
    ' Normal exit
    On Error GoTo 0
    
    ' Return content hash
    SaveComponentAsText = strHash
    Exit Function
    
ErrHandler:
    If Err.Number = 2950 And intType = acTableDataMacro Then
        ' This table apparently didn't have a Table Data Macro.
        Exit Function
    Else
        ' Some other error.
        Err.Raise Err.Number
    End If
    
End Function


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
                ' Look for saved build path
                Set dContents = ReadJsonFile(FSO.BuildPath(strFolder, "proj-properties.json"))
                strPath = dNZ(dContents, "Items\VCS Build Path")
                If strPath <> vbNullString Then
                    GetOriginalDbFullPathFromSource = strPath & PathSep & strFile
                Else
                    ' Unable to determine the original file location.
                    Exit Function
                End If
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
' Procedure : BuildJsonFile
' Author    : Adam Waller
' Date      : 2/5/2022
' Purpose   : Creates json file content with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'           : Set the file format version only when the dictionary structure changes
'           : with potentially breaking changes for prior versions.
'---------------------------------------------------------------------------------------
'
Public Function BuildJsonFile(strClassName As String, dItems As Dictionary, strDescription As String, _
    Optional dblExportFormatVersion As Double) As String
    
    Dim dContents As Dictionary
    Dim dHeader As Dictionary
    
    Set dContents = New Dictionary
    Set dHeader = New Dictionary
    
    ' Build dictionary structure
    dHeader.Add "Class", strClassName
    dHeader.Add "Description", strDescription
    If dblExportFormatVersion <> 0 Then dHeader.Add "Export File Format", dblExportFormatVersion
    dContents.Add "Info", dHeader
    dContents.Add "Items", dItems
    
    ' Return assembled content in Json format
    BuildJsonFile = ConvertToJson(dContents, JSON_WHITESPACE)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CompileAndSaveAllModules
' Author    : Adam Waller
' Date      : 7/10/2021
' Purpose   : Compile and save the modules in the current database
'---------------------------------------------------------------------------------------
'
Public Sub CompileAndSaveAllModules()
    Perf.OperationStart "Compile/Save Modules"
    ' Make sure we are running this in the CurrentDB, not the CodeDB
    Set VBE.ActiveVBProject = GetVBProjectForCurrentDB
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    DoEvents
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    For Each oProj In VBE.VBProjects
        If StrComp(oProj.FileName, GetAddInFileName, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Public Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    Application.Run GetAddInFileName & "!DummyFunction"
End Sub


