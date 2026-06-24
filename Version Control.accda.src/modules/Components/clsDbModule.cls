VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Components")

Private m_Module As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary
Private m_strSourceFile As String
Private m_colBatchItems As Collection

Private Type udtVbaFileContent
    strContent As String
    blnIsClass As Boolean
    blnPublicCreatable As Boolean
End Type

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strExt As String
    Dim strAlternateFile As String
    Dim strExportFolder As String
    Dim strAnnotation As String
    Dim strSafeName As String

    ' Read the live @Folder annotation once (bypassing any stale cache)
    strAnnotation = GetFolderAnnotation(Me.Parent, blnBypassCache:=True)
    strSafeName = GetSafeFileName(m_Module.Name)

    ' Build export path using the live annotation
    m_strSourceFile = vbNullString
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        m_strSourceFile = IDbComponent_BaseFolder _
            & strAnnotation & strSafeName & GetExtension
    End If

    ' Export as sanitized UTF-8 file
    ExportCodeModule m_Module.Name, Nz2(strAlternatePath, IDbComponent_SourceFile)

    ' Export object metadata (doc properties + hidden attribute) into companion .json.
    ' Produced for both normal and alternate-path (temp) exports so that conflict
    ' detection can compare all files and the conflict dialog can offer per-file diffs.
    ExportObjectMetadata FSO.GetParentFolderName(Nz2(strAlternatePath, IDbComponent_SourceFile)) & PathSep _
        & strSafeName & ".json", "Modules", m_Module.Name, acModule

    ' Clean up stale files only during normal exports (not alternate/temp exports)
    If strAlternatePath = vbNullString Then

        ' Determine the actual export folder (may include @Folder subfolder)
        strExportFolder = FSO.GetParentFolderName(IDbComponent_SourceFile) & PathSep

        ' Remove any file with the same name but alternate extension in the export folder
        strExt = IIf(GetExtension = ".bas", ".cls", ".bas")
        strAlternateFile = strExportFolder & strSafeName & strExt
        If FSO.FileExists(strAlternateFile) Then DeleteFile strAlternateFile

        ' Remove duplicate files at any other location in the base folder tree
        CleanupDuplicateSourceFiles IDbComponent_BaseFolder, strExportFolder, _
            strSafeName, GetExtension, strExt, ".json"

    End If

    ' Update the index with the current VBA hash. (Note, this will not show
    ' changes to the hidden VBE properties that might have been added.)
    Dim cIdx As clsVCSIndexItem
    Set cIdx = VCSIndex.Update(Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), _
        GetFileHash(Nz2(strAlternatePath, IDbComponent_SourceFile)), GetCodeModuleHash(edbModule, m_Module.Name), _
        strMetaHash:=GetMetadataHash("Modules", m_Module.Name, acModule))

    ' Cache the @Folder annotation in the index so subsequent scans
    ' can skip the expensive VBE COM call for unchanged objects.
    cIdx.FolderAnnotation = IIf(Len(strAnnotation) > 0, strAnnotation, FOLDER_ANNOTATION_NONE)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim FunctionName As String
    Dim strName As String
    Dim strTempFile As String
    Dim udtFile As udtVbaFileContent

    FunctionName = ModuleName(Me) & ".Import"

    ' Full build reuses one instance; do not reuse a prior object's cached path.
    m_strSourceFile = vbNullString

    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' Only import files with the correct extension.
    If Not (strFile Like "*.bas" Or strFile Like "*.cls") Then Exit Sub

    ' Parse source file
    strName = GetObjectNameFromFileName(strFile)
    udtFile = ParseSourceFile(strFile, strName)

    ' Write to a new file using system encoding (converting from UTF-8)
    strTempFile = GetTempFile
    WriteFile udtFile.strContent, strTempFile, GetSystemEncoding

    ' Import the source code
    LoadVbeModuleFromFile strTempFile, strName
    If Operation.ErrorLevel >= eelError Then GoTo ImportCleanUp

    ' Make sure the correct project is active before saving.
    Set VBE.ActiveVBProject = CurrentVBProject
    DoCmd.Save acModule, strName
    If CatchAny(eelError, T("Error saving module: {0}", var0:=strName), FunctionName) Then GoTo ImportCleanUp

    DoEvents

    FinalizeOneModule strFile, strName, udtFile.blnPublicCreatable, FunctionName

ImportCleanUp:
    DeleteFile strTempFile

    CatchAny eelError, T("Error importing module: {0}", var0:=strName), FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportFast
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Pass 1 of full-build module import: parse and load via VBE only. The
'           : caller must invoke FinalizeImports once after all files are imported.
'---------------------------------------------------------------------------------------
'
Public Sub ImportFast(strFile As String)

    Dim strName As String
    Dim strTempFile As String
    Dim udtFile As udtVbaFileContent
    Dim dBatch As Dictionary
    Dim FunctionName As String

    FunctionName = ModuleName(Me) & ".ImportFast"

    LogUnhandledErrors FunctionName
    On Error Resume Next

    If m_colBatchItems Is Nothing Then Set m_colBatchItems = New Collection

    If Not (strFile Like "*.bas" Or strFile Like "*.cls") Then Exit Sub

    strName = GetObjectNameFromFileName(strFile)
    udtFile = ParseSourceFile(strFile, strName)

    strTempFile = GetTempFile
    WriteFile udtFile.strContent, strTempFile, GetSystemEncoding

    LoadVbeModuleFromFile strTempFile, strName
    DeleteFile strTempFile

    If Operation.ErrorLevel >= eelError Then Exit Sub

    If Not VbeModuleExists(strName) Then
        Log.Error eelError, _
            T("VBE import did not create module {0}. Check Attribute VB_Name in the source file.", _
                var0:=strName), FunctionName
        Exit Sub
    End If

    Set dBatch = New Dictionary
    dBatch.Add "File", strFile
    dBatch.Add "Name", strName
    dBatch.Add "PublicCreatable", udtFile.blnPublicCreatable
    m_colBatchItems.Add dBatch

    CatchAny eelError, T("Error importing module: {0}", var0:=strName), FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FinalizeImports
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Pass 2 of full-build module import: one project save and Documents
'           : refresh, then resolve, metadata, and index updates per batched file.
'---------------------------------------------------------------------------------------
'
Public Sub FinalizeImports()

    Dim varItem As Variant
    Dim dBatch As Dictionary
    Dim FunctionName As String

    FunctionName = ModuleName(Me) & ".FinalizeImports"

    If m_colBatchItems Is Nothing Then Exit Sub
    If m_colBatchItems.Count = 0 Then
        Set m_colBatchItems = Nothing
        Exit Sub
    End If

    LogUnhandledErrors FunctionName
    On Error Resume Next

    Set VBE.ActiveVBProject = CurrentVBProject

    ' Apply instancing before saving so a single save pass publishes every module.
    For Each varItem In m_colBatchItems
        Set dBatch = varItem
        If dBatch("PublicCreatable") Then
            CurrentVBProject.VBComponents(CStr(dBatch("Name"))).Properties("Instancing") = 5
            Err.Clear
        End If
    Next varItem

    SaveAllBatchModules FunctionName
    If Operation.ErrorLevel >= eelCritical Then GoTo FinalizeCleanUp

    Perf.OperationStart "Refresh Documents"
    SharedDb.Containers("Modules").Documents.Refresh
    Perf.OperationEnd

    For Each varItem In m_colBatchItems
        Set dBatch = varItem
        FinalizeOneModule CStr(dBatch("File")), CStr(dBatch("Name")), CBool(dBatch("PublicCreatable")), _
            FunctionName, blnSkipDocumentsRefresh:=True, blnInstancingAlreadyApplied:=True
        If Operation.ErrorLevel = eelCritical Then Exit For
    Next varItem

FinalizeCleanUp:
    Set m_colBatchItems = Nothing

    CatchAny eelError, T("Error finalizing module imports"), FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveAllBatchModules
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Save each batched module via Access (publishes to AllModules). DoEvents
'           : before and after the loop lets VBE publish imported modules before Access
'           : save and lets AllModules catch up afterward.
'---------------------------------------------------------------------------------------
'
Private Sub SaveAllBatchModules(FunctionName As String)

    Dim varItem As Variant
    Dim dBatch As Dictionary
    Dim strName As String

    Set VBE.ActiveVBProject = CurrentVBProject
    DoEvents

    Perf.OperationStart "Save Modules"
    For Each varItem In m_colBatchItems
        Set dBatch = varItem
        strName = CStr(dBatch("Name"))
        DoCmd.Save acModule, strName
        If CatchAny(eelError, T("Error saving module: {0}", var0:=strName), FunctionName) Then
            Perf.OperationEnd
            Exit Sub
        End If
    Next varItem
    Perf.OperationEnd

    DoEvents

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FinalizeOneModule
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Resolve AllModules, apply metadata, and update the index for one import.
'---------------------------------------------------------------------------------------
'
Private Sub FinalizeOneModule(strFile As String, strName As String, blnPublicCreatable As Boolean, _
    FunctionName As String, Optional blnSkipDocumentsRefresh As Boolean = False, _
    Optional blnInstancingAlreadyApplied As Boolean = False)

    Const cintPublicCreateable As Integer = 5   ' Instancing

    ' Full build reuses one instance; bind this file explicitly for index updates.
    m_strSourceFile = strFile

    If blnPublicCreatable And Not blnInstancingAlreadyApplied Then
        CurrentVBProject.VBComponents(strName).Properties("Instancing") = cintPublicCreateable
        DoCmd.Save acModule, strName
        If CatchAny(eelError, T("Error saving module instancing: {0}", var0:=strName), FunctionName) Then Exit Sub
    End If

    Set m_Module = GetAccessModuleObject(strName, IIf(blnInstancingAlreadyApplied, 5, 3))
    If m_Module Is Nothing Then
        Log.Error eelCritical, _
            T("Imported module not found after import: {0}" & vbNewLine & _
                "Ensure 'Attribute VB_Name = {0}' is found at the start of the file. " & _
                "Filenames (before prefix) must match the 'Attribute VB_Name' value." _
                , var0:=strName), FunctionName
        Exit Sub
    End If

    ImportObjectMetadata SwapExtension(strFile, "json"), "Modules", strName, acModule, _
        blnSkipDocumentsRefresh:=blnSkipDocumentsRefresh

    VCSIndex.Update Me, eatImport, GetFileHash(strFile), _
        GetCodeModuleHash(IDbComponent_ComponentType, strName), _
        strMetaHash:=GetMetadataHash("Modules", strName, acModule)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseSourceFile
' Author    : Adam Waller
' Date      : 7/13/2021
' Purpose   : Parse the source file to build VBE content
'---------------------------------------------------------------------------------------
'
Private Function ParseSourceFile(strFile As String, strName As String) As udtVbaFileContent

    Dim strLines() As String
    Dim lngLine As Long
    Dim cData As clsConcat
    Dim blnIsClass As Boolean
    Dim blnHasHeader As Boolean
    Dim blnCreatable As Boolean
    Dim blnPublic As Boolean

    Perf.OperationStart "Parse VBA Module"

    ' Read file contents into array of lines
    strLines = Split(ReadFile(strFile), vbCrLf)

    ' Loop through first several lines to determine type and header
    For lngLine = 0 To UBound(strLines)
        If strLines(lngLine) = "VERSION 1.0 CLASS" Then
            ' Class with VBE header
            blnIsClass = True
            blnHasHeader = True
        ElseIf StartsWith(strLines(lngLine), "Attribute VB_Name = """) Then
            ' Module with VBE header
            blnHasHeader = True
        ElseIf StartsWith(strLines(lngLine), "Attribute VB_GlobalNameSpace = ") Then
            ' Class with no header
            blnIsClass = True
        ElseIf strLines(lngLine) = "Attribute VB_Creatable = True" Then
            blnCreatable = True
        ElseIf strLines(lngLine) = "Attribute VB_Exposed = True" Then
            blnPublic = True
        End If
        ' Exit after 9 lines
        If lngLine > 8 Then Exit For
    Next lngLine

    ' Use concatenation class to build file contents
    Set cData = New clsConcat
    With cData
        .AppendOnAdd = vbCrLf

        ' Build header, if needed
        If Not blnHasHeader Then
            If blnIsClass Then
                .Add "VERSION 1.0 CLASS"
                .Add "BEGIN"
                .Add "  MultiUse = -1  'True"
                .Add "END"
            End If
            .Add "Attribute VB_Name = """, strName, """"
        End If

        ' Add in file contents
        For lngLine = 0 To UBound(strLines)
            .Add strLines(lngLine)
        Next lngLine

        ' Remove trailing vbCrLf
        .Remove 2
    End With

    ' Return values
    With ParseSourceFile
        .blnIsClass = blnIsClass
        .strContent = cData.GetStr
        .blnPublicCreatable = (blnCreatable And blnPublic)
    End With

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVbeModuleFromFile
' Author    : Adam Waller
' Date      : 7/12/2021
' Purpose   : Load the VBA standard/class module from a file through VBE. (This allows
'           : us to preserve hidden attributes not recognized in then LoadFromText
'           : import of code modules and classes.)
'---------------------------------------------------------------------------------------
'
Private Sub LoadVbeModuleFromFile(strFile As String, strName As String)

    Dim proj As VBProject
    Dim strSource As String

    strSource = ModuleName(Me) & ".LoadVbeModuleFromFile"

    Set proj = CurrentVBProject
    Perf.OperationStart "Import VBE Module"
    With proj.VBComponents

        ' Remove any existing component (In most cases the module will exist)
        LogUnhandledErrors
        On Error Resume Next
        .Remove .Item(strName)
        Err.Clear

        ' Load from the file
        .Import strFile
        If CatchAny(eelError, T("Error importing VBA code for {0}", var0:=strName), strSource) Then GoTo CleanUp

        If Not VbeModuleExists(strName) Then
            Log.Error eelError, _
                T("VBE import did not create module {0}. Check Attribute VB_Name in the source file.", _
                    var0:=strName), strSource
        End If

    End With

CleanUp:
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VbeModuleExists
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Return true when strName exists in the current project's VBComponents.
'---------------------------------------------------------------------------------------
'
Private Function VbeModuleExists(strName As String) As Boolean

    LogUnhandledErrors
    On Error Resume Next
    VbeModuleExists = (CurrentVBProject.VBComponents(strName).Name = strName)
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAccessModuleObject
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Resolve an imported module in CurrentProject.AllModules, retrying briefly
'           : because AllModules can lag behind VBE after a remove/re-import cycle.
'---------------------------------------------------------------------------------------
'
Private Function GetAccessModuleObject(strName As String, Optional intMaxTries As Integer = 3) As AccessObject

    Dim intTry As Integer

    For intTry = 1 To intMaxTries
        LogUnhandledErrors
        On Error Resume Next
        Set GetAccessModuleObject = CurrentProject.AllModules(strName)
        If Not GetAccessModuleObject Is Nothing Then Exit Function
        Err.Clear
        DoEvents
    Next intTry

End Function


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    DeleteObjectIfExists acModule, GetObjectNameFromFileName(strFile)
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    VerifyPath strToFolder
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".json", strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim oMod As AccessObject
    Dim cModule As IDbComponent
    Dim proj As VBProject
    Dim blnNeedAll As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary
        Set proj = CurrentVBProject
        For Each oMod In CurrentProject.AllModules
            Set cModule = New clsDbModule
            Set cModule.DbObject = oMod
            If blnNeedAll Then m_Items(False).Add cModule.SourceFile, cModule
            If blnModifiedOnly Then
                If cModule.IsModified Then m_Items(True).Add cModule.SourceFile, cModule
            End If
        Next oMod
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then
        Set m_FileList = GetFilePathsInFolderRecursive(IDbComponent_BaseFolder, "*.bas", "*.cls")
    End If
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export/import of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean

    ' Fast path: if the VBE project has no unsaved changes and the
    ' VBA project date matches the stored VBAProjectDate, skip the
    ' expensive VBA code hash. (Access saves all modules as a monolithic
    ' project, so all module dates move in lockstep -- we only need to
    ' check one date against a single top-level stored value.)
    If CurrentVBProject.Saved Then
        If VCSIndex.VBAProjectDate <> 0 Then
            If DateTruncToSeconds(CurrentProject.AllModules(0).DateModified) = VCSIndex.VBAProjectDate Then
                GoTo CheckMetadata
            End If
        End If
    End If

    ' VBE is dirty or date changed -- hash the code to detect changes.
    ' NOTE: This will not detect changes to hidden VBE attributes
    IDbComponent_IsModified = VCSIndex.Item(Me).OtherHash <> GetCodeModuleHash(IDbComponent_ComponentType, m_Module.Name)

CheckMetadata:
    ' Check metadata (Description/Hidden) which don't update DateModified
    If Not IDbComponent_IsModified Then
        IDbComponent_IsModified = VCSIndex.Item(Me).MetaHash <> _
            GetMetadataHash("Modules", m_Module.Name, acModule)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    If m_Module Is Nothing Then Exit Function
    IDbComponent_DateModified = CurrentProject.AllModules(m_Module.Name).DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Modules"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "modules" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "bas"
    IDbComponent_FileExtensions.Add "cls"
    IDbComponent_FileExtensions.Add "json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Module Is Nothing Then Exit Property
    IDbComponent_Name = m_Module.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If Len(m_strSourceFile) > 0 Then
        IDbComponent_SourceFile = m_strSourceFile
        Exit Property
    End If
    If m_Module Is Nothing Then Exit Property
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        m_strSourceFile = IDbComponent_BaseFolder _
            & GetFolderAnnotation(Me.Parent) _
            & GetSafeFileName(m_Module.Name) & GetExtension
    Else
        m_strSourceFile = IDbComponent_BaseFolder _
            & GetSafeFileName(m_Module.Name) & GetExtension
    End If
    IDbComponent_SourceFile = m_strSourceFile
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetExtension
' Author    : Adam Waller
' Date      : 6/4/2021
' Purpose   : Return the extension (".cls" or ".bas") based on the component type.
'---------------------------------------------------------------------------------------
'
Private Function GetExtension() As String
    If CurrentVBProject.VBComponents(m_Module.Name).Type = vbext_ct_StdModule Then
        GetExtension = ".bas"
    Else
        GetExtension = ".cls"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    Static lngCount As Long
    If lngCount = 0 Then lngCount = CurrentProject.AllModules.Count - 1
    IDbComponent_QuickCount = lngCount + 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbModule
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Module
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Module = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
