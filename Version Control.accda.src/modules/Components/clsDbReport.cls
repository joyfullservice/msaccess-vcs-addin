VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbReport"
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

Private m_Report As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary
Private m_strSourceFile As String

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

    Dim strHash As String
    Dim strExportFolder As String
    Dim strSafeName As String
    Dim strAnnotation As String

    ' Read the live @Folder annotation once (bypassing any stale cache)
    strAnnotation = GetFolderAnnotation(Me.Parent, blnBypassCache:=True)
    strSafeName = GetSafeFileName(m_Report.Name)

    ' Build export path using the live annotation
    m_strSourceFile = vbNullString
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        m_strSourceFile = IDbComponent_BaseFolder _
            & strAnnotation & strSafeName & ".report"
    End If

    strHash = SaveComponentAsText(acReport, m_Report.Name, Nz2(strAlternatePath, IDbComponent_SourceFile), Me)

    ' Export object metadata (doc properties + hidden attribute) into companion .json
    If strAlternatePath = vbNullString Then
        ExportObjectMetadata SwapExtension(IDbComponent_SourceFile, "json"), "Reports", m_Report.Name, Me
    End If

    ' Generate layout SVG if enabled
    If strAlternatePath = vbNullString Then
        TryExportLayoutSvg IDbComponent_SourceFile, acReport
    End If

    ' Remove duplicate files at any other location in the base folder tree
    strExportFolder = FSO.GetParentFolderName(IDbComponent_SourceFile) & PathSep
    CleanupDuplicateSourceFiles IDbComponent_BaseFolder, strExportFolder, _
        strSafeName, ".report", ".json", ".cls", ".svg"

    Dim cIdx As clsVCSIndexItem
    Set cIdx = VCSIndex.Update(Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), _
        strHash, GetCodeModuleHash(IDbComponent_ComponentType, m_Report.Name), _
        strMetaHash:=GetMetadataHash("Reports", m_Report.Name, acReport))

    ' Cache the @Folder annotation in the index
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

    Dim strName As String
    Dim blnImportCheck As Boolean

    m_strSourceFile = vbNullString

    ' Only import files with the correct extension.
    If Not (strFile Like "*.report" Or strFile Like "*.bas") Then Exit Sub

    strName = GetObjectNameFromFileName(strFile)
    blnImportCheck = LoadComponentFromText(acReport, strName, strFile)
    If Not blnImportCheck Or Operation.ErrorLevel = eelCritical Then Exit Sub
    Set m_Report = CurrentProject.AllReports(strName)

    ' Apply object metadata (doc properties + hidden attribute) from companion .json
    ImportObjectMetadata SwapExtension(strFile, "json"), "Reports", strName, acReport

    VCSIndex.Update Me, eatImport, GetFileHash(strFile), _
        GetCodeModuleHash(IDbComponent_ComponentType, strName), _
        strMetaHash:=GetMetadataHash("Reports", strName, acReport)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    DeleteObjectIfExists acReport, GetObjectNameFromFileName(strFile)
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
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".cls", strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim rpt As AccessObject
    Dim cReport As IDbComponent
    Dim blnNeedAll As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary
        For Each rpt In CurrentProject.AllReports
            Set cReport = New clsDbReport
            Set cReport.DbObject = rpt
            If blnNeedAll Then m_Items(False).Add cReport.SourceFile, cReport
            If blnModifiedOnly Then
                If cReport.IsModified Then m_Items(True).Add cReport.SourceFile, cReport
            End If
        Next rpt
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
        Set m_FileList = GetFilePathsInFolderRecursive(IDbComponent_BaseFolder, "*.report", "*.bas")
    End If
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean

    ' Item is considered modified unless proven otherwise.
    IDbComponent_IsModified = True

    ' Check the modified date first.
    If DateTruncToSeconds(m_Report.DateModified) = VCSIndex.Item(Me).ObjectDate Then

        ' Layout unchanged. Only hash code if the VBE project is dirty.
        If Not CurrentVBProject.Saved Then
            IDbComponent_IsModified = VCSIndex.Item(Me).OtherHash <> GetCodeModuleHash(IDbComponent_ComponentType, m_Report.Name)
        Else
            IDbComponent_IsModified = False
        End If
    End If

    ' Check metadata (Description/Hidden) which don't update DateModified
    If Not IDbComponent_IsModified Then
        IDbComponent_IsModified = VCSIndex.Item(Me).MetaHash <> _
            GetMetadataHash("Reports", m_Report.Name, acReport)
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
    If m_Report Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Report.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Reports"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "reports" & PathSep
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
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        IDbComponent_FileExtensions.Add "report"
    Else
        IDbComponent_FileExtensions.Add "bas"
    End If
    IDbComponent_FileExtensions.Add "cls"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Report Is Nothing Then Exit Property
    IDbComponent_Name = m_Report.Name
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
    If m_Report Is Nothing Then Exit Property
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        m_strSourceFile = IDbComponent_BaseFolder _
            & GetFolderAnnotation(Me.Parent) _
            & GetSafeFileName(m_Report.Name) & ".report"
    Else
        m_strSourceFile = IDbComponent_BaseFolder _
            & GetSafeFileName(m_Report.Name) & ".bas"
    End If
    IDbComponent_SourceFile = m_strSourceFile
End Property


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
    If lngCount = 0 Then lngCount = CurrentProject.AllReports.Count - 1
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
    IDbComponent_ComponentType = edbReport
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Report
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Report = RHS
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
