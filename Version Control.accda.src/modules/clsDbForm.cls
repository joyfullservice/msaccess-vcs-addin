VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbForm"
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

Private m_Form As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

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
    strHash = SaveComponentAsText(acForm, m_Form.Name, Nz2(strAlternatePath, IDbComponent_SourceFile), Me)
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), _
        strHash, GetCodeModuleHash(IDbComponent_ComponentType, m_Form.Name)
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
    Dim blnStage As Boolean

    ' Only import files with the correct extension.
    If Not strFile Like "*.bas" Then Exit Sub

    strName = GetObjectNameFromFileName(strFile)

    ' We can't import a form with the same name as a running form, even though
    ' it is technically in a different database.
    If strName = "frmVCSMain" Then
        If IsLoaded(acForm, "frmVCSMain") Then
            StageMainForm
            blnStage = True
        End If
    End If

    ' Load the form from the source file
    LoadComponentFromText acForm, strName, strFile, Me
    Set m_Form = CurrentProject.AllForms(strName)
    VCSIndex.Update Me, eatImport, GetCodeModuleHash(IDbComponent_ComponentType, strName)

    ' Restore main form if staged
    If blnStage Then RestoreMainForm

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

    Dim strName As String

    ' Get object name
    strName = GetObjectNameFromFileName(strFile)

    ' Don't attempt to merge forms with same names as add-in forms.
    If ObjectExists(acForm, strName, True) Then
        Log.Error eelError, "Merging not supported for add-in forms. Use full build instead.", _
            ModuleName(Me) & ".Merge"
        Exit Sub
    End If

    ' Import the form if the source file exists.
    If FSO.FileExists(strFile) Then
        ' Importing the form will replace the existing object
        IDbComponent_Import strFile
    Else
        DeleteObjectIfExists acForm, strName
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

    Dim frm As AccessObject
    Dim cForm As IDbComponent
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        For Each frm In CurrentProject.AllForms
            Set cForm = New clsDbForm
            Set cForm.DbObject = frm
            blnAdd = True
            If blnModifiedOnly Then blnAdd = cForm.IsModified
            If blnAdd Then m_Items(blnModifiedOnly).Add cForm.SourceFile, cForm
        Next frm
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
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.bas")
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    If Not Options.SavePrintVars Then ClearFilesByExtension IDbComponent_BaseFolder, "json"
    ClearOrphanedSourceFiles Me, "bas", "json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedDatabaseObjects
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Remove database objects that are not represented by existing source files.
'---------------------------------------------------------------------------------------
'
Public Sub IDbComponent_ClearOrphanedDatabaseObjects()
    RemoveOrphanedDatabaseObjects Me
End Sub


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
    ' (This may not reflect some code changes)
    If m_Form.DateModified <= VCSIndex.Item(Me).ExportDate Then

        ' Date is okay, check hash
        IDbComponent_IsModified = VCSIndex.Item(Me).OtherHash <> GetCodeModuleHash(IDbComponent_ComponentType, m_Form.Name)
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
    If m_Form Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Form.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = GetLastModifiedDate(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Forms"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "forms" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Form Is Nothing Then Exit Property
    IDbComponent_Name = m_Form.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Form Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Form.Name) & ".bas"
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
    If lngCount = 0 Then lngCount = CurrentProject.AllForms.Count - 1
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
    IDbComponent_ComponentType = edbForm
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed for forms.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Form
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Form = RHS
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
