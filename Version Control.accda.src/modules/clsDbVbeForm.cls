VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbVbeForm"
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

Private m_Form As VBIDE.VBComponent
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
' Date      : 5/17/2023
' Purpose   : Export the individual database component (table, form, query, etc...)
'           : NOTE: This has been expanded to include an export (for reference only) of
'           : a serialized representation of the form properties, controls, and layout.
'           : ALSO NOTE: This behaves a little differently than other objects. Because
'           : the binary *.frx file changes EVERY time you export it, we should avoid
'           : actually exporting it if nothing has changed in the serialized file.
'           : It is *technically* possible to make a change on the form that would not
'           : be included in the serialized output (such as a property change on an
'           : active-x control) but this is extremely unlikely in contrast to the noise
'           : produced by the binary file showing changes on every full export.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strContent As String
    Dim strBasePath As String
    Dim strJsonPath  As String
    Dim strFormPath As String
    Dim strBinaryPath As String
    Dim blnUnchanged As Boolean

    ' Get the JSON representation of the current database object
    strContent = GetSource

    ' Build out all three normal source file paths.
    strBasePath = IDbComponent_BaseFolder & GetSafeFileName(m_Form.Name)
    strJsonPath = strBasePath & ".json"
    strBinaryPath = strBasePath & ".frx"
    strFormPath = strBasePath & ".frm"

    ' Before exporting the VBE file, let's see if the source files exist, just in case
    ' the database object already matches the source file.
    If FSO.FileExists(strJsonPath) And FSO.FileExists(strBinaryPath) And FSO.FileExists(strFormPath) Then
        ' If so, compare the hash of the json file to the database object.
        blnUnchanged = (GetStringHash(strContent, True) = GetFileHash(strJsonPath))
    End If

    ' This is the serialized output file in JSON format to track changes in version control
    ' Always ouput this file during an export
    WriteFile strContent, strJsonPath

    ' Only export the VBE object (including binary source) if the form has changed
    ' or if the VBE files are missing in the original source location
    If Not blnUnchanged Then
        If Len(strAlternatePath) Then
            ' Save to alternate path
            m_Form.Export SwapExtension(strAlternatePath, "frm")
        Else
            ' This is the binary export file used when building from source
            m_Form.Export strFormPath
        End If
    End If

    ' Update index with a hash of the serialized content. (Since the binary content changes frequently)
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), GetStringHash(strContent, True)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/30/2021
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim proj As VBProject
    Dim strTestFile As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct (primary) extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Make sure the other two companion files also exist
    strTestFile = SwapExtension(strFile, "frm")
    If Not FSO.FileExists(strTestFile) Then
        Log.Error eelError, "VBE Form definition file not found: " & strTestFile, ModuleName(Me) & ".Import"
        Exit Sub
    Else
        ' Check binary file
        strTestFile = SwapExtension(strFile, "frx")
        If Not FSO.FileExists(strTestFile) Then
            Log.Error eelError, "VBE Form binary file not found: " & strTestFile, ModuleName(Me) & ".Import"
            Exit Sub
        End If
    End If

    ' With the files verified, we can move forward with the actual import
    Set proj = CurrentVBProject
    With proj.VBComponents

        ' Import the VBE source file
        .Import SwapExtension(strFile, "frm")

        ' Set reference to form after import
        Set m_Form = .Item(GetObjectNameFromFileName(strFile))

        ' Check for extra blank line that may get added during import (VBE bug?)
        With m_Form
            If Not .CodeModule Is Nothing Then
                If .CodeModule.Lines(1, 1) = vbNullString Then
                    ' Remove blank line added during import (Known VBA Bug)
                    .CodeModule.DeleteLines 1
                    ' Note, this change is not saved at this time. It will be saved during
                    ' the next compile/save operation.
                    Log.Add "Removed blank line from the top of the code module for " & .Name, False
                End If
            End If
        End With
    End With

    ' Update index (based on serialized representation)
    VCSIndex.Update Me, eatImport, GetStringHash(GetSource, True)

    ' Log any errors while importing the VBE form
    CatchAny eelError, "Error importing " & strFile, ModuleName(Me) & ".Import"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, replacing or removing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    DeleteVbeFormIfExists GetObjectNameFromFileName(strFile)
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DeleteVbeFormIfExists
' Author    : Adam Waller
' Date      : 2/5/2022
' Purpose   : Remove any existing VBE form
'---------------------------------------------------------------------------------------
'
Private Sub DeleteVbeFormIfExists(strName)
    LogUnhandledErrors
    On Error Resume Next
    With CurrentVBProject.VBComponents
        .Remove .Item(strName)
    End With
    Catch 9 ' Subscript out of range (name not found)
    CatchAny eelError, "Deleting VBE Form " & strName
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
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".frm", strToFolder
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".frx", strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim frm As VBComponent
    Dim cForm As IDbComponent
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        For Each frm In CurrentVBProject.VBComponents
            If frm.Type = vbext_ct_MSForm Then
                Set cForm = New clsDbVbeForm
                Set cForm.DbObject = frm
                blnAdd = True
                If blnModifiedOnly Then blnAdd = cForm.IsModified
                If blnAdd Then m_Items(blnModifiedOnly).Add cForm.SourceFile, cForm
            End If
        Next frm
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "Serialized VBE Form")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Return a dictionary of the serialized form structure
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary() As Dictionary
    Set GetDictionary = SerializeMSForm(m_Form)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
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
    ClearOrphanedSourceFiles Me, "frm", "frx", "json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedDatabaseObjects
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Remove database objects that are not represented by existing source files.
'---------------------------------------------------------------------------------------
'
Public Sub IDbComponent_ClearOrphanedDatabaseObjects()
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
    ' Compare the hash of the serialized JSON output, not the ever-changing binary source file.
    ' NOTE: This will not detect changes in unserialized objects such as ActiveX controls.
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceIsModified
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Returns true if the source file appears to be modified since the last
'           : import or export. (When compared to the index)
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceIsModified(strFile As String) As Boolean

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
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "VBE Forms"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "vbeforms" & PathSep
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
    IDbComponent_FileExtensions.Add "json"
    IDbComponent_FileExtensions.Add "frx"
    IDbComponent_FileExtensions.Add "frm"
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
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Form.Name) & ".json"
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
    ' No simple way to pull a quick count of these
    IDbComponent_QuickCount = 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbVbeForm
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
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
