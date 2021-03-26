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

Private Const ModuleName As String = "clsDbVbeProject"

Private m_Project As VBIDE.VBProject
Private m_AllItems As Collection

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
Private Sub IDbComponent_Export()

    Dim dProject As Dictionary
    
    Set dProject = GetDictionary
    
    ' Save in JSON format.
    WriteJsonFile TypeName(Me), dProject, IDbComponent_SourceFile, "VBE Project"
    
    ' Save to index
    VCSIndex.Update Me, eatExport, GetDictionaryHash(dProject)
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ValidHelpContextId
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Don't attempt to set the help context id to anything other than a number.
'---------------------------------------------------------------------------------------
'
Private Function ValidHelpContextId(ByVal helpID As String) As Long
    If helpID = vbNullString Then
        ValidHelpContextId = 0
    ElseIf Not IsNumeric(helpID) Then
        Log.Error eelWarning, "HelpContextID should be a number. " & _
            "Found '" & helpID & "' instead.", ModuleName & ".ValidHelpContextId"
        ValidHelpContextId = 0
    Else
        ValidHelpContextId = CLng(helpID)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidHelpFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Get help file path saved as a relative path.
'---------------------------------------------------------------------------------------
'
Private Function ValidHelpFile(ByVal helpFile As String) As String
    ValidHelpFile = GetPathFromRelative(helpFile)
    If ValidHelpFile <> vbNullString Then
        ' Make sure this is a valid help file
        If ValidHelpFile Like "*.hlp" Or ValidHelpFile Like "*.chm" Then
            If Not FSO.FileExists(ValidHelpFile) Then
                Log.Error eelWarning, "Help file not found: " & ValidHelpFile, ModuleName & ".ValidHelpFile"
            End If
        Else
            ' Does not appear to be a help file extension
            Log.Error eelWarning, "'" & ValidHelpFile & "' is not a valid help file name. " & _
                "(Expecting *.hlp or *.chm)", ModuleName & ".ValidHelpFile"
            ValidHelpFile = vbNullString
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Update project properties
    Dim dProject As Dictionary
    Set dProject = ReadJsonFile(strFile)
    Set m_Project = GetVBProjectForCurrentDB
    With m_Project
        .Name = dNZ(dProject, "Items\Name")
        .Description = dNZ(dProject, "Items\Description")
        
        ' Setting the HelpContextId can throw random automation errors.
        ' The setting does change despite the error.
        If DebugMode Then On Error Resume Next Else On Error Resume Next
        Dim newHelpID As Long
        newHelpID = ValidHelpContextId(dNZ(dProject, "Items\HelpContextId"))
        .HelpContextId = newHelpID
        If DebugMode Then On Error GoTo 0 Else On Error Resume Next
        
        ' If we failed to set the ID then it was a real error, throw it
        If CStr(.HelpContextId) <> newHelpID Then CatchAny eelError, "Failed to set help context ID"
        
        If DebugMode Then On Error Resume Next Else On Error Resume Next
        Dim newHelpFile As String
        newHelpFile = ValidHelpFile(dNZ(dProject, "Items\HelpFile"))
        .helpFile = newHelpFile
        If DebugMode Then On Error GoTo 0 Else On Error Resume Next
        If .helpFile <> newHelpFile Then CatchAny eelError, "Failed to set help file"
        
        ' // Read-only properties
        '.FileName = dNZ(dProject, "Items\FileName")
        '.Mode = dNZ(dProject, "Items\Mode")
        '.Protection = dNZ(dProject, "Items\Protection")
        '.Type = dNZ(dProject, "Items\Type")
    End With
    
    CatchAny eelError, "Importing VBE Project", ModuleName & ".Import"
    
    ' Save to index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Return a dictionary object of project properties.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary() As Dictionary

    ' Make sure we have a reference to the VB project
    If m_Project Is Nothing Then Set m_Project = GetVBProjectForCurrentDB
    
    ' Read project properties
    Set GetDictionary = New Dictionary
    With GetDictionary
        .Add "Name", m_Project.Name
        .Add "Description", m_Project.Description
        .Add "FileName", GetRelativePath(m_Project.FileName)
        .Add "HelpFile", ValidHelpFile(m_Project.helpFile)
        .Add "HelpContextId", ValidHelpContextId(m_Project.HelpContextId)
        .Add "Mode", m_Project.Mode
        .Add "Protection", m_Project.Protection
        .Add "Type", m_Project.Type
    End With
    
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
    IDbComponent_Import strFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection
    
    Dim cProj As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        ' Load class details
        Set m_Project = GetVBProjectForCurrentDB
        Set m_AllItems = New Collection
        Set cProj = New clsDbVbeProject
        Set cProj.DbObject = m_Project
        m_AllItems.Add cProj, m_Project.Name
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    Set IDbComponent_GetFileList = New Collection
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_GetFileList.Add IDbComponent_SourceFile
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
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
    IDbComponent_Category = "VB Project"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Project.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "vbe-project.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbVbeProject
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
    Set IDbComponent_DbObject = m_Project
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Project = RHS
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
    IDbComponent_SingleFile = True
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Helps us know whether we have already counted the objects.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    'm_Count = -1
End Sub


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