VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbVbeProject"
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
Private m_Items(True To False) As Dictionary
Private m_dItems As Dictionary

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
    Dim strContent As String
    strContent = GetSource
    WriteFile strContent, Nz2(strAlternatePath, IDbComponent_SourceFile)
    If strAlternatePath = vbNullString Then VCSIndex.Update Me, eatExport, GetStringHash(strContent, True)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dProject As Dictionary

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub
    Set dProject = ReadJsonFile(strFile)
    Set m_Project = GetVBProjectForCurrentDB
    
    ' Update project properties
    With m_Project
        .Name = dNZ(dProject, "Items\Name")
        .Description = dNZ(dProject, "Items\Description")
        
        ' Setting the HelpContextId can throw random automation errors.
        SafeSetProperty m_Project, "HelpContextId", ValidHelpContextId(dNZ(dProject, "Items\HelpContextId"))
        SafeSetProperty m_Project, "HelpFile", ValidHelpFile(dNZ(dProject, "Items\HelpFile"))
        
        ' This property is not exposed through the VBProject object model
        Application.SetOption "Conditional Compilation Arguments", dNZ(dProject, "Items\ConditionalCompilationArguments")
        
        ' // Read-only properties
        '.FileName = dNZ(dProject, "Items\FileName")
        '.Mode = dNZ(dProject, "Items\Mode")
        '.Protection = dNZ(dProject, "Items\Protection")
        '.Type = dNZ(dProject, "Items\Type")
    End With
    
    CatchAny eelError, "Importing VBE Project", ModuleName & ".Import"
    
    ' Save to index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary(False))

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SafeSetProperty
' Author    : Adam Waller
' Date      : 3/26/2021
' Purpose   : For some reason the help properties can sometimes throw strange runtime
'           : errors when setting them. This function handles the extra error handling
'           : involved in setting and verifying these properties.
'---------------------------------------------------------------------------------------
'
Private Sub SafeSetProperty(cProj As VBProject, strProperty As String, varValue As Variant)

    Dim varCurrent As Variant
    
    ' Get current property value
    varCurrent = CallByName(cProj, strProperty, VbGet)
    
    ' No need to set if the current value already matches
    If varValue = varCurrent Then Exit Sub
    
    ' Switch to on error resume next after checking for current errors
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    
    ' Attempt to set the property
    CallByName cProj, strProperty, VbLet, varValue
    
    ' Read the value after setting it
    varCurrent = CallByName(cProj, strProperty, VbGet)
    
    ' Verify that the property was set correctly
    If varCurrent <> varValue Then
        ' We might have thrown an actual error.
        If Not CatchAny(eelError, "Failed to set " & strProperty & " to '" & _
            CStr(varValue) & "'", ModuleName & ".SafeSetProperty") Then
            ' No error, but property not set correctly.
            Log.Error eelError, "Failed to set " & strProperty & ". Set value to '" & _
                varValue & "' but afterwards it returned '" & varCurrent & "'.", _
                ModuleName & ".SafeSetProperty"
        End If
    Else
        ' Clear any errors that may have been thrown, even if the change was successful.
        If Err Then Err.Clear
    End If
    
    ' Log any uncaught errors
    CatchAny eelError, "Setting Property '" & strProperty & "' to value '" & CStr(varValue) & "'", ModuleName & ".SafeSetProperty"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ValidHelpContextId
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Don't attempt to set the help context id to anything other than a number.
'---------------------------------------------------------------------------------------
'
Private Function ValidHelpContextId(strHelpID As String) As Long
    If strHelpID = vbNullString Then
        ValidHelpContextId = 0
    ElseIf Not IsNumeric(strHelpID) Then
        Log.Error eelWarning, "HelpContextID should be a number. " & _
            "Found '" & strHelpID & "' instead.", ModuleName & ".ValidHelpContextId"
        ValidHelpContextId = 0
    Else
        ValidHelpContextId = CLng(strHelpID)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidHelpFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Get help file path saved as a relative path.
'---------------------------------------------------------------------------------------
'
Private Function ValidHelpFile(strHelpFile As String) As String

    Dim strValid As String
    
    If strHelpFile <> vbNullString Then
    
        ' Check for a Rubber Duck Identifier value
        If InStr(1, strHelpFile, ".") = 0 And IsNumeric(strHelpFile) Then
            If Options.PreserveRubberDuckID Then
                ' Allow user to save this identifier in the exported source, if they really want to.
                ' See issue #197 for more details on this.
                Log.Add "RubberDuck Identifier " & strHelpFile & " found in VBE Project HelpFile field. " & _
                    "If you don't want to save this to source, " & _
                    "set PreserveRubberDuckID = False in the options file.", False
                    
                strValid = strHelpFile
            End If
            
        Else
            ' Might actually be a help file name/path
        
            ' Build out any relative path
            strValid = GetPathFromRelative(strHelpFile)
        
            ' Make sure this is a valid help file
            If strValid Like "*.hlp" Or strValid Like "*.chm" Then
                If Not FSO.FileExists(strValid) Then
                    Log.Error eelWarning, "Help file not found: " & strValid, ModuleName & ".ValidHelpFile"
                End If
            Else
                ' Does not appear to be a help file extension
                Log.Error eelWarning, "'" & strValid & "' is not a valid help file name. " & _
                    "(Expecting *.hlp or *.chm)", ModuleName & ".ValidHelpFile"
                strValid = vbNullString
            End If
        
        End If
    End If

    ' Return validated help file string
    ValidHelpFile = strValid
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "VBE Project")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Return a dictionary object of project properties.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary(Optional blnUseCache As Boolean = True) As Dictionary

    ' Check cache parameter
    If blnUseCache And Not m_dItems Is Nothing Then
        ' Return cached dictionary
        Set GetDictionary = m_dItems
        Exit Function
    End If

    ' Make sure we have a reference to the VB project
    If m_Project Is Nothing Then Set m_Project = GetVBProjectForCurrentDB
    
    ' Read project properties
    Set GetDictionary = New Dictionary
    With GetDictionary
        .Add "Name", m_Project.Name
        .Add "Description", m_Project.Description
        .Add "FileName", FSO.GetFileName(m_Project.FileName)
        .Add "HelpFile", ValidHelpFile(m_Project.HelpFile)
        .Add "HelpContextId", ValidHelpContextId(m_Project.HelpContextId)
        .Add "ConditionalCompilationArguments", Application.GetOption("Conditional Compilation Arguments")
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
    ' Import will replace existing values
    IDbComponent_Import strFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary
    
    Dim cProj As IDbComponent
    Dim blnAdd As Boolean
    
    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        ' Load class details
        Set m_Project = GetVBProjectForCurrentDB
        Set m_Items(blnModifiedOnly) = New Dictionary
        Set cProj = New clsDbVbeProject
        Set cProj.DbObject = m_Project
        blnAdd = True
        If blnModifiedOnly Then blnAdd = cProj.IsModified
        If blnAdd Then m_Items(blnModifiedOnly).Add cProj.SourceFile, cProj
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
    Set IDbComponent_GetFileList = New Dictionary
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_GetFileList.Add IDbComponent_SourceFile, vbNullString
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
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
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
    If m_Project Is Nothing Then Exit Property
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
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
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



