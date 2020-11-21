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

Private m_Property As AccessObjectProperty
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
    
    Dim prp As AccessObjectProperty
    Dim dCollection As Dictionary
    Dim strPath As String
    Dim varValue As Variant
    
    Set dCollection = New Dictionary
    
    ' Loop through all properties
    For Each prp In CurrentProject.Properties
        Select Case prp.Name
            Case "Last VCS Export", "Last VCS Version"
                ' Reduce noise by ignoring these values.
                ' (We already have this information in the header.)
            Case "AppIcon"
                ' Try to use a relative path
                strPath = GetRelativePath(CStr(prp.Value))
                If Left(strPath, 4) = "rel:" Then
                    varValue = strPath
                Else
                    ' The full path may contain sensitive info. Encrypt the path but not the file name.
                    varValue = SecurePath(CStr(varValue))
                End If
                ' ADP projects may have this property
                dCollection.Add prp.Name, SecurePath(CStr(varValue))
            Case Else
                dCollection.Add prp.Name, prp.Value
        End Select
    Next prp
    
    ' Write to file. The order of properties may change, so sort them to keep the order consistent.
    WriteJsonFile Me, SortDictionaryByKeys(dCollection), IDbComponent_SourceFile, "Project Properties (Access)"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dExisting As Dictionary
    Dim prp As AccessObjectProperty
    Dim dImport As Dictionary
    Dim dItems As Dictionary
    Dim projCurrent As CurrentProject
    Dim varKey As Variant
    Dim varValue As Variant
    
    Set projCurrent = CurrentProject
    
    ' Pull a list of the existing properties so we know whether
    ' to add or update the existing property.
    Set dExisting = New Dictionary
    For Each prp In projCurrent.Properties
        Select Case prp.Name
            Case "Connection"
            Case Else
                dExisting.Add prp.Name, prp.Value
        End Select
    Next prp

    ' Read properties from source file
    Set dImport = ReadJsonFile(strFile)
    If Not dImport Is Nothing Then
        Set dItems = dImport("Items")
        For Each varKey In dItems.Keys
            Select Case varKey
                Case "Name", "Connection"
                    ' Skip these properties
                Case Else
                    varValue = Decrypt(dItems(varKey))
                    If Left$(varValue, 4) = "rel:" Then varValue = GetPathFromRelative(CStr(varValue))
                    If dExisting.Exists(varKey) Then
                        If dItems(varKey) <> dExisting(varKey) Then
                            ' Update value of existing property if different.
                            projCurrent.Properties(varKey).Value = varValue
                        End If
                    Else
                        ' Add properties that don't exist.
                        projCurrent.Properties.Add varKey, varValue
                    End If
            End Select
        Next varKey
    End If
    
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

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection
    
    Dim prp As AccessObjectProperty
    Dim cProp As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
        For Each prp In CurrentProject.Properties
            Set cProp = New clsDbProjProperty
            Set cProp.DbObject = prp
            m_AllItems.Add cProp, prp.Name
        Next prp
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
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = New Collection
    IDbComponent_GetFileList.Add IDbComponent_SourceFile
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
    ' Modified date unknown.
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
    IDbComponent_Category = "proj properties"
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
    IDbComponent_Name = "Current Project Properties (Access)"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "proj-properties.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbProjectProperty
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
    Set IDbComponent_DbObject = m_Property
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Property = RHS
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