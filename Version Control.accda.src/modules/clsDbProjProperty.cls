VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbProjProperty"
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
Private m_AllItems As Dictionary
Private m_dItems As Dictionary
Private m_blnModifiedOnly As Boolean

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
    WriteJsonFile TypeName(Me), GetDictionary, Nz2(strAlternatePath, IDbComponent_SourceFile), "Project Properties (Access)"
    VCSIndex.Update Me, eatExport, GetDictionaryHash(GetDictionary)
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
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Pull a list of the existing properties so we know whether
    ' to add or update the existing property.
    Set projCurrent = CurrentProject
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
                    varValue = dItems(varKey)
                    If IsRelativePath(CStr(varValue)) Then varValue = GetPathFromRelative(CStr(varValue))
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
    
    ' Update index
    VCSIndex.Update Me, eatExport, GetDictionaryHash(GetDictionary(False))
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 5/29/2021
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)

    Dim dFile As Dictionary
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub
    
    ' Remove any document properties that don't exist in the incoming file,
    ' then import the file.
    Set dFile = ReadJsonFile(strFile)
    If dFile Is Nothing Then Set dFile = New Dictionary
    RemoveMissing dFile("Items"), GetDictionary
    IDbComponent_Import strFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/29/2021
' Purpose   : Return a dictionary of the ordered database properties
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary(Optional blnUseCache As Boolean) As Dictionary

    Dim prp As AccessObjectProperty
    Dim dCollection As Dictionary
    Dim dItem As Dictionary
    Dim varValue As Variant
    
    ' Check cache parameter
    If blnUseCache And Not m_dItems Is Nothing Then
        ' Return cached dictionary
        Set GetDictionary = m_dItems
        Exit Function
    End If
    
    Set dCollection = New Dictionary
    
    ' Loop through all properties
    For Each prp In CurrentProject.Properties
        Select Case prp.Name
            Case "Connection"
                ' Connection object for ODBCDirect workspaces. Not needed.
            Case "Last VCS Export", "Last VCS Version"
                ' Legacy properties no longer needed.
            Case "AppIcon"
                ' ADP projects may have this property
                dCollection.Add prp.Name, GetRelativePath(CStr(prp.Value))
            Case Else
                dCollection.Add prp.Name, prp.Value
        End Select
    Next prp
    
    ' Return sorted dictionary
    Set GetDictionary = SortDictionaryByKeys(dCollection)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveMissing
' Author    : Adam Waller
' Date      : 5/29/2021
' Purpose   : Removes current document properties missing from the master dictionary.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveMissing(dMaster As Dictionary, dTarget As Dictionary)

    Dim proj As CurrentProject
    Dim varProp As Variant
    
    ' Go through target dictionary, removing properties that don't exist
    ' in the master dictionary. (Note that this is only checking the
    ' properties we are actually interested in tracking.)
    Set proj = CurrentProject
    For Each varProp In dTarget.Keys
        ' Check to see if this key exists in the master
        If Not KeyExists(dMaster, varProp) Then
            ' Remove the property from the current project
            proj.Properties.Remove CStr(varProp)
        End If
    Next varProp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary
    
    Dim prp As AccessObjectProperty
    Dim cProp As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Dictionary
        If Not blnModifiedOnly Or IDbComponent_IsModified Then
            ' Return all the properties, since we don't know which ones
            ' were modified.
            For Each prp In CurrentProject.Properties
                Set cProp = New clsDbProjProperty
                Set cProp.DbObject = prp
                m_AllItems.Add cProp.SourceFile & ":" & prp.Name, cProp
            Next prp
        End If
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
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = VCSIndex.Item(Me).Hash <> GetDictionaryHash(GetDictionary)
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
    IDbComponent_Category = "Proj Properties"
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

