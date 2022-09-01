VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbHiddenAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ModuleName As String = "clsDbHiddenAttribute"

Private m_Items(True To False) As Dictionary
Private m_dItems As Dictionary

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)
    Dim strContent As String
    strContent = GetSource
    WriteFile strContent, Nz2(strAlternatePath, IDbComponent_SourceFile)
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), GetStringHash(strContent, True)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dbs As Database
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim objType As AcObjectType

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        Set dbs = CurrentDb
        Set dItems = dFile("Items")
        For Each varCont In dItems.Keys
            objType = GetObjectTypeFromContainer(dbs.Containers(varCont))
            If objType <> acDefault Then
                For Each varDoc In dItems(varCont)
                    ' Set object to hidden
                    Application.SetHiddenAttribute objType, varDoc, True
                    CatchAny eelError, "Error setting hidden attribute for " & varDoc, ModuleName
                Next varDoc
            End If
        Next varCont
    End If

    ' Update index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary(False))

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 5/28/2021
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
    
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary
    
    Dim cDoc As IDbComponent
    Dim colCont As Collection
    Dim dbs As Database
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim blnNoChanges As Boolean
    Dim strKey As String
    
    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then

        Set m_Items(blnModifiedOnly) = New Dictionary
        Set m_dItems = GetDictionary
        
        ' Return no objects if we match the hash from the cache
        If (Not blnModifiedOnly) Or IDbComponent_IsModified Then
            
            ' Loop through all the containers, documents, and check hidden property
            Set dbs = CurrentDb
            For Each varCont In m_dItems.Keys
                Set colCont = m_dItems(varCont)
                For Each varDoc In colCont
                    Set cDoc = New clsDbHiddenAttribute
                    Set cDoc.DbObject = dbs.Containers(varCont).Documents(varDoc)
                    strKey = cDoc.SourceFile & ":" & varCont & ":" & varDoc
                    m_Items(blnModifiedOnly).Add strKey, cDoc
                Next varDoc
            Next varCont
        End If
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
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "Database objects hidden attribute")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Return a dictionary object with the items
'---------------------------------------------------------------------------------------
'
Public Function GetDictionary(Optional blnUseCache As Boolean = True) As Dictionary

    Dim dItems As Dictionary
    Dim dCont As Dictionary
    Dim cont As DAO.Container
    Dim doc As DAO.Document
    Dim dbs As Database
    Dim contType As AcObjectType
    Dim colItems As Collection
        
    ' Check cache parameter
    If blnUseCache And Not m_dItems Is Nothing Then
        ' Return cached dictionary
        Set GetDictionary = m_dItems
        Exit Function
    End If

    ' Create dictionary object to hold all the items
    Set dItems = New Dictionary
    Set dbs = CurrentDb

    ' Loop through all the containers, documents, and check hidden property
    For Each cont In dbs.Containers
        Set dCont = New Dictionary
        Set dCont = New Dictionary
        Set colItems = New Collection
        contType = GetObjectTypeFromContainer(cont)
        For Each doc In cont.Documents
            If contType <> acDefault _
                And Not (contType = acTable _
                And (doc.Name Like "MSys*" Or doc.Name Like "~*")) Then
                ' Check Hidden Attribute property (only exposed here)
                If Application.GetHiddenAttribute(contType, doc.Name) Then
                    ' Add to collection of hidden item item names of this type.
                    colItems.Add doc.Name
                End If
            End If
        Next doc
        If colItems.Count > 0 Then dItems.Add cont.Name, SortCollectionByValue(colItems)
    Next cont
    
    ' Return assembled dictionary
    Set GetDictionary = dItems
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveMissing
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Unsets the hidden flag for documents that should not have it.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveMissing(dMaster As Dictionary, dTarget As Dictionary)

    Dim dCont As Dictionary
    Dim dDoc As Dictionary
    Dim dbs As Database
    Dim contType As AcObjectType
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim strDoc As String
    
    ' Go through target dictionary, removing the flag that doesn't exist
    ' in the master dictionary.
    Set dbs = CurrentDb
    For Each varCont In dTarget.Keys
        Set dCont = dTarget(varCont)
        contType = GetObjectTypeFromContainer(dbs.Containers(varCont))
        For Each varDoc In dCont.Keys
            strDoc = CStr(varDoc)
            If contType <> acDefault _
                And Not (contType = acTable _
                And (strDoc Like "MSys*" Or strDoc Like "~*")) Then
                ' Unset the Hidden Attribute property (only exposed here)
                Application.SetHiddenAttribute contType, strDoc, False
            End If
        Next varDoc
    Next varCont

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetObjectTypeFromContainer
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Get an object type from a DAO container
'---------------------------------------------------------------------------------------
'
Private Function GetObjectTypeFromContainer(ByRef cont As DAO.Container) As AcObjectType
    Select Case cont.Name
        Case "Tables"
            GetObjectTypeFromContainer = acTable
        Case "Forms"
            GetObjectTypeFromContainer = acForm
        Case "Scripts"
            GetObjectTypeFromContainer = acMacro
        Case "Queries"
            GetObjectTypeFromContainer = acQuery
        Case "Reports"
            GetObjectTypeFromContainer = acReport
        Case "Modules"
            GetObjectTypeFromContainer = acModule
        Case Else
            ' Unknown
            GetObjectTypeFromContainer = acDefault
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    Set IDbComponent_GetFileList = New Dictionary
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_GetFileList.Add IDbComponent_SourceFile, vbNullString
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    If FSO.FileExists(IDbComponent_SourceFile) Then
        If IDbComponent_GetAllFromDB.Count = 0 Then
            ' Remove existing file, since we don't have any hidden objects.
            DeleteFile IDbComponent_SourceFile
        End If
    End If
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
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
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
' Author    : Adam Waller / Indigo744
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
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Hidden Attributes"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = "Database hidden attributes"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "hidden-attributes.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
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
    ' Single file output
    IDbComponent_QuickCount = 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbHiddenAttribute
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller / Indigo744
' Date      : 11/14/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = Nothing
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller / Indigo744
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
' Author    : Adam Waller / Indigo744
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property

