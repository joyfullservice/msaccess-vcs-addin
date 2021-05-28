VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbDocument"
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

Private m_AllItems As Dictionary
Public m_dItems As Dictionary

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
    WriteJsonFile TypeName(Me), GetDictionary, IDbComponent_SourceFile, "Database Documents Properties (DAO)"
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

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dCont As Dictionary
    Dim dDoc As Dictionary
    Dim dbs As Database
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim varProp As Variant
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        ClearDatabaseSummaryProperties
        Set dbs = CurrentDb
        Set dItems = dFile("Items")
        For Each varCont In dItems.Keys
            Set dCont = dItems(varCont)
            For Each varDoc In dCont.Keys
                Set dDoc = dCont(varDoc)
                For Each varProp In dDoc.Keys
                    ' Attempt to add or update the property value on the object.
                    SetDAOProperty dbs.Containers(varCont).Documents(varDoc), dbText, CStr(varProp), dDoc(varProp)
                Next varProp
            Next varDoc
        Next varCont
    End If

    ' Update index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary(False))
    
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
' Procedure : ClearDatabaseSummaryProperties
' Author    : Adam Waller
' Date      : 5/13/2020
' Purpose   : When creating a new database, some properties may be filled out by
'           : default. Since the imported file only sets the ones that have values,
'           : it won't clear existing values that don't exist in the import file.
'           : I.e. `Company` may be already filled out as "Microsoft". This value would
'           : not be changed if the imported file did not specify this field.
'---------------------------------------------------------------------------------------
'
Private Sub ClearDatabaseSummaryProperties()

    Dim doc As DAO.Document
    Dim prp As DAO.Property
    Dim dbs As DAO.Database
    Dim intProp As Integer
    
    Set dbs = CurrentDb
    Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
    ' Loop backwards through the collection since we may be removing items.
    For intProp = doc.Properties.Count - 1 To 0 Step -1
        Set prp = doc.Properties(intProp)
        Select Case prp.Type
            Case dbText, dbMemo
                ' Text properties
                Select Case prp.Name
                    Case "Name", "Owner", "UserName", "Container" ' Leave these properties
                    Case Else
                        ' Remove other properties that might contain sensitive info.
                        ' They will be recreated from source files if they were in use.
                        doc.Properties.Delete prp.Name
                End Select
        End Select
    Next intProp
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Build a dictionary object of the document properties, just as you would
'           : use for the export content.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary(Optional blnUseCache As Boolean = True) As Dictionary

    Dim prp As DAO.Property
    Dim dItems As Dictionary    ' All Items
    Dim dCont As Dictionary     ' Container
    Dim dDoc As Dictionary      ' Document
    Dim cont As DAO.Container
    Dim dbs As Database
    Dim doc As DAO.Document
    Dim blnSave As Boolean
    
    ' Check cache parameter
    If blnUseCache And Not m_dItems Is Nothing Then
        ' Return cached dictionary
        Set GetDictionary = m_dItems
        Exit Function
    End If
    
    ' Create dictionary object to hold all the items
    Set dItems = New Dictionary
    Set dbs = CurrentDb
            
    ' Loop through all the containers, documents, and properties.
    ' Note, we don't want to collect everything here. We are taking
    ' a whitelist approach to specify the ones we want to save and
    ' write back to the database when importing.
    For Each cont In dbs.Containers
        Set dCont = New Dictionary
        For Each doc In cont.Documents
            Set dDoc = New Dictionary
            For Each prp In doc.Properties
                blnSave = False
                If cont.Name = "Databases" And doc.Name = "SummaryInfo" Then
                    ' Keep most of this information (Blacklist approach)
                    Select Case prp.Name
                        Case "AllPermissions", "Container", "DateCreated", "LastUpdated", _
                            "Name", "Owner", "GUID", "Permissions", "UserName" ' Ignore these
                        Case Else
                            blnSave = True
                    End Select
                Else
                    ' For other documents, use the whitelist approach, primarily
                    ' gathering navigation pane item descriptions and hidden status.
                    Select Case prp.Name
                        Case "Description"
                            blnSave = True
                    End Select
                End If
                ' If save flag set, save the property
                If blnSave Then dDoc.Add prp.Name, prp.Value
            Next prp
            If dDoc.Count > 0 Then dCont.Add doc.Name, SortDictionaryByKeys(dDoc)
        Next doc
        If dCont.Count > 0 Then dItems.Add cont.Name, SortDictionaryByKeys(dCont)
    Next cont

    ' Return assembled dictionary
    Set GetDictionary = dItems
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary
    
    Dim cDoc As IDbComponent
    Dim dCont As Dictionary
    Dim dDoc As Dictionary
    Dim dbs As Database
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim varProp As Variant
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then

        Set m_AllItems = New Dictionary
        Set m_dItems = GetDictionary
        Set dbs = CurrentDb
        
        ' Loop through all the containers, documents, and properties.
        For Each varCont In m_dItems.Keys
            Set dCont = m_dItems(varCont)
            For Each varDoc In dCont.Keys
                Set dDoc = dCont(varDoc)
                For Each varProp In dDoc.Keys
                    ' Add as class instance
                    Set cDoc = New clsDbDocument
                    Set cDoc.DbObject = dbs.Containers(varCont).Documents(varDoc).Properties(varProp)
                    m_AllItems.Add cDoc, vbNullString
                Next varProp
            Next varDoc
        Next varCont
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
    Set IDbComponent_GetFileList = New Collection
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
    IDbComponent_IsModified = VCSIndex.Item(Me)("Hash") <> GetDictionaryHash(GetDictionary)
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
    IDbComponent_Category = "Doc Properties"
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
    IDbComponent_Name = "Database Documents"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "documents.json"
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
    IDbComponent_ComponentType = edbDocument
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
    Set IDbComponent_DbObject = Nothing
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
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

