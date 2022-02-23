VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbProperty"
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

Private m_Property As DAO.Property
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
    If strAlternatePath <> vbNullString Then VCSIndex.Update Me, eatExport, GetStringHash(strContent, True)
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
    Dim prp As DAO.Property
    Dim dImport As Dictionary
    Dim dItems As Dictionary
    Dim dbs As DAO.Database
    Dim varKey As Variant
    Dim varValue As Variant
    Dim strDecrypted As String
    Dim blnAdd As Boolean
    Dim varItem As Variant
    Dim bArray() As Byte
    Dim i As Long
    Dim bUpdate As Boolean

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    Set dbs = CurrentDb
    
    ' Pull a list of the existing properties so we know whether
    ' to add or update the existing property.
    Set dExisting = New Dictionary
    For Each prp In dbs.Properties
        Select Case prp.Name
            Case "Connection"   ' This is an object.
            Case Else
                dExisting.Add prp.Name, Array(prp.Value, prp.Type)
        End Select
    Next prp

    ' Read properties from source file
    Set dImport = ReadJsonFile(strFile)
    If Not dImport Is Nothing Then
        Set dItems = dImport("Items")
        For Each varKey In dItems.Keys
            Select Case varKey
                Case "Connection", "Name", "Version", "CollatingOrder" ' Can't set these properties
                Case Else
                    blnAdd = False
                    bUpdate = False
                    ' Check if value is as Collection
                    If Not TypeOf dItems(varKey)("Value") Is Collection Then
                        varValue = dItems(varKey)("Value")
                        ' Check for relative path
                        If IsRelativePath(CStr(varValue)) Then varValue = GetPathFromRelative(CStr(varValue))
                    Else
                        ReDim bArray(0 To dItems(varKey)("Value").Count - 1)
                        For Each varItem In dItems(varKey)("Value")
                            bArray(i) = CByte(varItem)
                            i = i + 1
                        Next
                    End If
                    ' Check for existing value
                    If dExisting.Exists(varKey) Then
                        If dItems(varKey)("Type") <> dExisting(varKey)(1) Then
                            ' Type is different. Need to remove and add as correct type.
                            dbs.Properties.Delete varKey
                            blnAdd = True
                        Else
                            ' Check if value is a Collection
                            If Not TypeOf dItems(varKey)("Value") Is Collection Then
                                ' Check the value, and update if different
                                If varValue <> dExisting(varKey)(0) Then
                                    ' Update value of existing property if different.
                                    dbs.Properties(varKey).Value = varValue
                                End If
                            Else
                                ' Check the arrays, and update if different
                                If (LBound(bArray) <> LBound(dExisting(varKey)(0))) Or (UBound(bArray) <> UBound(dExisting(varKey)(0))) Then
                                    ' Different size
                                    bUpdate = True
                                Else
                                    ' Same size
                                    ' Check content
                                    For i = LBound(bArray) To UBound(bArray)
                                        If (bArray(i) <> dExisting(varKey)(0)(i)) Then
                                            bUpdate = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If bUpdate Then
                                    ' Update value of existing property if different.
                                    dbs.Properties(varKey).Value = bArray
                                End If
                            End If
                        End If
                    Else
                        ' Add properties that don't exist.
                        blnAdd = True
                    End If
                    
                    ' Can't add a text property with a null value. See issue #126
                    If dItems(varKey)("Type") = 10 Then
                        If varValue = vbNullChar Then blnAdd = False
                    End If
                    ' Add the property if the flag has been set.
                    If blnAdd Then
                        ' Check if value is a Collection
                        If Not TypeOf dItems(varKey)("Value") Is Collection Then
                            ' Create property
                            Set prp = dbs.CreateProperty(varKey, dItems(varKey)("Type"), varValue)
                        Else
                            ' Create property from array
                            Set prp = dbs.CreateProperty(varKey, dItems(varKey)("Type"), bArray)
                        End If
                        
                        ' Append property to collection
                        dbs.Properties.Append prp
                    End If
            End Select
        Next varKey
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
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "Database Properties (DAO)")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Return a dictionary of the ordered, unique VBE references
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary(Optional blnUseCache As Boolean) As Dictionary

    Dim prp As DAO.Property
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
    For Each prp In CurrentDb.Properties
        Select Case prp.Name
            Case "Connection"
                ' Connection object for ODBCDirect workspaces. Not needed.
            Case "Last VCS Export", "Last VCS Version"
                ' Legacy properties no longer needed.
            Case Else
                varValue = prp.Value
                If prp.Name = "AppIcon" Or prp.Name = "Name" Then
                    If Len(varValue) > 0 Then
                        ' Try to use a relative path
                        varValue = GetRelativePath(CStr(varValue))
                    End If
                End If
                Set dItem = New Dictionary
                dItem.Add "Value", varValue
                dItem.Add "Type", prp.Type
                dCollection.Add prp.Name, dItem
        End Select
    Next prp
    
    ' Return sorted dictionary
    Set GetDictionary = SortDictionaryByKeys(dCollection)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveMissing
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Removes current document properties missing from the master dictionary.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveMissing(dMaster As Dictionary, dTarget As Dictionary)

    Dim dbs As Database
    Dim varProp As Variant
    
    ' Go through target dictionary, removing properties that don't exist
    ' in the master dictionary. (Note that this is only checking the
    ' properties we are actually interested in tracking.)
    Set dbs = CurrentDb
    For Each varProp In dTarget.Keys
        ' Check to see if this key exists in the master
        If Not KeyExists(dMaster, varProp) Then
            ' Remove the property from the current database
            dbs.Properties.Delete CStr(varProp)
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
    
    Dim prp As DAO.Property
    Dim cProp As IDbComponent

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        If Not blnModifiedOnly Or IDbComponent_IsModified Then
            ' Return all the properties, since we don't know which ones
            ' were modified.
            For Each prp In CurrentDb.Properties
                Set cProp = New clsDbProperty
                Set cProp.DbObject = prp
                m_Items(blnModifiedOnly).Add cProp.SourceFile & ":" & prp.Name, cProp
            Next prp
        End If
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
    Dim strFile As String
    strFile = IDbComponent_BaseFolder & "properties.txt"
    If FSO.FileExists(strFile) Then DeleteFile strFile, True     ' Remove legacy file
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
    IDbComponent_Category = "DB Properties"
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
    IDbComponent_Name = "Database Properties (DAO)"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "dbs-properties.json"
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
    IDbComponent_ComponentType = edbDbsProperty
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
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Helps us know whether we have already counted the tables.
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

