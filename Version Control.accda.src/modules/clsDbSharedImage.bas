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


Private m_AllItems As Collection
Private m_Dbs As DAO.Database

' This is used to pass a reference to the record back
' into the class for loading the private variables
' with the actual file information.
Private m_Rst As DAO.Recordset

' File details used for exporting/importing
Private m_Name As String
Private m_FileName As String
Private m_Extension As String
Private m_FileData() As Byte

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the shared image as a json file with file details, and a copy
'           : of the binary image file saved as an image.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()

    Dim strFile As String
    Dim dItem As Dictionary
    Dim stm As ADODB.Stream
    
    ' Build header file
    Set dItem = New Dictionary
    dItem.Add "Name", m_Name
    dItem.Add "FileName", m_FileName
    dItem.Add "Extension", m_Extension
    
    ' Save json file with header details
    WriteJsonFile TypeName(Me), dItem, IDbComponent_SourceFile, "Shared Image Gallery Item"
    
    ' Save image file using extension from embedded file.
    strFile = IDbComponent_BaseFolder & FSO.GetBaseName(IDbComponent_SourceFile) & "." & FSO.GetExtensionName(m_FileName)
    Set stm = New ADODB.Stream
    With stm
        .Type = adTypeBinary
        .Open
        .Write StripOLEHeader(m_FileData)     ' Binary data
        VerifyPath strFile
        Perf.OperationStart "Write to Disk"
        .SaveToFile strFile, adSaveCreateOverWrite
        Perf.OperationEnd
        .Close
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StripOLEHeader
' Author    : Adam Waller
' Date      : 5/12/2020
' Purpose   : Strip out the OLE header so we can save the raw binary image as a
'           : readable file. (First 20 bytes (10 chars) of the data)
'           : Tested with jpg, gif, png, tiff
'---------------------------------------------------------------------------------------
'
Private Function StripOLEHeader(bteData() As Byte) As Byte()

    Dim strData As String
    
    ' Convert to string
    strData = bteData
    
    ' Strip off header, and convert back to byte array
    StripOLEHeader = Mid$(strData, 11)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dFile As Dictionary
    Dim dItem As Dictionary
    Dim strTemp As String
    Dim strImageFile As String
    Dim strOriginalName As String
    Dim strBase As String
    Dim lngIndex As Long
    Dim proj As CurrentProject
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Read json header file
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        Set dItem = dFile("Items")
        ' Check for an existing file with the same name
        strOriginalName = IDbComponent_BaseFolder & dItem("FileName")
        strImageFile = IDbComponent_BaseFolder & FSO.GetBaseName(strFile) & "." & FSO.GetExtensionName(strOriginalName)
        If strOriginalName <> strImageFile Then
            If FSO.FileExists(strOriginalName) Then
                strTemp = IDbComponent_BaseFolder & FSO.GetTempName
                ' Rename to temp file
                Name strOriginalName As strTemp
            End If
            Name strImageFile As strOriginalName
        End If
        ' Rename image to original name
        ' Import as image, then rename back to image file name that matches json file.
        Set proj = CurrentProject
        With proj
            lngIndex = .Resources.Count
            ' Import using the original file name as the resource name so the
            ' embedded file has the correct name.
            strBase = FSO.GetBaseName(strOriginalName)
            .AddSharedImage strBase, strOriginalName
            If .Resources.Count = lngIndex + 1 Then
                ' Rename shared resource to saved name if different.
                If strBase <> dItem("Name") Then
                    .Resources(GetResourceIndexByName(strBase)).Name = dItem("Name")
                End If
            End If
        End With
        ' Restore temp file if needed
        If strTemp <> vbNullString Then
            Name strTemp As strImageFile
        Else
            ' Restore image file name if needed.
            If strOriginalName <> strImageFile Then Name strOriginalName As strImageFile
        End If
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
' Procedure : GetResourceIndexByName
' Author    : Adam Waller
' Date      : 5/29/2020
' Purpose   : Return the index of the shared resource after locating by name.
'           : (This is needed because the new resource doesn't always have the
'           :  highest index.)
'---------------------------------------------------------------------------------------
'
Private Function GetResourceIndexByName(strName As String) As Long

    Dim lngIndex As Long
    Dim resShared As SharedResources
    
    Set resShared = CurrentProject.Resources
    For lngIndex = 0 To resShared.Count - 1
        If resShared(lngIndex).Name = strName Then
            GetResourceIndexByName = lngIndex
            Exit For
        End If
    Next lngIndex
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection

    Dim cImg As IDbComponent
    Dim rst As DAO.Recordset
    Dim strSql As String

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
        
        ' This system table should exist, but just in case...
        If TableExists("MSysResources") Then
            
            Set m_Dbs = CurrentDb
            strSql = "SELECT * FROM MSysResources WHERE Type='img'"
            Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)
            With rst
                Do While Not .EOF
                    Set cImg = New clsDbSharedImage
                    Set cImg.DbObject = rst    ' Reference to OLE object recordset2
                    m_AllItems.Add cImg
                    .MoveNext
                Loop
                .Close
            End With
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
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFiles Me, "json", "jpg", "jpeg", "jpe", "gif", "png", "ico"
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
    IDbComponent_Category = "Shared Images"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "images\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'           : In this case, we are building the name to include the info needed to
'           : recreate the record in the MSysResource table.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Name) & ".json"
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
    IDbComponent_ComponentType = edbSharedImage
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
    ' Not used
    Set IDbComponent_DbObject = Nothing
End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_DbObject
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Load in the class values from the recordset
'---------------------------------------------------------------------------------------
'
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)

    Dim fld2 As DAO.Field2
    Dim rst2 As DAO.Recordset2
    
    Set m_Rst = RHS
    
    ' Load in the object details.
    m_Name = m_Rst!Name
    m_Extension = m_Rst!Extension
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set fld2 = m_Rst!Data
    Set rst2 = fld2.Value
    m_FileName = rst2.Fields("FileName")
    m_FileData = rst2.Fields("FileData")
    
    ' Clear the object references
    Set rst2 = Nothing
    Set fld2 = Nothing
    Set m_Rst = Nothing

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


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 5/13/2020
' Purpose   : Clear reference to database object.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_Dbs = Nothing
End Sub
