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

Private m_Ref As VBIDE.Reference
Public AllItems As Collection


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

    Dim dRef As Dictionary
    Dim dItems As Dictionary
    Dim cRef As clsDbVbeReference
    Dim ref As VBIDE.Reference
    Dim strPath As String
    
    Set dItems = New Dictionary
    
    ' Loop through cached references (Duplicates have already been removed)
    For Each cRef In Me.AllItems
        Set dRef = New Dictionary
        Set ref = cRef.Parent.DbObject
        With dRef
            If ref.Type = vbext_rk_Project Then
                ' references of types mdb,accdb,mde etc don't have a GUID
                strPath = GetRelativePath(ref.FullPath)
                If strPath <> ref.FullPath Or Options.Security = esNone Then
                    ' Use relative path, or full path if not secured.
                    .Add "FullPath", strPath
                Else
                    ' Found a non-relative path.
                    .Add "File", FSO.GetFileName(ref.FullPath)
                    If Options.Security <> esRemove Then .Add "FullPath", Secure(ref.FullPath)
                End If
            Else
                If ref.Guid <> vbNullString Then .Add "GUID", ref.Guid
                .Add "Version", CStr(ref.Major) & "." & CStr(ref.Minor)
            End If
        End With
        dItems.Add ref.Name, dRef
    Next cRef
    
    ' Write to a json file.
    WriteJsonFile Me, dItems, IDbComponent_SourceFile, "VBE References"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dRef As Dictionary
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim ref As VBIDE.Reference
    Dim dFile As Dictionary
    Dim proj As VBProject
    Dim varVersion As Variant
    Dim strPath As String
    Dim dExisting As Dictionary
    
    ' Read in references from file
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        
        ' Build list of current references so we can avoid conflicts
        Set proj = GetVBProjectForCurrentDB
        Set dExisting = New Dictionary
        For Each ref In proj.References
            dExisting.Add ref.Name, ref.Guid
        Next ref
        
        ' Add any references from file that don't already exist
        Set dItems = dFile("Items")
        For Each varKey In dItems.Keys
            Set dRef = dItems(varKey)
            If Not dExisting.Exists(CStr(varKey)) Then
                If dRef.Exists("GUID") Then
                    varVersion = Split(dRef("Version"), ".")
                    AddFromGuid proj, CStr(varKey), dRef("GUID"), CLng(varVersion(0)), CLng(varVersion(1))
                ElseIf dRef.Exists("FullPath") Then
                    strPath = GetPathFromRelative(Decrypt(dRef("FullPath")))
                    If FSO.FileExists(strPath) Then
                        proj.References.AddFromFile strPath
                    Else
                        Log.Add "ERROR: Failed to add reference " & strPath
                    End If
                End If
            End If
        Next varKey
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddFromGuid
' Author    : Adam Waller
' Date      : 10/21/2020
' Purpose   : Return a GUID compatible with the current version of Microsoft Access.
'           : Only add exceptions here when they cause compile errors on older versions.
'           : Further reading: https://stackoverflow.com/questions/45088306
'           : https://www.fmsinc.com/microsoftaccess/history/features.htm
'           : https://kb.palisade.com/index.php?pg=kb.page&id=528
'---------------------------------------------------------------------------------------
'
Private Sub AddFromGuid(proj As VBIDE.VBProject, strName As String, strGuid As String, lngMajor As Long, lngMinor As Long)

    ' We might encounter a reference that is not available
    On Error GoTo ErrHandler
    
    Select Case strGuid
        Case "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"   ' Office
            Select Case Application.Version
                Case "14.0": proj.References.AddFromGuid strGuid, 2, 5  ' Access 2010
                Case "15.0": proj.References.AddFromGuid strGuid, 2, 6  ' Access 2013 (May need v.2.7)
                Case "16.0": proj.References.AddFromGuid strGuid, 2, 8  ' Access 2016, 2019, 365
                Case Else:   proj.References.AddFromGuid strGuid, lngMajor, lngMinor
            End Select
        Case "{00020813-0000-0000-C000-000000000046}"   ' Excel
            Select Case Application.Version
                Case "14.0": proj.References.AddFromGuid strGuid, 1, 7  ' Excel 2010
                Case "15.0": proj.References.AddFromGuid strGuid, 1, 8  ' Excel 2013
                Case "16.0": proj.References.AddFromGuid strGuid, 1, 9  ' Excel 2016, (2019, 365)?
                Case Else:   proj.References.AddFromGuid strGuid, lngMajor, lngMinor
            End Select
        Case Else
            ' Use specified GUID
            proj.References.AddFromGuid strGuid, lngMajor, lngMinor
    End Select
    
    ' Normal exit
    On Error GoTo 0
    Exit Sub

ErrHandler:

    ' Log error
    Log.Add "ERROR: Could not add VBE reference to " & strName
    
    If Err.Number = -2147319779 Then
        ' Object library not registered
        Log.Add "Encountered error " & Err.Number & ": '" & Err.Description & _
            "' while attempting to add GUID " & strGuid & " version " & lngMajor & "." & lngMinor & _
            " to this project. This may occur when the library does not exist on the build machine," & _
            " or when the version on the build machine is lower than the source file reference version." & _
            " See GitHub issue #96 for an example of how to resolve this problem.", Options.ShowDebug
        
    Else
        ' Other error
        Log.Add "Encountered error " & Err.Number & ": '" & Err.Description & _
            "' while attempting to add GUID " & strGuid & " version " & lngMajor & "." & lngMinor & _
            " to this project.", Options.ShowDebug
    End If
    
    ' Resume on next line
    Err.Clear
    Resume Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection
    
    Dim ref As VBIDE.Reference
    Dim cRef As clsDbVbeReference
    Dim colNames As Collection

    ' Build collection if not already cached
    If Me.AllItems Is Nothing Then
        Set Me.AllItems = New Collection
        Set colNames = New Collection
        For Each ref In GetVBProjectForCurrentDB.References
            If Not ref.BuiltIn Then
                Set cRef = New clsDbVbeReference
                Set cRef.Parent.DbObject = ref
                ' Export outputs single file, so every item needs a reference
                ' to the whole collection of references.
                Set cRef.AllItems = Me.AllItems
                ' Don't attempt add two references with the same name.
                ' (Take the first one, but ignore subsequent ones with the same name.)
                If Not InCollection(colNames, ref.Name) Then
                    Me.AllItems.Add cRef, ref.Name
                    colNames.Add ref.Name
                End If
            End If
        Next ref
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = Me.AllItems
        
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
    Dim strFile As String
    strFile = IDbComponent_BaseFolder & "references.csv"
    If FSO.FileExists(strFile) Then DeleteFile strFile, True    ' Remove legacy file
End Sub


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
    ' No date value here.
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
    IDbComponent_Category = "vbe references"
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
    IDbComponent_Name = m_Ref.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "vbe-references.json"
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
    IDbComponent_ComponentType = edbVbeReference
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
    Set IDbComponent_DbObject = m_Ref
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Ref = RHS
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