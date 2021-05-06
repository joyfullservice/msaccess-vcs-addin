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

Private Const ModuleName As String = "clsDbTheme"

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
' Purpose   : Export the theme file as either a zipped thmx file, or an extracted
'           : folder with the theme source files. (Depending on the specified options.)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()

    Dim strFile As String
    Dim strZip As String
    Dim strFolder As String
    Dim rst As Recordset2
    Dim rstAtc As Recordset2
    Dim strSql As String
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Query theme file details
    strSql = "SELECT [Data] FROM MSysResources WHERE [Name]='" & m_Name & "' AND Extension='" & m_Extension & "'"
    Set m_Dbs = CurrentDb
    Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)
    
    ' If we get multiple records back we don't know which to use
    If rst.RecordCount > 1 Then
        Log.Error eelCritical, "Multiple records in MSysResources table were found that matched this name. " & _
            "Compact and repair database and try again. Theme Name: " & LCase(m_Name) & "." & m_Extension, ModuleName & ".Export"
        Exit Sub
    End If

    ' Get full name of theme file. (*.thmx)
    strFile = IDbComponent_SourceFile
    
    ' Save as file
    If Not rst.EOF Then
        Set rstAtc = rst!Data.Value
        If FSO.FileExists(strFile) Then DeleteFile strFile, True
        VerifyPath strFile
        Perf.OperationStart "Export Theme"
        rstAtc!FileData.SaveToFile strFile
        Perf.OperationEnd
        rstAtc.Close
        Set rstAtc = Nothing
    End If
    rst.Close
    Set rst = Nothing

    CatchAny eelError, "Error exporting theme file: " & strFile, ModuleName & ".Export", True, True

    ' See if we need to extract the theme source files.
    ' (Only really needed when you are tracking themes customizations.)
    If Options.ExtractThemeFiles Then
        Perf.OperationStart "Extract Theme"
        ' Extract to folder and delete zip file.
        strFolder = FSO.BuildPath(FSO.GetParentFolderName(strFile), FSO.GetBaseName(strFile))
        If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
        DoEvents ' Make sure the folder is deleted before we recreate it.
        ' Rename to zip file before extracting
        strZip = strFolder & ".zip"
        Name strFile As strZip
        ExtractFromZip strZip, strFolder, False
        ' Rather than holding up the export while we extract the file,
        ' use a cleanup sub to do this after the export.
        Perf.OperationEnd
        CatchAny eelError, "Error extracting theme. Folder: " & strFolder, ModuleName & ".Export", True, True
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim rstResources As DAO.Recordset2
    Dim rstAttachment As DAO.Recordset2
    Dim fldFile As DAO.Field2
    Dim strZip As String
    Dim strThemeFile As String
    Dim strThemeName As String
    Dim strSql As String
    Dim blnIsFolder As Boolean
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Are we dealing with a folder, or a file?
    blnIsFolder = (Right$(strFile, 5) <> ".thmx")

    If blnIsFolder Then
        ' We need to compress this folder back into a zipped theme file.
        ' Build zip file name; if it's a folder, just add the extension.
        strZip = strFile & ".zip"
        ' Get theme name
        strThemeName = GetObjectNameFromFileName(FSO.GetBaseName(strZip))
        ' Remove any existing zip file
        If FSO.FileExists(strZip) Then DeleteFile strZip, True
        ' Copy source files into new zip file
        CreateZipFile strZip
        CopyFolderToZip strFile, strZip
        DoEvents
        strThemeFile = strFile & ".thmx"
        If FSO.FileExists(strThemeFile) Then DeleteFile strThemeFile, True
        Name strZip As strThemeFile
    Else
        ' Skip if file no longer exists. (Such as if we already
        ' imported this theme from a folder.)
        If Not FSO.FileExists(strFile) Then Exit Sub
        ' Theme file is ready to go
        strThemeFile = strFile
    End If

    ' Log any errors encountered.
    CatchAny eelError, "Error getting theme file. File: " & strThemeFile & ", IsFolder: " & blnIsFolder, ModuleName & ".Import", True, True

    ' Create/edit record in resources table.
    strThemeName = GetObjectNameFromFileName(FSO.GetBaseName(strFile))
    ' Make sure we have a resources table before we try to query the records.
    If VerifyResourcesTable(True) Then
        strSql = "SELECT * FROM MSysResources WHERE [Type] = 'thmx' AND [Name]=""" & strThemeName & """"
        Set rstResources = CurrentDb.OpenRecordset(strSql, dbOpenDynaset)
        With rstResources
            If .EOF Then
                ' No existing record found. Add a record
                .AddNew
                !Name = strThemeName
                !Extension = "thmx"
                !Type = "thmx"
                Set rstAttachment = .Fields("Data").Value
            Else
                ' Found theme record with the same name.
                ' Remove the attached theme file.
                .Edit
                Set rstAttachment = .Fields("Data").Value
                If Not rstAttachment.EOF Then rstAttachment.Delete
            End If
            
            ' Upload theme file into OLE field
            DoEvents
            With rstAttachment
                .AddNew
                Set fldFile = .Fields("FileData")
                fldFile.LoadFromFile strThemeFile
                .Update
                .Close
            End With
            
            ' Save and close record
            .Update
            .Close
        End With
    End If
    
    ' Remove compressed theme file if we are using a folder.
    If blnIsFolder Then DeleteFile strThemeFile, True
    
    ' Log any errors
    CatchAny eelError, "Error importing theme. File: " & strThemeFile & ", IsFolder: " & blnIsFolder, ModuleName & ".Import", True, True

    ' Clear object (Important with DAO/ADO)
    Set rstAttachment = Nothing
    Set rstResources = Nothing

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
    
    Dim cTheme As IDbComponent
    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim strKey As String
    Dim dItems As Dictionary
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
            
        ' Use dictionary to make sure we don't add duplicate records if we have
        ' both a folder and a theme file for the same theme.
        Set dItems = New Dictionary
        
        ' This system table should exist, but just in case...
        If TableExists("MSysResources") Then

            Set m_Dbs = CurrentDb
            strSql = "SELECT * FROM MSysResources WHERE Type='thmx'"
            Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)
            With rst
                Do While Not .EOF
                    Set cTheme = New clsDbTheme
                    Set cTheme.DbObject = rst    ' Reference to OLE object recordset2
                    strKey = Nz(!Name)
                    If Not dItems.Exists(strKey) Then
                        m_AllItems.Add cTheme, strKey
                        dItems.Add strKey, strKey
                    End If
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
' Procedure : VerifyResourceTable
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Make sure the resources table exists, creating it if needed.
'---------------------------------------------------------------------------------------
'
Public Function VerifyResourcesTable(blnClearThemes As Boolean) As Boolean

    Dim strName As String
    
    ' Make sure we actually have a resources table.
    If Not TableExists("MSysResources") Then
        ' It would be nice to find a magical system command for this, but for now
        ' we can create it by creating a temporary form object.
        strName = CreateForm().Name
        ' Close without saving
        DoCmd.Close acForm, strName, acSaveNo
        ' Remove any potential default theme
        If TableExists("MSysResources") Then
            If blnClearThemes Then CurrentDb.Execute "DELETE * FROM MSysResources WHERE [Type]='thmx'", dbFailOnError
            VerifyResourcesTable = True
        Else
            Log.Add "WARNING: Unable to create MSysResources table."
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    ' Get list of folders (extracted files) as well as zip files.
    Set IDbComponent_GetFileList = GetSubfolderPaths(IDbComponent_BaseFolder)
    MergeCollection IDbComponent_GetFileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*.thmx")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFolders Me
    ClearOrphanedSourceFiles Me, "thmx"
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
    '// TODO: Recursively identify the most recent file modified date.
    'If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = GetLastModifiedDate(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Themes"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "themes" & PathSep
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
'           : In this case, we are returning the theme file name.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Name) & ".thmx"
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
    IDbComponent_ComponentType = edbTheme
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