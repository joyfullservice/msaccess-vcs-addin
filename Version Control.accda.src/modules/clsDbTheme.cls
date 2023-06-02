VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTheme"
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

Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary
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
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strFile As String
    Dim strZip As String
    Dim strFolder As String
    Dim rst As Recordset2
    Dim rstAtc As Recordset2
    Dim strSql As String
    Dim strHash As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Query theme file details
    strSql = "SELECT [Data] FROM MSysResources WHERE [Name]='" & m_Name & "' AND Extension='" & m_Extension & "'"
    Set m_Dbs = CurrentDb
    Set rst = m_Dbs.OpenRecordset(strSql, dbOpenSnapshot, dbOpenForwardOnly)

    ' If we get multiple records back we don't know which to use
    If rst.RecordCount > 1 Then
        Log.Error eelCritical, "Multiple records in MSysResources table were found that matched this name. " & _
            "Compact and repair database and try again. Theme Name: " & LCase(m_Name) & "." & m_Extension, ModuleName(Me) & ".Export"
        Exit Sub
    End If

    ' Get full name of theme file. (*.thmx)
    strFile = Nz2(strAlternatePath, IDbComponent_SourceFile)

    ' Save as file
    If Not rst.EOF Then
        Set rstAtc = rst!Data.Value
        If FSO.FileExists(strFile) Then DeleteFile strFile, True
        If FSO.FolderExists(strFile) Then FSO.DeleteFolder strFile, True
        VerifyPath strFile
        Perf.OperationStart "Export Theme"
        rstAtc!FileData.SaveToFile strFile
        Perf.OperationEnd
        rstAtc.Close
        Set rstAtc = Nothing
    End If
    rst.Close
    Set rst = Nothing

    CatchAny eelError, "Error exporting theme file: " & strFile, ModuleName(Me) & ".Export", True, True

    ' Get hash before extracting the theme
    strHash = GetFileHash(strFile)

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
        If FSO.FileExists(strZip) Then DeleteFile (strZip)
        Name strFile As strZip
        ExtractFromZip strZip, strFolder, False
        ' Rather than holding up the export while we extract the file,
        ' use a cleanup sub to do this after the export.
        Perf.OperationEnd
        CatchAny eelError, "Error extracting theme. Folder: " & strFolder, ModuleName(Me) & ".Export", True, True
    End If

    ' Update index after exporting theme
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), strHash, GetSignatureHash

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    ' IMPORTANT: These variables must be declared as the
    ' generic parent classes to avoid problems in newer
    ' releases of Microsoft Access. (See issue #341)
    Dim rstResources As DAO.Recordset   ' Recordset2
    Dim rstAttachment As DAO.Recordset  ' Recordset2
    Dim fldFile As DAO.Field            ' Field2

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
    CatchAny eelError, "Error getting theme file. File: " & strThemeFile & ", IsFolder: " & blnIsFolder, ModuleName(Me) & ".Import", True, True

    ' Create/edit record in resources table.
    strThemeName = GetObjectNameFromFileName(FSO.GetBaseName(strFile))
    ' Make sure we have a resources table before we try to query the records.
    If VerifyResourcesTable Then
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
                m_FileData = fldFile
                .Update
                .Close
            End With

            ' Update class variables so we can recalculate the index hash
            m_Name = strThemeName
            m_FileName = strThemeName & ".thmx"
            m_Extension = "thmx"

            ' Save and close record
            .Update
            .Close
        End With
    End If

    ' Remove compressed theme file if we are using a folder.
    If blnIsFolder Then DeleteFile strThemeFile, True

    ' Update index
    VCSIndex.Update Me, eatImport, GetSignatureHash

    ' Log any errors
    CatchAny eelError, "Error importing theme. File: " & strThemeFile & ", IsFolder: " & blnIsFolder, ModuleName(Me) & ".Import", True, True

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
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        ' The import will replace the existing theme with the same name.
        IDbComponent_Import strFile
    Else
        'TODO: Remove existing theme
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSignatureHash
' Author    : Adam Waller
' Date      : 5/28/2021
' Purpose   : Return a simple hash that should give us a pretty good idea if the theme
'           : has changed since it was last exported.
'---------------------------------------------------------------------------------------
'
Private Function GetSignatureHash() As String

    Dim intCnt As Integer
    Dim lngLen As Long

    ' Get length (bytes) of file data
    If StrPtr(m_FileData) <> 0 Then lngLen = UBound(m_FileData)

    With New clsConcat
        ' Compile string with file information
        .Add m_Name, m_FileName, m_Extension, CStr(lngLen)
        ' Add in last 100 characters of file content (byte numbers)
        ' Since this is a compressed (zip) file, we will probably
        ' see a change here if the contents were modified.
        If lngLen > 100 Then
            For intCnt = lngLen To lngLen - 100 Step -1
                .Add CStr(m_FileData(intCnt))
            Next intCnt
        End If
        ' Return hash of this string
        GetSignatureHash = GetStringHash(.GetStr)
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
    ' Check for extracted theme files
    MoveFolderIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim cTheme As IDbComponent
    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim strKey As String
    Dim dItems As Dictionary
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary

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
                    strKey = Nz(!Name)
                    If Not dItems.Exists(strKey) Then
                        ' Load theme details so we can check modified status
                        Set cTheme = New clsDbTheme
                        Set cTheme.DbObject = rst    ' Reference to OLE object recordset2
                        blnAdd = True
                        If blnModifiedOnly Then blnAdd = cTheme.IsModified
                        If blnAdd Then m_Items(blnModifiedOnly).Add cTheme.SourceFile, cTheme
                        dItems.Add strKey, strKey
                    End If
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyResourceTable
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Make sure the resources table exists, creating it if needed.
'---------------------------------------------------------------------------------------
'
Public Function VerifyResourcesTable() As Boolean

    Dim blnExists As Boolean
    Dim strName As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure we actually have a resources table.
    blnExists = TableExists("MSysResources")
    If Not blnExists Then
        CreateResourcesTable
        blnExists = TableExists("MSysResources")
    End If

    ' Return true if the table exists
    VerifyResourcesTable = blnExists

    ' Log any errors
    CatchAny eelError, "Error verifying MSysResources table", ModuleName(Me) & ".VerifyResourcesTable"

End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateResourcesTable
' Author    : Adam Waller
' Date      : 7/9/2021
' Purpose   : The resources table is also used for shared images. We can add and remove
'           : a temporary shared image to cause this table to be generated.
'---------------------------------------------------------------------------------------
'
Private Sub CreateResourcesTable()

    Dim strTempFile As String
    Dim strName As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    Perf.OperationStart "Create MSysResources Table"

    ' Create a temporary image file
    strTempFile = GetTempFile("IMG")
    Create1x1pxImage strTempFile

    ' Add to database as shared image
    strName = FSO.GetBaseName(strTempFile)
    CurrentProject.AddSharedImage strName, strTempFile

    ' Now that we have the table, remove temporary shared image record and temp file.
    CurrentDb.Execute "DELETE * FROM MSysResources WHERE Name='" & strName & "'", dbFailOnError
    DeleteFile strTempFile

    ' Log any errors
    CatchAny eelError, "Error creating MSysResources table", ModuleName(Me) & ".CreateResourcesTable"
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Create1x1pxImage
' Author    : Adam Waller
' Date      : 7/9/2021
' Purpose   : Creates a 1 pixel by 1 pixel png image.
'---------------------------------------------------------------------------------------
'
Private Sub Create1x1pxImage(strPath As String)

    Dim intCnt As Integer
    Dim varBytes As Variant
    Dim bteImg() As Byte

    ' Byte array for binary 1x1 pixel png file.
    varBytes = Array((137), (80), (78), (71), (13), (10), (26), (10), (0), (0), (0), (13), (73), (72), (68), _
        (82), (0), (0), (0), (1), (0), (0), (0), (1), (1), (3), (0), (0), (0), (37), (219), (86), (202), (0), _
        (0), (0), (3), (80), (76), (84), (69), (0), (0), (0), (167), (122), (61), (218), (0), (0), (0), (1), _
        (116), (82), (78), (83), (0), (64), (230), (216), (102), (0), (0), (0), (10), (73), (68), (65), (84), _
        (8), (215), (99), (96), (0), (0), (0), (2), (0), (1), (226), (33), (188), (51), (0), (0), (0), (0), _
        (73), (69), (78), (68), (174), (66), (96), (130))

    ' Convert to byte array
    ReDim bteImg(0 To UBound(varBytes)) As Byte
    For intCnt = 0 To UBound(varBytes)
        bteImg(intCnt) = varBytes(intCnt)
    Next intCnt

    ' Write to file
    WriteBinaryFile strPath, bteImg

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    ' Get list of folders (extracted files) as well as zip files.
    If m_FileList Is Nothing Then
        Set m_FileList = GetSubfolderPaths(IDbComponent_BaseFolder)
        MergeDictionary m_FileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*.thmx")
    End If
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFiles Me, "thmx"
    ClearOrphanedSourceFolders Me
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
    ' Compare the index hash with the current theme signature hash
    IDbComponent_IsModified = VCSIndex.Item(Me).OtherHash <> GetSignatureHash
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
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    IDbComponent_QuickCount = 1
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

    ' IMPORTANT: These variables must be declared as the
    ' generic parent classes to avoid problems in newer
    ' releases of Microsoft Access. (See issue #341)
    Dim fld2 As DAO.Field       ' Field2
    Dim rst2 As DAO.Recordset   ' Recordset2

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
