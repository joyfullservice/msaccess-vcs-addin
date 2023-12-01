VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'           :
'           : This class actually doesn't import or export anything but provides
'           : useful metadata about the connections that the database project has.
'           : During "export", all possible connections are harvested while "import"
'           : will prime Access' internal cache, which helps reduce the numbers of
'           : ODBC login dialogs that may pop open during the imports of associated
'           : database objects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_Items(True To False) As Dictionary   ' Instances of this class
Private m_dItems As Dictionary  ' Connection items for JSON output in GetDictionary

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 5/17/2021
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
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dConnections As Dictionary
    Dim varConnect As Variant

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub
    Set dConnections = ReadJsonFile(strFile).Item("Items")

    ' Prime Access' internal cache. If the connection string is incomplete, it will
    ' pop open a dialog from the driver for the user to then fill in.
    For Each varConnect In dConnections
        ' We only need to process the sanitized connection strings, rather than possible
        ' connection strings which help reduce number of logins user may need to
        ' complete.
        Dim x As Variant
        For Each x In dConnections(varConnect)
            CacheConnection CStr(x)
        Next
    Next

    CatchAny eelError, "Importing Connection", ModuleName(Me) & ".Import"

    ' Save to index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, _
        "Database connections used within the project used to prime Access' internal cache during import.")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a dictionary object of project connections.
'---------------------------------------------------------------------------------------
'
Public Function GetDictionary(Optional blnUseCache As Boolean = True) As Dictionary

    ' Check cache first
    If blnUseCache And Not m_dItems Is Nothing Then
        Set GetDictionary = m_dItems
        Exit Function
    End If

    ' Read project connections
    Set GetDictionary = New Dictionary

    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim cConnection As clsConnectionODBC

    Set cConnection = New clsConnectionODBC
    With CurrentDb
        For Each tdf In .TableDefs
            If StartsWith(tdf.Connect, "ODBC;", vbTextCompare) And (Len(tdf.Connect) > 5) Then
                If StartsWith(tdf.Name, "~") Then
                    ' Ignore temporary table objects
                Else
                    cConnection.ParseOdbcConnectionString tdf.Connect
                    AddOdbcConnectionString GetDictionary, cConnection, tdf.Name
                End If
            End If
        Next
        For Each qdf In .QueryDefs
            ' We have to handle two possible scenarios:
            ' 1. Sometimes a certain querydef will throw an unrelated error when trying to load
            '    querydef by accessing one of its properties for the first time. In this case,
            '    we simply retry and we are usually able to read the properties of the querydef.
            ' 2. Sometimes the querydef is simply unusable and will always throw errors when
            '    trying to read the property. We should just skip over them and hope for the best.
            Dim strQdfConnect As String
            On Error Resume Next
            strQdfConnect = qdf.Connect
            If Err.Number Then
                Err.Clear
                strQdfConnect = qdf.Connect
                If Err.Number Then
                    ' Even though we cannot read the Connect property, Name property should be always available.
                    CatchAny eelWarning, "Unable to read the Connect property for query '" & qdf.Name & "'. Error " & Err.Number & ": " & Err.Description, ModuleName(Me) & ".GetDictionary"
                    strQdfConnect = vbNullString
                End If
            End If
            On Error GoTo 0
            If StartsWith(strQdfConnect, "ODBC;", vbTextCompare) And (Len(strQdfConnect) > 5) Then
                If StartsWith(qdf.Name, "~") Then
                    ' Ignore temporary queries
                Else
                    cConnection.ParseOdbcConnectionString strQdfConnect
                    AddOdbcConnectionString GetDictionary, cConnection, qdf.Name
                End If
            End If
        Next
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddOdbcConnectionString
' Author    : bclothier
' Date      : 4/1/2023
' Purpose   : The values of the GetDictionary are themselves a dictionary of the original
'           : ODBC connection strings. This procedure helps handle creating a new
'           : dictionary if there's not an entry for the sanitized connection string
'           : already and then adds the original connection string under the key of
'           : sanitized connection string.
'---------------------------------------------------------------------------------------
'
Private Function AddOdbcConnectionString(dSanitizedConnections As Dictionary, cConnection As clsConnectionODBC, strObjectName As String)
    Dim dConnections As Dictionary

    If dSanitizedConnections.Exists(cConnection.SanitizedConnectionString) Then
        Set dConnections = dSanitizedConnections(cConnection.SanitizedConnectionString)
    Else
        Set dConnections = New Dictionary
        dConnections.CompareMode = TextCompare
        dSanitizedConnections.Add cConnection.SanitizedConnectionString, dConnections
    End If
    If dConnections.Exists(cConnection.OriginalConnectionString) = False Then
        dConnections.Add cConnection.OriginalConnectionString, strObjectName
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim varKey As Variant

    ' Build a dictionary of component class objects if not already cached.
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        If Not blnModifiedOnly Or IDbComponent_IsModified Then
            ' Return all the items since we are not concerned about which ones changed.
            For Each varKey In GetDictionary.Keys
                ' Since we are only processing one of these instances, we can
                ' just reference the existing instance and avoid rescanning the
                ' tables and queries.
                m_Items(blnModifiedOnly).Add varKey, Me
            Next varKey
        End If
    End If

    ' Return cached dictionary
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 5/17/2021
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
' Date      : 5/17/2021
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    If FSO.FileExists(IDbComponent_SourceFile) Then
        ' Delete the source file when we no longer have any connection strings
        If IDbComponent_GetAllFromDB.Count = 0 Then DeleteFile IDbComponent_SourceFile
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
' Date      : 5/17/2021
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceIsModified
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Returns true if the source file appears to be modified since the last
'           : import or export. (When compared to the index)
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceIsModified(strFile As String) As Boolean

End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "DB Connections"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = "Database Connection"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "db-connection.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 5/17/2021
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
' Date      : 5/17/2021
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbConnection
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    'Not applicable
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    'Not applicable
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 5/17/2021
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
' Date      : 5/17/2021
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
