Attribute VB_Name = "modContainers"
'---------------------------------------------------------------------------------------
' Module    : modContainers
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Component type registry, container lookup, and source file metadata
'           : functions for the IDbComponent system.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modFileAccess, modHash, modFunctions
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName = "modContainers"


'---------------------------------------------------------------------------------------
' Procedure : GetAllContainers
' Author    : Adam Waller
' Date      : 5/4/2020
' Purpose   : Return a collection of all containers.
'           : NOTE: The order doesn't matter for export, but is VERY important
'           : when building the project from source.
'---------------------------------------------------------------------------------------
'
Public Function GetContainers(Optional intFilter As eContainerFilter = ecfAllObjects) As Collection

    Dim blnADP As Boolean
    Dim blnMDB As Boolean

    blnADP = (CurrentProject.ProjectType = acADP)
    blnMDB = (CurrentProject.ProjectType = acMDB)

    Set GetContainers = New Collection
    With GetContainers
        Select Case intFilter

            ' Primary case for processing all objects
            Case ecfAllObjects
                If blnMDB Then
                    ' Handle the connections early as possible but only for MDB formats
                    .Add New clsDbConnection
                End If

                ' Shared objects in both MDB and ADP formats
                .Add New clsDbProject
                .Add New clsDbVbeProject
                .Add New clsDbVbeReference
                .Add New clsDbVbeForm
                .Add New clsDbProjProperty
                .Add New clsDbSavedSpec
                .Add New clsDbModule
                .Add New clsDbCommandBar
                If blnADP Then
                    ' Some types of objects only exist in ADP projects
                    .Add New clsAdpFunction
                    .Add New clsAdpServerView
                    .Add New clsAdpProcedure
                    .Add New clsAdpTable
                    .Add New clsAdpTrigger
                ElseIf blnMDB Then
                    ' These objects only exist in DAO databases
                    .Add New clsDbSharedImage
                    .Add New clsDbTheme
                    .Add New clsDbProperty
                    .Add New clsDbImexSpec
                    .Add New clsDbTableDef
                    .Add New clsDbQuery
                End If
                ' Additional objects to import after ADP/MDB specific items
                .Add New clsDbForm
                .Add New clsDbMacro
                .Add New clsDbReport
                .Add New clsDbTableData
                If blnMDB Then
                    .Add New clsDbTableDataMacro
                    .Add New clsDbRelation
                    .Add New clsDbDocument
                    .Add New clsDbNavPaneGroup
                    .Add New clsDbHiddenAttribute
                End If

            ' Process only items that may contain VBA code
            Case ecfVBAItems

                .Add New clsDbModule
                .Add New clsDbForm
                .Add New clsDbReport

        End Select
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetClassFromObject
' Author    : Adam Waller
' Date      : 2/22/2023
' Purpose   : Returns a database component class from a database object. (Used when
'           : exporting a single object.)
'           : Note that not all component types are supported.
'---------------------------------------------------------------------------------------
'
Public Function GetClassFromObject(objItem As AccessObject) As IDbComponent

    Dim cItem As IDbComponent

    ' Map to correct component class
    Select Case objItem.Type
        Case acForm:    Set cItem = New clsDbForm
        Case acMacro:   Set cItem = New clsDbMacro
        Case acModule:  Set cItem = New clsDbModule
        Case acQuery:   Set cItem = New clsDbQuery
        Case acReport:  Set cItem = New clsDbReport
        Case acTable:   Set cItem = New clsDbTableDef
        Case Else
            ' Not currently supported
    End Select

    ' Set database item and return class instance
    If Not cItem Is Nothing Then
        Set cItem.DbObject = objItem
        Set GetClassFromObject = cItem
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetComponentClass
' Author    : Adam Waller
' Date      : 6/3/2023
' Purpose   : Return an instance of the component class from the component type.
'---------------------------------------------------------------------------------------
'
Public Function GetComponentClass(intType As eDatabaseComponentType) As IDbComponent
    Dim cls As IDbComponent
    Select Case intType
        Case edbAdpFunction:        Set cls = New clsAdpFunction
        Case edbAdpServerView:      Set cls = New clsAdpServerView
        Case edbAdpStoredProcedure: Set cls = New clsAdpProcedure
        Case edbAdpTable:           Set cls = New clsAdpTable
        Case edbAdpTrigger:         Set cls = New clsAdpTrigger
        Case edbConnection:         Set cls = New clsDbConnection
        Case edbDbsProperty:        Set cls = New clsDbProperty
        Case edbDocument:           Set cls = New clsDbDocument
        Case edbForm:               Set cls = New clsDbForm
        Case edbHiddenAttribute:    Set cls = New clsDbHiddenAttribute
        Case edbImexSpec:           Set cls = New clsDbImexSpec
        Case edbMacro:              Set cls = New clsDbMacro
        Case edbModule:             Set cls = New clsDbModule
        Case edbNavPaneGroup:       Set cls = New clsDbNavPaneGroup
        Case edbProject:            Set cls = New clsDbProject
        Case edbProjectProperty:    Set cls = New clsDbProjProperty
        Case edbQuery:              Set cls = New clsDbQuery
        Case edbRelation:           Set cls = New clsDbRelation
        Case edbReport:             Set cls = New clsDbReport
        Case edbSavedSpec:          Set cls = New clsDbSavedSpec
        Case edbSharedImage:        Set cls = New clsDbSharedImage
        Case edbTableData:          Set cls = New clsDbTableData
        Case edbTableDataMacro:     Set cls = New clsDbTableDataMacro
        Case edbTableDef:           Set cls = New clsDbTableDef
        Case edbTheme:              Set cls = New clsDbTheme
        Case edbVbeForm:            Set cls = New clsDbVbeForm
        Case edbVbeProject:         Set cls = New clsDbVbeProject
        Case edbVbeReference:       Set cls = New clsDbVbeReference
        Case edbCommandBar:         Set cls = New clsDbCommandBar
        Case Else ' No match, return nothing
    End Select
    Set GetComponentClass = cls
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveComponentType
' Author    : Adam Waller
' Date      : 4/15/2026
' Purpose   : Map a user-supplied type string to an eDatabaseComponentType enum value.
'           : Accepts singular, plural, and common alias forms so that callers
'           : (MCP tools, direct API, etc.) don't need to know the canonical name.
'           : Returns -1 if the type string is not recognized.
'---------------------------------------------------------------------------------------
'
Public Function ResolveComponentType(strType As String) As eDatabaseComponentType
    ResolveComponentType = -1
    Select Case LCase$(strType)
        ' Core Access objects
        Case "query", "queries":                                                ResolveComponentType = edbQuery
        Case "form", "forms":                                                   ResolveComponentType = edbForm
        Case "report", "reports":                                               ResolveComponentType = edbReport
        Case "module", "modules":                                               ResolveComponentType = edbModule
        Case "table", "tables":                                                 ResolveComponentType = edbTableDef
        Case "macro", "macros":                                                 ResolveComponentType = edbMacro
        ' Extended multi-file types
        Case "table_data":                                                      ResolveComponentType = edbTableData
        Case "table_data_macro", "table_data_macros":                           ResolveComponentType = edbTableDataMacro
        Case "relation", "relations":                                           ResolveComponentType = edbRelation
        Case "saved_spec", "saved_specs":                                       ResolveComponentType = edbSavedSpec
        Case "imex_spec", "imex_specs", "import_spec", "import_specs":          ResolveComponentType = edbImexSpec
        Case "theme", "themes":                                                 ResolveComponentType = edbTheme
        Case "shared_image", "shared_images":                                   ResolveComponentType = edbSharedImage
        Case "vbe_form", "vbe_forms", "userform", "userforms":                  ResolveComponentType = edbVbeForm
        Case "command_bar", "command_bars", "commandbar", "commandbars":        ResolveComponentType = edbCommandBar
        ' Single-file types (object_name ignored by callers)
        Case "vbe_project", "vb_project":                                       ResolveComponentType = edbVbeProject
        Case "vbe_reference", "vbe_references":                                 ResolveComponentType = edbVbeReference
        Case "project":                                                         ResolveComponentType = edbProject
        Case "connection", "connections":                                        ResolveComponentType = edbConnection
        Case "db_property", "db_properties":                                    ResolveComponentType = edbDbsProperty
        Case "project_property", "project_properties":                          ResolveComponentType = edbProjectProperty
        Case "document", "documents":                                           ResolveComponentType = edbDocument
        Case "hidden_attribute", "hidden_attributes":                           ResolveComponentType = edbHiddenAttribute
        Case "nav_pane_group", "nav_pane_groups":                               ResolveComponentType = edbNavPaneGroup
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ContainerHasObject
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Returns true if the dictionary container has one or more of the specified
'           : object type within it. (This is useful when determining whether we need
'           : to run certain post-build operations.)
'---------------------------------------------------------------------------------------
'
Public Function ContainerHasObject(dContainer As Dictionary, intComponentType As eDatabaseComponentType) As Boolean

    Dim dCategory As Dictionary
    Dim varKey As Variant

    ' Get category (key) name of component type
    varKey = GetComponentClass(intComponentType).Category

    ' Check for any objects in this category
    If dContainer.Exists(varKey) Then
        Set dCategory = dContainer(varKey)
        If dCategory.Exists("Files") Then ContainerHasObject = (dCategory("Files").Count > 0)
        If dCategory.Exists("Objects") Then ContainerHasObject = (dCategory("Objects").Count > 0)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ContainerHasAnyObject
' Author    : Adam Waller
' Date      : 6/3/2023
' Purpose   : Return true if ANY of the specified objects types are found.
'---------------------------------------------------------------------------------------
'
Public Function ContainerHasAnyObject(dContainer As Dictionary, ParamArray intOtherTypes() As Variant) As Boolean

    Dim intType As Integer
    Dim cCategory As IDbComponent
    Dim dCategory As Dictionary
    Dim varKey As Variant
    Dim blnFound As Boolean

    ' Loop through types
    For intType = LBound(intOtherTypes) To UBound(intOtherTypes)
        ' Loop through containers
        For Each varKey In dContainer.Keys
            If TypeOf varKey Is IDbComponent Then
                Set cCategory = varKey
                ' Look for matching component type
                If cCategory.ComponentType = intOtherTypes(intType) Then
                    Set dCategory = dContainer(varKey)
                    If dCategory.Exists("Files") Then blnFound = (dCategory("Files").Count > 0)
                    If dCategory.Exists("Objects") Then blnFound = (dCategory("Objects").Count > 0)
                    If blnFound Then Exit For
                End If
            End If
        Next varKey
        If blnFound Then Exit For
    Next intType

    ' Return true if any matching object was found.
    ContainerHasAnyObject = blnFound

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeIfChanged
' Author    : Adam Waller
' Date      : 6/3/2023
' Purpose   : Merge the component type if the database object seems to be different
'           : than the source file. (For example, after merging a form, you may need
'           : to merge document properties to get the form description.)
'---------------------------------------------------------------------------------------
'
Public Sub MergeIfChanged(intComponentType As eDatabaseComponentType)

    Dim cComponent As IDbComponent
    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim cItem As IDbComponent

    ' Convert enum to component class
    Set cComponent = GetComponentClass(intComponentType)

    ' Check component items for changed database objects.
    With cComponent
        ' Get dictionary of modified items
        Set dItems = .GetAllFromDB(True)
        For Each varKey In dItems.Keys
            Set cItem = dItems(varKey)
            Log.Add "  " & FSO.GetFileName(cItem.SourceFile)
            cItem.Merge cItem.SourceFile
            If .SingleFile Then Exit For
        Next varKey
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetQuickObjectCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a quick, non-iterative object count for the collection of
'           : database components. (Used for progress bar)
'---------------------------------------------------------------------------------------
'
Public Function GetQuickObjectCount(colContainers As Collection) As Long

    Dim lngTotal As Long
    Dim cCont As IDbComponent

    Perf.OperationStart "Quick Count Objects"
    For Each cCont In colContainers
        lngTotal = lngTotal + cCont.QuickCount
    Next cCont
    Perf.OperationEnd

    GetQuickObjectCount = lngTotal

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetQuickFileCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a quick count of the files in each folder so we can increment
'           : the progress of scanning through files in a folder.
'---------------------------------------------------------------------------------------
'
Public Function GetQuickFileCount(colContainers As Collection) As Long

    Dim lngTotal As Long
    Dim strBase As String
    Dim strFolder As String
    Dim cCont As IDbComponent

    ' Get base folder path
    Perf.OperationStart "Quick Count Files"
    strBase = Options.GetExportFolder

    For Each cCont In colContainers
        strFolder = cCont.BaseFolder
        If StrComp(strBase, strFolder, vbTextCompare) = 0 Then
            ' Add a single count for the single file
            lngTotal = lngTotal + 1
        Else
            ' Make sure the folder actually exists before getting a file count
            If FSO.FolderExists(strFolder) Then
                lngTotal = lngTotal + CountFilesRecursive(FSO.GetFolder(strFolder))
            End If
        End If
    Next cCont
    Perf.OperationEnd

    ' Return total number of files in all source folders
    GetQuickFileCount = lngTotal

End Function


'---------------------------------------------------------------------------------------
' Procedure : CountFilesRecursive
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Return the total count of files in a folder and all subfolders.
'           : Used by GetQuickFileCount for @Folder annotation support.
'---------------------------------------------------------------------------------------
'
Private Function CountFilesRecursive(oFolder As Scripting.Folder) As Long

    Dim oSubFolder As Scripting.Folder

    CountFilesRecursive = oFolder.Files.Count
    For Each oSubFolder In oFolder.SubFolders
        CountFilesRecursive = CountFilesRecursive + CountFilesRecursive(oSubFolder)
    Next oSubFolder

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSourceModifiedDate
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Return the largest modified date for the source files associated with
'           : this component type.
'---------------------------------------------------------------------------------------
'
Public Function GetSourceModifiedDate(cmp As IDbComponent, Optional strFile As String) As Date

    Dim varExt As Variant
    Dim dteLatest As Date
    Dim strBaseFile As String

    ' Build base file path without extension
    If Len(strFile) Then
        ' Use provided file name first
        strBaseFile = FSO.GetBaseName(strFile)
    Else
        ' Otherwise use default source file name
        strBaseFile = FSO.GetBaseName(cmp.SourceFile)
    End If

    ' Check each possible file extension to find the most recent date
    For Each varExt In cmp.FileExtensions
        dteLatest = Largest(dteLatest, GetLastModifiedDate(strBaseFile & "." & varExt))
    Next varExt

    ' Return most recent date
    GetSourceModifiedDate = dteLatest

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLastModifiedSourceFile
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Return the path of the last modified source file for the specified
'           : component. (This may not be the primary source file used by the index.)
'---------------------------------------------------------------------------------------
'
Public Function GetLastModifiedSourceFile(cmp As IDbComponent, Optional strFile As String)

    Dim varExt As Variant
    Dim dteLatest As Date
    Dim dteTest As Date
    Dim strSourceFile As String
    Dim strLastModifiedFile As String
    Dim strBaseFile As String

    ' Build base file path without extension
    If Len(strFile) Then
        ' Use provided file name first
        strBaseFile = FSO.GetBaseName(strFile)
    Else
        ' Otherwise use default source file name
        strBaseFile = FSO.GetBaseName(cmp.SourceFile)
    End If

    ' Check each possible file extension to find the most recent date
    For Each varExt In cmp.FileExtensions
        strSourceFile = strBaseFile & "." & varExt
        dteTest = GetLastModifiedDate(strSourceFile)
        If dteTest > dteLatest Then
            dteLatest = dteTest
            strLastModifiedFile = strSourceFile
        End If
    Next varExt

    ' Return file path for most recently modified file
    GetLastModifiedSourceFile = strLastModifiedFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSourceFilesPropertyHash
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Return a has of the source file dates and sizes for this component type.
'           : (Useful for determining if any of the source files have been modified.)
'---------------------------------------------------------------------------------------
'
Public Function GetSourceFilesPropertyHash(cmp As IDbComponent, Optional strFile As String) As String

    Dim varExt As Variant
    Dim strSourceFile As String
    Dim strBaseFile As String
    Dim oFile As Scripting.File

    Perf.OperationStart "Get File Property Hash"

    ' Build base file path without extension
    If Len(strFile) Then
        ' Use provided file name first
        strBaseFile = FSO.BuildPath(FSO.GetParentFolderName(strFile), FSO.GetBaseName(strFile))
    Else
        ' Otherwise use default source file name
        strBaseFile = FSO.BuildPath(FSO.GetParentFolderName(cmp.SourceFile), FSO.GetBaseName(cmp.SourceFile))
    End If

    ' Build a combined string with all the properties
    With New clsConcat

        ' Check each possible file extension to find all the source files
        For Each varExt In cmp.FileExtensions
            strSourceFile = strBaseFile & "." & varExt
            If FSO.FileExists(strSourceFile) Then
                Set oFile = FSO.GetFile(strSourceFile)
                .Add oFile.DateLastModified, oFile.Size
            End If
        Next varExt

        ' Return hash of combined string
        GetSourceFilesPropertyHash = GetStringHash(.GetStr)
        Perf.OperationEnd
    End With

End Function
