Attribute VB_Name = "modVCSUtility"
'---------------------------------------------------------------------------------------
' Module    : modVCSUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Utility functions specific to the VCS project but not publicly exposed.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

' Control the interaction mode for the add-in
Public InteractionMode As eInteractionMode

Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetKeyboardState Lib "user32" (lppbKeyState As Any) As Long
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Any) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As LongPtr) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Private Const ModuleName = "modVCSUtility"


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
' Procedure : ContainerHasObject
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Returns true if the dictionary container has one or more of the specified
'           : object type within it. (This is useful when determining whether we need
'           : to run certain post-build operations.)
'---------------------------------------------------------------------------------------
'
Public Function ContainerHasObject(dContainer As Dictionary, intComponentType As eDatabaseComponentType) As Boolean

    Dim cCategory As IDbComponent
    Dim dCategory As Dictionary
    Dim varKey As Variant

    ' Loop through containers
    For Each varKey In dContainer.Keys
        If TypeOf varKey Is IDbComponent Then
            Set cCategory = varKey
            ' Look for matching component type
            If cCategory.ComponentType = intComponentType Then
                Set dCategory = dContainer(varKey)
                If dCategory.Exists("Files") Then ContainerHasObject = (dCategory("Files").Count > 0)
                If dCategory.Exists("Objects") Then ContainerHasObject = (dCategory("Objects").Count > 0)
                Exit For
            End If
        End If
    Next varKey

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
                ' Add a count of the files in the folder
                lngTotal = lngTotal + FSO.GetFolder(strFolder).Files.Count
            End If
        End If
    Next cCont
    Perf.OperationEnd

    ' Return total number of files in all source folders
    GetQuickFileCount = lngTotal

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


'---------------------------------------------------------------------------------------
' Procedure : GetVCSVersion
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Gets the version of the version control system. (Used to turn off fast
'           : save until a full export has been run with the current version of
'           : the MSAccessVCS addin.)
'---------------------------------------------------------------------------------------
'
Public Function GetVCSVersion() As String

    Dim dbs As Database
    Dim prp As DAO.Property

    Set dbs = CodeDb

    For Each prp In dbs.Properties
        If prp.Name = "AppVersion" Then
            ' Return version
            GetVCSVersion = prp.Value
        End If
    Next prp

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveComponentAsText
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Wrapper for Application.SaveAsText that verifies that the path exists,
'           : and then removes any existing file before saving the object as text.
'           : Returns a hash of the file content (if applicable) to track changes.
'---------------------------------------------------------------------------------------
'
Public Function SaveComponentAsText(intType As AcObjectType, _
                                strName As String, _
                                strFile As String, _
                                Optional cDbObjectClass As IDbComponent = Nothing) As String

    Dim strTempFile As String
    Dim strAltFile As String
    Dim strContent As String
    Dim strPrintSettingsFile As String
    Dim strHash As String
    Dim cParser As clsSourceParser

    On Error GoTo ErrHandler

    ' Export to temporary file
    strTempFile = GetTempFile
    Perf.OperationStart "App.SaveAsText()"
    Application.SaveAsText intType, strName, strTempFile
    Perf.OperationEnd
    VerifyPath strFile

    ' Delete any existing source file
    If FSO.FileExists(strFile) Then DeleteFile strFile

    ' Sanitize certain object types
    Set cParser = New clsSourceParser
    Select Case intType
        Case acForm, acReport

            ' Load content from file
            strContent = ReadSourceFile(strTempFile)

            ' Process any saved devmode settings
            With New clsDevMode
                ' Build print settings file name.
                strPrintSettingsFile = SwapExtension(strFile, "json")
                ' See if we are exporting print vars.
                If Options.SavePrintVars = True Then
                    ' Grab the printer settings before sanitizing the file.
                    .LoadFromExportFile strContent
                    ' Only need to save print settings if they are different
                    ' from the default printer settings.
                    If (.GetHash <> VCSIndex.DefaultDevModeHash) And .HasData Then
                        WriteFile BuildJsonFile(TypeName(cDbObjectClass), .GetDictionary, _
                          strName & " Print Settings"), strPrintSettingsFile
                    Else
                        ' No print settings in this object.
                        If FSO.FileExists(strPrintSettingsFile) Then DeleteFile strPrintSettingsFile
                    End If
                Else
                    ' Remove any existing (now orphaned) print settings file.
                    If FSO.FileExists(strPrintSettingsFile) Then DeleteFile strPrintSettingsFile
                End If
            End With

            ' Sanitizing converts to UTF-8
            With cParser
                .LoadString strContent
                .ObjectName = FSO.GetBaseName(strFile)
                WriteFile .Sanitize(ectObjectDefinition), strFile
                strHash = .Hash

                ' Process any VBA
                strAltFile = SwapExtension(strFile, "cls")
                If Options.SplitLayoutFromVBA And Len(.GetObjectVBA) Then
                    ' Write VBA code as separate .cls file.
                    WriteFile .GetObjectVBA, strAltFile
                Else
                    ' Remove any split VBA file
                    If FSO.FileExists(strAltFile) Then DeleteFile strAltFile
                End If
            End With

        Case acQuery, acMacro
            ' Sanitizing converts to UTF-8
            With cParser
                .LoadSourceFile strTempFile
                WriteFile .Sanitize(ectObjectDefinition), strFile
                strHash = .Hash
            End With

        ' Case acModule - Use VBE export instead.

        Case acTableDataMacro
            ' Table data macros are stored in XML format
            ' The file may not exist if no TD Macro was found
            If FSO.FileExists(strTempFile) Then
                With cParser
                    .LoadSourceFile strTempFile
                    WriteFile .Sanitize(ectXML), strFile
                    strHash = .Hash
                End With
            End If

        Case Else
            ' Handle UCS conversion if needed
            ConvertUcs2Utf8 strTempFile, strFile

    End Select

    ' Remove any leftover temp file.
    If FSO.FileExists(strTempFile) Then DeleteFile strTempFile

    ' Normal exit
    On Error GoTo 0

    ' Return content hash
    SaveComponentAsText = strHash
    Exit Function

ErrHandler:
    If Err.Number = 2950 And intType = acTableDataMacro Then
        ' This table apparently didn't have a Table Data Macro.
        Exit Function
    Else
        ' Some other error.
        Err.Raise Err.Number
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadComponentFromText
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Load the object into the database from the saved source file.
'---------------------------------------------------------------------------------------
'
Public Sub LoadComponentFromText(intType As AcObjectType, strName As String, strFile As String)

    Dim strTempFile As String
    Dim strSourceFile As String
    Dim strPrefix As String
    Dim strAltFile As String
    Dim strContent As String
    Dim blnVbaOverlay As Boolean
    Dim blnConvert As Boolean

    ' In most cases we are importing/converting the actual source file.
    strSourceFile = strFile

    ' Add DevMode structures and VBA code back into forms/reports
    Select Case intType
        Case acForm, acReport

            ' Read file content. (Should be UTF-8)
            strContent = ReadFile(strFile)
            With New clsSourceParser
                .LoadString strContent

                ' Check for print settings file
                strAltFile = SwapExtension(strFile, "json")
                If FSO.FileExists(strAltFile) Then
                    ' Merge the print settings into the source file content
                    .MergePrintSettings ReadFile(strAltFile)
                End If

                ' For forms and reports, check for VBA code file that needs to be merged
                strAltFile = SwapExtension(strFile, "cls")
                If FSO.FileExists(strAltFile) Then
                    ' Found a companion class file.
                    .MergeVBA ReadFile(strAltFile)
                    blnVbaOverlay = RequiresOverlay(.GetObjectVBA)
                End If

                ' Write ouput to a new file if anything has changed
                If .OutputModified Then
                    strSourceFile = GetTempFile
                    WriteFile .GetOutput, strSourceFile
                End If

            End With

    End Select

    ' Check UCS-2-LE requirement for the current database.
    ' (Cached after first call)
    Select Case intType
        Case acForm, acReport, acQuery, acMacro, acTableDataMacro
            blnConvert = RequiresUcs2
        Case acModule
            ' Always convert from UTF-8 in case the file contains
            ' UTF-8 encoded characters but does not have a BOM.
            blnConvert = True
    End Select

    ' Only run conversion if needed.
    If blnConvert Then
        ' Perform file conversion, and import from temp file.
        strTempFile = GetTempFile
        If intType = acModule Then
            ' Convert back to ANSI for VBA modules
            ConvertUtf8Ansi strSourceFile, strTempFile, False
        Else
            ' Other objects converted to UCS2
            ConvertUtf8Ucs2 strSourceFile, strTempFile, False
        End If
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strTempFile
        Perf.OperationEnd
        DeleteFile strTempFile, True
    Else
        ' Load UTF-8 file
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strSourceFile
        Perf.OperationEnd
    End If

    ' Clean up any additional temp file used in the building process
    If strFile <> strSourceFile Then
        If FSO.FileExists(strSourceFile) Then DeleteFile strSourceFile
    End If

    ' Check for VBA overlay
    If blnVbaOverlay Then
        strPrefix = IIf(intType = acForm, "Form_", "Report_")
        OverlayCodeModule strPrefix & strName, SwapExtension(strFile, "cls")
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVbComponent
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Export the code module VB component and convert to UTF-8
'---------------------------------------------------------------------------------------
'
Public Sub ExportCodeModule(strName As String, strFile As String)

    Dim strTempFile As String
    Dim strContent As String

    Perf.OperationStart "Export VBE Module"

    ' Export to a temp file so we can convert to UTF-8 encoding
    strTempFile = GetTempFile
    CurrentVBProject.VBComponents(strName).Export strTempFile

    ' Sanitize the VBA code while reading the temp file
    With New clsSourceParser
        .LoadString ReadFile(strTempFile, GetSystemEncoding)
        strContent = .Sanitize(ectVBA)
    End With

    ' Write the content as UTF-8 to the final destination
    WriteFile strContent, strFile
    DeleteFile strTempFile

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OverlayCodeModule
' Author    : Adam Waller
' Date      : 10/24/2023
' Purpose   : Overlay VBA code from an object's *.cls file to the form or report
'           : Note that this opens the object in design view, which may slow the build
'           : process if a large number of items are invovled.
'---------------------------------------------------------------------------------------
'
Public Sub OverlayCodeModule(strName As String, strClassFile As String)

    Dim objModule As VBIDE.CodeModule
    Dim strContent As String
    Dim intType As AcObjectType
    Dim strShortName As String
    Dim cParser As clsSourceParser

    LogUnhandledErrors
    'On Error Resume Next
    Set objModule = CurrentVBProject.VBComponents(strName).CodeModule
    If CatchAny(eelError, "Could not find code module for " & strName, ModuleName & ".OverlayCodeModule") Then Exit Sub

    ' Read class file content
    strContent = ReadFile(strClassFile)
    If strContent = vbNullString Then
        Log.Error eelError, "Unable to read " & strClassFile, ModuleName & ".OverlayCodeModule"
        Exit Sub
    End If

    ' Get object type and short name
    If strName Like "Form_*" Then
        intType = acForm
        strShortName = Mid$(strName, 6)
        DoCmd.OpenForm strShortName, acDesign, , , , acHidden
    ElseIf strName Like "Report_*" Then
        intType = acReport
        strShortName = Mid$(strName, 8)
        DoCmd.OpenReport strShortName, acViewDesign, , , acHidden
    End If

    ' Overlay the VBA code, replacing any existing code.
    Set cParser = New clsSourceParser
    objModule.DeleteLines 1, objModule.CountOfLines
    objModule.AddFromString cParser.StripClassHeader(strContent, False)

    ' Close any form or report object
    Select Case intType
        Case acForm, acReport
            DoCmd.Close intType, strShortName, acSaveYes
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RequiresOverlay
' Author    : Adam Waller
' Date      : 11/2/2023
' Purpose   : Returns true if we need to overlay the VBA code through VBE for a form
'           : or report object.
'---------------------------------------------------------------------------------------
'
Private Function RequiresOverlay(strVbaCode As String) As Boolean
    If modEncoding.GetSystemEncoding(True) = "utf-8" Then
        RequiresOverlay = StringHasExtendedASCII(strVbaCode)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveNonBuiltInReferences
' Author    : Adam Waller
' Date      : 10/20/2020
' Purpose   : Remove any references that are not built-in. (Sometimes additional
'           : references are added when creating a new database, but not not really
'           : needed in the completed database when building the project from source.)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveNonBuiltInReferences()

    Dim intCnt As Integer
    Dim strName As String
    Dim ref As Access.Reference

    Perf.OperationStart "Clear References"
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If Not ref.BuiltIn Then
            strName = ref.Name
            Application.References.Remove ref
            Log.Add "  Removed " & strName, False
        End If
        Set ref = Nothing
    Next intCnt
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetOriginalDbFullPathFromSource
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Determine the original full path of the database, based on the files
'           : in the source folder. (Assumes that options have been loaded)
'---------------------------------------------------------------------------------------
'
Public Function GetOriginalDbFullPathFromSource(strFolder As String) As String

    Dim strPath As String
    Dim dContents As Dictionary
    Dim strFile As String
    Dim strExportFolder As String
    Dim lngLevel As Long

    strPath = FSO.BuildPath(strFolder, "vbe-project.json")
    If Not FSO.FileExists(strPath) Then
        Log.Error eelCritical, "Unable to find source file: " & strPath, "GetOriginalDbFullPathFromSource"
        GetOriginalDbFullPathFromSource = vbNullString
    Else
        ' Look up file name from VBE project file name
        Set dContents = ReadJsonFile(strPath)
        strFile = dNZ(dContents, "Items\FileName")

        ' Convert legacy relative path
        If Left$(strFile, 4) = "rel:" Then strFile = Mid$(strFile, 5)

        ' Trim off any tailing slash
        strExportFolder = StripSlash(strFolder)

        ' Check export folder settings
        If Options.ExportFolder = vbNullString Then
            ' Default setting, using parent folder of source directory
            GetOriginalDbFullPathFromSource = strExportFolder & PathSep & ".." & PathSep & strFile
        Else
            ' Check to see if we are using an absolute export path  (\\* or *:*)
            If StartsWith(Options.ExportFolder, PathSep & PathSep) _
                Or (InStr(2, Options.ExportFolder, ":") > 0) Then
                ' Look for saved build path
                Set dContents = ReadJsonFile(FSO.BuildPath(strFolder, "proj-properties.json"))
                strPath = dNZ(dContents, "Items\VCS Build Path")
                If strPath <> vbNullString Then
                    GetOriginalDbFullPathFromSource = strPath & PathSep & strFile
                Else
                    ' Unable to determine the original file location.
                    Exit Function
                End If
            Else
                ' Calculate how many levels deep to create original path
                lngLevel = UBound(Split(StripSlash(Options.ExportFolder), PathSep))
                If lngLevel < 0 Then lngLevel = 0   ' Handle "\" to export in current folder.
                GetOriginalDbFullPathFromSource = strExportFolder & PathSep & _
                    Repeat(".." & PathSep, lngLevel) & strFile
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderHasVcsOptionsFile
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Returns true if the folder as a vcs-options.json file, which is required
'           : to build a project from source files.
'---------------------------------------------------------------------------------------
'
Public Function FolderHasVcsOptionsFile(strFolder As String) As Boolean
    FolderHasVcsOptionsFile = FSO.FileExists(FSO.BuildPath(strFolder, "vcs-options.json"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildJsonFile
' Author    : Adam Waller
' Date      : 2/5/2022
' Purpose   : Creates json file content with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'           : Set the file format version only when the dictionary structure changes
'           : with potentially breaking changes for prior versions.
'---------------------------------------------------------------------------------------
'
Public Function BuildJsonFile(strClassName As String, dItems As Dictionary, strDescription As String, _
    Optional dblExportFormatVersion As Double) As String

    Dim dContents As Dictionary
    Dim dHeader As Dictionary

    Set dContents = New Dictionary
    Set dHeader = New Dictionary

    ' Build dictionary structure
    dHeader.Add "Class", strClassName
    dHeader.Add "Description", strDescription
    If dblExportFormatVersion <> 0 Then dHeader.Add "Export File Format", dblExportFormatVersion
    dContents.Add "Info", dHeader
    dContents.Add "Items", dItems

    ' Return assembled content in Json format
    BuildJsonFile = ConvertToJson(dContents, JSON_WHITESPACE)

End Function


'---------------------------------------------------------------------------------------
' Procedure : CompileAndSaveAllModules
' Author    : Adam Waller
' Date      : 7/10/2021
' Purpose   : Compile and save the modules in the current database
'---------------------------------------------------------------------------------------
'
Public Sub CompileAndSaveAllModules()
    Perf.OperationStart "Compile/Save Modules"
    ' Make sure we are running this in the CurrentDB, not the CodeDB
    Set VBE.ActiveVBProject = CurrentVBProject
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    DoEvents
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    For Each oProj In VBE.VBProjects
        If StrComp(oProj.FileName, GetAddInFileName, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Public Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    LogUnhandledErrors
    On Error Resume Next
    Application.Run GetAddInFileName & "!DummyFunction"
    If Err Then Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AfterBuild
' Author    : Adam Waller
' Date      : 12/18/2023
' Purpose   : Run this code after building the add-in from source.
'---------------------------------------------------------------------------------------
'
Public Sub AfterBuild()
    modResource.VerifyResources
    Translation.LoadTranslations
    ImportCommandBarsTemplate
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckGitFiles
' Author    : Adam Waller
' Date      : 5/23/2022
' Purpose   : If this project appears to be a git repository, this checks to see if
'           : it contains a .gitignore and .gitattributes file. If it doesn't, then
'           : the default files are extracted and added to the project, and the user
'           : notified that these have been added.
'           : Checks both the export folder and the current folder.
'---------------------------------------------------------------------------------------
'
Public Sub CheckGitFiles()

    Dim strPath As String
    Dim strFile As String
    Dim blnAdded As Boolean

    ' Check export folder
    strPath = Options.GetExportFolder
    If Not FSO.FolderExists(strPath & ".git") Then
        ' Check current folder for repository root
        ' (This would be the default usage)
        strPath = CurrentProject.Path & PathSep
        If Not FSO.FolderExists(strPath & ".git") Then
            ' No git folder found.
            Exit Sub
        End If
    End If

    ' gitignore file
    strFile = strPath & ".gitignore"
    If Not FSO.FileExists(strFile) Then
        ExtractResource "Default .gitignore", strPath
        Name strFile & ".default" As strFile
        Log.Add "Added default .gitignore file", , , "blue"
        blnAdded = True
    End If

    ' gitattributes file
    strFile = strPath & ".gitattributes"
    If Not FSO.FileExists(strFile) Then
        ExtractResource "Default .gitattributes", strPath
        Name strFile & ".default" As strFile
        Log.Add "Added default .gitattributes file", , , "blue"
        blnAdded = True
    End If

    ' Notify user
    If blnAdded Then MsgBox2 "Added Default Git File(s)", _
        "Added a default .gitignore and/or .gitattributes file to your project.", _
        "By default these files exclude the binary database files from version control," & vbCrLf & _
        "allowing you to track changes at the source file level." & vbCrLf & vbCrLf & _
        "You may wish to customize these further for your environment.", vbInformation

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ShiftOpenDatabase
' Author    : Adam Waller
' Date      : 2/25/2022
' Purpose   : Open a database with the shift key held down so we can (hopefully)
'           : bypass the startup code.
'---------------------------------------------------------------------------------------
'
Public Sub ShiftOpenDatabase(strPath As String, Optional blnExclusive As Boolean = False)

    Const VK_SHIFT = &H10

    ' Skip open if we are already on the correct database
    If CurrentProject.FullName = strPath And Not blnExclusive Then Exit Sub

    ' Close any open database before we try to open another one.
    If DatabaseFileOpen Then
        StageMainForm
        CloseCurrentDatabase2
        DoCmd.OpenForm "frmVCSMain", , , , , acHidden
        RestoreMainForm
    End If

    On Error GoTo Error_Handler

    Dim abytCodesSrc(0 To 255) As Byte
    Dim abytCodesDest(0 To 255) As Byte

    If (FSO.FileExists(strPath) = False) Then
        Err.Raise 53
    End If

    SetForegroundWindow Application.hWndAccessApp
    SetFocus Application.hWndAccessApp

    ' Set Shift state
    GetKeyboardState abytCodesSrc(0)
    GetKeyboardState abytCodesDest(0)
    abytCodesDest(VK_SHIFT) = 128
    SetKeyboardState abytCodesDest(0)

    ' Open the database with shift key down
    Application.OpenCurrentDatabase strPath, blnExclusive

    ' Revert back keyboard state and restore focus
    SetKeyboardState abytCodesSrc(0)
    SetForegroundWindow Application.hWndAccessApp
    SetFocus Application.hWndAccessApp

    Exit Sub

Error_Handler:
    SetForegroundWindow Application.hWndAccessApp

    With Err
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSchemaParams
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Return the schema initialization parameters for dependency injection.
'---------------------------------------------------------------------------------------
'
Public Function GetSchemaInitParams(strName As String) As Dictionary

    Dim dParams As Dictionary
    Dim strFile As String

    ' Load parameters for initializing the connection
    If Options.SchemaExports.Exists(strName) Then
        Set dParams = CloneDictionary(Options.SchemaExports(strName))
    Else
        ' Could be a new schema not yet saved
        Set dParams = New Dictionary
        dParams.CompareMode = TextCompare
    End If
    dParams("Name") = strName

    ' Check for `Connect` or other parameters in .env file
    strFile = BuildPath2(Options.GetExportFolder & "databases", GetSafeFileName(strName), ".env")
    If FSO.FileExists(strFile) Then
        With New clsDotEnv
            .LoadFromFile strFile
            .MergeIntoDictionary dParams, False
        End With
    End If

    ' Return initialization parameters
    Set GetSchemaInitParams = dParams

End Function


'---------------------------------------------------------------------------------------
' Procedure : PassesSchemaFilter
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Returns true if this item passed any user-defined filter rules.
'           : The current implementation processes rules sequentially, applying each
'           : rule in order. Last matching rule will apply to the object.
'---------------------------------------------------------------------------------------
'
Public Function PassesSchemaFilter(strItem As String, colFilters As Collection) As Boolean

    Dim blnPass As Boolean
    Dim varRule As Variant
    Dim strRule As String

    If colFilters Is Nothing Then
        blnPass = True
    ElseIf colFilters.Count = 0 Then
        blnPass = True
    Else
        ' Loop through rules
        For Each varRule In colFilters
            strRule = CStr(varRule)
            Select Case Left$(strRule, 1)
                Case "#", vbNullString
                    ' Ignore comments and blank lines
                Case "!"
                    ' Negative rule (do not include)
                    If strItem Like Mid$(strRule, 2) Then blnPass = False
                Case Else
                    ' Positive rule
                    If strItem Like strRule Then blnPass = True
            End Select
        Next varRule
    End If

    ' Return final result
    PassesSchemaFilter = blnPass

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadSourceFile
' Author    : Adam Waller
' Date      : 11/8/2023
' Purpose   : Load source file content into a string. (Considers BOM and file type)
'---------------------------------------------------------------------------------------
'
Public Function ReadSourceFile(strPath As String) As String

    Dim strTempFile As String
    Dim strExt As String

    ' Read text from file, and split into lines
    If HasUcs2Bom(strPath) Then
        ReadSourceFile = ReadFile(strPath, "Unicode")
    Else
        ' ADP projects may contain mixed Unicode content
        If CurrentProject.ProjectType = acADP Then
            strTempFile = GetTempFile
            ConvertUcs2Utf8 strPath, strTempFile, False
            ReadSourceFile = ReadFile(strTempFile)
            DeleteFile strTempFile
        Else
            strExt = LCase(Right(strPath, 4))
            ' ExportXML defaults to UTF-8 regardless of DB version
            If DbVersion <= 4 And strExt <> ".xml" Then
                ' Access 2000 format exports using system codepage
                ' See issue #217
                ReadSourceFile = ReadFile(strPath, GetSystemEncoding)
            Else
                ' Newer versions export as UTF-8
                ReadSourceFile = ReadFile(strPath)
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportCommandBarsTemplate
' Author    : bclothier
' Date      : 02/14/2025
' Purpose   : Import the template command bar from the template file for the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub ImportCommandBarsTemplate()
    Dim strTemplatePath As String

    strTemplatePath = BuildPath2(CurrentProject.Path, "\Template\CommandBars.bin")

    If FSO.FileExists(strTemplatePath) Then
        Select Case ImportCommandBars(strTemplatePath, strTemplateCommandBarName)
            Case eicImportedVerified
                ' All good
            Case eicImportedUnableToVerify
                Log.Error eelWarning, "Template command bar was imported but we cannot verify if it was imported successfully."
            Case eicImportedNotVerified
                Log.Error eelError, "Template command bar was imported  but we didn't find it in the built file. This indicates something went wrong with the import."
            Case Else
                Log.Error eelError, "Template command bar was not imported!"
        End Select
    Else
        MsgBox2 "Unable to import the template command bar", "The add-in could not locate the '\Template\CommandBars.bin' in the repository which is required for the add-in to function correctly. Ensure that you have pulled the latest from the git repository and the file is present before building the add-in.", , vbCritical, "Error importing command bar template"
        Log.Error eelCritical, "Template command bar could not be imported because the source file is missing."
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportCommandBars
' Author    : bclothier
' Date      : 02/14/2025
' Purpose   : Import the command bars from a source database to the specified project.
'           : By default the current Application is used but it can be another instance
'           : (e.g., an automated Access.Appplication instance.)
'           : A zero or a negative return indicates an error with the import.
'---------------------------------------------------------------------------------------
'
Public Function ImportCommandBars(strSourceDatabasePath As String, strCommandBarNameToVerify As String, Optional objTargetApplication As Application = Nothing) As eImportCommandBarsResult
    Dim strSql As String
    Dim blnResult As Boolean

    If objTargetApplication Is Nothing Then
        Set objTargetApplication = Application
    End If

    With objTargetApplication
        ' If we do not delete the command bar with the same name, it will not import. Deleting it from
        ' application will not actually delete it from its original database so even if we delete some
        ' another database's command bar, it won't actually remove it from that database and it'll be
        ' restored next time it is opened.
        On Error Resume Next
        Do
            .CommandBars(strCommandBarNameToVerify).Delete
        Loop Until Err.Number
        On Error GoTo 0

        ' Note that we are manipulating the application's WizHook which might not be necessarily
        ' the same one in modWizHook.
        .WizHook.Key = 51488399
        .WizHook.WizCopyCmdbars strSourceDatabasePath

        ' Verify we have the command bar imported.
        On Error Resume Next
        ' Application.CommandBars is the union of all loaded databases' command bars; just because we can find a
        ' command bar with same name does not mean the database has the command bar loaded into the binary file.
        ' However, this is a good first step in verifying the import since a negative result definitely mean it
        ' wasn't imported at all. We use the target project's Application in case it's an automated instance
        ' independent of the current Application object.
        blnResult = Not (.CommandBars(strCommandBarNameToVerify) Is Nothing)
        If Err.Number Then
            ImportCommandBars = eicFailed
            Exit Function
        End If
        On Error GoTo 0

        If blnResult Then
            blnResult = False
            ' Not all versions of MDB files or ADP files have MSysAccessStorage table.
            strSql = _
                "SELECT o.Name " & _
                "FROM MSysObjects AS o " & _
                "WHERE o.Name = 'MSysAccessStorage' " & _
                "  AND o.Type = 1;"
            With .CurrentProject.Connection.Execute(strSql)
                If Not .EOF Then
                    If .Fields(0).Value = "MSysAccessStorage" Then
                        blnResult = True
                    End If
                End If
            End With
        End If

        If blnResult Then
            ' This project has MSysAccessStorage table so we can determine if it contains the commandbar.
            ' We only need to check the virtual directory for the CmdBars entry. If the virtual directory
            ' has the command bar's name in its listing, we can assume the command bar was succcessfully
            ' imported into this specific database file. The directory entry is delimited as following:
            '   * Chr(4)
            '   * <byte length of the command bar's name> plus 4 more for the ending delimiters
            '   * the command bar's name (Unicode)
            '   * 4 null bytes (or 2 vbNullChars)
            ' Because a string reads in little endian order, we need to swap the Chr(4) and the byte length,
            ' so it becomes ChrW(Hex(((<length> + 4) * 256) + 4).
            strSql = _
                "SELECT s1.Lv " & _
                "FROM MSysAccessStorage AS s1 " & _
                "WHERE s1.Name = (Chr(3) & 'DirData') " & _
                "  AND s1.Type = 2 " & _
                "  AND s1.ParentId = (" & _
                "    SELECT s2.Id " & _
                "    FROM MSysAccessStorage AS s2 " & _
                "    WHERE s2.Name = 'CmdBars' " & _
                "      AND s2.Type = 1 " & _
                ");"

            With .CurrentProject.Connection.Execute(strSql)
                If Not .EOF Then
                    If InStr(1, .Fields(0).Value, ChrW(((LenB(strCommandBarNameToVerify) + 4) * 256) + 4) & strCommandBarNameToVerify & vbNullChar & vbNullChar, vbTextCompare) > 0 Then
                        ImportCommandBars = eicImportedVerified
                    Else
                        ImportCommandBars = eicImportedNotVerified
                    End If
                End If
            End With
        Else
            ' We can only tenatively assume success since we don't have the MSysAccessStorage table that can be
            ' easily inspected. Older MDB files use MSysAccessObjects which are even more opaque. No clue how
            ' we'd inspect an ADP file since it won't have any system tables to describe its contents.
            ImportCommandBars = eicImportedUnableToVerify
        End If
    End With
End Function
