Attribute VB_Name = "modLoadSaveText"
'---------------------------------------------------------------------------------------
' Module    : modLoadSaveText
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Wrappers around Access Application.SaveAsText and LoadFromText with
'           : sanitization, encoding conversion, print settings, and VBA overlay.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modFileAccess, modEncoding, modLoadFromText,
'           : modErrorHandling, modVbeUtility, modFunctions
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName = "modLoadSaveText"


'---------------------------------------------------------------------------------------
' Procedure : SaveComponentAsText
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Wrapper for Application.SaveAsText that verifies that the path exists,
'           : and then removes any existing file before saving the object as text.
'           : Returns a hash of the file content (if applicable) to track changes.
'---------------------------------------------------------------------------------------
'
Public Function SaveComponentAsText(intType As AcObjectType _
                                    , strName As String _
                                    , strFile As String _
                                    , cDbObjectClass As IDbComponent) As String

    Const FunctionName As String = ModuleName & ".SaveComponentAsText"

    Dim strTempFile As String
    Dim strAltFile As String
    Dim strContent As String
    Dim strPrintSettingsFile As String
    Dim strHash As String
    Dim cParser As clsSourceParser

    LogUnhandledErrors FunctionName
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
            With New clsPrinterSettings
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
                .LoadString strContent, intType
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

                ' Extract per-control conditional formatting into the companion JSON.
                ' (Merges with any existing print settings / metadata in the same file.)
                If Options.DecodeConditionalFormatting _
                    And Options.ExportFormatVersion >= EFV_5_0_0 Then
                    WriteConditionalFormatting strPrintSettingsFile, _
                        .GetConditionalFormats, strName, TypeName(cDbObjectClass)
                End If
            End With

        Case acQuery, acMacro
            ' Sanitizing converts to UTF-8
            With cParser
                .LoadSourceFile strTempFile, intType
                WriteFile .Sanitize(ectObjectDefinition), strFile
                strHash = .Hash
            End With

        ' Case acModule - Use VBE export instead.

        Case acTableDataMacro
            ' Table data macros are stored in XML format
            ' The file may not exist if no TD Macro was found
            If FSO.FileExists(strTempFile) Then
                With cParser
                    .LoadSourceFile strTempFile, intType
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
    If Catch(2950) And intType = acTableDataMacro Then
        ' This table apparently didn't have a Table Data Macro.
        Exit Function
    Else
        ' Some other error.
        Log.Error eelError, "Issue creating output file.", FunctionName
        'Err.Raise Err.Number
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadComponentFromText
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Load the object into the database from the saved source file.
'           : Returns True if the loading worked; False if an error occured or other
'           : issue was detected.
'---------------------------------------------------------------------------------------
'
Public Function LoadComponentFromText(intType As AcObjectType _
                                    , ByRef strName As String _
                                    , ByRef strFile As String _
                                    , Optional blnSuppressError As Boolean = False _
                                    , Optional strSourceDisplayFile As String = vbNullString) As Boolean

    Const FunctionName As String = ModuleName & ".LoadComponentFromText"

    Dim blnErrInFunction As Boolean
    Dim strTempFile As String
    Dim strSourceFile As String
    Dim strPrefix As String
    Dim strAltFile As String
    Dim strContent As String
    Dim strJsonContent As String
    Dim blnVbaOverlay As Boolean
    Dim blnConvert As Boolean

    If Len(strSourceDisplayFile) = 0 Then strSourceDisplayFile = strFile

    LogUnhandledErrors FunctionName
    On Error GoTo ErrHandler
    Perf.OperationStart FunctionName

RetryImport:
    ' In most cases we are importing/converting the actual source file.
    strSourceFile = strFile

    ' Add DevMode structures and VBA code back into forms/reports
    Select Case intType
        Case acForm, acReport

            ' Read file content. (Should be UTF-8)
            strContent = ReadFile(strFile)
            With New clsSourceParser
                .LoadString strContent, intType

                ' Check for companion JSON (print settings and conditional formatting)
                strAltFile = SwapExtension(strFile, "json")
                If FSO.FileExists(strAltFile) Then
                    strJsonContent = ReadFile(strAltFile)
                    ' Merge the print settings into the source file content
                    .MergePrintSettings strJsonContent
                    ' Rebuild and reinsert per-control conditional formatting blocks
                    .MergeConditionalFormat strJsonContent
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

        Case acQuery
            ' Resolve any env: references in pass-through query connection strings
            strContent = ReadFile(strFile)
            If InStr(1, strContent, "env:conn_", vbTextCompare) > 0 Then
                strContent = ResolveEnvReferencesInText(strContent)
                strSourceFile = GetTempFile
                WriteFile strContent, strSourceFile
            End If

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
        Perf.OperationStart "modLoadFromText.LoadFromText"
        modLoadFromText.LoadFromText intType, strName, strTempFile
        Perf.OperationEnd
        DeleteFile strTempFile, True

    Else
        ' Load UTF-8 file
        Perf.OperationStart "modLoadFromText.LoadFromText"
        modLoadFromText.LoadFromText intType, strName, strSourceFile
        Perf.OperationEnd
    End If

CleanUp:
    If Perf.CurrentOperationName = "modLoadFromText.LoadFromText" Then
        Perf.OperationEnd
    End If

    ' Clean up any additional temp file used in the building process
    If strFile <> strSourceFile Then
        If FSO.FileExists(strSourceFile) Then DeleteFile strSourceFile
    End If

    ' Check for VBA overlay
    If blnVbaOverlay And Not Operation.ErrorLevel = eelCritical Then ' don't do this if we're trying to bail out.
        strPrefix = IIf(intType = acForm, "Form_", "Report_")
        OverlayCodeModule strPrefix & strName, SwapExtension(strFile, "cls")
    End If

Exit_Here:
    ' Only set output to true when import and function didn't have any issues.
    LoadComponentFromText = (Not blnErrInFunction) And (Not Operation.ErrorLevel = eelCritical)
    Perf.OperationEnd
    Exit Function

ErrHandler:
    Dim strErrDescription As String
    strErrDescription = Err.Description
    Err.Clear ' FIXME: A temporary hack to avoid spurious output in logs when calling Log.Error

    If blnSuppressError Then
        ' Generate warning entries for suppressed errors
        Log.Error eelWarning, T("Import issue with '{0}'; {1}", var0:=strName, var1:=strErrDescription), FunctionName
        blnErrInFunction = True
        Resume CleanUp
    End If

    ' Log import details to the log file only; console gets one summary on Ignore.
    Log.Add T("Import issue with '{0}'; {1}", var0:=strName, var1:=strErrDescription), False
    Log.Add T("Source file: {0}", var0:=strSourceDisplayFile), False

    Select Case MsgBox2(T("Could not import '{0}'.", var0:=strName) _
            , T("Abort build, retry importing, or skip?") _
            , T("[Abort] = Abort build process entirely." & vbNewLine & _
                "[Retry] = Retry importing the item." & vbNewLine & _
                "[Ignore] = Skip this item.") _
            , vbAbortRetryIgnore, "Error importing!", vbAbort)

    Case vbAbort
        Log.Error eelCritical, "Aborted build.", FunctionName
        blnErrInFunction = True
        Resume CleanUp

    Case vbRetry
        Log.Add T("Retrying import for: {0}", var0:=strName)
        Resume RetryImport

    Case Else ' this also includes ignore.
        Log.Add T("Skipping import of '{0}'. Your application may not run or complile.", var0:=strName), False
        Log.Error eelError, T("Failed to import '{0}'.", var0:=strName), FunctionName
        blnErrInFunction = True
        Resume CleanUp

    End Select

End Function


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
' Procedure : GetMetadataHash
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Returns a lightweight hash of the Description property and Hidden
'           : attribute for a database object. Used by IsModified to detect metadata
'           : changes that do not update the object's DateModified timestamp.
'---------------------------------------------------------------------------------------
'
Public Function GetMetadataHash(strContainerName As String, _
    strObjectName As String, intObjType As AcObjectType) As String

    Dim dbs As Database
    Dim doc As DAO.Document
    Dim strDesc As String
    Dim blnHidden As Boolean

    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Function

    Set dbs = SharedDb

    LogUnhandledErrors
    On Error Resume Next
    Set doc = dbs.Containers(strContainerName).Documents(strObjectName)
    If Err.Number = 0 Then
        strDesc = CStr(doc.Properties("Description").Value)
        If Err.Number <> 0 Then
            strDesc = vbNullString
            Err.Clear
        End If
    Else
        Err.Clear
    End If

    blnHidden = Application.GetHiddenAttribute(intObjType, strObjectName)
    If Err.Number <> 0 Then
        blnHidden = False
        Err.Clear
    End If
    On Error GoTo 0

    GetMetadataHash = GetStringHash(strDesc & "|" & CStr(blnHidden), True)

End Function


'---------------------------------------------------------------------------------------
' Procedure : CollectObjectMetadata
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Reads document properties and hidden attribute from the database object
'           : and adds "Properties" and/or "Hidden" keys to the provided Items dictionary.
'           : Any existing Properties/Hidden keys are removed first.
'           : This is the shared logic used by both ExportObjectMetadata (file-based path)
'           : and inline callers like clsDbTableDef that build their dictionary in memory.
'---------------------------------------------------------------------------------------
'
Public Sub CollectObjectMetadata(dItems As Dictionary, strContainerName As String, _
                                 strObjectName As String, intObjType As AcObjectType)

    Dim dProps As Dictionary
    Dim dProp As Dictionary
    Dim doc As DAO.Document
    Dim prp As DAO.Property
    Dim dbs As Database
    Dim blnHidden As Boolean

    ' Gate behind export format version
    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Sub

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Remove any existing metadata keys before rebuilding
    If dItems.Exists("Properties") Then dItems.Remove "Properties"
    If dItems.Exists("Hidden") Then dItems.Remove "Hidden"

    ' Read document properties from DAO
    Set dbs = SharedDb
    Set dProps = New Dictionary

    If Options.SaveAllDocumentProperties Then
        ' Deep scan: iterate all properties, skipping standard DAO ones
        Set doc = dbs.Containers(strContainerName).Documents(strObjectName)
        For Each prp In doc.Properties
            Select Case prp.Name
                Case "AllPermissions", "Container", "DateCreated", _
                     "LastUpdated", "Name", "Owner", "GUID", _
                     "Permissions", "UserName"
                    ' Skip standard DAO properties
                Case Else
                    Set dProp = New Dictionary
                    dProp.CompareMode = TextCompare
                    dProp.Add "Type", prp.Type
                    dProp.Add "Value", prp.Value
                    dProps.Add prp.Name, dProp
            End Select
        Next prp
        CatchAny eelError, "Error reading document properties for " & strObjectName, _
            ModuleName & ".CollectObjectMetadata"
    Else
        ' Fast path: only check for Description property
        LogUnhandledErrors
        On Error Resume Next
        Set doc = dbs.Containers(strContainerName).Documents(strObjectName)
        If Err.Number = 0 Then
            Set prp = doc.Properties("Description")
            If Err.Number = 0 Then
                Set dProp = New Dictionary
                dProp.CompareMode = TextCompare
                dProp.Add "Type", prp.Type
                dProp.Add "Value", prp.Value
                dProps.Add "Description", dProp
            End If
            Err.Clear
        Else
            Err.Clear
        End If
        If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    End If

    ' Add Properties section if any properties were found
    If dProps.Count > 0 Then
        dItems.Add "Properties", SortDictionaryByKeys(dProps)
    End If

    ' Check hidden attribute
    LogUnhandledErrors
    On Error Resume Next
    blnHidden = Application.GetHiddenAttribute(intObjType, strObjectName)
    If Err.Number <> 0 Then blnHidden = False
    Err.Clear
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    If blnHidden Then
        dItems.Add "Hidden", True
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObjectMetadata
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Reads document properties and hidden attribute for a single database
'           : object and writes/updates the "Properties" and "Hidden" keys in the
'           : companion .json file. Merges with any existing content (print settings,
'           : linked table data) rather than replacing the whole file.
'           : Removes the "Properties"/"Hidden" keys when empty, and deletes the
'           : .json file entirely if no content remains.
'---------------------------------------------------------------------------------------
'
Public Sub ExportObjectMetadata(strJsonFile As String, strContainerName As String, _
                                strObjectName As String, cDbObjectClass As IDbComponent)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dHeader As Dictionary
    Dim strClass As String
    Dim intObjType As AcObjectType

    ' Gate behind export format version
    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Sub

    ' The Access object type is derived from the component (edb* enum members
    ' equal their AcObjectType counterparts for these standard object types).
    intObjType = cDbObjectClass.ComponentType

    ' The owning component class (e.g. clsDbForm) is recorded in Info.Class for
    ' EFV_5_1_0+; earlier formats leave it empty for backward-compatible bytes.
    If Options.ExportFormatVersion >= EFV_5_1_0 Then
        strClass = TypeName(cDbObjectClass)
    Else
        strClass = vbNullString
    End If

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Read existing .json file content (may contain linked table data, print settings, etc.)
    ' We modify the existing dictionary in-place so that all existing content (Info header,
    ' Items keys like Connect/SourceTableName, etc.) is preserved automatically.
    If FSO.FileExists(strJsonFile) Then
        Set dFile = ReadJsonFile(strJsonFile)
    End If
    If dFile Is Nothing Then Set dFile = New Dictionary
    If dFile.Exists("Items") Then
        Set dItems = dFile("Items")
    Else
        Set dItems = New Dictionary
        Set dFile("Items") = dItems
    End If

    ' Collect metadata into the Items dictionary
    CollectObjectMetadata dItems, strContainerName, strObjectName, intObjType

    ' Determine if the file has any remaining content worth keeping
    If dItems.Exists("Properties") Or dItems.Exists("Hidden") Or HasNonMetadataKeys(dItems) Then
        ' Ensure the file has an Info header (needed for new companion files;
        ' existing files like linked tables already have one and it is preserved)
        If Not dFile.Exists("Info") Then
            Set dHeader = New Dictionary
            dHeader.Add "Class", strClass
            dHeader.Add "Description", strObjectName & " Metadata"
            ' Build new dictionary with Info before Items for correct JSON key order
            Set dFile = New Dictionary
            dFile.Add "Info", dHeader
            dFile.Add "Items", dItems
        ElseIf Options.ExportFormatVersion >= EFV_5_1_0 Then
            Set dHeader = dFile("Info")
            dHeader("Class") = strClass
        End If
        WriteFile ConvertToJson(dFile, JSON_WHITESPACE), strJsonFile
    Else
        ' No metadata and no other content -- remove the file
        If FSO.FileExists(strJsonFile) Then DeleteFile strJsonFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteConditionalFormatting
' Author    : Adam Waller
' Date      : 6/17/2026
' Purpose   : Write/update the "ConditionalFormatting" section in the companion .json
'           : file, keyed by control name. Merges with any existing content (print
'           : settings, metadata) rather than replacing the whole file. Removes the
'           : section when empty and deletes the file if no content remains.
'---------------------------------------------------------------------------------------
'
Public Sub WriteConditionalFormatting(strJsonFile As String, dCF As Dictionary, _
                                      strObjectName As String, strComponentClass As String)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dHeader As Dictionary
    Dim strClass As String

    ' Use the owning component class (e.g. clsDbForm) for EFV_5_1_0+; earlier
    ' formats keep the legacy "clsSourceParser" value for backward-compatible bytes.
    If Options.ExportFormatVersion >= EFV_5_1_0 And Len(strComponentClass) > 0 Then
        strClass = strComponentClass
    Else
        strClass = "clsSourceParser"
    End If

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Read existing companion JSON (may hold print settings / metadata)
    If FSO.FileExists(strJsonFile) Then Set dFile = ReadJsonFile(strJsonFile)
    If dFile Is Nothing Then Set dFile = New Dictionary
    If dFile.Exists("Items") Then
        Set dItems = dFile("Items")
    Else
        Set dItems = New Dictionary
        Set dFile("Items") = dItems
    End If

    ' Replace any prior section
    If dItems.Exists("ConditionalFormatting") Then dItems.Remove "ConditionalFormatting"
    If Not dCF Is Nothing Then
        If dCF.Count > 0 Then dItems.Add "ConditionalFormatting", dCF
    End If

    ' Write the file if it has content, otherwise clean up
    If dItems.Count > 0 Then
        If Not dFile.Exists("Info") Then
            Set dHeader = New Dictionary
            dHeader.Add "Class", strClass
            dHeader.Add "Description", strObjectName & " Conditional Formatting"
            Set dFile = New Dictionary
            dFile.Add "Info", dHeader
            dFile.Add "Items", dItems
        ElseIf Options.ExportFormatVersion >= EFV_5_1_0 Then
            Set dHeader = dFile("Info")
            dHeader("Class") = strClass
        End If
        WriteFile ConvertToJson(dFile, JSON_WHITESPACE), strJsonFile
    Else
        If FSO.FileExists(strJsonFile) Then DeleteFile strJsonFile
    End If

    CatchAny eelError, "Error writing conditional formatting for " & strObjectName, _
        ModuleName & ".WriteConditionalFormatting"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportObjectMetadata
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Reads "Properties" and "Hidden" from a companion .json file and applies
'           : them to the database object. Called after the object has been created
'           : via LoadFromText or equivalent. When blnSkipDocumentsRefresh is True,
'           : the caller has already refreshed the container (batch module import).
'---------------------------------------------------------------------------------------
'
Public Sub ImportObjectMetadata(strJsonFile As String, strContainerName As String, _
                                strObjectName As String, intObjType As AcObjectType, _
                                Optional blnSkipDocumentsRefresh As Boolean = False)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dProps As Dictionary
    Dim dProp As Dictionary
    Dim dbs As Database
    Dim varProp As Variant

    ' Only process .json files that exist
    If Not FSO.FileExists(strJsonFile) Then Exit Sub

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set dFile = ReadJsonFile(strJsonFile)
    If dFile Is Nothing Then Exit Sub
    If Not dFile.Exists("Items") Then Exit Sub

    Set dItems = dFile("Items")
    Set dbs = SharedDb

    ' Apply document properties
    If dItems.Exists("Properties") Then
        If Not blnSkipDocumentsRefresh Then
            Perf.OperationStart "Refresh Documents"
            dbs.Containers(strContainerName).Documents.Refresh
            Perf.OperationEnd
        End If
        Set dProps = dItems("Properties")
        For Each varProp In dProps.Keys
            If TypeOf dProps(varProp) Is Dictionary Then
                Set dProp = dProps(varProp)
                SetDAOProperty dbs.Containers(strContainerName).Documents(strObjectName), _
                    dProp("Type"), CStr(varProp), dProp("Value")
            Else
                SetDAOProperty dbs.Containers(strContainerName).Documents(strObjectName), _
                    dbText, CStr(varProp), dProps(varProp)
            End If
            CatchAny eelError, "Error setting document property " & strContainerName & _
                "." & strObjectName & "." & varProp, ModuleName & ".ImportObjectMetadata"
        Next varProp
    End If

    ' Apply hidden attribute
    If dItems.Exists("Hidden") Then
        If dItems("Hidden") = True Then
            Application.SetHiddenAttribute intObjType, strObjectName, True
            CatchAny eelError, "Error setting hidden attribute for " & strObjectName, _
                ModuleName & ".ImportObjectMetadata"
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasNonMetadataKeys
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Returns True if the Items dictionary contains keys other than
'           : "Properties" and "Hidden" (i.e. print settings, linked table data).
'---------------------------------------------------------------------------------------
'
Private Function HasNonMetadataKeys(dItems As Dictionary) As Boolean
    Dim varKey As Variant
    For Each varKey In dItems.Keys
        Select Case varKey
            Case "Properties", "Hidden"
                ' These are metadata keys managed by ExportObjectMetadata
            Case Else
                HasNonMetadataKeys = True
                Exit Function
        End Select
    Next varKey
End Function
