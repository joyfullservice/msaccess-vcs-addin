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
                                    , Optional cDbObjectClass As IDbComponent = Nothing) As String

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
                                    , Optional blnSuppressError As Boolean = False) As Boolean

    Const FunctionName As String = ModuleName & ".LoadComponentFromText"

    Dim blnErrInFunction As Boolean
    Dim strTempFile As String
    Dim strSourceFile As String
    Dim strPrefix As String
    Dim strAltFile As String
    Dim strContent As String
    Dim blnVbaOverlay As Boolean
    Dim blnConvert As Boolean

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

    ' Issue importing form. We need to prompt user to see if we continue on or not.
    Log.Error eelError, T("Import issue with '{0}'; {1}", var0:=strName, var1:=strErrDescription), FunctionName

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
        ' Clear out strName because we're going to use it to detect if the import failed.
        Log.Error eelError, T("Skipping import of '{0}'. Your application may not run or complile.", var0:=strName), FunctionName
        blnErrInFunction = True
        Resume Next

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
