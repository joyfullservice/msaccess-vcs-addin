Attribute VB_Name = "modVCSUtility"
'---------------------------------------------------------------------------------------
' Module    : modVCSUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : VCS-specific utility functions: version helpers, JSON file building,
'           : path resolution, git file checks, schema filters, and command bar import.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modFileAccess, modFunctions, modCollectionUtil,
'           : modStringUtil, modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

' (InteractionMode now lives on the Operation singleton --
' see clsOperation.InteractionMode and modAPI.SetInteractionMode.)

Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetKeyboardState Lib "user32" (lppbKeyState As Any) As Long
Private Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Any) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As LongPtr) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Private Const ModuleName = "modVCSUtility"


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
' Procedure : VersionToExportFormat
' Author    : Adam Waller
' Date      : 3/6/2026
' Purpose   : Convert a version string like "4.1.2" to a packed integer (40102).
'           : Uses Major * 10000 + Minor * 100 + Patch.
'---------------------------------------------------------------------------------------
'
Public Function VersionToExportFormat(strVersion As String) As Long

    Dim varParts As Variant

    varParts = Split(strVersion, ".")
    If UBound(varParts) = 2 Then
        VersionToExportFormat = CLng(varParts(0)) * 10000 + CLng(varParts(1)) * 100 + CLng(varParts(2))
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportFormatToVersion
' Author    : Adam Waller
' Date      : 3/6/2026
' Purpose   : Convert a packed integer (40102) back to a version string ("4.1.2").
'---------------------------------------------------------------------------------------
'
Public Function ExportFormatToVersion(lngFormat As Long) As String
    ExportFormatToVersion = (lngFormat \ 10000) & "." & ((lngFormat \ 100) Mod 100) & "." & (lngFormat Mod 100)
End Function


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
            strPath = strExportFolder & PathSep & ".." & PathSep & strFile
        Else
            ' Check to see if we are using an absolute export path  (\\* or *:*)
            If StartsWith(Options.ExportFolder, PathSep & PathSep) _
                Or (InStr(2, Options.ExportFolder, ":") > 0) Then
                ' Look for saved build path
                Set dContents = ReadJsonFile(FSO.BuildPath(strFolder, "proj-properties.json"))
                strPath = dNZ(dContents, "Items\VCS Build Path")
                If strPath <> vbNullString Then
                    strPath = strPath & PathSep & strFile
                Else
                    ' We may have a source path override in effect. Build in parent folder
                    ' since the source does not specify an absolute build path.
                    strPath = strExportFolder & PathSep & ".." & PathSep & strFile
                End If
            Else
                ' Calculate how many levels deep to create original path
                lngLevel = UBound(Split(StripSlash(Options.ExportFolder), PathSep))
                If lngLevel < 0 Then lngLevel = 0   ' Handle "\" to export in current folder.
                strPath = strExportFolder & PathSep & Repeat(".." & PathSep, lngLevel) & strFile
            End If
        End If

        ' Expand absolute path
        GetOriginalDbFullPathFromSource = FSO.GetAbsolutePathName(strPath)
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

    ' Return empty string if we don't have any items in the dictionary.
    ' (This also gives us an easy way to test the return value for items.)
    If dItems.Count = 0 Then Exit Function

    ' Create dictionary objects
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
    Else
        UpgradeGitignoreIndexEntry strFile
    End If

    ' Ensure *.env is excluded (protects credentials saved to .env files)
    If FSO.FileExists(strFile) Then
        If EnsureGitignoreLine(strFile, "*.env", vbNullString) Then
            Log.Add T("Added *.env to .gitignore"), , , "blue"
        End If
    End If

    ' Ensure logs/ is excluded (VCS add-in writes operation logs there)
    If FSO.FileExists(strFile) Then
        If EnsureGitignoreLineRespectComment(strFile, "logs/", "*.env") Then
            Log.Add T("Added logs/ to .gitignore"), , , "blue"
        End If
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
' Procedure : ReplaceGitignoreLine
' Author    : Adam Waller
' Date      : 5/4/2026
' Purpose   : Finds a whole line matching strOldPattern in the .gitignore file and
'           : replaces it with strNewPattern. Skips the replacement if strNewPattern
'           : already exists (prevents duplicates). Returns True if the file was
'           : modified. Matching is trimmed and case-insensitive.
'---------------------------------------------------------------------------------------
'
Private Function ReplaceGitignoreLine(strFile As String, strOldPattern As String, strNewPattern As String) As Boolean

    Dim strContent As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim blnFoundOld As Boolean
    Dim blnFoundNew As Boolean

    strContent = ReadFile(strFile)
    If Len(strContent) = 0 Then Exit Function

    varLines = Split(strContent, vbCrLf)

    For lngLine = LBound(varLines) To UBound(varLines)
        If StrComp(Trim$(varLines(lngLine)), strNewPattern, vbTextCompare) = 0 Then
            blnFoundNew = True
            Exit For
        End If
        If StrComp(Trim$(varLines(lngLine)), strOldPattern, vbTextCompare) = 0 Then
            blnFoundOld = True
        End If
    Next lngLine

    ' New pattern already present — nothing to do
    If blnFoundNew Then Exit Function

    ' Old pattern not found — nothing to replace
    If Not blnFoundOld Then Exit Function

    ' Replace old with new (re-scan to preserve original whitespace context)
    For lngLine = LBound(varLines) To UBound(varLines)
        If StrComp(Trim$(varLines(lngLine)), strOldPattern, vbTextCompare) = 0 Then
            varLines(lngLine) = strNewPattern
            Exit For
        End If
    Next lngLine

    WriteFile Join(varLines, vbCrLf), strFile
    ReplaceGitignoreLine = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : EnsureGitignoreLine
' Author    : Adam Waller
' Date      : 5/4/2026
' Purpose   : Ensures strNewPattern exists as a line in the .gitignore file. If
'           : missing, inserts it after the first line matching strAfterPattern
'           : (or appends to the end if strAfterPattern is empty or not found).
'           : Returns True if the file was modified. Matching is trimmed and
'           : case-insensitive.
'---------------------------------------------------------------------------------------
'
Private Function EnsureGitignoreLine(strFile As String, strNewPattern As String, strAfterPattern As String) As Boolean

    Dim strContent As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim lngInsertAfter As Long
    Dim strResult As String

    strContent = ReadFile(strFile)
    If Len(strContent) = 0 Then Exit Function

    varLines = Split(strContent, vbCrLf)

    ' Check if the pattern already exists
    For lngLine = LBound(varLines) To UBound(varLines)
        If StrComp(Trim$(varLines(lngLine)), strNewPattern, vbTextCompare) = 0 Then
            Exit Function
        End If
    Next lngLine

    ' Find insertion point
    lngInsertAfter = -1
    If Len(strAfterPattern) > 0 Then
        For lngLine = LBound(varLines) To UBound(varLines)
            If StrComp(Trim$(varLines(lngLine)), strAfterPattern, vbTextCompare) = 0 Then
                lngInsertAfter = lngLine
                Exit For
            End If
        Next lngLine
    End If

    If lngInsertAfter >= 0 Then
        ' Insert after the matched line
        strResult = Join(ArraySlice(varLines, LBound(varLines), lngInsertAfter), vbCrLf) & _
            vbCrLf & strNewPattern
        If lngInsertAfter < UBound(varLines) Then
            strResult = strResult & vbCrLf & _
                Join(ArraySlice(varLines, lngInsertAfter + 1, UBound(varLines)), vbCrLf)
        End If
    Else
        ' Append to end
        strResult = strContent
        If Right$(strResult, 2) <> vbCrLf Then strResult = strResult & vbCrLf
        strResult = strResult & strNewPattern
    End If

    WriteFile strResult, strFile
    EnsureGitignoreLine = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ArraySlice
' Author    : Adam Waller
' Date      : 5/4/2026
' Purpose   : Returns a portion of a Variant array from lngStart to lngEnd (inclusive).
'---------------------------------------------------------------------------------------
'
Private Function ArraySlice(varArr As Variant, lngStart As Long, lngEnd As Long) As Variant

    Dim varResult() As Variant
    Dim lngIdx As Long

    If lngStart > lngEnd Then
        ArraySlice = Array()
        Exit Function
    End If

    ReDim varResult(0 To lngEnd - lngStart)
    For lngIdx = lngStart To lngEnd
        varResult(lngIdx - lngStart) = varArr(lngIdx)
    Next lngIdx
    ArraySlice = varResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : EnsureGitignoreLineRespectComment
' Author    : Adam Waller
' Date      : 5/13/2026
' Purpose   : Like EnsureGitignoreLine, but first checks whether the pattern already
'           : exists in commented-out form (e.g. "# logs/"). If the user has
'           : deliberately commented out the line, this respects that choice and
'           : does not re-add it. Returns True only if the file was modified.
'---------------------------------------------------------------------------------------
'
Private Function EnsureGitignoreLineRespectComment( _
        strFile As String, strNewPattern As String, strAfterPattern As String) As Boolean

    Dim strContent As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim strTrimmed As String

    strContent = ReadFile(strFile)
    If Len(strContent) = 0 Then Exit Function

    varLines = Split(strContent, vbCrLf)
    For lngLine = LBound(varLines) To UBound(varLines)
        strTrimmed = Trim$(varLines(lngLine))
        ' Strip leading # and whitespace to detect commented-out patterns
        If Left$(strTrimmed, 1) = "#" Then
            strTrimmed = Trim$(Mid$(strTrimmed, 2))
        End If
        If StrComp(strTrimmed, strNewPattern, vbTextCompare) = 0 Then
            Exit Function
        End If
    Next lngLine

    EnsureGitignoreLineRespectComment = EnsureGitignoreLine(strFile, strNewPattern, strAfterPattern)

End Function


'---------------------------------------------------------------------------------------
' Procedure : UpgradeGitignoreIndexEntry
' Author    : Adam Waller
' Date      : 5/4/2026
' Purpose   : Existing repos set up from older templates may have only vcs-index.json
'           : in their .gitignore. This upgrades the entry to the wildcard pattern
'           : vcs-index.* so the new binary .idx format is also excluded.
'---------------------------------------------------------------------------------------
'
Private Sub UpgradeGitignoreIndexEntry(strFile As String)
    If ReplaceGitignoreLine(strFile, "vcs-index.json", "vcs-index.*") Then
        Log.Add T("Updated .gitignore: vcs-index.json -> vcs-index.*"), , , "blue"
    End If
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
        RestoreMainForm
    End If

    LogUnhandledErrors
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
        LogUnhandledErrors
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
        LogUnhandledErrors
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
