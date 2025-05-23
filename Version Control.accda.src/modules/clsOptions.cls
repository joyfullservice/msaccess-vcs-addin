﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Note, this class could be described as Options, Settings or Preferences. After some
' deliberation and review, I chose to go with Options because it is most widely used
' in Microsoft products (Office, Visual Studio, VBA IDE, etc...)
' Further reading: https://english.stackexchange.com/questions/59058/

Private Const cstrOptionsFilename As String = "vcs-options.json"
Private Const cstrSourcePathProperty As String = "VCS Source Path"
Private Const cstrBuildPathProperty As String = "VCS Build Path"

' Options
Public ExportFolder As String
Public ShowDebug As Boolean
Public UseFastSave As Boolean
Public UseMergeBuild As Boolean
Public UseGitIntegration As Boolean
Public GitSettings As Dictionary
Public SavePrintVars As Boolean
Public ExportPrintSettings As Dictionary
Public SaveQuerySQL As Boolean
Public FormatSQL As Boolean
Public ForceImportOriginalQuerySQL As Boolean
Public SaveTableSQL As Boolean
Public SplitLayoutFromVBA As Boolean
Public StripPublishOption As Boolean
Public SanitizeColors As eSanitizeLevel
Public SanitizeLevel As eSanitizeLevel
Public ExtractThemeFiles As Boolean
Public TablesToExportData As Dictionary
Public SchemaExports As Dictionary
Public RunBeforeExport As String
Public RunAfterExport As String
Public RunBeforeBuild As String
Public RunAfterBuild As String
Public RunBeforeMerge As String
Public RunAfterMerge As String
Public ShowVCSLegacy As Boolean
Public HashAlgorithm As String
Public UseShortHash As Boolean
Public BreakOnError As Boolean
Public PreserveRubberDuckID As Boolean

' Enum to compare loaded version to installed version.
Public Enum eVersionCompare
    evcUnknown
    evcOlderVersion
    evcSameVersion
    evcNewerVersion
End Enum

' Private collections for options and enum values.
Private m_colOptions As Collection
Private m_dEnum As Dictionary
Private m_strOptionsFilePath As String
Private m_strLoadedVersion As String


'---------------------------------------------------------------------------------------
' Procedure : LoadDefaults
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Loads the default values. Define system defaults here.
'           : (Some values not defined here when they initialize to the default state.)
'---------------------------------------------------------------------------------------
'
Public Sub LoadDefaults()

    With Me
        ' Top level settings
        .ExportFolder = vbNullString
        .ShowDebug = False
        .UseFastSave = True
        .UseMergeBuild = False
        .UseGitIntegration = False
        .SavePrintVars = True
        .SaveQuerySQL = True
        .FormatSQL = True
        .ForceImportOriginalQuerySQL = False
        .SaveTableSQL = True
        .SplitLayoutFromVBA = True
        .StripPublishOption = True
        .SanitizeLevel = eslStandard
        .SanitizeColors = eslMinimal
        .ShowVCSLegacy = True
        .HashAlgorithm = DefaultHashAlgorithm
        .UseShortHash = True

        ' Table data export
        Set .TablesToExportData = New Dictionary
        ' Set CompareMode to textual comparison
        .TablesToExportData.CompareMode = vbTextCompare
        ' Save specific tables by default
        AddTableToExportData "USysRegInfo", etdTabDelimited
        AddTableToExportData "USysRibbons", etdTabDelimited

        ' External Database Schemas
        Set .SchemaExports = New Dictionary
        .SchemaExports.CompareMode = TextCompare

        ' Print settings to export
        Set .ExportPrintSettings = New Dictionary
        With .ExportPrintSettings
            .CompareMode = TextCompare
            .Add "Orientation", True
            .Add "PaperSize", True
            .Add "Duplex", False
            .Add "PrintQuality", False
            .Add "DisplayFrequency", False
            .Add "Collate", False
            .Add "Resolution", False
            .Add "DisplayFlags", False
            .Add "Color", False
            .Add "Copies", False
            .Add "ICMMethod", False
            .Add "DefaultSource", False
            .Add "Scale", False
            .Add "ICMIntent", False
            .Add "FormName", False
            .Add "PaperLength", False
            .Add "DitherType", False
            .Add "MediaType", False
            .Add "PaperWidth", False
            .Add "TTOption", False
        End With

        ' Git integration settings
        Set .GitSettings = New Dictionary
        With .GitSettings
            .CompareMode = TextCompare
            .Add "MergeUntrackedFiles", True
            .Add "ImportTableData", False
            .Add "MergeQuerySQL", False
            .Add "MergeConflicts", "Cancel Merge"
            .Add "RunBeforeMerge", vbNullString
            .Add "RunAfterMerge", vbNullString
            .Add "InspectSharedImages", False
            .Add "InspectThemeFiles", False
        End With
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddTableToExportData
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Add a table to the list of saved tables
'---------------------------------------------------------------------------------------
'
Public Sub AddTableToExportData(strName As String, intExportFormat As eTableDataExportFormat)

    Dim strFormat(etdTabDelimited To etdXML) As String
    Dim dTable As Dictionary

    Set dTable = New Dictionary

    strFormat(etdTabDelimited) = "TabDelimited"
    strFormat(etdXML) = "XMLFormat"
    With Me.TablesToExportData
        Set .Item(strName) = dTable
        .Item(strName)("Format") = GetTableExportFormatName(intExportFormat)
        ' Could add ExcludeColumns here later...
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsToFile
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Save the loaded Options to a file in JSON format
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsToFile(strPath As String)
    WriteFile modJsonConverter.ConvertToJson(SerializeOptions, JSON_WHITESPACE) & vbCrLf, strPath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsAsDefault
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save these options as default for projects.
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsAsDefault()
    Me.SaveOptionsToFile GetDefaultOptionsFilePath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsForProject
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save these options for this particular project.
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsForProject()
    ' Save source path option in current database.
    SavedSourcePath = Me.ExportFolder
    ' Set the build path, if needed.
    SetBuildPath
    ' Save options to the export folder location
    Me.SaveOptionsToFile Me.GetExportFolder & cstrOptionsFilename
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptionsFromFile
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load matching options from json file.
'---------------------------------------------------------------------------------------
'
Public Sub LoadOptionsFromFile(strFile As String)

    Dim dFile As Dictionary
    Dim dOptions As Dictionary
    Dim varOption As Variant
    Dim strKey As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Save file path, in case we need to use it to determine
    ' the export folder location with no database open.
    m_strOptionsFilePath = strFile

    ' Read in the options from the json file.
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        ' Get version of loaded file
        m_strLoadedVersion = dNZ(dFile, "Info\AddinVersion")
        If dFile.Exists("Options") Then
            Set dOptions = dFile("Options")
            ' Perform any relevant option upgrades
            Upgrade dOptions
            ' Attempt to set any matching options in this class.
            For Each varOption In m_colOptions
                strKey = CStr(varOption)
                If dOptions.Exists(strKey) Then
                    ' Set class property with value read from file.
                    Select Case strKey
                        Case "ExportPrintSettings"
                            Set Me.ExportPrintSettings = dOptions(strKey)
                        Case "TablesToExportData"
                            Set Me.TablesToExportData = CloneDictionary(dOptions(strKey), ecmTextCompare)
                        Case "SchemaExports"
                            Set Me.SchemaExports = CloneDictionary(dOptions(strKey), ecmTextCompare)
                        Case Else
                            ' Regular top-level properties
                            CallByName Me, strKey, VbLet, dOptions(strKey)
                    End Select
                End If
            Next varOption
        End If
    End If

    ' Set the build path property when loading options
    SetBuildPath

    CatchAny eelError, "Loading options from " & strFile, ModuleName(Me) & ".LoadOptionsFromFile"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLoadedVersion
' Author    : Adam Waller
' Date      : 11/21/2023
' Purpose   : Return the VCS version that was loaded for this project.
'---------------------------------------------------------------------------------------
'
Public Function GetLoadedVersion() As String
    GetLoadedVersion = m_strLoadedVersion
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadedVersion
' Author    : Adam Waller
' Date      : 11/21/2023
' Purpose   : Compares the loaded version to the currently installed version of the
'           : add-in.
'---------------------------------------------------------------------------------------
'
Public Function CompareLoadedVersion() As eVersionCompare

    Dim strLoaded As String
    Dim strInstalled As String

    ' Build comparison strings
    strLoaded = BuildVersionCompareString(m_strLoadedVersion)
    strInstalled = BuildVersionCompareString(GetVCSVersion)

    If Len(strLoaded) And Len(strInstalled) Then
        If strLoaded < strInstalled Then CompareLoadedVersion = evcOlderVersion
        If strLoaded = strInstalled Then CompareLoadedVersion = evcSameVersion
        If strLoaded > strInstalled Then CompareLoadedVersion = evcNewerVersion
    Else
        CompareLoadedVersion = evcUnknown
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildVersionCompareString
' Author    : Adam Waller
' Date      : 11/21/2023
' Purpose   : Reformat the string to a padded string that can be used to compare
'           : versions.  I.e. "4.1.345" becomes "000040000100345"
'---------------------------------------------------------------------------------------
'
Private Function BuildVersionCompareString(strVersion As String) As String

    Dim varParts As Variant

    ' Build version strings to compare
    varParts = Split(strVersion, ".")
    If UBound(varParts) = 2 Then
        BuildVersionCompareString = _
            Format(varParts(0), "00000") & _
            Format(varParts(1), "00000") & _
            Format(varParts(2), "00000")
    Else
        ' Unable to parse version
        BuildVersionCompareString = strVersion
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 11/21/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Upgrade(ByRef dOptions As Dictionary)

    Dim varKey As Variant
    Dim dItem As Dictionary
    Dim varItems As Variant
    Dim colItems As Collection
    Dim varLine As Variant

    ' 6/16/2021
    ' Aggressive sanitize to sanitize levels
    If dOptions.Exists("AggressiveSanitize") Then
        If Not dOptions.Exists("SanitizeLevel") Then
            ' Check for non-default level
            If Not dOptions("AggressiveSanitize") Then
                ' Set to basic level
                dOptions.Add "SanitizeLevel", eslMinimal
            End If
        End If
    End If

    ' 11/3/2023
    ' Check option to split VBA from object layout
    If Not dOptions.Exists("SplitLayoutFromVBA") Then
        ' The existing options file does not have this option defined.
        ' See if we have any source files from previous exports.
        If HasUnifiedLayoutFilesInGit(Me.GetExportFolder) Then
            ' Set the option as false by default, and let the user
            ' turn it on explicitly for this project.
            ' (That way they are not forced to make a decision immediately)
            Me.SplitLayoutFromVBA = False
        Else
            ' If we already have split files, or if this project is
            ' being exported for the first time, leave the option at
            ' the default setting.
        End If
    End If

    ' 1/9/2024 (4.0.31)
    ' Use collection to store schema filter entries
    If dOptions.Exists("SchemaExports") Then
        If dOptions("SchemaExports").Count > 0 Then
            For Each varKey In dOptions("SchemaExports").Keys
                Set dItem = dOptions("SchemaExports").Item(varKey)
                If dItem.Exists("Filter") Then
                    ' If this is a string value, upgrade it to a collection
                    If Not IsObject(dItem("Filter")) Then
                        Set colItems = New Collection
                        varItems = Split(dItem("Filter"), "\r\n")
                        Set dItem("Filter") = New Collection
                        ' Build collection items
                        For Each varLine In varItems
                            colItems.Add CStr(varLine)
                        Next varLine
                        Set dItem("Filter") = colItems
                    End If
                End If
            Next varKey
        End If
    End If

    ' 2/8/2025 (4.0.39)
    ' Remove UI option for SanitizeLevel = None (#554)
    If dOptions.Exists("SanitizeLevel") Then
        If dOptions("SanitizeLevel") = eSanitizeLevel.eslNone Then
            dOptions("SanitizeLevel") = eSanitizeLevel.eslMinimal
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasUnifiedLayoutFilesInGit
' Author    : Adam Waller
' Date      : 11/3/2023
' Purpose   : Returns true if the current project seems to have existing form or report
'           : source files AND appears to be in a .git repository.
'           : (This function is used when determining the default for splitting VBA
'           :  from layout files in new projects.)
'           : For performance reasons this is not a fully comprehensive check of every
'           : possible source file, but should be a pretty good indication of whether
'           : existing source files need to be split in git to preserve the history in
'           : both source files.
'---------------------------------------------------------------------------------------
'
Private Function HasUnifiedLayoutFilesInGit(strExportPath As String) As Boolean

    Dim blnHasFiles As Boolean

    ' See if we have any ".bas" files, but no corresponding ".cls" files in the
    ' forms and reports export folders.
    ' Hard-coding the folder names to avoid calling options.
    If GetFileList(BuildPath2(strExportPath, "forms"), "*.bas").Count > 0 Then
        blnHasFiles = (GetFileList(BuildPath2(strExportPath, "forms"), "*.cls").Count = 0)
    ElseIf GetFileList(BuildPath2(strExportPath, "reports"), "*.bas").Count > 0 Then
        blnHasFiles = (GetFileList(BuildPath2(strExportPath, "reports"), "*.cls").Count = 0)
    End If

    If blnHasFiles Then
        ' Check to see if this folder is in a git repository
        If Git.Installed Then
            ' Check export path
            Git.WorkingFolder = strExportPath
            HasUnifiedLayoutFilesInGit = Git.IsInsideRepository
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadProjectOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Loads the project options from a saved file. (If any)
'---------------------------------------------------------------------------------------
'
Public Sub LoadProjectOptions()

    Dim strSaved As String

    ' We can only load the options for the current project if we
    ' have a database file open.
    If Not DatabaseFileOpen Then Exit Sub

    ' Get saved path from database (if defined)
    strSaved = SavedSourcePath

    ' Attempt to load the project options file.
    If strSaved <> vbNullString Then Me.ExportFolder = strSaved
    LoadOptionsFromFile Me.GetExportFolder & cstrOptionsFilename

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadDefaultOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load in the default options from the saved json file in the addin path.
'---------------------------------------------------------------------------------------
'
Public Sub LoadDefaultOptions()
    LoadOptionsFromFile GetDefaultOptionsFilePath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDefaultOptionsFilePath
' Author    : Adam Waller
' Date      : 5/7/2021
' Purpose   : Return the full path of the default options file
'---------------------------------------------------------------------------------------
'
Private Function GetDefaultOptionsFilePath() As String
    GetDefaultOptionsFilePath = FSO.BuildPath(CodeProject.Path, cstrOptionsFilename)
End Function


'---------------------------------------------------------------------------------------
' Procedure : PrintOptionsToDebugWindow
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Output the current options.
'---------------------------------------------------------------------------------------
'
Public Sub PrintOptionsToDebugWindow()
    Debug.Print modJsonConverter.ConvertToJson(SerializeOptions, JSON_WHITESPACE)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetExportFolder
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Returns the actual export folder path from the ExportFolder option.
'           : NOTE: We cannot return a relative export folder path without an open
'           : database file.
'---------------------------------------------------------------------------------------
'
Public Function GetExportFolder() As String

    Dim strFullPath As String

    If Me.ExportFolder = vbNullString Then
        If DatabaseFileOpen Then
            ' Build default path using project file name
            strFullPath = CurrentProject.FullName & ".src" & PathSep
        Else
            ' Check options file path
            If Len(m_strOptionsFilePath) Then
                strFullPath = FSO.GetParentFolderName(m_strOptionsFilePath)
            End If
        End If
    Else
        If Left$(Me.ExportFolder, 2) = PathSep & PathSep Then
            ' UNC path
            strFullPath = Me.ExportFolder
        ElseIf Left$(Me.ExportFolder, 1) = PathSep Then
            ' Relative path (from database file location)
            strFullPath = CurrentProject.Path & Me.ExportFolder
        Else
            ' Other absolute path (i.e. c:\myfiles\)
            strFullPath = Me.ExportFolder
        End If
        ' Environment variable replacements
        If InStr(1, strFullPath, "%") > 0 Then
            strFullPath = ExpandEnvironmentVariables(strFullPath)
            ' Other placeholder replacements
            If InStr(1, strFullPath, "%dbName%", vbTextCompare) > 0 Then
                strFullPath = Replace(strFullPath, "%dbName%", CurrentProject.Name, , , vbTextCompare)
            End If
        End If
    End If

    ' Check to make sure we have built a valid path.
    Select Case strFullPath
        Case PathSep & PathSep, PathSep, ".src" & PathSep, vbNullString
            ' Invalid paths
            MsgBox2 "Cannot determine export path", _
                "A database file must be open to return a relative export path.", _
                "This is probably due to a problem in the build sequence logic.", vbExclamation
            Log.Add "ERROR: Cannot build relative export folder path without an open database file."
            GetExportFolder = vbNullString
        Case Else
            ' Return export path with a trailing slash
            GetExportFolder = StripSlash(strFullPath) & PathSep
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : SerializeOptions
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Serializes Options into a dictionary array for saving to file as JSON.
'---------------------------------------------------------------------------------------
'
Private Function SerializeOptions() As Dictionary

    Dim dOptions As Dictionary
    Dim dInfo As Dictionary
    Dim dWrapper As Dictionary
    Dim varOption As Variant
    Dim strBit As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set dOptions = New Dictionary
    Set dInfo = New Dictionary
    Set dWrapper = New Dictionary

    ' Add some header information (For debugging or upgrading)
    #If Win64 Then
        strBit = " 64-bit"
    #Else
        strBit = " 32-bit"
    #End If
    dInfo.Add "AddinVersion", AppVersion
    dInfo.Add "AccessVersion", Application.Version & strBit

    ' Loop through options
    For Each varOption In m_colOptions
        ' Simulate reflection to serialize properties.
        dOptions.Add CStr(varOption), CallByName(Me, CStr(varOption), VbGet)
    Next varOption

    'Set SerializeOptions = new Dictionary
    Set dWrapper("Info") = dInfo
    Set dWrapper("Options") = dOptions
    Set SerializeOptions = dWrapper

    CatchAny eelError, "Serializing options", ModuleName(Me) & ".SerializeOptions"

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHash
' Author    : Adam Waller
' Date      : 2/16/2021
' Purpose   : Return a hash of the current options. Used to detect if options have
'           : changed, which may require a full export to reflect the change.
'           : Note that this only checks a hash of the options that affect general
'           : exports, not changes to external database connections.
'---------------------------------------------------------------------------------------
'
Public Function GetHash() As String
    Dim dExportOptions As Dictionary
    Set dExportOptions = CloneDictionary(SerializeOptions)
    With dExportOptions("Options")
        ' Remove options not related to exports
        If .Exists("SchemaExports") Then .Remove ("SchemaExports")
    End With
    GetHash = GetDictionaryHash(dExportOptions)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSanitizeLevelName
' Author    : Hecon5 (as adapted from Adam Waller's below)
' Date      : 6/09/2021
' Purpose   : Return the name used to read and write to the JSON options files.
'---------------------------------------------------------------------------------------
'
Public Function GetSanitizeLevelName(intSanitizeLevel As eSanitizeLevel) As String
    Select Case intSanitizeLevel
        Case eslNone:           GetSanitizeLevelName = "None (Off)"
        Case eslMinimal:        GetSanitizeLevelName = "Minimal"
        Case eslStandard:       GetSanitizeLevelName = "Standard"
        Case eslExtended:       GetSanitizeLevelName = "Extended"
        Case Else:              GetSanitizeLevelName = vbNullString
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSanitizeLevel
' Author    : Hecon5 (as adapted from Adam Waller's below)
' Date      : 4/17/2020
' Purpose   : Translate the table export format key to the corresponding enum value.
'---------------------------------------------------------------------------------------
'
Public Function GetSanitizeLevel(strKey As String) As eSanitizeLevel
    Dim intSanitizeLevel As eSanitizeLevel
    Dim strName As String
    For intSanitizeLevel = eslNone To eSanitizeLevel.[_Last]
        strName = Me.GetSanitizeLevelName(intSanitizeLevel)
        If strName = strKey Then
            GetSanitizeLevel = intSanitizeLevel
            Exit For
        End If
    Next intSanitizeLevel
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportFormatName
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Return the name used to read and write to the JSON options files.
'---------------------------------------------------------------------------------------
'
Public Function GetTableExportFormatName(intFormat As eTableDataExportFormat) As String
    Select Case intFormat
        Case etdTabDelimited:   GetTableExportFormatName = "Tab Delimited"
        Case etdXML:            GetTableExportFormatName = "XML Format"
        Case Else:              GetTableExportFormatName = vbNullString
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportFormat
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Translate the table export format key to the corresponding enum value.
'---------------------------------------------------------------------------------------
'
Public Function GetTableExportFormat(strKey As String) As eTableDataExportFormat
    Dim intFormat As eTableDataExportFormat
    Dim strName As String
    For intFormat = etdNoData To eTableDataExportFormat.[_Last] - 1
        strName = Me.GetTableExportFormatName(intFormat)
        If strName = strKey Then
            GetTableExportFormat = intFormat
            Exit For
        End If
    Next intFormat
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Load default values when creating class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    ' Initialize the options collection
    Set m_colOptions = New Collection

    ' Load enum values
    Set m_dEnum = New Dictionary

    ' Load list of property names for reflection type behavior.
    With m_colOptions
        .Add "ExportFolder"
        .Add "ShowDebug"
        .Add "UseFastSave"
        .Add "UseMergeBuild"
        .Add "UseGitIntegration"
        .Add "SavePrintVars"
        .Add "ExportPrintSettings"
        .Add "SaveQuerySQL"
        .Add "FormatSQL"
        .Add "ForceImportOriginalQuerySQL"
        .Add "SaveTableSQL"
        .Add "SplitLayoutFromVBA"
        .Add "StripPublishOption"
        .Add "SanitizeColors"
        .Add "SanitizeLevel"
        .Add "ExtractThemeFiles"
        .Add "TablesToExportData"
        .Add "SchemaExports"
        .Add "RunBeforeExport"
        .Add "RunAfterExport"
        .Add "RunBeforeBuild"
        .Add "RunAfterBuild"
        .Add "RunBeforeMerge"
        .Add "RunAfterMerge"
        .Add "ShowVCSLegacy"
        .Add "HashAlgorithm"
        .Add "UseShortHash"
        .Add "BreakOnError"
        .Add "PreserveRubberDuckID"
    End With

    ' Load default values
    Me.LoadDefaults

    ' Load saved defaults
    LoadDefaultOptions

    ' Other run-time options
    JsonOptions.AllowUnicodeChars = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetBuildPath
' Author    : Adam Waller
' Date      : 10/28/2021
' Purpose   : Set a build path property value if the export folder is not relative
'           : to the current database path.
'---------------------------------------------------------------------------------------
'
Private Sub SetBuildPath()

    Dim prp As AccessObjectProperty
    Dim proj As CurrentProject

    Set proj = CurrentProject
    Set prp = GetPropertyByName(cstrBuildPathProperty)

    ' This should only be done on an open database/project
    If DatabaseFileOpen Then
        ' Check the export folder option
        If Me.ExportFolder = vbNullString Then
            ' When using the default of a blank export path, we don't need
            ' to store the build folder location. (It will be different on
            ' different computers, and not needed for a relative export path.)
            If Not prp Is Nothing Then proj.Properties.Remove prp.Name
        Else
            If prp Is Nothing Then
                ' Create the property
                proj.Properties.Add cstrBuildPathProperty, proj.Path
            Else
                ' Update the value.
                If prp.Value <> proj.Path Then prp.Value = proj.Path
            End If
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SavedSourcePath
' Author    : Adam Waller
' Date      : 7/13/2020
' Purpose   : Get any saved path for VCS source files. (In case we are using a
'           : different location for the files.) This is stored as a property
'           : under the currentproject. (Works for both ADP and MDB)
'---------------------------------------------------------------------------------------
'
Private Property Get SavedSourcePath() As String
    Dim prp As AccessObjectProperty
    Set prp = GetPropertyByName(cstrSourcePathProperty)
    If Not prp Is Nothing Then SavedSourcePath = prp.Value
End Property


'---------------------------------------------------------------------------------------
' Procedure : SavedSourcePath
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Save the source path as a property in the current database.
'---------------------------------------------------------------------------------------
'
Private Property Let SavedSourcePath(strPath As String)

    Dim prp As AccessObjectProperty
    Dim proj As CurrentProject

    Set proj = CurrentProject
    Set prp = GetPropertyByName(cstrSourcePathProperty)

    If strPath = vbNullString Then
        ' Remove the property when no longer used.
        If Not prp Is Nothing Then proj.Properties.Remove prp.Name
    Else
        If prp Is Nothing Then
            ' Create the property
            proj.Properties.Add cstrSourcePathProperty, strPath
        Else
            ' Update the value.
            prp.Value = strPath
        End If
    End If

End Property


'---------------------------------------------------------------------------------------
' Procedure : GetPropertyByName
' Author    : Adam Waller
' Date      : 10/28/2021
' Purpose   : Return a project property
'---------------------------------------------------------------------------------------
'
Private Function GetPropertyByName(strName As String) As AccessObjectProperty

    Dim prp As AccessObjectProperty
    Dim proj As CurrentProject

    If DatabaseFileOpen Then
        Set proj = CurrentProject
        For Each prp In proj.Properties
            If prp.Name = strName Then
                Set GetPropertyByName = prp
                Exit For
            End If
        Next prp
    End If

End Function
