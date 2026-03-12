Attribute VB_Name = "modSourceUpgrade"
'---------------------------------------------------------------------------------------
' Module    : modSourceUpgrade
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Source file migration and upgrade routines for handling legacy file
'           : formats and extension changes across add-in versions.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modFileAccess, modFunctions, modErrorHandling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName As String = "modSourceUpgrade"


'---------------------------------------------------------------------------------------
' Procedure : CheckForLegacyModules
' Author    : Adam Waller
' Date      : 7/16/2020
' Purpose   : Informs the user if the database contains a legacy module from another
'           : fork of this project. (Some users might not realize that these are not
'           : needed anymore.)
'---------------------------------------------------------------------------------------
'
Public Sub CheckForLegacyModules()

    ' Check for legacy file
    If Options.ShowVCSLegacy Then
        If FSO.FileExists(Options.GetExportFolder & FSO.BuildPath("modules", "VCS_ImportExport.bas")) Then
            MsgBox2 T("Legacy Files not Needed"), _
                T("Other forks of the MSAccessVCS project used additional VBA modules to export code.") & vbNewLine & _
                T("This is no longer needed when using the installed Version Control Add-in.") & vbNewLine & vbNewLine & _
                T("Feel free to remove the legacy VCS_* modules from your database project and enjoy" & vbNewLine & _
                "a simpler, cleaner code base for ongoing development.  :-)"), _
                T("NOTE: This message can be disabled in 'Options -> Show Legacy Prompt'."), _
                vbInformation, T("Just a Suggestion...")
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpgradeSourceFiles
' Author    : Adam Waller
' Date      : 8/7/2024
' Purpose   : Removes any legacy file formats used by earlier versions of this add-in.
'---------------------------------------------------------------------------------------
'
Public Sub UpgradeSourceFiles()

    Dim strBase As String

    strBase = Options.GetExportFolder

    ' Remove legacy files by extension
    ClearFilesByExtension strBase & "sqltables", "tdf"
    ClearFilesByExtension strBase & "relations", "txt"      ' Relationships (pre-json)
    ClearFilesByExtension strBase & "report", "pv"          ' Print vars text file (pre-json)
    ClearFilesByExtension strBase & "tbldefs", "LNKD"       ' Formerly used for linked tables
    ClearFilesByExtension strBase & "tbldefs", "bas"        ' Moved to XML format
    ClearFilesByExtension strBase & "tbldefs", "tdf"

    ' Remove old .bas files from folders that now use descriptive extensions
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        ClearFilesByExtension strBase & "forms", "bas"
        ClearFilesByExtension strBase & "reports", "bas"
        ClearFilesByExtension strBase & "queries", "bas"
        ClearFilesByExtension strBase & "macros", "bas"

        ' Migrate document properties and hidden attributes from singleton files
        ' into per-object companion .json files
        MigrateMetadataToCompanionFiles
    End If

    ' Clear any print settings files if not using this option
    If Not Options.SavePrintVars Then
        ClearFilesByExtension "forms", "json"
        ClearFilesByExtension "reports", "json"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MigrateFileExtensions
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Renames source files from old .bas extension to descriptive extensions
'           : (.form, .report, .qdef, .macro) for forms, reports, queries, and macros.
'           : Also updates the VCS index keys so the next export doesn't treat
'           : every object as modified.
'---------------------------------------------------------------------------------------
'
Public Sub MigrateFileExtensions()

    Dim strBase As String
    Dim lngCount As Long

    strBase = Options.GetExportFolder

    ' Rename .bas files to new extensions in each affected folder
    lngCount = lngCount + RenameFilesInFolder(strBase & "forms", "bas", "form")
    lngCount = lngCount + RenameFilesInFolder(strBase & "reports", "bas", "report")
    lngCount = lngCount + RenameFilesInFolder(strBase & "queries", "bas", "qdef")
    lngCount = lngCount + RenameFilesInFolder(strBase & "macros", "bas", "macro")

    If lngCount > 0 Then
        ' Update VCS index keys to match the new file extensions
        VCSIndex.MigrateIndexExtension "Forms", "bas", "form"
        VCSIndex.MigrateIndexExtension "Reports", "bas", "report"
        VCSIndex.MigrateIndexExtension "Queries", "bas", "qdef"
        VCSIndex.MigrateIndexExtension "Macros", "bas", "macro"
        Log.Add T("Migrated {0} source files to new extensions", var0:=lngCount)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenameFilesInFolder
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Renames files in a folder from one extension to another.
'           : Returns the number of files renamed.
'---------------------------------------------------------------------------------------
'
Private Function RenameFilesInFolder(strFolder As String, strOldExt As String, strNewExt As String) As Long

    Dim dFiles As Dictionary
    Dim varKey As Variant
    Dim strNewPath As String

    Set dFiles = GetFilePathsInFolder(strFolder, "*." & strOldExt)
    For Each varKey In dFiles.Keys
        strNewPath = SwapExtension(CStr(varKey), strNewExt)
        If Not FSO.FileExists(strNewPath) Then
            FSO.MoveFile CStr(varKey), strNewPath
            RenameFilesInFolder = RenameFilesInFolder + 1
        End If
    Next varKey

End Function


'---------------------------------------------------------------------------------------
' Procedure : RevertFileExtensions
' Author    : Adam Waller
' Date      : 3/11/2026
' Purpose   : Reverts source files from descriptive extensions (.form, .report, .qdef,
'           : .macro) back to .bas, and flattens @Folder subfolders back to the base
'           : folder. This is the reverse of MigrateFileExtensions and is called when
'           : the export format version is downgraded below 5.0.0.
'---------------------------------------------------------------------------------------
'
Public Sub RevertFileExtensions()

    Dim strBase As String
    Dim lngCount As Long

    strBase = Options.GetExportFolder

    ' Flatten @Folder subfolders back to the base folder for all types that support them
    lngCount = lngCount + FlattenSubfolders(strBase & "forms")
    lngCount = lngCount + FlattenSubfolders(strBase & "reports")
    lngCount = lngCount + FlattenSubfolders(strBase & "modules")
    lngCount = lngCount + FlattenSubfolders(strBase & "vbeforms")

    ' Rename descriptive extensions back to .bas
    lngCount = lngCount + RenameFilesInFolder(strBase & "forms", "form", "bas")
    lngCount = lngCount + RenameFilesInFolder(strBase & "reports", "report", "bas")
    lngCount = lngCount + RenameFilesInFolder(strBase & "queries", "qdef", "bas")
    lngCount = lngCount + RenameFilesInFolder(strBase & "macros", "macro", "bas")

    If lngCount > 0 Then
        ' Update VCS index keys to match the reverted file extensions
        VCSIndex.MigrateIndexExtension "Forms", "form", "bas"
        VCSIndex.MigrateIndexExtension "Reports", "report", "bas"
        VCSIndex.MigrateIndexExtension "Queries", "qdef", "bas"
        VCSIndex.MigrateIndexExtension "Macros", "macro", "bas"
        Log.Add T("Reverted {0} source files to legacy extensions", var0:=lngCount)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MigrateMetadataToCompanionFiles
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Migrates document properties from documents.json and hidden attributes
'           : from hidden-attributes.json into per-object companion .json files.
'           : Rewrites documents.json with only "Databases" container entries, and
'           : deletes hidden-attributes.json. This is a one-time migration that runs
'           : during UpgradeSourceFiles when ExportFormatVersion >= EFV_5_0_0.
'---------------------------------------------------------------------------------------
'
Public Sub MigrateMetadataToCompanionFiles()

    Dim strBase As String
    Dim strDocFile As String
    Dim strHiddenFile As String
    Dim dDocFile As Dictionary
    Dim dDocItems As Dictionary
    Dim dHiddenFile As Dictionary
    Dim dHiddenItems As Dictionary
    Dim dDbOnly As Dictionary
    Dim varCont As Variant
    Dim varDoc As Variant
    Dim strFolder As String
    Dim strJsonFile As String
    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim lngCount As Long

    strBase = Options.GetExportFolder

    ' --- Migrate documents.json ---
    strDocFile = strBase & "documents.json"
    If FSO.FileExists(strDocFile) Then
        Set dDocFile = ReadJsonFile(strDocFile)
        If Not dDocFile Is Nothing Then
            If dDocFile.Exists("Items") Then
                Set dDocItems = dDocFile("Items")
                Set dDbOnly = New Dictionary

                Dim dCont As Dictionary
                For Each varCont In dDocItems.Keys
                    If varCont = "Databases" Then
                        ' Keep Databases container in the file
                        dDbOnly.Add varCont, dDocItems(varCont)
                    Else
                        ' Migrate non-Databases entries to companion files
                        Set dCont = dDocItems(varCont)
                        For Each varDoc In dCont.Keys
                            ' Resolve folder: "Tables" container needs per-object lookup
                            If varCont = "Tables" Then
                                strFolder = ResolveFolderForTablesContainer(CStr(varDoc), strBase)
                            Else
                                strFolder = GetFolderForContainer(CStr(varCont), strBase)
                            End If
                            If Len(strFolder) > 0 Then
                                strJsonFile = strFolder & GetSafeFileName(CStr(varDoc)) & ".json"
                                MergeMetadataIntoFile strJsonFile, "Properties", dCont(varDoc)
                                lngCount = lngCount + 1
                            End If
                        Next varDoc
                    End If
                Next varCont

                ' Rewrite documents.json with only Databases entries
                If dDbOnly.Count > 0 Then
                    WriteFile BuildJsonFile("clsDbDocument", dDbOnly, _
                        "Database Documents Properties (DAO)"), strDocFile
                Else
                    DeleteFile strDocFile
                End If
            End If
        End If
    End If

    ' --- Migrate hidden-attributes.json ---
    strHiddenFile = strBase & "hidden-attributes.json"
    If FSO.FileExists(strHiddenFile) Then
        Set dHiddenFile = ReadJsonFile(strHiddenFile)
        If Not dHiddenFile Is Nothing Then
            If dHiddenFile.Exists("Items") Then
                Set dHiddenItems = dHiddenFile("Items")
                Dim colHiddenItems As Object
                For Each varCont In dHiddenItems.Keys
                    Set colHiddenItems = dHiddenItems(varCont)
                    For Each varDoc In colHiddenItems
                        ' Resolve folder: "Tables" container needs per-object lookup
                        If varCont = "Tables" Then
                            strFolder = ResolveFolderForTablesContainer(CStr(varDoc), strBase)
                        Else
                            strFolder = GetFolderForContainer(CStr(varCont), strBase)
                        End If
                        If Len(strFolder) > 0 Then
                            strJsonFile = strFolder & GetSafeFileName(CStr(varDoc)) & ".json"
                            MergeHiddenIntoFile strJsonFile
                            lngCount = lngCount + 1
                        End If
                    Next varDoc
                Next varCont
            End If
        End If
        DeleteFile strHiddenFile
    End If

    If lngCount > 0 Then
        Log.Add T("Migrated {0} metadata entries to companion .json files", var0:=lngCount)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFolderForContainer
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Maps a DAO container name to its source file folder path.
'---------------------------------------------------------------------------------------
'
Private Function GetFolderForContainer(strContainerName As String, strBase As String) As String
    Select Case strContainerName
        Case "Forms"
            GetFolderForContainer = strBase & "forms" & PathSep
        Case "Reports"
            GetFolderForContainer = strBase & "reports" & PathSep
        Case "Scripts"
            GetFolderForContainer = strBase & "macros" & PathSep
        Case "Modules"
            GetFolderForContainer = strBase & "modules" & PathSep
        Case "Tables"
            ' Both queries and tables live in the "Tables" DAO container.
            ' During migration we check for existing source files to determine the folder.
            ' Tables -> tbldefs/, Queries -> queries/
            ' This ambiguity is resolved per-object in MigrateMetadataToCompanionFiles.
            GetFolderForContainer = vbNullString
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveFolderForTablesContainer
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : For objects in the "Tables" DAO container, determines whether the object
'           : is a query (queries/ folder) or table (tbldefs/ folder) by checking for
'           : existing source files.
'---------------------------------------------------------------------------------------
'
Private Function ResolveFolderForTablesContainer(strObjectName As String, strBase As String) As String

    Dim strSafe As String

    strSafe = GetSafeFileName(strObjectName)

    ' Check queries folder first (qdef, bas)
    If FSO.FileExists(strBase & "queries" & PathSep & strSafe & ".qdef") Or _
       FSO.FileExists(strBase & "queries" & PathSep & strSafe & ".bas") Then
        ResolveFolderForTablesContainer = strBase & "queries" & PathSep
    ' Then check tbldefs folder (xml, json)
    ElseIf FSO.FileExists(strBase & "tbldefs" & PathSep & strSafe & ".xml") Or _
           FSO.FileExists(strBase & "tbldefs" & PathSep & strSafe & ".json") Then
        ResolveFolderForTablesContainer = strBase & "tbldefs" & PathSep
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeMetadataIntoFile
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Merges a Properties dictionary into an existing or new companion .json file.
'---------------------------------------------------------------------------------------
'
Private Sub MergeMetadataIntoFile(strJsonFile As String, strKey As String, dValue As Object)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim strObjectName As String

    ' Read existing file or create new structure
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

    ' Add or replace the key
    If dItems.Exists(strKey) Then dItems.Remove strKey
    dItems.Add strKey, dValue

    ' Write the file
    strObjectName = FSO.GetBaseName(strJsonFile)
    VerifyPath strJsonFile
    WriteFile BuildJsonFile(vbNullString, dItems, strObjectName & " Metadata"), strJsonFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeHiddenIntoFile
' Author    : Adam Waller
' Date      : 3/12/2026
' Purpose   : Adds "Hidden": true to a companion .json file.
'---------------------------------------------------------------------------------------
'
Private Sub MergeHiddenIntoFile(strJsonFile As String)

    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim strObjectName As String

    ' Read existing file or create new structure
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

    ' Add Hidden flag
    If dItems.Exists("Hidden") Then dItems.Remove "Hidden"
    dItems.Add "Hidden", True

    ' Write the file
    strObjectName = FSO.GetBaseName(strJsonFile)
    VerifyPath strJsonFile
    WriteFile BuildJsonFile(vbNullString, dItems, strObjectName & " Metadata"), strJsonFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FlattenSubfolders
' Author    : Adam Waller
' Date      : 3/11/2026
' Purpose   : Moves all files from subfolders back to the base folder, then removes
'           : empty subfolders. Used when reverting from @Folder-based organization.
'           : Returns the number of files moved.
'---------------------------------------------------------------------------------------
'
Private Function FlattenSubfolders(strFolder As String) As Long

    Dim oFolder As Scripting.Folder
    Dim oSubFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim colSubFolders As New Collection
    Dim varItem As Variant
    Dim strDestPath As String
    Dim strBaseFolder As String

    strBaseFolder = StripSlash(strFolder)
    If Not FSO.FolderExists(strBaseFolder) Then Exit Function

    Set oFolder = FSO.GetFolder(strBaseFolder)

    ' Collect subfolders (avoid modifying collection during iteration)
    For Each oSubFolder In oFolder.SubFolders
        colSubFolders.Add oSubFolder
    Next oSubFolder

    For Each varItem In colSubFolders
        Set oSubFolder = varItem
        ' Recursively flatten nested subfolders first
        FlattenSubfolders = FlattenSubfolders + FlattenSubfolders(oSubFolder.Path)
        ' Move files from this subfolder to the base folder
        For Each oFile In oSubFolder.Files
            strDestPath = FSO.BuildPath(strBaseFolder, oFile.Name)
            If Not FSO.FileExists(strDestPath) Then
                FSO.MoveFile oFile.Path, strDestPath
                FlattenSubfolders = FlattenSubfolders + 1
            End If
        Next oFile
        ' Remove subfolder if empty
        If oSubFolder.Files.Count = 0 And oSubFolder.SubFolders.Count = 0 Then
            oSubFolder.Delete True
        End If
    Next varItem

End Function
