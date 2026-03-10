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
        VCSIndex.MigrateIndexExtension "Forms", "form"
        VCSIndex.MigrateIndexExtension "Reports", "report"
        VCSIndex.MigrateIndexExtension "Queries", "qdef"
        VCSIndex.MigrateIndexExtension "Macros", "macro"
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
