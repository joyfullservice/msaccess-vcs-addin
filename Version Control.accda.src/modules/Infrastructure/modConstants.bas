Attribute VB_Name = "modConstants"
'---------------------------------------------------------------------------------------
' Module    : modConstants
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Constants and enums used throughout the project, but not exposed publicly.
'           : If constants are only used in a single module, then they should be
'           : declared in that module. If they are shared among multiple modules, they
'           : should be declared here.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Infrastructure")


' VBE Project Name for this Add-In
Public Const PROJECT_NAME As String = "MSAccessVCS"

' Base filename (without extension) for the installed add-in.
' The COM ribbon DLL and worker scripts depend on this exact name.
Public Const ADDIN_BASENAME As String = "Version Control"

' Read/write chunks of text, rather than the whole thing at once for massive
' performance gains when reading large files.
' See https://docs.microsoft.com/is-is/sql/ado/reference/ado-api/readtext-method
Public Const CHUNK_SIZE As Long = 131072    ' (128K)

' Whitespace indentation for outputing JSON content
Public Const JSON_WHITESPACE As Integer = 2

' BOM characters for UTF-8/UTF-16 files
Public Const UTF8_BOM As String = "ï»¿"

' Default hashing algorithm
Public Const DefaultHashAlgorithm As String = "SHA256"

' Sentinel value stored in the VCS index to indicate that a component was
' checked for an @Folder annotation but none was found. Distinguishes
' "no annotation" from "not yet cached" (empty string).
Public Const FOLDER_ANNOTATION_NONE As String = "."

' This value seems to indicate that the theme was not used.
Public Const NO_THEME_INDEX As Integer = -1

' Object types used when determining SQL modification date.
Public Enum eSqlObjectType
    estView
    estStoredProcedure
    estTable
    estTrigger
    estOther
End Enum

' Types of objects that can be exported/imported from a database.
' (Use corresponding constants wherever possible)
' Be careful not to create collisions with two members sharing the
' same value.
Public Enum eDatabaseComponentType
    ' Standard database objects
    edbForm = acForm
    edbMacro = acMacro
    edbModule = acModule
    edbQuery = acQuery
    edbReport = acReport
    edbTableDef = acTable
    edbTableDataMacro = acTableDataMacro
    ' ADP specific
    edbAdpTable
    edbAdpFunction
    edbAdpServerView
    edbAdpStoredProcedure
    edbAdpTrigger
    ' Custom object types we are also handling.
    edbTableData
    edbRelation
    edbDbsProperty
    edbProjectProperty
    edbSharedImage
    edbDocument
    edbHiddenAttribute
    edbSavedSpec
    edbImexSpec
    edbNavPaneGroup
    edbTheme
    edbVbeForm
    edbVbeProject
    edbVbeReference
    edbProject
    edbConnection
    edbCommandBar
    ' Future implementation?
    'edbLinkedTable
    'edbFileProperty
    [_Last]
End Enum

' How a merge can bring a table's data in line with its source file.
Public Enum eTableMergeStrategy
    etmsNone = 0        ' Cannot be merged; the table is left untouched
    etmsReconcile = 1   ' Row by row against the merge key: insert, update, delete
    etmsReload = 2      ' No merge key: replace every row (only when nothing references it)
End Enum

' Database server types for external databases
Public Enum eDatabaseServerType
    estUnknown
    estMsSql
    estMySql
End Enum

' Applications used to interact with source code repository
' (This list is also hard-coded as a drop-down in the options form)
Public Enum eRepositoryApp
    eraUndefined = 0
    eraGitHubDesktop = 1
    eraVSCode = 2
    eraSourceTree = 3
    eraTortoiseGit = 4
End Enum

' Error levels used for logging and monitoring the status
' of the current operation.
Public Enum eErrorLevel
    eelNoError
    eelWarning      ' Logged to file
    eelError        ' Displayed and logged
    eelCritical     ' Cancel operation
End Enum

' Compare mode for cloning dictionary object
' See CloneDictionary function
Public Enum eCompareMethod2
    ecmBinaryCompare = 0
    ecmTextCompare = 1
    ecmDatabaseCompare = 2
    ' Added this to use original compare method
    ecmSourceMethod = 3
End Enum

' VBE Tools > Options > General > Error Trapping
Public Enum eVbeErrorTrapping
    eetBreakOnAllErrors = 0
    eetBreakInClassModule = 1
    eetBreakOnUnhandledErrors = 2
End Enum

' Type of operation in progress
Public Enum eOperationType
    eotExport = 1
    eotBuild = 2
    eotMerge = 3
    eotTestRun = 4   ' VCS.RunTests / clsTestRunner suite
    eotOther = 9     ' Other / catch-all operations
End Enum

' State of operation
Public Enum eOperationState
    eosReady    ' Ready to begin a new operation
    eosStaged   ' Staged and waiting to be resumed
    eosRunning  ' Currently running an operation
End Enum

' Result of operation
Public Enum eOperationResult
    eorUnknown
    eorSuccess
    eorFailed
    eorTimeout
    eorCanceled
End Enum

' Source of the current operation
Public Enum eOperationSource
    eosUserInterface = 0    ' Default - ribbon, menu, immediate window
    eosExternalAPI = 1      ' Application.Run API calls (non-MCP)
    eosMCPTool = 2          ' MCP server with callbacks
End Enum

' Per-connection state during build/import (modConnect)
Public Enum eConnStatus
    csUnknown = 0
    csOK
    csIgnored
End Enum

' User response to a failed connection attempt during build/import
Public Enum eConnFailAction
    cfaRetry
    cfaIgnore
    cfaAbort
End Enum

' Options for resolving file conflicts
Public Enum eResolveConflict
    ercNone
    ercSkip
    ercOverwrite
    ercDelete
End Enum

' Conflict types for import/export conflicts
Public Enum eConflictType
    ectOrphanedSourceFile
    ectNewerDatabaseObject
    ectNewerSourceFile
End Enum

' Release type used when updating version
' or deploying add-in
Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
    Same_Version = 3
End Enum

' Types of objects to include in GetContainers()
Public Enum eContainerFilter
    ecfAllObjects
    ecfVBAItems
    ecfSchemas
End Enum

' Scope for IDbComponent.FileExtensions: indexed (hash/conflict) vs all (move/orphan cleanup)
Public Enum eFileExtensionScope
    efesIndexed = 0   ' Tracked/indexed files (hash, change detection, conflict). Default.
    efesAll = 1       ' All files the component writes (adds derived sidecars, e.g. svg)
End Enum

' Used for handling custom built-in command bar controls. See clsDbCommandBar for details.
Public Const strTemplateCommandBarName As String = "MSAccessVCSCustomBuiltinCommandBarTemplate"

' Export format versions using packed integers (Major * 10000 + Minor * 100 + Patch)
' Used to gate export behavior changes so users can upgrade on their own schedule.
' When adding a member here, add a matching line to GetExportFormatVersions below.
' (modTestExportFormat fails if the enum and that list disagree.)
Public Enum eExportFormatVersion
    EFV_4_1_2 = 40102
    EFV_5_0_0 = 50000      ' v5 baseline: extensions, @Folder, CF decode-to-JSON, command bar replica export, etc.
    EFV_5_1_0 = 50100      ' Sidecar Info.Class names; canonical tbldefs property order; IMEX spec SpecID.
End Enum

' Bump this whenever the SVG layout generator output changes
' to force regeneration of cached SVG files.
Public Const LAYOUT_SVG_GENERATOR_VERSION As Long = 1

' Revisions for the external database schema exporters (clsSchemaMsSql, clsSchemaMySql).
' Bump when a DDL-shape fix changes exported output. Folded into the per-schema state
' fingerprint stored in the index, which forces a one-time re-export of that schema.
' Schema exports do not participate in CategoryHashes, so GetExporterRevisions does
' not apply to them. History (one line per bump):
'   MsSql = 1  7/27/2026  Initial fingerprint (also tracks sp_GetDDL availability)
'   MySql = 1  7/27/2026  Initial fingerprint
Public Const SCHEMA_EXPORTER_REVISION_MSSQL As Long = 1
Public Const SCHEMA_EXPORTER_REVISION_MYSQL As Long = 1

' Used for ImportCommandBars function; negative/zero result should be treated as an error.
Public Enum eImportCommandBarsResult
    eicImportedNotVerified = -1
    eicFailed = 0
    eicImportedVerified = 1
    eicImportedUnableToVerify = 2
End Enum


'---------------------------------------------------------------------------------------
' Function  : GetExportFormatVersions
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Selectable export format versions, in ascending order. Add one line here
'           : for each new eExportFormatVersion member. The options form combo box and
'           : LatestExportFormat both read this list, so nothing else needs updating.
'---------------------------------------------------------------------------------------
'
Public Function GetExportFormatVersions() As Collection

    Dim col As Collection
    Set col = New Collection
    col.Add EFV_4_1_2
    col.Add EFV_5_0_0
    col.Add EFV_5_1_0
    Set GetExportFormatVersions = col

End Function


'---------------------------------------------------------------------------------------
' Function  : LatestExportFormat
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : The newest export format version (last entry in GetExportFormatVersions).
'---------------------------------------------------------------------------------------
'
Public Function LatestExportFormat() As Long

    Dim col As Collection

    Set col = GetExportFormatVersions
    LatestExportFormat = col.Item(col.Count)

End Function


'---------------------------------------------------------------------------------------
' Function  : GetExporterRevisions
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Per-category exporter revisions. Bump when a bug fix changes a category's
'           : exported output in a way the change index cannot otherwise detect (sidecar
'           : files, or components whose IsModified relies on DateModified). This forces
'           : a one-time full re-export of that category via CategoryHashes invalidation.
'           : NOT for opt-in output changes (use eExportFormatVersion) and NOT needed for
'           : content-hashed primary output (IsModified self-heals).
'           : History (one line per bump):
'           :   CommandBars = 1  7/27/2026  Fixed _Images sidecar export
'           :   Forms = 1        7/28/2026  Fixed CF14 data bar decode (issue #730)
'           :   Reports = 1      7/28/2026  Fixed CF14 data bar decode (issue #730)
'---------------------------------------------------------------------------------------
'
Public Function GetExporterRevisions() As Dictionary

    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "CommandBars", 1
    d.Add "Forms", 1
    d.Add "Reports", 1
    Set GetExporterRevisions = d

End Function
