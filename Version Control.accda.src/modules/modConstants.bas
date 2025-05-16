﻿Attribute VB_Name = "modConstants"
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


' VBE Project Name for this Add-In
Public Const PROJECT_NAME As String = "MSAccessVCS"

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

' Type of operation in progress
Public Enum eOperationType
    eotExport = 1
    eotBuild = 2
    eotMerge = 3
    eotOther = 4
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

' Used for handling custom built-in command bar controls. See clsDbCommandBar for details.
Public Const strTemplateCommandBarName As String = "MSAccessVCSCustomBuiltinCommandBarTemplate"

' Used for ImportCommandBars function; negative/zero result should be treated as an error.
Public Enum eImportCommandBarsResult
    eicImportedNotVerified = -1
    eicFailed = 0
    eicImportedVerified = 1
    eicImportedUnableToVerify = 2
End Enum
