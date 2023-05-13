Attribute VB_Name = "modConstants"
'---------------------------------------------------------------------------------------
' Module    : modConstants
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Constants and enums used throughout the project, but not exposed publicly.
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
    edbLinkedTable
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
    edbFileProperty
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
End Enum
