'---------------------------------------------------------------------------------------
' Module    : modConstants
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Constants and enums used throughout the project, but not exposed publicly.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


Public Const JSON_WHITESPACE As Integer = 2
Public Const UTF8_BOM As String = "ï»¿"
Public Const UCS2_BOM As String = "ÿþ"

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
    edbForm
    edbMacro
    edbModule
    edbQuery
    edbReport
    edbTableDef
    edbTableDataMacro
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
End Enum

' Error levels used for logging and monitoring the status
' of the current operation.
Public Enum eErrorLevel
    eelNoError
    eelWarning      ' Logged to file
    eelError        ' Displayed and logged
    eelCritical     ' Cancel operation
End Enum
