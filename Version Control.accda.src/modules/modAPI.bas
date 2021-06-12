Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Note, some enums are listed here when they are directly exposed
' through the Options class. (Allowing them to be used externally)

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last] = 2
End Enum

Public Enum eSanitizeColors
    escNone = 0     ' Do not remove any color information
    escBasic        ' Remove dynamic theme colors when clearly safe to do so
    escAdvanced     ' Remove additional color values that are likely to be dynamic theme colors
                    ' (May remove some static colors if object definition blocks are incomplete,
                    '  which may occur in databases upgraded from prior versions of Access.)
    [_Last]
End Enum

Private m_VCS As clsVersionControl


'---------------------------------------------------------------------------------------
' Procedure : VCS
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Wrapper for the VCS class, providing easy API access to VCS functions.
'---------------------------------------------------------------------------------------
'
Public Function VCS() As clsVersionControl
    If m_VCS Is Nothing Then Set m_VCS = New clsVersionControl
    Set VCS = m_VCS
End Function


'---------------------------------------------------------------------------------------
' Procedure : RepairColors
' Author    : Adam Waller
' Date      : 6/8/2021
' Purpose   : Reapply the color properties in the current database
'---------------------------------------------------------------------------------------
'
Public Sub RepairColors()
    RepairColorDefinitionBlocks
End Sub
