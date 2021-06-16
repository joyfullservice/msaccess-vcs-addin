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

Public Enum eSanitizeLevel
    eslNone = 0     ' Sanitize only items which won't build correctly unless you sanitize them.
    eslBasic        ' Strip out excess items (like GUIDs) that are just noise and no effect can be found.
    eslAgressive    ' Strip out anything that can be reliably rebuilt by Access during Build (themed control colors).

    ' WARNING: AdvancedBeta introduces sanitzation that may or may not work in all environments, and has known
    '          (or highly suspected) edge cases where it does not always operate correctly. Do not use this level in
    '          production databases.
    eslAdvancedBeta ' Remove all excess noise. Try out new sanitize features that still have ragged edges.
    [_Last]         ' DO NOT REMOVE: This is a "Fake" level, and must be at the end.
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
