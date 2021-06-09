Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last] = 2
End Enum

Public Enum eSanitizeLevel
    eslNone ' Sanitize nothing except for critical items (items that if not sanitized, won't build correctly).
    eslMinimal ' Strip out excess items (like GUIDs) that don't let you see what's up.
    eslRobust ' Strip out anything that can be easily (and reliably) rebuilt by access on the other end (themed control colors).
    eslAgressive ' Strip out most code-change noise. This may introduce small variations on your re-built database in SOME cases.
    
    ' Warning: below introduces sanitzation that may or may not work in all environments, and has known (or highly suspected) edge
    ' cases. Do not use this level on production databases.
    eslToTheBoneBeta ' Cut to the bone, remove all excess stuff. Try out new sanitize features that still have ragged edges.
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
