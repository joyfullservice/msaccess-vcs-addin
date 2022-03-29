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
    eslAggressive    ' Strip out anything that can be reliably rebuilt by Access during Build (themed control colors).

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
' Date      : 3/28/2022
' Purpose   : Wrapper for the VCS class, providing easy API access to VCS functions.
'           : *NOTE* that this class is not persisted. This allows us to wrap up and
'           : remove any object references after the call completes.
'---------------------------------------------------------------------------------------
'
Public Function VCS() As clsVersionControl
    Set VCS = New clsVersionControl
End Function

