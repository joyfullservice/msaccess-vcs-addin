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