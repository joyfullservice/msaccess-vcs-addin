VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStandardLetterCasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsStandardLetterCasing
' Author    : Mike Wolfe (minor tweaks by Adam Waller)
' Date      : 10/29/2025
' Purpose   : A simple integrated solution for enforcing variable casing in VBA.
'           : See: https://nolongerset.com/standardizelettercasing/
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")


' Set letter-casing for all other identifiers here:
Dim i           'i
Dim strSql      'strSql
Dim lngAs       'lngAs
Dim strId       'strId
Dim lngId       'lngId
Dim Id          'Id
Dim CleanUp     'CleanUp
Dim strXml      'strXml
