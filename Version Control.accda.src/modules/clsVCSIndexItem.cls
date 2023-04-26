VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSIndexItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVCSIndexItem
' Author    : Adam Waller
' Date      : 9/23/2021
' Purpose   : An index item, representing a single database object and corresponding
'           : source file.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public dParent As Dictionary
Public ImportDate As Date
Public ExportDate As Date
Public SourceModified As Date
Public FileHash As String
Public OtherHash As String
Public FilePropertiesHash As String

