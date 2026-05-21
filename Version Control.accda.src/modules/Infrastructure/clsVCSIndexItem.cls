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
'@Folder("Infrastructure")

Public ImportDate As Date
Public ExportDate As Date
Public SourceModified As Date
Public FileHash As String
Public OtherHash As String
Public ObjectDate As Date
Public FilePropertiesHash As String
Public MetaHash As String
Public FolderAnnotation As String


'---------------------------------------------------------------------------------------
' Procedure : SyncDate
' Author    : Adam Waller
' Date      : 5/20/2026
' Purpose   : Returns the date we last synced with the source file, whether by exporting
'           : to it or importing from it. Used by conflict detection to determine if the
'           : source file has been modified externally since we last touched it.
'---------------------------------------------------------------------------------------
'
Public Property Get SyncDate() As Date
    SyncDate = Largest(ExportDate, ImportDate)
End Property
