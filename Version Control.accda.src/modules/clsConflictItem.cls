VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConflictItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ItemList As Dictionary
Public Category As String
Public FileName As String
Public ObjectDate As Date
Public IndexDate As Date
Public FileDate As Date
Public Operation As eIndexActionType
Public Resolution As eResolveConflict


'---------------------------------------------------------------------------------------
' Procedure : Resolve
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Resolve the conflict
'---------------------------------------------------------------------------------------
'
Public Function Resolve()

    If Resolution = ercOverwrite Then
        Log.Add "  " & FSO.GetFileName(FileName) & " (Overwrite)", False
    ElseIf Resolution = ercSkip Then
        RemoveFromItemList
        Log.Add "  " & FSO.GetFileName(FileName) & " (Skip)", False
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveFromCollection
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Remove this item from the parent collection
'---------------------------------------------------------------------------------------
'
Private Function RemoveFromItemList()
    Select Case Operation
        Case eatImport
            ' Remove from list of files to import
            If ItemList.Exists(Me.FileName) Then ItemList.Remove (Me.FileName)
        Case eatExport
            ' Remove from object list
            If ItemList.Exists(Me.FileName) Then ItemList.Remove (Me.FileName)
    End Select
End Function

