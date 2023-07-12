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

Public CategoryList As Dictionary
Public CategoryName As String
Public ItemKey As String
Public FileName As String
Public ObjectDate As Date
Public IndexDate As Date
Public FileDate As Date
Public Operation As eIndexOperationType
Public ActionType As eResolveConflict
Public Resolution As eResolveConflict
Public Suggestion As eResolveConflict


'---------------------------------------------------------------------------------------
' Procedure : Resolve
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Resolve the conflict
'---------------------------------------------------------------------------------------
'
Public Function Resolve(dCategories As Dictionary)

    Select Case Me.Resolution
        Case ercOverwrite
            If Me.ActionType = ercDelete Then
                Log.Add "  " & FSO.GetFileName(FileName) & " (Delete)", False
                DeleteFile Me.FileName
            ElseIf Me.ActionType = ercOverwrite Then
                Log.Add "  " & FSO.GetFileName(FileName) & " (Overwrite)", False
                ' Leave in list of files to overwrite.
            End If
        Case ercSkip
            RemoveFromItemList
            Log.Add "  " & FSO.GetFileName(FileName) & " (Skip)", False
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveFromCollection
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Remove this item from the parent collection of items to import/export
'---------------------------------------------------------------------------------------
'
Private Function RemoveFromItemList()
    If CategoryList.Exists(Me.CategoryName) Then
        If CategoryList(Me.CategoryName).Exists("Objects") Then
            With CategoryList(Me.CategoryName)("Objects")
                If .Exists(Me.ItemKey) Then
                    Select Case Operation
                        Case eatImport
                            ' Remove from list of files to import
                            .Remove Me.FileName
                        Case eatExport
                            ' Remove from object list
                            .Remove Me.ItemKey
                    End Select
                    ' Remove category if no more objects of this type,
                    ' or if this type exports a single file.
                    If .Count = 0 Or CategoryList(Me.CategoryName)("Class").SingleFile Then
                        CategoryList.Remove Me.CategoryName
                    End If
                End If
            End With
        End If
    End If
End Function
