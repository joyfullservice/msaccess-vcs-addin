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
Public SingleFile As Boolean
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
Public Sub Resolve()

    Select Case Me.Operation

        ' Merge operation (import)
        Case eatImport
            Select Case Me.Resolution
                Case ercDelete
                    ' Delete orphaned object by merging it from a non-existent source
                    ' file. (Object deletion is handled during the merge)
                    ' Leave in list of objects to merge
                    Log.Add "  " & FSO.GetFileName(FileName) & " (Delete)", False
                Case ercOverwrite
                    ' Leave in list of objects to overwrite
                    Log.Add "  " & FSO.GetFileName(FileName) & " (Overwrite)", False
                Case ercSkip
                    ' Don't delete this item
                    RemoveFromItemList
                    Log.Add "  " & FSO.GetFileName(FileName) & " (Skip)", False
            End Select

        ' Export operation
        Case eatExport, eatAltExport
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
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveFromCollection
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Remove this item from the parent collection of items to import/export
'---------------------------------------------------------------------------------------
'
Private Function RemoveFromItemList()

    Dim lngCount As Long

    If CategoryList.Exists(Me.CategoryName) Then
        ' Check operation type to see if we are dealing with objects or source files
        Select Case Operation

            Case eatImport  ' Merge
                If CategoryList(Me.CategoryName).Exists("Files") Then
                    With CategoryList(Me.CategoryName)("Files")
                        If .Exists(Me.FileName) Then
                            ' Remove from list of files to import
                            .Remove Me.FileName
                        End If
                        lngCount = .Count
                    End With
                End If

            Case eatExport
                If CategoryList(Me.CategoryName).Exists("Objects") Then
                    With CategoryList(Me.CategoryName)("Objects")
                        If .Exists(Me.ItemKey) Then
                            ' Remove from object list
                            .Remove Me.ItemKey
                        End If
                        lngCount = .Count
                    End With
                End If
        End Select

        ' Remove category if no more objects of this type,
        ' or if this type exports a single file.
        If lngCount = 0 Or CategoryList(Me.CategoryName)("Class").SingleFile Then
            CategoryList.Remove Me.CategoryName
        End If
    End If

End Function
