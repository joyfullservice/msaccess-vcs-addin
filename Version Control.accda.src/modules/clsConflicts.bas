VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConflicts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsConflicts
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : This class is used to manage the conflicts during import/export
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


Private m_Items As Collection


'---------------------------------------------------------------------------------------
' Procedure : ShowDialog
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Show the merge conflict dialog
'---------------------------------------------------------------------------------------
'
Public Sub ShowDialog()
    SaveToTable
    DoCmd.OpenForm "frmVCSConflict", , , , , acDialog
    ClearTable
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Add an item to the conflict list.
'---------------------------------------------------------------------------------------
'
Public Sub Add(cComponent As IDbComponent, dteIndexDate As Date, dteFileDate As Date)
    Dim cItem As clsConflictItem
    Set cItem = New clsConflictItem
    With cItem
        .Category = cComponent.Category
        .FileName = cComponent.SourceFile
        .ObjectDate = cComponent.DateModified
        .IndexDate = dteIndexDate
        .FileDate = dteFileDate
    End With
    m_Items.Add cItem
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveToTable
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Save collection to table
'---------------------------------------------------------------------------------------
'
Private Sub SaveToTable()

    Dim dbs As Database
    Dim rst As Recordset
    Dim cItem As clsConflictItem
    
    ' Clear any existing records first
    ClearTable
    If m_Items.Count = 0 Then Exit Sub
    
    ' Open table to add items
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblConflicts")
    With rst

        ' Loop through items
        For Each cItem In m_Items
            .AddNew
                !Component = cItem.Category
                !FileName = Mid$(cItem.FileName, Len(Options.GetExportFolder))
                !ObjectDate = cItem.ObjectDate
                !IndexDate = cItem.IndexDate
                !FileDate = cItem.FileDate
                !Resolution = cItem.Resolution
            .Update
        Next cItem
        .Close
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTable
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Clear the records from the conflicts table.
'---------------------------------------------------------------------------------------
'
Private Sub ClearTable()
    Dim dbs As Database
    Set dbs = CodeDb
    dbs.Execute "delete from tblConflicts", dbFailOnError
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Return count of conflicts
'---------------------------------------------------------------------------------------
'
Public Function Count() As Long
    Count = m_Items.Count
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Create new collection of items
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub
