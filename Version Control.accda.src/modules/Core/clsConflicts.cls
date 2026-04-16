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
'@Folder("Core")

' Status to approve and continue.
Public ApproveResolutions As Boolean

' Set to True when the user explicitly cancels the conflict dialog.
' Used by callers to distinguish a user-initiated cancel from other early exits
' (e.g. critical errors) so they can skip persisting index state on cancel.
Public UserCanceled As Boolean

Private m_intOperationType As eIndexOperationType
Private m_dCategories As Dictionary
Private m_dItems As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : ShowDialog
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Show the merge conflict dialog
'---------------------------------------------------------------------------------------
'
Public Sub ShowDialog()
    SaveToTable
    With DoCmd
        ' The following gives a smoother load when using the datasheet subform.
        .Hourglass True
        '.OpenForm "frmVCSConflict", , , , , acHidden
        DoEvents
        .Hourglass False
        ' Open the form in dialog mode so that we return a status when it closes.
        MCP.Log "Waiting for user to resolve " & Me.Count & " conflict(s)..."
        Perf.PauseTiming
        .OpenForm "frmVCSConflict", , , , , acDialog
        UpdateResolutionsFromTable
        Perf.ResumeTiming
        If Me.ApproveResolutions Then
            MCP.Log "User resolved conflicts, continuing operation..."
        Else
            Me.UserCanceled = True
            MCP.Log "User canceled conflict resolution."
        End If
    End With
    ClearTable
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveOrPrompt
' Author    : Adam Waller
' Date      : 4/15/2026
' Purpose   : For agent/API operations, auto-resolve all conflicts (source is truth for
'           : imports, database is truth for exports). For user-initiated operations,
'           : show the interactive conflict dialog as usual.
'---------------------------------------------------------------------------------------
'
Public Sub ResolveOrPrompt()

    Dim varKey As Variant
    Dim cItem As clsConflictItem

    If Operation.Source = eosMCPTool Or Operation.Source = eosExternalAPI Then
        For Each varKey In m_dItems.Keys
            Set cItem = m_dItems(varKey)
            If cItem.ActionType <> ercNone Then
                cItem.Resolution = cItem.ActionType
            Else
                cItem.Resolution = ercOverwrite
            End If
        Next varKey
        Log.Add T("Auto-resolved {0} conflict(s) for agent operation.", _
            var0:=Me.Count), False
        MCP.Log "Auto-resolved " & Me.Count & " conflict(s)."
        Me.ApproveResolutions = True
    Else
        ShowDialog
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Resolve
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Resolve the conflicts, removing any objects that should be skipped in the
'           : import/export process.
'---------------------------------------------------------------------------------------
'
Public Sub Resolve()
    Dim cItem As clsConflictItem
    Dim varKey As Variant
    For Each varKey In m_dItems.Keys
        Set cItem = m_dItems(varKey)
        cItem.Resolve
        ' Clear reference to categories dictionary
        Set cItem.CategoryList = Nothing
    Next varKey
    ' Clear additional reference to list
    Set m_dCategories = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Add an item to the conflict list.
'---------------------------------------------------------------------------------------
'
Public Sub Add(cComponent As IDbComponent, strKey As String, dteIndexDate As Date, dteFileDate As Date, intActionType As eResolveConflict, _
    Optional strSourceFile As String, Optional intSuggestedResolution As eResolveConflict)

    Dim cItem As clsConflictItem

    Set cItem = New clsConflictItem
    With cItem
        .Operation = m_intOperationType
        .CategoryName = cComponent.Category
        .ItemKey = strKey
        .SingleFile = cComponent.SingleFile
        .FileName = Nz2(strSourceFile, cComponent.SourceFile)
        .ObjectDate = DateTruncToSeconds(cComponent.DateModified)
        .IndexDate = dteIndexDate
        .FileDate = dteFileDate
        .ActionType = intActionType
        .Suggestion = intSuggestedResolution
        Set .CategoryList = m_dCategories
        m_dItems.Add .FileName, cItem
    End With

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
    Dim varKey As Variant
    Dim cItem As clsConflictItem

    ' Clear any existing records first
    ClearTable
    If m_dItems.Count = 0 Then Exit Sub

    ' Open table to add items
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblConflicts")
    With rst

        ' Loop through items
        For Each varKey In m_dItems.Keys
            Set cItem = m_dItems(varKey)
            .AddNew
                !Component = cItem.CategoryName
                !ItemKey = cItem.ItemKey
                ' Save the file name, including object type folder
                !FileName = Mid$(cItem.FileName, Len(Options.GetExportFolder) + 1)
                !ObjectDate = ZN(cItem.ObjectDate)
                !IndexDate = ZN(cItem.IndexDate)
                !FileDate = ZN(cItem.FileDate)
                !Suggestion = cItem.Suggestion
                !Resolution = cItem.Resolution
                ' Set default resolution, if provided
                If cItem.Resolution = ercNone Then
                    If cItem.Suggestion <> ercNone Then
                        !Resolution = cItem.Suggestion
                    ElseIf cItem.ActionType <> ercNone Then
                        !Resolution = cItem.ActionType
                    End If
                End If
            .Update
        Next varKey
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpdateResolutionsFromTable
' Author    : Adam Waller
' Date      : 11/1/2021
' Purpose   : Update the class items with the specified resolution set by the user
'           : in the work table.
'---------------------------------------------------------------------------------------
'
Private Sub UpdateResolutionsFromTable()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset

    If m_dItems.Count = 0 Then Exit Sub

    ' Open table to read records
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblConflicts")
    With rst
        Do While Not .EOF
            ' Update matching class objects
            If m_dItems.Exists(Nz(!ItemKey)) Then
                m_dItems(Nz(!ItemKey)).Resolution = Nz(!Resolution, 0)
            End If
            .MoveNext
        Loop
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
    Count = m_dItems.Count
End Function


'---------------------------------------------------------------------------------------
' Procedure : Reset
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Reset the class, clearing any existing conflicts
'           : Preserve reference to categories dictionary so we can remove items later.
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(ByRef dCategories As Dictionary, intOperationType As eIndexOperationType)
    Set m_dCategories = dCategories
    Set m_dItems = New Dictionary
    m_intOperationType = intOperationType
    Me.ApproveResolutions = False
    Me.UserCanceled = False
End Sub
