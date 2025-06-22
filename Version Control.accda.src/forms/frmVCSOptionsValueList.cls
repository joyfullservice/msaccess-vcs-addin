Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_Data As ADODB.Recordset

Public Event CloseOptionList()

Public OptionValueListCode As String

Public Sub SetData(ByRef Values() As String)

    Dim i As Long

    Set m_Data = New ADODB.Recordset
    m_Data.CursorLocation = adUseClient
    m_Data.CursorType = adOpenKeyset
    m_Data.LockType = adLockOptimistic
   ' m_Data.Fields.Append "id", adInteger, 8, adFldRowID
    m_Data.Fields.Append "OptionValue", adVarChar, 255, adFldUpdatable + adFldIsNullable
    m_Data.Open

    If UBound(Values) >= 0 Then
        For i = LBound(Values) To UBound(Values)
            m_Data.AddNew
            ' m_Data.Fields(0).Value = i
            m_Data.Fields(0).Value = Values(i)
            m_Data.Update
        Next i
        m_Data.MoveFirst
    End If

    Set Me.Recordset = m_Data

End Sub

Public Function GetData() As String()

    Dim Values() As String
    Dim rs As ADODB.Recordset

    If m_Data.RecordCount = 0 Then
        ReDim Values(0)
        GetData = Values
        Exit Function
    End If

    Set rs = m_Data.Clone

    rs.MoveLast
    rs.MoveFirst
    ReDim Values(rs.RecordCount - 1)

    Dim i As Long
    Do While Not rs.EOF
        Values(i) = rs.Fields(0).Value
        rs.MoveNext
        i = i + 1
    Loop

    GetData = Values

End Function

Private Sub Form_Current()
    If Me.NewRecord Then
        With Me.txtValueInsert
            .Value = Me.txtValue.Value
            .SetFocus
        End With
    End If
End Sub

Private Sub txtValueInsert_Change()
    If Me.NewRecord Then
        m_Data.AddNew
        m_Data.Fields(0).Value = Me.txtValueInsert.Text
        Me.Recordset.Bookmark = m_Data.Bookmark

        Me.txtValue.SetFocus
        Me.txtValue.Text = Me.txtValueInsert.Value
        Me.txtValue.SelStart = Len(Me.txtValue.Text)
    End If
End Sub

Private Sub cmdClose_Click()
    RaiseEvent CloseOptionList
End Sub
