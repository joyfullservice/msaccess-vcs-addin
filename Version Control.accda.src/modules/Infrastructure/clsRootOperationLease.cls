VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRootOperationLease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsRootOperationLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Exclusive ownership handle for a root operation. Only the lease holder
'           : (or a validated timer continuation) may complete the root.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private m_objOp As clsOperation
Private m_strToken As String
Private m_blnValid As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Called by clsOperation when issuing a new root lease. The lease talks back
'           : to the instance that issued it, never to the session singleton by name, so
'           : tests can exercise lifecycle rules on a throwaway operation instance.
'---------------------------------------------------------------------------------------
'
Public Sub Init(objOwner As clsOperation, strToken As String)
    Set m_objOp = objOwner
    m_strToken = strToken
    m_blnValid = (Len(strToken) > 0) And Not (objOwner Is Nothing)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Token
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Opaque root token (for timer continuations only).
'---------------------------------------------------------------------------------------
'
Public Property Get Token() As String
    Token = m_strToken
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsValid
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : True while this lease may still complete or detach the root.
'---------------------------------------------------------------------------------------
'
Public Property Get IsValid() As Boolean
    IsValid = m_blnValid And m_objOp.RootLeaseIsActive(m_strToken)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Complete
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Complete the root operation exactly once.
'---------------------------------------------------------------------------------------
'
Public Sub Complete(Optional intResult As eOperationResult = eorUnknown)
    If Not m_blnValid Then Exit Sub
    m_objOp.CompleteRootLease m_strToken, intResult
    m_blnValid = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DetachForContinuation
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Transfer root ownership across an asynchronous timer boundary.
'---------------------------------------------------------------------------------------
'
Public Function DetachForContinuation() As String
    If Not m_blnValid Then Exit Function
    DetachForContinuation = m_objOp.DetachRootLease(m_strToken)
    m_blnValid = False
End Function


Private Sub Class_Terminate()
    On Error Resume Next
    If m_objOp Is Nothing Then Exit Sub
    If m_blnValid And m_objOp.RootLeaseIsActive(m_strToken) Then
        m_objOp.AbandonRootLease m_strToken, "clsRootOperationLease.Class_Terminate"
    End If
End Sub
