VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperationPause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOperationPause
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Short synchronous suspension around foreign/user code (hooks, form close).
'           : Distinct from timer-based root continuation.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private m_objOp As clsOperation
Private m_blnActive As Boolean


'---------------------------------------------------------------------------------------
' Procedure : Init
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Begin a pause on the active root.
'---------------------------------------------------------------------------------------
'
Public Sub Init(objOwner As clsOperation)
    If objOwner Is Nothing Then Exit Sub
    Set m_objOp = objOwner
    m_objOp.BeginPauseScope
    m_blnActive = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResumePause
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : End the pause and restore root error trapping.
'---------------------------------------------------------------------------------------
'
Public Sub ResumePause()
    If Not m_blnActive Then Exit Sub
    m_objOp.EndPauseScope
    m_blnActive = False
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    If m_blnActive Then ResumePause
End Sub
