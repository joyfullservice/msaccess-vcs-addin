VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubRdAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Tests"
Option Compare Database
Option Explicit

Private Const OUTCOME_FAIL As Integer = 0
Private Const OUTCOME_PASS As Integer = 1
Private Const OUTCOME_INCONCLUSIVE As Integer = -1

Private m_Real As Object
Private m_RealTried As Boolean

Public Sub AreEqual(ByVal Expected As Variant, ByVal Actual As Variant, _
                    Optional ByVal Message As String = vbNullString)
    If Route(EvalEqual(Expected, Actual), Message) Then Exit Sub
    RealAssert.AreEqual Expected, Actual, Message
End Sub

Public Sub AreNotEqual(ByVal Expected As Variant, ByVal Actual As Variant, _
                       Optional ByVal Message As String = vbNullString)
    If Route(EvalNotEqual(Expected, Actual), Message) Then Exit Sub
    RealAssert.AreNotEqual Expected, Actual, Message
End Sub

Public Sub AreSame(ByVal Expected As Variant, ByVal Actual As Variant, _
                   Optional ByVal Message As String = vbNullString)
    If Route(EvalSame(Expected, Actual), Message) Then Exit Sub
    RealAssert.AreSame Expected, Actual, Message
End Sub

Public Sub AreNotSame(ByVal Expected As Variant, ByVal Actual As Variant, _
                      Optional ByVal Message As String = vbNullString)
    If Route(EvalNotSame(Expected, Actual), Message) Then Exit Sub
    RealAssert.AreNotSame Expected, Actual, Message
End Sub

Public Sub IsTrue(ByVal condition As Boolean, _
                  Optional ByVal Message As String = vbNullString)
    If Route(IIf(condition, OUTCOME_PASS, OUTCOME_FAIL), Message) Then Exit Sub
    RealAssert.IsTrue condition, Message
End Sub

Public Sub IsFalse(ByVal condition As Boolean, _
                   Optional ByVal Message As String = vbNullString)
    If Route(IIf(condition, OUTCOME_FAIL, OUTCOME_PASS), Message) Then Exit Sub
    RealAssert.IsFalse condition, Message
End Sub

Public Sub IsNothing(ByVal Value As Variant, _
                     Optional ByVal Message As String = vbNullString)
    If Route(EvalIsNothing(Value, True), Message) Then Exit Sub
    RealAssert.IsNothing Value, Message
End Sub

Public Sub IsNotNothing(ByVal Value As Variant, _
                        Optional ByVal Message As String = vbNullString)
    If Route(EvalIsNothing(Value, False), Message) Then Exit Sub
    RealAssert.IsNotNothing Value, Message
End Sub

Public Sub Succeed()
    If Route(OUTCOME_PASS, vbNullString) Then Exit Sub
    RealAssert.Succeed
End Sub

Public Sub Fail(Optional ByVal Message As String = vbNullString)
    If Route(OUTCOME_FAIL, Message) Then Exit Sub
    RealAssert.Fail Message
End Sub

Public Sub Inconclusive(Optional ByVal Message As String = vbNullString)
    If Route(OUTCOME_INCONCLUSIVE, Message) Then Exit Sub
    RealAssert.Inconclusive Message
End Sub

Private Function Route(ByVal intOutcome As Integer, ByVal Message As String) As Boolean
    If ReportedToVcs(intOutcome, Message) Then
        Route = True
    ElseIf RealAssert Is Nothing Then
        Debug.Assert (intOutcome <> OUTCOME_FAIL)
        Route = True
    End If
End Function

Private Function ReportedToVcs(ByVal intOutcome As Integer, _
                               ByVal Message As String) As Boolean
    On Error Resume Next
    ReportedToVcs = Application.Run(AddInPath() & ".HandleTestOutcome", _
                                    VcsOutcome(intOutcome), CVar(Message))
    On Error GoTo 0
End Function

Private Function VcsOutcome(ByVal intOutcome As Integer) As Long
    Select Case intOutcome
        Case OUTCOME_INCONCLUSIVE: VcsOutcome = 3
        Case OUTCOME_FAIL: VcsOutcome = 2
        Case Else: VcsOutcome = 1
    End Select
End Function

Private Function AddInPath() As String
    AddInPath = VBA.GetSetting("MSAccessVCS", "Install", "Install Folder", _
                               Environ$("AppData") & "\MSAccessVCS") & "\Version Control"
End Function

Private Function RealAssert() As Object
    If m_Real Is Nothing And Not m_RealTried Then
        m_RealTried = True
        On Error Resume Next
        Set m_Real = CreateObject("Rubberduck.AssertClass")
        On Error GoTo 0
    End If
    Set RealAssert = m_Real
End Function

Private Function EvalEqual(ByVal Expected As Variant, ByVal Actual As Variant) As Integer
    If IsObject(Expected) Or IsObject(Actual) Then
        EvalEqual = OUTCOME_INCONCLUSIVE
        Exit Function
    End If
    If VarType(Expected) = vbNull And VarType(Actual) = vbNull Then
        EvalEqual = OUTCOME_PASS
    ElseIf VarType(Expected) <> VarType(Actual) Then
        EvalEqual = OUTCOME_INCONCLUSIVE
    Else
        EvalEqual = IIf(Expected = Actual, OUTCOME_PASS, OUTCOME_FAIL)
    End If
End Function

Private Function EvalNotEqual(ByVal Expected As Variant, ByVal Actual As Variant) As Integer
    Dim intEqual As Integer
    intEqual = EvalEqual(Expected, Actual)
    If intEqual = OUTCOME_INCONCLUSIVE Then
        EvalNotEqual = OUTCOME_INCONCLUSIVE
    Else
        EvalNotEqual = IIf(intEqual = OUTCOME_PASS, OUTCOME_FAIL, OUTCOME_PASS)
    End If
End Function

Private Function EvalSame(ByVal Expected As Variant, ByVal Actual As Variant) As Integer
    If IsObject(Expected) And IsObject(Actual) Then
        EvalSame = IIf(Expected Is Actual, OUTCOME_PASS, OUTCOME_FAIL)
    Else
        EvalSame = OUTCOME_INCONCLUSIVE
    End If
End Function

Private Function EvalNotSame(ByVal Expected As Variant, ByVal Actual As Variant) As Integer
    Dim intSame As Integer
    intSame = EvalSame(Expected, Actual)
    If intSame = OUTCOME_INCONCLUSIVE Then
        EvalNotSame = OUTCOME_INCONCLUSIVE
    Else
        EvalNotSame = IIf(intSame = OUTCOME_PASS, OUTCOME_FAIL, OUTCOME_PASS)
    End If
End Function

Private Function EvalIsNothing(ByVal Value As Variant, _
                               ByVal blnWantNothing As Boolean) As Integer
    Dim blnIsNothing As Boolean
    If IsObject(Value) Then blnIsNothing = (Value Is Nothing)
    EvalIsNothing = IIf(blnIsNothing = blnWantNothing, OUTCOME_PASS, OUTCOME_FAIL)
End Function
