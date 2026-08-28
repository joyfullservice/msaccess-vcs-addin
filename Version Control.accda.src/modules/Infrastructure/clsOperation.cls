VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsOperation
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Root operation ownership, validated continuations, and pause scopes.
'           : Exactly one root controls registry state, heartbeat, logs, callbacks, and
'           : teardown. Core work never begins or finishes an operation directly.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")


Private Const HEARTBEAT_TIMEOUT_SECONDS As Integer = 60 * 10
Private Const REG_KEY_ROOT_TOKEN As String = "RootToken"
Private Const REG_KEY_ROOT_TYPE As String = "RootOperationType"


Public Heartbeat As Date
Public ErrorLevel As eErrorLevel
Public Result As eOperationResult
Public Source As eOperationSource

' Set by headless entry points before TryBeginRoot/Begin so Attended stays False even
' when Source is still eosUserInterface (Immediate Window callers).
Public ForceUnattended As Boolean


Private Type udtRoot
    Token As String
    OperationType As eOperationType
    RequestedInteractionMode As eInteractionMode
    Attended As Boolean
    StartTime As Date
    FinishTime As Date
    State As eOperationState
    AppHwnd As Long
    ErrorTrapping As eVbeErrorTrapping
    Detached As Boolean
    Completed As Boolean
    PauseDepth As Long
End Type
Private m_Root As udtRoot

Private m_eimSticky As eInteractionMode


'---------------------------------------------------------------------------------------
' Procedure : TryBeginRoot
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Acquire exclusive root ownership. Returns Nothing when refused.
'---------------------------------------------------------------------------------------
'
Public Function TryBeginRoot(intType As eOperationType) As clsRootOperationLease

    Dim cLease As clsRootOperationLease

    If Not CanBeginRoot Then Exit Function

    InitializeRoot intType
    Set cLease = New clsRootOperationLease
    cLease.Init Me, m_Root.Token
    Set TryBeginRoot = cLease

End Function


'---------------------------------------------------------------------------------------
' Procedure : Begin
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Synchronous root entry for work that begins and ends in one procedure.
'           : Wraps TryBeginRoot and discards the lease handle; the caller completes
'           : with Finish on the same stack. Use TryBeginRoot when ownership crosses a
'           : timer, a form lifetime, or a procedure boundary instead.
'---------------------------------------------------------------------------------------
'
Public Function Begin(intType As eOperationType) As Boolean

    Dim cLease As clsRootOperationLease
    Set cLease = TryBeginRoot(intType)
    Begin = (Not cLease Is Nothing)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResumeRoot
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Resume a detached root from a timer continuation. Fails closed on mismatch.
'---------------------------------------------------------------------------------------
'
Public Function ResumeRoot(strToken As String) As clsRootOperationLease

    Dim cLease As clsRootOperationLease

    If Len(strToken) = 0 Then Exit Function
    If m_Root.Completed Then Exit Function
    If m_Root.State <> eosStaged Then Exit Function
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Function
    If m_Root.AppHwnd <> hWndAccessApp Then Exit Function

    m_Root.Detached = False
    m_Root.State = eosRunning
    SetErrorTrapping
    UpdateRegistry
    Pulse

    Set cLease = New clsRootOperationLease
    cLease.Init Me, m_Root.Token
    Set ResumeRoot = cLease

End Function


'---------------------------------------------------------------------------------------
' Procedure : CompleteRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Complete the root when the caller holds a matching lease token.
'---------------------------------------------------------------------------------------
'
Public Sub CompleteRootLease(strToken As String, Optional intResult As eOperationResult = eorUnknown)

    If m_Root.Completed Then Exit Sub
    If Len(m_Root.Token) = 0 Then Exit Sub
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Sub

    FinishRootInternal intResult

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DetachRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Stage the root for an asynchronous continuation.
'---------------------------------------------------------------------------------------
'
Public Function DetachRootLease(strToken As String) As String

    If m_Root.Completed Then Exit Function
    If Len(m_Root.Token) = 0 Then Exit Function
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Function
    If m_Root.State <> eosRunning Then Exit Function

    m_Root.Detached = True
    m_Root.State = eosStaged
    RestoreUserErrorTrapping
    UpdateRegistry
    Pulse
    DetachRootLease = m_Root.Token

End Function


'---------------------------------------------------------------------------------------
' Procedure : AbandonRootLease
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Safety-net cleanup when a lease is dropped without Complete.
'---------------------------------------------------------------------------------------
'
Public Sub AbandonRootLease(strToken As String, strSource As String)

    If m_Root.Completed Then Exit Sub
    If Len(m_Root.Token) = 0 Then Exit Sub
    If StrComp(m_Root.Token, strToken, vbBinaryCompare) <> 0 Then Exit Sub

    LogDiagnostic "Root lease abandoned", strSource
    FinishRootInternal eorFailed

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RootLeaseIsActive
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : True when the token matches the live, incomplete root.
'---------------------------------------------------------------------------------------
'
Public Function RootLeaseIsActive(strToken As String) As Boolean

    RootLeaseIsActive = (Not m_Root.Completed) _
        And Len(m_Root.Token) > 0 _
        And StrComp(m_Root.Token, strToken, vbBinaryCompare) = 0 _
        And (m_Root.State = eosRunning Or m_Root.State = eosStaged)

End Function


'---------------------------------------------------------------------------------------
' Procedure : Finish
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Synchronous root completion for work begun with Begin on the same stack.
'           : Finish supplies the live root token to itself, so it performs no ownership
'           : check: only call it from the procedure that began the root.
'---------------------------------------------------------------------------------------
'
Public Function Finish(Optional intResult As eOperationResult = eorUnknown)

    If Len(m_Root.Token) = 0 Then Exit Function
    CompleteRootLease m_Root.Token, intResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : TryPause
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Begin a short synchronous pause around foreign code: user VBA invoked
'           : through Application.Run, or closing a form whose unload handler would
'           : otherwise read a running operation as a cancel request. The root is
'           : suspended, not detached - no continuation token is issued, and the pause
'           : ends on the same stack that started it.
'---------------------------------------------------------------------------------------
'
Public Function TryPause() As clsOperationPause

    Dim cPause As clsOperationPause

    ' Pauses nest: an already-paused root is staged, so accept that too, but never a
    ' root that was detached for an asynchronous continuation.
    If m_Root.State <> eosRunning Then
        If m_Root.PauseDepth = 0 Then Exit Function
        If m_Root.State <> eosStaged Or m_Root.Detached Then Exit Function
    End If

    Set cPause = New clsOperationPause
    cPause.Init Me
    Set TryPause = cPause

End Function


'---------------------------------------------------------------------------------------
' Procedure : BeginPauseScope
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Internal entry for clsOperationPause.
'---------------------------------------------------------------------------------------
'
Public Sub BeginPauseScope()

    m_Root.PauseDepth = m_Root.PauseDepth + 1
    If m_Root.PauseDepth > 1 Then Exit Sub

    If m_Root.State = eosRunning Then
        m_Root.State = eosStaged
        m_Root.Detached = False
        RestoreUserErrorTrapping
        UpdateRegistry
        Pulse
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EndPauseScope
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Internal exit for clsOperationPause.
'---------------------------------------------------------------------------------------
'
Public Sub EndPauseScope()

    If m_Root.PauseDepth <= 0 Then Exit Sub
    m_Root.PauseDepth = m_Root.PauseDepth - 1
    If m_Root.PauseDepth > 0 Then Exit Sub

    ' A root detached for a continuation while paused is no longer ours to resume.
    If m_Root.State = eosStaged And Not m_Root.Detached Then
        m_Root.State = eosRunning
        SetErrorTrapping
        UpdateRegistry
        Pulse
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Pulse
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Pulse the heartbeat
'---------------------------------------------------------------------------------------
'
Public Function Pulse()
    Heartbeat = Now
End Function


'---------------------------------------------------------------------------------------
' Procedure : Status
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Return the current status of the operation.
'---------------------------------------------------------------------------------------
'
Public Property Get Status() As eOperationState
    If m_Root.State <> eosReady Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, Me.Heartbeat) < Now Then
            m_Root.State = eosReady
            Me.Result = eorTimeout
            RestoreUserErrorTrapping
        End If
    End If
    Status = m_Root.State
End Property


'---------------------------------------------------------------------------------------
' Procedure : OperationType
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : The type of work the active root is doing.
'---------------------------------------------------------------------------------------
'
Public Property Get OperationType() As eOperationType
    OperationType = m_Root.OperationType
End Property


Public Property Let OperationType(intType As eOperationType)
    m_Root.OperationType = intType
End Property


' eInteractionMode values are ordered by strictness (Normal < Silent).
Public Property Get InteractionMode() As eInteractionMode
    If m_Root.State = eosReady Or Len(m_Root.Token) = 0 Then
        InteractionMode = m_eimSticky
    Else
        InteractionMode = m_Root.RequestedInteractionMode
    End If
End Property


' Interaction can only get stricter while a root is active: a caller that asked for
' silence must not have it revoked by unrelated code deeper in the same operation.
' Between operations the value is just the default for the next root, so relaxing it is
' allowed - that is how a headless caller returns the session to normal.
Public Property Let InteractionMode(eimMode As eInteractionMode)

    Dim blnRootActive As Boolean
    blnRootActive = (m_Root.State <> eosReady) And (Len(m_Root.Token) > 0)

    If eimMode < m_eimSticky And blnRootActive Then
        LogDiagnostic "Ignored attempt to relax interaction mode", "clsOperation.InteractionMode"
        Exit Property
    End If

    m_eimSticky = eimMode
    If blnRootActive Then
        If eimMode > m_Root.RequestedInteractionMode Then
            m_Root.RequestedInteractionMode = eimMode
        End If
    End If

End Property


Public Property Get Attended() As Boolean
    Attended = m_Root.Attended
End Property


Public Property Get CurrentRootToken() As String
    CurrentRootToken = m_Root.Token
End Property


Public Property Get SourceName() As String
    Select Case Me.Source
        Case eosMCPTool:     SourceName = "MCP Tool"
        Case eosExternalAPI: SourceName = "External API"
    End Select
End Property


Public Property Get AutomationSource() As Boolean
    AutomationSource = (Me.Source = eosExternalAPI Or Me.Source = eosMCPTool)
End Property


Private Function CanBeginRoot() As Boolean

    Select Case m_Root.State
        Case eosRunning, eosStaged
            ReportRefusal T("Another Operation Already Running"), _
                T("Please complete the current operation before beginning a new one.")
        Case eosReady
            CanBeginRoot = True
        Case Else
            ReportRefusal T("Unknown Operation State"), T("Current state: {0}", , , , m_Root.State)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReportRefusal
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Explain a refused root. This runs before a root exists, so the usual
'           : Attended/InteractionMode state describes the previous operation and cannot
'           : be trusted - decide from the incoming caller's own intent instead. An
'           : automated or explicitly unattended caller gets a log entry, never a dialog.
'---------------------------------------------------------------------------------------
'
Private Sub ReportRefusal(strTitle As String, strMessage As String)

    If IsSessionSingleton _
        And m_eimSticky = eimNormal _
        And Not (Me.AutomationSource Or Me.ForceUnattended) Then
        MsgBox2 strTitle, strMessage, , vbExclamation
    Else
        LogDiagnostic strTitle & ": " & strMessage, "clsOperation.CanBeginRoot"
    End If

End Sub


Private Sub InitializeRoot(intType As eOperationType)

    m_Root.Token = NewRootToken()
    m_Root.OperationType = intType
    m_Root.RequestedInteractionMode = m_eimSticky
    m_Root.Attended = Not (Me.AutomationSource Or Me.ForceUnattended)
    m_Root.StartTime = Now
    m_Root.FinishTime = 0
    m_Root.State = eosRunning
    m_Root.Detached = False
    m_Root.Completed = False
    m_Root.PauseDepth = 0
    m_Root.AppHwnd = hWndAccessApp

    Me.ErrorLevel = eelNoError
    Me.Result = eorUnknown

    SetErrorTrapping
    UpdateRegistry
    Pulse

End Sub


Private Sub FinishRootInternal(Optional intResult As eOperationResult = eorUnknown)

    Dim blnWasRunning As Boolean

    If m_Root.Completed Then Exit Sub
    If m_Root.State = eosReady And Len(m_Root.Token) = 0 Then Exit Sub

    If intResult > eorUnknown Then Me.Result = intResult
    blnWasRunning = (m_Root.State = eosRunning)

    m_Root.Completed = True
    m_Root.State = eosReady
    m_Root.FinishTime = Now

    If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagBegin "op.registry"
    UpdateRegistry
    If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagEnd "op.registry"

    If MCP.IsActive And IsSessionSingleton Then
        Dim dExtra As Dictionary
        Set dExtra = New Dictionary
        dExtra.Add "log_path", Log.SavedLogFilePath
        ' Failed tests complete as eorFailed (MCP "error"), so the results file
        ' has to ride on every terminal callback or the client only sees
        ' "Operation failed" with no JSON to rerun from.
        If m_Root.OperationType = eotTestRun Then
            If Len(TestRunner.LastResultsPath) > 0 Then
                dExtra.Add "results_path", TestRunner.LastResultsPath
            End If
        End If

        Select Case Me.Result
            Case eorSuccess
                MCP.PostCallback "complete", 100, 100, "Operation completed successfully", dExtra
            Case eorFailed
                MCP.PostCallback "error", -1, -1, "Operation failed", dExtra
            Case eorCanceled
                MCP.PostCallback "cancelled", -1, -1, "Operation was cancelled", dExtra
            Case eorTimeout
                MCP.PostCallback "error", -1, -1, "Operation timed out", dExtra
            Case Else
                MCP.PostCallback "complete", 100, 100, "Operation completed", dExtra
        End Select
    End If

    If IsSessionSingleton Then
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagBegin "op.release"
        modObjects.ReleaseObjects
        If modTestRunnerDiag.DiagEnabled Then modTestRunnerDiag.DiagEnd "op.release"
    End If

    If blnWasRunning Then RestoreUserErrorTrapping

    Me.ForceUnattended = False
    m_eimSticky = eimNormal

    m_Root.Token = vbNullString

End Sub


Private Function NewRootToken() As String
    NewRootToken = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Timer * 1000000@)
End Function


Private Sub LogDiagnostic(strMessage As String, strSource As String)
    On Error Resume Next
    Log.Add "[Operation] " & strMessage & IIf(Len(strSource) > 0, " (" & strSource & ")", vbNullString), False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsSessionSingleton
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : True only for the one instance published by modObjects. Everything that
'           : reaches outside this instance - the registry state used for crash recovery,
'           : the MCP completion callback, VBE error trapping, and singleton teardown -
'           : is gated on this so a test can drive a private instance through a full root
'           : lifecycle without disturbing the operation that is actually running.
'---------------------------------------------------------------------------------------
'
Private Function IsSessionSingleton() As Boolean
    On Error Resume Next
    IsSessionSingleton = (Me Is modObjects.Operation)
End Function


Private Sub UpdateRegistry()
    If Not IsSessionSingleton Then Exit Sub
    If m_Root.AppHwnd <> 0 Then
        CheckSetting "Heartbeat", Me.Heartbeat
        CheckSetting "Result", Me.Result
        CheckSetting "StartTime", m_Root.StartTime
        CheckSetting "FinishTime", m_Root.FinishTime
        CheckSetting "State", m_Root.State
        CheckSetting "AppHwnd", m_Root.AppHwnd
        CheckSetting "ErrorTrapping", m_Root.ErrorTrapping
        CheckSetting REG_KEY_ROOT_TOKEN, m_Root.Token
        CheckSetting REG_KEY_ROOT_TYPE, m_Root.OperationType
    End If
End Sub


Private Sub CheckSetting(strKey As String, varValue)
    If GetSetting(PROJECT_NAME, "State", strKey) <> CStr(varValue) Then
        SaveSetting PROJECT_NAME, "State", strKey, CStr(varValue)
    End If
End Sub


Private Function RestoreFromRegistry() As Boolean

    If ReadReg("AppHwnd") = CStr(hWndAccessApp) Then
        If DateAdd("s", HEARTBEAT_TIMEOUT_SECONDS, CDate(Nz2(ReadReg("Heartbeat"), 0))) > Now Then
            Me.Result = ReadReg("Result")
            m_Root.StartTime = ReadReg("StartTime")
            m_Root.FinishTime = ReadReg("FinishTime")
            m_Root.State = ReadReg("State")
            m_Root.AppHwnd = ReadReg("AppHwnd")
            m_Root.ErrorTrapping = CInt(ReadReg("ErrorTrapping"))
            m_Root.Token = ReadReg(REG_KEY_ROOT_TOKEN)
            m_Root.OperationType = Val(ReadReg(REG_KEY_ROOT_TYPE))
            m_Root.Completed = False
            RestoreFromRegistry = True
        End If
    End If

End Function


Private Function ReadReg(strKey As String) As String
    ReadReg = GetSetting(PROJECT_NAME, "State", strKey)
End Function


Private Sub SetErrorTrapping()
    If Not IsSessionSingleton Then Exit Sub
    m_Root.ErrorTrapping = SaveUserErrorTrapping()
    ApplyUserErrorTrapping eetBreakInClassModule
End Sub


Private Sub RestoreUserErrorTrapping()
    If Not IsSessionSingleton Then Exit Sub
    modErrorHandling.RestoreUserErrorTrapping m_Root.ErrorTrapping
End Sub


Private Sub Class_Initialize()
    m_Root.State = eosReady
    m_eimSticky = eimNormal
    m_Root.RequestedInteractionMode = eimNormal
    m_Root.Attended = True
    If modObjects.CreatingOperationSingleton Then
        If RestoreFromRegistry Then Exit Sub
    End If
    m_Root.AppHwnd = Application.hWndAccessApp
End Sub
