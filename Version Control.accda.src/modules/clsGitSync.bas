Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_State As clsGitState


Public Property Get State() As clsGitState
    If m_State Is Nothing Then Set m_State = New clsGitState
    Set State = m_State
End Property