VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestDotEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestDotEnv
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsDotEnv .env file handling.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Connect")

Private m_strTempFile As String


Private Sub Class_Initialize()
    m_strTempFile = GetTempFile() & ".env"
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    If FSO.FileExists(m_strTempFile) Then DeleteFile m_strTempFile
    On Error GoTo 0
End Sub


Public Sub TestSetVarAndGetVar()
    Dim cEnv As New clsDotEnv
    cEnv.SetVar "MY_KEY", "my_value"
    TestAssert cEnv.GetVar("MY_KEY", False) = "my_value", "set and get"
    TestAssert cEnv.GetVar("MY_KEY", False) = cEnv.GetVar("my_key", False), _
        "case insensitive keys"
End Sub


Public Sub TestSaveAndLoad()
    Dim cWrite As New clsDotEnv
    cWrite.SetVar "DB_HOST", "localhost"
    cWrite.SetVar "DB_PORT", "5432"
    cWrite.SaveToFile m_strTempFile

    Dim cRead As New clsDotEnv
    cRead.LoadFromFile m_strTempFile
    TestAssert cRead.GetVar("DB_HOST", False) = "localhost", "DB_HOST preserved"
    TestAssert cRead.GetVar("DB_PORT", False) = "5432", "DB_PORT preserved"
End Sub


Public Sub TestCommentPreservation()
    Dim strContent As String
    strContent = "# This is a comment" & vbCrLf & "KEY=value"
    WriteFile strContent, m_strTempFile

    Dim cEnv As New clsDotEnv
    cEnv.LoadFromFile m_strTempFile
    TestAssert cEnv.GetVar("KEY", False) = "value", "key after comment"
    ' WriteFile appends trailing vbCrLf; Split produces an extra blank entry
    TestAssert cEnv.Lines.Count = 3, "comment + key + trailing blank"
End Sub


Public Sub TestMissingKeyReturnsEmpty()
    Dim cEnv As New clsDotEnv
    TestAssert cEnv.GetVar("NONEXISTENT", False) = "", "missing key returns empty"
End Sub


Public Sub TestMergeIntoDictionary()
    Dim cEnv As New clsDotEnv
    cEnv.SetVar "A", "1"
    cEnv.SetVar "B", "2"

    Dim dTarget As New Dictionary
    dTarget.Add "B", "old"
    dTarget.Add "C", "3"
    cEnv.MergeIntoDictionary dTarget, False

    TestAssert dTarget("A") = "1", "new key added"
    TestAssert dTarget("B") = "2", "existing key replaced"
    TestAssert dTarget("C") = "3", "untouched key preserved"
End Sub
