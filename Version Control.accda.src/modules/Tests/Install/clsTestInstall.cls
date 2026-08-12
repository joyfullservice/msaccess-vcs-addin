VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestInstall
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Tests for unattended add-in rebuild helpers: process guard, held-file
'           : matching, build target derivation, trusted location check, /cmd token
'           : parse, command-line database detection, rebuild-status.json round trip,
'           : and the worker handoff check. Nothing here may close a process: the
'           : guard only reports, and a test that terminated an Access instance would
'           : take the developer's other work with it.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Install")
'@Tag("unit")

Private m_strTempFolder As String


Private Sub Class_Initialize()
    m_strTempFolder = GetTempFolder("vcs_rebuild_status")
End Sub


Private Sub Class_Terminate()
    If Len(m_strTempFolder) > 0 Then
        If FSO.FolderExists(m_strTempFolder) Then FSO.DeleteFolder StripSlash(m_strTempFolder), True
    End If
End Sub


Public Sub TestParseInstallCommand_Install()
    TestAssert ParseInstallCommand("INSTALL") = "INSTALL", "INSTALL"
    TestAssert ParseInstallCommand("install") = "INSTALL", "case insensitive"
    TestAssert ParseInstallCommand("  INSTALL  ") = "INSTALL", "trimmed"
End Sub


Public Sub TestParseInstallCommand_InstallSilent()
    TestAssert ParseInstallCommand("INSTALL SILENT") = "INSTALL SILENT", "INSTALL SILENT"
    TestAssert ParseInstallCommand("install silent") = "INSTALL SILENT", "case insensitive"
    TestAssert ParseInstallCommand("INSTALL  SILENT") = "INSTALL SILENT", "collapsed spaces"
End Sub


Public Sub TestParseInstallCommand_NotInstall()
    TestAssert ParseInstallCommand(vbNullString) = vbNullString, "empty"
    TestAssert ParseInstallCommand("EXPORT") = vbNullString, "EXPORT"
    TestAssert ParseInstallCommand("INSTALLING") = vbNullString, "INSTALLING is not INSTALL"
    TestAssert ParseInstallCommand("INSTALL LOUDLY") = vbNullString, "unknown modifier"
End Sub


Public Sub TestParseInstallCommand_StatusFile()

    Dim strFile As String

    TestAssert ParseInstallCommand("INSTALL SILENT C:\ci\status.json", strFile) = _
        "INSTALL SILENT", "keyword still recognized with a path"
    TestAssert strFile = "C:\ci\status.json", "unquoted path"

    ParseInstallCommand "INSTALL SILENT ""C:\Program Files\status.json""", strFile
    TestAssert strFile = "C:\Program Files\status.json", "quotes stripped"

    ' Taken whole rather than split on spaces, so repeated spaces in a path survive.
    ParseInstallCommand "INSTALL SILENT C:\two  spaces\status.json", strFile
    TestAssert strFile = "C:\two  spaces\status.json", "interior spaces preserved"

    ParseInstallCommand "INSTALL SILENT", strFile
    TestAssert strFile = vbNullString, "no path supplied"

    ParseInstallCommand "INSTALL", strFile
    TestAssert strFile = vbNullString, "plain install carries no path"

End Sub


Public Sub TestGetOtherAccessInstances_DoesNotRaise()

    Dim strDetail As String
    Dim lngCount As Long
    Dim lngBlocking As Long

    lngCount = GetOtherAccessInstances(strDetail, lngBlocking, _
        GetInstalledAddInFileName())
    TestAssert lngCount >= -1, "returns -1 or a count"
    If lngCount = 0 Then
        TestAssert Len(strDetail) = 0, "no detail when count is zero"
        TestAssert lngBlocking = 0, "nothing blocks when no other instance exists"
    ElseIf lngCount > 0 Then
        TestAssert Len(strDetail) > 0, "detail lists other instances"
        TestAssert InStr(strDetail, "PID ") > 0, "detail includes PID"
        TestAssert lngBlocking >= 0 And lngBlocking <= lngCount, _
            "blocking count is a subset of the instances found"
        ' An instance only blocks a rebuild if it holds a file the rebuild replaces, so
        ' log what real instances reported rather than assuming how they classify.
        Log.Add "Other Access instances seen by the rebuild guard:" & vbCrLf & strDetail, False
    Else
        TestAssert Len(strDetail) > 0, "query failure includes a reason"
        TestAssert lngBlocking = -1, "query failure blocks the rebuild"
    End If

End Sub


Public Sub TestIsTargetFilePath()

    Const TARGETS As String = "C:\Users\Me\AppData\Roaming\MSAccessVCS\Version Control.accda" & _
        vbCrLf & "C:\Repos\addin\Version Control.accda"

    ' The VBE reports a project's FileName in upper case, which is the whole reason this
    ' comparison cannot be an equality test.
    TestAssert IsTargetFilePath( _
        "C:\USERS\ME\APPDATA\ROAMING\MSACCESSVCS\VERSION CONTROL.ACCDA", TARGETS), _
        "upper case FileName matches"
    TestAssert IsTargetFilePath("C:\Repos\addin\Version Control.accda", TARGETS), _
        "second entry in the list"
    TestAssert Not IsTargetFilePath("C:\Repos\addin\Other.accdb", TARGETS), "unrelated file"
    TestAssert Not IsTargetFilePath(vbNullString, TARGETS), "empty file name"
    TestAssert Not IsTargetFilePath("C:\a.accda", vbNullString), "empty target list"
    TestAssert Not IsTargetFilePath("C:\a.accda", vbCrLf & vbCrLf), "blank entries are skipped"

End Sub


Public Sub TestGetBuildTargetFileName()

    TestAssert GetBuildTargetFileName("C:\Repos\addin\Version Control.accda.src") = _
        "C:\Repos\addin\Version Control.accda", "source folder names its database"
    TestAssert GetBuildTargetFileName("C:\Repos\addin\Version Control.accda.src\") = _
        "C:\Repos\addin\Version Control.accda", "trailing separator"
    TestAssert GetBuildTargetFileName("C:\Repos\addin\NotAnExport") = vbNullString, _
        "folder without .src infers nothing"
    TestAssert GetBuildTargetFileName(vbNullString) = vbNullString, "empty"

End Sub


Public Sub TestExtractOpenDatabaseFromCommandLine_QuotedPath()
    TestAssert ExtractOpenDatabaseFromCommandLine( _
        """C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"" ""C:\My Data\db.accdb""") _
        = "C:\My Data\db.accdb", "quoted path with spaces"
End Sub


Public Sub TestExtractOpenDatabaseFromCommandLine_UnquotedPath()
    TestAssert ExtractOpenDatabaseFromCommandLine("MSACCESS.EXE C:\dev\db.accdb") = _
        "C:\dev\db.accdb", "unquoted path"
    TestAssert ExtractOpenDatabaseFromCommandLine( _
        """C:\Office16\MSACCESS.EXE"" C:\dev\db.accdb /nostartup") = _
        "C:\dev\db.accdb", "unquoted path followed by a switch"
End Sub


Public Sub TestExtractOpenDatabaseFromCommandLine_NoDatabase()
    TestAssert ExtractOpenDatabaseFromCommandLine(vbNullString) = vbNullString, "empty"
    TestAssert ExtractOpenDatabaseFromCommandLine("""C:\Office16\MSACCESS.EXE""") = _
        vbNullString, "executable only"
    TestAssert ExtractOpenDatabaseFromCommandLine( _
        """C:\Office16\MSACCESS.EXE"" /cmd ""INSTALL SILENT""") = vbNullString, _
        "switch arguments are not databases"
    TestAssert ExtractOpenDatabaseFromCommandLine("MSACCESS.EXE C:\dev\notes.txt") = _
        vbNullString, "non-database extension"
End Sub


Public Sub TestExtractOpenDatabaseFromCommandLine_Extensions()
    TestAssert ExtractOpenDatabaseFromCommandLine("MSACCESS.EXE C:\a.accda") = "C:\a.accda", "accda"
    TestAssert ExtractOpenDatabaseFromCommandLine("MSACCESS.EXE C:\a.MDB") = "C:\a.MDB", "mdb, any case"
    TestAssert ExtractOpenDatabaseFromCommandLine("MSACCESS.EXE C:\a.adp") = "C:\a.adp", "adp"
End Sub


Public Sub TestRebuildStatusFile_RoundTrip()
    Dim strFile As String
    Dim dStatus As Dictionary
    Dim strStarted As String

    strFile = GetRebuildStatusFilePath(m_strTempFolder)
    TestAssert Right$(strFile, Len("logs" & PathSep & "rebuild-status.json")) = _
        "logs" & PathSep & "rebuild-status.json", "status file lives under logs/"

    WriteRebuildStatusFile strFile, "launched", vbNullString, vbNullString, "2026-08-12 10:00:00"
    Set dStatus = ReadRebuildStatusFile(strFile)
    TestAssert Not dStatus Is Nothing, "reads back a dictionary"
    TestAssert dStatus("status") = "launched", "status"
    TestAssert dStatus("phaseStarted") = "2026-08-12 10:00:00", "phaseStarted preserved"
    TestAssert dStatus.Exists("updated"), "updated written"
    TestAssert dStatus("error") = vbNullString, "empty error"
    TestAssert dStatus("buildLog") = vbNullString, "empty buildLog"
    strStarted = dStatus("phaseStarted")

    WriteRebuildStatusFile strFile, "building", "in progress", "C:\logs\Build_1.log"
    Set dStatus = ReadRebuildStatusFile(strFile)
    TestAssert dStatus("status") = "building", "updated status"
    TestAssert dStatus("error") = "in progress", "error field"
    TestAssert dStatus("buildLog") = "C:\logs\Build_1.log", "buildLog field"
    TestAssert dStatus("phaseStarted") = strStarted, "phaseStarted kept on update"

    ' Inheriting the previous value is right for a phase update and wrong for the first
    ' write of a new attempt, which is why RebuildAddIn stamps "launched" explicitly.
    ' Without the override a fresh attempt would carry the last rebuild's timestamp and
    ' read as long stalled to anyone timing it against the clock.
    WriteRebuildStatusFile strFile, "launched", vbNullString, vbNullString, "2026-08-12 11:30:00"
    Set dStatus = ReadRebuildStatusFile(strFile)
    TestAssert dStatus("phaseStarted") = "2026-08-12 11:30:00", _
        "an explicit phaseStarted replaces a stale one"
End Sub


Public Sub TestRebuildStatusFile_LaunchFailedRoundTrip()
    Dim strFile As String
    Dim dStatus As Dictionary
    Dim strError As String

    strError = "The rebuild worker did not start within 15 seconds."
    strFile = GetRebuildStatusFilePath(m_strTempFolder)

    WriteRebuildStatusFile strFile, "launch-failed", strError
    Set dStatus = ReadRebuildStatusFile(strFile)
    TestAssert dStatus("status") = "launch-failed", "terminal launch-failed status"
    TestAssert dStatus("error") = strError, "error text survives the round trip"
End Sub


Public Sub TestWaitForRebuildHandoff_NoFile()
    ' A worker that never wrote anything cannot be confirmed, and neither can one whose
    ' status file was removed between the launch and the check.
    TestAssert Not WaitForRebuildHandoff(StripSlash(m_strTempFolder) & PathSep & _
        "no-such-status.json", 0.4), "absent status file is not a handoff"
End Sub


Public Sub TestWaitForRebuildHandoff_StillLaunched()
    Dim strFile As String

    ' This is the failure the wait exists to catch: the launching code wrote "launched"
    ' and the worker died before writing anything of its own.
    strFile = GetRebuildStatusFilePath(m_strTempFolder)
    WriteRebuildStatusFile strFile, "launched"
    TestAssert Not WaitForRebuildHandoff(strFile, 0.4), "still launched is not a handoff"
End Sub


Public Sub TestWaitForRebuildHandoff_WorkerResponded()
    Dim strFile As String

    strFile = GetRebuildStatusFilePath(m_strTempFolder)
    WriteRebuildStatusFile strFile, "building"
    TestAssert WaitForRebuildHandoff(strFile, 0.4), "building confirms the handoff"

    ' A worker that started and immediately failed still started, and its own status is
    ' more useful to the caller than a generic launch failure.
    WriteRebuildStatusFile strFile, "build-failed", "source folder was empty"
    TestAssert WaitForRebuildHandoff(strFile, 0.4), "a terminal worker status also confirms it"
End Sub


Public Sub TestIsFolderTrusted_InstallFolder()
    If HasTrustedLocationKey Then
        TestAssert IsFolderTrusted(FSO.GetParentFolderName(GetAddInFileName)), _
            "install folder is trusted when the add-in registered a trusted location"
    Else
        TestAssert Not IsFolderTrusted("C:\Windows\Temp\vcs-not-a-trusted-location"), _
            "an arbitrary folder is not trusted"
    End If
End Sub
