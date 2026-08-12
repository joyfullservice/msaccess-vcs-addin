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
' Purpose   : Tests for unattended add-in rebuild helpers: process guard, trusted
'           : location check, /cmd token parse, and rebuild-status.json round trip.
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
End Sub


Public Sub TestGetOtherAccessInstances_DoesNotRaise()
    Dim strDetail As String
    Dim lngCount As Long
    lngCount = GetOtherAccessInstances(strDetail)
    TestAssert lngCount >= -1, "returns -1 or a count"
    If lngCount = 0 Then
        TestAssert Len(strDetail) = 0, "no detail when count is zero"
    ElseIf lngCount > 0 Then
        TestAssert Len(strDetail) > 0, "detail lists other instances"
        TestAssert InStr(strDetail, "PID ") > 0, "detail includes PID"
    Else
        TestAssert Len(strDetail) > 0, "query failure includes a reason"
    End If
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
