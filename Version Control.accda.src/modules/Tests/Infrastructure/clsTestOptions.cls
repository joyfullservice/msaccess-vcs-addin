VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestOptions
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsOptions defaults, hashing, and enum conversions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestLoadDefaults_HashStable()
    Dim cOpt As New clsOptions
    cOpt.LoadDefaults
    Dim strHash1 As String
    strHash1 = cOpt.GetHash
    TestAssert Len(strHash1) > 0, "hash is non-empty"

    Dim cOpt2 As New clsOptions
    cOpt2.LoadDefaults
    TestAssert cOpt2.GetHash = strHash1, "same defaults produce same hash"
End Sub


Public Sub TestGetCategoryHashes()
    Dim cOpt As New clsOptions
    cOpt.LoadDefaults
    Dim dHashes As Dictionary
    Set dHashes = cOpt.GetCategoryHashes
    TestAssert Not dHashes Is Nothing, "returns dictionary"
    TestAssert dHashes.Count > 0, "non-empty category hashes"
End Sub


Public Sub TestSkipReopenBeforeMergeIsNotAnExportOption()
    Dim cOpt As New clsOptions
    cOpt.LoadDefaults
    TestAssert cOpt.SkipReopenBeforeMerge = False, "defaults to off"

    ' This option only affects how a merge prepares the database, so flipping it must
    ' never change an export hash (which would re-export every affected category).
    Dim strHash As String
    strHash = cOpt.GetHash
    cOpt.SkipReopenBeforeMerge = True
    TestAssert cOpt.GetHash = strHash, "toggling does not change the export hash"
End Sub


Public Sub TestSanitizeLevelNames()
    Dim cOpt As New clsOptions
    TestAssert cOpt.GetSanitizeLevelName(eslNone) = "None (Off)", "None level"
    TestAssert cOpt.GetSanitizeLevelName(eslStandard) = "Standard", "Standard level"
    TestAssert cOpt.GetSanitizeLevel("None (Off)") = eslNone, "None from name"
    TestAssert cOpt.GetSanitizeLevel("Standard") = eslStandard, "Standard from name"
End Sub


Public Sub TestTableExportFormatNames()
    Dim cOpt As New clsOptions
    Dim strName As String
    strName = cOpt.GetTableExportFormatName(etdTabDelimited)
    TestAssert Len(strName) > 0, "format name non-empty"
    TestAssert cOpt.GetTableExportFormat(strName) = etdTabDelimited, "round trip format name"
End Sub


Public Sub TestGetExportFolder()
    TestAssert Len(Options.GetExportFolder) > 0, "export folder non-empty"
    TestAssert Right$(Options.GetExportFolder, 1) = "\", "ends with separator"
End Sub


Public Sub TestEndSession_DiscardsCachedOptions()

    Dim cVcs As clsVersionControl
    Dim cFresh As clsOptions
    Dim dResult As Dictionary
    Dim strSession As String
    Dim blnPriorShowDebug As Boolean

    ' A dummy session id so this does not delete the live MCP override file or
    ' clear modObjects.SessionId. EndSession still discards the Options cache.
    strSession = "end-session-shadow-test"
    blnPriorShowDebug = Options.ShowDebug
    Options.ShowDebug = Not blnPriorShowDebug

    Set cVcs = VCS
    Set dResult = ParseJson(cVcs.EndSession(strSession))

    TestAssert dResult("success") = True, "EndSession succeeds (Options is settable)"

    Set cFresh = New clsOptions
    cFresh.LoadProjectOptions
    TestAssert Options.ShowDebug = cFresh.ShowDebug, _
        "cached Options was discarded and reloaded from the project file"

    If Len(modObjects.SessionId) > 0 Then Options.LoadOptionOverrides
    Options.ShowDebug = blnPriorShowDebug

End Sub
