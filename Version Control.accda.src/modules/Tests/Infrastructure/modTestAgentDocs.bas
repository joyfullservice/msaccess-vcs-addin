Attribute VB_Name = "modTestAgentDocs"
'---------------------------------------------------------------------------------------
' Module    : modTestAgentDocs
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Guard the agent documentation that ships into every user's export folder.
'           : AGENTS.md and vcs-agent-docs\*.md are extracted on every export and sit in
'           : the reading agent's context on every turn, so they are budgeted. Every
'           : failure mode these tests cover is otherwise silent: a reference nobody
'           : links to is simply never read, a repo-relative path resolves to nothing
'           : once the file leaves this repo, and a document missing from
'           : GetAgentDocFiles is skipped by VerifyResource without a word.
'           :
'           : See docs\agent-docs-maintenance.md for the rules these enforce.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")

' Line budgets. See docs\agent-docs-maintenance.md for where these come from.
Private Const MAX_ENTRY_LINES As Long = 150
Private Const MAX_REFERENCE_LINES As Long = 110

Private Const SOURCE_FOLDER As String = "Version Control.accda.src"
Private Const DOCS_FOLDER As String = "vcs-agent-docs"
Private Const ENTRY_FILE As String = "AGENTS.md"

' Identifiers that only exist inside this add-in's own project or repository, and so
' mean nothing to an agent reading these files in somebody else's database.
Private Const BANNED_IDENTIFIERS As String = "VCSIndex|Testing/Fixtures|CodeDb|CodeProject"

' modTestAssert is installed into the user's project, so it is a legitimate mention
' even though it matches the modTest* pattern used for this repo's own test modules.
Private Const ALLOWED_TEST_MODULE As String = "modTestAssert"


'---------------------------------------------------------------------------------------
' Procedure : TestAgentDocsWithinLineBudget
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : The entry file and each reference stay within budget. An addition that
'           : breaks this must remove something in the same edit.
'---------------------------------------------------------------------------------------
'
Public Sub TestAgentDocsWithinLineBudget()

    Dim varFile As Variant
    Dim lngLines As Long

    If Not SourceIsAvailable Then Exit Sub

    lngLines = CountLines(EntryFilePath)
    TestAssert lngLines <= MAX_ENTRY_LINES, _
        ENTRY_FILE & " is " & lngLines & " lines, over the " & _
        MAX_ENTRY_LINES & " line budget"

    For Each varFile In modResource.GetAgentDocFiles
        lngLines = CountLines(DocsFolderPath & varFile)
        TestAssert lngLines <= MAX_REFERENCE_LINES, _
            varFile & " is " & lngLines & " lines, over the " & _
            MAX_REFERENCE_LINES & " line budget"
    Next varFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAgentDocsAreLinkedFromEntryFile
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : A reference the entry file never names is read in a small fraction of
'           : sessions. Linking it is what makes the content reachable at all.
'---------------------------------------------------------------------------------------
'
Public Sub TestAgentDocsAreLinkedFromEntryFile()

    Dim strEntry As String
    Dim varFile As Variant

    If Not SourceIsAvailable Then Exit Sub

    strEntry = ReadFile(EntryFilePath)

    For Each varFile In modResource.GetAgentDocFiles
        TestAssert InStr(1, strEntry, DOCS_FOLDER & "/" & varFile, vbTextCompare) > 0, _
            varFile & " is not referenced from " & ENTRY_FILE
    Next varFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAgentDocsHaveNoRepoRelativeLinks
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : A "../" link points outside the export folder once these files ship, so
'           : it resolves to nothing in the user's repository.
'---------------------------------------------------------------------------------------
'
Public Sub TestAgentDocsHaveNoRepoRelativeLinks()

    Dim varFile As Variant

    If Not SourceIsAvailable Then Exit Sub

    AssertNoRepoRelativeLinks ENTRY_FILE, ReadFile(EntryFilePath)

    For Each varFile In modResource.GetAgentDocFiles
        AssertNoRepoRelativeLinks CStr(varFile), ReadFile(DocsFolderPath & varFile)
    Next varFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAgentDocsHaveNoInternalIdentifiers
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Names that exist only in this project or repository are unreachable and
'           : meaningless from a user's database.
'---------------------------------------------------------------------------------------
'
Public Sub TestAgentDocsHaveNoInternalIdentifiers()

    Dim varFile As Variant

    If Not SourceIsAvailable Then Exit Sub

    AssertNoInternalIdentifiers ENTRY_FILE, ReadFile(EntryFilePath)

    For Each varFile In modResource.GetAgentDocFiles
        AssertNoInternalIdentifiers CStr(varFile), ReadFile(DocsFolderPath & varFile)
    Next varFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestAgentDocsAreRegisteredAsResources
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : GetAgentDocFiles drives both resource registration and extraction, and
'           : VerifyResource skips a missing source file silently. A document present
'           : on disk but absent from the list would never reach a user.
'---------------------------------------------------------------------------------------
'
Public Sub TestAgentDocsAreRegisteredAsResources()

    Dim dRegistered As Dictionary
    Dim varFile As Variant
    Dim oFile As Scripting.File

    If Not SourceIsAvailable Then Exit Sub

    Set dRegistered = New Dictionary
    dRegistered.CompareMode = TextCompare

    ' Every registered document exists on disk, or the resource is skipped silently
    For Each varFile In modResource.GetAgentDocFiles
        dRegistered(CStr(varFile)) = True
        TestAssert FSO.FileExists(DocsFolderPath & varFile), _
            varFile & " is registered but missing from " & DOCS_FOLDER
    Next varFile

    ' Every document on disk is registered, or it never ships
    For Each oFile In FSO.GetFolder(StripSlash(DocsFolderPath)).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "md", vbTextCompare) = 0 Then
            TestAssert dRegistered.Exists(oFile.Name), _
                oFile.Name & " is missing from modResource.GetAgentDocFiles"
        End If
    Next oFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertNoRepoRelativeLinks
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Fail on any markdown link whose target climbs out of the export folder.
'---------------------------------------------------------------------------------------
'
Private Sub AssertNoRepoRelativeLinks(strFile As String, strContent As String)
    TestAssert InStr(1, strContent, "](../") = 0, _
        strFile & " contains a repo-relative link that will not resolve for a user"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertNoInternalIdentifiers
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Fail on add-in internals and on this repo's own test module names. The
'           : modTest* check strips modTestAssert first, since that module is installed
'           : into the user's project and is fair game.
'---------------------------------------------------------------------------------------
'
Private Sub AssertNoInternalIdentifiers(strFile As String, strContent As String)

    Dim varTerm As Variant
    Dim strStripped As String
    Dim strNext As String
    Dim lngChar As Long

    For Each varTerm In Split(BANNED_IDENTIFIERS, "|")
        TestAssert InStr(1, strContent, CStr(varTerm), vbTextCompare) = 0, _
            strFile & " mentions " & varTerm & ", which does not exist in a user's database"
    Next varTerm

    ' "modTest" followed by an uppercase letter names one of this repo's test modules.
    ' Compare by character code rather than Like, which Option Compare Database would
    ' make case-insensitive and therefore useless here.
    strStripped = Replace(strContent, ALLOWED_TEST_MODULE, vbNullString, , , vbTextCompare)
    lngChar = InStr(1, strStripped, "modTest", vbTextCompare)
    Do While lngChar > 0
        strNext = Mid$(strStripped, lngChar + 7, 1)
        TestAssert Len(strNext) = 0 Or Asc(strNext & " ") < 65 Or Asc(strNext & " ") > 90, _
            strFile & " names a test module from this repository: " & _
            Mid$(strStripped, lngChar, 24)
        lngChar = InStr(lngChar + 1, strStripped, "modTest", vbTextCompare)
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SourceIsAvailable
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : These tests read the add-in's own source tree, which is only present when
'           : running from a repository checkout. Report a passing note rather than a
'           : failure when the add-in is running from its installed location, matching
'           : how VerifyResource treats the same condition.
'---------------------------------------------------------------------------------------
'
Private Function SourceIsAvailable() As Boolean
    SourceIsAvailable = FSO.FileExists(EntryFilePath)
    If Not SourceIsAvailable Then
        TestAssert True, "skipped: add-in source not available at " & EntryFilePath
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : EntryFilePath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Full path to the shipped entry file in the add-in's source tree.
'---------------------------------------------------------------------------------------
'
Private Function EntryFilePath() As String
    EntryFilePath = SourceFolderPath & ENTRY_FILE
End Function


'---------------------------------------------------------------------------------------
' Procedure : DocsFolderPath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Full path to the reference document folder, with a trailing separator.
'---------------------------------------------------------------------------------------
'
Private Function DocsFolderPath() As String
    DocsFolderPath = SourceFolderPath & DOCS_FOLDER & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceFolderPath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : The add-in's own export folder inside the repository. CodeProject.Path is
'           : the repo root when running from "<repo>\Version Control.accda", which is
'           : the same assumption VerifyResource makes.
'---------------------------------------------------------------------------------------
'
Private Function SourceFolderPath() As String
    SourceFolderPath = CodeProject.Path & PathSep & SOURCE_FOLDER & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : CountLines
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Count lines in a text file, ignoring a single trailing newline.
'---------------------------------------------------------------------------------------
'
Private Function CountLines(strPath As String) As Long

    Dim strContent As String

    strContent = ReadFile(strPath)
    If Len(strContent) = 0 Then Exit Function

    ' Drop one trailing newline so a file ending in CRLF is not counted as having an
    ' extra empty line at the end.
    If Right$(strContent, 2) = vbCrLf Then
        strContent = Left$(strContent, Len(strContent) - 2)
    ElseIf Right$(strContent, 1) = vbLf Then
        strContent = Left$(strContent, Len(strContent) - 1)
    End If

    CountLines = UBound(Split(Replace(strContent, vbCrLf, vbLf), vbLf)) + 1

End Function
