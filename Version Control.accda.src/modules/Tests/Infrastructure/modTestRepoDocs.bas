Attribute VB_Name = "modTestRepoDocs"
'---------------------------------------------------------------------------------------
' Module    : modTestRepoDocs
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Guard this repository's own agent-facing documentation. The root AGENTS.md
'           : is loaded on every turn of every session here, so it is budgeted the same
'           : way the shipped entry file is, and the Cursor rules are budgeted because a
'           : glob-scoped rule is cheap only while it stays short.
'           :
'           : The two structural checks cover silent failures: a broken docs/ link sends
'           : an agent looking for a file that is not there, and a reference document no
'           : file links to is simply never opened.
'           :
'           : Sibling of modTestAgentDocs, which guards the docs that ship to users.
'           : See docs\agent-docs-maintenance.md (Part 2) for the rules these enforce.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")

' Line budgets. See docs\agent-docs-maintenance.md for where these come from.
Private Const MAX_ROOT_LINES As Long = 150
Private Const MAX_RULE_LINES As Long = 120

Private Const ROOT_ENTRY_FILE As String = "AGENTS.md"
Private Const DOCS_FOLDER As String = "docs"
Private Const DOCS_INDEX_FILE As String = "README.md"
Private Const RULES_FOLDER As String = ".cursor\rules"

' Marker that opens a markdown link target pointing into the docs folder
Private Const DOCS_LINK_PREFIX As String = "](docs/"


'---------------------------------------------------------------------------------------
' Procedure : TestRootAgentsFileWithinLineBudget
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : The root AGENTS.md is a router into docs\, not a manual. An addition that
'           : breaks this budget must remove something or move content into docs\.
'---------------------------------------------------------------------------------------
'
Public Sub TestRootAgentsFileWithinLineBudget()

    Dim lngLines As Long

    If Not RepoIsAvailable Then Exit Sub

    lngLines = CountLines(RepoRootPath & ROOT_ENTRY_FILE)
    TestAssert lngLines <= MAX_ROOT_LINES, _
        ROOT_ENTRY_FILE & " is " & lngLines & " lines, over the " & _
        MAX_ROOT_LINES & " line budget"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCursorRulesWithinLineBudget
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Rules are triggers, not references. Depth belongs in docs\, which the rule
'           : links to, so that agents outside Cursor can still reach it.
'---------------------------------------------------------------------------------------
'
Public Sub TestCursorRulesWithinLineBudget()

    Dim oFile As Scripting.File
    Dim lngLines As Long

    If Not RepoIsAvailable Then Exit Sub
    If Not FSO.FolderExists(StripSlash(RulesFolderPath)) Then Exit Sub

    For Each oFile In FSO.GetFolder(StripSlash(RulesFolderPath)).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "mdc", vbTextCompare) = 0 Then
            lngLines = CountLines(oFile.Path)
            TestAssert lngLines <= MAX_RULE_LINES, _
                oFile.Name & " is " & lngLines & " lines, over the " & _
                MAX_RULE_LINES & " line budget"
        End If
    Next oFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestRootAgentsDocLinksResolve
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Every docs\ link in the routing table points at a file that exists. A stale
'           : link costs an agent a failed read and a fallback to searching.
'---------------------------------------------------------------------------------------
'
Public Sub TestRootAgentsDocLinksResolve()

    Dim varTarget As Variant

    If Not RepoIsAvailable Then Exit Sub

    For Each varTarget In GetDocLinkTargets(ReadFile(RepoRootPath & ROOT_ENTRY_FILE))
        TestAssert FSO.FileExists(RepoRootPath & Replace(varTarget, "/", PathSep)), _
            ROOT_ENTRY_FILE & " links to " & varTarget & ", which does not exist"
    Next varTarget

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestReferenceDocsAreLinked
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : A document nothing links to is read in a small fraction of sessions. Each
'           : one must be reachable from the root routing table, the docs\ index, or a
'           : Cursor rule that fires when the relevant files are opened.
'---------------------------------------------------------------------------------------
'
Public Sub TestReferenceDocsAreLinked()

    Dim strLinkSources As String
    Dim oFile As Scripting.File

    If Not RepoIsAvailable Then Exit Sub
    If Not FSO.FolderExists(StripSlash(DocsFolderPath)) Then Exit Sub

    strLinkSources = ReadFile(RepoRootPath & ROOT_ENTRY_FILE) & GetRuleFileContents
    If FSO.FileExists(DocsFolderPath & DOCS_INDEX_FILE) Then
        strLinkSources = strLinkSources & ReadFile(DocsFolderPath & DOCS_INDEX_FILE)
    End If

    For Each oFile In FSO.GetFolder(StripSlash(DocsFolderPath)).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "md", vbTextCompare) = 0 Then
            ' The index cannot vouch for its own reachability
            If StrComp(oFile.Name, DOCS_INDEX_FILE, vbTextCompare) <> 0 Then
                TestAssert InStr(1, strLinkSources, oFile.Name, vbTextCompare) > 0, _
                    oFile.Name & " is not linked from " & ROOT_ENTRY_FILE & _
                    ", a Cursor rule, or the docs index"
            End If
        End If
    Next oFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDocLinkTargets
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Return the repo-relative target of every markdown link into docs\, with any
'           : trailing anchor removed.
'---------------------------------------------------------------------------------------
'
Private Function GetDocLinkTargets(strContent As String) As Collection

    Dim colTargets As Collection
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strTarget As String

    Set colTargets = New Collection

    lngStart = InStr(1, strContent, DOCS_LINK_PREFIX, vbTextCompare)
    Do While lngStart > 0
        lngStart = lngStart + 2     ' step past the "](" that opens the target
        lngEnd = InStr(lngStart, strContent, ")")
        If lngEnd = 0 Then Exit Do
        strTarget = Mid$(strContent, lngStart, lngEnd - lngStart)
        If InStr(1, strTarget, "#") > 0 Then
            strTarget = Split(strTarget, "#")(0)
        End If
        If Len(strTarget) > 0 Then colTargets.Add strTarget
        lngStart = InStr(lngEnd, strContent, DOCS_LINK_PREFIX, vbTextCompare)
    Loop

    Set GetDocLinkTargets = colTargets

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRuleFileContents
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Concatenate every Cursor rule so link checks can search them in one pass.
'---------------------------------------------------------------------------------------
'
Private Function GetRuleFileContents() As String

    Dim oFile As Scripting.File

    If Not FSO.FolderExists(StripSlash(RulesFolderPath)) Then Exit Function

    For Each oFile In FSO.GetFolder(StripSlash(RulesFolderPath)).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "mdc", vbTextCompare) = 0 Then
            GetRuleFileContents = GetRuleFileContents & ReadFile(oFile.Path)
        End If
    Next oFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : RepoIsAvailable
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : These tests read the repository working tree, which is only present when the
'           : add-in runs from a checkout. Report a passing note rather than a failure
'           : otherwise, matching how modTestAgentDocs handles the same condition.
'---------------------------------------------------------------------------------------
'
Private Function RepoIsAvailable() As Boolean
    RepoIsAvailable = FSO.FileExists(RepoRootPath & ROOT_ENTRY_FILE)
    If Not RepoIsAvailable Then
        TestAssert True, "skipped: repository not available at " & RepoRootPath
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : RepoRootPath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Repository root, with a trailing separator. CodeProject.Path is the root
'           : when running from "<repo>\Version Control.accda".
'---------------------------------------------------------------------------------------
'
Private Function RepoRootPath() As String
    RepoRootPath = CodeProject.Path & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : DocsFolderPath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Full path to the reference document folder, with a trailing separator.
'---------------------------------------------------------------------------------------
'
Private Function DocsFolderPath() As String
    DocsFolderPath = RepoRootPath & DOCS_FOLDER & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : RulesFolderPath
' Author    : Adam Waller
' Date      : 8/6/2026
' Purpose   : Full path to the Cursor rules folder, with a trailing separator.
'---------------------------------------------------------------------------------------
'
Private Function RulesFolderPath() As String
    RulesFolderPath = RepoRootPath & RULES_FOLDER & PathSep
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
