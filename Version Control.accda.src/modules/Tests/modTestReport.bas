Attribute VB_Name = "modTestReport"
'---------------------------------------------------------------------------------------
' Module    : modTestReport
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Generate a self-contained HTML test-results dashboard from the durable
'           : test-state.json. The JSON is inlined into an embedded template so the
'           : report opens offline via file:// (no fetch/CORS).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Private Const ModuleName As String = "modTestReport"
Private Const RESULTS_FILE As String = "test-results.html"
Private Const RESOURCE_KEY As String = "Test Results HTML"
Private Const PLACEHOLDER As String = "__VCS_TEST_STATE_JSON__"
Private Const REPORT_CACHE_PREFIX As String = "MSAccessVCS_TestReport"

Private m_strTemplateCacheFolder As String


'---------------------------------------------------------------------------------------
' Procedure : ExportResultsHtml
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Write test-results.html from test-state.json and the embedded template.
'           : Returns the output path, or empty string when state/template is missing.
'---------------------------------------------------------------------------------------
'
Public Function ExportResultsHtml(Optional ByVal strPath As String = vbNullString) As String

    Const FunctionName As String = ModuleName & ".ExportResultsHtml"

    Dim strStatePath As String
    Dim strJson As String
    Dim strTemplate As String
    Dim strHtml As String
    Dim strEscaped As String

    strStatePath = modTestState.GetStateFilePath()
    If Not FSO.FileExists(strStatePath) Then Exit Function

    strJson = ReadFile(strStatePath)
    If Len(strJson) = 0 Then Exit Function

    strTemplate = ResolveResultsTemplate()
    If Len(strTemplate) = 0 Then Exit Function

    If Len(strPath) = 0 Then
        strPath = modTestState.GetTestResultsFolder() & RESULTS_FILE
    End If

    strEscaped = EscapeJsonForHtmlScript(strJson)
    strHtml = Replace(strTemplate, PLACEHOLDER, strEscaped, , , vbBinaryCompare)
    WriteFile strHtml, strPath

    Log.Add strPath
    ExportResultsHtml = strPath

    CatchAny eelWarning, vbNullString, FunctionName

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveResultsTemplate
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Return the results.html template text (embedded resource or dev source).
'---------------------------------------------------------------------------------------
'
Private Function ResolveResultsTemplate() As String

    Const FunctionName As String = ModuleName & ".ResolveResultsTemplate"

    Dim strFileName As String
    Dim strTarget As String
    Dim strSource As String

    strFileName = "results.html"

    If Len(m_strTemplateCacheFolder) = 0 Or Not FSO.FolderExists(m_strTemplateCacheFolder) Then
        m_strTemplateCacheFolder = GetTempFolder(REPORT_CACHE_PREFIX) & PathSep
        VerifyPath m_strTemplateCacheFolder & "placeholder"
    End If
    strTarget = m_strTemplateCacheFolder & strFileName

    If modResource.GetResourceHash(RESOURCE_KEY) <> vbNullString Then
        modResource.ExtractResource RESOURCE_KEY, m_strTemplateCacheFolder
    End If

    strSource = CodeProject.Path & PathSep & "TestRunner" & PathSep & strFileName

    If FSO.FileExists(strSource) Then
        If Not FSO.FileExists(strTarget) Then
            FSO.CopyFile strSource, strTarget, True
        ElseIf FSO.GetFile(strSource).DateLastModified > FSO.GetFile(strTarget).DateLastModified Then
            FSO.CopyFile strSource, strTarget, True
        End If
    End If

    If Not FSO.FileExists(strTarget) Then
        Log.Add T("Test results HTML template not found (resource '{0}' and source both missing).", _
            var0:=RESOURCE_KEY), , , "red"
        Exit Function
    End If

    ResolveResultsTemplate = ReadFile(strTarget)

    CatchAny eelWarning, vbNullString, FunctionName

End Function


'---------------------------------------------------------------------------------------
' Procedure : EscapeJsonForHtmlScript
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Prevent </script> breakout when inlining JSON inside a script tag.
'---------------------------------------------------------------------------------------
'
Private Function EscapeJsonForHtmlScript(ByVal strJson As String) As String

    EscapeJsonForHtmlScript = Replace$(strJson, "<", "\u003c")

End Function
