Attribute VB_Name = "modTestFolderPlacement"
'---------------------------------------------------------------------------------------
' Module    : modTestFolderPlacement
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Tests for @Folder placement helpers and duplicate source cleanup.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestGetFolderAnnotationFromText_ParsesDotSegments()
    Dim strCode As String
    strCode = "'@Folder(""Core.Utility"")" & vbCrLf & "Option Explicit"
    TestAssert GetFolderAnnotationFromText(vbCrLf & strCode, "C:\repo\modules\") = _
        "Core" & PathSep & "Utility" & PathSep, "@Folder dots become path separators"
End Sub


Public Sub TestGetFolderAnnotationFromText_EmptyWhenMissing()
    TestAssert Len(GetFolderAnnotationFromText("Option Explicit", "C:\repo\modules\")) = 0, _
        "no annotation returns empty"
End Sub


Public Sub TestRemoveDuplicateModuleFiles_DeletesMisplacedCopy()
    Dim strRoot As String
    Dim strBase As String
    Dim strMisplaced As String
    Dim strCanonical As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "modules" & PathSep
    VerifyPath strBase & "Tests" & PathSep

    strMisplaced = strBase & "modDupPlacement.bas"
    strCanonical = strBase & "Tests\modDupPlacement.bas"
    WriteFile BuildTestModuleSource("modDupPlacement", "Tests"), strMisplaced
    WriteFile BuildTestModuleSource("modDupPlacement", "Tests"), strCanonical

    RemoveDuplicateModuleFiles strBase

    TestAssert FSO.FileExists(strCanonical), "canonical @Folder copy kept"
    TestAssert Not FSO.FileExists(strMisplaced), "misplaced root copy removed"

    DeleteFolderPlacementFixture strRoot
End Sub


Public Sub TestRemoveDuplicateModuleFiles_PreservesAmbiguousDuplicates()
    Dim strRoot As String
    Dim strBase As String
    Dim strCopy1 As String
    Dim strCopy2 As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "modules" & PathSep
    VerifyPath strBase & "Core" & PathSep

    ' Both copies claim @Folder Tests but sit in wrong folders -> no single canonical winner.
    strCopy1 = strBase & "modAmbiguousDup.bas"
    strCopy2 = strBase & "Core\modAmbiguousDup.bas"
    WriteFile BuildTestModuleSource("modAmbiguousDup", "Tests"), strCopy1
    WriteFile BuildTestModuleSource("modAmbiguousDup", "Tests"), strCopy2

    RemoveDuplicateModuleFiles strBase

    TestAssert FSO.FileExists(strCopy1), "ambiguous duplicate 1 preserved"
    TestAssert FSO.FileExists(strCopy2), "ambiguous duplicate 2 preserved"

    DeleteFolderPlacementFixture strRoot
End Sub


Public Sub TestRemoveDuplicateFormFiles_DeletesMisplacedCopy()
    Dim strRoot As String
    Dim strBase As String
    Dim strMisplacedFolder As String
    Dim strCanonicalFolder As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "forms" & PathSep
    strMisplacedFolder = strBase
    strCanonicalFolder = strBase & "Tests" & PathSep
    VerifyPath strCanonicalFolder

    WriteFormPlacementFixture strMisplacedFolder, "frmDupPlacement", "Tests"
    WriteFormPlacementFixture strCanonicalFolder, "frmDupPlacement", "Tests"

    RemoveDuplicateFormFiles strBase

    TestAssert FSO.FileExists(strCanonicalFolder & "frmDupPlacement.form"), "canonical form kept"
    TestAssert FSO.FileExists(strCanonicalFolder & "frmDupPlacement.cls"), "canonical cls kept"
    TestAssert FSO.FileExists(strCanonicalFolder & "frmDupPlacement.json"), "canonical json kept"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "frmDupPlacement.form"), "misplaced form removed"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "frmDupPlacement.cls"), "misplaced cls removed"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "frmDupPlacement.json"), "misplaced json removed"

    DeleteFolderPlacementFixture strRoot
End Sub


Public Sub TestRemoveDuplicateFormFiles_IgnoresSingleFolderGroup()
    Dim strRoot As String
    Dim strBase As String
    Dim strFolder As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "forms" & PathSep
    strFolder = strBase & "Tests" & PathSep
    VerifyPath strFolder

    WriteFormPlacementFixture strFolder, "frmSingleGroup", "Tests"

    RemoveDuplicateFormFiles strBase

    TestAssert FSO.FileExists(strFolder & "frmSingleGroup.form"), "single-folder form untouched"
    TestAssert FSO.FileExists(strFolder & "frmSingleGroup.cls"), "single-folder cls untouched"
    TestAssert FSO.FileExists(strFolder & "frmSingleGroup.json"), "single-folder json untouched"

    DeleteFolderPlacementFixture strRoot
End Sub


Public Sub TestRemoveDuplicateFormFiles_AnnotationFromClsSidecar()
    Dim strRoot As String
    Dim strBase As String
    Dim strMisplacedFolder As String
    Dim strCanonicalFolder As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "forms" & PathSep
    strMisplacedFolder = strBase
    strCanonicalFolder = strBase & "Core" & PathSep
    VerifyPath strCanonicalFolder

    ' Annotation only in .cls; .form has no @Folder line.
    WriteFile BuildTestFormPrimarySource("frmClsAnnotation"), strMisplacedFolder & "frmClsAnnotation.form"
    WriteFile BuildTestFormClsSource("frmClsAnnotation", "Core"), strMisplacedFolder & "frmClsAnnotation.cls"
    WriteFile "{}", strMisplacedFolder & "frmClsAnnotation.json"

    WriteFile BuildTestFormPrimarySource("frmClsAnnotation"), strCanonicalFolder & "frmClsAnnotation.form"
    WriteFile BuildTestFormClsSource("frmClsAnnotation", "Core"), strCanonicalFolder & "frmClsAnnotation.cls"
    WriteFile "{}", strCanonicalFolder & "frmClsAnnotation.json"

    RemoveDuplicateFormFiles strBase

    TestAssert FSO.FileExists(strCanonicalFolder & "frmClsAnnotation.form"), "cls-annotated canonical kept"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "frmClsAnnotation.form"), "cls-annotated misplaced removed"

    DeleteFolderPlacementFixture strRoot
End Sub


Public Sub TestRemoveDuplicateReportFiles_DeletesMisplacedCopy()
    Dim strRoot As String
    Dim strBase As String
    Dim strMisplacedFolder As String
    Dim strCanonicalFolder As String

    strRoot = GetTempFolder("vcs_folder_placement") & PathSep
    strBase = strRoot & "reports" & PathSep
    strMisplacedFolder = strBase
    strCanonicalFolder = strBase & "Tests" & PathSep
    VerifyPath strCanonicalFolder

    WriteReportPlacementFixture strMisplacedFolder, "rptDupPlacement", "Tests"
    WriteReportPlacementFixture strCanonicalFolder, "rptDupPlacement", "Tests"

    RemoveDuplicateReportFiles strBase

    TestAssert FSO.FileExists(strCanonicalFolder & "rptDupPlacement.report"), "canonical report kept"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "rptDupPlacement.report"), "misplaced report removed"
    TestAssert Not FSO.FileExists(strMisplacedFolder & "rptDupPlacement.cls"), "misplaced report cls removed"

    DeleteFolderPlacementFixture strRoot
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildTestModuleSource
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Minimal .bas source with Attribute VB_Name and optional @Folder annotation.
'---------------------------------------------------------------------------------------
'
Private Function BuildTestModuleSource(strModuleName As String, _
    Optional strFolder As String = vbNullString) As String

    Dim cOut As New clsConcat
    cOut.Add "Attribute VB_Name = """ & strModuleName & """" & vbCrLf
    If Len(strFolder) > 0 Then cOut.Add "'@Folder(""" & strFolder & """)" & vbCrLf
    cOut.Add "Option Explicit" & vbCrLf
    BuildTestModuleSource = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteFormPlacementFixture
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Write a minimal form source group (.form, .cls, .json) for placement tests.
'---------------------------------------------------------------------------------------
'
Private Sub WriteFormPlacementFixture(strFolder As String, strFormName As String, strAnnotationFolder As String)

    WriteFile BuildTestFormPrimarySource(strFormName), strFolder & strFormName & ".form"
    WriteFile BuildTestFormClsSource(strFormName, strAnnotationFolder), strFolder & strFormName & ".cls"
    WriteFile "{}", strFolder & strFormName & ".json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteReportPlacementFixture
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Write a minimal report source group (.report, .cls, .json) for placement tests.
'---------------------------------------------------------------------------------------
'
Private Sub WriteReportPlacementFixture(strFolder As String, strReportName As String, strAnnotationFolder As String)

    WriteFile BuildTestReportPrimarySource(strReportName), strFolder & strReportName & ".report"
    WriteFile BuildTestFormClsSource(strReportName, strAnnotationFolder), strFolder & strReportName & ".cls"
    WriteFile "{}", strFolder & strReportName & ".json"

End Sub


'---------------------------------------------------------------------------------------
' Function  : BuildTestFormPrimarySource
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Minimal .form layout text without @Folder (annotation lives in .cls).
'---------------------------------------------------------------------------------------
'
Private Function BuildTestFormPrimarySource(strFormName As String) As String

    BuildTestFormPrimarySource = "Version =21" & vbCrLf & "Begin Form" & vbCrLf & _
        "    Name =""" & strFormName & """" & vbCrLf & "End" & vbCrLf

End Function


'---------------------------------------------------------------------------------------
' Function  : BuildTestReportPrimarySource
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Minimal .report layout text without @Folder (annotation lives in .cls).
'---------------------------------------------------------------------------------------
'
Private Function BuildTestReportPrimarySource(strReportName As String) As String

    BuildTestReportPrimarySource = "Version =21" & vbCrLf & "Begin Report" & vbCrLf & _
        "    Name =""" & strReportName & """" & vbCrLf & "End" & vbCrLf

End Function


'---------------------------------------------------------------------------------------
' Function  : BuildTestFormClsSource
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Minimal form/report code-behind .cls with optional @Folder annotation.
'---------------------------------------------------------------------------------------
'
Private Function BuildTestFormClsSource(strObjectName As String, _
    Optional strFolder As String = vbNullString) As String

    Dim cOut As New clsConcat
    cOut.Add "VERSION 1.0 CLASS" & vbCrLf
    cOut.Add "BEGIN" & vbCrLf
    cOut.Add "  MultiUse = -1  'True" & vbCrLf
    cOut.Add "END" & vbCrLf
    cOut.Add "Attribute VB_Name = ""Form_" & strObjectName & """" & vbCrLf
    If Len(strFolder) > 0 Then cOut.Add "'@Folder(""" & strFolder & """)" & vbCrLf
    cOut.Add "Option Explicit" & vbCrLf
    BuildTestFormClsSource = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteFolderPlacementFixture
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Remove temp folder tree created by folder-placement tests.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteFolderPlacementFixture(strRoot As String)

    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strRoot) Then FSO.DeleteFolder StripSlash(strRoot), True
    Err.Clear

End Sub
