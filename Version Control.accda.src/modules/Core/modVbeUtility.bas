Attribute VB_Name = "modVbeUtility"
'---------------------------------------------------------------------------------------
' Module    : modVbeUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : VBA Editor (VBE) operations including code module export, overlay,
'           : compilation, reference management, and add-in project loading.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modDatabase, modFileAccess, modEncoding,
'           : modErrorHandling, modInstall
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName = "modVbeUtility"


'---------------------------------------------------------------------------------------
' Procedure : ExportVbComponent
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Export the code module VB component and convert to UTF-8
'---------------------------------------------------------------------------------------
'
Public Sub ExportCodeModule(strName As String, strFile As String)

    Dim strTempFile As String
    Dim strContent As String

    Perf.OperationStart "Export VBE Module"

    ' Export to a temp file so we can convert to UTF-8 encoding
    strTempFile = GetTempFile
    CurrentVBProject.VBComponents(strName).Export strTempFile

    ' Sanitize the VBA code while reading the temp file
    With New clsSourceParser
        .LoadString ReadFile(strTempFile, GetSystemEncoding), edbModule
        strContent = .Sanitize(ectVBA)
    End With

    ' Write the content as UTF-8 to the final destination
    WriteFile strContent, strFile
    DeleteFile strTempFile

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFolderAnnotation
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Return the subfolder path from a Rubberduck-style '@Folder annotation
'           : in a VBE code module. Dots are converted to path separators.
'           : Returns empty string if no annotation found or component has no code.
'           : Example: '@Folder("Core.Utility") returns "Core\Utility\"
'           : When the folder already exists on disk, the returned path uses the
'           : actual disk casing rather than the annotation casing.
'---------------------------------------------------------------------------------------
'
Public Function GetFolderAnnotation(cComponent As IDbComponent, _
    Optional blnBypassCache As Boolean = False) As String

    Dim cmpItem As VBComponent
    Dim strCode As String
    Dim strVBEName As String
    Dim strCached As String

    LogUnhandledErrors
    On Error Resume Next

    ' Check index cache before making any VBE COM calls.
    ' The sentinel FOLDER_ANNOTATION_NONE means "checked, no annotation found."
    ' Export callers pass blnBypassCache:=True to read the live annotation
    ' from VBE, since the cached value may be stale after an annotation edit.
    If Not blnBypassCache Then
        strCached = TryGetCachedAnnotation(cComponent)
        If Len(strCached) > 0 Then
            If strCached <> FOLDER_ANNOTATION_NONE Then
                GetFolderAnnotation = strCached
            End If
            Exit Function
        End If
    End If

    ' Determine VBE component name from the database component type
    Select Case cComponent.ComponentType
        Case edbForm:   strVBEName = "Form_" & cComponent.Name
        Case edbReport: strVBEName = "Report_" & cComponent.Name
        Case Else:      strVBEName = cComponent.Name
    End Select

    Perf.OperationStart "Get @Folder Annotation"

    ' Attempt to locate the component in the VBE
    Set cmpItem = CurrentVBProject.VBComponents(strVBEName)
    If cmpItem Is Nothing Then GoTo CleanUp
    If cmpItem.CodeModule Is Nothing Then GoTo CleanUp
    If cmpItem.CodeModule.CountOfLines = 0 Then GoTo CleanUp

    ' Read all code in a single COM call and prepend vbCrLf so that
    ' a line-1 annotation is found by the same pattern as any other line.
    strCode = vbCrLf & cmpItem.CodeModule.Lines(1, 999999)
    GetFolderAnnotation = GetFolderAnnotationFromText(strCode, cComponent.BaseFolder, strVBEName)

    CatchAny eelError, "Error reading @Folder annotation for " & strVBEName, ModuleName & ".GetFolderAnnotation"

CleanUp:
    Perf.OperationEnd
    If Err Then Err.Clear


End Function


'---------------------------------------------------------------------------------------
' Function  : GetFolderAnnotationFromText
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Parse a Rubberduck-style '@Folder annotation from module source text.
'           : Dots become path separators; each segment is sanitized. When the
'           : folder already exists on disk, the returned path uses actual casing.
'           : strContext is optional (module name) for duplicate-annotation warnings.
'---------------------------------------------------------------------------------------
'
Public Function GetFolderAnnotationFromText(ByVal strCode As String, _
    ByVal strBaseFolder As String, Optional ByVal strContext As String = vbNullString) As String

    Dim strUpper As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    Const TAG As String = "'@FOLDER("

    If Len(strCode) = 0 Then Exit Function

    ' Match the VBE reader: prepend vbCrLf so a line-1 annotation is found.
    If Left$(strCode, 2) <> vbCrLf Then strCode = vbCrLf & strCode
    strUpper = UCase$(strCode)

    lngPos = InStr(1, strUpper, vbCrLf & TAG)
    If lngPos = 0 Then Exit Function
    lngPos = lngPos + 2

    If InStr(lngPos + Len(TAG), strUpper, vbCrLf & TAG) > 0 Then
        Log.Add T("WARNING: Multiple @Folder annotations found in {0}. Using first annotation.", _
            var0:=IIf(Len(strContext) > 0, strContext, T("module source"))), Options.ShowDebug
    End If

    lngStart = InStr(lngPos, strCode, """")
    If lngStart > 0 Then
        lngEnd = InStr(lngStart + 1, strCode, """")
        If lngEnd > lngStart + 1 Then
            Dim varSegments As Variant
            Dim lngSeg As Long
            varSegments = Split(Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1), ".")
            For lngSeg = LBound(varSegments) To UBound(varSegments)
                varSegments(lngSeg) = GetSafeFileName(CStr(varSegments(lngSeg)))
            Next lngSeg
            GetFolderAnnotationFromText = Join(varSegments, PathSep) & PathSep
        End If
    End If

    If Len(GetFolderAnnotationFromText) > 0 Then
        Dim varParts As Variant
        Dim strResolved As String
        Dim strCheckPath As String
        Dim lngPart As Long
        varParts = Split(Left$(GetFolderAnnotationFromText, Len(GetFolderAnnotationFromText) - 1), PathSep)
        strCheckPath = StripSlash(strBaseFolder)
        For lngPart = LBound(varParts) To UBound(varParts)
            strCheckPath = strCheckPath & PathSep & varParts(lngPart)
            If FSO.FolderExists(strCheckPath) Then
                strResolved = strResolved & FSO.GetFolder(strCheckPath).Name & PathSep
                strCheckPath = FSO.GetFolder(strCheckPath).Path
            Else
                strResolved = strResolved & varParts(lngPart) & PathSep
            End If
        Next lngPart
        GetFolderAnnotationFromText = strResolved
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveDuplicateModuleFiles
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Before build/import, scan the modules tree for duplicate basenames
'           : (.bas/.cls). When exactly one copy sits in its @Folder-derived path,
'           : delete the misplaced copies. Ambiguous groups are left in place with
'           : a warning so import does not silently last-one-wins.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveDuplicateModuleFiles(strBaseFolder As String)
    RemoveDuplicateComponentFiles strBaseFolder, Array("bas", "cls"), vbNullString, Array("json")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveDuplicateFormFiles
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Build-time duplicate cleanup for forms. @Folder is read from the .cls
'           : code-behind when present, otherwise from the primary .form/.bas file.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveDuplicateFormFiles(strBaseFolder As String)
    RemoveDuplicateComponentFiles strBaseFolder, Array("form", "bas"), "cls", Array("cls", "json", "svg")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveDuplicateReportFiles
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Build-time duplicate cleanup for reports. @Folder is read from the .cls
'           : code-behind when present, otherwise from the primary .report/.bas file.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveDuplicateReportFiles(strBaseFolder As String)
    RemoveDuplicateComponentFiles strBaseFolder, Array("report", "bas"), "cls", Array("cls", "json", "svg")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveDuplicateComponentFiles
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Before build/import, scan a component tree for duplicate basenames in
'           : distinct folders. When exactly one instance sits in its @Folder-derived
'           : path, delete misplaced copies (primary + companions). Ambiguous groups
'           : are left in place with a warning so import does not silently last-one-wins.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveDuplicateComponentFiles(strBaseFolder As String, _
    varPrimaryExts As Variant, strAnnotationExt As String, varCompanionExts As Variant)

    Dim dGroups As Dictionary
    Dim dInstances As Dictionary
    Dim varKey As Variant

    If Not FSO.FolderExists(strBaseFolder) Then Exit Sub

    Set dGroups = New Dictionary
    CollectComponentInstancesForDuplicateScan StripSlash(strBaseFolder), dGroups, varPrimaryExts

    For Each varKey In dGroups.Keys
        Set dInstances = dGroups(varKey)
        If dInstances.Count > 1 Then
            ResolveDuplicateComponentGroup strBaseFolder, CStr(varKey), dInstances, _
                strAnnotationExt, varCompanionExts
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarnDuplicateModuleBasenames
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Log a warning when duplicate module basenames remain in the tree.
'           : Used after export as a safety net when per-module cleanup missed a copy.
'---------------------------------------------------------------------------------------
'
Public Sub WarnDuplicateModuleBasenames(strBaseFolder As String)
    WarnDuplicateComponentBasenames strBaseFolder, Array("bas", "cls"), "module"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarnDuplicateFormBasenames
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Log a warning when duplicate form basenames remain in the tree.
'---------------------------------------------------------------------------------------
'
Public Sub WarnDuplicateFormBasenames(strBaseFolder As String)
    WarnDuplicateComponentBasenames strBaseFolder, Array("form", "bas"), "form"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarnDuplicateReportBasenames
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Log a warning when duplicate report basenames remain in the tree.
'---------------------------------------------------------------------------------------
'
Public Sub WarnDuplicateReportBasenames(strBaseFolder As String)
    WarnDuplicateComponentBasenames strBaseFolder, Array("report", "bas"), "report"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarnDuplicateComponentBasenames
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Log a warning when duplicate component basenames remain in distinct folders.
'---------------------------------------------------------------------------------------
'
Public Sub WarnDuplicateComponentBasenames(strBaseFolder As String, _
    varPrimaryExts As Variant, strComponentLabel As String)

    Dim dGroups As Dictionary
    Dim dInstances As Dictionary
    Dim varKey As Variant

    If Not FSO.FolderExists(strBaseFolder) Then Exit Sub

    Set dGroups = New Dictionary
    CollectComponentInstancesForDuplicateScan StripSlash(strBaseFolder), dGroups, varPrimaryExts

    For Each varKey In dGroups.Keys
        Set dInstances = dGroups(varKey)
        If dInstances.Count > 1 Then
            Log.Add T("WARNING: Duplicate {0} source files for {1}: {2}", _
                var0:=strComponentLabel, var1:=CStr(varKey), _
                var2:=JoinDictionaryInstancePaths(dInstances)), Options.ShowDebug
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CollectComponentInstancesForDuplicateScan
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Recursively collect primary source files grouped by basename and folder.
'           : Companion files (.cls/.json/.svg for forms) do not create instances.
'---------------------------------------------------------------------------------------
'
Private Sub CollectComponentInstancesForDuplicateScan(strFolder As String, _
    dGroups As Dictionary, varPrimaryExts As Variant)

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim dInstances As Dictionary
    Dim varItem As Variant
    Dim strName As String
    Dim strBaseName As String
    Dim strInstanceFolder As String

    ScanFolderContents strFolder, colFiles, colSubFolders

    For Each varItem In colFiles
        strName = FSO.GetFileName(CStr(varItem))
        If FileMatchesPrimaryExt(strName, varPrimaryExts) Then
            strBaseName = FSO.GetBaseName(strName)
            strInstanceFolder = AddSlash(FSO.GetParentFolderName(CStr(varItem)))
            If dGroups.Exists(strBaseName) Then
                Set dInstances = dGroups(strBaseName)
            Else
                Set dInstances = New Dictionary
                dGroups.Add strBaseName, dInstances
            End If
            If Not dInstances.Exists(strInstanceFolder) Then
                dInstances.Add strInstanceFolder, CStr(varItem)
            End If
        End If
    Next varItem

    For Each varItem In colSubFolders
        CollectComponentInstancesForDuplicateScan CStr(varItem), dGroups, varPrimaryExts
    Next varItem

End Sub


'---------------------------------------------------------------------------------------
' Function  : FileMatchesPrimaryExt
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Return True when strFileName ends with one of the primary extensions.
'---------------------------------------------------------------------------------------
'
Private Function FileMatchesPrimaryExt(strFileName As String, varPrimaryExts As Variant) As Boolean

    Dim varExt As Variant
    For Each varExt In varPrimaryExts
        If strFileName Like "*." & CStr(varExt) Then
            FileMatchesPrimaryExt = True
            Exit Function
        End If
    Next varExt

End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveDuplicateComponentGroup
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : For one duplicate basename, delete misplaced folder instances when
'           : exactly one instance is in its @Folder-derived location.
'---------------------------------------------------------------------------------------
'
Private Sub ResolveDuplicateComponentGroup(strBaseFolder As String, _
    strBaseName As String, dInstances As Dictionary, _
    strAnnotationExt As String, varCompanionExts As Variant)

    Dim colCanonical As New Collection
    Dim colMisplaced As New Collection
    Dim varFolder As Variant
    Dim strPrimaryFile As String
    Dim strCode As String
    Dim strAnnotation As String
    Dim strExpectedFolder As String
    Dim strActualFolder As String

    strBaseFolder = AddSlash(StripSlash(strBaseFolder))

    For Each varFolder In dInstances.Keys
        strPrimaryFile = CStr(dInstances(varFolder))
        strActualFolder = CStr(varFolder)
        strCode = ReadComponentAnnotationCode(strActualFolder, strBaseName, _
            strPrimaryFile, strAnnotationExt)
        strAnnotation = GetFolderAnnotationFromText(strCode, strBaseFolder, strBaseName)
        strExpectedFolder = strBaseFolder & strAnnotation
        If StrComp(strActualFolder, strExpectedFolder, vbTextCompare) = 0 Then
            colCanonical.Add strPrimaryFile
        Else
            colMisplaced.Add strPrimaryFile
        End If
    Next varFolder

    If colCanonical.Count = 1 And colMisplaced.Count > 0 Then
        For Each varFolder In colMisplaced
            strPrimaryFile = CStr(varFolder)
            DeleteMisplacedComponentInstance AddSlash(FSO.GetParentFolderName(strPrimaryFile)), _
                strBaseName, strPrimaryFile, varCompanionExts
            Log.Add T("Removed duplicate source file: {0}", var0:=strPrimaryFile), Options.ShowDebug
        Next varFolder
        RemoveEmptyModuleSubfolders StripSlash(strBaseFolder)
    ElseIf dInstances.Count > 1 Then
        Log.Add T("WARNING: Ambiguous duplicate source files for {0}; not auto-deleting. Paths: {1}", _
            var0:=strBaseName, var1:=JoinDictionaryInstancePaths(dInstances)), Options.ShowDebug
    End If

End Sub


'---------------------------------------------------------------------------------------
' Function  : ReadComponentAnnotationCode
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Read source text containing @Folder for a component instance. Prefer
'           : the annotation sidecar (.cls for forms/reports) when present.
'---------------------------------------------------------------------------------------
'
Private Function ReadComponentAnnotationCode(strFolder As String, strBaseName As String, _
    strPrimaryFile As String, strAnnotationExt As String) As String

    Dim strAnnotationFile As String

    If Len(strAnnotationExt) > 0 Then
        strAnnotationFile = strFolder & strBaseName & "." & strAnnotationExt
        If FSO.FileExists(strAnnotationFile) Then
            ReadComponentAnnotationCode = ReadFile(strAnnotationFile)
            Exit Function
        End If
    End If
    ReadComponentAnnotationCode = ReadFile(strPrimaryFile)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteMisplacedComponentInstance
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Delete a misplaced component instance: primary file plus companions.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteMisplacedComponentInstance(strFolder As String, strBaseName As String, _
    strPrimaryFile As String, varCompanionExts As Variant)

    Dim varExt As Variant
    Dim strCompanion As String

    DeleteFile strPrimaryFile
    For Each varExt In varCompanionExts
        strCompanion = strFolder & strBaseName & "." & CStr(varExt)
        If FSO.FileExists(strCompanion) Then DeleteFile strCompanion
    Next varExt

End Sub


'---------------------------------------------------------------------------------------
' Function  : JoinDictionaryInstancePaths
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Join primary file paths from an instance dictionary for diagnostics.
'---------------------------------------------------------------------------------------
'
Private Function JoinDictionaryInstancePaths(dInstances As Dictionary) As String

    Dim varItem As Variant
    Dim cOut As New clsConcat
    For Each varItem In dInstances.Items
        cOut.Add CStr(varItem), "; "
    Next varItem
    JoinDictionaryInstancePaths = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveEmptyModuleSubfolders
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Remove empty subfolders under the modules base folder after duplicate cleanup.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveEmptyModuleSubfolders(strBaseFolder As String)

    Dim colSubFolders As New Collection
    Dim colFiles As New Collection
    Dim varItem As Variant

    ScanFolderContents strBaseFolder, colFiles, colSubFolders
    For Each varItem In colSubFolders
        RemoveEmptyModuleSubfolders CStr(varItem)
        If FSO.FolderExists(CStr(varItem)) Then
            If FSO.GetFolder(CStr(varItem)).Files.Count = 0 _
                And FSO.GetFolder(CStr(varItem)).SubFolders.Count = 0 Then
                LogUnhandledErrors
                On Error Resume Next
                FSO.DeleteFolder CStr(varItem), True
                CatchAny eelWarning, "Unable to delete empty folder: " & CStr(varItem), _
                    ModuleName & ".RemoveEmptyModuleSubfolders"
            End If
        End If
    Next varItem

End Sub


'---------------------------------------------------------------------------------------
' Function  : JoinCollectionPaths
' Author    : Adam Waller
' Date      : 6/18/2026
' Purpose   : Join collection items into a semicolon-delimited diagnostic string.
'---------------------------------------------------------------------------------------
'
Private Function JoinCollectionPaths(colItems As Collection) As String

    Dim varItem As Variant
    Dim cOut As New clsConcat
    For Each varItem In colItems
        cOut.Add CStr(varItem), "; "
    Next varItem
    JoinCollectionPaths = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : CleanupDuplicateSourceFiles
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : After exporting a component, scan the base folder tree for any file
'           : matching strSafeName + one of the given extensions that lives in a
'           : subfolder OTHER than strCorrectFolder. Deletes duplicates regardless of
'           : how they were introduced (stale cache, manual copy, git merge, etc.).
'           : Empty subfolders are removed after cleanup.
'---------------------------------------------------------------------------------------
'
Public Sub CleanupDuplicateSourceFiles(strBaseFolder As String, _
    strCorrectFolder As String, strSafeName As String, _
    ParamArray varExtensions() As Variant)

    ' Copy ParamArray into a plain Variant so it can be forwarded to the recursive helper
    Dim varExts As Variant
    varExts = varExtensions

    If StrComp(StripSlash(strBaseFolder), StripSlash(strCorrectFolder), vbTextCompare) = 0 Then Exit Sub
    If Not FSO.FolderExists(strBaseFolder) Then Exit Sub

    ' Recursive scan of the base folder tree
    ScanForDuplicates StripSlash(strBaseFolder), strCorrectFolder, strSafeName, varExts

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ScanForDuplicates
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Recursively scan a folder and subfolders, deleting any file matching
'           : the target name + extensions that is not in the correct folder.
'---------------------------------------------------------------------------------------
'
Private Sub ScanForDuplicates(strFolder As String, strCorrectFolder As String, _
    strSafeName As String, varExtensions As Variant)

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim varItem As Variant
    Dim strName As String
    Dim strParent As String
    Dim i As Long

    ScanFolderContents strFolder, colFiles, colSubFolders

    ' Only check files if this is NOT the correct folder
    strParent = AddSlash(strFolder)
    If StrComp(strParent, strCorrectFolder, vbTextCompare) <> 0 Then
        For Each varItem In colFiles
            strName = FSO.GetFileName(CStr(varItem))
            For i = LBound(varExtensions) To UBound(varExtensions)
                If StrComp(strName, strSafeName & CStr(varExtensions(i)), vbTextCompare) = 0 Then
                    DeleteFile CStr(varItem)
                    Exit For
                End If
            Next i
        Next varItem
    End If

    ' Recurse into subfolders
    For Each varItem In colSubFolders
        ScanForDuplicates CStr(varItem), strCorrectFolder, strSafeName, varExtensions
        ' Remove subfolder if empty after cleanup
        If FSO.FolderExists(CStr(varItem)) Then
            If FSO.GetFolder(CStr(varItem)).Files.Count = 0 _
                And FSO.GetFolder(CStr(varItem)).SubFolders.Count = 0 Then
                LogUnhandledErrors
                On Error Resume Next
                FSO.DeleteFolder CStr(varItem), True
                CatchAny eelWarning, "Unable to delete empty folder: " & CStr(varItem), _
                    ModuleName & ".ScanForDuplicates"
            End If
        End If
    Next varItem

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TryGetCachedAnnotation
' Author    : Adam Waller
' Date      : 5/5/2026
' Purpose   : Try to retrieve a cached @Folder annotation from the VCS index using the
'           : component's category and name, avoiding the SourceFile property (which
'           : depends on the annotation). Returns vbNullString on cache miss.
'           : For modules, tries both .bas and .cls extensions since the type is not
'           : known without a VBE COM call.
'---------------------------------------------------------------------------------------
'
Private Function TryGetCachedAnnotation(cComponent As IDbComponent) As String

    Dim strSafeName As String
    Dim strResult As String

    strSafeName = GetSafeFileName(cComponent.Name)

    Select Case cComponent.ComponentType
        Case edbModule
            strResult = VCSIndex.GetCachedAnnotation("Modules", strSafeName & ".bas")
            If Len(strResult) = 0 Then
                strResult = VCSIndex.GetCachedAnnotation("Modules", strSafeName & ".cls")
            End If
        Case edbForm
            strResult = VCSIndex.GetCachedAnnotation("Forms", strSafeName & ".form")
        Case edbReport
            strResult = VCSIndex.GetCachedAnnotation("Reports", strSafeName & ".report")
        Case edbVbeForm
            strResult = VCSIndex.GetCachedAnnotation("VBE Forms", strSafeName & ".json")
    End Select

    TryGetCachedAnnotation = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : OverlayCodeModule
' Author    : Adam Waller
' Date      : 10/24/2023
' Purpose   : Overlay VBA code from an object's *.cls file to the form or report
'           : Note that this opens the object in design view, which may slow the build
'           : process if a large number of items are invovled.
'---------------------------------------------------------------------------------------
'
Public Sub OverlayCodeModule(strName As String _
                            , strClassFile As String)

    Const FunctionName As String = ModuleName & ".OverlayCodeModule"

    Dim objModule As VBIDE.CodeModule
    Dim strContent As String
    Dim intType As AcObjectType
    Dim strShortName As String
    Dim cParser As clsSourceParser

    LogUnhandledErrors FunctionName
    'On Error Resume Next

    Set objModule = CurrentVBProject.VBComponents(strName).CodeModule
    If CatchAny(eelError, T("Could not find code module for {0}", var0:=strName), FunctionName) Then Exit Sub

    ' Read class file content
    strContent = ReadFile(strClassFile)
    If strContent = vbNullString Then
        Log.Error eelError, T("Unable to read {0}", var0:=strClassFile), FunctionName
        Exit Sub
    End If

    ' Get object type and short name
    If strName Like "Form_*" Then
        intType = acForm
        strShortName = Mid$(strName, 6)
        DoCmd.OpenForm strShortName, acDesign, , , , acHidden
    ElseIf strName Like "Report_*" Then
        intType = acReport
        strShortName = Mid$(strName, 8)
        DoCmd.OpenReport strShortName, acViewDesign, , , acHidden
    End If

    ' Overlay the VBA code, replacing any existing code.
    Set cParser = New clsSourceParser
    objModule.DeleteLines 1, objModule.CountOfLines
    objModule.AddFromString cParser.StripClassHeader(strContent, False)

    ' Close any form or report object
    Select Case intType
        Case acForm, acReport
            DoCmd.Close intType, strShortName, acSaveYes
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveNonBuiltInReferences
' Author    : Adam Waller
' Date      : 10/20/2020
' Purpose   : Remove any references that are not built-in. (Sometimes additional
'           : references are added when creating a new database, but not not really
'           : needed in the completed database when building the project from source.)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveNonBuiltInReferences()

    Dim intCnt As Integer
    Dim strName As String
    Dim ref As Access.Reference

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Perf.OperationStart "Clear References"
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If Not ref.BuiltIn Then
            strName = ref.Name
            Application.References.Remove ref
            If CatchAny(eelError, "Failed to remove reference: " & strName, _
                ModuleName & ".RemoveNonBuiltInReferences", True, True, True) Then
                Log.Add "  ERROR removing " & strName, False
            Else
                Log.Add "  Removed " & strName, False
            End If
        End If
        Set ref = Nothing
    Next intCnt
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompileAndSaveAllModules
' Author    : Adam Waller
' Date      : 7/10/2021
' Purpose   : Compile and save the modules in the current database
'---------------------------------------------------------------------------------------
'
Public Sub CompileAndSaveAllModules()
    Perf.OperationStart "Compile/Save Modules"
    ' Make sure we are running this in the CurrentDB, not the CodeDB
    Set VBE.ActiveVBProject = CurrentVBProject
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    DoEvents
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MinimizeVBEWindow
' Author    : Adam Waller
' Date      : 4/9/2026
' Purpose   : Minimize the VBE main window if it is currently visible. This prevents
'           : the VBE from stealing focus during builds and exports when VBA components
'           : are imported. The window is left minimized so the user can restore it
'           : from the taskbar after the operation completes.
'---------------------------------------------------------------------------------------
'
Public Sub MinimizeVBEWindow()
    LogUnhandledErrors
    On Error Resume Next
    If VBE.MainWindow.Visible Then
        VBE.MainWindow.WindowState = vbext_ws_Minimize
    End If
    If Err Then Err.Clear
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSafeProjectFileName
' Author    : Adam Waller
' Date      : 6/9/2026
' Purpose   : Return the FileName of a VBProject, or an empty string if it cannot be
'           : read. The VBE.VBProjects collection can contain entries that are not
'           : true VBA projects (registered type libraries or wizards injected by
'           : third-party VBE add-ins such as MZ-Tools, VBWatchdog, or VBExtras).
'           : Reading .FileName on those raises "Requested Type Library or Wizard is
'           : not a VBA Project" (#709). This guards against that so callers can scan
'           : the collection safely.
'           : NOTE: Do NOT short-circuit on Protection here. A locked project (such
'           : as the compiled .accde add-in) still has a readable .FileName and must
'           : remain matchable.
'---------------------------------------------------------------------------------------
'
Public Function GetSafeProjectFileName(oProj As VBProject) As String
    LogUnhandledErrors
    On Error Resume Next
    GetSafeProjectFileName = oProj.FileName
    If Err Then Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    Dim strAddInFile As String
    strAddInFile = GetInstalledAddInFileName
    For Each oProj In VBE.VBProjects
        If StrComp(GetSafeProjectFileName(oProj), strAddInFile, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Public Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    LogUnhandledErrors
    On Error Resume Next
    Application.Run GetInstalledAddInFileName & "!DummyFunction"
    If Err Then Err.Clear
End Sub
