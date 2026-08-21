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

' Snapshot of the export trees consulted by CleanupDuplicateSourceFiles, so a category
' walks each tree once instead of once per component. Only live between a
' BeginDuplicateScanCache/EndDuplicateScanCache pair; outside one, cleanup falls back to
' scanning the tree directly. Keyed by base folder (no trailing separator).
Private m_blnDupScanCache As Boolean
Private m_dDupScanTrees As Dictionary


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
        RemoveEmptyFolder CStr(varItem)
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

    ' Copy ParamArray into a plain Variant so it can be forwarded to the helpers
    Dim varExts As Variant
    varExts = varExtensions

    If StrComp(StripSlash(strBaseFolder), StripSlash(strCorrectFolder), vbTextCompare) = 0 Then Exit Sub
    If Not FSO.FolderExists(strBaseFolder) Then Exit Sub

    Perf.OperationStart "Cleanup Duplicate Files"
    If m_blnDupScanCache Then
        CleanupFromScanCache GetDupScanTree(StripSlash(strBaseFolder)), _
            strCorrectFolder, strSafeName, varExts
    Else
        ' Recursive scan of the base folder tree
        ScanForDuplicates StripSlash(strBaseFolder), strCorrectFolder, strSafeName, varExts
    End If
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BeginDuplicateScanCache
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Start a session in which CleanupDuplicateSourceFiles answers from a single
'           : snapshot of each export tree rather than rescanning the tree for every
'           : component. Exporting 177 modules otherwise walked the modules folder 177
'           : times. Safe because an export only ever writes into a component's correct
'           : folder, and the correct folder is recomputed live per component, so a
'           : snapshot can miss a duplicate but can never name one wrongly.
'           : Must be paired with EndDuplicateScanCache, which prunes emptied folders.
'---------------------------------------------------------------------------------------
'
Public Sub BeginDuplicateScanCache()
    m_blnDupScanCache = True
    Set m_dDupScanTrees = New Dictionary
    m_dDupScanTrees.CompareMode = TextCompare
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EndDuplicateScanCache
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Close the snapshot session, removing folders left empty. Safe to call when
'           : no session is open, so callers can use it defensively on an error path.
'---------------------------------------------------------------------------------------
'
Public Sub EndDuplicateScanCache()

    Dim varKey As Variant

    m_blnDupScanCache = False
    If m_dDupScanTrees Is Nothing Then Exit Sub

    For Each varKey In m_dDupScanTrees.Keys
        PruneEmptyScanCacheFolders m_dDupScanTrees(varKey)
    Next varKey
    Set m_dDupScanTrees = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDupScanTree
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Return the snapshot for a base folder, indexing the tree on first use.
'           : The snapshot holds file paths grouped by file name, plus a file and
'           : subfolder count per folder so emptiness can be tracked without going back
'           : to disk. "Order" lists folders parents-first, as the indexing queue.
'---------------------------------------------------------------------------------------
'
Private Function GetDupScanTree(strBaseFolder As String) As Dictionary

    Dim dTree As Dictionary

    If m_dDupScanTrees.Exists(strBaseFolder) Then
        Set GetDupScanTree = m_dDupScanTrees(strBaseFolder)
        Exit Function
    End If

    Set dTree = New Dictionary
    dTree.Add "Files", NewTextDictionary
    dTree.Add "FileCount", NewTextDictionary
    dTree.Add "SubCount", NewTextDictionary
    dTree.Add "Order", New Collection
    IndexScanCacheTree strBaseFolder, dTree

    m_dDupScanTrees.Add strBaseFolder, dTree
    Set GetDupScanTree = dTree

End Function


'---------------------------------------------------------------------------------------
' Procedure : IndexScanCacheTree
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Walk a base folder once, recording every file grouped by name and the file
'           : and subfolder counts of every folder. Iterative so the folder list comes
'           : out parents-first, which lets the teardown prune it in reverse.
'---------------------------------------------------------------------------------------
'
Private Sub IndexScanCacheTree(strBaseFolder As String, dTree As Dictionary)

    Dim colOrder As Collection
    Dim colFiles As Collection
    Dim colSubFolders As Collection
    Dim dFiles As Dictionary
    Dim dPaths As Dictionary
    Dim varItem As Variant
    Dim strFolder As String
    Dim strPath As String
    Dim strName As String
    Dim lngPos As Long

    Set dFiles = dTree("Files")
    Set colOrder = dTree("Order")
    colOrder.Add strBaseFolder

    lngPos = 1
    Do While lngPos <= colOrder.Count
        strFolder = CStr(colOrder(lngPos))
        Set colFiles = New Collection
        Set colSubFolders = New Collection
        ScanFolderContents strFolder, colFiles, colSubFolders

        dTree("FileCount")(strFolder) = colFiles.Count
        dTree("SubCount")(strFolder) = colSubFolders.Count

        For Each varItem In colFiles
            strPath = CStr(varItem)
            strName = Mid$(strPath, InStrRev(strPath, PathSep) + 1)
            If Not dFiles.Exists(strName) Then Set dFiles(strName) = NewTextDictionary
            Set dPaths = dFiles(strName)
            If Not dPaths.Exists(strPath) Then dPaths.Add strPath, strFolder
        Next varItem

        For Each varItem In colSubFolders
            colOrder.Add CStr(varItem)
        Next varItem

        lngPos = lngPos + 1
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CleanupFromScanCache
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Delete the misplaced copies of one component using the tree snapshot,
'           : which turns the per-component tree walk into a dictionary lookup.
'---------------------------------------------------------------------------------------
'
Private Sub CleanupFromScanCache(dTree As Dictionary, strCorrectFolder As String, _
    strSafeName As String, varExtensions As Variant)

    Dim dFiles As Dictionary
    Dim dPaths As Dictionary
    Dim dFileCount As Dictionary
    Dim varPath As Variant
    Dim strKey As String
    Dim strPath As String
    Dim strFolder As String
    Dim i As Long

    Set dFiles = dTree("Files")
    Set dFileCount = dTree("FileCount")

    For i = LBound(varExtensions) To UBound(varExtensions)
        strKey = strSafeName & CStr(varExtensions(i))
        If dFiles.Exists(strKey) Then
            Set dPaths = dFiles(strKey)
            For Each varPath In dPaths.Keys
                strPath = CStr(varPath)
                strFolder = CStr(dPaths(strPath))
                If StrComp(AddSlash(strFolder), strCorrectFolder, vbTextCompare) <> 0 Then
                    DeleteFile strPath
                    ' Only count it out when it actually went away, so a locked file
                    ' cannot make its folder look empty at teardown.
                    If Not FSO.FileExists(strPath) Then
                        dPaths.Remove strPath
                        dFileCount(strFolder) = dFileCount(strFolder) - 1
                    End If
                End If
            Next varPath
            If dPaths.Count = 0 Then dFiles.Remove strKey
        End If
    Next i

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PruneEmptyScanCacheFolders
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Remove folders the snapshot shows as empty, walking children before parents
'           : so a folder emptied by its own children going away is still caught. Index 1
'           : is the base folder, which is never removed.
'---------------------------------------------------------------------------------------
'
Private Sub PruneEmptyScanCacheFolders(dTree As Dictionary)

    Dim colOrder As Collection
    Dim dFileCount As Dictionary
    Dim dSubCount As Dictionary
    Dim strFolder As String
    Dim strParent As String
    Dim lngPos As Long

    Set colOrder = dTree("Order")
    Set dFileCount = dTree("FileCount")
    Set dSubCount = dTree("SubCount")

    For lngPos = colOrder.Count To 2 Step -1
        strFolder = CStr(colOrder(lngPos))
        If dFileCount(strFolder) = 0 And dSubCount(strFolder) = 0 Then
            If RemoveEmptyFolder(strFolder) Then
                strParent = Left$(strFolder, InStrRev(strFolder, PathSep) - 1)
                If dSubCount.Exists(strParent) Then dSubCount(strParent) = dSubCount(strParent) - 1
            End If
        End If
    Next lngPos

End Sub


'---------------------------------------------------------------------------------------
' Procedure : NewTextDictionary
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : A dictionary that matches keys case-insensitively, mirroring the way the
'           : file system treats the paths and file names used as keys here.
'---------------------------------------------------------------------------------------
'
Private Function NewTextDictionary() As Dictionary
    Set NewTextDictionary = New Dictionary
    NewTextDictionary.CompareMode = TextCompare
End Function


'---------------------------------------------------------------------------------------
' Procedure : ScanForDuplicates
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Recursively scan a folder and subfolders, deleting any file matching
'           : the target name + extensions that is not in the correct folder.
'           : Returns True when this folder holds nothing once its own cleanup and its
'           : children are done, which is what tells the caller to remove it. Emptiness
'           : is counted from the directory listing this level already read, so a tree
'           : with no duplicates costs nothing beyond the listing itself.
'---------------------------------------------------------------------------------------
'
Private Function ScanForDuplicates(strFolder As String, strCorrectFolder As String, _
    strSafeName As String, varExtensions As Variant) As Boolean

    Dim colFiles As New Collection
    Dim colSubFolders As New Collection
    Dim varItem As Variant
    Dim strPath As String
    Dim strName As String
    Dim strParent As String
    Dim lngFiles As Long
    Dim lngSubFolders As Long
    Dim i As Long

    ScanFolderContents strFolder, colFiles, colSubFolders
    lngFiles = colFiles.Count
    lngSubFolders = colSubFolders.Count

    ' Only check files if this is NOT the correct folder
    strParent = AddSlash(strFolder)
    If StrComp(strParent, strCorrectFolder, vbTextCompare) <> 0 Then
        For Each varItem In colFiles
            strPath = CStr(varItem)
            strName = Mid$(strPath, InStrRev(strPath, PathSep) + 1)
            For i = LBound(varExtensions) To UBound(varExtensions)
                If StrComp(strName, strSafeName & CStr(varExtensions(i)), vbTextCompare) = 0 Then
                    DeleteFile strPath
                    ' Only count it out when it actually went away, so a locked file
                    ' cannot make this folder look empty to the caller.
                    If Not FSO.FileExists(strPath) Then lngFiles = lngFiles - 1
                    Exit For
                End If
            Next i
        Next varItem
    End If

    ' Recurse into subfolders, removing any the recursion reports empty
    For Each varItem In colSubFolders
        If ScanForDuplicates(CStr(varItem), strCorrectFolder, strSafeName, varExtensions) Then
            If RemoveEmptyFolder(CStr(varItem)) Then lngSubFolders = lngSubFolders - 1
        End If
    Next varItem

    ScanForDuplicates = (lngFiles = 0) And (lngSubFolders = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveEmptyFolder
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Delete a folder the caller believes is empty, returning True once it is
'           : gone. Callers derive emptiness from a directory listing they already hold,
'           : so this re-confirms it before the forced delete, which would otherwise
'           : take any contents with it. The confirmation only runs for folders already
'           : believed empty, which on a normal export is none of them.
'---------------------------------------------------------------------------------------
'
Private Function RemoveEmptyFolder(strFolder As String) As Boolean

    Dim objFolder As Scripting.Folder

    LogUnhandledErrors
    On Error Resume Next

    Set objFolder = FSO.GetFolder(strFolder)
    If objFolder Is Nothing Then GoTo CleanUp
    If objFolder.Files.Count > 0 Then GoTo CleanUp
    If objFolder.SubFolders.Count > 0 Then GoTo CleanUp
    Set objFolder = Nothing
    Err.Clear

    FSO.DeleteFolder strFolder, True
    RemoveEmptyFolder = Not CatchAny(eelWarning, "Unable to delete empty folder: " & strFolder, _
        ModuleName & ".RemoveEmptyFolder")

CleanUp:
    If Err Then Err.Clear

End Function


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
' Procedure : ResetCurrentVBProjectState
' Author    : Adam Waller
' Date      : 7/6/2026
' Purpose   : Reset the current database VBA project via the VBE Reset command,
'           : clearing any lingering run-state (module-level/global/Static vars)
'           : without closing the database. This prevents "This action will reset
'           : your project" prompts (and intermittent module-import failures) that
'           : occur when VBComponents are modified while the project holds run-state.
'           : Acts on the active project only; it does NOT reset library/add-in
'           : projects, so the add-in's own singletons remain intact.
'           : Returns True if the Reset control was found and executed without error.
'           :
'           : Set blnTrace when a fault here would take the process down (see
'           : LogCrashTrace). Callers that already run with a warm VBE and an idle target
'           : project have no need for it.
'---------------------------------------------------------------------------------------
'
Public Function ResetCurrentVBProjectState(Optional blnTrace As Boolean) As Boolean

    Const VBE_CMD_RESET_ID As Long = 228   ' VBE Standard toolbar Reset (language-independent)
    Dim ctl As CommandBarControl

    LogUnhandledErrors
    On Error Resume Next
    If blnTrace Then LogCrashTrace "reset: setting active project"
    Set VBE.ActiveVBProject = CurrentVBProject
    If blnTrace Then LogCrashTrace "reset: finding Reset control"
    Set ctl = Application.VBE.CommandBars.FindControl(, VBE_CMD_RESET_ID)
    If Not ctl Is Nothing Then
        If blnTrace Then LogCrashTrace "reset: executing Reset control"
        ctl.Execute
        If Err.Number = 0 Then ResetCurrentVBProjectState = True
    End If
    If blnTrace Then LogCrashTrace "reset: returned " & CStr(ResetCurrentVBProjectState)
    If Err Then Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveCurrentVBProject
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Save the current database's VBA project, and report whether the project ended
'           : up clean. Delegates to the worker script, because the VBE Save command only
'           : works when nothing sits below it on the VBA stack (see below).
'           :
'           : This does NOT reset the project first, which was tried and reverted -- see the
'           : list of dropped mechanisms below.
'           :
'           : `DoCmd.Save acModule, <one module>` is not equivalent, despite the long-held
'           : assumption that saving one module saves the whole project. When form and
'           : report class modules are dirty — the usual state after startup code has run,
'           : where dozens of form classes report unsaved — it leaves the project dirty.
'           : modLetterCasing had been logging exactly that ("VBA project still has unsaved
'           : changes after letter casing corrections") for a long time before the cause
'           : was understood.
'           :
'           : It also matters for database locking: a partial save leaves the database
'           : inaccessible to other clients, where a complete one does not, which is the
'           : difference between an in-place merge and one that has to reopen.
'           :
'           : Three mechanisms were tried and dropped, and should not be reintroduced
'           : without new evidence:
'           :
'           :  * Saving individual modules. Form and report class modules cannot be saved
'           :    this way at all, so it reports success while leaving dirty precisely the
'           :    components that matter. This was the original bug.
'           :  * `DoCmd.RunCommand acCmdSaveAllModules`. Raises 2046 ("isn't available now")
'           :    unless a module window is active, and is widely reported to do nothing even
'           :    when it does run. An expected error also has to be captured and cleared
'           :    before anything is logged, or LogUnhandledErrors reports it and a modal
'           :    dialog stops an unattended merge.
'           :  * Executing the VBE Save command (ID 3) in process. Reports success and saves
'           :    nothing — no error, correct project active, caption confirming the right
'           :    document, before and after a project reset alike. Running the identical
'           :    command from the worker saves the project, so the caller's own VBA stack is
'           :    what it objects to. Do not add this back as a "free" first attempt: it never
'           :    succeeds, and targeting it correctly means showing a code pane, which pops
'           :    the VBE window open mid-merge.
'           :  * Resetting the project in the same worker job, immediately before the save,
'           :    to avoid the VBE's "this action will reset your project" prompt on a project
'           :    holding run-state. This broke export and cannot work from here. A VBE reset
'           :    ends whatever code is *running*; setting ActiveVBProject does not scope it
'           :    away from us. During an export the running code is this add-in, waiting in
'           :    Worker.WaitForQueue's DoEvents loop for the very job issuing the reset — so
'           :    it terminated its own caller, taking the job queue with it ("Returned worker
'           :    not found in job queue", then 40040 from the ribbon command). The merge is
'           :    immune only because its next stage arrives on a Windows timer, so nothing of
'           :    ours has to survive; a save called mid-procedure has no such re-entry. Note
'           :    also that the prompt this was meant to prevent has never been observed here.
'           :
'           : Compiling is deliberately avoided: `acCmdCompileAndSaveAllModules` is the
'           : mechanism usually recommended, but a project that does not compile still has
'           : to be mergeable.
'           :
'           : Returns the project's actual `Saved` state, so a caller never has to trust
'           : that this worked.
'---------------------------------------------------------------------------------------
'
Public Function SaveCurrentVBProject(Optional blnTrace As Boolean) As Boolean

    Dim strDetail As String

    LogUnhandledErrors
    On Error Resume Next

    If CurrentVBProject.Saved Then GoTo Verify

    If blnTrace Then LogCrashTrace "save: saving VBA project out of process"
    Worker.Run_SaveVbaProject
    strDetail = ErrDetail
    If Err Then Err.Clear
    If blnTrace Then LogCrashTrace "save: worker returned" & strDetail & _
        ", saved: " & CStr(CurrentVBProject.Saved)

Verify:
    ' Report what actually happened rather than what was attempted.
    SaveCurrentVBProject = CurrentVBProject.Saved
    If Err Then Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : ErrDetail
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Render the current error for a trace line, or an empty string when there is
'           : none, so that a step which fails silently can be told apart from one that
'           : raises. Reads Err without clearing it, so callers must capture the text and
'           : clear Err before logging -- see the note in SaveCurrentVBProject.
'---------------------------------------------------------------------------------------
'
Private Function ErrDetail() As String
    If Err.Number <> 0 Then ErrDetail = " [err " & Err.Number & ": " & Err.Description & "]"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetWouldEndOurOwnCode
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Returns True when the project that ResetCurrentVBProjectState would reset
'           : is also the project running this code — i.e. the add-in is open as the
'           : current database rather than loaded as a library add-in.
'           :
'           : A VBE reset is equivalent to the End statement for the project it acts on.
'           : Resetting a project that has frames on the call stack destroys the stack
'           : underneath the running code, which crashes Access inside VBE7.DLL
'           : (access violation, no trappable error, no chance to fall back). This is
'           : why the reset is safe in RunVBA, where the reset target is the current
'           : database and the caller lives in the add-in library project.
'           :
'           : Callers that reset as a side effect must check this first and choose a
'           : different strategy when it returns True.
'           :
'           : False is NOT a guarantee that a reset is harmless. Setting ActiveVBProject
'           : does not confine a reset to that project: it ends running code generally, so
'           : an add-in-side caller can lose its own module-level state while resetting a
'           : different project. Resetting during an export demonstrated this by wiping the
'           : worker job queue mid-operation. What makes the merge safe is not this check
'           : alone but that its next stage re-enters on a Windows timer, so no state has to
'           : survive the reset. A caller that needs to keep running afterwards needs the
'           : same choreography, not just a False here.
'---------------------------------------------------------------------------------------
'
Public Function ResetWouldEndOurOwnCode() As Boolean
    ResetWouldEndOurOwnCode = (StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0)
End Function


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
