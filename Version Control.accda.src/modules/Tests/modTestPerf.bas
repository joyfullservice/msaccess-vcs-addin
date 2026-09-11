Attribute VB_Name = "modTestPerf"
'---------------------------------------------------------------------------------------
' Module    : modTestPerf
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Micro-benchmarks for the source file scanning path used by merge builds.
'           : These are measurement tools rather than assertions, so every public entry
'           : point takes a parameter. (The test runner only discovers parameterless
'           : Public Sub procedures, so these stay out of the test tree.)
'           :
'           : Run from the Immediate Window in the add-in's own VBE:
'           :   ?modTestPerf.BenchmarkHashPrimitives()
'           :   ?modTestPerf.BenchmarkLegacyIndexBackfill()
'           :
'           : The end-to-end measurement remains the PERFORMANCE REPORTS section of an
'           : actual Export/Merge log. This module isolates the individual primitives so
'           : a regression can be attributed to one of them.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Private Const ModuleName As String = "modTestPerf"

' Report layout
Private Const clngLabelWidth As Long = 46
Private Const clngLineWidth As Long = 78

' Repetition counts for the cheap primitives, so a single measurement is not
' dominated by timer granularity.
Private Const clngStringHashReps As Long = 1000
Private Const clngFileHashReps As Long = 200
Private Const clngComponentHashReps As Long = 200


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkHashPrimitives
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Time each primitive used by merge change detection and return a report.
'           : lngIterations applies to the expensive whole-folder/whole-category
'           : operations; the cheap per-file primitives use their own repeat counts.
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkHashPrimitives(Optional lngIterations As Long = 3) As String

    Dim cOut As clsConcat
    Dim cModules As IDbComponent
    Dim dMeta As Dictionary
    Dim dModuleMeta As Dictionary
    Dim strRoot As String
    Dim strSmallFile As String
    Dim strLargeFile As String
    Dim strTestFile As String
    Dim lngIdx As Long
    Dim dblStart As Double
    Dim strJunk As String

    If lngIterations < 1 Then lngIterations = 1

    strRoot = Options.GetExportFolder
    If Not FSO.FolderExists(strRoot) Then
        BenchmarkHashPrimitives = "Export folder does not exist: " & strRoot
        Exit Function
    End If

    Set cModules = GetCategoryContainer("Modules")
    If cModules Is Nothing Then
        BenchmarkHashPrimitives = "Could not resolve the Modules container."
        Exit Function
    End If

    Set cOut = New clsConcat
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "SCAN PATH BENCHMARK", vbCrLf
    cOut.Add "  Project:   ", CurrentProject.Name, vbCrLf
    cOut.Add "  Source:    ", strRoot, vbCrLf
    ' Parentheses force ByVal so these Variant results coerce to clsConcat.Add's
    ' ByRef String parameters.
    cOut.Add "  Algorithm: ", (Nz2(Options.HashAlgorithm, DefaultHashAlgorithm)), _
        (IIf(Options.UseShortHash, " (short)", vbNullString)), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("Operation", clngLabelWidth), PadLeft("Calls", 8), _
        PadLeft("Seconds", 10), PadLeft("ms/call", 12), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' --- Folder metadata scans -------------------------------------------------------
    ' The recursive root scan is what a merge now performs once for the whole phase.
    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dMeta = ScanFolderMetadata(strRoot)
    Next lngIdx
    AddResult cOut, "ScanFolderMetadata (export root, recursive)", lngIterations, dblStart
    cOut.Add PadRight("  files found", clngLabelWidth), PadLeft(CStr(dMeta.Count), 8), vbCrLf

    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dModuleMeta = ScanFolderMetadata(cModules.BaseFolder)
    Next lngIdx
    AddResult cOut, "ScanFolderMetadata (modules folder)", lngIterations, dblStart

    ' --- Raw hashing primitives ------------------------------------------------------
    dblStart = MicroSeconds
    For lngIdx = 1 To clngStringHashReps
        strJunk = GetStringHash("benchmark content for a short string")
    Next lngIdx
    AddResult cOut, "GetStringHash (short string)", clngStringHashReps, dblStart

    ' Pick the smallest and largest source file in the modules folder
    GetSizeExtremes dModuleMeta, strSmallFile, strLargeFile

    If Len(strSmallFile) Then
        dblStart = MicroSeconds
        For lngIdx = 1 To clngFileHashReps
            strJunk = GetFileHash(strSmallFile)
        Next lngIdx
        AddResult cOut, "GetFileHash (smallest module file)", clngFileHashReps, dblStart
    End If

    If Len(strLargeFile) Then
        dblStart = MicroSeconds
        For lngIdx = 1 To clngFileHashReps
            strJunk = GetFileHash(strLargeFile)
        Next lngIdx
        AddResult cOut, "GetFileHash (largest module file)", clngFileHashReps, dblStart
    End If

    ' --- Per-component hash helpers --------------------------------------------------
    strTestFile = GetFirstSourceFile(cModules)
    If Len(strTestFile) Then

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesPropertyHash(cModules, strTestFile)
        Next lngIdx
        AddResult cOut, "GetSourceFilesPropertyHash (FSO)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesPropertyHash(cModules, strTestFile, dModuleMeta)
        Next lngIdx
        AddResult cOut, "GetSourceFilesPropertyHash (scan map)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesContentHash(cModules, strTestFile)
        Next lngIdx
        AddResult cOut, "GetSourceFilesContentHash (FSO)", clngComponentHashReps, dblStart

        dblStart = MicroSeconds
        For lngIdx = 1 To clngComponentHashReps
            strJunk = GetSourceFilesContentHash(cModules, strTestFile, dModuleMeta)
        Next lngIdx
        AddResult cOut, "GetSourceFilesContentHash (scan map)", clngComponentHashReps, dblStart

    End If

    ' --- End to end change detection per category ------------------------------------
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    AddCategoryBenchmark cOut, "Modules", dMeta, lngIterations
    AddCategoryBenchmark cOut, "Forms", dMeta, lngIterations
    AddCategoryBenchmark cOut, "Queries", dMeta, lngIterations

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    BenchmarkHashPrimitives = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkLegacyIndexBackfill
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : Measure the one-time upgrade-path cost of backfilling AllFilesHash on
'           : legacy multi-file entries, then confirm the second scan returns to the
'           : property-hash short-circuit. Temporarily clears AllFilesHash on every
'           : Forms index entry, times two GetModifiedSourceFiles calls, and restores
'           : the saved hashes. Run from the Immediate Window after rebuilding.
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkLegacyIndexBackfill() As String

    Dim cOut As clsConcat
    Dim cForm As IDbComponent
    Dim dFiles As Dictionary
    Dim dSaved As Dictionary
    Dim dMeta As Dictionary
    Dim dResult As Dictionary
    Dim cIdx As clsVCSIndexItem
    Dim varFile As Variant
    Dim strFile As String
    Dim lngBackfilled As Long
    Dim dblStart As Double
    Dim dblFirst As Double
    Dim dblSecond As Double

    Set cForm = GetCategoryContainer("Forms")
    If cForm Is Nothing Then
        BenchmarkLegacyIndexBackfill = "Could not resolve the Forms container."
        Exit Function
    End If

    Set dFiles = cForm.GetFileList
    Set dSaved = New Dictionary
    Set dMeta = ScanFolderMetadata(Options.GetExportFolder)

    ' Snapshot and clear AllFilesHash so the next scan takes the legacy path
    For Each varFile In dFiles
        strFile = CStr(varFile)
        Set cIdx = VCSIndex.Item(cForm, strFile)
        dSaved.Add strFile, cIdx.AllFilesHash
        cIdx.AllFilesHash = vbNullString
        ' Keep FilePropertiesHash current so the first scan hits the clean fast path
        ' (backfill) rather than the conservative "report modified" branch.
        cIdx.FilePropertiesHash = GetSourceFilesPropertyHash(cForm, strFile, dMeta)
    Next varFile

    Set cOut = New clsConcat
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "LEGACY INDEX BACKFILL BENCHMARK (Forms)", vbCrLf
    cOut.Add "  Source files: ", CStr(dFiles.Count), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    dblStart = MicroSeconds
    Set dResult = VCSIndex.GetModifiedSourceFiles(cForm, dMeta)
    dblFirst = MicroSeconds - dblStart
    AddResult cOut, "First scan (legacy backfill)", 1, dblStart

    For Each varFile In dFiles
        If Len(VCSIndex.Item(cForm, CStr(varFile)).AllFilesHash) > 0 Then
            lngBackfilled = lngBackfilled + 1
        End If
    Next varFile

    dblStart = MicroSeconds
    Set dResult = VCSIndex.GetModifiedSourceFiles(cForm, dMeta)
    dblSecond = MicroSeconds - dblStart
    AddResult cOut, "Second scan (short-circuit)", 1, dblStart

    cOut.Add PadRight("  reported modified (second scan)", clngLabelWidth), _
        PadLeft(CStr(dResult.Count), 8), vbCrLf
    cOut.Add PadRight("  entries with AllFilesHash after first", clngLabelWidth), _
        PadLeft(CStr(lngBackfilled), 8), vbCrLf
    If dblSecond > 0 Then
        cOut.Add PadRight("  first/second ratio", clngLabelWidth), _
            PadLeft(Format$(dblFirst / dblSecond, "0.00") & "x", 8), vbCrLf
    End If
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' Restore original hashes so the live index is not left in a stripped state
    For Each varFile In dSaved.Keys
        VCSIndex.Item(cForm, CStr(varFile)).AllFilesHash = CStr(dSaved(varFile))
    Next varFile

    BenchmarkLegacyIndexBackfill = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkFormGeometry
' Author    : Adam Waller
' Date      : 9/9/2026
' Purpose   : Measure the cost the form geometry canonicalizer adds to a form export,
'           : through the real sanitize path rather than isolated primitives. The
'           : 5.0.0 row runs the identical parser with the canonicalizer gated off, so
'           : the difference between the two sanitize rows is the production cost of
'           : the pass. Phase rows attribute that cost across the four passes.
'           :
'           :   ?modTestPerf.BenchmarkFormGeometry("C:\path\to\Some.form")
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkFormGeometry(strFormPath As String, Optional lngReps As Long = 5, _
    Optional lngRounds As Long = 3) As String

    Dim cOut As clsConcat
    Dim cParser As clsSourceParser
    Dim cCanon As clsFormGeometryCanonicalizer
    Dim strText As String
    Dim strResult As String
    Dim astrCanon() As String
    Dim astrBench() As String
    Dim lngIdx As Long
    Dim lngRound As Long
    Dim dblStart As Double
    Dim dblRound As Double
    Dim dblOn As Double
    Dim dblOff As Double
    Dim dblCanon As Double
    Dim dblSplit As Double
    Dim lngOldFormat As Long
    Dim intOldSanitize As eSanitizeLevel
    Dim eimPrior As eInteractionMode
    Dim dblCopyIn As Double
    Dim dblParse As Double
    Dim dblGroups As Double
    Dim dblTabs As Double
    Dim dblEnvelopes As Double
    Dim dblCopyOut As Double
    Dim strError As String

    If Not FSO.FileExists(strFormPath) Then
        BenchmarkFormGeometry = "File not found: " & strFormPath
        Exit Function
    End If
    If lngReps < 1 Then lngReps = 1
    If lngRounds < 1 Then lngRounds = 1
    dblOn = 1E+30
    dblOff = 1E+30
    dblCanon = 1E+30
    dblSplit = 1E+30

    strText = ReadFile(strFormPath)
    lngOldFormat = Options.ExportFormatVersion
    intOldSanitize = Options.SanitizeLevel
    Options.SanitizeLevel = eslStandard

    ' A form whose geometry cannot be planned logs a warning. With no console attached
    ' that warning becomes a modal dialog, which blocks the run and lands in the
    ' timings as if it were work.
    eimPrior = Operation.InteractionMode
    Operation.InteractionMode = eimSilent

    ' The overrides above outlive this procedure, so a failure below has to fall
    ' through the restore rather than leave the session on a benchmark's settings.
    On Error GoTo CleanUp

    ' Warm up, and capture the canonical output so an optimization can be checked
    ' for byte identity rather than just speed.
    Options.ExportFormatVersion = EFV_5_1_0
    Set cParser = New clsSourceParser
    cParser.LoadString strText, edbForm
    strResult = cParser.Sanitize(ectObjectDefinition)

    Set cOut = New clsConcat
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "FORM GEOMETRY BENCHMARK", vbCrLf
    cOut.Add "  File:  ", FSO.GetFileName(strFormPath), vbCrLf
    cOut.Add "  Lines: ", CStr(UBound(Split(strText, vbCrLf)) + 1), vbCrLf
    cOut.Add "  Reps:  ", CStr(lngReps), " x ", CStr(lngRounds), " rounds (min reported)", vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("Operation", clngLabelWidth), PadLeft("Calls", 8), _
        PadLeft("Seconds", 10), PadLeft("ms/call", 12), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' Rounds interleave the measurements so drift affects them alike; the minimum is
    ' the least noise-sensitive estimator available on a busy desktop.
    For lngRound = 1 To lngRounds

        ' Full sanitize, canonicalizer on
        Options.ExportFormatVersion = EFV_5_1_0
        dblStart = MicroSeconds
        For lngIdx = 1 To lngReps
            Set cParser = New clsSourceParser
            cParser.LoadString strText, edbForm
            cParser.Sanitize ectObjectDefinition
        Next lngIdx
        dblRound = MicroSeconds - dblStart
        If dblRound < dblOn Then dblOn = dblRound

        ' Full sanitize, canonicalizer gated off by export format version
        Options.ExportFormatVersion = EFV_5_0_0
        dblStart = MicroSeconds
        For lngIdx = 1 To lngReps
            Set cParser = New clsSourceParser
            cParser.LoadString strText, edbForm
            cParser.Sanitize ectObjectDefinition
        Next lngIdx
        dblRound = MicroSeconds - dblStart
        If dblRound < dblOff Then dblOff = dblRound
        Options.ExportFormatVersion = EFV_5_1_0

        ' Canonicalize alone, capturing phase attribution from the fastest round
        dblStart = MicroSeconds
        For lngIdx = 1 To lngReps
            astrCanon = Split(strText, vbCrLf)
            Set cCanon = New clsFormGeometryCanonicalizer
            cCanon.Canonicalize astrCanon, "benchmark"
        Next lngIdx
        dblRound = MicroSeconds - dblStart
        If dblRound < dblCanon Then
            dblCanon = dblRound
            dblCopyIn = cCanon.PhaseCopyIn
            dblParse = cCanon.PhaseParse
            dblGroups = cCanon.PhaseGroups
            dblTabs = cCanon.PhaseTabs
            dblEnvelopes = cCanon.PhaseEnvelopes
            dblCopyOut = cCanon.PhaseCopyOut
        End If

        ' Control: the Split alone, with no canonicalization
        dblStart = MicroSeconds
        For lngIdx = 1 To lngReps
            astrBench = Split(strText, vbCrLf)
        Next lngIdx
        dblRound = MicroSeconds - dblStart
        If dblRound < dblSplit Then dblSplit = dblRound

    Next lngRound

    AddPhase cOut, "Sanitize form (5.1.0, canonicalize on)", lngReps, dblOn
    AddPhase cOut, "Sanitize form (5.0.0, canonicalize off)", lngReps, dblOff
    AddPhase cOut, "  => canonicalization overhead", lngReps, dblOn - dblOff
    If dblOff > 0 Then
        cOut.Add PadRight("  => sanitize cost multiple", clngLabelWidth), _
            PadLeft(Format$(dblOn / dblOff, "0.00") & "x", 8), vbCrLf
    End If

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    AddPhase cOut, "Canonicalize only (excludes sanitize)", lngReps, dblCanon
    AddPhase cOut, "  phase: copy array in", 1, dblCopyIn
    AddPhase cOut, "  phase: ParseDocument", 1, dblParse
    AddPhase cOut, "  phase: CanonicalizeGroups", 1, dblGroups
    AddPhase cOut, "  phase: CanonicalizeTabs", 1, dblTabs
    AddPhase cOut, "  phase: GrowEnvelopes", 1, dblEnvelopes
    AddPhase cOut, "  phase: copy array out", 1, dblCopyOut
    AddPhase cOut, "Split control (no canonicalization)", lngReps, dblSplit

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("  blocks parsed", clngLabelWidth), _
        PadLeft(CStr(cCanon.BlockCount), 8), vbCrLf
    cOut.Add PadRight("  groups / PlanAxis calls / rewrites", clngLabelWidth), _
        PadLeft(CStr(cCanon.GroupCount) & " / " & CStr(cCanon.PlanAxisCalls) & _
        " / " & CStr(cCanon.RewriteCalls), 20), vbCrLf
    cOut.Add PadRight("  output length / hash", clngLabelWidth), _
        PadLeft(CStr(Len(strResult)), 8), " ", GetStringHash(strResult), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' Clears Err, so the restore below can tell a failure from a clean finish.
    On Error GoTo 0

CleanUp:
    If Err.Number <> 0 Then
        strError = "Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Options.ExportFormatVersion = lngOldFormat
    Options.SanitizeLevel = intOldSanitize
    Operation.InteractionMode = eimPrior
    If Len(strError) > 0 Then
        BenchmarkFormGeometry = strError
    Else
        BenchmarkFormGeometry = cOut.GetStr
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkFormCorpus
' Author    : Adam Waller
' Date      : 9/9/2026
' Purpose   : Canonicalize every .form file in a folder, reporting total and per-form
'           : cost against a real corpus rather than a single file.
'           :
'           : Already-exported source files are a fixed point of the canonicalizer, so
'           : any file this reports as CHANGED means the canonicalizer no longer agrees
'           : with the output committed to that repository. That makes this an
'           : equivalence check across the whole corpus, not just a timing run.
'           :
'           :   ?modTestPerf.BenchmarkFormCorpus("C:\repo\db.accdb.src\forms")
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkFormCorpus(strFolder As String, Optional lngLimit As Long = 0) As String

    Dim cOut As clsConcat
    Dim cCanon As clsFormGeometryCanonicalizer
    Dim oFile As Object
    Dim strIn As String
    Dim astrLines() As String
    Dim curStart As Currency
    Dim dblOne As Double
    Dim dblTotal As Double
    Dim dblWorst As Double
    Dim strWorst As String
    Dim lngForms As Long
    Dim lngChanged As Long
    Dim lngErrors As Long
    Dim lngLines As Long
    Dim eimPrior As eInteractionMode
    Dim strError As String

    If Not FSO.FolderExists(strFolder) Then
        BenchmarkFormCorpus = "Folder not found: " & strFolder
        Exit Function
    End If

    ' Any form whose geometry cannot be planned logs a warning, and with no console
    ' attached each one becomes a modal dialog that stalls the whole corpus run.
    eimPrior = Operation.InteractionMode
    Operation.InteractionMode = eimSilent

    ' The override above outlives this procedure, so a failure reading or splitting a
    ' form has to fall through the restore rather than leave the session silent.
    On Error GoTo CleanUp

    Set cOut = New clsConcat
    For Each oFile In FSO.GetFolder(strFolder).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), "form", vbTextCompare) = 0 Then
            strIn = ReadFile(oFile.Path)
            astrLines = Split(strIn, vbCrLf)
            lngLines = lngLines + UBound(astrLines) + 1
            Set cCanon = New clsFormGeometryCanonicalizer
            curStart = Perf.MicroTimer
            On Error Resume Next
            Err.Clear
            cCanon.Canonicalize astrLines, oFile.Name
            If Err.Number <> 0 Then
                lngErrors = lngErrors + 1
                If lngErrors <= 15 Then
                    cOut.Add "  ERROR ", CStr(Err.Number), " ", Err.Description, _
                        " in ", oFile.Name, vbCrLf
                End If
                Err.Clear
            End If
            On Error GoTo CleanUp
            dblOne = CDbl(Perf.MicroTimer - curStart)
            dblTotal = dblTotal + dblOne
            lngForms = lngForms + 1
            If dblOne > dblWorst Then
                dblWorst = dblOne
                strWorst = oFile.Name
            End If
            If cCanon.ChangedLines > 0 Then
                lngChanged = lngChanged + 1
                If lngChanged <= 15 Then
                    cOut.Add "  CHANGED: ", oFile.Name, " (", _
                        CStr(cCanon.ChangedLines), " lines)", vbCrLf
                End If
            End If
            ' A long uninterrupted VBA loop starves the message pump, which makes Access
            ' look unresponsive to the automation client driving this run.
            If lngForms Mod 20 = 0 Then DoEvents
            If lngLimit > 0 And lngForms >= lngLimit Then Exit For
        End If
    Next oFile

    ' Clears Err, so the restore below can tell a failure from a clean finish.
    On Error GoTo 0

    If lngForms = 0 Then
        strError = "No .form files found in " & strFolder
        GoTo CleanUp
    End If

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "FORM CORPUS CANONICALIZATION", vbCrLf
    cOut.Add "  Folder:        ", strFolder, vbCrLf
    cOut.Add "  Forms:         ", CStr(lngForms), vbCrLf
    cOut.Add "  Lines:         ", CStr(lngLines), vbCrLf
    cOut.Add "  Changed:       ", CStr(lngChanged), _
        IIf(lngChanged = 0, "  (idempotent: matches committed output)", "  <-- REVIEW"), vbCrLf
    cOut.Add "  Errors:        ", CStr(lngErrors), vbCrLf
    cOut.Add "  Total:         ", Format$(dblTotal, "0.000"), " sec", vbCrLf
    cOut.Add "  Per form:      ", Format$(dblTotal / lngForms * 1000, "0.0"), " ms", vbCrLf
    cOut.Add "  Slowest:       ", strWorst, " (", Format$(dblWorst * 1000, "0.0"), " ms)", vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

CleanUp:
    If Err.Number <> 0 Then
        strError = "Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Operation.InteractionMode = eimPrior
    If Len(strError) > 0 Then
        BenchmarkFormCorpus = strError
    Else
        BenchmarkFormCorpus = cOut.GetStr
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkSanitizeCorpus
' Author    : Adam Waller
' Date      : 9/11/2026
' Purpose   : Run the full Sanitize path over every matching source file in a folder and
'           : compare the result to the file on disk. Committed source is a fixed point
'           : of the sanitizer that produced it, so a DIFFERS count above zero means the
'           : sanitizer no longer agrees with that repository's committed output.
'           :
'           : This covers the whole Case Else chain, not just the canonicalizer that
'           : BenchmarkFormCorpus exercises. A refactor that broke report sanitizing left
'           : that harness at 0 changed while every report in a real project silently
'           : kept lines it should have dropped, so run this one for reports and macros
'           : too. Match the export format and sanitize options to the ones the corpus
'           : was exported with, or every file will differ for uninteresting reasons.
'           :
'           :   ?modTestPerf.BenchmarkSanitizeCorpus("C:\repo\db.accdb.src\reports", "report", edbReport)
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkSanitizeCorpus(strFolder As String, _
    Optional strExtension As String = "form", _
    Optional intObjectType As eDatabaseComponentType = edbForm, _
    Optional lngFormat As Long = EFV_5_1_0, _
    Optional intSanitize As eSanitizeLevel = eslStandard, _
    Optional intColors As eSanitizeLevel = eslMinimal, _
    Optional lngLimit As Long = 0) As String

    Dim cOut As clsConcat
    Dim cParser As clsSourceParser
    Dim oFile As Object
    Dim strIn As String
    Dim strOut As String
    Dim curStart As Currency
    Dim dblOne As Double
    Dim dblTotal As Double
    Dim dblWorst As Double
    Dim strWorst As String
    Dim lngFiles As Long
    Dim lngDiffer As Long
    Dim lngErrors As Long
    Dim lngOldFormat As Long
    Dim intOldSanitize As eSanitizeLevel
    Dim intOldColors As eSanitizeLevel
    Dim eimPrior As eInteractionMode
    Dim strError As String

    If Not FSO.FolderExists(strFolder) Then
        BenchmarkSanitizeCorpus = "Folder not found: " & strFolder
        Exit Function
    End If

    lngOldFormat = Options.ExportFormatVersion
    intOldSanitize = Options.SanitizeLevel
    intOldColors = Options.SanitizeColors
    eimPrior = Operation.InteractionMode
    Options.ExportFormatVersion = lngFormat
    Options.SanitizeLevel = intSanitize
    Options.SanitizeColors = intColors
    Operation.InteractionMode = eimSilent

    ' The overrides above outlive this procedure, so a failure reading or sanitizing a
    ' file has to fall through the restore rather than leave the session silent.
    On Error GoTo CleanUp

    Set cOut = New clsConcat
    For Each oFile In FSO.GetFolder(strFolder).Files
        If StrComp(FSO.GetExtensionName(oFile.Name), strExtension, vbTextCompare) = 0 Then
            strIn = ReadFile(oFile.Path)
            Set cParser = New clsSourceParser
            cParser.LoadString strIn, intObjectType
            ' Matches the export path, which names the object after the file. The
            ' split-VBA marker line embeds this name, so leaving it blank makes every
            ' file with code behind it differ.
            cParser.ObjectName = FSO.GetBaseName(oFile.Name)
            curStart = Perf.MicroTimer
            On Error Resume Next
            Err.Clear
            strOut = cParser.Sanitize(ectObjectDefinition)
            If Err.Number <> 0 Then
                lngErrors = lngErrors + 1
                If lngErrors <= 15 Then
                    cOut.Add "  ERROR ", CStr(Err.Number), " ", Err.Description, _
                        " in ", oFile.Name, vbCrLf
                End If
                Err.Clear
            End If
            On Error GoTo CleanUp
            dblOne = CDbl(Perf.MicroTimer - curStart)
            dblTotal = dblTotal + dblOne
            lngFiles = lngFiles + 1
            If dblOne > dblWorst Then
                dblWorst = dblOne
                strWorst = oFile.Name
            End If
            ' WriteFile terminates the file with a newline that Sanitize does not
            ' return, so compare the content without trailing blank lines.
            If StrComp(TrimTrailingCrLf(strIn), TrimTrailingCrLf(strOut), _
                vbBinaryCompare) <> 0 Then
                lngDiffer = lngDiffer + 1
                If lngDiffer <= 15 Then
                    cOut.Add "  DIFFERS: ", oFile.Name, " (in ", CStr(Len(strIn)), _
                        " / out ", CStr(Len(strOut)), " chars)", vbCrLf
                End If
            End If
            ' A long uninterrupted VBA loop starves the message pump, which makes Access
            ' look unresponsive to the automation client driving this run.
            If lngFiles Mod 20 = 0 Then DoEvents
            If lngLimit > 0 And lngFiles >= lngLimit Then Exit For
        End If
    Next oFile

    ' Clears Err, so the restore below can tell a failure from a clean finish.
    On Error GoTo 0

    If lngFiles = 0 Then
        strError = "No ." & strExtension & " files found in " & strFolder
        GoTo CleanUp
    End If

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "SOURCE CORPUS SANITIZE", vbCrLf
    cOut.Add "  Folder:        ", strFolder, vbCrLf
    cOut.Add "  Files:         ", CStr(lngFiles), " (.", strExtension, ")", vbCrLf
    cOut.Add "  Format:        ", CStr(lngFormat), "  Sanitize: ", CStr(intSanitize), _
        "  Colors: ", CStr(intColors), vbCrLf
    cOut.Add "  Differs:       ", CStr(lngDiffer), _
        IIf(lngDiffer = 0, "  (idempotent: matches committed output)", "  <-- REVIEW"), vbCrLf
    cOut.Add "  Errors:        ", CStr(lngErrors), vbCrLf
    cOut.Add "  Total:         ", Format$(dblTotal, "0.000"), " sec", vbCrLf
    cOut.Add "  Per file:      ", Format$(dblTotal / lngFiles * 1000, "0.0"), " ms", vbCrLf
    cOut.Add "  Slowest:       ", strWorst, " (", Format$(dblWorst * 1000, "0.0"), " ms)", vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

CleanUp:
    If Err.Number <> 0 Then
        strError = "Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Options.ExportFormatVersion = lngOldFormat
    Options.SanitizeLevel = intOldSanitize
    Options.SanitizeColors = intOldColors
    Operation.InteractionMode = eimPrior
    If Len(strError) > 0 Then
        BenchmarkSanitizeCorpus = strError
    Else
        BenchmarkSanitizeCorpus = cOut.GetStr
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : TrimTrailingCrLf
' Author    : Adam Waller
' Date      : 9/11/2026
' Purpose   : Return the text without any trailing CRLF pairs.
'---------------------------------------------------------------------------------------
'
Private Function TrimTrailingCrLf(strText As String) As String

    Dim strResult As String

    strResult = strText
    Do While Right$(strResult, 2) = vbCrLf
        strResult = Left$(strResult, Len(strResult) - 2)
    Loop
    TrimTrailingCrLf = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : BenchmarkSanitize
' Author    : Adam Waller
' Date      : 9/11/2026
' Purpose   : Measure sanitize cost through the real clsSourceParser path at EFV_5_0_0
'           : so the form geometry canonicalizer is gated off. Reports min-of-rounds
'           : timings and output hashes for a form, a module, and a table-definition XML
'           : file so non-form sanitize paths are represented.
'           :
'           :   ?modTestPerf.BenchmarkSanitize()
'           :   ?modTestPerf.BenchmarkSanitize("<form source file>")
'---------------------------------------------------------------------------------------
'
Public Function BenchmarkSanitize(Optional strFormPath As String, _
    Optional strModulePath As String, Optional strXmlPath As String, _
    Optional lngReps As Long = 5, Optional lngRounds As Long = 3) As String

    Dim cOut As clsConcat
    Dim cParser As clsSourceParser
    Dim strText As String
    Dim strResult As String
    Dim lngIdx As Long
    Dim lngRound As Long
    Dim dblStart As Double
    Dim dblRound As Double
    Dim dblForm As Double
    Dim dblModule As Double
    Dim dblXml As Double
    Dim lngOldFormat As Long
    Dim intOldSanitize As eSanitizeLevel
    Dim eimPrior As eInteractionMode
    Dim strFormHash As String
    Dim strModuleHash As String
    Dim strXmlHash As String
    Dim strError As String
    Dim strSourceFolder As String

    strSourceFolder = FSO.BuildPath(CurrentProject.Path, "Version Control.accda.src")
    If Len(strFormPath) = 0 Then
        strFormPath = FSO.BuildPath(FSO.BuildPath(strSourceFolder, "forms"), _
            "frmVCSMain.form")
    End If
    If Len(strModulePath) = 0 Then
        strModulePath = FSO.BuildPath(FSO.BuildPath(strSourceFolder, _
            "modules\Utility"), "modStringUtil.bas")
    End If
    If Len(strXmlPath) = 0 Then
        strXmlPath = FSO.BuildPath(FSO.BuildPath(strSourceFolder, "tbldefs"), _
            "tblConflicts.xml")
    End If
    If lngReps < 1 Then lngReps = 1
    If lngRounds < 1 Then lngRounds = 1
    dblForm = 1E+30
    dblModule = 1E+30
    dblXml = 1E+30

    lngOldFormat = Options.ExportFormatVersion
    intOldSanitize = Options.SanitizeLevel
    Options.SanitizeLevel = eslStandard
    Options.ExportFormatVersion = EFV_5_0_0

    eimPrior = Operation.InteractionMode
    Operation.InteractionMode = eimSilent

    ' The three overrides above outlive this procedure, so a failure below has to fall
    ' through the restore rather than leave the session on a benchmark's settings.
    On Error GoTo CleanUp

    ' Warm up and capture reference hashes at the canonicalizer-off format gate.
    If FSO.FileExists(strFormPath) Then
        strText = ReadFile(strFormPath)
        Set cParser = New clsSourceParser
        cParser.LoadString strText, edbForm
        strResult = cParser.Sanitize(ectObjectDefinition)
        strFormHash = GetStringHash(strResult)
    End If
    If FSO.FileExists(strModulePath) Then
        strText = ReadFile(strModulePath)
        Set cParser = New clsSourceParser
        cParser.LoadString strText, edbModule
        strResult = cParser.Sanitize(ectVBA)
        strModuleHash = GetStringHash(strResult)
    End If
    If FSO.FileExists(strXmlPath) Then
        strText = ReadFile(strXmlPath)
        Set cParser = New clsSourceParser
        cParser.LoadString strText, edbTableDef
        strResult = cParser.Sanitize(ectXML)
        strXmlHash = GetStringHash(strResult)
    End If

    Set cOut = New clsConcat
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add "SANITIZE BENCHMARK (EFV_5_0_0, canonicalizer off)", vbCrLf
    cOut.Add "  Reps:  ", CStr(lngReps), " x ", CStr(lngRounds), " rounds (min reported)", vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("Object", clngLabelWidth), PadLeft("Calls", 8), _
        PadLeft("Seconds", 10), PadLeft("ms/call", 12), vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    For lngRound = 1 To lngRounds
        If FSO.FileExists(strFormPath) Then
            strText = ReadFile(strFormPath)
            dblStart = MicroSeconds
            For lngIdx = 1 To lngReps
                Set cParser = New clsSourceParser
                cParser.LoadString strText, edbForm
                cParser.Sanitize ectObjectDefinition
            Next lngIdx
            dblRound = MicroSeconds - dblStart
            If dblRound < dblForm Then dblForm = dblRound
        End If

        If FSO.FileExists(strModulePath) Then
            strText = ReadFile(strModulePath)
            dblStart = MicroSeconds
            For lngIdx = 1 To lngReps
                Set cParser = New clsSourceParser
                cParser.LoadString strText, edbModule
                cParser.Sanitize ectVBA
            Next lngIdx
            dblRound = MicroSeconds - dblStart
            If dblRound < dblModule Then dblModule = dblRound
        End If

        If FSO.FileExists(strXmlPath) Then
            strText = ReadFile(strXmlPath)
            dblStart = MicroSeconds
            For lngIdx = 1 To lngReps
                Set cParser = New clsSourceParser
                cParser.LoadString strText, edbTableDef
                cParser.Sanitize ectXML
            Next lngIdx
            dblRound = MicroSeconds - dblStart
            If dblRound < dblXml Then dblXml = dblRound
        End If
    Next lngRound

    If FSO.FileExists(strFormPath) Then
        AddPhase cOut, "Form: " & FSO.GetFileName(strFormPath), lngReps, dblForm
    Else
        cOut.Add PadRight("Form: (missing)", clngLabelWidth), vbCrLf
    End If
    If FSO.FileExists(strModulePath) Then
        AddPhase cOut, "Module: " & FSO.GetFileName(strModulePath), lngReps, dblModule
    Else
        cOut.Add PadRight("Module: (missing)", clngLabelWidth), vbCrLf
    End If
    If FSO.FileExists(strXmlPath) Then
        AddPhase cOut, "XML: " & FSO.GetFileName(strXmlPath), lngReps, dblXml
    Else
        cOut.Add PadRight("XML: (missing)", clngLabelWidth), vbCrLf
    End If

    cOut.Add String$(clngLineWidth, "-"), vbCrLf
    cOut.Add PadRight("  form hash", clngLabelWidth), strFormHash, vbCrLf
    cOut.Add PadRight("  module hash", clngLabelWidth), strModuleHash, vbCrLf
    cOut.Add PadRight("  xml hash", clngLabelWidth), strXmlHash, vbCrLf
    cOut.Add String$(clngLineWidth, "-"), vbCrLf

    ' Clears Err, so the restore below can tell a failure from a clean finish.
    On Error GoTo 0

CleanUp:
    If Err.Number <> 0 Then
        strError = "Error " & CStr(Err.Number) & ": " & Err.Description
    End If
    Options.ExportFormatVersion = lngOldFormat
    Options.SanitizeLevel = intOldSanitize
    Operation.InteractionMode = eimPrior
    If Len(strError) > 0 Then
        BenchmarkSanitize = strError
    Else
        BenchmarkSanitize = cOut.GetStr
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddPhase
' Author    : Adam Waller
' Date      : 9/9/2026
' Purpose   : Append a phase row where the elapsed time was accumulated by the caller
'           : rather than measured from a start marker.
'---------------------------------------------------------------------------------------
'
Private Sub AddPhase(cOut As clsConcat, ByVal strLabel As String, _
    ByVal lngCalls As Long, ByVal dblElapsed As Double)

    cOut.Add PadRight(strLabel, clngLabelWidth), _
        PadLeft(CStr(lngCalls), 8), _
        PadLeft(Format$(dblElapsed, "0.000"), 10), _
        PadLeft(Format$((dblElapsed / lngCalls) * 1000, "0.0000"), 12), vbCrLf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddCategoryBenchmark
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Time GetModifiedSourceFiles for one category.
'           :
'           : Reported in three rows, because a container caches GetFileList and
'           : GetAllFromDB on first use and a naive back-to-back comparison therefore
'           : charges the whole cache fill to whichever variant runs first:
'           :   cold        - a fresh container, one call. This is what a merge actually
'           :                 pays: the source file enumeration and the database object
'           :                 scan, plus the per-file change detection loop.
'           :   shared map  - warm, using the caller-supplied folder metadata map.
'           :   own scan    - warm, letting the category scan its own folder. The gap
'           :                 against "shared map" is the per-category folder walk that
'           :                 a merge build now avoids.
'---------------------------------------------------------------------------------------
'
Private Sub AddCategoryBenchmark(cOut As clsConcat, ByVal strCategory As String, _
    dMeta As Dictionary, ByVal lngIterations As Long)

    Dim cCategory As IDbComponent
    Dim lngIdx As Long
    Dim dblStart As Double
    Dim dResult As Dictionary

    ' Cold: a container that has never enumerated source files or database objects
    Set cCategory = GetCategoryContainer(strCategory)
    If cCategory Is Nothing Then Exit Sub

    dblStart = MicroSeconds
    Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory, dMeta)
    AddResult cOut, strCategory & ": change scan (cold, shared map)", 1, dblStart

    ' Warm: the container's file list and object list are now cached, so these rows
    ' isolate the per-file detection loop and the folder metadata source.
    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory, dMeta)
    Next lngIdx
    AddResult cOut, strCategory & ": change scan (warm, shared map)", lngIterations, dblStart

    dblStart = MicroSeconds
    For lngIdx = 1 To lngIterations
        Set dResult = VCSIndex.GetModifiedSourceFiles(cCategory)
    Next lngIdx
    AddResult cOut, strCategory & ": change scan (warm, own scan)", lngIterations, dblStart

    cOut.Add PadRight("  source files / reported modified", clngLabelWidth), _
        PadLeft(CStr(cCategory.GetFileList.Count) & " / " & CStr(dResult.Count), 14), vbCrLf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddResult
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Append one measured row to the report.
'---------------------------------------------------------------------------------------
'
Private Sub AddResult(cOut As clsConcat, ByVal strLabel As String, _
    ByVal lngCalls As Long, ByVal dblStart As Double)

    Dim dblElapsed As Double

    dblElapsed = MicroSeconds - dblStart
    cOut.Add PadRight(strLabel, clngLabelWidth), _
        PadLeft(CStr(lngCalls), 8), _
        PadLeft(Format$(dblElapsed, "0.000"), 10), _
        PadLeft(Format$((dblElapsed / lngCalls) * 1000, "0.0000"), 12), vbCrLf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MicroSeconds
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : High resolution timer value in seconds. (Perf.MicroTimer returns Currency,
'           : which would truncate the arithmetic used to average per-call times.)
'---------------------------------------------------------------------------------------
'
Private Function MicroSeconds() As Double
    MicroSeconds = CDbl(Perf.MicroTimer)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCategoryContainer
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the container class for a category name, or Nothing if this project
'           : does not include that category.
'---------------------------------------------------------------------------------------
'
Private Function GetCategoryContainer(ByVal strCategory As String) As IDbComponent

    Dim cCategory As IDbComponent

    For Each cCategory In GetContainers()
        If StrComp(cCategory.Category, strCategory, vbTextCompare) = 0 Then
            Set GetCategoryContainer = cCategory
            Exit Function
        End If
    Next cCategory

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFirstSourceFile
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the first existing source file path for a category.
'---------------------------------------------------------------------------------------
'
Private Function GetFirstSourceFile(cCategory As IDbComponent) As String

    Dim varFile As Variant

    For Each varFile In cCategory.GetFileList
        If FSO.FileExists(CStr(varFile)) Then
            GetFirstSourceFile = CStr(varFile)
            Exit Function
        End If
    Next varFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSizeExtremes
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Return the smallest and largest file in a ScanFolderMetadata map, so the
'           : file hashing measurement covers both ends of the size range.
'---------------------------------------------------------------------------------------
'
Private Sub GetSizeExtremes(dMeta As Dictionary, ByRef strSmallest As String, ByRef strLargest As String)

    Dim varKey As Variant
    Dim varMeta As Variant
    Dim dblSize As Double
    Dim dblMin As Double
    Dim dblMax As Double

    If dMeta Is Nothing Then Exit Sub

    dblMin = -1
    For Each varKey In dMeta.Keys
        varMeta = dMeta(varKey)
        dblSize = CDbl(varMeta(1))
        If dblMin < 0 Or dblSize < dblMin Then
            dblMin = dblSize
            strSmallest = CStr(varKey)
        End If
        If dblSize > dblMax Then
            dblMax = dblSize
            strLargest = CStr(varKey)
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Left justify text in a fixed width column.
'---------------------------------------------------------------------------------------
'
Private Function PadRight(ByVal strText As String, ByVal lngLen As Long) As String

    Dim strResult As String
    Dim strTrimmed As String

    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - 1)
    Mid$(strResult, 1, Len(strTrimmed)) = strTrimmed
    PadRight = strResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : PadLeft
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Right justify text in a fixed width column.
'---------------------------------------------------------------------------------------
'
Private Function PadLeft(ByVal strText As String, ByVal lngLen As Long) As String

    Dim strResult As String
    Dim strTrimmed As String

    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - 1)
    Mid$(strResult, lngLen - Len(strTrimmed) + 1, Len(strTrimmed)) = strTrimmed
    PadLeft = strResult

End Function
