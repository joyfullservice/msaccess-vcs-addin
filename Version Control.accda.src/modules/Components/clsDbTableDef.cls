VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTableDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Components")

' Saved queries in the database before the DAO table-def fast path is worth using.
' Application.ImportXML costs roughly the total number of table references across the
' saved-query catalog, so query count is the cheapest available proxy. Measured cost of
' one acStructureOnly call against a growing catalog:
'   4 queries                       0.02 sec
'   2,000 queries, no table refs    2.00 sec
'   2,000 queries, table refs      19.19 sec
'   4,000 queries, table refs      43.13 sec
'   3,692 real production queries    276-281 sec
' Below this line ImportXML is cheap and is the better-proven path; above it, every
' table avoided is seconds to minutes. (See DECISIONS.md 2026-07-30.)
Private Const MIN_QUERIES_FOR_FAST_TABLEDEF As Long = 500

' Consecutive expensive failures before the fast path stands down for the rest of the
' operation. Source that predates the current export format could fail verification on
' every table, and a merge touching many of them should not pay for the attempt twice.
Private Const MAX_FAST_PATH_FAILURES As Long = 3

Private m_Table As AccessObject
Private m_lngTableType As Long
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

' Fast-path eligibility, resolved once and cached. 0 = not yet checked, 1 = yes, 2 = no.
' These live on the instance because one instance serves one operation: modContainers
' builds a fresh set per build, and LoadSingleObject is handed a fresh one per call.
Private m_intFastPathState As Integer
Private m_lngFastPathFailures As Long

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strFile As String
    Dim dbs As Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim dItem As Dictionary
    Dim rst As DAO.Recordset
    Dim bteLvProp() As Byte
    Dim parser As clsLvPropParser
    Dim dProps As Dictionary
    Dim dTblProps As Dictionary
    Dim strHash As String
    Dim strContent As String
    Dim strConnect As String
    Dim strEnvKey As String

    ' For linked tables, check if the back-end is already known to be unavailable
    strConnect = SharedDb.TableDefs(m_Table.Name).Connect
    If Len(strConnect) > 0 Then
        If IsBackEndUnavailable(strConnect) Then Exit Sub
    End If

    ' Make sure the table exists, and is accessible
    If Not TableExists(m_Table.Name) Then
        If Len(strConnect) > 0 Then
            ' Linked table failed - test whether the back-end itself is reachable
            If TestBackEndConnection(strConnect) Then
                Log.Error eelError, T("Unable to access table {0} on the back-end database.", _
                    var0:=m_Table.Name), ModuleName(Me) & ".Export"
            Else
                MarkBackEndUnavailable strConnect, m_Table.Name
            End If
        Else
            Log.Error eelError, T("Unable to access table {0}." & _
                " Please ensure that this table exists and is accessible in the current database.", _
                var0:=m_Table.Name), ModuleName(Me) & ".Export"
        End If
        Exit Sub
    End If

    ' Get the export file name
    strFile = Nz2(strAlternatePath, IDbComponent_SourceFile)

    ' For internal tables, we can export them as XML.
    If Not IsLinkedTable Then

        ' Save structure in XML format
        VerifyPath strFile
        Perf.OperationStart "App.ExportXML()"
        ' Note that the additional properties are important to accurately reconstruct the table.
        Application.ExportXML acExportTable, m_Table.Name, , strFile, , , , acExportAllTableAndFieldProperties
        Perf.OperationEnd

        ' Rewrite sanitized XML as formatted UTF-8 content
        With New clsSourceParser
            .LoadSourceFile strFile, IDbComponent_ComponentType
            DeleteFile strFile
            WriteFile .Sanitize(ectXML), strFile
            strHash = .Hash
        End With

    Else
        ' Linked table - Save as JSON
        Set dbs = SharedDb
        Set tdf = dbs.TableDefs(m_Table.Name)
        Set dItem = New Dictionary
        With dItem
            .Add "Name", m_Table.Name
            strConnect = SanitizeConnectionString(tdf.Connect)
            strEnvKey = ShouldUseEnvForConnection(tdf.Connect)
            If Len(strEnvKey) Then
                SaveConnectionToEnv strEnvKey, strConnect, m_Table.Name & " (linked table)"
                strConnect = "env:" & strEnvKey
            Else
                ' SECURITY: a connection kept in source (not externalized to .env) must
                ' still never carry a raw password (v5.0.0+). GetSourceSafeConnect is the
                ' single chokepoint that enforces this - do NOT bypass it when refactoring
                ' this branch, or plaintext credentials can leak into committed source.
                ' Regression coverage: modTestConnect.TestGetSourceSafeConnectGating /
                ' TestStripConnectionCredentials. (No end-to-end export-to-file test exists;
                ' that risk is accepted - see DECISIONS.md 2026-06-19.)
                strConnect = GetSourceSafeConnect(strConnect, m_Table.Name & " (linked table)")
            End If
            .Add "Connect", strConnect
            .Add "SourceTableName", tdf.SourceTableName
            .Add "Attributes", tdf.Attributes
            ' indexes (Find primary key)
            If TableIndexesAvailable(tdf) Then
                For Each idx In tdf.Indexes
                    If idx.Primary Then
                        .Add "PrimaryKey", BracketFieldList(CStr(idx.Fields))
                        Exit For
                    End If
                Next idx
            End If
        End With

        ' Export front-end display properties by reading the LvProp binary blob
        ' from MSysObjects. This is sub-millisecond, works offline, and captures
        ' all locally-stored properties (column widths, lookup combos, formatting,
        ' datasheet colors, custom properties, etc.)
        If Options.ExportFormatVersion >= EFV_5_0_0 Then
            Perf.OperationStart "Export Linked Props"
            Set rst = dbs.OpenRecordset( _
                "SELECT LvProp FROM MSysObjects WHERE Name=""" & tdf.Name & """", _
                dbOpenSnapshot)
            If Not rst.EOF Then
                If Not IsNull(rst!LvProp) Then
                    bteLvProp = rst!LvProp
                    Set parser = New clsLvPropParser
                    Set dProps = parser.ParseLvProp(bteLvProp, Options.SanitizeLevel)
                    If dProps.Exists("FieldProperties") Then
                        If dProps("FieldProperties").Count > 0 Then
                            dItem.Add "FieldProperties", dProps("FieldProperties")
                        End If
                    End If
                    If dProps.Exists("TableProperties") Then
                        Set dTblProps = dProps("TableProperties")
                        ' Keep engine-managed version stamps (FCMin*) out of source. They
                        ' are build/version-specific and churn across machines. (Import
                        ' already tolerates them via SetDAOProperty; this just keeps the
                        ' exported .json clean.)
                        Set dTblProps = FilterEngineManagedProps(dTblProps)
                        If dTblProps.Count > 0 Then
                            dItem.Add "TableProperties", SortDictionaryByKeys(dTblProps)
                        End If
                    End If
                End If
            End If
            rst.Close
            Perf.OperationEnd
        End If

        ' Collect object metadata (doc properties + hidden attribute) directly into
        ' the dictionary before serializing. This avoids the write-read-merge-rewrite
        ' cycle that ExportObjectMetadata would otherwise perform for linked tables.
        CollectObjectMetadata dItem, "Tables", m_Table.Name, acTable

        ' Write export file.
        strContent = BuildJsonFile(TypeName(Me), dItem, "Linked Table")
        strHash = GetStringHash(strContent, True)
        WriteFile strContent, strFile

    End If

    ' Export companion files (metadata .json, optional .sql).
    ' These are produced for both normal and alternate-path (temp) exports
    ' so that conflict detection can compare all files and the conflict
    ' dialog can offer per-file diffs.
    Dim strExportFolder As String
    strExportFolder = FSO.GetParentFolderName(strFile) & PathSep

    ' Export object metadata for local tables into a companion .json file.
    ' Linked tables already have metadata included in their primary .json above.
    If Not IsLinkedTable Then
        ExportObjectMetadata strExportFolder & GetSafeFileName(m_Table.Name) & ".json", _
            "Tables", m_Table.Name, Me
    End If

    ' Remove any alternate source file in case we have switched formats
    If strAlternatePath = vbNullString Then
        RemoveAlternateFormatSourceFile
    End If

    ' Optionally save in SQL format
    Dim strSqlFile As String
    strSqlFile = strExportFolder & GetSafeFileName(m_Table.Name) & ".sql"
    If Options.SaveTableSQL And ((Not IsLinkedTable) Or Options.SaveLinkedTableDef) Then
        Log.Add "  " & m_Table.Name & " (SQL)", Options.ShowDebug
        SaveTableSqlDef m_Table.Name, strExportFolder
    ElseIf strAlternatePath = vbNullString Then
        ' Remove stale SQL files only during real exports, not temp exports
        If FSO.FileExists(strSqlFile) Then DeleteFile strSqlFile
    End If

    ' Update index
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), strHash, _
        strMetaHash:=GetMetadataHash("Tables", m_Table.Name, acTable)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim blnUseTemp As Boolean
    Dim blnImportFailed As Boolean
    Dim strTempFile As String
    Dim strName As String
    Dim strFileName As String
    Dim colBigIntFields As Collection

    ' Full build reuses one instance; reset link-type cache per file.
    m_lngTableType = 0
    strName = GetObjectNameFromFileName(strFile)
    strFileName = FSO.GetFileName(strFile)

    If FileHasGitConflictMarkers(strFile, ModuleName(Me) & ".Import") Then Exit Sub

    ' Determine import type from extension
    Select Case LCase$(FSO.GetExtensionName(strFile))

        Case "json"
            If Not ImportLinkedTable(strFile) Then Exit Sub

        Case "xml"
            ' Build the table directly through DAO where that is worth doing, and only
            ' keep the result if it re-exports to exactly the source file. Everything
            ' below is the fallback, unchanged.
            If TryFastTableDefImport(strFile, strName) Then
                ApplyMetadataAndUpdateIndex strName, strFile
                Exit Sub
            End If

            ' Watch out for errors when importing XML
            LogUnhandledErrors
            On Error Resume Next

            ' Application.ImportXML has no schema representation for dbBigInt and
            ' silently re-creates such fields as dbDecimal(38,0) instead. Detect them
            ' from the source XML before the (possibly temp-copied) file is gone.
            Set colBigIntFields = GetBigIntRepairFieldNamesFromTableDefXmlFile(strFile)

            ' The ImportXML function does not properly handle UrlEncoded paths
            blnUseTemp = (InStr(1, strFile, "%") > 0)
            Perf.OperationStart "App.ImportXML() Structure"
            If blnUseTemp Then
                ' Import from (safe) temporary file name.
                strTempFile = GetTempFile
                FSO.CopyFile strFile, strTempFile
                Application.ImportXML strTempFile, acStructureOnly
                DeleteFile strTempFile
            Else
                Application.ImportXML strFile, acStructureOnly
            End If
            Perf.OperationEnd

            ' Importing XML can throw an error if ActiveX security is too restrictive.
            If Catch(31521) Then
                Log.Error eelError, "ActiveX security settings blocked the import of " & _
                "table definition XML file: " & strFileName & ".", ModuleName(Me) & ".Import"
                Log.Add "Importing XML files requires ActiveX to be enabled."
                blnImportFailed = True
            ElseIf Catch(31550) Then
                ' Access can report failed data rows for a structure-only import even
                ' though this XML contains no data. Accept that warning only when the
                ' requested table was actually created.
                blnImportFailed = Not ObjectExists(acTable, strName)
                If blnImportFailed Then
                    Log.Error eelError, "Unable to import XML for " & strFileName, _
                        ModuleName(Me) & ".Import"
                End If
            Else
                ' Keep handler arguments side-effect free while Err is pending.
                blnImportFailed = CatchAny(eelError, "Unable to import XML for " & _
                    strFileName, ModuleName(Me) & ".Import")
            End If

            ' Application.ImportXML creates the table outside SharedDb's TableDefs
            ' snapshot. Invalidate before FixCorruptedBigIntFields / metadata apply /
            ' single-object callers that have no per-category ReleaseDbReferences.
            If Not blnImportFailed Then ReleaseDbReferences

            If Not blnImportFailed And colBigIntFields.Count > 0 Then
                FixCorruptedBigIntFields strName, colBigIntFields
            End If

        Case Else
            ' Unsupported file
            Exit Sub

    End Select

    ' Apply metadata and update index
    ApplyMetadataAndUpdateIndex strName, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TryFastTableDefImport
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Create the table through DAO instead of Application.ImportXML, returning
'           : true only when the result is provably identical to the source file.
'           :
'           : The verification is the whole point. A DAO build reproduces what we know
'           : how to read, and the strict parser in modTableDefBuilder is meant to
'           : decline anything else -- but "meant to" is not a guarantee, so the finished
'           : table is exported and hash-compared before we accept it. On any doubt the
'           : table is dropped and the caller falls through to the proven path, which
'           : means the worst case is the old timing plus a few cheap milliseconds.
'---------------------------------------------------------------------------------------
'
Private Function TryFastTableDefImport(strFile As String, strName As String) As Boolean

    Dim strBuiltFile As String

    If Not UseFastTableDefImport Then Exit Function

    LogUnhandledErrors
    On Error Resume Next

    If TryBuildTableFromDefXml(strFile, strName) Then
        If StoredDefinitionMatchesSource(strName, strFile, "Verify Table Def (DAO)", strBuiltFile) Then
            m_lngFastPathFailures = 0
            TryFastTableDefImport = True
            Log.Add T("    Built {0} directly, without importing the XML.", var0:=strName), _
                Options.ShowDebug
        Else
            LogDefinitionMismatch strBuiltFile, strFile
            Set m_Table = Nothing
            m_lngTableType = 0
            DeleteObjectIfExists acTable, strName
            RecordVerificationFailure T("the rebuilt table did not match the source file")
            RecordFastPathFailure strName, GetLastDeclineReason
        End If
    ElseIf GetLastAttemptReachedBuild Then
        RecordFastPathFailure strName, GetLastDeclineReason
    Else
        ' Declined during parsing, which costs about a millisecond. Not worth counting
        ' against the breaker -- an unsupported construct in one table says nothing
        ' about the next one.
        Log.Add T("    Using XML import for {0}: {1}.", var0:=strName, _
            var1:=GetLastDeclineReason), Options.ShowDebug
    End If

    ' Leave no residual error for the fallback path's Catch/CatchAny checks to find.
    Err.Clear
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : LogDefinitionMismatch
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Record which lines of the rebuilt table's export differ from the source
'           : file, so a fallback explains itself.
'           :
'           : Without this the log says only that the rebuilt table "did not match the
'           : source file", which is a dead end: the temp export is swept at the end of
'           : the operation, so by the time anyone reads the log the evidence is gone and
'           : the only way to learn more is to reproduce -- on the very databases where
'           : reproducing costs minutes per attempt.
'           :
'           : Written to the log file but not echoed to the console; this is diagnostic
'           : detail for after the fact, and the fallback itself is already reported.
'           : Both files come from the same exporter, so the lines correspond one to one
'           : and a positional comparison is enough -- no need to diff properly.
'---------------------------------------------------------------------------------------
'
Private Sub LogDefinitionMismatch(strBuiltFile As String, strSourceFile As String)

    ' Enough to characterize a mismatch without flooding the log on a wholesale
    ' difference (a stale export format, say, where nearly every line differs).
    Const MAX_REPORTED_LINES As Long = 8

    Dim varBuilt As Variant
    Dim varSource As Variant
    Dim lngIndex As Long
    Dim lngLast As Long
    Dim lngShown As Long

    If Not FSO.FileExists(strBuiltFile) Then Exit Sub
    If Not FSO.FileExists(strSourceFile) Then Exit Sub

    varBuilt = Split(ReadFile(strBuiltFile), vbCrLf)
    varSource = Split(ReadFile(strSourceFile), vbCrLf)

    lngLast = UBound(varBuilt)
    If UBound(varSource) > lngLast Then lngLast = UBound(varSource)

    For lngIndex = 0 To lngLast
        If GetLineAt(varSource, lngIndex) <> GetLineAt(varBuilt, lngIndex) Then
            If lngShown = 0 Then Log.Add "      Differences (source vs. rebuilt):", False
            Log.Add "        " & lngIndex + 1 & " source: " & Trim$(GetLineAt(varSource, lngIndex)), False
            Log.Add "        " & lngIndex + 1 & " built : " & Trim$(GetLineAt(varBuilt, lngIndex)), False
            lngShown = lngShown + 1
            If lngShown >= MAX_REPORTED_LINES Then
                Log.Add "        ... additional differences not shown.", False
                Exit For
            End If
        End If
    Next lngIndex

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLineAt
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return one line from a Split() result, or an empty string past the end,
'           : so two files of different lengths can be walked with one index.
'---------------------------------------------------------------------------------------
'
Private Function GetLineAt(varLines As Variant, lngIndex As Long) As String
    If lngIndex <= UBound(varLines) Then GetLineAt = varLines(lngIndex)
End Function


'---------------------------------------------------------------------------------------
' Procedure : UseFastTableDefImport
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Returns true when the DAO table-def path should be attempted for this
'           : operation. Resolved once and cached, since a merge may ask per table.
'---------------------------------------------------------------------------------------
'
Private Function UseFastTableDefImport() As Boolean

    If m_lngFastPathFailures >= MAX_FAST_PATH_FAILURES Then Exit Function
    If FastPathTestOverride Then
        UseFastTableDefImport = True
        Exit Function
    End If

    If m_intFastPathState = 0 Then
        If FastTableDefImportApplies Then
            m_intFastPathState = 1
        Else
            m_intFastPathState = 2
        End If
    End If

    UseFastTableDefImport = (m_intFastPathState = 1)

End Function


'---------------------------------------------------------------------------------------
' Procedure : FastTableDefImportApplies
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Decide whether this operation is one where Application.ImportXML is slow
'           : enough to be worth replacing.
'---------------------------------------------------------------------------------------
'
Private Function FastTableDefImportApplies() As Boolean

    ' A DAO-built table and an ImportXML-built one export the same properties in a
    ' different order, and only the canonical ordering introduced in EFV_5_1_0 makes the
    ' two agree. On an older format the table would import correctly but re-export
    ' differently from the file it came from, rewriting tbldefs on the next export. Not
    ' worth trading source churn for speed, so the path waits for the format.
    If Options.ExportFormatVersion < EFV_5_1_0 Then Exit Function

    ' A full build creates every table before the first query exists (GetContainers
    ' orders tables at #14 and queries at #15), so the pathology never appears there --
    ' measured at 21.80 seconds for 372 tables. Leave that path on the proven call.
    If Operation.OperationType = eotBuild Then Exit Function

    ' Everything else -- merge build, ImportByType, single-object import -- runs against
    ' a populated catalog, where the cost scales with it.
    FastTableDefImportApplies = (CurrentData.AllQueries.Count >= MIN_QUERIES_FOR_FAST_TABLEDEF)

End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordFastPathFailure
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Count a fast-path attempt that cost real time, and stand the path down for
'           : the rest of the operation once they stop looking like one-offs.
'---------------------------------------------------------------------------------------
'
' ByVal on strReason: callers pass the result of T(), which is a Variant.
Private Sub RecordFastPathFailure(ByVal strName As String, ByVal strReason As String)

    m_lngFastPathFailures = m_lngFastPathFailures + 1

    Log.Add T("    Using XML import for {0}: {1}.", var0:=strName, var1:=strReason), _
        Options.ShowDebug

    If m_lngFastPathFailures = MAX_FAST_PATH_FAILURES Then
        Log.Add T("Direct table creation failed {0} times. Using XML import from here on.", _
            var0:=CStr(MAX_FAST_PATH_FAILURES))
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyMetadataAndUpdateIndex
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Apply the object metadata (doc properties + hidden attribute) from the
'           : companion .json file and record the import in the index. For linked
'           : tables, the metadata is in the primary .json file itself. For local
'           : tables (.xml import), this is a companion .json.
'           : Shared by Import and by the Merge path that leaves the existing table in
'           : place, since both need the table to end up in the same recorded state.
'---------------------------------------------------------------------------------------
'
Private Sub ApplyMetadataAndUpdateIndex(strName As String, strFile As String)

    Dim strJsonFile As String

    m_lngTableType = 0
    Set m_Table = CurrentData.AllTables(strName)

    strJsonFile = IDbComponent_BaseFolder & GetSafeFileName(strName) & ".json"
    ImportObjectMetadata strJsonFile, "Tables", strName, acTable

    VCSIndex.Update Me, eatImport, GetFileHash(strFile), _
        strMetaHash:=GetMetadataHash("Tables", strName, acTable)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'           :
'           : Replacing the table means Application.ImportXML, which dominates the cost
'           : of this path once a database holds a lot of objects. Measured on a
'           : three-field lookup table in a project with roughly 5,000 objects:
'           : 281 seconds for the ImportXML call, against under 0.01 seconds each for
'           : DoCmd.DeleteObject, creating the same table through DDL,
'           : Application.ExportXML, and appending the rows. The cost tracks the object
'           : count rather than the table, so it cannot be tuned from this side --
'           : the only remedy is not to make the call. A table whose stored definition
'           : already matches its source file has nothing to rebuild, so that case is
'           : short-circuited before any of the destructive work begins.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)

    Dim strName As String
    Dim dRelations As Dictionary

    strName = GetObjectNameFromFileName(strFile)

    ' Nothing to rebuild when the table already matches the source file.
    If StoredDefinitionMatchesSource(strName, strFile) Then
        Log.Add T("    {0} already matches the source file. Skipping table rebuild.", _
            var0:=strName), Options.ShowDebug
        ApplyMetadataAndUpdateIndex strName, strFile
        Exit Sub
    End If

    ' Stage any affected relationships
    Perf.OperationStart "Stage Relations"
    Set dRelations = StageRelations(strName)
    Perf.OperationEnd

    ' Delete any existing table with this name
    Set m_Table = Nothing
    m_lngTableType = 0
    Perf.OperationStart "Delete Table Object"
    DeleteObjectIfExists acTable, strName
    Perf.OperationEnd

    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
        ' Restore any affected relationships
        Perf.OperationStart "Restore Relations"
        RestoreRelations dRelations
        Perf.OperationEnd
    Else
        ' If the table is deleted, discard any staged relationships
        ' since we can't restore them when the table does not exist.

        ' Also remove table def from index.
        VCSIndex.Remove Me, strFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StoredDefinitionMatchesSource
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Returns true when the table currently in the database exports to exactly
'           : the content of the source file, meaning a merge would rebuild it into
'           : what it already is.
'           :
'           : Deliberately limited to local tables (.xml source). Those are the ones
'           : rebuilt through Application.ImportXML, which is the expensive call this
'           : check exists to avoid. Linked tables import from .json without it, and
'           : exporting one can reach across to the back end, so checking them could
'           : cost more than the import it would skip.
'           :
'           : A false negative is harmless -- the caller just rebuilds as before -- so
'           : anything unexpected (missing table, export that produces no file) returns
'           : false. Source files outside the export folder are still verifiable: the
'           : check falls back to a flat name in the temp folder rather than giving up.
'           :
'           : Also used to verify a table built by the DAO fast path, which is the same
'           : question asked at a different moment. strTimerLabel keeps the two apart in
'           : the performance report.
'---------------------------------------------------------------------------------------
'
Private Function StoredDefinitionMatchesSource(strName As String, strFile As String, _
    Optional strTimerLabel As String = "Compare Table Definition", _
    Optional ByRef strComparedFile As String) As Boolean

    Dim strTempFile As String

    ' Only local table definitions take the Application.ImportXML path.
    If LCase$(FSO.GetExtensionName(strFile)) <> "xml" Then Exit Function
    If Not FSO.FileExists(strFile) Then Exit Function
    If Not TableExists(strName) Then Exit Function

    ' Prefer a mirrored path under the conflict-detection temp folder so the export
    ' lands beside the other comparison copies and is cleaned up with them. When the
    ' source file is not under the export folder (single-object import from an
    ' arbitrary path, or a test fixture) there is no prefix to rewrite -- fall back
    ' to a flat name in the same temp folder. The last check is the real invariant:
    ' never export onto the source file being compared against.
    strTempFile = Replace(strFile, Options.GetExportFolder, VCSIndex.GetTempExportFolder)
    If StrComp(strTempFile, strFile, vbTextCompare) = 0 Then
        strTempFile = VCSIndex.GetTempExportFolder & FSO.GetFileName(strFile)
    End If
    If StrComp(strTempFile, strFile, vbTextCompare) = 0 Then Exit Function

    Perf.OperationStart strTimerLabel

    Set m_Table = CurrentData.AllTables(strName)
    m_lngTableType = 0
    If FSO.FileExists(strTempFile) Then DeleteFile strTempFile
    IDbComponent_Export strTempFile

    ' Hand the export back so a caller that cares *why* the two differ can diff them.
    ' It stays in the temp folder until the operation's ClearTempExportFolder.
    strComparedFile = strTempFile

    If FSO.FileExists(strTempFile) Then
        StoredDefinitionMatchesSource = (GetFileHash(strTempFile) = GetFileHash(strFile))
    End If

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : StageRelations
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Return a dictionary of the database relationships involving this table.
'           : Each relation is represented by a string value of the JSON source file
'           : content for that relation. (This way we can release the object reference
'           : and delete the actual relation object, but still recreate it later.)
'---------------------------------------------------------------------------------------
'
Private Function StageRelations(strTableName As String) As Dictionary

    Dim dbs As Database
    Dim rel As Relation
    Dim cRel As IDbComponent
    Dim strTempFile As String
    Dim varKey As Variant

    Set StageRelations = New Dictionary

    ' Loop through relations, looking for matching table name
    Set dbs = SharedDb
    For Each rel In dbs.Relations
        If StrComp(rel.Table, strTableName, vbTextCompare) = 0 Or _
            StrComp(rel.ForeignTable, strTableName, vbTextCompare) = 0 Then
            If strTempFile = vbNullString Then
                ' Create temp file for export
                strTempFile = GetTempFile
                Name strTempFile As strTempFile & ".json"
                strTempFile = strTempFile & ".json"
            End If
            ' Set up DB component class targeting this object
            Set cRel = New clsDbRelation
            Set cRel.DbObject = rel
            cRel.Export strTempFile
            ' Add file content to dictionary
            StageRelations.Add rel.Name, ReadFile(strTempFile)
        End If
    Next rel

    ' Remove the existing relationships affecting this table
    For Each varKey In StageRelations.Keys
        dbs.Relations.Delete varKey
    Next varKey

    ' Remove any temp file
    If FSO.FileExists(strTempFile) Then DeleteFile strTempFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : RestoreRelations
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Restore any staged relationships.
'---------------------------------------------------------------------------------------
'
Private Function RestoreRelations(dRelations As Dictionary)

    Dim varKey As Variant
    Dim strTempFile As String

    If dRelations.Count > 0 Then
        strTempFile = GetTempFile
        ' Add .json file extension
        Name strTempFile As strTempFile & ".json"
        strTempFile = strTempFile & ".json"
        ' Loop through relations, importing back to database from the temp file.
        For Each varKey In dRelations.Keys
            WriteFile dRelations.Item(varKey), strTempFile
            With New clsDbRelation
                .Parent.Import strTempFile
            End With
        Next varKey
        DeleteFile strTempFile
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveAlternateFormatSourceFile
' Author    : Adam Waller
' Date      : 1/25/2022
' Purpose   : Remove any existing alternate format source file. (When switching a table
'           : from local to linked, or vice versa.)
'---------------------------------------------------------------------------------------
'
Private Sub RemoveAlternateFormatSourceFile()
    Dim strAltFile As String
    strAltFile = IDbComponent_BaseFolder & GetSourceFileName(m_Table.Name, Not IsLinkedTable)
    If FSO.FileExists(strAltFile) Then DeleteFile strAltFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableSqlDef
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Save a version of the table formatted as a SQL statement.
'           : (Makes it easier to see table changes in version control systems.)
'---------------------------------------------------------------------------------------
'
Public Sub SaveTableSqlDef(strTable As String, strFolder As String)

    Dim cData As New clsConcat
    Dim cAttr As New clsConcat
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strFile As String
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

    Perf.OperationStart "Save Table SQL"

    Set dbs = SharedDb
    Set tdf = dbs.TableDefs(strTable)

    With cData
        .Add "CREATE TABLE [", DblQ(strTable), "] (", vbCrLf

        ' Loop through fields
        For Each fld In tdf.Fields
            .Add "  [", DblQ(fld.Name), "] "
            If (fld.Attributes And dbAutoIncrField) Then
                .Add "AUTOINCREMENT"
            ElseIf fld.Type = dbDecimal Then
                .Add "DECIMAL(", fld.Precision, ",", fld.Scale, ")"
            Else
                .Add GetTypeString(fld.Type)
            End If
            Select Case fld.Type
                Case dbText, dbVarBinary
                    .Add " (", fld.Size, ")"
            End Select

            ' Indexes
            For Each idx In tdf.Indexes
                Set cAttr = New clsConcat
                If idx.Fields.Count = 1 And idx.Fields(0).Name = fld.Name Then
                    If idx.Primary Then cAttr.Add " PRIMARY KEY"
                    If idx.Unique Then cAttr.Add " UNIQUE"
                    If idx.Required Then cAttr.Add " NOT NULL"
                    If idx.Foreign Then AddFieldReferences dbs, idx.Fields, strTable, cAttr
                    If Len(cAttr.GetStr) > 0 Then .Add " CONSTRAINT [", DblQ(idx.Name), "]"
                End If
                .Add cAttr.GetStr
            Next
            .Add ",", vbCrLf
        Next fld
        .Remove 3   ' strip off last comma and crlf

        ' Constraints
        If TableIndexesAvailable(tdf) Then
            Set cAttr = New clsConcat
            For Each idx In tdf.Indexes
                If idx.Fields.Count > 1 Then
                    If Len(cAttr.GetStr) = 0 Then cAttr.Add " CONSTRAINT "
                    If idx.Primary Then
                        cAttr.Add "[", DblQ(idx.Name), "] PRIMARY KEY ("
                        For Each fld In idx.Fields
                            cAttr.Add "[", DblQ(fld.Name), "], "
                        Next fld
                        cAttr.Remove 2
                        cAttr.Add ")"
                    End If
                    If Not idx.Foreign Then
                        If Len(cAttr.GetStr) > 0 Then
                            .Add ",", vbCrLf
                            .Add "  ", cAttr.GetStr
                            AddFieldReferences dbs, idx.Fields, strTable, cData
                        End If
                    End If
                End If
            Next idx
        End If
        .Add vbCrLf, ")"

        ' Build file name and create file.
        strFile = strFolder & GetSafeFileName(strTable) & ".sql"
        WriteFile .GetStr, strFile
        Perf.OperationEnd

    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddFieldReferences
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add references to other fields in table definition.
'---------------------------------------------------------------------------------------
'
Private Sub AddFieldReferences(dbs As Database, fld As Object, strTable As String, cData As clsConcat)

    Dim rel As DAO.Relation
    Dim fld2 As DAO.Field

    For Each rel In dbs.Relations
        If (rel.ForeignTable = strTable) Then
            If FieldsIdentical(fld, rel.Fields) Then

                ' References
                cData.Add " REFERENCES [", DblQ(rel.Table), "] ("
                For Each fld2 In rel.Fields
                    cData.Add "[", DblQ(fld2.Name), "],"
                Next fld2
                ' Remove trailing comma
                If rel.Fields.Count > 0 Then cData.Remove 1
                cData.Add ")"

                ' Attributes for cascade update or delete
                If rel.Attributes And dbRelationUpdateCascade Then cData.Add " ON UPDATE CASCADE "
                If rel.Attributes And dbRelationDeleteCascade Then cData.Add " ON DELETE CASCADE "

                ' Exit now that we have found the matching relationship.
                Exit For

            End If
        End If
    Next rel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FieldsIdentical
' Author    : Adam Waller
' Date      : 1/21/2019
' Purpose   : Return true if the two collections of fields have the same field names.
'           : (Even if the order of the fields is different.)
'---------------------------------------------------------------------------------------
'
Private Function FieldsIdentical(oFields1 As Object, oFields2 As Object) As Boolean

    Dim fld As Object
    Dim fld2 As Object
    Dim blnMismatch As Boolean
    Dim blnFound As Boolean

    If oFields1.Count <> oFields2.Count Then
        blnMismatch = True
    Else
        ' Set this flag to false after going through each field.
        For Each fld In oFields1
            blnFound = False
            For Each fld2 In oFields2
                If fld.Name = fld2.Name Then
                    blnFound = True
                    Exit For
                End If
            Next fld2
            If Not blnFound Then
                blnMismatch = True
                Exit For
            End If
        Next
    End If

    ' Return result
    FieldsIdentical = Not blnMismatch

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBigIntRepairFieldNamesFromTableDefXmlFile
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Read a table-definition XML file and return dbBigInt field names that
'           : Application.ImportXML would corrupt. Skips DOM parsing when the export
'           : signature is absent.
'---------------------------------------------------------------------------------------
'
Private Function GetBigIntRepairFieldNamesFromTableDefXmlFile(strXmlFile As String) As Collection

    Dim colNames As New Collection
    Dim strXml As String

    Set GetBigIntRepairFieldNamesFromTableDefXmlFile = colNames
    If Not FSO.FileExists(strXmlFile) Then Exit Function

    strXml = ReadFile(strXmlFile)
    If InStr(1, strXml, "totalDigits value=""0""", vbTextCompare) = 0 Then Exit Function

    Perf.OperationStart "Scan table XML for BigInt fields"
    Set GetBigIntRepairFieldNamesFromTableDefXmlFile = GetBigIntRepairFieldNamesFromTableDefXml(strXml)
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : FixCorruptedBigIntFields
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Correct fields that Application.ImportXML mis-created as dbDecimal(38,0)
'           : back to genuine dbBigInt fields via ALTER COLUMN ... BIGINT.
'---------------------------------------------------------------------------------------
'
Private Sub FixCorruptedBigIntFields(strTableName As String, colFieldNames As Collection)

    Dim varName As Variant
    Dim dbs As DAO.Database
    Dim dbsVerify As DAO.Database
    Dim lngVerifyType As Long

    Set dbs = SharedDb

    For Each varName In colFieldNames
        On Error Resume Next
        Err.Clear
        dbs.Execute "ALTER TABLE [" & strTableName & "] ALTER COLUMN [" & varName & "] BIGINT", dbFailOnError
        If Err Then
            CatchAny eelWarning, "Unable to restore Big Integer field type for " & _
                strTableName & "." & varName, ModuleName(Me) & ".FixCorruptedBigIntFields"
        Else
            Set dbsVerify = CurrentDb
            Err.Clear
            lngVerifyType = dbsVerify.TableDefs(strTableName).Fields(CStr(varName)).Type
            If Err Then
                CatchAny eelWarning, "Unable to verify Big Integer field type restoration for " & _
                    strTableName & "." & varName, ModuleName(Me) & ".FixCorruptedBigIntFields"
            ElseIf lngVerifyType <> dbBigInt Then
                Log.Error eelWarning, "Big Integer field type restoration did not take effect for " & _
                    strTableName & "." & varName, ModuleName(Me) & ".FixCorruptedBigIntFields"
                Err.Clear
            End If
        End If
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTypeString
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Get the type string used by Access SQL
'---------------------------------------------------------------------------------------
'
Private Function GetTypeString(intType As DAO.DataTypeEnum) As String
    Select Case intType
        Case dbLongBinary:      GetTypeString = "LONGBINARY"
        Case dbBinary:          GetTypeString = "BINARY"
        Case dbBoolean:         GetTypeString = "BIT"
        Case dbBigInt:          GetTypeString = "BIGINT"
        Case dbCurrency:        GetTypeString = "CURRENCY"
        Case dbDate, dbTime:    GetTypeString = "DATETIME"
        Case dbGUID:            GetTypeString = "GUID"
        Case dbMemo:            GetTypeString = "LONGTEXT"
        Case dbDouble:          GetTypeString = "DOUBLE"
        Case dbSingle:          GetTypeString = "SINGLE"
        Case dbByte:            GetTypeString = "BYTE"
        Case dbInteger:         GetTypeString = "SHORT"
        Case dbLong:            GetTypeString = "LONG"
        Case dbNumeric:         GetTypeString = "NUMERIC"
        Case dbText:            GetTypeString = "VARCHAR"
        Case Else:              GetTypeString = "VARCHAR"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportLinkedTable
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Recreate a linked table from the JSON source file.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function ImportLinkedTable(strFile As String) As Boolean

    Dim dTable As Dictionary
    Dim dItem As Dictionary
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strSql As String
    Dim strKeyFields As String
    Dim strConnect As String
    Dim strEnvKey As String
    Dim strKey As String
    Dim dState As Dictionary
    Dim blnLinked As Boolean
    Dim lngErrNum As Long
    Dim strErrDesc As String
    Dim eAction As eConnFailAction
    Dim dFieldsImport As Dictionary
    Dim dFieldPropsImport As Dictionary
    Dim dTablePropsImport As Dictionary
    Dim dProp As Dictionary
    Dim varField As Variant
    Dim varPropName As Variant

    LogUnhandledErrors
    On Error Resume Next

    ' Read json file
    Set dTable = ReadJsonFile(strFile)
    If Not dTable Is Nothing Then

        ' Link the table
        Set dItem = dTable("Items")

        ' A tbldefs .json is a linked-table definition only when it carries Connect.
        ' Without Connect it is a local-table metadata sidecar (schema is in the .xml;
        ' metadata is applied when that .xml imports).
        If Not dItem.Exists("Connect") Then Exit Function

        Set dbs = SharedDb

        strConnect = dItem("Connect")
        strEnvKey = vbNullString
        If IsEnvReference(strConnect) Then
            strEnvKey = Mid$(strConnect, 5)
            strConnect = ResolveEnvConnection(strConnect)
        End If
        strKey = IIf(Len(strEnvKey) > 0, strEnvKey, GetConnectionEnvKey(strConnect))
        Set dState = GetConnState(strKey)
        If dState Is Nothing Then Exit Function

        If dState("Status") = csIgnored Then
            Log.Add "  " & T("Skipped {0} (connection ignored)", var0:=dItem("Name")), Options.ShowDebug
            Exit Function
        ElseIf dState("Status") = csOK And Len(dState("Completed")) > 0 Then
            strConnect = dState("Completed")
        ElseIf Len(strConnect) = 0 Then
            If Operation.InteractionMode = eimNormal And Options.UseEnvForConnections <> uecNever Then
                strConnect = "ODBC;"
            Else
                Log.Add "  " & T("Skipped {0} (missing .env key)", var0:=dItem("Name")), Options.ShowDebug
                Exit Function
            End If
        End If
        strConnect = GetFullConnect(strConnect)

        blnLinked = False
        Do
            Set tdf = dbs.CreateTableDef(dItem("Name"))
            With tdf
                .Connect = strConnect
                .SourceTableName = dItem("SourceTableName")
                ' Honor the source table's Save-Password flag (in Attributes). When the
                ' user chose to save the password, the linked table carries the full
                ' credentials so RecordAuthenticatedConnection can persist them to .env;
                ' otherwise the password is not stored. The authentication method
                ' (e.g. Authentication=...) is retained either way, which is enough to
                ' capture passwordless (AD/trusted) connections.
                .Attributes = SafeAttributes(dItem("Attributes"))
            End With
            dbs.TableDefs.Append tdf
            lngErrNum = Err.Number
            strErrDesc = Err.Description
            If lngErrNum = 0 Then
                blnLinked = True
                Exit Do
            End If
            If Catch(3011, 3024, 3044) Then
                Log.Error eelError, "Could not link table '" & dItem("SourceTableName") & "'", _
                    ModuleName(Me) & ".ImportLinkedTable"
                Log.Add "Linked table object not found in " & strFile, False
                Log.Add "Connection String: " & strConnect, False
                Exit Function
            ElseIf IsConnectionError(lngErrNum) Then
                Err.Clear
                eAction = HandleConnectionFailure(strKey, dItem("Name") & " (linked table)", lngErrNum, strErrDesc)
                Select Case eAction
                    Case cfaRetry
                    Case cfaIgnore
                        Log.Add "  " & T("Skipped {0} (connection ignored)", var0:=dItem("Name")), Options.ShowDebug
                        Exit Function
                    Case cfaAbort
                        Log.Error eelCritical, _
                            T("Connection canceled by user. Canceling operation."), _
                            ModuleName(Me) & ".ImportLinkedTable"
                        Exit Function
                End Select
            Else
                CatchAny eelError, vbNullString, ModuleName(Me) & ".ImportLinkedTable"
                Exit Function
            End If
        Loop

        If Not blnLinked Then Exit Function

        ' Verify that the connection matches the source file. (Issue #192)
        If tdf.Connect <> strConnect Then
            Do
                tdf.Connect = strConnect
                tdf.RefreshLink
                lngErrNum = Err.Number
                strErrDesc = Err.Description
                If lngErrNum = 0 Then Exit Do
                If Catch(3283) Then
                    Log.Add "Encountered error 3283 when refreshing link for " & tdf.Name & ". " & _
                        "You can probably safely ignore this error. (See issue #484)"
                    Exit Do
                End If
                If IsConnectionError(lngErrNum) Then
                    Err.Clear
                    eAction = HandleConnectionFailure(strKey, dItem("Name") & " (linked table)", lngErrNum, strErrDesc)
                    Select Case eAction
                        Case cfaRetry
                        Case cfaIgnore
                            Log.Add "  " & T("Skipped {0} (connection ignored)", var0:=dItem("Name")), Options.ShowDebug
                            Exit Function
                        Case cfaAbort
                            Log.Error eelCritical, _
                                T("Connection canceled by user. Canceling operation."), _
                                ModuleName(Me) & ".ImportLinkedTable"
                            Exit Function
                    End Select
                Else
                    CatchAny eelError, "Error refreshing link for " & tdf.Name, ModuleName(Me) & ".ImportLinkedTable"
                    Exit Function
                End If
            Loop
        End If
        dbs.TableDefs.Refresh

        ' Set index on linked table.
        If InStr(1, tdf.Connect, ";DATABASE=", vbTextCompare) = 1 Then
            ' Can't create a key on a linked Access database table.
            ' Presumably this would use the Access index instead of needing the pseudo index
        Else
            ' Check for a primary key index (Linked SQL tables may bring over the index, but linked views won't.)
            If dItem.Exists("PrimaryKey") And Not HasUniqueIndex(tdf) Then
                strKeyFields = BracketFieldList(CStr(dItem("PrimaryKey")))
                If Len(strKeyFields) Then
                    ' Create a pseudo index on the linked table
                    strSql = "CREATE UNIQUE INDEX __uniqueindex ON [" & DblQ(tdf.Name) & "] (" & strKeyFields & ") WITH PRIMARY"
                    dbs.Execute strSql, dbFailOnError
                    dbs.TableDefs.Refresh
                End If
            End If
        End If

        ' Restore front-end display properties for fields
        If dItem.Exists("FieldProperties") Then
            Perf.OperationStart "Import Linked Field Props"
            Set dFieldsImport = dItem("FieldProperties")
            For Each varField In dFieldsImport.Keys
                Set dFieldPropsImport = dFieldsImport(varField)
                For Each varPropName In dFieldPropsImport.Keys
                    If TypeOf dFieldPropsImport(varPropName) Is Dictionary Then
                        Set dProp = dFieldPropsImport(varPropName)
                        SetDAOProperty tdf.Fields(CStr(varField)), _
                            dProp("Type"), CStr(varPropName), dProp("Value")
                    End If
                Next varPropName
            Next varField
            Perf.OperationEnd
        End If

        ' Restore front-end display properties for the table
        If dItem.Exists("TableProperties") Then
            Perf.OperationStart "Import Linked Table Props"
            ' Engine-managed properties (e.g. FCMin* version stamps) that older source
            ' may still contain are tolerated by SetDAOProperty (error 3916) rather than
            ' aborting the import, so no name filtering is needed here.
            Set dTablePropsImport = dItem("TableProperties")
            For Each varPropName In dTablePropsImport.Keys
                If TypeOf dTablePropsImport(varPropName) Is Dictionary Then
                    Set dProp = dTablePropsImport(varPropName)
                    SetDAOProperty tdf, dProp("Type"), CStr(varPropName), dProp("Value")
                End If
            Next varPropName
            Perf.OperationEnd
        End If

        RecordAuthenticatedConnection strKey, strEnvKey, tdf.Connect, dItem("Name") & " (linked table)"
        ImportLinkedTable = (Err.Number = 0)
    End If

    ' Report any unhandled errors
    CatchAny eelError, "Error importing " & strFile, ".ImportLinkedTable"

End Function


'---------------------------------------------------------------------------------------
' Procedure : SafeAttributes
' Author    : Adam Waller
' Date      : 6/29/2020
' Purpose   : Rebuild new attributes flag using attributes that we can actually set.
'---------------------------------------------------------------------------------------
'
Private Function SafeAttributes(lngAttributes As Long) As Long

    Dim colAtts As Collection
    Dim varAtt As Variant
    Dim lngNew As Long

    Set colAtts = New Collection
    With colAtts
        '.Add dbAttachedODBC
        '.Add dbAttachedTable
        .Add dbAttachExclusive
        .Add dbAttachSavePWD
        .Add dbHiddenObject
        .Add dbSystemObject
    End With

    For Each varAtt In colAtts
        ' Use boolean logic to check for bit flag
        If CBool((lngAttributes And varAtt) = varAtt) Then
            ' Add to our rebuilt flag value.
            lngNew = lngNew + varAtt
        End If
    Next varAtt

    ' Return attributes value after rebuilding from scratch.
    SafeAttributes = lngNew

End Function


'---------------------------------------------------------------------------------------
' Procedure : HasUniqueIndex
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function HasUniqueIndex(tdf As TableDef) As Boolean

    Dim idx As DAO.Index

    If TableIndexesAvailable(tdf) Then
        For Each idx In tdf.Indexes
            If idx.Unique Then
                HasUniqueIndex = True
                Exit For
            End If
        Next idx
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveComponentSource Me, strFromFolder, strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim tdf As AccessObject
    Dim cTable As IDbComponent
    Dim blnNeedAll As Boolean
    Dim sngStart As Single
    Dim lngScanCount As Long
    Dim dSysTables As Dictionary

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary
        If blnModifiedOnly Then
            lngScanCount = CurrentData.AllTables.Count
            sngStart = Timer
        End If
        ' Batch-load table Types from MSysObjects once, so per-table SourceFile /
        ' IsLinkedTable checks avoid hundreds of slow single-row Type lookups.
        BuildTableTypeCache
        Set dSysTables = GetSystemTableNames
        For Each tdf In CurrentData.AllTables
            If dSysTables.Exists(tdf.Name) Or tdf.Name Like "~*" Then
                ' Skip system and temporary tables
            Else
                Set cTable = New clsDbTableDef
                Set cTable.DbObject = tdf
                If blnNeedAll Then m_Items(False).Add cTable.SourceFile, cTable
                If blnModifiedOnly Then
                    If cTable.IsModified Then m_Items(True).Add cTable.SourceFile, cTable
                    Log.IncrementObjectScanProgress sngStart, lngScanCount, "Tables"
                End If
            End If
        Next tdf
        ClearTableTypeCache
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'           : Local tables export schema as .xml and metadata as a sibling .json sidecar;
'           : only .json files without a matching .xml are linked-table definitions.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary

    Dim dAll As Dictionary
    Dim dXmlBaseNames As Dictionary
    Dim varKey As Variant
    Dim strExt As String

    If m_FileList Is Nothing Then
        Set dAll = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.xml", "*.json")
        Set dXmlBaseNames = New Dictionary
        For Each varKey In dAll.Keys
            If LCase$(FSO.GetExtensionName(CStr(varKey))) = "xml" Then
                dXmlBaseNames.Add FSO.GetBaseName(CStr(varKey)), vbNullString
            End If
        Next varKey
        Set m_FileList = New Dictionary
        For Each varKey In dAll.Keys
            strExt = LCase$(FSO.GetExtensionName(CStr(varKey)))
            If strExt = "json" Then
                If dXmlBaseNames.Exists(FSO.GetBaseName(CStr(varKey))) Then
                    ' Metadata sidecar for a local table; schema imports from the .xml.
                Else
                    m_FileList.Add varKey, vbNullString
                End If
            Else
                m_FileList.Add varKey, vbNullString
            End If
        Next varKey
    End If
    Set IDbComponent_GetFileList = m_FileList

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = (DateTruncToSeconds(m_Table.DateModified) <> VCSIndex.Item(Me).ObjectDate)

    ' Check metadata (Description/Hidden) which don't update DateModified
    If Not IDbComponent_IsModified Then
        IDbComponent_IsModified = VCSIndex.Item(Me).MetaHash <> _
            GetMetadataHash("Tables", m_Table.Name, acTable)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    If m_Table Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Table.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Tables"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "tbldefs" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions(Optional Scope As eFileExtensionScope = efesIndexed) As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "xml"
    IDbComponent_FileExtensions.Add "json"
    IDbComponent_FileExtensions.Add "sql"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_Name = m_Table.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSourceFileName(m_Table.Name, IsLinkedTable)
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsLinkedTable
' Author    : Adam Waller
' Date      : 2/12/2022
' Purpose   : Returns true if this table is a linked table
'---------------------------------------------------------------------------------------
'
Private Property Get IsLinkedTable() As Boolean

    Dim dbs As Database
    Dim rst As Recordset

    ' Nothing to check if we don't have a table name
    If m_Table Is Nothing Then Exit Property

    ' Use cached value if possible (per table object on this instance)
    If m_lngTableType > 0 Then
        IsLinkedTable = (m_lngTableType <> 1)
        Exit Property
    End If

    ' Use module-level batch cache when active (built by GetAllFromDB).
    m_lngTableType = GetCachedTableType(m_Table.Name)
    If m_lngTableType > 0 Then
        IsLinkedTable = (m_lngTableType <> 1)
        Exit Property
    End If

    Perf.OperationStart "Check for linked table"

    ' This is the simple way to check for a linked table, but it is not very performant
    ' when called hundreds of times while scanning tables for changes.
    'IsLinkedTable = CurrentDb.TableDefs(m_Table.Name).Connect <> vbNullString

    ' A faster way is to read the object type from MSysObjects
    Set dbs = SharedDb
    Set rst = dbs.OpenRecordset( _
        "select Type from MSysObjects" & _
        " where name =""" & m_Table.Name & """" & _
        " and Type in(1,4,6)", _
        dbOpenDynaset, dbReadOnly)
    ' Read the table type from the system objects record.
    m_lngTableType = Nz(rst!Type, 1)
    rst.Close
    Set rst = Nothing
    Set dbs = Nothing

    ' A value of 1 means a local Access table. Any other value
    ' is almost certainly a linked table of some type. (ODBC, Excel, etc)
    IsLinkedTable = (m_lngTableType <> 1)
    Perf.OperationEnd

End Property


'---------------------------------------------------------------------------------------
' Procedure : GetSourceFileName
' Author    : Adam Waller
' Date      : 1/25/2022
' Purpose   : Return a safe file name with the correct file extension.
'---------------------------------------------------------------------------------------
'
Private Function GetSourceFileName(strName As String, blnLinkedTable As Boolean) As String
    If blnLinkedTable Then
        GetSourceFileName = GetSafeFileName(strName) & ".json"
    Else
        GetSourceFileName = GetSafeFileName(strName) & ".xml"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    Static lngCount As Long
    If lngCount = 0 Then lngCount = CurrentData.AllTables.Count - 1
    IDbComponent_QuickCount = lngCount + 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbTableDef
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Table
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Table = RHS
    m_lngTableType = 0
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
