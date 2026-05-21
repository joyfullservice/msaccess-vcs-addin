VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVCSIndex
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   : Maintain index of source files and database objects so that changes
'           : can be detected.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Infrastructure")

' File name for index
Private Const cstrFileName As String = "vcs-index.idx"

' General properties
Public MergeBuildDate As Date
Public FullBuildDate As Date
Public ExportDate As Date
Public FullExportDate As Date
Public VBAProjectDate As Date
Public Disabled As Boolean

' Git integration
Public LastMergedCommit As String

' Action types for update function
Public Enum eIndexOperationType
    eatExport
    eatImport
    eatAltExport    ' Alternate export folder, such as a scan for changes
End Enum

' Index of component/file information, based on source files
Private m_dIndex As Dictionary
Private m_dGitIndex As Dictionary
Private m_strFile As String
Private m_Git As clsGitIntegration
Private m_Conflicts As clsConflicts
Private m_strTempExportFolderPath As String

' Binary index format
Private Const cstrIdxMagic As String = "VCSI"
Private Const cintIdxFormatVersion As Integer = 3
Private Const cstrLegacyFileName As String = "vcs-index.json"

' Optional field flags for component entries
Private Const cbytFlagFileHash As Byte = 1
Private Const cbytFlagOtherHash As Byte = 2
Private Const cbytFlagMetaHash As Byte = 4
Private Const cbytFlagFilePropertiesHash As Byte = 8
Private Const cbytFlagFolderAnnotation As Byte = 16

' UDTs for LSet Double <-> Byte() punning
Private Type tDoubleBytes
    b(0 To 7) As Byte
End Type
Private Type tDoubleVal
    Value As Double
End Type

' Write buffer state (used during Save)
Private m_bBuffer() As Byte
Private m_lngBufPos As Long
Private m_lngBufSize As Long


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Load the state for the project.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromFile(Optional strFile As String)

    Dim bData() As Byte
    Dim lngPos As Long
    Dim lngCatCount As Long
    Dim lngItemCount As Long
    Dim lngHashCount As Long
    Dim lngCat As Long
    Dim lngItem As Long
    Dim lngHash As Long
    Dim strCat As String
    Dim strFileName As String
    Dim strLegacyPath As String
    Dim bFlags As Byte
    Dim cEntry As clsVCSIndexItem
    Dim lngVersion As Long
    Dim intCnt As Integer
    Dim dCategory As Dictionary

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Reset class to blank values
    Call Class_Initialize

    ' Determine file path
    m_strFile = strFile
    If m_strFile = vbNullString Then m_strFile = DefaultFilePath
    If m_strFile = vbNullString Then Exit Sub

    ' Check for legacy JSON index and clean up
    If Not FSO.FileExists(m_strFile) Then
        strLegacyPath = Replace(m_strFile, cstrFileName, cstrLegacyFileName)
        If FSO.FileExists(strLegacyPath) Then
            LogUnhandledErrors
            On Error Resume Next
            FSO.DeleteFile strLegacyPath, True
            CatchAny eelWarning, "Removed legacy index: " & cstrLegacyFileName, _
                ModuleName(Me) & ".LoadFromFile"
        End If
        Exit Sub
    End If

    Perf.OperationStart "Load Index"

    ' Any read error (truncated file, bad data) is treated as corruption.
    On Error GoTo CorruptIndex

    ' Read entire file as byte array
    bData = GetFileBytes(m_strFile)
    lngPos = 0

    ' Validate magic bytes
    If UBound(bData) < Len(cstrIdxMagic) + 1 Then
        GoTo CorruptIndex
    End If
    For intCnt = 1 To Len(cstrIdxMagic)
        If bData(intCnt - 1) <> CByte(Asc(Mid$(cstrIdxMagic, intCnt, 1))) Then
            GoTo CorruptIndex
        End If
    Next intCnt
    lngPos = Len(cstrIdxMagic)

    ' Read and validate format version
    lngVersion = ReadUInt16(bData, lngPos)
    If lngVersion <> cintIdxFormatVersion Then GoTo CorruptIndex

    ' Read global properties (UTC dates converted to local)
    Me.MergeBuildDate = ReadUtcDate(bData, lngPos)
    Me.FullBuildDate = ReadUtcDate(bData, lngPos)
    Me.ExportDate = ReadUtcDate(bData, lngPos)
    Me.FullExportDate = ReadUtcDate(bData, lngPos)
    Me.LastMergedCommit = ReadLPS(bData, lngPos)
    Me.VBAProjectDate = ReadUtcDate(bData, lngPos)

    ' Read category hashes
    lngHashCount = ReadUInt32(bData, lngPos)
    For lngHash = 1 To lngHashCount
        strCat = ReadLPS(bData, lngPos)
        m_dIndex("CategoryHashes").Add strCat, ReadLPS(bData, lngPos)
    Next lngHash

    ' Read components
    lngCatCount = ReadUInt32(bData, lngPos)
    For lngCat = 1 To lngCatCount
        strCat = ReadLPS(bData, lngPos)
        lngItemCount = ReadUInt32(bData, lngPos)

        ' Create category dictionary
        Set dCategory = New Dictionary
        m_dIndex("Components").Add strCat, dCategory

        For lngItem = 1 To lngItemCount
            strFileName = ReadLPS(bData, lngPos)
            bFlags = ReadByte(bData, lngPos)

            Set cEntry = New clsVCSIndexItem
            cEntry.ExportDate = ReadUtcDate(bData, lngPos)
            cEntry.ImportDate = ReadUtcDate(bData, lngPos)
            cEntry.SourceModified = ReadUtcDate(bData, lngPos)
            cEntry.ObjectDate = ReadUtcDate(bData, lngPos)

            If (bFlags And cbytFlagFileHash) Then cEntry.FileHash = ReadLPS(bData, lngPos)
            If (bFlags And cbytFlagOtherHash) Then cEntry.OtherHash = ReadLPS(bData, lngPos)
            If (bFlags And cbytFlagMetaHash) Then cEntry.MetaHash = ReadLPS(bData, lngPos)
            If (bFlags And cbytFlagFilePropertiesHash) Then cEntry.FilePropertiesHash = ReadLPS(bData, lngPos)
            If (bFlags And cbytFlagFolderAnnotation) Then cEntry.FolderAnnotation = ReadLPS(bData, lngPos)

            dCategory.Add strFileName, cEntry
        Next lngItem
    Next lngCat

    Perf.OperationEnd
    Exit Sub

CorruptIndex:
    ' Reset to empty state and remove the corrupt file
    Call Class_Initialize
    m_strFile = strFile
    If m_strFile = vbNullString Then m_strFile = DefaultFilePath
    LogUnhandledErrors
    On Error Resume Next
    If FSO.FileExists(m_strFile) Then FSO.DeleteFile m_strFile, True
    CatchAny eelWarning, "Removed corrupt index file. A full export will regenerate it.", _
        ModuleName(Me) & ".LoadFromFile"
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Save
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Save to a file
'---------------------------------------------------------------------------------------
'
Public Sub Save(Optional strInFolder As String)

    Dim strFile As String
    Dim varCat As Variant
    Dim varFile As Variant
    Dim dComponents As Dictionary
    Dim dCategory As Dictionary
    Dim dCatHashes As Dictionary
    Dim cEntry As clsVCSIndexItem
    Dim bFlags As Byte

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    Perf.OperationStart "Save Index"

    ' Sort files and components for deterministic output
    SortComponentSection m_dIndex, "Components"

    ' Remove the AlternateExport section, since this is not
    ' needed after the completion of the export process.
    If m_dIndex.Exists("AlternateExport") Then m_dIndex.Remove "AlternateExport"

    ' Build file path
    If strInFolder = vbNullString Then
        strFile = m_strFile
    Else
        strFile = StripSlash(strInFolder) & PathSep & cstrFileName
    End If

    If strFile = vbNullString Then
        Perf.OperationEnd
        Exit Sub
    End If

    ' Initialize the write buffer
    InitWriteBuffer

    ' Header: magic + format version (UInt16 LE)
    WriteMagic cstrIdxMagic
    WriteUInt16 CLng(cintIdxFormatVersion)

    ' Global properties (dates stored as UTC)
    WriteUtcDate Me.MergeBuildDate
    WriteUtcDate Me.FullBuildDate
    WriteUtcDate Me.ExportDate
    WriteUtcDate Me.FullExportDate
    WriteLPS Me.LastMergedCommit
    WriteUtcDate Me.VBAProjectDate

    ' Category hashes
    Set dCatHashes = m_dIndex("CategoryHashes")
    WriteUInt32 dCatHashes.Count
    For Each varCat In dCatHashes.Keys
        WriteLPS CStr(varCat)
        WriteLPS CStr(dCatHashes(varCat))
    Next varCat

    ' Components
    Set dComponents = m_dIndex("Components")
    WriteUInt32 dComponents.Count
    For Each varCat In dComponents.Keys
        WriteLPS CStr(varCat)
        Set dCategory = dComponents(varCat)
        WriteUInt32 dCategory.Count
        For Each varFile In dCategory.Keys
            Set cEntry = dCategory(varFile)
            WriteLPS CStr(varFile)

            ' Build flags byte for optional hash fields
            bFlags = 0
            If Len(cEntry.FileHash) > 0 Then bFlags = bFlags Or cbytFlagFileHash
            If Len(cEntry.OtherHash) > 0 Then bFlags = bFlags Or cbytFlagOtherHash
            If Len(cEntry.MetaHash) > 0 Then bFlags = bFlags Or cbytFlagMetaHash
            If Len(cEntry.FilePropertiesHash) > 0 Then bFlags = bFlags Or cbytFlagFilePropertiesHash
            If Len(cEntry.FolderAnnotation) > 0 Then bFlags = bFlags Or cbytFlagFolderAnnotation
            WriteByte bFlags

            ' Fixed date fields (UTC)
            WriteUtcDate cEntry.ExportDate
            WriteUtcDate cEntry.ImportDate
            WriteUtcDate cEntry.SourceModified
            WriteUtcDate cEntry.ObjectDate

            ' Conditional hash fields
            If (bFlags And cbytFlagFileHash) Then WriteLPS cEntry.FileHash
            If (bFlags And cbytFlagOtherHash) Then WriteLPS cEntry.OtherHash
            If (bFlags And cbytFlagMetaHash) Then WriteLPS cEntry.MetaHash
            If (bFlags And cbytFlagFilePropertiesHash) Then WriteLPS cEntry.FilePropertiesHash
            If (bFlags And cbytFlagFolderAnnotation) Then WriteLPS cEntry.FolderAnnotation
        Next varFile
    Next varCat

    ' Write the buffer to file
    WriteBinaryFile strFile, GetWriteBuffer()

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SortComponentSection
' Author    : Adam Waller
' Date      : 9/1/2022
' Purpose   : Save a named section of components as
'---------------------------------------------------------------------------------------
'
Private Sub SortComponentSection(ByRef dIndex As Dictionary, strSection As String)

    Dim dComponents As Dictionary
    Dim varCat As Variant

    If dIndex.Exists(strSection) Then
        Set dComponents = dIndex(strSection)
        For Each varCat In dComponents.Keys
            ' Sort list of files in each category
            Set dComponents(varCat) = SortDictionaryByKeys(dComponents(varCat))
        Next varCat

        ' Sort list of categories
        Set dIndex(strSection) = SortDictionaryByKeys(dComponents)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Updates an item in the index.
'---------------------------------------------------------------------------------------
'
Public Function Update(cItem As IDbComponent, intAction As eIndexOperationType, _
    strFileHash As String, Optional strOtherHash As String, _
    Optional dteDateTime As Date, Optional strMetaHash As String) As clsVCSIndexItem

    Dim cEntry As clsVCSIndexItem
    Dim strSection As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Function

    ' Determine section
    If intAction = eatAltExport Then
        strSection = "AlternateExport"
    Else
        strSection = "Components"
    End If

    ' Look up item, creating if needed.
    Set cEntry = Me.Item(cItem, , strSection)

    With cEntry
        .FileHash = strFileHash
        .OtherHash = strOtherHash
        .MetaHash = strMetaHash

        ' Add timestamp, defaulting to now
        If dteDateTime = 0 Then dteDateTime = Now
        Select Case intAction
            Case eatExport, eatAltExport
                .ExportDate = dteDateTime
            Case eatImport
                .ImportDate = dteDateTime
        End Select

        ' Save the object's DateModified for accurate change detection.
        ' Truncate to seconds since serialization cannot preserve sub-second
        ' precision, and Access DateModified may carry fractional seconds.
        ' For modules, use the top-level VBAProjectDate instead (all module
        ' dates are always identical since Access saves VBA as a monolith).
        If cItem.ComponentType = edbModule Then
            Me.VBAProjectDate = DateTruncToSeconds(cItem.DateModified)
        Else
            .ObjectDate = DateTruncToSeconds(cItem.DateModified)
        End If

        ' Save timestamp and property hash of exported source file.
        ' (Skip for alternate exports -- the canonical source file may not
        '  exist yet, and these values are recomputed when the alt entry
        '  is promoted via UpdateFromAltExport.)
        If intAction <> eatAltExport Then
            dteDateTime = GetSourceModifiedDate(cItem, cItem.SourceFile)
            .SourceModified = dteDateTime
            .FilePropertiesHash = GetSourceFilesPropertyHash(cItem)
        End If
    End With

    Set Update = cEntry

End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateFromAltExport
' Author    : Adam Waller
' Date      : 9/1/2022
' Purpose   : Update the index entry from the alternate export section. This is done
'           : after resolving a conflict where the exported file in the temp folder is
'           : moved to the primary export folder. (This way we don't have to compute
'           : the hashes and index values again.
'---------------------------------------------------------------------------------------
'
Public Sub UpdateFromAltExport(cItem As IDbComponent)

    Dim strFile As String
    Dim cAltItem As clsVCSIndexItem
    Dim cTarget As clsVCSIndexItem

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Get file name from component
    strFile = FSO.GetFileName(cItem.SourceFile)

    ' Look for entry in AlternateExport section
    With m_dIndex("AlternateExport")
        If .Exists(cItem.Category) Then
            If .Item(cItem.Category).Exists(strFile) Then
                Set cAltItem = .Item(cItem.Category).Item(strFile)
            End If
        End If
    End With

    ' Copy all properties from the alternate export entry
    ' to the main component entry.
    If Not cAltItem Is Nothing Then
        Set cTarget = LoadItem(cItem, strFile, "Components")
        With cTarget
            .FileHash = cAltItem.FileHash
            .OtherHash = cAltItem.OtherHash
            .MetaHash = cAltItem.MetaHash
            .FilePropertiesHash = cAltItem.FilePropertiesHash
            .FolderAnnotation = cAltItem.FolderAnnotation
            .ExportDate = cAltItem.ExportDate
            .ImportDate = cAltItem.ImportDate
            .SourceModified = cAltItem.SourceModified
            .ObjectDate = cAltItem.ObjectDate
        End With
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Remove an item from the index when the object and file no longer exist.
'---------------------------------------------------------------------------------------
'
Public Sub Remove(cItem As IDbComponent, Optional strSourceFile As String)

    Dim strFile As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(Nz2(strSourceFile, cItem.SourceFile))

    ' Remove dictionary objects.
    With m_dIndex("Components")
        If .Exists(cItem.Category) Then
            If .Item(cItem.Category).Exists(strFile) Then
                .Item(cItem.Category).Remove strFile
                ' Remove category if no more items
                If .Item(cItem.Category).Count = 0 Then
                    .Remove cItem.Category
                End If
            End If
        End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MigrateIndexExtension
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Renames the file extension on index keys for a given category.
'           : Only renames keys whose current extension matches strOldExt,
'           : leaving secondary extensions (e.g. .cls, .sql) untouched.
'---------------------------------------------------------------------------------------
'
Public Sub MigrateIndexExtension(strCategory As String, strOldExt As String, strNewExt As String)

    Dim dSection As Dictionary
    Dim varSection As Variant
    Dim varKey As Variant
    Dim strExpected As String

    For Each varSection In Array(m_dIndex("Components"), m_dIndex("AlternateExport"))
        Set dSection = varSection
        If dSection.Exists(strCategory) Then
            For Each varKey In dSection(strCategory).Keys
                If StrComp(FSO.GetExtensionName(CStr(varKey)), strOldExt, vbTextCompare) = 0 Then
                    strExpected = SwapExtension(CStr(varKey), strNewExt)
                    If dSection(strCategory).Exists(strExpected) Then
                        ' Target key already exists; remove the stale old-extension entry
                        dSection(strCategory).Remove CStr(varKey)
                    Else
                        dSection(strCategory).Key(CStr(varKey)) = strExpected
                    End If
                End If
            Next
        End If
    Next

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetItem
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Returns a dictionary object with the saved values, creating if needed.
'---------------------------------------------------------------------------------------
'
Public Function Item(cItem As IDbComponent, Optional strSourceFile As String, _
    Optional strSection As String = "Components") As clsVCSIndexItem

    Dim strFile As String

    ' Return an empty item when disabled so callers can safely
    ' access properties (they will see blank/default values).
    If Me.Disabled Then
        Set Item = New clsVCSIndexItem
        Exit Function
    End If

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(Nz2(strSourceFile, cItem.SourceFile))
    Set Item = LoadItem(cItem, strFile, strSection)

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadItem
' Author    : Adam Waller
' Date      : 9/23/2021
' Purpose   : Loads the item from the index, creating if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Function LoadItem(cItem As IDbComponent, strSourceFile As String, strSection As String) As clsVCSIndexItem

    ' Get or create category dictionary, then get or create the item.
    With m_dIndex(strSection)
        If Not .Exists(cItem.Category) Then Set .Item(cItem.Category) = New Dictionary
        If Not .Item(cItem.Category).Exists(strSourceFile) Then
            Set .Item(cItem.Category)(strSourceFile) = New clsVCSIndexItem
        End If
        Set LoadItem = .Item(cItem.Category)(strSourceFile)
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCachedAnnotation
' Author    : Adam Waller
' Date      : 5/5/2026
' Purpose   : Return the cached @Folder annotation for a component, looked up by
'           : category name and source file name. This avoids the SourceFile property
'           : (which itself depends on the annotation) to break the circular dependency.
'           : Returns vbNullString when the item is not in the index.
'---------------------------------------------------------------------------------------
'
Public Function GetCachedAnnotation(strCategory As String, strFileName As String) As String

    Dim dCat As Dictionary

    If Me.Disabled Then Exit Function

    With m_dIndex("Components")
        If .Exists(strCategory) Then
            Set dCat = .Item(strCategory)
            If dCat.Exists(strFileName) Then
                GetCachedAnnotation = dCat(strFileName).FolderAnnotation
            End If
        End If
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Returns true if the file exists in the index.
'---------------------------------------------------------------------------------------
'
Public Function Exists(cCategory As IDbComponent, strSourceFilePath As String) As Boolean

    Dim strFile As String
    Dim blnExists

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(strSourceFilePath)

    ' See if the entry exists in the index
    With m_dIndex("Components")
        If .Exists(cCategory.Category) Then
            blnExists = .Item(cCategory.Category).Exists(strFile)
        End If
    End With

    ' Return result
    Exists = blnExists

End Function


'---------------------------------------------------------------------------------------
' Procedure : DefaultDevModeHash
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Return a hash of the printer settings from the default printer. (Used to
'           : determine whether a form or report is using any specific or custom
'           : print settings, thereby requiring us to save the print settings.
'---------------------------------------------------------------------------------------
'
Public Property Get DefaultDevModeHash() As String

    ' Cache the result for future calls
    Static strHash As String

    If strHash = vbNullString Then
        With New clsPrinterSettings
            .LoadFromDefaultPrinter
            strHash = .GetHash
        End With
    End If

    ' Return hash
    DefaultDevModeHash = strHash

End Property


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Return a collection of source files that appear to be different from
'           : the index hash of the previous export. File modified dates will be used
'           ' to determine which files have changed.
'           : NOTE: This will also include source file paths for database objects that
'           : no longer exist. (Orphaned database objects not represented in source.)
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCategory As IDbComponent) As Dictionary

    Dim dAllFiles As Dictionary
    Dim dItems As Dictionary
    Dim varItem As Variant
    Dim cItem As IDbComponent
    Dim strHash As String
    Dim strFile As String
    Dim blnModified As Boolean
    Dim blnMissing As Boolean
    Dim strStoredHash As String

    Set GetModifiedSourceFiles = New Dictionary
    With GetModifiedSourceFiles

        ' Get a full list of source files and database objects
        Set dAllFiles = cCategory.GetFileList
        Set dItems = cCategory.GetAllFromDB

        ' Loop through files and check modified dates.
        ' Get a list of all the files for this component.
        For Each varItem In dAllFiles
            strFile = varItem

            ' Reset flags
            blnModified = True
            blnMissing = False

            ' Single-file components always return their source path from
            ' GetFileList even when the file was never exported or was
            ' removed during a format migration.  Skip non-existent files
            ' so they are not passed to Merge (which could be destructive).
            If cCategory.SingleFile And (Not FSO.FileExists(strFile)) Then
                blnModified = False
            Else
                ' See if the database object currently exists
                If Not cCategory.SingleFile Then
                    blnMissing = Not dItems.Exists(strFile)
                End If

                ' Check the index for any existing entry, so we can compare the saved modified date.
                ' NOTE: We intentionally do NOT gate this on blnMissing. The
                ' path-based check (dItems.Exists) can fail due to casing,
                ' @Folder annotation drift, or stale index state. As long as
                ' the index has an entry keyed by filename we can still compare
                ' hashes and avoid unnecessary merges.
                If Me.Exists(cCategory, strFile) Then
                    ' Get hash of current source files.
                    strHash = GetSourceFilesPropertyHash(cCategory, strFile)
                    ' File is considered not modified if the index hash matches the current hash.
                    blnModified = Not (Me.Item(cCategory, strFile).FilePropertiesHash = strHash)
                    ' Content-hash fallback: if property hash differs (e.g. git
                    ' checkout touched timestamps without changing content), verify
                    ' by comparing the stored file content hash.
                    If blnModified Then
                        strStoredHash = Me.Item(cCategory, strFile).FileHash
                        If Len(strStoredHash) > 0 Then
                            blnModified = (strStoredHash <> GetFileHash(strFile))
                            If Not blnModified Then
                                ' Content unchanged -- refresh stale property hash so this
                                ' fallback doesn't repeat on the next operation.
                                Me.Item(cCategory, strFile).FilePropertiesHash = strHash
                                Me.Item(cCategory, strFile).SourceModified = GetSourceModifiedDate(cCategory, strFile)
                            End If
                        End If
                    End If
                End If
            End If

            ' Add modified files to collection
            If blnModified Then .Add strFile, vbNullString
        Next varItem

        ' Now loop through database objects, and compare to source file list
        ' to identify orphaned database objects no longer represented in source.
        For Each varItem In dItems.Items
            Set cItem = varItem
            ' We are only concerned about non single-file items.
            If Not cItem.SingleFile Then
                ' See if the path exists in the list of source files
                If Not dAllFiles.Exists(cItem.SourceFile) Then
                    ' This object does not appear to exist in the source files.
                    ' Add to list of files to process. (This will remove the
                    ' existing database object during the merge.)
                    .Add cItem.SourceFile, vbNullString
                End If
            End If
        Next varItem
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Conflicts
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Expose conflicts as an extension of the index
'---------------------------------------------------------------------------------------
'
Public Function Conflicts() As clsConflicts
    If m_Conflicts Is Nothing Then Set m_Conflicts = New clsConflicts
    Set Conflicts = m_Conflicts
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckExportConflicts
' Author    : Adam Waller
' Date      : 9/15/2021
' Purpose   : Check for potential conflicts when exporting these items
'---------------------------------------------------------------------------------------
'
Public Sub CheckExportConflicts(dDbComponents As Dictionary)

    Dim varKey As Variant
    Dim cItem As IDbComponent
    Dim sngStart As Single
    Dim colFiles As Collection

    ' Capture start time so we can display some progress if this
    ' scan takes a while.
    sngStart = Timer

    ' Loop through all components
    For Each varKey In dDbComponents.Keys
        Set cItem = dDbComponents(varKey)
        Set colFiles = GetExportConflictFiles(cItem)
        If Not colFiles Is Nothing Then
            ' Add to list of conflicts with per-file detail
            Me.Conflicts.Add cItem, _
                CStr(varKey), _
                Me.Item(cItem).SyncDate, _
                GetSourceModifiedDate(cItem, cItem.SourceFile), _
                ercOverwrite, , , colFiles
        End If
        ' Increment the progress bar
        Log.Increment
        ' Exit loop if single source file
        If cItem.SingleFile Then Exit For
        ' Display more detail if export runs more than 1 second.
        If sngStart > 0 And (Timer - sngStart > 1) Then
            Log.Add " - Checking " & dDbComponents.Count & " " & LCase(cItem.Category) & "...", , , , , True
            sngStart = 0
        End If
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetExportConflictFiles
' Author    : Adam Waller
' Date      : 9/15/2021
' Purpose   : Returns a Collection of file names that differ between the source folder
'           : and a fresh temp export, or Nothing if there is no conflict. Used by
'           : CheckExportConflicts to populate per-file conflict rows.
'---------------------------------------------------------------------------------------
'
Private Function GetExportConflictFiles(cItem As IDbComponent) As Collection

    Dim strHash As String
    Dim strFile As String
    Dim strTempFile As String
    Dim blnCompareTemp As Boolean

    ' Check for special cases based on component type
    If cItem.ComponentType = edbTableData Then
        ' Table data is not considered an export conflict, and should
        ' always be overwritten.
        Exit Function
    End If

    ' See if the export file even exists before we worry about conflicts.
    strFile = cItem.SourceFile
    If FSO.FileExists(strFile) Then

        ' Get a hash of the file properties
        strHash = GetSourceFilesPropertyHash(cItem)

        ' Check to see if this object is in the index
        ' (Has been exported/imported before.)
        If Me.Exists(cItem, strFile) Then

            ' Item found in index.
            ' Check to see if it matches the property hash in the index.
            If Me.Item(cItem).FilePropertiesHash = strHash Then
                ' Source file has not changed since last export/import.
                ' We should be good to overwrite it.
            Else
                ' Doesn't match. This could be a conflict, depending on the dates.
                ' If the source file has been modified since we last synced with
                ' it (either by exporting to it or importing from it)...
                If Me.Item(cItem).SyncDate < GetSourceModifiedDate(cItem) Then
                    ' ...and the content is different, then we have a conflict.
                    blnCompareTemp = True
                End If
            End If
        Else
            ' Without an existing index entry, we will need to compare a fresh
            ' export of the item to the source file and check for differing content.
            blnCompareTemp = True
        End If
    End If

    ' Compare files to identify which ones actually differ.
    If blnCompareTemp Then
        strTempFile = Replace(strFile, Options.GetExportFolder, GetTempExportFolder)
        Log.Add "  Exporting " & cItem.Name, False
        cItem.Export strTempFile
        Set GetExportConflictFiles = GetDifferingFiles(cItem, strFile, strTempFile)
        If GetExportConflictFiles Is Nothing Then
            ' No actual conflict -- refresh stale property hash so this
            ' file doesn't trigger an expensive temp export next time.
            Me.Item(cItem).FilePropertiesHash = strHash
            Me.Item(cItem).SourceModified = GetSourceModifiedDate(cItem)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckImportConflicts
' Author    : Adam Waller
' Date      : 2/24/2023
' Purpose   : Check for potential conflicts when merging into this database.
'---------------------------------------------------------------------------------------
'
Public Sub CheckMergeConflicts(cComponent As IDbComponent, dFiles As Dictionary)

    Dim varKey As Variant
    Dim cItem As IDbComponent
    Dim sngStart As Single
    Dim strFile As String
    Dim intConflict As eResolveConflict
    Dim colFiles As Collection

    ' Capture start time so we can display some progress if this
    ' scan takes a while.
    sngStart = Timer
    Perf.OperationStart "Check Merge Conflicts"

    ' Loop through the list of source files
    ' (Each key is a full path to a source file)
    For Each varKey In dFiles.Keys
        strFile = varKey
        Set colFiles = Nothing
        intConflict = IsMergeConflict(cComponent, strFile, colFiles)
        If intConflict Then
            ' Get database item
            If cComponent.GetAllFromDB.Exists(strFile) Then
                ' Get individual item (form, report)
                Set cItem = cComponent.GetAllFromDB.Item(strFile)
            Else
                ' Get parent class for single file items
                Set cItem = cComponent
            End If
            ' Add to list of conflicts with per-file detail
            Me.Conflicts.Add cItem, _
                strFile, _
                cItem.DateModified, _
                GetSourceModifiedDate(cComponent, strFile), _
                intConflict, _
                strFile, _
                ercSkip, colFiles
        End If
        ' Increment the progress bar
        Log.Increment
        ' Exit loop if single source file
        If cComponent.SingleFile Then Exit For
        ' Display more detail if export runs more than 1 second.
        If sngStart > 0 And (Timer - sngStart > 1) Then
            Log.Add " - Checking " & dFiles.Count & " " & LCase(cComponent.Category) & "...", , , , , True
            sngStart = 0
        End If
    Next varKey
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsMergeConflict
' Author    : Adam Waller
' Date      : 12/5/2023
' Purpose   : Returns the conflict type if we are about to overwrite changes in the
'           : database that don't match the current source file. When the conflict
'           : involves differing files, colDiffFiles is set to the list of file names.
'---------------------------------------------------------------------------------------
'
Private Function IsMergeConflict(cItemClass As IDbComponent, strSourceFile As String, _
    Optional ByRef colDiffFiles As Collection) As eResolveConflict

    Dim cItem As IDbComponent
    Dim strTempFile As String

    ' Check for special cases based on component type
    If cItemClass.ComponentType = edbTableData Then
        ' Table data is not considered an import conflict, and should
        ' always be overwritten.
        IsMergeConflict = ercNone
        Exit Function
    End If

    ' Single file objects use the same temp-export + compare approach as
    ' multi-file components, ensuring the temp file exists for diffing and
    ' eliminating false-positive conflicts from stale index hashes.
    If cItemClass.SingleFile Then
        If cItemClass.IsModified Then
            If Not FSO.FileExists(strSourceFile) Then
                ' Source file does not exist. For single-file components this
                ' means the file was never exported or was removed during a
                ' format migration — not a conflict to resolve.
            Else
                strTempFile = Replace(strSourceFile, Options.GetExportFolder, GetTempExportFolder)
                Log.Add "  Exporting " & cItemClass.Name, False
                cItemClass.Export strTempFile
                ' When content is empty (e.g. hidden-attributes with EFV >= 5.0.0),
                ' the export produces no file. Create an empty placeholder so the
                ' diff tool has something to compare against.
                If Not FSO.FileExists(strTempFile) Then
                    VerifyPath strTempFile
                    WriteFile vbNullString, strTempFile
                End If
                Set colDiffFiles = GetDifferingFiles(cItemClass, strSourceFile, strTempFile)
                If Not colDiffFiles Is Nothing Then
                    IsMergeConflict = ercOverwrite
                End If
            End If
        End If
    Else
        ' See if the database object even exists before we worry about conflicts.
        If cItemClass.GetAllFromDB.Exists(strSourceFile) Then

            ' Get class instance for database item
            Set cItem = cItemClass.GetAllFromDB.Item(strSourceFile)

            ' See if this object appears to be modified since the last import/export
            If cItem.IsModified Then
                If Not FSO.FileExists(strSourceFile) Then
                    ' Modified object, and source file does not exist. Could be a case
                    ' where we are removing an orphaned database object.
                    IsMergeConflict = ercDelete
                Else
                    ' Just because the modification date is different doesn't always
                    ' mean the object definition has actually changed. By comparing
                    ' a fresh export of the object, we can determine for sure if the
                    ' source files match.
                    strTempFile = Replace(strSourceFile, Options.GetExportFolder, GetTempExportFolder)
                    Log.Add "  Exporting " & cItem.Name, False
                    cItem.Export strTempFile
                    ' Because we are using a newly exported temp file, the file properties
                    ' hash won't match, since the file dates are going to be different.
                    Set colDiffFiles = GetDifferingFiles(cItemClass, strSourceFile, strTempFile)
                    If Not colDiffFiles Is Nothing Then
                        IsMergeConflict = ercOverwrite
                    End If
                End If
            Else
                ' Based on the index, it does not appear that this item
                ' has changed since the last export or import operation.
                ' We should be good to overwrite (or skip importing) this
                ' object in a merge operation.
            End If
        Else
            ' Database object does not exist. No conflict to merge this file.
        End If

    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDifferingFiles
' Author    : Adam Waller
' Date      : 5/5/2026
' Purpose   : Returns a Collection of relative file names (e.g. "MyTable.xml") that
'           : differ between two source paths. Returns Nothing when all files match.
'           : Files present on only one side are included as differences.
'---------------------------------------------------------------------------------------
'
Private Function GetDifferingFiles(cType As IDbComponent, strFile1 As String, strFile2 As String) As Collection

    Dim dPaths(1 To 2) As Dictionary
    Dim strFiles(1 To 2) As String
    Dim intCnt As Integer
    Dim varExt As Variant
    Dim strPath As String
    Dim varKey As Variant
    Dim colResult As Collection

    strFiles(1) = strFile1
    strFiles(2) = strFile2

    ' Start by retrieving the paths of existing files for this component.
    For intCnt = 1 To 2
        Set dPaths(intCnt) = New Dictionary
        For Each varExt In cType.FileExtensions
            strPath = SwapExtension(strFiles(intCnt), CStr(varExt))
            If FSO.FileExists(strPath) Then
                dPaths(intCnt).Add FSO.GetFileName(strPath), strPath
            End If
        Next varExt
    Next intCnt

    Set colResult = New Collection

    ' Files only in path 1 (source) but not path 2 (temp)
    For Each varKey In dPaths(1).Keys
        If Not dPaths(2).Exists(varKey) Then colResult.Add CStr(varKey)
    Next varKey

    ' Files only in path 2 (temp) but not path 1 (source)
    For Each varKey In dPaths(2).Keys
        If Not dPaths(1).Exists(varKey) Then colResult.Add CStr(varKey)
    Next varKey

    ' Files in both -- compare sizes, then hashes
    For Each varKey In dPaths(1).Keys
        If dPaths(2).Exists(varKey) Then
            If FSO.GetFile(CStr(dPaths(1).Item(varKey))).Size _
            <> FSO.GetFile(CStr(dPaths(2).Item(varKey))).Size Then
                colResult.Add CStr(varKey)
            ElseIf GetFileHash(CStr(dPaths(1).Item(varKey))) _
            <> GetFileHash(CStr(dPaths(2).Item(varKey))) Then
                colResult.Add CStr(varKey)
            End If
        End If
    Next varKey

    ' Return Nothing when everything matches
    If colResult.Count > 0 Then Set GetDifferingFiles = colResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTempExportFolder
' Author    : Adam Waller
' Date      : 8/30/2022
' Purpose   : Returns the temp folder used for temporary source exports when the
'           : modified status of the the component cannot be determined without running
'           : an export operation. Persisting this folder during the export operation
'           : allows us to perform a single export of the component while still giving
'           : the user the ability to decide whether to overwrite the existing file.
'           : NOTE: If the folder does not exist, a new one will be created.
'           : This should be cleared after the export operation using the corresponding
'           : ClearTempExportFolder sub.
'---------------------------------------------------------------------------------------
'
Public Property Get GetTempExportFolder() As String
    If m_strTempExportFolderPath = vbNullString Then m_strTempExportFolderPath = GetTempFolder("VCS")
    GetTempExportFolder = m_strTempExportFolderPath & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : ClearTempExportFolder
' Author    : Adam Waller
' Date      : 8/30/2022
' Purpose   : Removes the temporary export folder, including any source files exported
'           : into this folder while scanning for changes.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTempExportFolder()
    If m_strTempExportFolderPath <> vbNullString Then
        If FSO.FolderExists(m_strTempExportFolderPath) Then
            LogUnhandledErrors
            On Error Resume Next
            ' Use FSO to delete the folder and contents
            FSO.DeleteFolder m_strTempExportFolderPath, True
            CatchAny eelWarning, "Unable to delete temp folder: '" & m_strTempExportFolderPath & _
                "' You will need to manually remove this folder.", ModuleName(Me) & ".ClearTempExportFolder"
        End If
        m_strTempExportFolderPath = vbNullString
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DumpToJson
' Author    : Adam Waller
' Date      : 4/28/2026
' Purpose   : Writes the in-memory index to a human-readable JSON file for
'           : troubleshooting. Since per-entry storage uses clsVCSIndexItem
'           : (not Dictionary), this builds a temporary Dictionary tree and
'           : feeds it through the existing JSON serializer.
'---------------------------------------------------------------------------------------
'
Public Sub DumpToJson(Optional strPath As String)

    Dim dDump As Dictionary
    Dim dComponents As Dictionary
    Dim dCatOut As Dictionary
    Dim dEntryOut As Dictionary
    Dim varCat As Variant
    Dim varFile As Variant
    Dim cEntry As clsVCSIndexItem
    Dim blnSavedValue As Boolean

    ' Default output path alongside the .idx file
    If strPath = vbNullString Then
        If m_strFile = vbNullString Then m_strFile = DefaultFilePath
        If m_strFile = vbNullString Then Exit Sub
        strPath = Replace(m_strFile, cstrFileName, "vcs-index-debug.json")
    End If

    ' Build a Dictionary tree that mirrors the old JSON structure
    Set dDump = New Dictionary
    dDump.Add "MergeBuildDate", ZNDate(Me.MergeBuildDate)
    dDump.Add "FullBuildDate", ZNDate(Me.FullBuildDate)
    dDump.Add "ExportDate", ZNDate(Me.ExportDate)
    dDump.Add "FullExportDate", ZNDate(Me.FullExportDate)
    dDump.Add "LastMergedCommit", Me.LastMergedCommit
    dDump.Add "VBAProjectDate", ZNDate(Me.VBAProjectDate)
    Set dDump("CategoryHashes") = m_dIndex("CategoryHashes")

    ' Convert clsVCSIndexItem entries to Dictionaries
    Set dComponents = New Dictionary
    If m_dIndex.Exists("Components") Then
        For Each varCat In m_dIndex("Components").Keys
            Set dCatOut = New Dictionary
            For Each varFile In m_dIndex("Components")(varCat).Keys
                Set cEntry = m_dIndex("Components")(varCat)(varFile)
                Set dEntryOut = New Dictionary
                If Len(cEntry.FileHash) > 0 Then dEntryOut.Add "FileHash", cEntry.FileHash
                If cEntry.ExportDate <> 0 Then dEntryOut.Add "ExportDate", cEntry.ExportDate
                If cEntry.ImportDate <> 0 Then dEntryOut.Add "ImportDate", cEntry.ImportDate
                dEntryOut.Add "SourceModified", ZNDate(cEntry.SourceModified)
                If Len(cEntry.FilePropertiesHash) > 0 Then dEntryOut.Add "FilePropertiesHash", cEntry.FilePropertiesHash
                If Len(cEntry.OtherHash) > 0 Then dEntryOut.Add "OtherHash", cEntry.OtherHash
                If Len(cEntry.MetaHash) > 0 Then dEntryOut.Add "MetaHash", cEntry.MetaHash
                If cEntry.ObjectDate <> 0 Then dEntryOut.Add "ObjectDate", cEntry.ObjectDate
                If Len(cEntry.FolderAnnotation) > 0 Then dEntryOut.Add "FolderAnnotation", cEntry.FolderAnnotation
                dCatOut.Add CStr(varFile), dEntryOut
            Next varFile
            dComponents.Add CStr(varCat), dCatOut
        Next varCat
    End If
    Set dDump("Components") = dComponents

    ' Serialize to JSON with ISO date conversion
    blnSavedValue = JsonOptions.ConvertDateToIso
    JsonOptions.ConvertDateToIso = True
    WriteFile BuildJsonFile(TypeName(Me), dDump, "Version Control System Index (Debug Dump)"), strPath
    JsonOptions.ConvertDateToIso = blnSavedValue

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CategoryHashes
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Per-category option hashes for smart re-export detection.
'---------------------------------------------------------------------------------------
'
Public Property Get CategoryHashes() As Dictionary
    Set CategoryHashes = m_dIndex("CategoryHashes")
End Property

Public Property Set CategoryHashes(dHashes As Dictionary)
    Set m_dIndex("CategoryHashes") = dHashes
End Property


'---------------------------------------------------------------------------------------
' Procedure : Git
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Reference to the Git integration class
'---------------------------------------------------------------------------------------
'
Private Function Git() As clsGitIntegration
    If m_Git Is Nothing Then Set m_Git = New clsGitIntegration
    Set Git = m_Git
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function DefaultFilePath() As String
    If DatabaseFileOpen Then DefaultFilePath = Options.GetExportFolder & cstrFileName
End Function


'---------------------------------------------------------------------------------------
' Binary I/O Helpers
' Used for reading and writing the binary index format (.idx)
'---------------------------------------------------------------------------------------

Private Sub InitWriteBuffer(Optional lngInitialSize As Long = 65536)
    m_lngBufSize = lngInitialSize
    ReDim m_bBuffer(0 To m_lngBufSize - 1)
    m_lngBufPos = 0
End Sub

Private Sub EnsureBuffer(lngNeeded As Long)
    If m_lngBufPos + lngNeeded > m_lngBufSize Then
        Do While m_lngBufPos + lngNeeded > m_lngBufSize
            m_lngBufSize = m_lngBufSize * 2
        Loop
        ReDim Preserve m_bBuffer(0 To m_lngBufSize - 1)
    End If
End Sub

Private Function GetWriteBuffer() As Byte()
    If m_lngBufPos > 0 Then
        ReDim Preserve m_bBuffer(0 To m_lngBufPos - 1)
    End If
    GetWriteBuffer = m_bBuffer
    Erase m_bBuffer
    m_lngBufPos = 0
    m_lngBufSize = 0
End Function

Private Function ReadUInt32(bData() As Byte, ByRef lngPos As Long) As Long
    EnsureReadable bData, lngPos, 4
    ReadUInt32 = bData(lngPos) + _
                 bData(lngPos + 1) * 256& + _
                 CLng(bData(lngPos + 2)) * 65536 + _
                 CLng(bData(lngPos + 3)) * 16777216
    lngPos = lngPos + 4
End Function

Private Function ReadUInt16(bData() As Byte, ByRef lngPos As Long) As Long
    EnsureReadable bData, lngPos, 2
    ReadUInt16 = bData(lngPos) + bData(lngPos + 1) * 256&
    lngPos = lngPos + 2
End Function

Private Sub WriteUInt32(lngValue As Long)
    EnsureBuffer 4
    m_bBuffer(m_lngBufPos) = lngValue And &HFF&
    m_bBuffer(m_lngBufPos + 1) = (lngValue \ 256&) And &HFF&
    m_bBuffer(m_lngBufPos + 2) = (lngValue \ 65536) And &HFF&
    m_bBuffer(m_lngBufPos + 3) = (lngValue \ 16777216) And &HFF&
    m_lngBufPos = m_lngBufPos + 4
End Sub

Private Sub WriteUInt16(lngValue As Long)
    If lngValue < 0 Or lngValue > 65535 Then
        Err.Raise 5, ModuleName(Me) & ".WriteUInt16", "Value out of range for UInt16."
    End If
    EnsureBuffer 2
    m_bBuffer(m_lngBufPos) = lngValue And &HFF&
    m_bBuffer(m_lngBufPos + 1) = (lngValue \ 256&) And &HFF&
    m_lngBufPos = m_lngBufPos + 2
End Sub

Private Function ReadDouble(bData() As Byte, ByRef lngPos As Long) As Double
    Dim tB As tDoubleBytes
    Dim tV As tDoubleVal
    Dim i As Long
    EnsureReadable bData, lngPos, 8
    For i = 0 To 7
        tB.b(i) = bData(lngPos + i)
    Next i
    LSet tV = tB
    ReadDouble = tV.Value
    lngPos = lngPos + 8
End Function

Private Sub WriteDouble(dblValue As Double)
    Dim tB As tDoubleBytes
    Dim tV As tDoubleVal
    tV.Value = dblValue
    LSet tB = tV
    EnsureBuffer 8
    Dim i As Long
    For i = 0 To 7
        m_bBuffer(m_lngBufPos + i) = tB.b(i)
    Next i
    m_lngBufPos = m_lngBufPos + 8
End Sub

Private Function ReadUtcDate(bData() As Byte, ByRef lngPos As Long) As Date
    Dim dblValue As Double
    dblValue = ReadDouble(bData, lngPos)
    If dblValue <> 0# Then ReadUtcDate = ConvertToLocalDate(CDate(dblValue))
End Function

Private Sub WriteUtcDate(dteValue As Date)
    If dteValue = 0 Then
        WriteDouble 0#
    Else
        WriteDouble CDbl(ConvertToUtc(dteValue))
    End If
End Sub

Private Function ReadLPS(bData() As Byte, ByRef lngPos As Long) As String
    Dim lngLen As Long
    Dim bStr() As Byte
    Dim i As Long
    lngLen = ReadUInt32(bData, lngPos)
    If lngLen = 0 Then
        ReadLPS = vbNullString
    Else
        If (lngLen And 1) <> 0 Then Err.Raise 5, ModuleName(Me) & ".ReadLPS", "Invalid UTF-16 byte length in index file."
        EnsureReadable bData, lngPos, lngLen
        ReDim bStr(0 To lngLen - 1)
        For i = 0 To lngLen - 1
            bStr(i) = bData(lngPos + i)
        Next i
        ReadLPS = bStr
        lngPos = lngPos + lngLen
    End If
End Function

Private Sub WriteLPS(strValue As String)
    Dim bStr() As Byte
    Dim lngLen As Long
    Dim i As Long
    If Len(strValue) = 0 Then
        WriteUInt32 0
    Else
        bStr = strValue
        lngLen = UBound(bStr) + 1
        WriteUInt32 lngLen
        EnsureBuffer lngLen
        For i = 0 To lngLen - 1
            m_bBuffer(m_lngBufPos + i) = bStr(i)
        Next i
        m_lngBufPos = m_lngBufPos + lngLen
    End If
End Sub

Private Function ReadByte(bData() As Byte, ByRef lngPos As Long) As Byte
    EnsureReadable bData, lngPos, 1
    ReadByte = bData(lngPos)
    lngPos = lngPos + 1
End Function

Private Sub WriteByte(bValue As Byte)
    EnsureBuffer 1
    m_bBuffer(m_lngBufPos) = bValue
    m_lngBufPos = m_lngBufPos + 1
End Sub

Private Sub WriteMagic(strMagic As String)
    Dim intCnt As Integer

    For intCnt = 1 To Len(strMagic)
        WriteByte CByte(Asc(Mid$(strMagic, intCnt, 1)))
    Next intCnt
End Sub

Private Sub EnsureReadable(bData() As Byte, ByVal lngPos As Long, ByVal lngCount As Long)
    Dim lngMax As Long

    If lngCount < 0 Then
        Err.Raise 5, ModuleName(Me) & ".EnsureReadable", "Negative read length in index file."
    End If
    If lngCount = 0 Then Exit Sub
    If lngPos < 0 Then
        Err.Raise 5, ModuleName(Me) & ".EnsureReadable", "Negative read position in index file."
    End If

    lngMax = UBound(bData)
    If lngPos > lngMax Or (lngPos + lngCount - 1) > lngMax Then
        Err.Raise 5, ModuleName(Me) & ".EnsureReadable", "Truncated index file."
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Set up the dictionary object and keys for reflection.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set m_dIndex = New Dictionary
    With m_dIndex
        .Add "MergeBuildDate", Null
        .Add "FullBuildDate", Null
        .Add "ExportDate", Null
        .Add "FullExportDate", Null
        .Add "VBAProjectDate", Null
        .Add "LastMergedCommit", vbNullString
        Set .Item("CategoryHashes") = New Dictionary
        Set .Item("Components") = New Dictionary
        Set .Item("AlternateExport") = New Dictionary
    End With

    ' Load Git integration
    Set m_dGitIndex = Nothing

End Sub
