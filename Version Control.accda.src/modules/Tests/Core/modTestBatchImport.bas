Attribute VB_Name = "modTestBatchImport"
'---------------------------------------------------------------------------------------
' Module    : modTestBatchImport
' Author    : Adam Waller
' Date      : 9/15/2026
' Purpose   : Regression coverage for full-build two-pass component imports and the
'           : immediate metadata path retained by ordinary Import/Merge calls.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")


Public Sub TestMetadataComponentsImplementBatchImport()

    Dim colComponents As New Collection
    Dim cComponent As IDbComponent

    colComponents.Add New clsDbModule
    colComponents.Add New clsDbTableDef
    colComponents.Add New clsDbQuery
    colComponents.Add New clsDbForm
    colComponents.Add New clsDbMacro
    colComponents.Add New clsDbReport

    For Each cComponent In colComponents
        TestAssert TypeOf cComponent Is IDbBatchImport, _
            TypeName(cComponent) & " implements IDbBatchImport"
    Next cComponent

End Sub


Public Sub TestQueryBatchImport_AppliesMetadataAndIndexesEachFile()
    '@Tag("integration")

    Const strQuery1 As String = "vcs_test_batch_alpha"
    Const strQuery2 As String = "vcs_test_batch_beta"
    Const strDescription1 As String = "Batch metadata alpha"
    Const strDescription2 As String = "Batch metadata beta"
    Const strCustomValue As String = "Custom batch value"

    Dim cComponent As IDbComponent
    Dim cBatch As IDbBatchImport
    Dim strRoot As String
    Dim strFile1 As String
    Dim strFile2 As String
    Dim strSavedExport As String
    Dim lngSavedFormat As Long
    Dim blnSavedDeterministic As Boolean
    Dim cSavedIndex As clsVCSIndex
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    BeginQuerySandbox strRoot, strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    strFile1 = strRoot & "queries" & PathSep & strQuery1 & ".sql"
    strFile2 = strRoot & "queries" & PathSep & strQuery2 & ".sql"
    WriteFile "SELECT 1 AS One;", strFile1
    WriteFile "SELECT 2 AS Two;", strFile2
    WriteDescriptionSidecar SwapExtension(strFile1, "json"), strDescription1, True, strCustomValue
    WriteDescriptionSidecar SwapExtension(strFile2, "json"), strDescription2

    DeleteObjectIfExists acQuery, strQuery1
    DeleteObjectIfExists acQuery, strQuery2

    Set cComponent = New clsDbQuery
    Set cBatch = cComponent
    cBatch.ImportFast strFile1
    cBatch.ImportFast strFile2
    cBatch.FinalizeImports

    TestAssert QueryDescription(strQuery1) = strDescription1, _
        "first batch-imported query metadata applied"
    TestAssert QueryDescription(strQuery2) = strDescription2, _
        "second batch-imported query metadata applied"
    TestAssert QueryProperty(strQuery1, "BatchMarker") = strCustomValue, _
        "custom DAO property applied after the batch refresh"
    TestAssert Application.GetHiddenAttribute(acQuery, strQuery1), _
        "hidden attribute applied during batch finalization"
    TestAssert VCSIndex.Exists(cComponent, strFile1), _
        "first batch-imported query indexed under its own file"
    TestAssert VCSIndex.Exists(cComponent, strFile2), _
        "second batch-imported query indexed under its own file"
    TestAssert VCSIndex.Item(cComponent, strFile1).MetaHash = _
        GetMetadataHash("Tables", strQuery1, acQuery), _
        "first batch-imported query metadata hash indexed"
    TestAssert VCSIndex.Item(cComponent, strFile2).MetaHash = _
        GetMetadataHash("Tables", strQuery2, acQuery), _
        "second batch-imported query metadata hash indexed"

CleanUp:
    On Error Resume Next
    DeleteObjectIfExists acQuery, strQuery1
    DeleteObjectIfExists acQuery, strQuery2
    If Not cComponent Is Nothing Then
        If Len(strFile1) > 0 Then VCSIndex.Remove cComponent, strFile1
        If Len(strFile2) > 0 Then VCSIndex.Remove cComponent, strFile2
    End If
    RestoreQuerySandbox strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    If lngErr <> 0 Then TestAssert False, _
        "unexpected batch import error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestQueryBatchImport_RetriesDeferredQuery()
    '@Tag("integration")

    Const strCustomers As String = "vcs_test_batch_customers"
    Const strProducts As String = "vcs_test_batch_products"
    Const strEarly As String = "vcs_test_batch_early"
    Const strConsumer As String = "vcs_test_batch_consumer"
    Const strLate As String = "vcs_test_batch_late"
    Const strDescription As String = "Deferred batch query"

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim cComponent As IDbComponent
    Dim cBatch As IDbBatchImport
    Dim strRoot As String
    Dim strConsumerFile As String
    Dim strLateFile As String
    Dim strSavedExport As String
    Dim lngSavedFormat As Long
    Dim blnSavedDeterministic As Boolean
    Dim cSavedIndex As clsVCSIndex
    Dim eelSavedLevel As eErrorLevel
    Dim lngProbeErr As Long
    Dim strProbeErr As String
    Dim intLockedFile As Integer
    Dim blnFileLocked As Boolean
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    BeginQuerySandbox strRoot, strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    strConsumerFile = strRoot & "queries" & PathSep & strConsumer & ".sql"
    strLateFile = strRoot & "queries" & PathSep & strLate & ".sql"

    DeleteObjectIfExists acQuery, strConsumer
    DeleteObjectIfExists acQuery, strLate
    DeleteObjectIfExists acQuery, strEarly
    DeleteObjectIfExists acTable, strProducts
    DeleteObjectIfExists acTable, strCustomers

    Set dbs = SharedDb
    dbs.Execute "CREATE TABLE " & strCustomers & " (CustomerId LONG, DisplayText TEXT(20))", dbFailOnError
    dbs.Execute "CREATE TABLE " & strProducts & " (ProductCode LONG, ProductText TEXT(20))", dbFailOnError
    dbs.Execute "INSERT INTO " & strCustomers & " (CustomerId, DisplayText) VALUES (1, 'One')", dbFailOnError
    dbs.Execute "INSERT INTO " & strProducts & " (ProductCode, ProductText) VALUES (1, 'One')", dbFailOnError
    Set qdf = dbs.CreateQueryDef(strEarly, _
        "SELECT CustomerId, DisplayText FROM " & strCustomers & ";")
    dbs.TableDefs.Refresh
    dbs.QueryDefs.Refresh
    RefreshContainerDocuments "Tables"

    WriteFile "SELECT DISTINCTROW" & vbCrLf & _
        "  " & strEarly & ".CustomerId," & vbCrLf & _
        "  " & strEarly & ".DisplayText" & vbCrLf & _
        "FROM" & vbCrLf & _
        "  " & strEarly & vbCrLf & _
        "  INNER JOIN " & strLate & " ON (" & strEarly & ".CustomerId = " & strLate & ".ProductCode)" & vbCrLf & _
        "    AND (" & strEarly & ".DisplayText = " & strLate & ".ProductText);", _
        strConsumerFile
    WriteDescriptionSidecar SwapExtension(strConsumerFile, "json"), strDescription
    WriteFile "SELECT ProductCode, ProductText FROM " & strProducts & ";", strLateFile
    WriteDescriptionSidecar SwapExtension(strLateFile, "json"), "Late batch dependency"

    Set qdf = dbs.CreateQueryDef(strLate, ReadFile(strLateFile))
    Set qdf = Nothing

    ' A rejected SQL string here is the thing under test, not a test failure, so this
    ' one statement runs untrapped and hands the handler back immediately after.
    On Error Resume Next
    Set qdf = dbs.CreateQueryDef(strConsumer, ReadFile(strConsumerFile))
    lngProbeErr = Err.Number
    strProbeErr = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    TestAssert lngProbeErr = 0, _
        "sanitized consumer SQL is valid: " & lngProbeErr & " " & strProbeErr
    Set qdf = Nothing
    DeleteObjectIfExists acQuery, strConsumer
    DeleteObjectIfExists acQuery, strLate
    dbs.QueryDefs.Refresh
    RefreshContainerDocuments "Tables"

    ' This Access database accepts the sanitized missing-query reference, while the
    ' production shape from #783 rejects it. Lock the unchanged, valid source for the
    ' first attempt to exercise the same transient ImportFast failure boundary.
    intLockedFile = FreeFile
    Open strConsumerFile For Binary Access Read Write Lock Read Write As #intLockedFile
    blnFileLocked = True
    Set cComponent = New clsDbQuery
    Set cBatch = cComponent
    cBatch.ImportFast strConsumerFile
    Close #intLockedFile
    blnFileLocked = False
    TestAssert Not QueryDefExists(strConsumer), "transient first-pass failure is deferred"
    cBatch.ImportFast strLateFile
    cBatch.FinalizeImports

    TestAssert QueryDefExists(strLate), "late dependency imported"
    TestAssert QueryDefExists(strConsumer), "deferred consumer imported"
    TestAssert QueryDocumentExists(strConsumer), "deferred consumer published as a document"
    If QueryDocumentExists(strConsumer) Then
        TestAssert QueryDescription(strConsumer) = strDescription, _
            "deferred consumer metadata applied"
    End If
    TestAssert VCSIndex.Exists(cComponent, strConsumerFile), _
        "deferred consumer indexed"
    TestAssert VCSIndex.Exists(cComponent, strLateFile), _
        "late dependency indexed"

CleanUp:
    On Error Resume Next
    If blnFileLocked Then Close #intLockedFile
    Set qdf = Nothing
    Set dbs = Nothing
    DeleteObjectIfExists acQuery, strConsumer
    DeleteObjectIfExists acQuery, strLate
    DeleteObjectIfExists acQuery, strEarly
    DeleteObjectIfExists acTable, strProducts
    DeleteObjectIfExists acTable, strCustomers
    If Not cComponent Is Nothing Then
        If Len(strConsumerFile) > 0 Then VCSIndex.Remove cComponent, strConsumerFile
        If Len(strLateFile) > 0 Then VCSIndex.Remove cComponent, strLateFile
    End If
    RestoreQuerySandbox strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    If lngErr <> 0 Then TestAssert False, _
        "unexpected deferred import error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestQueryMerge_AppliesMetadataImmediately()
    '@Tag("integration")

    Const strQuery As String = "vcs_test_merge_metadata"
    Const strDescription As String = "Immediate merge metadata"

    Dim cComponent As IDbComponent
    Dim strRoot As String
    Dim strFile As String
    Dim strSavedExport As String
    Dim lngSavedFormat As Long
    Dim blnSavedDeterministic As Boolean
    Dim cSavedIndex As clsVCSIndex
    Dim eelSavedLevel As eErrorLevel
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    BeginQuerySandbox strRoot, strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    strFile = strRoot & "queries" & PathSep & strQuery & ".sql"
    WriteFile "SELECT 3 AS Three;", strFile
    WriteDescriptionSidecar SwapExtension(strFile, "json"), strDescription

    DeleteObjectIfExists acQuery, strQuery
    Set cComponent = New clsDbQuery
    cComponent.Merge strFile

    TestAssert QueryDescription(strQuery) = strDescription, _
        "merge applies metadata without batch finalization"
    TestAssert VCSIndex.Exists(cComponent, strFile), _
        "merge indexes the metadata-bearing query immediately"

CleanUp:
    On Error Resume Next
    DeleteObjectIfExists acQuery, strQuery
    If Not cComponent Is Nothing Then
        If Len(strFile) > 0 Then VCSIndex.Remove cComponent, strFile
    End If
    RestoreQuerySandbox strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex, _
        eelSavedLevel
    If lngErr <> 0 Then TestAssert False, _
        "unexpected query merge error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


' Start from the state an operation starts from. No operation begins in this project
' during a test run, so an earlier test can leave a SharedDb handle that predates the
' queries created here, or an eelCritical level that makes LoadComponentFromText
' report failure after a successful load.
Private Sub BeginQuerySandbox(ByRef strRoot As String, ByRef strSavedExport As String, _
    ByRef lngSavedFormat As Long, ByRef blnSavedDeterministic As Boolean, _
    ByRef cSavedIndex As clsVCSIndex, ByRef eelSavedLevel As eErrorLevel)

    Set cSavedIndex = VCSIndex
    strSavedExport = Options.ExportFolder
    lngSavedFormat = Options.ExportFormatVersion
    blnSavedDeterministic = Options.UseDeterministicQueryExport
    eelSavedLevel = Operation.ErrorLevel
    Operation.ErrorLevel = eelNoError
    ReleaseDbReferences

    strRoot = GetTempFolder("vcs_batch_import") & PathSep
    VerifyPath strRoot & "queries" & PathSep
    Options.ExportFolder = strRoot
    Options.ExportFormatVersion = EFV_5_1_0
    Options.UseDeterministicQueryExport = True
    Set VCSIndex = Nothing

End Sub


Private Sub RestoreQuerySandbox(strSavedExport As String, lngSavedFormat As Long, _
    blnSavedDeterministic As Boolean, cSavedIndex As clsVCSIndex, eelSavedLevel As eErrorLevel)

    Options.ExportFolder = strSavedExport
    Options.ExportFormatVersion = lngSavedFormat
    Options.UseDeterministicQueryExport = blnSavedDeterministic
    Set VCSIndex = cSavedIndex
    Operation.ErrorLevel = eelSavedLevel
    ReleaseDbReferences

End Sub


Private Sub WriteDescriptionSidecar(strFile As String, strDescription As String, _
    Optional blnHidden As Boolean = False, Optional strCustomValue As String = vbNullString)

    Dim dFile As New Dictionary
    Dim dItems As New Dictionary
    Dim dProperties As New Dictionary
    Dim dDescription As New Dictionary

    dDescription.Add "Type", dbText
    dDescription.Add "Value", strDescription
    dProperties.Add "Description", dDescription
    If Len(strCustomValue) > 0 Then
        Set dDescription = New Dictionary
        dDescription.Add "Type", dbText
        dDescription.Add "Value", strCustomValue
        dProperties.Add "BatchMarker", dDescription
    End If
    dItems.Add "Properties", dProperties
    If blnHidden Then dItems.Add "Hidden", True
    dFile.Add "Items", dItems

    WriteFile ConvertToJson(dFile, JSON_WHITESPACE), strFile

End Sub


Private Function QueryDefExists(strQueryName As String) As Boolean

    Dim qdf As DAO.QueryDef

    SharedDb.QueryDefs.Refresh
    On Error Resume Next
    Set qdf = SharedDb.QueryDefs(strQueryName)
    QueryDefExists = (Err.Number = 0)
    Err.Clear

End Function


Private Function QueryDocumentExists(strQueryName As String) As Boolean

    Dim doc As DAO.Document

    RefreshContainerDocuments "Tables"
    On Error Resume Next
    Set doc = SharedDb.Containers("Tables").Documents(strQueryName)
    QueryDocumentExists = (Err.Number = 0)
    Err.Clear

End Function


Private Function QueryDescription(strQueryName As String) As String

    Dim doc As DAO.Document

    Set doc = SharedDb.Containers("Tables").Documents(strQueryName)
    QueryDescription = CStr(doc.Properties("Description").Value)

End Function


Private Function QueryProperty(strQueryName As String, strPropertyName As String) As String

    Dim doc As DAO.Document

    Set doc = SharedDb.Containers("Tables").Documents(strQueryName)
    QueryProperty = CStr(doc.Properties(strPropertyName).Value)

End Function
