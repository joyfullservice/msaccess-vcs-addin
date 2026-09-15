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
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    BeginQuerySandbox strRoot, strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex
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
    RestoreQuerySandbox strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex
    On Error GoTo 0
    If lngErr <> 0 Then Err.Raise lngErr, , strErr
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
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    BeginQuerySandbox strRoot, strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex
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
    RestoreQuerySandbox strSavedExport, lngSavedFormat, blnSavedDeterministic, cSavedIndex
    On Error GoTo 0
    If lngErr <> 0 Then Err.Raise lngErr, , strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Private Sub BeginQuerySandbox(ByRef strRoot As String, ByRef strSavedExport As String, _
    ByRef lngSavedFormat As Long, ByRef blnSavedDeterministic As Boolean, _
    ByRef cSavedIndex As clsVCSIndex)

    Set cSavedIndex = VCSIndex
    strSavedExport = Options.ExportFolder
    lngSavedFormat = Options.ExportFormatVersion
    blnSavedDeterministic = Options.UseDeterministicQueryExport

    strRoot = GetTempFolder("vcs_batch_import") & PathSep
    VerifyPath strRoot & "queries" & PathSep
    Options.ExportFolder = strRoot
    Options.ExportFormatVersion = EFV_5_1_0
    Options.UseDeterministicQueryExport = True
    Set VCSIndex = Nothing

End Sub


Private Sub RestoreQuerySandbox(strSavedExport As String, lngSavedFormat As Long, _
    blnSavedDeterministic As Boolean, cSavedIndex As clsVCSIndex)

    Options.ExportFolder = strSavedExport
    Options.ExportFormatVersion = lngSavedFormat
    Options.UseDeterministicQueryExport = blnSavedDeterministic
    Set VCSIndex = cSavedIndex

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
