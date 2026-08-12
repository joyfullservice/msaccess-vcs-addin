Attribute VB_Name = "modTestImexSpec"
'---------------------------------------------------------------------------------------
' Module    : modTestImexSpec
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Import/merge tests for IMEX specs, including the unnamed SpecName
'           : singleton and explicit SpecID allocation.
'           : Run: ?VCS.RunTests("modTestImexSpec")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")


Private Const TEST_PREFIX As String = "vcs_test_imex"
Private Const TEST_COL As String = "vcs_test_imex_col"
Private Const TEST_NAMED As String = "vcs_test_imex_named"
Private Const TEST_NEW As String = "vcs_test_imex_new"


'---------------------------------------------------------------------------------------
' Procedure : TestUnnamedSpecExportMergeRoundtrip
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : An unnamed spec survives export and merge as a single header row with
'           : its columns intact and not attached to any other spec.
'---------------------------------------------------------------------------------------
'
Public Sub TestUnnamedSpecExportMergeRoundtrip()
    '@Tag("integration")

    Dim cSpec As IDbComponent
    Dim strFile As String
    Dim strExport As String
    Dim lngId As Long
    Dim lngBlank As Long
    Dim lngCols As Long

    DeleteTestSpecs
    strFile = WriteTestSpecFile(vbNullString)

    Set cSpec = New clsDbImexSpec
    cSpec.Import strFile
    lngId = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0)
    TestAssert lngId > 0, "unnamed spec was imported"
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""""") = 1, "exactly one unnamed spec after import"
    TestAssert DCount("*", "MSysIMEXColumns", "SpecID=" & lngId & " AND FieldName=""" & TEST_COL & """") = 1, _
        "test column attached to unnamed spec"
    TestAssert cSpec.Name = "Spec " & lngId, "display name is Spec {id}, not blank"

    strExport = FSO.GetParentFolderName(strFile) & PathSep & "exported.json"
    cSpec.Export strExport
    TestAssert FSO.FileExists(strExport), "export wrote a file"

    cSpec.Merge strExport
    lngBlank = DCount("*", "MSysIMEXSpecs", "SpecName=""""")
    TestAssert lngBlank = 1, "exactly one unnamed spec after merge"
    lngId = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0)
    lngCols = DCount("*", "MSysIMEXColumns", "SpecID=" & lngId & " AND FieldName=""" & TEST_COL & """")
    TestAssert lngCols = 1, "test column still on the unnamed spec after merge"
    TestAssert DCount("*", "MSysIMEXColumns", "FieldName=""" & TEST_COL & """") = 1, _
        "test column is not attached to any other spec"

    DeleteTestSpecs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUnnamedSpecMergeAfterIdDrift
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Merging an unnamed spec still replaces the blank-named row after its
'           : SpecID (and therefore its derived file name) has changed.
'---------------------------------------------------------------------------------------
'
Public Sub TestUnnamedSpecMergeAfterIdDrift()
    '@Tag("integration")

    Dim cSpec As IDbComponent
    Dim dbs As DAO.Database
    Dim strFile As String
    Dim lngOld As Long
    Dim lngNew As Long
    Dim rst As DAO.Recordset

    DeleteTestSpecs
    strFile = WriteTestSpecFile(vbNullString)

    Set cSpec = New clsDbImexSpec
    cSpec.Import strFile
    lngOld = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0)
    TestAssert lngOld > 0, "unnamed spec was imported"

    ' Move the unnamed spec to a different SpecID so the derived file name no longer
    ' matches. Merge must still find it by blank SpecName.
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("SELECT Max(SpecID) FROM MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
    lngNew = Nz(rst(0), 0) + 50
    rst.Close
    dbs.Execute "DELETE FROM MSysIMEXColumns WHERE SpecID=" & lngOld, dbFailOnError
    dbs.Execute "DELETE FROM MSysIMEXSpecs WHERE SpecID=" & lngOld, dbFailOnError
    dbs.Execute "INSERT INTO MSysIMEXSpecs (SpecID, SpecName, SpecType, FileType, StartRow) " & _
        "VALUES (" & lngNew & ", """", 0, 0, 1)", dbFailOnError
    dbs.Execute "INSERT INTO MSysIMEXColumns (SpecID, FieldName, Attributes, DataType, IndexType, SkipColumn, Start, Width) " & _
        "VALUES (" & lngNew & ", """ & TEST_COL & """, 0, 10, 0, False, 1, 10)", dbFailOnError
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""""") = 1, "still one unnamed spec after ID change"
    TestAssert Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0) = lngNew, "unnamed spec now has drifted ID"

    Set cSpec = New clsDbImexSpec
    cSpec.Merge strFile
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""""") = 1, "exactly one unnamed spec after merge across ID drift"
    TestAssert DCount("*", "MSysIMEXColumns", "FieldName=""" & TEST_COL & """") = 1, _
        "test column exists once after merge"
    TestAssert DCount("*", "MSysIMEXColumns", "SpecID=" & lngNew) = 0, _
        "drifted SpecID was replaced, not left behind"

    DeleteTestSpecs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestNamedSpecMergeLeavesUnnamed
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Merging a named spec does not disturb an unnamed spec in the same database.
'---------------------------------------------------------------------------------------
'
Public Sub TestNamedSpecMergeLeavesUnnamed()
    '@Tag("integration")

    Dim cSpec As IDbComponent
    Dim strNamed As String
    Dim strUnnamed As String
    Dim lngUnnamedId As Long

    DeleteTestSpecs
    strNamed = WriteTestSpecFile(TEST_NAMED)
    strUnnamed = WriteTestSpecFile(vbNullString)

    Set cSpec = New clsDbImexSpec
    cSpec.Import strNamed
    Set cSpec = New clsDbImexSpec
    cSpec.Import strUnnamed
    lngUnnamedId = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0)
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""" & TEST_NAMED & """") = 1, "named spec imported"
    TestAssert lngUnnamedId > 0, "unnamed spec imported"

    Set cSpec = New clsDbImexSpec
    cSpec.Merge strNamed
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""" & TEST_NAMED & """") = 1, "named spec still present after merge"
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""""") = 1, "unnamed spec undisturbed"
    TestAssert Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0) = lngUnnamedId, _
        "unnamed SpecID unchanged"
    TestAssert DCount("*", "MSysIMEXColumns", "SpecID=" & lngUnnamedId & " AND FieldName=""" & TEST_COL & """") = 1, _
        "unnamed spec columns intact"

    DeleteTestSpecs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSpecIdAllocationIgnoresAutonumberSeed
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : After restoring specs in an order that ends on a low SpecID, the next
'           : insert must not reuse a live ID. The AutoNumber seed would; Max+1 does not.
'---------------------------------------------------------------------------------------
'
Public Sub TestSpecIdAllocationIgnoresAutonumberSeed()
    '@Tag("integration")

    Dim cSpec As IDbComponent
    Dim rst As DAO.Recordset
    Dim lngBase As Long
    Dim lngOccupy As Long
    Dim lngHigh As Long
    Dim lngLast As Long
    Dim lngNew As Long
    Dim strOccupy As String
    Dim strHigh As String
    Dim strLast As String
    Dim strNew As String

    DeleteTestSpecs
    Set rst = CurrentDb.OpenRecordset("SELECT Max(SpecID) FROM MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
    lngBase = Nz(rst(0), 0)
    rst.Close
    lngOccupy = lngBase + 10
    lngHigh = lngBase + 11
    lngLast = lngBase + 9

    strOccupy = WriteTestSpecFile(TEST_PREFIX & "_occ", lngOccupy)
    strHigh = WriteTestSpecFile(TEST_PREFIX & "_hi", lngHigh)
    strLast = WriteTestSpecFile(TEST_PREFIX & "_lo", lngLast)
    strNew = WriteTestSpecFile(TEST_NEW)

    Set cSpec = New clsDbImexSpec
    cSpec.Import strOccupy
    Set cSpec = New clsDbImexSpec
    cSpec.Import strHigh
    Set cSpec = New clsDbImexSpec
    cSpec.Import strLast
    ' Seed is now lngLast+1 = lngOccupy, which is already live.
    Set cSpec = New clsDbImexSpec
    cSpec.Import strNew

    Set rst = CurrentDb.OpenRecordset( _
        "SELECT SpecID, COUNT(*) AS Cnt FROM MSysIMEXSpecs GROUP BY SpecID HAVING COUNT(*)>1", _
        dbOpenSnapshot, dbReadOnly)
    TestAssert rst.EOF, "no duplicate SpecIDs after seed-trap sequence"
    rst.Close

    lngNew = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName=""" & TEST_NEW & """"), 0)
    TestAssert lngNew > 0, "new spec was imported"
    TestAssert lngNew <> lngOccupy, "new spec did not reuse the occupied seed target"
    TestAssert lngNew = lngHigh + 1, "new spec used Max(SpecID)+1"

    DeleteTestSpecs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDuplicateDerivedFileNameWarns
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : A spec named "Spec N" alongside an unnamed spec with SpecID N must not
'           : raise on GetAllFromDB; the duplicate derived file name is logged instead.
'---------------------------------------------------------------------------------------
'
Public Sub TestDuplicateDerivedFileNameWarns()
    '@Tag("integration")

    Dim cSpec As IDbComponent
    Dim dItems As Dictionary
    Dim lngId As Long
    Dim lngErrBefore As Long
    Dim strNamed As String
    Dim strUnnamed As String

    DeleteTestSpecs
    strUnnamed = WriteTestSpecFile(vbNullString)
    Set cSpec = New clsDbImexSpec
    cSpec.Import strUnnamed
    lngId = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName="""""), 0)
    TestAssert lngId > 0, "unnamed spec imported"

    strNamed = WriteTestSpecFile("Spec " & lngId)
    Set cSpec = New clsDbImexSpec
    cSpec.Import strNamed
    TestAssert DCount("*", "MSysIMEXSpecs", "SpecName=""Spec " & lngId & """") = 1, _
        "named Spec N imported alongside unnamed spec"

    lngErrBefore = Log.ErrorCount
    Set cSpec = New clsDbImexSpec
    Set dItems = cSpec.GetAllFromDB
    TestAssert Log.ErrorCount > lngErrBefore, "duplicate derived file name logged a warning"
    TestAssert dItems.Exists(cSpec.BaseFolder & GetSafeFileName("Spec " & lngId) & ".json"), _
        "derived file name is present once in the collection"

    DeleteTestSpecs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteTestSpecFile
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Write a minimal IMEX spec JSON file and return its path.
'---------------------------------------------------------------------------------------
'
Private Function WriteTestSpecFile(strName As String, Optional varSpecId As Variant) As String

    Dim dItems As Dictionary
    Dim dCols As Dictionary
    Dim dCol As Dictionary
    Dim strFolder As String
    Dim strFile As String
    Dim strBase As String

    Set dItems = New Dictionary
    If Not IsMissing(varSpecId) Then
        If Len(CStr(varSpecId)) > 0 Then dItems.Add "SpecID", CLng(varSpecId)
    End If
    dItems.Add "DateDelim", "/"
    dItems.Add "DateFourDigitYear", False
    dItems.Add "DateLeadingZeros", False
    dItems.Add "DateOrder", 2
    dItems.Add "DecimalPoint", "."
    dItems.Add "FieldSeparator", ","
    dItems.Add "FileType", 0
    dItems.Add "SpecName", strName
    dItems.Add "SpecType", 0
    dItems.Add "StartRow", 1
    dItems.Add "TextDelim", """"
    dItems.Add "TimeDelim", ":"

    Set dCol = New Dictionary
    dCol.Add "Attributes", 0
    dCol.Add "DataType", 10
    dCol.Add "IndexType", 0
    dCol.Add "SkipColumn", False
    dCol.Add "Start", 1
    dCol.Add "Width", 10
    Set dCols = New Dictionary
    dCols.Add TEST_COL, dCol
    dItems.Add "Columns", dCols

    strFolder = GetTempFolder("vcs_imex") & PathSep
    If strName = vbNullString Then
        strBase = "unnamed"
    Else
        strBase = GetSafeFileName(strName)
    End If
    strFile = strFolder & strBase & ".json"
    WriteFile BuildJsonFile("clsDbImexSpec", dItems, "Import/Export Specification from MSysIMEXSpecs"), strFile
    WriteTestSpecFile = strFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteTestSpecs
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Remove specs created by this module, including an unnamed spec that
'           : carries the test column. Leaves any unrelated unnamed spec alone.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteTestSpecs()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim colIds As Collection
    Dim varId As Variant
    Dim strSql As String

    If Not TableExists("MSysIMEXSpecs") Then Exit Sub
    Set dbs = CurrentDb
    Set colIds = New Collection

    strSql = "SELECT SpecID FROM MSysIMEXSpecs WHERE SpecName Like """ & TEST_PREFIX & "*"""
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    Do While Not rst.EOF
        colIds.Add rst!SpecID
        rst.MoveNext
    Loop
    rst.Close

    strSql = "SELECT DISTINCT s.SpecID FROM MSysIMEXSpecs s INNER JOIN MSysIMEXColumns c " & _
        "ON s.SpecID=c.SpecID WHERE (s.SpecName="""" OR s.SpecName Is Null) AND c.FieldName=""" & TEST_COL & """"
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    Do While Not rst.EOF
        colIds.Add rst!SpecID
        rst.MoveNext
    Loop
    rst.Close

    For Each varId In colIds
        dbs.Execute "DELETE FROM MSysIMEXColumns WHERE SpecID=" & varId, dbFailOnError
        dbs.Execute "DELETE FROM MSysIMEXSpecs WHERE SpecID=" & varId, dbFailOnError
    Next varId

End Sub
