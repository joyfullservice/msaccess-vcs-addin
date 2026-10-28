Attribute VB_Name = "modTestDatabase"
'---------------------------------------------------------------------------------------
' Module    : modTestDatabase
' Author    : Adam Waller
' Date      : 6/19/2026
' Purpose   : Unit tests for modDatabase utility functions (engine-managed property
'           : handling for table/object property round-tripping).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Utility")
'@Tag("unit")


Public Sub TestIsEngineManagedProperty()
    ' The FCMin* feature-compatibility version stamps are engine-managed.
    TestAssert IsEngineManagedProperty("FCMinDesignVer"), "FCMinDesignVer"
    TestAssert IsEngineManagedProperty("FCMinReadVer"), "FCMinReadVer"
    TestAssert IsEngineManagedProperty("FCMinWriteVer"), "FCMinWriteVer"
    TestAssert IsEngineManagedProperty("fcminwritever"), "case-insensitive match"
    ' Ordinary display/custom properties are not.
    TestAssert Not IsEngineManagedProperty("Description"), "Description is settable"
    TestAssert Not IsEngineManagedProperty("ColumnWidth"), "ColumnWidth is settable"
    TestAssert Not IsEngineManagedProperty("FC"), "shorter than FCMin prefix"
    TestAssert Not IsEngineManagedProperty(vbNullString), "empty string"
End Sub


Public Sub TestFilterEngineManagedProps()
    Dim dIn As Dictionary
    Dim dOut As Dictionary

    Set dIn = New Dictionary
    dIn.CompareMode = TextCompare
    dIn.Add "Description", "desc"
    dIn.Add "FCMinDesignVer", "16.0.12600.10000"
    dIn.Add "FCMinReadVer", "16.0.12600.10000"
    dIn.Add "FCMinWriteVer", "16.0.12600.10000"
    dIn.Add "ColumnWidth", 1440

    Set dOut = FilterEngineManagedProps(dIn)
    TestAssert dOut.Count = 2, "only non-engine props remain"
    TestAssert dOut.Exists("Description"), "Description kept"
    TestAssert dOut.Exists("ColumnWidth"), "ColumnWidth kept"
    TestAssert Not dOut.Exists("FCMinDesignVer"), "FCMinDesignVer removed"
    TestAssert Not dOut.Exists("FCMinWriteVer"), "FCMinWriteVer removed"
    TestAssert dOut.CompareMode = TextCompare, "compare mode preserved"
    TestAssert dIn.Count = 5, "input dictionary not mutated"
End Sub


Public Sub TestFilterEngineManagedProps_NoEngineProps()
    Dim dIn As Dictionary
    Dim dOut As Dictionary

    Set dIn = New Dictionary
    dIn.Add "Description", "desc"
    dIn.Add "Caption", "cap"

    Set dOut = FilterEngineManagedProps(dIn)
    TestAssert dOut.Count = 2, "all properties retained when none are engine-managed"
End Sub


Public Sub TestCloseOpenObjectsForTypeNoOpTypes()
    TestAssert CloseOpenObjectsForType(edbCommandBar, acSaveYes), "command bars no-op succeeds"
    TestAssert CloseOpenObjectsForType(edbModule, acSaveYes), "modules no-op succeeds"
    TestAssert CloseOpenObjectsForType(edbVbeReference, acSaveYes), "references no-op succeeds"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableLookupsWithQuotesInName
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/5/2026
' Purpose   : TableExists and IsLocalTable look the name up in MSysObjects inside a
'           : double-quoted literal. A double quote in a table name must be escaped
'           : there; the apostrophe tables are the control, and must keep matching.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableLookupsWithQuotesInName()
    '@Tag("integration")

    Const cLocalQuote As String = "tblVcsTest""Quote"
    Const cLocalApos As String = "tblVcsTest'Apos"
    Const cLinkedQuote As String = "tblVcsTestLink""Quote"
    Const cLinkedApos As String = "tblVcsTestLink'Apos"

    Dim dbs As DAO.Database
    Dim dbsBack As DAO.Database
    Dim strFolder As String
    Dim strBackEnd As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    DropQuoteTestTables Array(cLocalQuote, cLocalApos, cLinkedQuote, cLinkedApos)

    ' Back end for the linked tables
    strFolder = GetTempFolder("vcs_quoted_names")
    strBackEnd = strFolder & PathSep & "backend.accdb"
    Set dbsBack = DBEngine.CreateDatabase(strBackEnd, dbLangGeneral)
    dbsBack.Execute "CREATE TABLE [Source] ([ID] LONG)", dbFailOnError
    dbsBack.Close
    Set dbsBack = Nothing

    Set dbs = CurrentDb
    CreateQuoteTestTable dbs, cLocalQuote, vbNullString
    CreateQuoteTestTable dbs, cLocalApos, vbNullString
    CreateQuoteTestTable dbs, cLinkedQuote, strBackEnd
    CreateQuoteTestTable dbs, cLinkedApos, strBackEnd
    dbs.TableDefs.Refresh
    ReleaseDbReferences

    ' Controls first: the apostrophe must keep working
    TestAssert ProbeTableExists(cLocalApos), "control: local table with an apostrophe exists"
    TestAssert ProbeTableExists(cLinkedApos), "control: linked table with an apostrophe exists"
    TestAssert ProbeIsLocalTable(cLocalApos), "control: local table with an apostrophe is local"
    TestAssert ProbeIsNotLocalTable(cLinkedApos), "control: linked table with an apostrophe is not local"

    TestAssert ProbeTableExists(cLocalQuote), "local table with a double quote exists"
    TestAssert ProbeTableExists(cLinkedQuote), "linked table with a double quote exists"
    TestAssert ProbeIsLocalTable(cLocalQuote), "local table with a double quote is local"
    TestAssert ProbeIsNotLocalTable(cLinkedQuote), "linked table with a double quote is not local"

CleanUp:
    On Error Resume Next
    DropQuoteTestTables Array(cLocalQuote, cLocalApos, cLinkedQuote, cLinkedApos)
    If Len(strFolder) > 0 Then If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    Err.Clear
    If lngErr <> 0 Then TestAssert False, "quoted table name error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDataMacroScanWithQuotesInName
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/5/2026
' Purpose   : The data macro scan queries MSysObjects for every local table. A double
'           : quote in one table name must not stop the scan.
'---------------------------------------------------------------------------------------
'
Public Sub TestDataMacroScanWithQuotesInName()
    '@Tag("integration")

    Const cLocalQuote As String = "tblVcsTestMacro""Quote"

    Dim cCategory As IDbComponent
    Dim lngScanErr As Long
    Dim strScanErr As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    DropQuoteTestTables Array(cLocalQuote)
    CreateQuoteTestTable CurrentDb, cLocalQuote, vbNullString
    CurrentDb.TableDefs.Refresh
    ReleaseDbReferences

    Set cCategory = New clsDbTableDataMacro
    On Error Resume Next
    cCategory.GetAllFromDB False
    lngScanErr = Err.Number
    strScanErr = Err.Description
    Err.Clear
    On Error GoTo ErrHandler

    TestAssert lngScanErr = 0, "data macro scan with a double quote in a table name (error " & _
        lngScanErr & ": " & strScanErr & ")"

CleanUp:
    On Error Resume Next
    DropQuoteTestTables Array(cLocalQuote)
    Err.Clear
    If lngErr <> 0 Then TestAssert False, "data macro scan error " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProbeTableExists / ProbeIsLocalTable / ProbeIsNotLocalTable
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/5/2026
' Purpose   : Call the function under test and return False on a runtime error, so a
'           : failing lookup is reported as its own assertion and the rest still run.
'           : ProbeIsNotLocalTable exists so that a runtime error cannot pass as the
'           : expected False.
'---------------------------------------------------------------------------------------
'
Private Function ProbeIsNotLocalTable(strName As String) As Boolean
    On Error Resume Next
    ProbeIsNotLocalTable = Not IsLocalTable(strName)
    If Err.Number <> 0 Then
        Debug.Print "IsLocalTable(" & strName & ") error " & Err.Number & ": " & Err.Description
        ProbeIsNotLocalTable = False
    End If
    Err.Clear
End Function


Private Function ProbeTableExists(strName As String) As Boolean
    On Error Resume Next
    ProbeTableExists = TableExists(strName)
    If Err.Number <> 0 Then Debug.Print "TableExists(" & strName & ") error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function


Private Function ProbeIsLocalTable(strName As String) As Boolean
    On Error Resume Next
    ProbeIsLocalTable = IsLocalTable(strName)
    If Err.Number <> 0 Then Debug.Print "IsLocalTable(" & strName & ") error " & Err.Number & ": " & Err.Description
    Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateQuoteTestTable
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/5/2026
' Purpose   : Create a one-field local table, or a table linked to strBackEnd.
'---------------------------------------------------------------------------------------
'
Private Sub CreateQuoteTestTable(dbs As DAO.Database, strTable As String, strBackEnd As String)

    Dim tdf As DAO.TableDef

    Set tdf = dbs.CreateTableDef(strTable)
    If Len(strBackEnd) = 0 Then
        tdf.Fields.Append tdf.CreateField("ID", dbLong)
    Else
        tdf.Connect = ";DATABASE=" & strBackEnd
        tdf.SourceTableName = "Source"
    End If
    dbs.TableDefs.Append tdf

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DropQuoteTestTables
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 10/5/2026
' Purpose   : Drop the named tables if present. Walks TableDefs instead of calling
'           : TableExists, which is the function under test.
'---------------------------------------------------------------------------------------
'
Private Sub DropQuoteTestTables(varNames As Variant)

    Dim dbs As DAO.Database
    Dim varName As Variant
    Dim tdf As DAO.TableDef
    Dim blnFound As Boolean

    Set dbs = CurrentDb
    For Each varName In varNames
        blnFound = False
        For Each tdf In dbs.TableDefs
            If tdf.Name = CStr(varName) Then
                blnFound = True
                Exit For
            End If
        Next tdf
        If blnFound Then dbs.TableDefs.Delete CStr(varName)
    Next varName
    dbs.TableDefs.Refresh
    ReleaseDbReferences

End Sub
