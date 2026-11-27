Attribute VB_Name = "modTestConnect"
'---------------------------------------------------------------------------------------
' Module    : modTestConnect
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Unit tests for modConnect connection string functions.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Connect")
'@Tag("unit")


Public Sub TestSanitizeConnectionString()
    TestAssert SanitizeConnectionString(";test;test;") = ";test;test;", "preserves semicolons"
    TestAssert SanitizeConnectionString("test;test") = "test;test", "middle semicolons"
    TestAssert SanitizeConnectionString("test") = "test", "no semicolons"
    TestAssert SanitizeConnectionString(vbNullString) = vbNullString, "empty string"
End Sub


Public Sub TestStripConnectionCredentials()

    Dim strResult As String

    ' SQL auth: both UID and PWD removed, everything else retained
    strResult = StripConnectionCredentials( _
        "ODBC;DRIVER={SQL Server};SERVER=svr;UID=user;PWD=secret;DATABASE=db")
    TestAssert InStr(1, strResult, "PWD=", vbTextCompare) = 0, "SQL auth: PWD removed"
    TestAssert InStr(1, strResult, "UID=", vbTextCompare) = 0, "SQL auth: UID removed"
    TestAssert InStr(1, strResult, "secret") = 0, "SQL auth: password value gone"
    TestAssert InStr(strResult, "SERVER=svr") > 0, "SQL auth: SERVER retained"
    TestAssert InStr(strResult, "DATABASE=db") > 0, "SQL auth: DATABASE retained"

    ' Access back-end database password removed
    strResult = StripConnectionCredentials("MS Access;PWD=secret;DATABASE=C:\data\be.accdb")
    TestAssert InStr(1, strResult, "PWD=", vbTextCompare) = 0, "Access: PWD removed"
    TestAssert InStr(strResult, "DATABASE=C:\data\be.accdb") > 0, "Access: DATABASE retained"

    ' Case-insensitive key matching
    strResult = StripConnectionCredentials("ODBC;DRIVER=x;pwd=secret;uid=user;DATABASE=d")
    TestAssert InStr(strResult, "secret") = 0, "lowercase pwd= removed"
    TestAssert InStr(1, strResult, "uid=", vbTextCompare) = 0, "lowercase uid= removed"

    ' Auth method (AD/integrated) is preserved; only the empty PWD= is dropped
    strResult = StripConnectionCredentials( _
        "ODBC;DRIVER=x;SERVER=s;PWD=;DATABASE=d;Authentication=ActiveDirectoryIntegrated")
    TestAssert InStr(1, strResult, "PWD=", vbTextCompare) = 0, "empty PWD= dropped"
    TestAssert InStr(strResult, "Authentication=ActiveDirectoryIntegrated") > 0, _
        "Authentication method retained"

    ' Connection without credentials is unchanged in substance
    strResult = StripConnectionCredentials("ODBC;DRIVER=x;SERVER=s;DATABASE=d")
    TestAssert InStr(strResult, "SERVER=s") > 0 And InStr(strResult, "DATABASE=d") > 0, _
        "no-credential string retains all parts"

    TestAssert StripConnectionCredentials(vbNullString) = vbNullString, "empty string"

End Sub


Public Sub TestGetSourceSafeConnectGating()

    Dim lngSaved As Long
    Dim eimPriorMode As eInteractionMode
    Dim strConn As String
    Dim strAd As String
    Dim strResult As String

    strConn = "ODBC;DRIVER={SQL Server};SERVER=svr;UID=user;PWD=secret;DATABASE=db"
    strAd = "ODBC;DRIVER=x;SERVER=s;PWD=;DATABASE=d;Authentication=ActiveDirectoryIntegrated"

    ' Stripping a real password logs one eelWarning by design. The VCS.RunTests
    ' harness already runs eimSilent, but force it here too so this test never
    ' pops a MsgBox when run standalone (F5 / Immediate window). Cache + restore
    ' like modTestErrorHandling.TestCatch. The logged warning is expected.
    eimPriorMode = Operation.InteractionMode
    Operation.InteractionMode = eimSilent

    ' Preserve and restore the shared option (TestAssert is non-fatal, so the
    ' restore at the end always runs even if an assertion fails).
    lngSaved = Options.ExportFormatVersion

    ' Below 5.0.0: behavior unchanged - credentials pass through to source.
    Options.ExportFormatVersion = EFV_4_1_2
    strResult = GetSourceSafeConnect(strConn, "test (linked table)")
    TestAssert InStr(strResult, "PWD=secret") > 0, "pre-5.0.0 leaves password untouched"

    ' 5.0.0+: real password stripped from anything bound for source.
    Options.ExportFormatVersion = EFV_5_0_0
    strResult = GetSourceSafeConnect(strConn, "test (linked table)")
    TestAssert InStr(1, strResult, "PWD=", vbTextCompare) = 0, "5.0.0 strips PWD"
    TestAssert InStr(strResult, "secret") = 0, "5.0.0 removes password value"

    ' 5.0.0+ with passwordless auth (empty PWD): no secret, returned unchanged.
    strResult = GetSourceSafeConnect(strAd, "test (linked table)")
    TestAssert strResult = strAd, "passwordless AD connection is not altered"

    ' 5.0.0+ with no credentials at all: returned unchanged.
    strResult = GetSourceSafeConnect("ODBC;DRIVER=x;SERVER=s;DATABASE=d", "test")
    TestAssert strResult = "ODBC;DRIVER=x;SERVER=s;DATABASE=d", "no-credential string unchanged"

    Options.ExportFormatVersion = lngSaved
    Operation.InteractionMode = eimPriorMode

End Sub


Public Sub TestGetConnectPart()
    Dim strConn As String
    strConn = "ODBC;DRIVER={SQL Server};SERVER=mysvr;DATABASE=mydb"
    TestAssert GetConnectPart(strConn, "SERVER") = "mysvr", "extracts SERVER"
    TestAssert GetConnectPart(strConn, "DATABASE") = "mydb", "extracts DATABASE (last part)"
    TestAssert GetConnectPart(strConn, "DRIVER") = "{SQL Server}", "extracts DRIVER"
    TestAssert GetConnectPart(strConn, "MISSING") = "", "missing part returns empty"
    TestAssert GetConnectPart("", "ANY") = "", "empty string"
End Sub


Public Sub TestIsEnvReference()
    TestAssert IsEnvReference("env:conn_mydb"), "valid env reference"
    TestAssert IsEnvReference("ENV:conn_mydb"), "case insensitive"
    TestAssert Not IsEnvReference("not_env"), "not an env reference"
    TestAssert Not IsEnvReference(""), "empty string"
End Sub


Public Sub TestAccessBackEndConnectKey()
    TestAssert GetBackEndConnectKey(";DATABASE=C:\Data\MyDatabase.accdb") = _
        GetBackEndConnectKey("MS Access;PWD=secret;DATABASE=C:\DATA\MYDATABASE.ACCDB"), _
        "connection string casing normalizes to same back-end key"
End Sub


Public Sub TestFileBackEndConnection()

    Const TEXT_LINK As String = "Text;FMT=Delimited;HDR=YES;IMEX=2;ACCDB=YES;DATABASE="
    Const EXCEL_LINK As String = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE="

    Dim strDir As String
    Dim strMissing As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strDir = GetTempFolder("FileLink")
    strMissing = strDir & PathSep & "missing"
    WriteFile "ID", strDir & PathSep & "book.xlsx"

    TestAssert TestBackEndConnection(TEXT_LINK & strDir), "text link to an existing folder"
    TestAssert Not TestBackEndConnection(TEXT_LINK & strMissing), "text link to a missing folder"
    TestAssert TestBackEndConnection(EXCEL_LINK & strDir & PathSep & "book.xlsx"), _
        "Excel link to an existing file"
    TestAssert Not TestBackEndConnection(EXCEL_LINK & strMissing & ".xlsx"), _
        "Excel link to a missing file"
    TestAssert TestBackEndConnection("WSS;HDR=NO;IMEX=2;ACCDB=YES;" & _
        "DATABASE=https://example.com/sites/team;LIST={00000000-0000-0000-0000-000000000000};"), _
        "a SharePoint list has no local path to check"

CleanUp:
    On Error Resume Next
    If Len(strDir) > 0 Then If FSO.FolderExists(strDir) Then FSO.DeleteFolder strDir, True
    If lngErr <> 0 Then TestAssert False, "file back-end connection " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestGetConnectionEnvKey()
    Dim strKey As String

    ' Access back-end: should use filename as key identity
    strKey = GetConnectionEnvKey(";DATABASE=C:\Data\MyDatabase.accdb")
    TestAssert Len(strKey) > 0, "non-empty key for Access connection"
    TestAssert Left$(strKey, 5) = "conn_", "starts with conn_ prefix"

    ' ODBC with DATABASE: should use database name
    strKey = GetConnectionEnvKey("ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=SalesDB")
    TestAssert Left$(strKey, 5) = "conn_", "ODBC key starts with conn_ prefix"
    TestAssert InStr(strKey, "salesdb") > 0, "ODBC key contains db name"

    ' No DATABASE or DSN: falls back to 7-char hash of connection string (credentials excluded)
    strKey = GetConnectionEnvKey("ODBC;DRIVER={SQL Server};SERVER=svr;UID=user;PWD=secret")
    TestAssert Left$(strKey, 5) = "conn_", "hash fallback starts with conn_ prefix"
    TestAssert Len(strKey) = 12, "hash fallback is conn_ plus 7-char hash"
    TestAssert GetConnectionEnvKey("ODBC;DRIVER={SQL Server};SERVER=svr;UID=user;PWD=secret") = strKey, _
        "hash fallback is deterministic"
End Sub


Public Sub TestGetConnStateKey()

    Const CONN_A As String = "ODBC;DRIVER={SQL Server};SERVER=host.domain;Trusted_Connection=Yes;DATABASE=Db1"
    Const CONN_B As String = "ODBC;DRIVER={SQL Server};SERVER=host;Trusted_Connection=Yes;DATABASE=Db1"
    Const CONN_USER As String = "ODBC;DRIVER={SQL Server};SERVER=host;DATABASE=Db1;UID=user"

    TestAssert GetConnectionEnvKey(CONN_A) = GetConnectionEnvKey(CONN_B), _
        "the .env key alone cannot tell the two servers apart"
    TestAssert GetConnStateKey(CONN_A, vbNullString) <> GetConnStateKey(CONN_B, vbNullString), _
        "connections that differ only by SERVER= keep separate state"
    TestAssert GetConnStateKey(";DATABASE=C:\a\be.accdb", vbNullString) <> _
        GetConnStateKey(";DATABASE=C:\b\be.accdb", vbNullString), _
        "Access back ends with the same file name keep separate state"
    TestAssert GetConnStateKey(CONN_A, "conn_db1") = "conn_db1", _
        "an .env reference is keyed by its .env key"
    TestAssert GetConnStateKey(CONN_USER & ";PWD=secret", vbNullString) = GetConnStateKey(CONN_USER, vbNullString), _
        "the password does not change the key"
    TestAssert InStr(1, GetConnStateKey(CONN_USER & ";PWD=secret", vbNullString), "secret", vbTextCompare) = 0, _
        "the key does not hold the password"
    TestAssert GetConnStateKey(CONN_USER, vbNullString) <> _
        GetConnStateKey(Replace(CONN_USER, "UID=user", "UID=other"), vbNullString), _
        "another user is another connection"
    TestAssert GetConnStateKey(UCase$(CONN_B), vbNullString) = GetConnStateKey(CONN_B, vbNullString), _
        "the key ignores case"
    TestAssert Len(GetConnStateKey(vbNullString, vbNullString)) = 0, "no connection, no key"

End Sub


Public Sub TestGetEnvKeyClash()

    Const ENV_KEY As String = "conn_db1"
    Const CONN_A As String = "ODBC;DRIVER={SQL Server};SERVER=host.domain;Trusted_Connection=Yes;DATABASE=Db1"
    Const CONN_B As String = "ODBC;DRIVER={SQL Server};SERVER=host;Trusted_Connection=Yes;DATABASE=Db1"

    Dim strFullBackEnd As String

    ClearEnvCache
    TestAssert GetEnvKeyClash(ENV_KEY, CONN_A, "tblFirst") = vbNullString, "the first save is not a clash"
    TestAssert GetEnvKeyClash(ENV_KEY, CONN_A & ";APP=Microsoft Office", "tblSame") = vbNullString, _
        "a driver-added setting is not a clash"
    TestAssert GetEnvKeyClash(ENV_KEY, CONN_B, "tblSecond") = "tblFirst", _
        "another server on the same key names the first object"
    TestAssert GetEnvKeyClash(ENV_KEY, CONN_B, "tblThird") = vbNullString, "a key is reported once"
    TestAssert GetEnvKeyClash("conn_db2", Replace(CONN_B, "Db1", "Db2"), "tblOther") = vbNullString, _
        "another key is tracked on its own"

    ClearEnvCache
    TestAssert GetEnvKeyClash(ENV_KEY, CONN_A, "tblFirst") = vbNullString, "clearing the cache starts a new export"

    ClearEnvCache
    strFullBackEnd = FSO.BuildPath(CurrentProject.Path, "be.accdb")
    GetEnvKeyClash "conn_be", ";DATABASE=" & strFullBackEnd, "tblRaw"
    TestAssert GetEnvKeyClash("conn_be", ";DATABASE=rel:be.accdb", "tblSanitized") = vbNullString, _
        "a relative path to the same back end is not a clash"
    ClearEnvCache

End Sub


Public Sub TestIsOracleOdbcConnect()
    TestAssert IsOracleOdbcConnect( _
        "ODBC;DRIVER={Oracle ODBC Driver};SERVER=ora;UID=u;PWD=p"), _
        "Oracle ODBC Driver detected"
    TestAssert IsOracleOdbcConnect( _
        "ODBC;DRIVER={Oracle in OraClient11g_home1};DBQ=tns;UID=u"), _
        "Oracle in OraClient detected"
    TestAssert IsOracleOdbcConnect( _
        "ODBC;DRIVER={Microsoft ODBC for Oracle};SERVER=ora"), _
        "Microsoft ODBC for Oracle detected"
    TestAssert Not IsOracleOdbcConnect( _
        "ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=db"), _
        "SQL Server not Oracle"
    TestAssert Not IsOracleOdbcConnect("ODBC;DSN=MyOracle;UID=u"), _
        "unknown DSN falls back to not Oracle"
    TestAssert Not IsOracleOdbcConnect(vbNullString), "empty string not Oracle"
End Sub


Public Sub TestIsOracleDriverName()
    TestAssert IsOracleDriverName("Oracle ODBC Driver"), "Oracle ODBC Driver"
    TestAssert IsOracleDriverName("Oracle in OraClient19Home1"), "OraClient"
    TestAssert IsOracleDriverName("Microsoft ODBC for Oracle"), "Microsoft ODBC for Oracle"
    TestAssert IsOracleDriverName("Devart ODBC Driver for Oracle"), "Devart"
    TestAssert IsOracleDriverName("{Oracle ODBC Driver}"), "braces retained"
    TestAssert Not IsOracleDriverName("SQL Server"), "SQL Server"
    TestAssert Not IsOracleDriverName(vbNullString), "empty"
    TestAssert IsOracleDriverName("C:\Oracle\bin\sqora32.dll"), "sqora32.dll path"
    TestAssert IsOracleDriverName("sqora64.dll"), "sqora64.dll"
    TestAssert IsOracleDriverName("msorcl32.dll"), "msorcl32.dll"
    TestAssert Not IsOracleDriverName("sqlncli11.dll"), "non-Oracle DLL"
End Sub


Public Sub TestGetOdbcDriverName()
    TestAssert GetOdbcDriverName("ODBC;DRIVER={Oracle ODBC Driver};SERVER=ora") = _
        "{Oracle ODBC Driver}", "DRIVER= fast path"
    TestAssert GetOdbcDriverName("ODBC;DRIVER={SQL Server};SERVER=svr") = _
        "{SQL Server}", "SQL Server DRIVER="
    TestAssert GetOdbcDriverName(vbNullString) = vbNullString, "empty string"
    TestAssert GetOdbcDriverName("ODBC;UID=u") = vbNullString, "no DRIVER or DSN"
    TestAssert GetOdbcDriverName("ODBC;DSN=NoSuchDsn12345") = vbNullString, _
        "unknown DSN returns empty"
End Sub


Public Sub TestGetConnectivityProbeSql()
    TestAssert GetConnectivityProbeSql( _
        "ODBC;DRIVER={Oracle ODBC Driver};SERVER=ora") = "SELECT 1 FROM DUAL;", _
        "Oracle uses FROM DUAL"
    TestAssert GetConnectivityProbeSql( _
        "ODBC;DRIVER={SQL Server};SERVER=svr") = "SELECT 1;", _
        "non-Oracle uses SELECT 1"
    TestAssert GetConnectivityProbeSql("ODBC;DSN=MyOracle;UID=u") = "SELECT 1;", _
        "unknown DSN keeps SELECT 1"
End Sub


Public Sub TestIsOracleSyntaxError()
    TestAssert IsOracleSyntaxError( _
        "[Oracle][ODBC][Ora]ORA-00923: FROM keyword not found where expected"), _
        "ORA-00923 is syntax error"
    TestAssert Not IsOracleSyntaxError( _
        "[Oracle][ODBC][Ora]ORA-01017: invalid username/password"), _
        "ORA-01017 is not syntax"
    TestAssert Not IsOracleSyntaxError("ODBC--call failed."), _
        "generic ODBC failure is not syntax"
    TestAssert Not IsOracleSyntaxError(vbNullString), "empty"
End Sub


Public Sub TestGetFileDsnDriverName()

    Dim strDir As String
    Dim strFile As String
    Dim strChain As String

    ClearConnState
    strDir = GetTempFolder("FileDsn") & PathSep
    strFile = strDir & "oracle.dsn"
    WriteFile "[ODBC]" & vbCrLf & "DRIVER=Oracle in OraClient19Home1" & vbCrLf & "DBQ=ORCL", strFile

    TestAssert GetFileDsnDriverName(strFile) = "Oracle in OraClient19Home1", _
        "reads DRIVER from File DSN"
    TestAssert IsOracleOdbcConnect("ODBC;FILEDSN=" & strFile), _
        "FILEDSN with Oracle driver is detected"
    TestAssert GetConnectivityProbeSql("ODBC;FILEDSN=" & strFile) = "SELECT 1 FROM DUAL;", _
        "FILEDSN Oracle uses FROM DUAL"
    TestAssert GetOdbcDriverName("ODBC;FILEDSN=""" & strFile & """") = _
        "Oracle in OraClient19Home1", "quoted FILEDSN= resolves"

    TestAssert GetFileDsnDriverName(strDir & "missing.dsn") = vbNullString, _
        "missing file returns empty"

    strChain = strDir & "chain.dsn"
    WriteFile "[ODBC]" & vbCrLf & "DSN=NoSuchOracleDsn", strChain
    TestAssert GetFileDsnDriverName(strChain) = vbNullString, _
        "FILEDSN with unknown DSN= returns empty"

    On Error Resume Next
    If FSO.FolderExists(StripSlash(strDir)) Then FSO.DeleteFolder StripSlash(strDir), True
    ClearConnState

End Sub


Public Sub TestDbConnectionExportKeepsFullEnvValue()

    Const SANITIZED As String = "ODBC;Driver={ODBC Driver 18 for SQL Server};SERVER=svr;DATABASE=dbExample"
    Const FULL As String = "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=svr;" & _
        "Trusted_Connection=Yes;TrustServerCertificate=Yes;DATABASE=dbExample"

    Dim strSavedFolder As String
    Dim lngSavedVersion As Long
    Dim eSavedUseEnv As eUseEnvConnections
    Dim blnSaved As Boolean
    Dim strRoot As String
    Dim dOriginal As Dictionary
    Dim dInner As Dictionary
    Dim dExport As Dictionary
    Dim cConn As clsDbConnection
    Dim cEnv As clsDotEnv
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strSavedFolder = Options.ExportFolder
    lngSavedVersion = Options.ExportFormatVersion
    eSavedUseEnv = Options.UseEnvForConnections
    blnSaved = True

    strRoot = GetTempFolder("DbConnEnv") & PathSep
    Options.ExportFolder = strRoot
    Options.ExportFormatVersion = EFV_5_0_0
    Options.UseEnvForConnections = uecAlways
    ClearEnvCache

    Set dInner = New Dictionary
    dInner.CompareMode = TextCompare
    dInner.Add FULL, "tblLinkedExample"
    Set dOriginal = New Dictionary
    dOriginal.Add SANITIZED, dInner

    Set cConn = New clsDbConnection
    Set dExport = cConn.GetExportItems(dOriginal)

    TestAssert dExport.Exists("env:conn_dbexample"), "outer key exported as env reference"
    If dExport.Exists("env:conn_dbexample") Then
        TestAssert dExport("env:conn_dbexample").Exists("env:conn_dbexample"), _
            "inner key exported as env reference"
    End If

    Set cEnv = New clsDotEnv
    cEnv.LoadFromFileIfExists strRoot & ".env"
    TestAssert cEnv.GetVar("conn_dbexample", blnUseEnviron:=False) = FULL, _
        ".env keeps the full connection string, not the sanitized outer key"

CleanUp:
    On Error Resume Next
    If blnSaved Then
        Options.ExportFolder = strSavedFolder
        Options.ExportFormatVersion = lngSavedVersion
        Options.UseEnvForConnections = eSavedUseEnv
    End If
    ClearEnvCache
    If Len(strRoot) Then CleanupTempDotEnvFolder strRoot
    If lngErr <> 0 Then TestAssert False, "db connection export " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestConnectionSettingsMatch()

    Const CONN_STORED As String = "ODBC;Driver={ODBC Driver 18 for SQL Server};Server=svr;Database=db;" & _
        "Trusted_Connection=yes;TrustServerCertificate=yes;"
    Const CONN_COMPLETED As String = "ODBC;DRIVER=ODBC Driver 18 for SQL Server;SERVER=svr;UID=DOMAIN\user;" & _
        "Trusted_Connection=yes;APP=Microsoft Office;DATABASE=db;TrustServerCertificate=yes;"

    TestAssert ConnectionSettingsMatch(CONN_STORED, CONN_COMPLETED), _
        "driver-added APP, trusted UID, braces, order, and key case are ignored"
    TestAssert Not ConnectionSettingsMatch(CONN_STORED, Replace(CONN_COMPLETED, "TrustServerCertificate=yes;", "")), _
        "a missing TrustServerCertificate is a difference"
    TestAssert Not ConnectionSettingsMatch("ODBC;DRIVER=x;SERVER=s;UID=a", "ODBC;DRIVER=x;SERVER=s;UID=b"), _
        "UID is compared when the connection is not trusted"
    TestAssert ConnectionSettingsMatch("ODBC;DRIVER=x;SERVER=s", "ODBC;DRIVER=x;SERVER=s;WSID=machine"), _
        "WSID is ignored"

End Sub


Public Sub TestShouldSaveCompletedConnect()

    Const CONN_TRUSTED As String = "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=svr;DATABASE=db;Trusted_Connection=yes"
    Const CONN_COMPLETED As String = "ODBC;DRIVER=ODBC Driver 18 for SQL Server;SERVER=svr;UID=DOMAIN\user;" & _
        "Trusted_Connection=yes;APP=Microsoft Office;DATABASE=db;TrustServerCertificate=yes;"
    Const CONN_BARE As String = "ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=svr;DATABASE=db"

    TestAssert ShouldSaveCompletedConnect(CONN_TRUSTED, CONN_COMPLETED, True), _
        "completed settings add TrustServerCertificate to an authenticating stored value"
    TestAssert Not ShouldSaveCompletedConnect(CONN_TRUSTED & ";TrustServerCertificate=yes", CONN_COMPLETED, True), _
        "completed settings that only add driver keys are not saved"
    TestAssert Not ShouldSaveCompletedConnect(CONN_TRUSTED, CONN_COMPLETED, False), _
        "without completed settings, an authenticating stored value is kept"
    TestAssert ShouldSaveCompletedConnect(CONN_BARE, CONN_COMPLETED, False), _
        "a stored value without authentication is replaced"
    TestAssert Not ShouldSaveCompletedConnect(CONN_TRUSTED, CONN_BARE & ";TrustServerCertificate=yes", True), _
        "a completed string without authentication is never saved"
    TestAssert ShouldSaveCompletedConnect(vbNullString, CONN_COMPLETED, False), _
        "a missing stored value is saved"

End Sub


Public Sub TestPrepareCompletedConnectForSave()

    Const CONN_STORED As String = "ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=db;UID=user;PWD=secret"
    Const CONN_COMPLETED As String = "ODBC;DRIVER=SQL Server;SERVER=svr;UID=user;APP=Microsoft Office;DATABASE=db"

    Dim strResult As String

    TestAssert PrepareCompletedConnectForSave("ODBC;DRIVER=x;Trusted_Connection=yes", CONN_COMPLETED) = CONN_COMPLETED, _
        "no stored password: completed string is saved as-is"
    TestAssert PrepareCompletedConnectForSave(CONN_STORED, CONN_COMPLETED & ";PWD=other") = CONN_COMPLETED & ";PWD=other", _
        "completed string with its own password is saved as-is"

    strResult = PrepareCompletedConnectForSave(CONN_STORED, CONN_COMPLETED)
    TestAssert InStr(strResult, "PWD=secret") > 0, "stored password carried over for the same user"
    TestAssert InStr(strResult, "APP=Microsoft Office") > 0, "completed settings kept"

    strResult = PrepareCompletedConnectForSave(CONN_STORED, CONN_COMPLETED & ";PWD=")
    TestAssert CountConnectKey(strResult, "PWD=") = 1, "empty PWD= replaced, not duplicated"
    TestAssert InStr(strResult, "PWD=secret") > 0, "replacement carries the stored password"

    TestAssert PrepareCompletedConnectForSave(CONN_STORED, Replace(CONN_COMPLETED, "UID=user", "UID=other")) = vbNullString, _
        "stored password is not carried over to a different user"

End Sub


Public Sub TestGetLinkConnect()

    Const CONN_STORED As String = "ODBC;DSN=MyDsn;Trusted_Connection=Yes;DATABASE=Db2"
    Const CONN_COMPLETED As String = "ODBC;DSN=MyDsn;Description=x;Trusted_Connection=Yes;" & _
        "APP=Microsoft Office;DATABASE=Db2;TrustServerCertificate=Yes;"

    TestAssert GetLinkConnect(CONN_STORED, CONN_COMPLETED) = CONN_STORED, _
        "links with the source string, not the driver-completed one"
    TestAssert GetLinkConnect(CONN_STORED, CONN_STORED) = CONN_STORED, _
        "an unchanged completed string links the same"
    TestAssert GetLinkConnect("ODBC;", CONN_COMPLETED) = CONN_COMPLETED, _
        "the bare placeholder for a missing .env key links with the dialog's string"
    TestAssert GetLinkConnect(vbNullString, CONN_COMPLETED) = CONN_COMPLETED, _
        "an empty source links with the completed string"

End Sub


Public Sub TestCompletedConnectSavesButDoesNotLink()

    Const ENV_KEY As String = "conn_db2"
    Const CONN_STORED As String = "ODBC;DSN=MyDsn;Trusted_Connection=Yes;DATABASE=Db2"
    Const CONN_COMPLETED As String = "ODBC;DSN=MyDsn;Description=x;Trusted_Connection=Yes;" & _
        "APP=Microsoft Office;DATABASE=Db2;TrustServerCertificate=Yes;"

    Dim strSavedFolder As String
    Dim eSavedUseEnv As eUseEnvConnections
    Dim blnSaved As Boolean
    Dim strRoot As String
    Dim cEnv As clsDotEnv
    Dim dState As Dictionary
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strSavedFolder = Options.ExportFolder
    eSavedUseEnv = Options.UseEnvForConnections
    blnSaved = True

    strRoot = GetTempFolder("DbConnLink") & PathSep
    Set cEnv = New clsDotEnv
    cEnv.SetVar ENV_KEY, CONN_STORED
    cEnv.SaveToFile strRoot & ".env"
    Options.ExportFolder = strRoot
    Options.UseEnvForConnections = uecAlways
    ClearEnvCache
    ClearConnState

    ' The same calls clsDbConnection.Import makes after CacheConnection succeeds
    RecordAuthenticatedConnection ENV_KEY, ENV_KEY, GetLinkConnect(CONN_STORED, CONN_COMPLETED), _
        CONN_COMPLETED, "db connection", Not ConnectionSettingsMatch(CONN_STORED, CONN_COMPLETED)
    Set dState = GetConnState(ENV_KEY)
    TestAssert dState("LinkConnect") = CONN_STORED, "later linked tables use the stored string"
    TestAssert dState("NeedsSave"), "the completed settings are offered for .env"
    TestAssert dState("SaveConnect") = CONN_COMPLETED, ".env would receive the completed string"

    ClearConnState
    Options.UseEnvForConnections = uecNever
    RecordAuthenticatedConnection ENV_KEY, ENV_KEY, GetLinkConnect(CONN_STORED, CONN_COMPLETED), _
        CONN_COMPLETED, "db connection", True
    Set dState = GetConnState(ENV_KEY)
    TestAssert dState("LinkConnect") = CONN_STORED, "with .env off, tables still use the stored string"
    TestAssert Not dState("NeedsSave"), "with .env off, nothing is saved"

CleanUp:
    On Error Resume Next
    If blnSaved Then
        Options.ExportFolder = strSavedFolder
        Options.UseEnvForConnections = eSavedUseEnv
    End If
    ClearConnState
    ClearEnvCache
    If Len(strRoot) Then CleanupTempDotEnvFolder strRoot
    If lngErr <> 0 Then TestAssert False, "completed connect roles " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Private Function CountConnectKey(strConnect As String, strKey As String) As Long
    CountConnectKey = (Len(strConnect) - Len(Replace(strConnect, strKey, vbNullString, , , vbTextCompare))) \ Len(strKey)
End Function


Public Sub TestResolveEnvReferencesInText()
    ' When no env: references exist, text should pass through unchanged
    Dim strInput As String
    strInput = "DRIVER={SQL Server};SERVER=mysvr"
    TestAssert ResolveEnvReferencesInText(strInput) = strInput, "no env refs unchanged"
End Sub


Public Sub TestDotEnvLoadFromFileIfExistsMissing()
    Dim cEnv As New clsDotEnv
    TestAssert Not cEnv.LoadFromFileIfExists("C:\nonexistent\vcs_dotenv_test\.env"), _
        "missing file returns False"
End Sub


Public Sub TestDotEnvLocalOverridesBase()
    Dim strDir As String
    strDir = SetupTempDotEnvFolder()
    WriteFile "KEY=base", strDir & ".env"
    WriteFile "KEY=local", strDir & ".env.local"

    Dim cEnv As New clsDotEnv
    cEnv.LoadFromFileIfExists strDir & ".env"
    cEnv.LoadFromFileIfExists strDir & ".env.local", blnMerge:=True
    TestAssert cEnv.GetVar("KEY", blnUseEnviron:=False) = "local", _
        ".env.local overrides .env"

    CleanupTempDotEnvFolder strDir
End Sub


Public Sub TestDotEnvNoAppEnvIgnoresEnvFiles()
    Dim strDir As String
    strDir = SetupTempDotEnvFolder()
    WriteFile "KEY=base", strDir & ".env"
    WriteFile "KEY=other", strDir & ".env.dev"

    Dim cEnv As New clsDotEnv
    cEnv.LoadFromFileIfExists strDir & ".env"
    cEnv.LoadFromFileIfExists strDir & ".env.local", blnMerge:=True
    TestAssert Len(cEnv.GetVar("APP_ENV", blnUseEnviron:=False)) = 0, _
        "APP_ENV unset in base config"
    TestAssert cEnv.GetVar("KEY", blnUseEnviron:=False) = "base", _
        "without APP_ENV, .env.dev is not loaded"

    CleanupTempDotEnvFolder strDir
End Sub


Public Sub TestDotEnvAppEnvLayeredPrecedence()
    Dim strDir As String
    Dim cEnv As New clsDotEnv
    Dim strAppEnv As String

    strDir = SetupTempDotEnvFolder()
    WriteFile "APP_ENV=dev" & vbCrLf & "KEY=base", strDir & ".env"
    WriteFile "KEY=base_local", strDir & ".env.local"
    WriteFile "KEY=dev", strDir & ".env.dev"
    WriteFile "KEY=dev_local", strDir & ".env.dev.local"

    cEnv.LoadFromFileIfExists strDir & ".env"
    cEnv.LoadFromFileIfExists strDir & ".env.local", blnMerge:=True
    strAppEnv = cEnv.GetVar("APP_ENV", blnUseEnviron:=False)
    If Len(strAppEnv) > 0 Then
        cEnv.LoadFromFileIfExists strDir & ".env." & strAppEnv, blnMerge:=True
        cEnv.LoadFromFileIfExists strDir & ".env." & strAppEnv & ".local", blnMerge:=True
    End If
    TestAssert cEnv.GetVar("KEY", blnUseEnviron:=False) = "dev_local", _
        ".env.{APP_ENV}.local wins full precedence chain"

    CleanupTempDotEnvFolder strDir
End Sub


Public Sub TestDotEnvMissingMergeFilesSkipped()
    Dim strDir As String
    Dim cEnv As New clsDotEnv

    strDir = SetupTempDotEnvFolder()
    WriteFile "KEY=base", strDir & ".env"

    TestAssert cEnv.LoadFromFileIfExists(strDir & ".env"), "base .env loads"
    TestAssert Not cEnv.LoadFromFileIfExists(strDir & ".env.local", blnMerge:=True), _
        "missing .env.local returns False without error"
    TestAssert cEnv.GetVar("KEY", blnUseEnviron:=False) = "base", _
        "base value retained after skipped merge"

    CleanupTempDotEnvFolder strDir
End Sub


Private Function SetupTempDotEnvFolder() As String
    SetupTempDotEnvFolder = GetTempFolder("DotEnv") & PathSep
End Function


Private Sub CleanupTempDotEnvFolder(strDir As String)
    On Error Resume Next
    If FSO.FolderExists(StripSlash(strDir)) Then FSO.DeleteFolder StripSlash(strDir), True
    On Error GoTo 0
End Sub


Public Sub TestDataSourceKind()

    Dim strProject As String
    Dim strOther As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strProject = GetTempFolder("vcs_datasource_project")
    strOther = GetTempFolder("vcs_datasource_other")
    FSO.CreateTextFile(strProject & "\be.accdb").Close
    FSO.CreateTextFile(strProject & "\lib.accda").Close
    FSO.CreateTextFile(strOther & "\shared.accdb").Close
    FSO.CreateFolder strProject & "-copy"
    FSO.CreateTextFile(strProject & "-copy\be.accdb").Close

    TestAssert GetDataSourceKind(";DATABASE=" & strProject & "\be.accdb", strProject) = "project", _
        "back end beside the database is per copy"
    TestAssert GetDataSourceKind(strProject & "\lib.accda", strProject) = "project", _
        "library path beside the database is per copy"
    TestAssert GetDataSourceKind("Text;DSN=x;DATABASE=" & strProject, strProject & "\") = "project", _
        "a text link to the database folder itself is per copy"
    TestAssert GetDataSourceKind("MS Access;PWD=secret;DATABASE=" & strOther & "\shared.accdb", strProject) = "outside", _
        "back end in another local folder is shared"
    TestAssert GetDataSourceKind(";DATABASE=" & strProject & "\absent.accdb", strProject) = "missing", _
        "back end that was not copied is missing"
    TestAssert GetDataSourceKind(";DATABASE=" & strProject & "-copy\be.accdb", strProject) = "outside", _
        "a sibling folder with the same prefix is not the project folder"
    TestAssert GetDataSourceKind("ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=db", strProject) = "server", _
        "ODBC is a server"
    TestAssert GetDataSourceKind("WSS;HDR=NO;DATABASE=https://example.com/site;LIST={1}", strProject) = "server", _
        "a SharePoint list is a server"

CleanUp:
    On Error Resume Next
    If Len(strProject) Then
        FSO.DeleteFolder strProject, True
        FSO.DeleteFolder strProject & "-copy", True
    End If
    If Len(strOther) Then FSO.DeleteFolder strOther, True
    If lngErr <> 0 Then TestAssert False, "data source kind " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestDataSourceLabelHasNoCredentials()
    Dim strLabel As String

    strLabel = GetDataSourceLabel("ODBC;DRIVER={SQL Server};SERVER=svr;UID=user;PWD=secret;DATABASE=db")
    TestAssert InStr(strLabel, "secret") = 0, "ODBC label has no password"
    TestAssert InStr(1, strLabel, "SVR", vbTextCompare) > 0, "ODBC label names the server"

    strLabel = GetDataSourceLabel("MS Access;PWD=secret;DATABASE=C:\data\be.accdb")
    TestAssert strLabel = "C:\data\be.accdb", "Access back end is labeled by its path"

    strLabel = GetDataSourceLabel("WSS;PWD=secret;DATABASE=https://example.com/site;LIST={1}")
    TestAssert strLabel = "https://example.com/site", "SharePoint list is labeled by its site"

    TestAssert GetDataSourceLabel("C:\lib\shared.accda") = "C:\lib\shared.accda", "library path is its own label"
End Sub


Public Sub TestLinkedDataSourcesIncludeUnlinked()

    Dim colSources As Collection
    Dim varSource As Variant
    Dim dSource As Dictionary
    Dim lngTableMissing As Long
    Dim lngEnvMissing As Long
    Dim lngLibraryMissing As Long
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    ClearConnState
    RecordUnlinkedSource "table", ";DATABASE=C:\nonexistent\vcs_datasource\be.accdb"
    RecordUnlinkedSource "table", ";DATABASE=C:\nonexistent\vcs_datasource\be.accdb"
    RecordUnlinkedSource "table", vbNullString, "conn_vcs_datasource_test"
    RecordUnlinkedSource "table", "ODBC;", "conn_vcs_datasource_test"
    RecordUnlinkedSource "table", vbNullString
    RecordUnlinkedSource "library", "C:\nonexistent\vcs_datasource\lib.accda"

    Set colSources = GetLinkedDataSources
    For Each varSource In colSources
        Set dSource = varSource
        Select Case dSource("source")
            Case "C:\nonexistent\vcs_datasource\be.accdb"
                TestAssert dSource("kind") = "missing" And dSource("type") = "table", "uncopied back end is missing"
                lngTableMissing = dSource("notLinked")
            Case "env:conn_vcs_datasource_test"
                TestAssert dSource("kind") = "missing", "a .env key with no value is missing"
                lngEnvMissing = dSource("notLinked")
            Case "C:\nonexistent\vcs_datasource\lib.accda"
                TestAssert dSource("kind") = "missing" And dSource("type") = "library", "uncopied library is missing"
                lngLibraryMissing = dSource("notLinked")
        End Select
    Next varSource

    TestAssert lngTableMissing = 2, "both tables on the missing back end are counted"
    TestAssert lngEnvMissing = 2, "both tables on the missing .env key are counted"
    TestAssert lngLibraryMissing = 1, "the missing library is listed"

CleanUp:
    On Error Resume Next
    ClearConnState
    If lngErr <> 0 Then TestAssert False, "linked data sources " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub
