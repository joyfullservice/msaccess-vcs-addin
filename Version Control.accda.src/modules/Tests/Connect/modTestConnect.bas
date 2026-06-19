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

    ' Below 5.1.0: behavior unchanged - credentials pass through to source.
    Options.ExportFormatVersion = EFV_5_0_0
    strResult = GetSourceSafeConnect(strConn, "test (linked table)")
    TestAssert InStr(strResult, "PWD=secret") > 0, "pre-5.1.0 leaves password untouched"

    ' 5.1.0+: real password stripped from anything bound for source.
    Options.ExportFormatVersion = EFV_5_1_0
    strResult = GetSourceSafeConnect(strConn, "test (linked table)")
    TestAssert InStr(1, strResult, "PWD=", vbTextCompare) = 0, "5.1.0 strips PWD"
    TestAssert InStr(strResult, "secret") = 0, "5.1.0 removes password value"

    ' 5.1.0+ with passwordless auth (empty PWD): no secret, returned unchanged.
    strResult = GetSourceSafeConnect(strAd, "test (linked table)")
    TestAssert strResult = strAd, "passwordless AD connection is not altered"

    ' 5.1.0+ with no credentials at all: returned unchanged.
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
