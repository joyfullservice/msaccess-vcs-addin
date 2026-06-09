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
