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


Public Sub TestGetConnectionEnvKey()
    Dim strKey As String

    ' Access back-end: should use filename as key identity
    strKey = GetConnectionEnvKey(";DATABASE=C:\Data\MyBackEnd.accdb")
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
