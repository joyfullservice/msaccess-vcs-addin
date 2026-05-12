VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestConnectionODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestConnectionODBC
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for clsConnectionODBC ODBC connection string parsing.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Connect")

Private m_odbc As clsConnectionODBC


Private Sub Class_Initialize()
    Set m_odbc = New clsConnectionODBC
End Sub


Private Sub Class_Terminate()
    Set m_odbc = Nothing
End Sub


Public Sub TestParseBasicODBC()
    m_odbc.ParseOdbcConnectionString "ODBC;DRIVER={SQL Server};SERVER=mysvr;DATABASE=mydb"
    TestAssert m_odbc.Driver = "{SQL Server}", "Driver parsed"
    TestAssert m_odbc.Server = "mysvr", "Server parsed"
    TestAssert m_odbc.Database = "mydb", "Database parsed"
End Sub


Public Sub TestParseWithCredentials()
    m_odbc.ParseOdbcConnectionString "ODBC;DRIVER={SQL Server};UID=sa;PWD=secret;DATABASE=test"
    TestAssert m_odbc.UID = "sa", "UID parsed"
    TestAssert m_odbc.PWD = "secret", "PWD parsed"
End Sub


Public Sub TestParseWithDSN()
    m_odbc.ParseOdbcConnectionString "ODBC;DSN=MyDataSource"
    TestAssert m_odbc.DSN = "MyDataSource", "DSN parsed"
End Sub


Public Sub TestGetKeyValues()
    m_odbc.ParseOdbcConnectionString "ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=db"
    Dim dKeys As Dictionary
    Set dKeys = m_odbc.GetKeyValues
    TestAssert dKeys.Count > 0, "key-value dictionary non-empty"
    TestAssert dKeys.Exists("DRIVER"), "DRIVER key present"
    TestAssert dKeys.Exists("SERVER"), "SERVER key present"
End Sub


Public Sub TestIsGenericKey()
    TestAssert m_odbc.IsGenericKey("DRIVER"), "DRIVER is generic"
    TestAssert m_odbc.IsGenericKey("DSN"), "DSN is generic"
    TestAssert m_odbc.IsGenericKey("UID"), "UID is generic"
    TestAssert m_odbc.IsGenericKey("PWD"), "PWD is generic"
    TestAssert Not m_odbc.IsGenericKey("CustomDriverKey"), "custom key is not generic"
End Sub


Public Sub TestNonOdbcStringIgnored()
    m_odbc.ParseOdbcConnectionString ";DATABASE=C:\data.accdb"
    TestAssert m_odbc.Database = "", "non-ODBC string not parsed"
End Sub


Public Sub TestOriginalConnectionString()
    Dim strConn As String
    strConn = "ODBC;DRIVER={SQL Server};SERVER=svr;DATABASE=db"
    m_odbc.ParseOdbcConnectionString strConn
    TestAssert m_odbc.OriginalConnectionString = strConn, "original string preserved"
End Sub
