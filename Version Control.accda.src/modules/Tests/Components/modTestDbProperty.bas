Attribute VB_Name = "modTestDbProperty"
'---------------------------------------------------------------------------------------
' Module    : modTestDbProperty
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Change-detection tests for DAO database properties, including the
'           : issue #773 case where unsetting a property deletes it.
'           : Run: ?VCS.RunTests("modTestDbProperty")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")

Private Const PHANTOM_PROP As String = "VcsTestIssue773Phantom"
Private Const LIVE_PROP As String = "VcsTestIssue773Live"


'---------------------------------------------------------------------------------------
' Procedure : TestDeletedPropertyCountsAsModified
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : A property that exists in dbs-properties.json but not in the live
'           : database is returned by GetAllFromDB(True). Access deletes a property
'           : when it is unset, so this is the incremental-export miss in #773.
'---------------------------------------------------------------------------------------
'
Public Sub TestDeletedPropertyCountsAsModified()

    Dim cProp As IDbComponent
    Dim dModified As Dictionary
    Dim dFile As Dictionary
    Dim dItem As Dictionary
    Dim strFile As String
    Dim strBackup As String
    Dim strKey As String

    Set cProp = New clsDbProperty
    strFile = cProp.SourceFile
    TestAssert FSO.FileExists(strFile), "dbs-properties.json exists for this project"
    If Not FSO.FileExists(strFile) Then Exit Sub

    strBackup = strFile & ".773test.bak"
    If FSO.FileExists(strBackup) Then RestorePropertyFile strFile, strBackup
    FSO.CopyFile strFile, strBackup, True

    Set dFile = ReadJsonFile(strFile)
    Set dItem = New Dictionary
    dItem.Add "Value", "phantom"
    dItem.Add "Type", dbText
    dFile("Items").Add PHANTOM_PROP, dItem
    WriteFile ConvertToJson(dFile, JSON_WHITESPACE), strFile

    Set cProp = New clsDbProperty
    Set dModified = cProp.GetAllFromDB(True)
    strKey = cProp.SourceFile & ":" & PHANTOM_PROP
    TestAssert dModified.Exists(strKey), "deleted property is in the modified set"

    RestorePropertyFile strFile, strBackup
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestNewPropertyCountsAsModified
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : A newly created database property is returned by GetAllFromDB(True).
'---------------------------------------------------------------------------------------
'
Public Sub TestNewPropertyCountsAsModified()

    Dim cProp As IDbComponent
    Dim dModified As Dictionary
    Dim dbs As DAO.Database
    Dim prp As DAO.Property
    Dim strKey As String

    DeleteTestLiveProperty
    Set dbs = SharedDb
    Set prp = dbs.CreateProperty(LIVE_PROP, dbText, "live")
    dbs.Properties.Append prp

    Set cProp = New clsDbProperty
    Set dModified = cProp.GetAllFromDB(True)
    strKey = cProp.SourceFile & ":" & LIVE_PROP
    TestAssert dModified.Exists(strKey), "new property is in the modified set"

    DeleteTestLiveProperty
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPhantomGoneAfterRestore
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : After the phantom key is removed from the source file, a fresh
'           : GetAllFromDB(True) no longer reports it.
'---------------------------------------------------------------------------------------
'
Public Sub TestPhantomGoneAfterRestore()

    Dim cProp As IDbComponent
    Dim dModified As Dictionary
    Dim strKey As String

    Set cProp = New clsDbProperty
    Set dModified = cProp.GetAllFromDB(True)
    strKey = cProp.SourceFile & ":" & PHANTOM_PROP
    TestAssert Not dModified.Exists(strKey), "phantom property is not reported after restore"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RestorePropertyFile
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Put dbs-properties.json back from the byte-identical backup.
'---------------------------------------------------------------------------------------
'
Private Sub RestorePropertyFile(strFile As String, strBackup As String)
    If FSO.FileExists(strBackup) Then
        FSO.CopyFile strBackup, strFile, True
        FSO.DeleteFile strBackup
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DeleteTestLiveProperty
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Remove the live test property if a previous run left it behind.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteTestLiveProperty()
    On Error Resume Next
    SharedDb.Properties.Delete LIVE_PROP
    Err.Clear
End Sub
