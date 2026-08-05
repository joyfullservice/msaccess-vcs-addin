Attribute VB_Name = "basTestHelpers"
'@Folder "Tests"
Option Compare Database
Option Explicit

Public Function CreateTestAssert() As Object
    Set CreateTestAssert = New StubRdAssert
End Function
