Attribute VB_Name = "ExampleTestModule"
'@TestModule
'@Folder "Tests"
Option Compare Database
Option Explicit

Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateTestAssert()
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
End Sub

'@TestCleanup
Public Sub TestCleanup()
End Sub

'@TestMethod("Example")
Public Sub TestAddition()
    Assert.AreEqual 2&, 1& + 1&, "one plus one should equal two"
End Sub
