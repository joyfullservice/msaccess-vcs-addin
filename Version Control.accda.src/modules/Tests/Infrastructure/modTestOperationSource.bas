Attribute VB_Name = "modTestOperationSource"
'---------------------------------------------------------------------------------------
' Module    : modTestOperationSource
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Tests for clsOperation.AutomationSource and SourceName. Each test
'           : caches and restores Operation.Source because the suite itself may be
'           : running under eosExternalAPI (MCP / Application.Run).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestAutomationSource_UserInterface()
    Dim ePrior As eOperationSource
    ePrior = Operation.Source
    Operation.Source = eosUserInterface
    TestAssert Not Operation.AutomationSource, "UI source is not automation"
    Operation.Source = ePrior
End Sub


Public Sub TestAutomationSource_ExternalAPI()
    Dim ePrior As eOperationSource
    ePrior = Operation.Source
    Operation.Source = eosExternalAPI
    TestAssert Operation.AutomationSource, "External API source is automation"
    Operation.Source = ePrior
End Sub


Public Sub TestAutomationSource_MCPTool()
    Dim ePrior As eOperationSource
    ePrior = Operation.Source
    Operation.Source = eosMCPTool
    TestAssert Operation.AutomationSource, "MCP tool source is automation"
    Operation.Source = ePrior
End Sub


Public Sub TestSourceName_UserInterfaceIsEmpty()
    Dim ePrior As eOperationSource
    ePrior = Operation.Source
    Operation.Source = eosUserInterface
    TestAssert Len(Operation.SourceName) = 0, "UI source name is empty (no log heading)"
    Operation.Source = ePrior
End Sub
