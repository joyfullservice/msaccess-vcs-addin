Attribute VB_Name = "modTestRunVBA"
'---------------------------------------------------------------------------------------
' Module    : modTestRunVBA
' Author    : Adam Waller
' Date      : 8/27/2026
' Purpose   : Unit tests for RunVBA temporary-module identification and safe sweeping.
'           : These tests never create or remove a component; reset and mutation behavior
'           : is covered by the MCP integration sequence against Testing.accdb.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestMcpTempModuleNameRequiresPrefixAndStandardModule()

    Dim cVcs As clsVersionControl
    Set cVcs = VCS

    TestAssert cVcs.IsMcpTempModuleName("MCP_Temp_12345", vbext_ct_StdModule), _
        "RunVBA standard module recognized"
    TestAssert cVcs.IsMcpTempModuleName("mcp_temp_12345", vbext_ct_StdModule), _
        "prefix comparison is case-insensitive"
    TestAssert Not cVcs.IsMcpTempModuleName("MCP_Temp_12345", vbext_ct_ClassModule), _
        "class module with matching prefix is not owned by RunVBA"
    TestAssert Not cVcs.IsMcpTempModuleName("MCP_Temporary", vbext_ct_StdModule), _
        "similar prefix is not treated as a RunVBA module"

End Sub


Public Sub TestMcpTempModuleSweepIsNoOpWhenNothingIsStale()

    Dim cVcs As clsVersionControl
    Dim colRemoved As Collection
    Dim colRemaining As Collection
    Dim intCountBefore As Long

    Set cVcs = VCS
    intCountBefore = CurrentVBProject.VBComponents.Count

    TestAssert cVcs.SweepMcpTempModules(colRemoved, colRemaining), _
        "project without stale RunVBA modules reports clean"
    TestAssert colRemoved.Count = 0, "no modules removed"
    TestAssert colRemaining.Count = 0, "no orphan modules remain"
    TestAssert CurrentVBProject.VBComponents.Count = intCountBefore, _
        "no-op sweep does not change component count"

End Sub
