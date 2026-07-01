Attribute VB_Name = "modTestCommandBarNames"
'---------------------------------------------------------------------------------------
' Module    : modTestCommandBarNames
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Tests for command bar control classification. The runtime addability probe
'           : (IsBuiltInControlAddable) is the classifier; IsNonAddableControl is an
'           : empty manual override. Probe tests are tagged slow/access because
'           : addability is tested against the running Access version.
'           : Run: ?VCS.RunTests("modTestCommandBarNames")
'           : Skip on older Access: ?VCS.RunTests("-access") or ?VCS.RunTests("-slow")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")


Public Sub TestNonAddableOverrideEmptyByDefault()
    ' The manual override is empty by default; the probe handles classification.
    TestAssert Not IsNonAddableControl(1839), "About (1839) not force-overridden"
    TestAssert Not IsNonAddableControl(3227), "Edit Hyperlink (3227) not force-overridden"
    TestAssert Not IsNonAddableControl(4), "Print (4) not force-overridden"
End Sub


Public Sub TestKnownNonAddableControls()
    '@Tag("slow")
    '@Tag("access")
    ' Access "custom built-in" and discontinued controls that cannot be recreated by Id.
    AssertNotAddable msoControlButton, 1835, "Open Table (1835)"
    AssertNotAddable msoControlButton, 1839, "About (1839)"
    AssertNotAddable msoControlEdit, 3227, "Edit Hyperlink (3227)"
    AssertNotAddable msoControlButton, 3887, "Diagrams (3887)"
    AssertNotAddable msoControlButton, 11108, "Sort Descending context (11108)"
End Sub


Public Sub TestKnownAddableControls()
    '@Tag("slow")
    '@Tag("access")
    ' Ids used as minimal built-in exports in the Menus round-trip corpus.
    AssertAddable msoControlButton, 4, "Print (4)"
    AssertAddable msoControlButton, 106, "Close (106)"
    AssertAddable msoControlButton, 247, "Page Setup (247)"
    AssertAddable msoControlButton, 752, "Exit (752)"
    AssertAddable msoControlButton, 211, "Sort Descending (211)"
    AssertAddable msoControlButton, 2071, "Compact and Repair (2071)"
End Sub


Public Sub TestAddableCacheMatchesFreshProbe()
    '@Tag("slow")
    '@Tag("access")
    ' The cached classifier must agree with a fresh, uncached probe.
    ResetAddableCache
    TestAssert IsBuiltInControlAddable(msoControlButton, 4) = ProbeBuiltInControlAddable(msoControlButton, 4), _
        "cached Print (4) matches fresh probe"
    TestAssert IsBuiltInControlAddable(msoControlEdit, 3227) = ProbeBuiltInControlAddable(msoControlEdit, 3227), _
        "cached Edit Hyperlink (3227) matches fresh probe"
    ' Second call hits the cache and returns the same answer.
    TestAssert IsBuiltInControlAddable(msoControlButton, 4) = ProbeBuiltInControlAddable(msoControlButton, 4), _
        "cached Print (4) stable on repeat call"
End Sub


Public Sub TestTemplateControlOpenTable()
    '@Tag("slow")
    '@Tag("access")
    ' Non-addable object-openers route through the template bar, not replicas.
    Dim cbTemplate As CommandBar

    LogUnhandledErrors
    On Error Resume Next
    Set cbTemplate = Application.CommandBars(strTemplateCommandBarName)
    If CatchAny(eelNoError, vbNullString) Or cbTemplate Is Nothing Then
        TestAssert True, "SKIP: template command bar not loaded"
        Exit Sub
    End If

    TestAssert Not (cbTemplate.FindControl(Id:=1835) Is Nothing), "Open Table (1835) on template bar"
    TestAssert Not IsBuiltInControlAddable(msoControlButton, 1835), "Open Table (1835) not addable"
End Sub


Private Sub AssertNotAddable(lngType As Long, lngId As Long, strLabel As String)
    TestAssert Not IsBuiltInControlAddable(lngType, lngId), strLabel & " should not be .Add-able"
End Sub


Private Sub AssertAddable(lngType As Long, lngId As Long, strLabel As String)
    TestAssert IsBuiltInControlAddable(lngType, lngId), strLabel & " should be .Add-able on this Access version"
End Sub
