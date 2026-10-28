VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestPrinterSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestPrinterSettings
' Author    : Adam Waller
' Date      : 9/16/2026
' Purpose   : Regression tests for restoring report/form printer settings.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Core")
'@Tag("unit")

Private Const MISSING_PRINTER As String = "VCS Test Missing Printer 784"
Private Const TEST_OBJECT As String = "rptPrinterSettingsTest"

Private m_Settings As clsPrinterSettings


Private Sub Class_Initialize()
    Set m_Settings = New clsPrinterSettings
End Sub


Private Sub Class_Terminate()
    Set m_Settings = Nothing
End Sub


Public Sub TestSpecificBindingToCurrentDefaultRemainsSpecific()

    Dim dActual As Dictionary
    Dim dDevice As Dictionary
    Dim strPrinter As String

    TestAssert Access.Printers.Count > 0, "test requires an installed default printer"
    If Access.Printers.Count = 0 Then Exit Sub

    strPrinter = Application.Printer.DeviceName
    m_Settings.ApplySettings BuildSettings(strPrinter), TEST_OBJECT
    Set dActual = m_Settings.GetDictionary

    TestAssert dActual.Exists("Device"), _
        "an explicit binding remains specific when the device is also the Windows default"
    If Not dActual.Exists("Device") Then Exit Sub

    Set dDevice = dActual("Device")
    TestAssert dDevice("DeviceName") = strPrinter, "specific device name retained"
    TestAssert dDevice("Default") = False, "specific binding is not encoded as default"

End Sub


Public Sub TestMissingPrinterPreservesPortableSettings()

    Dim dActual As Dictionary
    Dim dPrinter As Dictionary

    TestAssert Access.Printers.Count > 0, "test requires an installed default printer"
    If Access.Printers.Count = 0 Then Exit Sub

    m_Settings.ApplySettings BuildSettings(MISSING_PRINTER), TEST_OBJECT
    Set dActual = m_Settings.GetDictionary
    Set dPrinter = dActual("Printer")

    TestAssert Not dActual.Exists("Device"), "missing specific device falls back to default binding"
    TestAssert dPrinter("Orientation") = "Landscape", "orientation retained during fallback"
    TestAssert dPrinter("PaperSize") = "A4", "paper size retained during fallback"

End Sub


Public Sub TestMissingPrinterWarningNamesObject()

    Dim lngStart As Long
    Dim colLogged As Collection
    Dim varEntry As Variant
    Dim dEntry As Dictionary
    Dim blnFound As Boolean

    TestAssert Access.Printers.Count > 0, "test requires an installed default printer"
    If Access.Printers.Count = 0 Then Exit Sub

    lngStart = Log.ErrorJournalCount
    m_Settings.ApplySettings BuildSettings(MISSING_PRINTER), TEST_OBJECT
    Set colLogged = Log.GetErrorJournalSince(lngStart)

    For Each varEntry In colLogged
        Set dEntry = varEntry
        If CStr(dEntry("source")) = "clsPrinterSettings.LoadFromPrinter" _
            And InStr(1, CStr(dEntry("message")), TEST_OBJECT, vbBinaryCompare) > 0 Then
            blnFound = True
            Exit For
        End If
    Next varEntry

    TestAssert blnFound, "missing-printer warning identifies the affected object"

End Sub


Public Sub TestPrinterRestoreSummaryListsUniqueObjects()

    Dim dAffected As Dictionary
    Dim strSummary As String

    Set dAffected = New Dictionary
    dAffected.Add "rptFirst", True
    dAffected.Add "rptSecond", True

    strSummary = modBuild.FormatPrinterRestoreWarningSummary(dAffected)
    TestAssert InStr(strSummary, "default printer substituted") > 0, "fallback explained"
    TestAssert InStr(strSummary, "rptFirst, rptSecond") > 0, "unique object names listed"

End Sub


Private Function BuildSettings(strDeviceName As String) As Dictionary

    Dim dItems As Dictionary
    Dim dDevice As Dictionary
    Dim dPrinter As Dictionary
    Dim dMargins As Dictionary

    Set dItems = New Dictionary
    Set dDevice = New Dictionary
    dDevice.Add "DeviceName", strDeviceName
    dDevice.Add "Default", False
    dItems.Add "Device", dDevice

    Set dPrinter = New Dictionary
    dPrinter.Add "Orientation", "Landscape"
    dPrinter.Add "PaperSize", "A4"
    dItems.Add "Printer", dPrinter

    Set dMargins = New Dictionary
    dItems.Add "Margins", dMargins

    Set BuildSettings = dItems

End Function
