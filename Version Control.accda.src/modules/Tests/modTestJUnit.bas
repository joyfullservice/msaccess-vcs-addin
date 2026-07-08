Attribute VB_Name = "modTestJUnit"
'---------------------------------------------------------------------------------------
' Module    : modTestJUnit
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Project the durable test-state.json into JUnit XML for CI consumption
'           : (GitLab native reports, GitHub Actions via third-party reporters).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Private Const ModuleName As String = "modTestJUnit"
Private Const JUNIT_FILE As String = "test-results.xml"


'---------------------------------------------------------------------------------------
' Procedure : ExportFromState
' Author    : Adam Waller
' Date      : 7/8/2026
' Purpose   : Write JUnit XML from the current test-state.json. Callable without
'           : running tests (projection of durable state).
'---------------------------------------------------------------------------------------
'
Public Function ExportFromState(Optional ByVal strPath As String = vbNullString) As String

    Dim dRoot As Dictionary
    Dim dTests As Dictionary
    Dim dSuites As Dictionary
    Dim varKey As Variant
    Dim dTest As Dictionary
    Dim strModule As String
    Dim colSuiteTests As Collection
    Dim dSuite As Dictionary
    Dim strStatus As String
    Dim dblTotalSec As Double
    Dim lngTotalTests As Long
    Dim lngTotalFailures As Long
    Dim lngTotalErrors As Long
    Dim strXml As String

    If Len(strPath) = 0 Then
        strPath = modTestState.GetTestResultsFolder() & JUNIT_FILE
    End If

    Set dRoot = modTestState.LoadState()
    If dRoot Is Nothing Then
        ExportFromState = vbNullString
        Exit Function
    End If
    If Not dRoot.Exists("tests") Then Exit Function
    If TypeName(dRoot("tests")) <> "Dictionary" Then Exit Function

    Set dTests = dRoot("tests")
    Set dSuites = New Dictionary

    For Each varKey In dTests.Keys
        Set dTest = dTests(CStr(varKey))
        strStatus = UCase$(CStr(Nz(dTest("status"), "PENDING")))
        If strStatus = "PENDING" Then GoTo NextJUnitTest

        strModule = CStr(dTest("moduleName"))
        If Not dSuites.Exists(strModule) Then
            Set dSuite = New Dictionary
            dSuite.Add "name", strModule
            dSuite.Add "tests", 0&
            dSuite.Add "failures", 0&
            dSuite.Add "errors", 0&
            dSuite.Add "time", 0#
            Set colSuiteTests = New Collection
            Set dSuite("cases") = colSuiteTests
            Set dSuites(strModule) = dSuite
        End If

        Set dSuite = dSuites(strModule)
        Set colSuiteTests = dSuite("cases")
        colSuiteTests.Add dTest
        dSuite("tests") = CLng(dSuite("tests")) + 1
        dSuite("time") = CDbl(dSuite("time")) + (CLng(Nz(dTest("durationMs"), 0)) / 1000#)

        Select Case strStatus
            Case "FAILED"
                dSuite("failures") = CLng(dSuite("failures")) + 1
                lngTotalFailures = lngTotalFailures + 1
            Case "ERRORED"
                dSuite("errors") = CLng(dSuite("errors")) + 1
                lngTotalErrors = lngTotalErrors + 1
        End Select

        lngTotalTests = lngTotalTests + 1
        dblTotalSec = dblTotalSec + (CLng(Nz(dTest("durationMs"), 0)) / 1000#)
NextJUnitTest:
    Next varKey

    strXml = BuildJUnitXml(dSuites, lngTotalTests, lngTotalFailures, lngTotalErrors, dblTotalSec)
    WriteFile strXml, strPath
    ExportFromState = strPath

End Function


' ===================== Private helpers =====================


Private Function BuildJUnitXml(ByVal dSuites As Dictionary, _
    ByVal lngTotalTests As Long, ByVal lngTotalFailures As Long, _
    ByVal lngTotalErrors As Long, ByVal dblTotalSec As Double) As String

    Dim strXml As String
    Dim varKey As Variant
    Dim dSuite As Dictionary
    Dim colCases As Collection
    Dim dTest As Dictionary
    Dim i As Long

    strXml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strXml = strXml & "<testsuites tests=""" & lngTotalTests & """ failures=""" & lngTotalFailures & _
        """ errors=""" & lngTotalErrors & """ time=""" & FormatSeconds(dblTotalSec) & """>" & vbCrLf

    For Each varKey In dSuites.Keys
        Set dSuite = dSuites(CStr(varKey))
        Set colCases = dSuite("cases")
        strXml = strXml & "  <testsuite name=""" & EscapeXml(CStr(dSuite("name"))) & """ tests=""" & _
            CLng(dSuite("tests")) & """ failures=""" & CLng(dSuite("failures")) & _
            """ errors=""" & CLng(dSuite("errors")) & """ time=""" & _
            FormatSeconds(CDbl(dSuite("time"))) & """>" & vbCrLf

        For i = 1 To colCases.Count
            Set dTest = colCases(i)
            strXml = strXml & BuildTestCaseXml(dTest)
        Next i

        strXml = strXml & "  </testsuite>" & vbCrLf
    Next varKey

    strXml = strXml & "</testsuites>" & vbCrLf
    BuildJUnitXml = strXml

End Function


Private Function BuildTestCaseXml(ByVal dTest As Dictionary) As String

    Dim strModule As String
    Dim strProc As String
    Dim strStatus As String
    Dim dblSec As Double
    Dim strMessage As String
    Dim strDetail As String
    Dim strLine As String

    strModule = CStr(dTest("moduleName"))
    strProc = CStr(dTest("procName"))
    strStatus = UCase$(CStr(Nz(dTest("status"), "PENDING")))
    dblSec = CLng(Nz(dTest("durationMs"), 0)) / 1000#

    strLine = "    <testcase classname=""" & EscapeXml(strModule) & """ name=""" & _
        EscapeXml(strProc) & """ time=""" & FormatSeconds(dblSec) & """"

    Select Case strStatus
        Case "FAILED", "ERRORED", "EMPTY"
            strMessage = CStr(Nz(dTest("errorMessage"), strStatus))
            strDetail = BuildFailureDetail(dTest)
            If strStatus = "ERRORED" Then
                BuildTestCaseXml = strLine & ">" & vbCrLf & _
                    "      <error message=""" & EscapeXml(strMessage) & """>" & _
                    EscapeXml(strDetail) & "</error>" & vbCrLf & _
                    "    </testcase>" & vbCrLf
            ElseIf strStatus = "EMPTY" Then
                BuildTestCaseXml = strLine & ">" & vbCrLf & _
                    "      <skipped message=""" & EscapeXml(strMessage) & """/>" & vbCrLf & _
                    "    </testcase>" & vbCrLf
            Else
                BuildTestCaseXml = strLine & ">" & vbCrLf & _
                    "      <failure message=""" & EscapeXml(strMessage) & """>" & _
                    EscapeXml(strDetail) & "</failure>" & vbCrLf & _
                    "    </testcase>" & vbCrLf
            End If
        Case Else
            BuildTestCaseXml = strLine & "/>" & vbCrLf
    End Select

End Function


Private Function BuildFailureDetail(ByVal dTest As Dictionary) As String

    Dim colAssertions As Collection
    Dim dA As Dictionary
    Dim i As Long
    Dim strDetail As String

    If Not dTest.Exists("assertions") Then Exit Function
    If TypeName(dTest("assertions")) <> "Collection" Then Exit Function

    Set colAssertions = dTest("assertions")
    For i = 1 To colAssertions.Count
        Set dA = colAssertions(i)
        If Not CBool(Nz(dA("passed"), True)) Then
            If Len(CStr(Nz(dA("context"), vbNullString))) > 0 Then
                strDetail = strDetail & CStr(dA("context")) & vbCrLf
            Else
                strDetail = strDetail & "Assertion " & CStr(dA("seq")) & " failed" & vbCrLf
            End If
        End If
    Next i

    BuildFailureDetail = strDetail

End Function


Private Function FormatSeconds(ByVal dblSec As Double) As String

    FormatSeconds = Replace$(Format$(dblSec, "0.000"), ",", ".")

End Function


Private Function EscapeXml(ByVal strText As String) As String

    Dim strOut As String

    strOut = strText
    strOut = Replace$(strOut, "&", "&amp;")
    strOut = Replace$(strOut, "<", "&lt;")
    strOut = Replace$(strOut, ">", "&gt;")
    strOut = Replace$(strOut, """", "&quot;")
    strOut = Replace$(strOut, "'", "&apos;")
    EscapeXml = strOut

End Function
