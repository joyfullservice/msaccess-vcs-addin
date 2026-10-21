Attribute VB_Name = "vcs_test_casing_holder"
Option Compare Database
Option Explicit

' Fixture for the code hash tests: uses an identifier that another fixture declares
' with a different case. The word Probe is also in a comment and in a string.
Private vcscasingprobe As Long

Public Function vcsTestCasingHolderValue() As String
    vcscasingprobe = 2
    vcsTestCasingHolderValue = "Probe " & CStr(vcscasingprobe)
End Function
