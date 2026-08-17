Attribute VB_Name = "modTestLetterCasing"
'---------------------------------------------------------------------------------------
' Module    : modTestLetterCasing
' Author    : Adam Waller
' Date      : 8/17/2026
' Purpose   : Unit tests for modLetterCasing Lib-clause extraction.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")


Public Sub TestExtractDeclareLibNameKernel32()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz_kernel32 Lib ""kernel32"" () 'kernel32") = "kernel32", _
        "kernel32 without extension"
End Sub


Public Sub TestExtractDeclareLibNameKernel32Dll()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz_kernel32_dll Lib ""kernel32.dll"" () 'kernel32.dll") = "kernel32.dll", _
        "kernel32.dll with extension"
End Sub


Public Sub TestExtractDeclareLibNamePreservesCasing()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz Lib ""KERNEL32"" () 'kernel32") = "KERNEL32", _
        "quoted text casing preserved"
End Sub


Public Sub TestExtractDeclareLibNameMissingQuotes()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz_kernel32 Lib kernel32 () 'kernel32") = vbNullString, _
        "no quotes"
End Sub


Public Sub TestExtractDeclareLibNameEmptyQuotes()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz Lib """" () 'kernel32") = vbNullString, _
        "empty quotes"
End Sub


Public Sub TestExtractDeclareLibNameUnclosedQuote()
    TestAssert ExtractDeclareLibName( _
        "Private Declare PtrSafe Function zzz Lib ""kernel32 () 'kernel32") = vbNullString, _
        "unclosed quote"
End Sub
