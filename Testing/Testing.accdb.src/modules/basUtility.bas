Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : RunAfterBuild
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : Compile and run testing after build.
'---------------------------------------------------------------------------------------
'
Public Sub RunAfterBuild()

    ' Compile and save VBA code. Should prompt for any errors here.
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestRegEx
' Author    : Adam Waller
' Date      : 5/1/2020
' Purpose   : This tests VBA code that uses early binding with non-built-in
'           : VBE References. (Add references before importing code)
'---------------------------------------------------------------------------------------
'
Public Sub TestRegEx()

    Dim regEx As VBScript_RegExp_55.RegExp
    
End Sub