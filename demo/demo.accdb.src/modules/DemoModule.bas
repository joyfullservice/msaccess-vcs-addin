Option Compare Database
Option Explicit

Public Function DemoHello()
Debug.Print "Hello world."
End Function



'---------------------------------------------------------------------------------------
' Procedure : StartUp
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Code to run when the database opens. (Include a call to this function
'           : in a macro named "AutoExec", and it will run automatically.)
'---------------------------------------------------------------------------------------
'
Public Function StartUp()
    InitializeVersionControlSystem
End Function