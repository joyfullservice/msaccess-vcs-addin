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


'---------------------------------------------------------------------------------------
' Procedure : InitializeVersionControlSystem
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Attempt to relink the MSAccess-VCS library from the current folder
'           : if the reference is reported as broken. This allows you to either specify
'           : a location that stays the same across projects, or to use the relative
'           : path of the current folder.
'---------------------------------------------------------------------------------------
'
Public Sub InitializeVersionControlSystem()

    Const VB_PROJECT As Integer = 1

    Dim ref As Reference
    Dim intCnt As Integer
    Dim strPath As String
    Dim strFile As String
    Dim varParts As Variant
    
    ' Loop backwards through references, since libraries will be near the end.
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If ref.Kind = VB_PROJECT Then
            If ref.Name = "MSAccess-VCS" Then
                ' Make sure the path is valid.
                If Dir(ref.FullPath) = "" Then
                    ' Could not find at current path. Get file name
                    varParts = Split(ref.FullPath, "\")
                    strFile = varParts(UBound(varParts))
                    ' Check in current folder
                    strPath = CodeProject.Path & "\" & strFile
                    If Dir(strPath) <> "" Then
                        ' Found in current folder. Relink at this location
                        Application.References.Remove ref
                        Set ref = Nothing
                        Application.References.AddFromFile strFile
                        Debug.Print "Updated Version Control System reference to: " & ref.FullPath
                    Else
                        MsgBox "Unable to find required reference for Version Control System" & _
                            vbCrLf & ref.FullPath & vbCrLf & "Please contact your systems administrator for assistance.", vbExclamation
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next intCnt
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetAllClassesPublic
' Author    : Adam Waller
' Date      : 4/2/2015
' Purpose   : Set object classes to public
'---------------------------------------------------------------------------------------
'
Public Sub SetAllClassesPublic()
    Const vbext_ct_ClassModule = 2
    Const PUBLIC_CREATABLE As Integer = 5
    Dim cls As Object ' VBComponent
    For Each cls In VBE.ActiveVBProject.VBComponents
        If cls.Type = vbext_ct_ClassModule Then
            cls.Properties("Instancing") = PUBLIC_CREATABLE
        End If
    Next cls
End Sub