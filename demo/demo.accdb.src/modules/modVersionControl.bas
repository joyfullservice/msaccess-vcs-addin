Option Explicit
Option Private Module
Option Compare Database


'---------------------------------------------------------------------------------------
' Procedure : InitializeVersionControlSystem
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Call this function to activate or deactivate version control.
'---------------------------------------------------------------------------------------
'
Public Sub InitializeVersionControlSystem(Optional blnUseVersionControl As Boolean = True)

    '//////////////////////////////////////////////////////////
    ' SET THESE VALUES AS NEEDED FOR YOUR ENVIRONMENT
    '//////////////////////////////////////////////////////////
    
    ' Version Control Library
    Const cstrLibraryPath As String = "\..\"
    Const cstrLibraryFile As String = "Version Control.accda"
    Const cstrLibraryName As String = "MSAccess-VCS"
    
    ' VCS Settings for this database (Additional parameters may be added as needed)
    Dim colParams As New Collection
    With colParams
        .Add Array("System", "GitHub")     ' Set this first, before other settings.
        .Add Array("Export Folder", CurrentProject.Path & "\" & CurrentProject.Name & ".src\")
        .Add Array("Show Debug", True)
        .Add Array("Include VBE", False)
        .Add Array("Fast Save", True)
        .Add Array("Save Print Vars", False)
        '.Add Array("Save Table", "vObjectPermissions")
    End With

    ' Pass the parameters to the wrapper function
    LoadVersionControl blnUseVersionControl, cstrLibraryPath, cstrLibraryFile, cstrLibraryName, colParams

End Sub


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
Private Sub LoadVersionControl(blnUseVersionControl As Boolean, strLibraryPath As String, strLibraryFile As String, strLibraryName As String, colParams As Collection)
    
    Const VB_PROJECT As Integer = 1

    Dim ref As Reference
    Dim intCnt As Integer
    Dim strPath As String
    Dim strFile As String
    Dim varParts As Variant
    Dim blnInitialize As Boolean
    Dim blnLoaded As Boolean
    
    ' Loop backwards through references, since libraries will be near the end.
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If ref.Kind = VB_PROJECT Then
            If ref.Name = strLibraryName Then
                If blnUseVersionControl Then
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
                            blnInitialize = True
                            blnLoaded = True
                        Else
                            MsgBox "Unable to find required reference for Version Control System" & _
                                vbCrLf & ref.FullPath & vbCrLf & "Please contact your systems administrator for assistance.", vbExclamation
                        End If
                        Exit For
                    Else
                        ' Found existing reference
                        blnInitialize = True
                        blnLoaded = True
                    End If
                Else
                    ' Disable version control
                    Run "[" & strLibraryName & "].ReleaseObjectReferences"
                    Application.References.Remove ref
                    Set ref = Nothing
                    Debug.Print "Removed Version Control System"
                    Exit For
                End If
            End If
        End If
    Next intCnt
    
    
    ' Prepare to initialize version control.
    If blnUseVersionControl And Not blnLoaded Then
        ' Attempt to load the file
        If strLibraryPath <> "\..\" And Dir(strLibraryPath, vbDirectory) <> "" Then
            ' Use specified path
            strPath = strLibraryPath
        Else
            ' Use current folder
            strPath = CodeProject.Path & "\..\"
        End If
        
        ' Check to see if the library file exists
        strPath = strPath & strLibraryFile
        If Dir(strPath) <> "" Then
            ' File exists
            If strPath <> CodeDb.Name Then Application.References.AddFromFile strPath
            blnInitialize = True
        End If
        
    End If
    
    ' Initialize the VBE menu
    ' (Use the Run commmand to avoid compile errors if the library was not loaded)
    If blnInitialize Then
        Run "[" & strLibraryName & "].LoadVersionControl", colParams
        '''Run "LoadVersionControl", colParams
    End If
    
End Sub