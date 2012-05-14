Option Compare Database

' Access Module `AppCodeImportExport`
' -----------------------------------
'
' Brendan Kidwell
' This code is licensed under BSD-style terms.
'
' This is some code for importing and exporting Access Queries, Forms,
' Reports, Macros, and Modules to and from plain text files, for the
' purpose of syncing with a version control system.
'
'
' Use:
'
' BACKUP YOUR WORK BEFORE TRYING THIS CODE!
'
' To create and/or overwrite source text files for all database objects
' (except tables) in "$database-folder/source/", run
' `ExportAllSource()`.
'
' To load and/or overwrite  all database objects from source files in
' "$database-folder/source/", run `ImportAllSource()`.
'
'
' Future expansion:
' * Maybe integrate into a dialog box triggered by a menu item.
' * Warning of destructive overwrite.


' --------------------------------
' Begin declarations for ShellWait
' --------------------------------

Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long

' --------------------------------
' End declarations for ShellWait
' --------------------------------




Private Sub MkDirIfNotexist(path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

Private Sub ClearTextFilesFromDir(path As String, Optional doUCS2 As Boolean = True, Optional doUTF8 As Boolean = True)

If doUCS2 Then
    On Error GoTo ClearTextFilesFromDir_noop
    Kill path & "\*.data"
ClearTextFilesFromDir_noop:
End If

If doUTF8 Then
    On Error GoTo ClearTextFilesFromDir_noop2
    Kill path & "\*.txt"
ClearTextFilesFromDir_noop2:
End If

On Error GoTo 0
End Sub

Private Sub SanitizeTextFiles(path As String)

Dim fso, infile, outfile, Filename As String, txt As String

Dim ForReading As Long

ForReading = 1
Set fso = CreateObject("Scripting.FileSystemObject")

Filename = Dir(path & "\*.txt")
Do
    obj_name = Mid(Filename, 1, Len(Filename) - 4)

    Set infile = fso.OpenTextFile(path & "\" & obj_name & ".txt", ForReading)
    Set outfile = fso.CreateTextFile(path & "\" & obj_name & ".sanitize", True)
    Do Until infile.AtEndOfStream
        txt = infile.ReadLine
        If Left(txt, 10) = "Checksum =" Then
            ' Skip lines starting with Checksum
        ElseIf InStr(txt, "NoSaveCTIWhenDisabled =1") Then
            ' Skip lines containning NoSaveCTIWhenDisabled
        ElseIf InStr(txt, "PrtDevNames = Begin") > 0 Or _
            InStr(txt, "PrtDevNamesW = Begin") > 0 Or _
            InStr(txt, "PrtDevModeW = Begin") > 0 Or _
            InStr(txt, "PrtDevMode = Begin") > 0 Then

            ' skip this block of code
            Do Until infile.AtEndOfStream
                txt = infile.ReadLine
                If InStr(txt, "End") Then Exit Do
            Loop
        Else
            outfile.WriteLine txt
        End If
    Loop
    outfile.Close
    infile.Close

    Filename = Dir()
Loop Until Len(Filename) = 0

Filename = Dir(path & "\*.txt")
Do
    obj_name = Mid(Filename, 1, Len(Filename) - 4)
    Kill path & "\" & obj_name & ".txt"
    Name path & "\" & obj_name & ".sanitize" As path & "\" & obj_name & ".txt"
    Filename = Dir()
Loop Until Len(Filename) = 0


End Sub

Public Sub ExportAllSource()

Dim db As Database
Dim source_path As String
Dim obj_path As String
Dim qry As QueryDef
Dim doc As Document
Dim obj_type As Variant
Dim obj_type_split() As String
Dim obj_type_label As String
Dim obj_type_name As String
Dim obj_type_num As Integer

Set db = CurrentDb

source_path = CurrentProject.path
If Right(source_path, 1) <> "\" Then source_path = source_path & "\"
source_path = source_path & "source"
MkDirIfNotexist source_path

Debug.Print

obj_path = source_path & "\queries"
MkDirIfNotexist obj_path
ClearTextFilesFromDir obj_path
Debug.Print "Exporting queries..."
For Each qry In db.QueryDefs
    If Left(qry.Name, 1) <> "~" Then
        Application.SaveAsText acQuery, qry.Name, obj_path & "\" & qry.Name & ".data"
    End If
Next

For Each obj_type In Split( _
    "forms|Forms|" & acForm & "," & _
    "reports|Reports|" & acReport & "," & _
    "macros|Scripts|" & acMacro & "," & _
    "modules|Modules|" & acModule _
    , "," _
)
    obj_type_split = Split(obj_type, "|")
    obj_type_label = obj_type_split(0)
    obj_type_name = obj_type_split(1)
    obj_type_num = Val(obj_type_split(2))
    obj_path = source_path & "\" & obj_type_label
    MkDirIfNotexist obj_path
    ClearTextFilesFromDir obj_path
    Debug.Print "Exporting " & obj_type_label & "..."
    For Each doc In db.Containers(obj_type_name).Documents
        If Left(doc.Name, 1) <> "~" Then
            Application.SaveAsText obj_type_num, doc.Name, obj_path & "\" & doc.Name & ".data"
        End If
    Next
Next

ShellWait CurrentProject.path & "\scripts\ucs2-to-utf8.bat", vbNormalFocus

Debug.Print "Removing Checksum and NoSaveCTIWhenDisabled lines"
For Each obj_type In Split("forms,reports,macros", ",")
    SanitizeTextFiles source_path & "\" & obj_type
Next

Debug.Print "Done."

End Sub

Public Sub ImportAllSource()

Dim db As Database
Dim source_path As String
Dim obj_path As String
Dim qry As QueryDef
Dim doc As Document
Dim obj_type As Variant
Dim obj_type_split() As String
Dim obj_type_label As String
Dim obj_type_name As String
Dim obj_type_num As Integer
Dim Filename As String
Dim obj_name As String

ShellWait CurrentProject.path & "\scripts\utf8-to-ucs2.bat", vbNormalFocus

Set db = CurrentDb

source_path = CurrentProject.path
If Right(source_path, 1) <> "\" Then source_path = source_path & "\"
source_path = source_path & "source"
MkDirIfNotexist source_path

Debug.Print

obj_path = source_path & "\queries"
Debug.Print "Importing queries..."
Filename = Dir(obj_path & "\*.data")
Do
    obj_name = Mid(Filename, 1, Len(Filename) - 5)
    Application.LoadFromText acQuery, obj_name, obj_path & "\" & Filename
    Filename = Dir()
Loop Until Len(Filename) = 0
ClearTextFilesFromDir obj_path, True, False

For Each obj_type In Split( _
    "forms|" & acForm & "," & _
    "reports|" & acReport & "," & _
    "macros|" & acMacro & "," & _
    "modules|" & acModule _
    , "," _
)
    obj_type_split = Split(obj_type, "|")
    obj_type_label = obj_type_split(0)
    obj_type_num = Val(obj_type_split(1))
    obj_path = source_path & "\" & obj_type_label
    Debug.Print "Importing " & obj_type_label & "..."
    Filename = Dir(obj_path & "\*.data")
    Do
        obj_name = Mid(Filename, 1, Len(Filename) - 5)
        If obj_name <> "AppCodeImportExport" Then
            Application.LoadFromText obj_type_num, obj_name, obj_path & "\" & Filename
        End If
        Filename = Dir()
    Loop Until Len(Filename) = 0
    ClearTextFilesFromDir obj_path, True, False
Next

Debug.Print "Done."

End Sub



'***************** Code Start ******************
'http://access.mvps.org/access/api/api0004.htm
'
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Terry Kreft
Public Sub ShellWait(Pathname As String, Optional WindowStyle As Long)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    ' Start the shelled application:
    ret& = CreateProcessA(0&, Pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)
End Sub
'***************** Code End ****************