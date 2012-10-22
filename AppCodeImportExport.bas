Option Compare Database

' Access Module `AppCodeImportExport`
' -----------------------------------
'
' https://github.com/bkidwell/msaccess-vcs-integration
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
' See project home page (URL above) for more information.
'
'
' Future expansion:
' * Maybe integrate into a dialog box triggered by a menu item.
' * Warning of destructive overwrite.


' --------------------------------
' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
'
' Provide a comman separated list of table names, or an empty string
' ("") if no tables are to be exported with the source code.
' --------------------------------

Private Const INCLUDE_TABLES = ""

' --------------------------------
' Constants
' --------------------------------

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

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

' --------------------------------
' Beginning of main functions of this module
' --------------------------------

' Create folder `Path`. Silently do nothing if it already exists.
Private Sub MkDirIfNotexist(Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

' Erase all *.data and *.txt files in `Path`.
Private Sub ClearTextFilesFromDir(Path As String, Optional doUCS2 As Boolean = True, Optional doUTF8 As Boolean = True)
    If doUCS2 Then
        On Error GoTo ClearTextFilesFromDir_noop
        If Dir(Path & "\*.data") <> "" Then
            Kill Path & "\*.data"
        End If
ClearTextFilesFromDir_noop:
    End If
    
    If doUTF8 Then
        On Error GoTo ClearTextFilesFromDir_noop2
        If Dir(Path & "\*.txt") <> "" Then
            Kill Path & "\*.txt"
        End If
ClearTextFilesFromDir_noop2:
    End If
    
    On Error GoTo 0
End Sub

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Private Sub SanitizeTextFiles(Path As String)
    Dim Fso, Infile, OutFile, FileName As String, txt As String
    
    Dim ForReading As Long
    
    ForReading = 1
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    FileName = Dir(Path & "\*.txt")
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, Len(FileName) - 4)
        
        Set Infile = Fso.OpenTextFile(Path & "\" & obj_name & ".txt", ForReading)
        Set OutFile = Fso.CreateTextFile(Path & "\" & obj_name & ".sanitize", True)
        Do Until Infile.AtEndOfStream
            txt = Infile.ReadLine
            If Left(txt, 10) = "Checksum =" Then
                ' Skip lines starting with Checksum
            ElseIf InStr(txt, "NoSaveCTIWhenDisabled =1") Then
                ' Skip lines containning NoSaveCTIWhenDisabled
            ElseIf InStr(txt, "PrtDevNames = Begin") > 0 Or _
                InStr(txt, "PrtDevNamesW = Begin") > 0 Or _
                InStr(txt, "PrtDevModeW = Begin") > 0 Or _
                InStr(txt, "PrtDevMode = Begin") > 0 Then
    
                ' skip this block of code
                Do Until Infile.AtEndOfStream
                    txt = Infile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        Infile.Close
        
        FileName = Dir()
    Loop
    
    FileName = Dir(Path & "\*.txt")
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, Len(FileName) - 4)
        Kill Path & "\" & obj_name & ".txt"
        Name Path & "\" & obj_name & ".sanitize" As Path & "\" & obj_name & ".txt"
        FileName = Dir()
    Loop
End Sub

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
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
    Dim tblName As Variant
    
    Set db = CurrentDb
    
    source_path = CurrentProject.Path
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
    
    obj_path = source_path & "\tables"
    MkDirIfNotexist obj_path
    ClearTextFilesFromDir obj_path
    Debug.Print "Exporting tables..."
    For Each tblName In Split(INCLUDE_TABLES, ",")
        ExportTable CStr(tblName), obj_path
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
    
    ShellWait CurrentProject.Path & "\scripts\ucs2-to-utf8.bat", vbNormalFocus
    
    Debug.Print "Removing Checksum and NoSaveCTIWhenDisabled lines"
    For Each obj_type In Split("forms,reports,macros", ",")
        SanitizeTextFiles source_path & "\" & obj_type
    Next
    
    Debug.Print "Done."
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
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
    Dim FileName As String
    Dim obj_name As String
    
    ShellWait CurrentProject.Path & "\scripts\utf8-to-ucs2.bat", vbNormalFocus
    
    Set db = CurrentDb
    
    source_path = CurrentProject.Path
    If Right(source_path, 1) <> "\" Then source_path = source_path & "\"
    source_path = source_path & "source"
    MkDirIfNotexist source_path
    
    Debug.Print
    
    obj_path = source_path & "\queries"
    Debug.Print "Importing queries..."
    FileName = Dir(obj_path & "\*.data")
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, Len(FileName) - 5)
        Application.LoadFromText acQuery, obj_name, obj_path & "\" & FileName
        FileName = Dir()
    Loop
    ClearTextFilesFromDir obj_path, True, False
    
    '' read in table values
    obj_path = source_path & "\tables"
    Debug.Print "Importing tables..."
    FileName = Dir(obj_path & "\*.data")
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, Len(FileName) - 5)
        ImportTable CStr(obj_name), obj_path
        FileName = Dir()
    Loop
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
        FileName = Dir(obj_path & "\*.data")
        Do Until Len(FileName) = 0
            obj_name = Mid(FileName, 1, Len(FileName) - 5)
            If obj_name <> "AppCodeImportExport" Then
                Application.LoadFromText obj_type_num, obj_name, obj_path & "\" & FileName
            End If
            FileName = Dir()
        Loop
        ClearTextFilesFromDir obj_path, True, False
    Next
    
    Debug.Print "Done."
End Sub

' Export the lookup table `tblName` to `source\tables`.
Private Sub ExportTable(tblName As String, obj_path As String)
    Dim Fso, OutFile, rs As Recordset, fieldObj As Field, C As Long, Value As Variant
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    Set OutFile = Fso.CreateTextFile(obj_path & "\" & tblName & ".data", True, True)
    
    Set rs = CurrentDb.OpenRecordset("export_" & tblName)
    C = 0
    For Each fieldObj In rs.Fields
        If C <> 0 Then OutFile.Write vbTab
        C = C + 1
        OutFile.Write fieldObj.Name
    Next
    OutFile.Write vbCrLf
    
    rs.MoveFirst
    Do Until rs.EOF
        C = 0
        For Each fieldObj In rs.Fields
            If C <> 0 Then OutFile.Write vbTab
            C = C + 1
            Value = rs(fieldObj.Name)
            If IsNull(Value) Then
                Value = ""
            Else
                Value = Replace(Value, "\", "\\")
                Value = Replace(Value, vbCrLf, "\n")
                Value = Replace(Value, vbCr, "\n")
                Value = Replace(Value, vbLf, "\n")
                Value = Replace(Value, vbTab, "\t")
            End If
            OutFile.Write CStr(Nz(rs(fieldObj.Name), ""))
        Next
        OutFile.Write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close
End Sub

' Import the lookup table `tblName` from `source\tables`.
Private Sub ImportTable(tblName As String, obj_path As String)
    Dim db As Database, Fso, Infile, rs As Recordset, fieldObj As Field, C As Long
    Dim buf As String, Values() As String, Value As Variant, rsWrite As Recordset
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set Infile = Fso.OpenTextFile(obj_path & "\" & tblName & ".data", ForReading, False, TristateTrue)
    Set db = CurrentDb
    
    db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = db.OpenRecordset("export_" & tblName)
    Set rsWrite = db.OpenRecordset(tblName)
    buf = Infile.ReadLine()
    Do Until Infile.AtEndOfStream
        buf = Infile.ReadLine()
        If Len(Trim(buf)) > 0 Then
            Values = Split(buf, vbTab)
            C = 0
            rsWrite.AddNew
            For Each fieldObj In rs.Fields
                Value = Values(C)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rsWrite(fieldObj.Name) = Value
                C = C + 1
            Next
            rsWrite.Update
        End If
    Loop
    
    rsWrite.Close
    rs.Close
    Infile.Close
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
    Dim Ret As Long
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    ' Start the shelled application:
    Ret& = CreateProcessA(0&, Pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
End Sub
'***************** Code End ****************