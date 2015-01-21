Attribute VB_Name = "AppCodeImportExport"
' Access Module `AppCodeImportExport`
' -----------------------------------
' https://github.com/timabell/msaccess-vcs-integration

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
' in "$database-folder/source/", run
' `ExportAllSource()`.
'
' Table contents that shall be saved must be listed in the INCLUDE_TABLES variable
'
' To load and/or overwrite  all database objects from source files in
' "$database-folder/source/", run `ImportProject()`.
'
' See project home page (URL above) for more information.
'
'
Option Compare Database
Option Explicit

' --------------------------------
' Configuration
' --------------------------------

' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
'
' Provide a comma separated list of table names, or an empty string
' ("") if no tables are to be exported with the source code.

Private Const INCLUDE_TABLES = ""

' Do more aggressive removal of superfluous blobs from exported MS
' Access source code?

Private Const AggressiveSanitize = True
Private Const StripPublishOption = True
Private Const ArchiveMyself = True
Private Const DebugOutput = False
'
' --------------------------------
' Structures
' --------------------------------

' Structure to track buffered reading or writing of binary files
Private Type BinFile
    file_num As Integer
    file_len As Long
    file_pos As Long
    buffer As String
    buffer_len As Integer
    buffer_pos As Integer
    at_eof As Boolean
    mode As String
End Type
' Structure to keep track of "on Update" and "on Delete" clauses
' Access does not in all cases execute such queries
Private Type structEnforce
    foreignTable As String
    foreignFields() As String
    Table As String
    refFields() As String
    isUpdate As Boolean
End Type
' --------------------------------
' Constants
' --------------------------------

' Constants for Scripting.FileSystemObject API
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' --------------------------------
' Module variables
' --------------------------------
'
' Does the current database file write UCS2-little-endian when exporting
' Queries, Forms, Reports, Macros
Private UsingUcs2 As Boolean
'
' keeping "on Update" relations to be complemented after table creation
Private K() As structEnforce


' --------------------------------
' External Library Functions
' --------------------------------
Private Declare PtrSafe _
    Function getTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                               ByVal lpBuffer As String) As Long
Private Declare PtrSafe _
    Function getTempFileName Lib "kernel32" _
         Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                   ByVal lpPrefixString As String, _
                                   ByVal wUnique As Long, _
                                   ByVal lpTempFileName As String) As Long
' --------------------------------
' Basic functions missing from VB 6: buffered file read/write, string builder
' --------------------------------

' Open a binary file for reading (mode = 'r') or writing (mode = 'w').
Private Function BinOpen(file_path As String, mode As String) As BinFile
    Dim f As BinFile

    f.file_num = FreeFile
    f.mode = LCase(mode)
    If f.mode = "r" Then
        Open file_path For Binary Access Read As f.file_num
        f.file_len = LOF(f.file_num)
        f.file_pos = 0
        If f.file_len > &H4000 Then
            f.buffer = String(&H4000, " ")
            f.buffer_len = &H4000
        Else
            f.buffer = String(f.file_len, " ")
            f.buffer_len = f.file_len
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    Else
        DelIfExist file_path
        Open file_path For Binary Access Write As f.file_num
        f.file_len = 0
        f.file_pos = 0
        f.buffer = String(&H4000, " ")
        f.buffer_len = 0
        f.buffer_pos = 0
    End If

    BinOpen = f
End Function

' Buffered read one byte at a time from a binary file.
Private Function BinRead(ByRef f As BinFile) As Integer
    If f.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid(f.buffer, f.buffer_pos + 1, 1))

    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= f.buffer_len Then
        f.file_pos = f.file_pos + &H4000
        If f.file_pos >= f.file_len Then
            f.at_eof = True
            Exit Function
        End If
        If f.file_len - f.file_pos > &H4000 Then
            f.buffer_len = &H4000
        Else
            f.buffer_len = f.file_len - f.file_pos
            f.buffer = String(f.buffer_len, " ")
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    End If
End Function

' Buffered write one byte at a time from a binary file.
Private Sub BinWrite(ByRef f As BinFile, b As Integer)
    Mid(f.buffer, f.buffer_pos + 1, 1) = Chr(b)
    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= &H4000 Then
        Put f.file_num, , f.buffer
        f.buffer_pos = 0
    End If
End Sub

' Close binary file.
Private Sub BinClose(ByRef f As BinFile)
    If f.mode = "w" And f.buffer_pos > 0 Then
        f.buffer = Left(f.buffer, f.buffer_pos)
        Put f.file_num, , f.buffer
    End If
    Close f.file_num
End Sub

' String builder: Init
Private Function Sb_Init() As String()
    Dim x(-1 To -1) As String
    Sb_Init = x
End Function

' String builder: Clear
Private Sub Sb_Clear(ByRef sb() As String)
    ReDim Sb_Init(-1 To -1)
End Sub

' String builder: Append
Private Sub Sb_Append(ByRef sb() As String, Value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = Value
End Sub

' String builder: Get value
Private Function Sb_Get(ByRef sb() As String) As String
    Sb_Get = Join(sb, "")
End Function

' --------------------------------
' Beginning of main functions of this module
' --------------------------------

' Close all open forms.
Private Function CloseFormsReports()
    On Error GoTo errorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).Name
        DoEvents
    Loop
    Exit Function

errorHandler:
    Debug.Print "AppCodeImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & Err.Description
End Function

' Pad a string on the right to make it `count` characters long.
Public Function PadRight(Value As String, Count As Integer)
    PadRight = Value
    If Len(Value) < Count Then
        PadRight = PadRight & Space(Count - Len(Value))
    End If
End Function

' Path of the current database file.
Private Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function
'
' Generate Random / Unique tempprary file name.
Private Function TempFile(Optional sPrefix As String = "VBA") As String
Dim sTmpPath As String * 512
Dim sTmpName As String * 576
Dim nRet As Long
Dim sFileName As String
    
    nRet = getTempPath(512, sTmpPath)
    nRet = getTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then sFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    TempFile = sFileName
End Function

' Export a database object with optional UCS2-to-UTF-8 conversion.
Private Sub ExportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    MkDirIfNotExist Left(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub ImportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = TempFile()
        ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

' Binary convert a UCS2-little-endian encoded file to UTF-8.
Private Sub ConvertUcs2Utf8(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_low As Integer, in_high As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_low = BinRead(f_in)
        in_high = BinRead(f_in)
        If in_high = 0 And in_low < &H80 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_low
        ElseIf in_high < &H8 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            BinWrite f_out, &HC0 + ((in_high And &H7) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            BinWrite f_out, &HE0 + ((in_high And &HF0) / &H10)
            BinWrite f_out, &H80 + ((in_high And &HF) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Binary convert a UTF-8 encoded file to UCS2-little-endian.
Private Sub ConvertUtf8Ucs2(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_1 As Integer, in_2 As Integer, in_3 As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_1 = BinRead(f_in)
        If (in_1 And &H80) = 0 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_1
            BinWrite f_out, 0
        ElseIf (in_1 And &HE0) = &HC0 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            BinWrite f_out, ((in_1 And &H3) * &H40) + (in_2 And &H3F)
            BinWrite f_out, (in_1 And &H1C) / &H4
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            in_3 = BinRead(f_in)
            BinWrite f_out, ((in_2 And &H3) * &H40) + (in_3 And &H3F)
            BinWrite f_out, ((in_1 And &HF) * &H10) + ((in_2 And &H3C) / &H4)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Private Sub InitUsingUcs2()
    Dim obj_name As String, i As Integer, obj_type As Variant, fn As Integer, bytes As String
    Dim obj_type_split() As String, obj_type_name As String, obj_type_num As Integer
    Dim Db As Object ' DAO.Database

    If CurrentDb.QueryDefs.Count > 0 Then
        obj_type_num = acQuery
        obj_name = CurrentDb.QueryDefs(0).Name
    Else
        For Each obj_type In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
        )
            DoEvents
            obj_type_split = Split(obj_type, "|")
            obj_type_name = obj_type_split(0)
            obj_type_num = Val(obj_type_split(1))
            If CurrentDb.Containers(obj_type_name).Documents.Count > 0 Then
                obj_name = CurrentDb.Containers(obj_type_name).Documents(0).Name
                Exit For
            End If
        Next
    End If

    If obj_name = "" Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2 = True
        Exit Sub
    End If

    Dim tempFileName As String: tempFileName = TempFile()
    Application.SaveAsText obj_type_num, obj_name, tempFileName
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    If Asc(Mid(bytes, 1, 1)) = &HFF And Asc(Mid(bytes, 2, 1)) = &HFE Then
        UsingUcs2 = True
    Else
        UsingUcs2 = False
    End If
    Close fn
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFile (tempFileName)
End Sub

' Create folder `Path`. Silently do nothing if it already exists.
Private Sub MkDirIfNotExist(Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

' Delete a file if it exists.
Private Sub DelIfExist(Path As String)
    On Error GoTo DelIfNotExist_Noop
    Kill Path
DelIfNotExist_Noop:
    On Error GoTo 0
End Sub

' Erase all *.`ext` files in `Path`.
Private Sub ClearTextFilesFromDir(Path As String, Ext As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Path) Then Exit Sub

    On Error GoTo ClearTextFilesFromDir_noop
    If Dir(Path & "*." & Ext) <> "" Then
        FSO.DeleteFile Path & "*." & Ext
    End If
ClearTextFilesFromDir_noop:

    On Error GoTo 0
End Sub

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Private Sub SanitizeTextFiles(Path As String, Ext As String)


    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      srchPattern = "(?:" & srchPattern
      srchPattern = srchPattern & "|GUID|NameMap|dbLongBinary ""DOL"""
      srchPattern = srchPattern & ")"
    End If
    '  Ensure that this is the begining of a block.
    srchPattern = srchPattern & " = Begin"
'Debug.Print srchPattern
    rxBlock.Pattern = srchPattern
    '
    '  Setup Line Matching Regex.
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
'Debug.Print srchPattern
    rxLine.Pattern = srchPattern
    Dim fileName As String
    fileName = Dir(Path & "*." & Ext)
    
    Do Until Len(fileName) = 0
        DoEvents
        Dim obj_name As String
        obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)

        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(Path & obj_name & "." & Ext, ForReading)
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(Path & obj_name & ".sanitize", True)
    
        Dim getLine As Boolean: getLine = True
        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String
            '
            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If
            '
            ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"
                '
                ' Get indentation level.
                Dim matches As Object
                Set matches = rxIndent.Execute(txt)
                '
                ' Setup pattern to match current indent
                Select Case matches.Count
                    Case 0
                        rxIndent.Pattern = "^" & vbNullString
                    Case Else
                        rxIndent.Pattern = "^" & matches(0).SubMatches(0)
                End Select
                rxIndent.Pattern = rxIndent.Pattern + "\S"
                '
                ' Skip lines with deeper indentation
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If rxIndent.Test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False
            '
            ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FSO.DeleteFile (Path & fileName)

        Dim thisFile As Object
        Set thisFile = FSO.GetFile(Path & obj_name & ".sanitize")
        thisFile.Move (Path & fileName)
        fileName = Dir()
    Loop


End Sub

' Import References from a CSV, true=SUCCESS
Private Function ImportReferences(obj_path As String) As Boolean
    Dim FSO, InFile
    Dim line As String
    Dim item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim fileName As String
    fileName = Dir(obj_path & "references.csv")
    If Len(fileName) = 0 Then
        ImportReferences = False
        Exit Function
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(obj_path & fileName, ForReading)
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        item = Split(line, ",")
        GUID = Trim(item(0))
        Major = CLng(item(1))
        Minor = CLng(item(2))
        Application.References.AddFromGuid GUID, Major, Minor
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    ImportReferences = True
    Exit Function
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
    
End Function
' Export References to a CSV
Private Sub ExportReferences(obj_path As String)
    Dim FSO, OutFile
    Dim line As String
    Dim ref As Reference
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(obj_path & "references.csv", True)
    For Each ref In Application.References
        line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
        OutFile.WriteLine line
    Next
    OutFile.Close
End Sub
' Save a Table Definition as SQL statement
Public Sub ExportTableDef(Db As Database, td As TableDef, tableName As String, fileName As String)
    Dim sql As String
    Dim fieldAttributeSql As String
    Dim idx As Index
    Dim fi As Field
    Dim i As Integer
    Dim f As Field
    Dim rel As Relation
    Dim FSO, OutFile
    Dim ff As Object
    'Debug.Print tableName
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(fileName, True)
    sql = "CREATE TABLE " & strName(tableName) & " (" & vbCrLf
    For Each fi In td.Fields
        sql = sql & "  " & strName(fi.Name) & " "
        If (fi.Attributes And dbAutoIncrField) Then
            sql = sql & "AUTOINCREMENT"
        Else
            sql = sql & strType(fi.Type) & " "
        End If
        Select Case fi.Type
            Case dbText, dbVarBinary
                sql = sql & "(" & fi.Size & ")"
            Case Else
        End Select
        For Each idx In td.Indexes
            fieldAttributeSql = ""
            If idx.Fields.Count = 1 And idx.Fields(0).Name = fi.Name Then
                If idx.Primary Then fieldAttributeSql = fieldAttributeSql & " PRIMARY KEY "
                If idx.Unique Then fieldAttributeSql = fieldAttributeSql & " UNIQUE "
                If idx.Required Then fieldAttributeSql = fieldAttributeSql & " NOT NULL "
                If idx.Foreign Then
                    Set ff = idx.Fields
                    fieldAttributeSql = fieldAttributeSql & formatReferences(Db, ff, tableName)
                End If
                If Len(fieldAttributeSql) > 0 Then fieldAttributeSql = " CONSTRAINT " & strName(idx.Name) & fieldAttributeSql
            End If
        Next
        sql = sql + fieldAttributeSql
        sql = sql + "," & vbCrLf
    Next
    sql = Left(sql, Len(sql) - 3) ' strip off last comma and crlf
    
    Dim constraintSql As String
    For Each idx In td.Indexes
        If idx.Fields.Count > 1 Then
            If Len(constraintSql) = 0 Then constraintSql = constraintSql & " CONSTRAINT "
            If idx.Primary Then constraintSql = constraintSql & formatConstraint("PRIMARY KEY", idx)
            If idx.Foreign Then constraintSql = formatConstraint("FOREIGN KEY", idx)
            If Len(constraintSql) > 0 Then
                sql = sql & "," & vbCrLf & "  " & constraintSql
                sql = sql & formatReferences(Db, idx.Fields, tableName)
            End If
        End If
    Next
    sql = sql & vbCrLf & ")"

    'Debug.Print sql
    OutFile.WriteLine sql
    
    OutFile.Close
End Sub
Private Function formatReferences(Db As Database, ff As Object, tableName As String)
    Dim rel As Relation
    Dim sql As String
    Dim f As Field
    For Each rel In Db.Relations
        If (rel.foreignTable = tableName) Then
         If FieldsIdentical(ff, rel.Fields) Then
          sql = " REFERENCES "
          sql = sql & rel.Table & " ("
          For Each f In rel.Fields
            sql = sql & strName(f.Name) & ","
          Next
          sql = Left(sql, Len(sql) - 1) & ")"
          If rel.Attributes And dbRelationUpdateCascade Then
            sql = sql + " ON UPDATE CASCADE "
          End If
          If rel.Attributes And dbRelationDeleteCascade Then
            sql = sql + " ON DELETE CASCADE "
          End If
          Exit For
         End If
        End If
    Next
    formatReferences = sql
End Function

Private Function formatConstraint(keyw As String, idx As Index) As String
    Dim sql As String
    Dim fi As Field
    
    sql = strName(idx.Name) & " " & keyw & " ("
    For Each fi In idx.Fields
        sql = sql & strName(fi.Name) & ", "
    Next
    sql = Left(sql, Len(sql) - 2) & ")" 'strip off last comma and close brackets
    
    'return value
    formatConstraint = sql
End Function

Private Function strName(s As String) As String
    strName = "[" & s & "]"
End Function

Private Function strType(i As Integer) As String
    Select Case i
    Case dbLongBinary
        strType = "LONGBINARY"
    Case dbBinary
        strType = "BINARY"
    'Case dbBit missing enum
    '    strType = "BIT"
    Case dbAutoIncrField
        strType = "COUNTER"
    Case dbCurrency
        strType = "CURRENCY"
    Case dbDate, dbTime
        strType = "DATETIME"
    Case dbGUID
        strType = "GUID"
    Case dbMemo
        strType = "LONGTEXT"
    Case dbDouble
        strType = "DOUBLE"
    Case dbSingle
        strType = "SINGLE"
    Case dbByte
        strType = "UNSIGNED BYTE"
    Case dbInteger
        strType = "SHORT"
    Case dbLong
        strType = "LONG"
    Case dbNumeric
        strType = "NUMERIC"
    Case dbText
        strType = "VARCHAR"
    Case Else
        strType = "VARCHAR"
    End Select
End Function
Private Function FieldsIdentical(ff As Object, gg As Object) As Boolean
    Dim f As Field
    If ff.Count <> gg.Count Then
        FieldsIdentical = False
        Exit Function
    End If
    For Each f In ff
        If Not FieldInFields(f, gg) Then
        FieldsIdentical = False
        Exit Function
        End If
    Next
    FieldsIdentical = True
        
    
End Function

Private Function FieldInFields(fi As Field, ff As Fields) As Boolean
    Dim f As Field
    For Each f In ff
        If f.Name = fi.Name Then
            FieldInFields = True
            Exit Function
        End If
    Next
    FieldInFields = False
End Function

' Determine if a table or exists.
' based on sample code of support.microsoftcom
' ARGUMENTS:
'    TName: The name of a table or query.
'
' RETURNS: True (it exists) or False (it does not exist).
Function TableExists(TName As String) As Boolean
        Dim Db As Database, Found As Boolean, Test As String
        Const NAME_NOT_IN_COLLECTION = 3265

         ' Assume the table or query does not exist.
        Found = False
        Set Db = CurrentDb()

         ' Trap for any errors.
        On Error Resume Next
         
         ' See if the name is in the Tables collection.
        Test = Db.TableDefs(TName).Name
        If Err <> NAME_NOT_IN_COLLECTION Then Found = True

        ' Reset the error variable.
        Err = 0

        TableExists = Found

End Function

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim ucs2 As Boolean
    Dim tblName As Variant

    Set Db = CurrentDb

    CloseFormsReports
    InitUsingUcs2

    source_path = ProjectPath() & "source\"
    MkDirIfNotExist source_path

    Debug.Print

    obj_path = source_path & "queries\"
    ClearTextFilesFromDir obj_path, "bas"
    Debug.Print PadRight("Exporting queries...", 24);
    obj_count = 0
    For Each qry In Db.QueryDefs
        DoEvents
                If Left(qry.Name, 1) <> "~" Then
            ExportObject acQuery, qry.Name, obj_path & qry.Name & ".bas", UsingUcs2
            obj_count = obj_count + 1
        End If
    Next
    SanitizeTextFiles obj_path, "bas"
    Debug.Print "[" & obj_count & "]"

    obj_path = source_path & "tables\"
    ClearTextFilesFromDir obj_path, "txt"
    If (Len(Replace(INCLUDE_TABLES, " ", "")) > 0) Then
        Debug.Print PadRight("Exporting tables...", 24);
        obj_count = 0
        For Each tblName In Split(INCLUDE_TABLES, ",")
            DoEvents
            ExportTable CStr(tblName), obj_path
            If Len(Dir(obj_path & tblName & ".txt")) > 0 Then
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"
    End If

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
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        ClearTextFilesFromDir obj_path, "bas"
        Debug.Print PadRight("Exporting " & obj_type_label & "...", 24);
        For Each doc In Db.Containers(obj_type_name).Documents
            DoEvents
            If (Left(doc.Name, 1) <> "~") And _
               (doc.Name <> "AppCodeImportExport" Or ArchiveMyself) Then
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = UsingUcs2
                End If
                ExportObject obj_type_num, doc.Name, obj_path & doc.Name & ".bas", ucs2
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"

        If obj_type_label <> "modules" Then
            SanitizeTextFiles obj_path, "bas"
        Else
            ' Make sure all modules find their needed references
            If obj_count > 0 Then ExportReferences obj_path
        End If
    Next

    
    Dim td As TableDef
    Dim tds As TableDefs
    Set tds = Db.TableDefs

    obj_type_label = "tbldef"
    obj_type_name = "Table_Def"
    obj_type_num = acTable
    obj_path = source_path & obj_type_label & "\"
    obj_count = 0
    MkDirIfNotExist Left(obj_path, InStrRev(obj_path, "\"))
    ClearTextFilesFromDir obj_path, "sql"
    Debug.Print PadRight("Exporting " & obj_type_label & "...", 24);
    
    For Each td In tds
        ' This is not a system table
        ' this is not a temporary table
        ' this is not an external table
        If Left$(td.Name, 4) <> "MSys" And _
        Left(td.Name, 1) <> "~" _
        And Len(td.Connect) = 0 _
        Then
            'Debug.Print
            ExportTableDef Db, td, td.Name, obj_path & td.Name & ".sql"
            obj_count = obj_count + 1
        End If
    Next
    Debug.Print "[" & obj_count & "]"
    
    
    Debug.Print "Done."
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource()
    Dim Db As Object ' DAO.Database
    Dim FSO As Object
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    Set Db = CurrentDb
    Set FSO = CreateObject("Scripting.FileSystemObject")

    CloseFormsReports
    InitUsingUcs2

    source_path = ProjectPath() & "source\"
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print

    obj_path = source_path & "queries\"
    fileName = Dir(obj_path & "*.bas")
    If Len(fileName) > 0 Then
        Debug.Print PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            ImportObject acQuery, obj_name, obj_path & fileName, UsingUcs2
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If

    ' restore table definitions
    obj_path = source_path & "tbldef\"
    fileName = Dir(obj_path & "*.sql")
    If Len(fileName) > 0 Then
        Debug.Print PadRight("Importing tabledefs...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            If DebugOutput Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            ImportTableDef CStr(obj_name), obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    ' NOW we may load data
    obj_path = source_path & "tables\"
    fileName = Dir(obj_path & "*.txt")
    If Len(fileName) > 0 Then
        Debug.Print PadRight("Importing tables...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            ImportTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If

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
        obj_path = source_path & obj_type_label & "\"
        
        If obj_type_label = "modules" Then
            If Not ImportReferences(obj_path) Then
                Debug.Print
                Debug.Print "Info: no references file in " & obj_path
            End If
        End If
    
        
        fileName = Dir(obj_path & "*.bas")
        If Len(fileName) > 0 Then
            Debug.Print PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = UsingUcs2
                End If
                If obj_name <> "AppCodeImportExport" Then
                    ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
                    obj_count = obj_count + 1
                Else
                    If ArchiveMyself Then
                            MsgBox "Module AppCodeImportExport could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
                    End If
                End If
                fileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    DoEvents
    Debug.Print "Done."
End Sub

' Build SQL to export `tbl_name` sorted by each field from first to last
Private Function TableExportSql(tbl_name As String)
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, Count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)
    
    sb = Sb_Init()
    Sb_Append sb, "SELECT "
    Count = 0
    For Each fieldObj In rs.Fields
        If Count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next
    Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    Count = 0
    For Each fieldObj In rs.Fields
        DoEvents
        If Count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next

    TableExportSql = Sb_Get(sb)

End Function

' Export the lookup table `tblName` to `source\tables`.
Private Sub ExportTable(tbl_name As String, obj_path As String)
    Dim FSO, OutFile
    Dim rs As Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim c As Long, Value As Variant
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset(TableExportSql(tbl_name))
    If rs.RecordCount = 0 Then
        Debug.Print "Error: Table " & tbl_name & "  empty"
        rs.Close
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    MkDirIfNotExist obj_path
    Dim tempFileName As String: tempFileName = TempFile()

    Set OutFile = FSO.CreateTextFile(tempFileName, True, True)

    c = 0
    For Each fieldObj In rs.Fields
        If c <> 0 Then OutFile.Write vbTab
        c = c + 1
        OutFile.Write fieldObj.Name
    Next
    OutFile.Write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        c = 0
        For Each fieldObj In rs.Fields
            DoEvents
            If c <> 0 Then OutFile.Write vbTab
            c = c + 1
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
            OutFile.Write Value
        Next
        OutFile.Write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close

    ConvertUcs2Utf8 tempFileName, obj_path & tbl_name & ".txt"
    FSO.DeleteFile tempFileName
End Sub
' Kill Table if Exists
Private Sub KillTable(tblName As String, Db As Object)
    If TableExists(tblName) Then
        Db.Execute "DROP TABLE [" & tblName & "]"
    End If
End Sub
' Import Table Definition
Private Sub ImportTableDef(tblName As String, FilePath As String)
    Dim Db As Object ' DAO.Database
    Dim FSO, InFile As Object
    Dim buf As String
    Dim p As Integer
    Dim p1 As Integer
    Dim strFields() As String
    Dim strRef As String
    Dim strMsg As String
    Dim strForeignKeys() As String
    Dim s
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tempFileName As String: tempFileName = TempFile()

    n = -1
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ConvertUtf8Ucs2 FilePath, tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set Db = CurrentDb
    KillTable tblName, Db
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = buf & InFile.ReadLine()
    Loop
    
    ' The following block is needed because "on update" actions may cause problems
    For Each s In Split("UPDATE|DELETE", "|")
    p = InStr(buf, "ON " & s & " CASCADE")
    While p > 0
        n = n + 1
        ReDim Preserve K(n)
        K(n).Table = tblName
        K(n).isUpdate = (s = "UPDATE")
        
        buf = Left(buf, p - 1) & Mid(buf, p + 18)
        p = InStrRev(buf, "REFERENCES", p)
        p1 = InStr(p, buf, "(")
        K(n).foreignFields = Split(SubString(p1, buf, "(", ")"), ",")
        K(n).foreignTable = Trim(Mid(buf, p + 10, p1 - p - 10))
        p = InStrRev(buf, "CONSTRAINT", p1)
        p1 = InStrRev(buf, "FOREIGN KEY", p1)
        If (p1 > 0) And (p > 0) And (p1 > p) Then
        ' multifield index
            K(n).refFields = Split(SubString(p1, buf, "(", ")"), ",")
        ElseIf p1 = 0 Then
        ' single field
        End If
        p = InStr(p, "ON " & s & " CASCADE", buf)
    Wend
    Next
    On Error Resume Next
    For i = 0 To n
        strMsg = K(i).Table & " to " & K(i).foreignTable
        strMsg = strMsg & "(  "
        For j = 0 To UBound(K(i).refFields)
            strMsg = strMsg & K(i).refFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") to ("
        For j = 0 To UBound(K(i).foreignFields)
            strMsg = strMsg & K(i).foreignFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") Check "
        If K(i).isUpdate Then
            strMsg = strMsg & " on update cascade " & vbCrLf
        Else
            strMsg = strMsg & " on delete cascade " & vbCrLf
        End If
    Next
    On Error GoTo 0
    Db.Execute buf
    InFile.Close
    If Len(strMsg) > 0 Then MsgBox strMsg, vbOKOnly, "Correct manually"
End Sub
' returns substring between e.g. "(" and ")", internal brackets ar skippped
Public Function SubString(p As Integer, s As String, startsWith As String, endsWith As String)
    Dim start As Integer
    Dim last As Integer
    Dim cursor As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    Dim level As Integer
    start = InStr(p, s, startsWith)
    level = 1
    p1 = InStr(start + 1, s, startsWith)
    p2 = InStr(start + 1, s, endsWith)
    While level > 0
        If p1 > p2 And p2 > 0 Then
            cursor = p2
            level = level - 1
        ElseIf p2 > p1 And p1 > 0 Then
            cursor = p1
            level = level + 1
        ElseIf p2 > 0 And p1 = 0 Then
            cursor = p2
            level = level - 1
        ElseIf p1 > 0 And p1 = 0 Then
            cursor = p1
            level = level + 1
        ElseIf p1 = 0 And p2 = 0 Then
            SubString = ""
            Exit Function
        End If
        p1 = InStr(cursor + 1, s, startsWith)
        p2 = InStr(cursor + 1, s, endsWith)
    Wend
    SubString = Mid(s, start + 1, cursor - start - 1)
End Function
' Import the lookup table `tblName` from `source\tables`.
Private Sub ImportTable(tblName As String, obj_path As String)
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO, InFile As Object
    Dim c As Long, buf As String, Values() As String, Value As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFileName As String: tempFileName = TempFile()
    ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set Db = CurrentDb

    Db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = Db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim(buf)) > 0 Then
            Values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                Value = Values(c)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.Name) = Value
                c = c + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
End Sub

' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject()
On Error GoTo errorHandler

    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr(149) & " Forms" & vbCrLf & _
              Chr(149) & " Macros" & vbCrLf & _
              Chr(149) & " Modules" & vbCrLf & _
              Chr(149) & " Queries" & vbCrLf & _
              Chr(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Sub
    End If

    Dim Db As DAO.Database
    Set Db = CurrentDb
    CloseFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print

    Dim dbObject As Object
    For Each dbObject In Db.QueryDefs
        DoEvents
        If Left(dbObject.Name, 1) <> "~" Then
'            Debug.Print dbObject.Name
            Db.QueryDefs.Delete dbObject.Name
        End If
    Next

    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME = 0
    Const OTID = 1

    For Each objType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(objType, "|")
        DoEvents
        For Each doc In Db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left(doc.Name, 1) <> "~") And _
               (doc.Name <> "AppCodeImportExport") Then
'                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.Name
            End If
        Next
    Next
    
    Debug.Print "================="
    Debug.Print "Importing Project"
    ImportAllSource
    GoTo exitHandler

errorHandler:
  Debug.Print "AppCodeImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
               Err.Description

exitHandler:
End Sub
' Expose for use as function, can be called by query
Public Function make()
    ImportProject
End Function
