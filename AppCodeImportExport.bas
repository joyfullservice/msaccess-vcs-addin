Attribute VB_Name = "AppCodeImportExport"
Option Compare Database
Option Explicit

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
' 2013-Mar-30  Brett Maton
'   Replaced Wait for Shell external library and kernel references with
'   Windows Shell Host command.  WSH works on 32 and 64bit versions of
'   MSOffice, the external references did not, without adding a bunch of
'   conditional compilation and basically duplicating the code (once for
'   each platform).
'   Changed path '\' references, now added to the end of paths instead of
'   the beginning.  Avoids ending up in the "root" filesystem if the path is
'   missing for any reason.
'   source\xxx directories are only created if something is going to be
'   exported.
'   Added queries in call to SanitizeFiles and skipped yet another
'   apparently useless bit of MS bloat "DOL"
'   Removed requirement for all directories to be present in Encoding scripts.
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
' Provide a comma separated list of table names, or an empty string
' ("") if no tables are to be exported with the source code.
' --------------------------------

Private Const INCLUDE_TABLES = ""

' --------------------------------
' Constants
' --------------------------------

Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2
' --------------------------------
' Beginning of main functions of this module
' --------------------------------

' Create folder `Path`. Silently do nothing if it already exists.
Private Sub MkDirIfNotexist(Path As String)
  On Error Resume Next
  MkDir Path
End Sub

' Erase all *.data and *.txt files in `Path`.
Private Sub ClearTextFilesFromDir(Path As String, Optional doUCS2 As Boolean = True, Optional doUTF8 As Boolean = True)

  If doUCS2 Then
    On Error Resume Next
    If Dir(Path & "*.data") <> "" Then
      Kill Path & "*.data"
    End If
  End If
    
  If doUTF8 Then
    On Error Resume Next
    If Dir(Path & "\*.txt") <> "" Then
      Kill Path & "\*.txt"
    End If
 End If
    
End Sub

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Private Sub SanitizeTextFiles(Path As String)
Dim FSO As Object
Dim Infile As Object
Dim OutFile As Object
Dim fileName As String
Dim strLine As String
Dim objName As String
    
  Set FSO = CreateObject("Scripting.FileSystemObject")
    
  fileName = Dir(Path & "*.txt")
  Do Until Len(fileName) = 0
    objName = Mid(fileName, 1, Len(fileName) - 4)
        
    Set Infile = FSO.OpenTextFile(Path & objName & ".txt", vbReadOnly)
    Set OutFile = FSO.CreateTextFile(Path & objName & ".sanitize", True)

    Do Until Infile.AtEndOfStream
      strLine = Infile.ReadLine
      ' Skip lines starting with Checksum
      If Left(strLine, 10) = "Checksum =" Then
      ' Skip lines containing NoSaveCTIWhenDisabled
      ElseIf InStr(strLine, "NoSaveCTIWhenDisabled =1") Then
      ElseIf InStr(strLine, "PrtDevNames = Begin") > 0 Or _
                InStr(strLine, "PrtDevNamesW = Begin") > 0 Or _
                InStr(strLine, "PrtDevModeW = Begin") > 0 Or _
                InStr(strLine, "PrtDevMode = Begin") > 0 Or _
                InStr(strLine, "dbLongBinary ""DOL"" = Begin") > 0 Then
    
                ' skip this block of code
        Do Until Infile.AtEndOfStream
          strLine = Infile.ReadLine
          '
          '  BM This should be a reg-ex anchored to the end of the line.
          '     What happens (albeit unlikely) if a data line contains 'End' ?
          If InStr(strLine, "End") Then Exit Do
        Loop
      Else
        OutFile.WriteLine strLine
      End If
    Loop
    OutFile.Close
    Infile.Close
    '
    '  Delete the old file
    Kill Path & objName & ".txt"
    '
    '  Rename Sanitized file
    Name Path & objName & ".sanitize" As Path & objName & ".txt"
    
    fileName = Dir()
  Loop

End Sub

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
Dim db As Database
Dim qry As QueryDef
Dim sourcePath As String
Dim doc As Document
Dim objType As Variant
Dim objTypeSplit() As String
Dim objTypeLabel As String
Dim objPath As String
Dim objTypeName As String
Dim objTypeNum As Integer
Dim tblName As Variant
  
  Set db = CurrentDb
    
  sourcePath = CurrentProject.Path
  If Right(sourcePath, 1) <> "\" Then sourcePath = sourcePath & "\"
  sourcePath = sourcePath & "source\"
  MkDirIfNotexist sourcePath
    
  Debug.Print
    
  objPath = sourcePath & "queries\"
  If (db.QueryDefs.Count > 0) Then
    MkDirIfNotexist objPath
    ClearTextFilesFromDir objPath
  
    Debug.Print "Exporting queries..."
    For Each qry In db.QueryDefs
      If Left(qry.Name, 1) <> "~" Then
        Application.SaveAsText acQuery, qry.Name, objPath & qry.Name & ".data"
      End If
    Next
  End If
  '
  '  If we're exporting any tables, then create the directory and
  '  run the table export bit.
  If (Len(INCLUDE_TABLES) > 0) Then
    objPath = sourcePath & "tables\"
    MkDirIfNotexist objPath
    ClearTextFilesFromDir objPath
    Debug.Print "Exporting tables..."

    For Each tblName In Split(INCLUDE_TABLES, ",")
      ExportTable CStr(tblName), objPath
    Next
  End If
  
  For Each objType In Split( _
        "forms|Forms|" & acForm & "," & _
        "reports|Reports|" & acReport & "," & _
        "macros|Scripts|" & acMacro & "," & _
        "modules|Modules|" & acModule _
        , "," _
  )
    objTypeSplit = Split(objType, "|")
    objTypeLabel = objTypeSplit(0)
    objTypeName = objTypeSplit(1)
    objTypeNum = Val(objTypeSplit(2))
    objPath = sourcePath & objTypeLabel & "\"
    '
    '  If we haven't got anything to export, then don't.
    If (docsToExport(db.Containers(objTypeName).Documents) > 0) Then
      MkDirIfNotexist objPath
      ClearTextFilesFromDir objPath
      Debug.Print "Exporting " & objTypeLabel & "..."
      For Each doc In db.Containers(objTypeName).Documents
'        Debug.Print "Application.SaveAsText " & objTypeNum & ", " & _
'                                         """" & doc.Name & """, " & _
'                                         """" & objPath & doc.Name & ".data"
        If Left(doc.Name, 1) <> "~" Then
          Application.SaveAsText objTypeNum, doc.Name, objPath & doc.Name & ".data"
        End If
      Next
    End If
  Next
  '
  '  Convert to UTF8
  Debug.Print "Converting to UTF8"
  ShellWait CurrentProject.Path & "\scripts\ucs2-to-utf8.bat", vbNormalFocus
  
  Debug.Print "Removing unnecessary properties"
  For Each objType In Split("forms,reports,macros,queries", ",")
    SanitizeTextFiles sourcePath & objType & "\"
  Next

  Debug.Print "Done."
End Sub
'
'  Count the documents we want to export.
Private Function docsToExport(docs As Documents) As Integer
Dim doc As Document
  docsToExport = 0
  For Each doc In docs
    If Left(doc.Name, 1) <> "~" Then docsToExport = docsToExport + 1
  Next
End Function

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource()
Dim db As Database
Dim qry As QueryDef
Dim sourcePath As String
Dim objPath As String
Dim doc As Document
Dim objType As Variant
Dim objTypeSplit() As String
Dim objTypePath As String
Dim objTypeName As String
Dim objTypeNum As Integer
Dim objName As String
Dim fileName As String

  ShellWait CurrentProject.Path & "\scripts\utf8-to-ucs2.bat", vbNormalFocus

  Set db = CurrentDb

  sourcePath = CurrentProject.Path
  If Right(sourcePath, 1) <> "\" Then sourcePath = sourcePath & "\"
    sourcePath = sourcePath & "source\"
    '
    '  We're trying to read, why on earth create a missing path?
    '
    'MkDirIfNotexist sourcePath
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not (FSO.FolderExists(sourcePath)) Then
      Debug.Print "Cannot find source to import (" & sourcePath & ")"
      Exit Sub
    End If
    Debug.Print
    '
    '  Don't prefix the path with '\', you'll end up in the root filesystem if
    '  there is a problem.
    objPath = sourcePath & "queries\"
    Debug.Print "Importing Queries..."
    fileName = Dir(objPath & "*.data")
    Do Until Len(fileName) = 0
        objName = Mid(fileName, 1, Len(fileName) - 5)
        Application.LoadFromText acQuery, objName, objPath & fileName
        fileName = Dir()
    Loop
    ClearTextFilesFromDir objPath, True, False
    '
    ' Read in table values
    objPath = sourcePath & "tables\"
    Debug.Print "Importing Tables..."
    fileName = Dir(objPath & "*.data")
    Do Until Len(fileName) = 0
      objName = Mid(fileName, 1, Len(fileName) - 5)
      ImportTable CStr(objName), objPath
      fileName = Dir()
    Loop
    ClearTextFilesFromDir objPath, True, False
    
    For Each objType In Split( _
        "forms|" & acForm & "," & _
        "reports|" & acReport & "," & _
        "macros|" & acMacro & "," & _
        "modules|" & acModule _
        , "," _
    )
      objTypeSplit = Split(objType, "|")
      objTypePath = objTypeSplit(0)
      objTypeNum = Val(objTypeSplit(1))
      objPath = sourcePath & objTypePath & "\"
      
      Debug.Print "Importing " & StrConv(objTypePath, vbProperCase) & "..."

      fileName = Dir(objPath & "*.data")
      Do Until Len(fileName) = 0
        objName = Mid(fileName, 1, Len(fileName) - 5)
        If objName <> "AppCodeImportExport" Then
'          Debug.Print "Application.LoadFromText " & objTypeNum; ", " & _
'                                             """" & objName & """, """ & _
'                                             objPath & fileName & """"
          Application.LoadFromText objTypeNum, objName, objPath & fileName
        End If
        fileName = Dir()
      Loop
      ClearTextFilesFromDir objPath, True, False
    Next
    
    Debug.Print "Done."
End Sub

' Export the lookup table `tblName` to `source\tables`.
Private Sub ExportTable(tblName As String, obj_path As String)
    Dim FSO, OutFile, rs As Recordset, fieldObj As Field, C As Long, Value As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    Set OutFile = FSO.CreateTextFile(obj_path & "\" & tblName & ".data", True, True)
    
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
    Dim db As Database, FSO, Infile, rs As Recordset, fieldObj As Field, C As Long
    Dim buf As String, Values() As String, Value As Variant, rsWrite As Recordset
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set Infile = FSO.OpenTextFile(obj_path & "\" & tblName & ".data", vbReadOnly, False, TristateTrue)
    Set db = CurrentDb
    
    db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = db.OpenRecordset(tblName)
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
            rsWrite.update
        End If
    Loop
    
    rsWrite.Close
    rs.Close
    Infile.Close
End Sub
Public Sub ShellWait(appPath As String, Optional windowStyle As Long)
Dim wSHShell As Object
Dim shellCmd As String
Dim iResult As Integer

  shellCmd = "Cmd /c " & appPath

  Set wSHShell = CreateObject("WScript.Shell")
  iResult = wSHShell.Run(shellCmd, windowStyle, True)
  Set wSHShell = Nothing
End Sub

