Attribute VB_Name = "VCS_Reference"
Option Compare Database

Option Private Module
Option Explicit


' Import References from a CSV, true=SUCCESS
Public Function VCS_ImportReferences(ByVal obj_path As String) As Boolean
    Dim FSO As Object
    Dim InFile As Object
    Dim line As String
    Dim strParts() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim fileName As String
    Dim refName As String
    
    fileName = Dir$(obj_path & "references.csv")
    If Len(fileName) = 0 Then
        VCS_ImportReferences = False
        Exit Function
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(obj_path & fileName, iomode:=ForReading, create:=False, Format:=TristateFalse)
    
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        strParts = Split(line, ",")
        If UBound(strParts) = 2 Then 'a ref with a guid
          GUID = Trim$(strParts(0))
          Major = CLng(strParts(1))
          Minor = CLng(strParts(2))
          Application.References.AddFromGuid GUID, Major, Minor
        Else
          refName = Trim$(strParts(0))
          Application.References.AddFromFile refName
        End If
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    VCS_ImportReferences = True
    Exit Function
    
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
    
End Function

' Export References to a CSV
Public Sub VCS_ExportReferences(ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim line As String
    Dim ref As Reference

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(obj_path & "references.csv", overwrite:=True, Unicode:=False)
    For Each ref In Application.References
        If ref.GUID <> vbNullString Then ' references of types mdb,accdb,mde etc don't have a GUID
            If Not ref.BuiltIn Then
                line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
                OutFile.WriteLine line
            End If
        Else
            line = ref.FullPath
            OutFile.WriteLine line
        End If
    Next
    OutFile.Close
End Sub