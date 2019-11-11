Option Explicit
Option Compare Database
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ExportReferences
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Export references to a csv text file
'---------------------------------------------------------------------------------------
'
Public Sub ExportReferences(strFolder As String, cModel As IVersionControl)
    
    Dim cData As New clsConcat
    Dim ref As Reference
    Dim intCnt As Integer
    
    For Each ref In Application.References
        If ref.GUID = "" Then ' references of types mdb,accdb,mde etc don't have a GUID
            cData.Add ref.FullPath
            cData.Add vbCrLf
            cModel.Log "  [" & ref.Name & "]", cModel.ShowDebug
       Else
            With cData
                .Add ref.GUID
                .Add ","
                .Add ref.Name
                .Add ","
                .Add CStr(ref.Major)
                .Add ","
                .Add CStr(ref.Minor)
                .Add vbCrLf
            End With
            cModel.Log "  " & ref.Name & " " & ref.Major & "." & ref.Minor, cModel.ShowDebug
        End If
        intCnt = intCnt + 1
    Next ref
    
    ' Write ouput to file
    WriteFile cData.GetStr, strFolder & "references.csv"

    ' Show summary
    If cModel.ShowDebug Then
        cModel.Log "[" & intCnt & "] references exported."
    Else
        cModel.Log "[" & intCnt & "]"
    End If
    
End Sub


' Import References from a CSV, true=SUCCESS
Public Function ImportReferences(obj_path As String, Optional ShowDebugInfo As Boolean) As Boolean
    
    Dim InFile As Scripting.TextStream
    Dim line As String
    Dim Item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim FileName As String
    Dim refName As String
    Dim obj_count As Integer
    
    FileName = Dir(obj_path & "references.csv")
    If Len(FileName) = 0 Then
        ImportReferences = False
        Exit Function
    Else
        Debug.Print PadRight("Importing references...", 24);
        obj_count = 0
    End If
    Set InFile = FSO.OpenTextFile(obj_path & FileName, ForReading)
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        Item = Split(line, ",")
'        If UBound(Item) = 2 Then 'a ref with a guid
        If UBound(Item) = 3 Then 'a ref with a guid
          GUID = Trim(Item(0))
'          Major = CLng(Item(1))
          Major = CLng(Item(2))
'          Minor = CLng(Item(2))
          Minor = CLng(Item(3))
          Application.References.AddFromGuid GUID, Major, Minor
          refName = GUID
        Else
          refName = Trim(Item(0))
          Application.References.AddFromFile refName
        End If
        If ShowDebugInfo Then
            If obj_count = 0 Then
                Debug.Print
            End If
            Debug.Print "  [debug] reference " & refName;
            Debug.Print
        End If
        obj_count = obj_count + 1
go_on:
    Loop
    Debug.Print "[" & obj_count & "]"

On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    ImportReferences = True
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