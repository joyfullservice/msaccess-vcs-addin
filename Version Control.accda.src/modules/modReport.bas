Option Explicit
Option Compare Database
Option Private Module

Private Type str_DEVMODE
 RGB As String * 94
End Type

Private Type type_DEVMODE
    strDeviceName(31) As Byte 'vba strings are encoded in unicode (16 bit) not ascii
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    intOrientation As Integer
    intPaperSize As Integer
    intPaperLength As Integer
    intPaperWidth As Integer
    intScale As Integer
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    intResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName(31) As Byte
    lngPad As Long
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type



'Exports print vars for reports
Public Sub ExportPrintVars(obj_name As String, filePath As String)
    DoEvents
    Dim fso As New Scripting.FileSystemObject
    
    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String
    Dim DM As type_DEVMODE
    Dim rpt As Report
    
    'report must be open to access Report object
    'report must be opened in design view to save changes to the print vars
    Application.Echo False
    DoCmd.OpenReport obj_name, acViewDesign
    Set rpt = Reports(obj_name)
    rpt.Visible = False
    ' Move focus back to IDE
    VBE.ActiveCodePane.Show
    
    'read print vars into struct
    If Not IsNull(rpt.PrtDevMode) Then
       DevModeExtra = rpt.PrtDevMode
       DevModeString.RGB = DevModeExtra
       LSet DM = DevModeString
    Else
       Set rpt = Nothing
       DoCmd.Close acReport, obj_name, acSaveNo
       Debug.Print "Warning: PrtDevMode is null"
       Exit Sub
    End If
    
    Dim OutFile As Scripting.TextStream
    Set OutFile = fso.CreateTextFile(filePath, True)
    
    'print out print var values
    OutFile.WriteLine "Orientation=" & DM.intOrientation
    OutFile.WriteLine "PaperSize=" & DM.intPaperSize
    OutFile.WriteLine "PaperLength=" & DM.intPaperLength
    OutFile.WriteLine "PaperWidth=" & DM.intPaperWidth
    OutFile.WriteLine "Scale=" & DM.intScale
    OutFile.Close
    
    Set rpt = Nothing
    
    DoCmd.Close acReport, obj_name, acSaveNo ' acSaveYes
    Application.Echo True
    VBE.ActiveCodePane.Show
    
End Sub


Public Sub ImportPrintVars(obj_name As String, filePath As String)

Dim fso As New Scripting.FileSystemObject

 Dim DevModeString As str_DEVMODE
 Dim DevModeExtra As String
 Dim varLine As Variant
 
 Dim DM As type_DEVMODE
  Dim rpt As Report
 'report must be open to access Report object
 'report must be opened in design view to save changes to the print vars

  DoCmd.OpenReport obj_name, acViewDesign

 Set rpt = Reports(obj_name)

'read print vars into struct
 If Not IsNull(rpt.PrtDevMode) Then
    DevModeExtra = rpt.PrtDevMode
    DevModeString.RGB = DevModeExtra
    LSet DM = DevModeString
 Else
    Set rpt = Nothing
    DoCmd.Close acReport, obj_name, acSaveNo
    Debug.Print "Warning: PrtDevMode is null"
    Exit Sub
 End If
 
 Dim InFile As Scripting.TextStream ' Object
 Set InFile = fso.OpenTextFile(filePath, ForReading)
 
 ' Loop through lines
 Do While Not InFile.AtEndOfStream
    varLine = Split(InFile.ReadLine, "=")
    If UBound(varLine) = 1 Then
        Select Case varLine(0)
            Case "Orientation":     DM.intOrientation = varLine(1)
            Case "PaperSize":       DM.intPaperSize = varLine(1)
            Case "PaperLength":     DM.intPaperLength = varLine(1)
            Case "PaperWidth":      DM.intPaperWidth = varLine(1)
            Case "Scale":           DM.intScale = varLine(1)
            Case Else
                Debug.Print "* Unknown print var: '" & varLine(0) & "'"
        End Select
    End If
 Loop
 
 InFile.Close
 
'write print vars back into report
LSet DevModeString = DM
 Mid(DevModeExtra, 1, 94) = DevModeString.RGB
 rpt.PrtDevMode = DevModeExtra

 Set rpt = Nothing

DoCmd.Close acReport, obj_name, acSaveYes
End Sub