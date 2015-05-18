Option Compare Database
Option Private Module
Option Explicit

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
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
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
    
    Dim OutFile As Object
    Set OutFile = FSO.CreateTextFile(filePath, True)
    
    'print out print var values
    OutFile.WriteLine DM.intOrientation
    OutFile.WriteLine DM.intPaperSize
    OutFile.WriteLine DM.intPaperLength
    OutFile.WriteLine DM.intPaperWidth
    OutFile.WriteLine DM.intScale
    OutFile.Close
    
    Set rpt = Nothing
    
    DoCmd.Close acReport, obj_name, acSaveYes
    Application.Echo True
    VBE.ActiveCodePane.Show
    
End Sub

Public Sub ImportPrintVars(obj_name As String, filePath As String)

Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

 Dim DevModeString As str_DEVMODE
 Dim DevModeExtra As String
 
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
 
 Dim InFile As Object
 Set InFile = FSO.OpenTextFile(filePath, ForReading)
 
 'print out print var values
 DM.intOrientation = InFile.ReadLine
 DM.intPaperSize = InFile.ReadLine
 DM.intPaperLength = InFile.ReadLine
 DM.intPaperWidth = InFile.ReadLine
 DM.intScale = InFile.ReadLine
 InFile.Close
 
'write print vars back into report
LSet DevModeString = DM
 Mid(DevModeExtra, 1, 94) = DevModeString.RGB
 rpt.PrtDevMode = DevModeExtra

 Set rpt = Nothing

DoCmd.Close acReport, obj_name, acSaveYes
End Sub