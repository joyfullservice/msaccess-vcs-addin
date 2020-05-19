Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDevMode
' Author    : Adam Waller
' Date      : 5/15/2020
' Purpose   : Helper class to handle the parsing of saved print settings.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


' See the following links for additional technical details regarding the DEVMODE strcture:
' https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode
' https://stackoverflow.com/questions/49560317/64-bit-word-vba-devmode-dmduplex-returns-4
' http://toddmcdermid.blogspot.com/2009/02/microsoft-access-2003-and-printer.html


' DevMode for printer details
Private Type tDevModeBuffer
    strBuffer As String * 220 ' Pad with plenty of extra room.
End Type

Private Type tDevMode
    strDeviceName(1 To 32) As Byte
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
    strFormName(1 To 32) As Byte
    intUnusedPadding As Integer
    intBitsPerPel As Integer
    lngPelsWidth As Long
    lngPelsHeight As Long
    lngDisplayFlags As Long
    lngDisplayFrequency As Long
    lngICMMethod As Long
    lngICMIntent As Long
    lngMediaType As Long
    lngDitherType As Long
    lngReserved1 As Long
    lngReserved2 As Long
End Type


' Printer Margins
Private Type tMipBuffer
    strBuffer As String * 28
End Type

Private Type tMip
    xLeftMargin As Long
    yTopMargin As Long
    xRightMargin As Long
    yBotMargin As Long
    fDataOnly As Long
    xWidth As Long
    yHeight As Long
    fDefaultSize As Long
    cxColumns As Long
    yColumnSpacing As Long
    xRowSpacing As Long
    rItemLayout As Long
    fFastPrint As Long
    fDatasheet As Long
End Type


' Device Name (and if default printer)
Private Type tDevNamesBuffer
    strBuffer As String * 255
End Type

Private Type tDevNames
    intDriverOffset As Integer
    intDeviceOffset As Integer
    intOutputOffset As Integer
    intDefault As Integer
    strData(1 To 255) As Byte
End Type


' API calls for reading the DevMode for the default printer
Private Type PRINTER_DEFAULTS
   pDatatype As Long
   pDevmode As Long
   DesiredAccess As Long
End Type

Private Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevmode As Long               ' Pointer to DEVMODE
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long    ' Pointer to SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" _
    (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, _
    ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Private Declare PtrSafe Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
    (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" _
    (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare PtrSafe Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" _
    (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Byte, ByVal Command As Long) As Long
Private Declare PtrSafe Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" _
    (ByVal flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Long, _
    ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare PtrSafe Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
    (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare PtrSafe Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As LongPtr, pSource As LongPtr, ByVal cbLength As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" _
    (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, _
    lpOutput As Any, ByVal dev As Long) As Long



' Enum for types that can be expanded to friendly
' values for storing in version control.
Public Enum ePrintEnum
    ' Access constants
    epeColor
    epeDuplex
    epeOrientation
    epePaperBin
    epePaperSize
    epePrintQuality
    ' API values
    epeTTOption
    epeCollate
    epeDisplayFlags
    epeICMMethod
    epeICMIntent
    epeMediaType
    epeDitherType
    [_Last]
End Enum

' How to convert the enum value
Public Enum eEnumConversion
    eecAuto
    eecToEnum
    eecToName
End Enum


' Enums of printer constants for translating between
' values and friendly names.
Private m_dEnum(0 To ePrintEnum.[_Last] - 1) As Dictionary

' Form or Report object
Private m_objSource As Object

' Printer structures in native mode
' (Ready to write back to object)
Private m_tDevMode As tDevMode
Private m_tMip As tMip
Private m_tDevNames As tDevNames


'---------------------------------------------------------------------------------------
' Procedure : HasPrinterAssigned
' Author    : Adam Waller
' Date      : 5/18/2020
' Purpose   : Returns true if the report or form has a printer specifically assigned
'           : versus just using the default printer. (Make sure you load structures
'           : before checking this value.)
'---------------------------------------------------------------------------------------
'
Public Function HasPrinterAssigned() As Boolean
    HasPrinterAssigned = (m_tDevNames.intDefault = 0)
End Function


Public Sub LoadFromExportFile(strFile As String)

End Sub


Public Sub LoadFromJsonFile(strFile As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadFromDefaultPrinter
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Loads print settings from the default printer.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromDefaultPrinter()
    LoadFromPrinter Application.Printer.DeviceName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadFromPrinter
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Loads print settings for the specified printer through the API. This gives
'           : us a set of the basic print settings to which we can apply customizations.
'           : The following link was helpful in working out the details of doing this
'           : through the Windows API:
'           : http://www.lessanvaezi.com/changing-printer-settings-using-the-windows-api/
'---------------------------------------------------------------------------------------
'
Public Function LoadFromPrinter(strPrinter As String)

    ' API constants for reading printer properties
    Const READ_CONTROL = &H20000
    Const DM_IN_BUFFER = 8
    Const DM_OUT_BUFFER = 2
    Const PRINTER_ENUM_CONNECTIONS = &H4
    Const PRINTER_ENUM_LOCAL = &H2

    Dim hPrinter As Long
    Dim udtDefaults As PRINTER_DEFAULTS
    Dim lngReturn As Long
    Dim bteBuffer(0 To 219) As Byte
    Dim strBuffer As String
    Dim udtBuffer As tDevModeBuffer
    
    ' Clear our existing devmode structures
    ClearStructures
    
    ' Open a handle to read the default printer
    udtDefaults.DesiredAccess = READ_CONTROL
    lngReturn = OpenPrinter(strPrinter, hPrinter, udtDefaults)
    If lngReturn <> 0 And hPrinter <> 0 Then
        
        ' Check size of DevMode structure to make sure it fits in our buffer.
        lngReturn = DocumentProperties(0, hPrinter, strPrinter, 0, 0, 0)
        If lngReturn > 0 Then
        
            ' Read the devmode structure
            strBuffer = Space$(lngReturn + 100)
            lngReturn = DocumentProperties(0, hPrinter, strPrinter, StrPtr(strBuffer), 0, DM_OUT_BUFFER)
            If lngReturn > 0 Then
            
                ' Load into DevMode type
                udtBuffer.strBuffer = strBuffer
                LSet m_tDevMode = udtBuffer
            
            End If
        End If
    End If
    
    ' Close printer handle
    If hPrinter <> 0 Then ClosePrinter hPrinter

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFromReport
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Wrapper functions for loading objects by type.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromReport(strName As String)
    LoadFromObject acReport, strName
End Sub
Public Sub LoadFromForm(strName As String)
    LoadFromObject acForm, strName
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyTo
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Apply the adjustments to the specified class instance
'---------------------------------------------------------------------------------------
'
Public Sub ApplyToClass(ByRef cDevMode As clsDevMode)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyToObject
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Applies the settings to the object.
'---------------------------------------------------------------------------------------
'
Public Sub ApplyToObject(intType As AcObjectType, strName As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsDifferentFromDefault
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Returns true if the current settings differ from the default printer.
'---------------------------------------------------------------------------------------
'
Public Function IsDifferentFromDefault(Optional cDefault As clsDevMode) As Boolean

    ' Allow user to pass in default printer class to avoid reading it multiple times.
    If cDefault Is Nothing Then
        Set cDefault = New clsDevMode
        cDefault.LoadFromDefaultPrinter
    End If
    
    

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Return the loaded structures in a dictionary format. (For saving to
'           : Version Control.) Enums are translated to appropriate values.
'---------------------------------------------------------------------------------------
'
Public Function GetDictionary() As Dictionary
    
    Set GetDictionary = DevModeToDictionary()
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFromObject
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Load settings from a form or report object
'---------------------------------------------------------------------------------------
'
Private Sub LoadFromObject(intType As AcObjectType, strName As String)

    Dim udtDevModeBuffer As tDevModeBuffer
    Dim udtDevNamesBuffer As tDevNamesBuffer
    Dim udtMipBuffer As tMipBuffer
    Dim objSource As Object ' Could be a Form or Report

    ' Clear any existing structure data
    ClearStructures

    ' Open object
    Application.Echo False
    Select Case intType
        Case acForm
            DoCmd.OpenForm strName, acDesign, , , , acHidden
            Set objSource = Forms(strName)
        Case acReport
            DoCmd.OpenReport strName, acDesign, , , acHidden
            Set objSource = Reports(strName)
    End Select
    If objSource Is Nothing Then Exit Sub
    

    ' DevMode
    If Not IsNull(objSource.PrtDevMode) Then
        udtDevModeBuffer.strBuffer = objSource.PrtDevMode
        LSet m_tDevMode = udtDevModeBuffer
    End If
        
    ' DevNames
    If Not IsNull(objSource.PrtDevNames) Then
        udtDevNamesBuffer.strBuffer = objSource.PrtDevNames
        LSet m_tDevNames = udtDevNamesBuffer
    End If
    
    ' Mip (Margins)
    If Not IsNull(objSource.PrtMip) Then
        udtMipBuffer.strBuffer = objSource.PrtMip
        LSet m_tMip = udtMipBuffer
    End If

    ' Clean up
    Set objSource = Nothing
    DoCmd.Close intType, strName, acSaveNo
    Application.Echo True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearStructures
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Clear the existing structures.
'---------------------------------------------------------------------------------------
'
Private Sub ClearStructures()

    Dim tDevModeBlank As tDevMode
    Dim tMipBlank As tMip
    Dim tDevNamesBlank As tDevNames

    m_tDevMode = tDevModeBlank
    m_tMip = tMipBlank
    m_tDevNames = tDevNamesBlank
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DevModeToDictionary
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Convert a DEVMODE type to a dictionary. Only outputs meaningful fields.
'---------------------------------------------------------------------------------------
'
Private Function DevModeToDictionary() As Dictionary

    ' Constants to verify that the property is available
    Const DM_ORIENTATION = &H1
    Const DM_PAPERSIZE = &H2
    Const DM_PAPERLENGTH = &H4
    Const DM_PAPERWIDTH = &H8
    Const DM_SCALE = &H10
    Const DM_COPIES = &H100
    Const DM_DEFAULTSOURCE = &H200
    Const DM_PRINTQUALITY = &H400
    Const DM_COLOR = &H800
    Const DM_DUPLEX = &H1000
    Const DM_YRESOLUTION = &H2000
    Const DM_TTOPTION = &H4000
    Const DM_COLLATE = &H8000
    Const DM_FORMNAME = &H10000
    Const DM_LOGPIXELS = &H20000
    Const DM_BITSPERPEL = &H40000
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000
    Const DM_DISPLAYFLAGS = &H200000
    Const DM_DISPLAYFREQUENCY = &H400000
    Const DM_ICMMETHOD = &H800000
    Const DM_ICMINTENT = &H1000000
    Const DM_MEDIATYPE = &H2000000
    Const DM_DITHERTYPE = &H4000000
    Const DM_PANNINGWIDTH = &H20000000
    Const DM_PANNINGHEIGHT = &H40000000
    
    Dim lngFld As Long
    Dim cDM As tDevMode

    LSet cDM = m_tDevMode
    lngFld = cDM.lngFields
    
    Set DevModeToDictionary = New Dictionary
    
    With DevModeToDictionary
        .Add "DeviceName", NTrim(StrConv(cDM.strDeviceName, vbUnicode))
        '.Add "SpecVersion", cDM.intSpecVersion
        '.Add "DriverVersion", cDM.intDriverVersion
        '.Add "Size", cDM.intSize
        '.Add "DriverExtra", cDM.intDriverExtra
        '.Add "Fields", cDM.lngFields
        If BitSet(lngFld, DM_ORIENTATION) Then .Add "Orientation", GetEnum(epeOrientation, cDM.intOrientation)
        If BitSet(lngFld, DM_PAPERSIZE) Then .Add "PaperSize", GetEnum(epePaperSize, cDM.intPaperSize)
        If BitSet(lngFld, DM_PAPERLENGTH) Then .Add "PaperLength", cDM.intPaperLength
        If BitSet(lngFld, DM_PAPERWIDTH) Then .Add "PaperWidth", cDM.intPaperWidth
        If BitSet(lngFld, DM_SCALE) Then .Add "Scale", cDM.intScale
        If BitSet(lngFld, DM_COPIES) Then .Add "Copies", cDM.intCopies
        If BitSet(lngFld, DM_DEFAULTSOURCE) Then .Add "DefaultSource", GetEnum(epePaperBin, cDM.intDefaultSource)
        If BitSet(lngFld, DM_PRINTQUALITY) Then .Add "PrintQuality", GetEnum(epePrintQuality, cDM.intPrintQuality)
        If BitSet(lngFld, DM_COLOR) Then .Add "Color", GetEnum(epeColor, cDM.intColor)
        If BitSet(lngFld, DM_DUPLEX) Then .Add "Duplex", GetEnum(epeDuplex, cDM.intDuplex)
        If BitSet(lngFld, DM_YRESOLUTION) Then .Add "Resolution", cDM.intResolution
        If BitSet(lngFld, DM_TTOPTION) Then .Add "TTOption", GetEnum(epeTTOption, cDM.intTTOption)
        If BitSet(lngFld, DM_COLLATE) Then .Add "Collate", GetEnum(epeCollate, cDM.intCollate)
        If BitSet(lngFld, DM_FORMNAME) Then .Add "FormName", NTrim(StrConv(cDM.strFormName, vbUnicode))
        '.Add "UnusedPadding", cDM.intUnusedPadding
        '.Add "BitsPerPel", cDM.intBitsPerPel
        '.Add "PelsWidth", cDM.lngPelsWidth
        '.Add "PelsHeight", cDM.lngPelsHeight
        If BitSet(lngFld, DM_DISPLAYFLAGS) Then .Add "DisplayFlags", GetEnum(epeDisplayFlags, cDM.lngDisplayFlags)
        If BitSet(lngFld, DM_DISPLAYFREQUENCY) Then .Add "DisplayFrequency", cDM.lngDisplayFrequency
        If BitSet(lngFld, DM_ICMMETHOD) Then .Add "ICMMethod", GetEnum(epeICMMethod, cDM.lngICMMethod)
        If BitSet(lngFld, DM_ICMINTENT) Then .Add "ICMIntent", GetEnum(epeICMIntent, cDM.lngICMIntent)
        If BitSet(lngFld, DM_MEDIATYPE) Then .Add "MediaType", GetEnum(epeMediaType, cDM.lngMediaType)
        If BitSet(lngFld, DM_DITHERTYPE) Then .Add "DitherType", GetEnum(epeDitherType, cDM.lngDitherType)
        '.Add "Reserved1", cDM.lngReserved1
        '.Add "Reserved2", cDM.lngReserved2
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : BitSet
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Returns true if the flag is set.
'---------------------------------------------------------------------------------------
'
Private Function BitSet(lngFlags As Long, lngValue As Long) As Boolean
    BitSet = CBool((lngFlags And lngValue) = lngValue)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DictionaryToDevMode
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Excel formulas make it easy to edit these!
'---------------------------------------------------------------------------------------
'
Private Function DictionaryToDevMode(dDevMode As Dictionary) As tDevMode
    
    Dim intCnt As Integer
    Dim varDevice As Variant
    Dim varForm As Variant
    
    ' Assign regular properties
    With DictionaryToDevMode
        .intSpecVersion = dDevMode("SpecVersion")
        .intDriverVersion = dDevMode("DriverVersion")
        .intSize = dDevMode("Size")
        .intDriverExtra = dDevMode("DriverExtra")
        .lngFields = dDevMode("Fields")
        .intOrientation = GetEnum(epeOrientation, dDevMode("Orientation"), eecToEnum)
        .intPaperSize = GetEnum(epePaperSize, dDevMode("PaperSize"), eecToEnum)
        .intPaperLength = dDevMode("PaperLength")
        .intPaperWidth = dDevMode("PaperWidth")
        .intScale = dDevMode("Scale")
        .intCopies = dDevMode("Copies")
        .intDefaultSource = GetEnum(epePaperBin, dDevMode("DefaultSource"), eecToEnum)
        .intPrintQuality = GetEnum(epePrintQuality, dDevMode("PrintQuality"), eecToEnum)
        .intColor = GetEnum(epeColor, dDevMode("Color"), eecToEnum)
        .intDuplex = GetEnum(epeDuplex, dDevMode("Duplex"), eecToEnum)
        .intResolution = dDevMode("Resolution")
        .intTTOption = GetEnum(epeTTOption, dDevMode("TTOption"), eecToEnum)
        .intCollate = GetEnum(epeCollate, dDevMode("Collate"), eecToEnum)
        .intUnusedPadding = dDevMode("UnusedPadding")
        .intBitsPerPel = dDevMode("BitsPerPel")
        .lngPelsWidth = dDevMode("PelsWidth")
        .lngPelsHeight = dDevMode("PelsHeight")
        .lngDisplayFlags = GetEnum(epeDisplayFlags, dDevMode("DisplayFlags"), eecToEnum)
        .lngDisplayFrequency = dDevMode("DisplayFrequency")
        .lngICMMethod = GetEnum(epeICMMethod, dDevMode("ICMMethod"), eecToEnum)
        .lngICMIntent = GetEnum(epeICMIntent, dDevMode("ICMIntent"), eecToEnum)
        .lngMediaType = GetEnum(epeMediaType, dDevMode("MediaType"), eecToEnum)
        .lngDitherType = GetEnum(epeDitherType, dDevMode("DitherType"), eecToEnum)
        .lngReserved1 = dDevMode("Reserved1")
        .lngReserved2 = dDevMode("Reserved2")
    
        ' Assign byte arrays for string values
        varDevice = GetNullTermByteArray(dDevMode("DeviceName"), 32)
        varForm = GetNullTermByteArray(dDevMode("FormName"), 32)
        For intCnt = 1 To 32
            .strDeviceName(intCnt) = varDevice(intCnt)
            .strFormName(intCnt) = varForm(intCnt)
        Next intCnt
    End With
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetEnum
' Author    : Adam Waller
' Date      : 5/15/2020
' Purpose   : Return an enum value, 0 or UNKNOWN if not found.
'---------------------------------------------------------------------------------------
'
Public Function GetEnum(eType As ePrintEnum, varValue, Optional Convert As eEnumConversion = eecAuto) As Variant

    Dim varKey As Variant
    Dim varReturn As Variant
    
    ' Build cached enum on first request
    If m_dEnum(eType) Is Nothing Then BuildEnum eType
    
    ' Determine conversion type
    If Convert = eecAuto Then
        If IsNumeric(varValue) Then
            Convert = eecToName
        Else
            Convert = eecToEnum
        End If
    End If
    
    ' By default, just return the original value if we don't find a match.
    varReturn = varValue
    
    ' See if we are trying to return the description or the enum value
    If Convert = eecToName Then
        If m_dEnum(eType).Exists(varValue) Then varReturn = m_dEnum(eType)(varValue)
    Else
        ' Search for matching description
        For Each varKey In m_dEnum(eType).Keys
            If m_dEnum(eType)(varKey) = varValue Then
                varReturn = varKey
                Exit For
            End If
        Next varKey
    End If
    
    ' Return value (if found)
    GetEnum = varReturn
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildEnum
' Author    : Adam Waller
' Date      : 5/15/2020
' Purpose   : Builds out the pseudo enum values for constants. (Makes the values easier
'           : to understand when stored in version control.)
'           : These values were from the Access 2010 help file.
'           : Additional enums can be added later if desired, such as media type or
'           : true type font handling. See the following link for constants:
'           : http://www.jasinskionline.com/windowsapi/ref/d/devmode.html
'---------------------------------------------------------------------------------------
'
Private Function BuildEnum(eType As ePrintEnum)

    Set m_dEnum(eType) = New Dictionary
    With m_dEnum(eType)
        
        Select Case eType
            Case epeColor
                .Add acPRCMColor, "Color"
                .Add acPRCMMonochrome, "Monochrome"
                
            Case epeDuplex
                .Add acPRDPHorizontal, "Horizontal"
                .Add acPRDPSimplex, "Simplex"
                .Add acPRDPVertical, "Vertical"
                
            Case epePrintQuality
                .Add acPRPQDraft, "Draft"
                .Add acPRPQHigh, "High"
                .Add acPRPQLow, "Low"
                .Add acPRPQMedium, "Medium"
                
            Case epeOrientation
                .Add acPRORLandscape, "Landscape"
                .Add acPRORPortrait, "Portrait"
                
            Case epePaperBin
                .Add acPRBNAuto, "Auto"
                .Add acPRBNCassette, "Cassette"
                .Add acPRBNEnvelope, "Envelope"
                .Add acPRBNEnvManual, "Envelope Manual"
                .Add acPRBNFormSource, "Form Source"
                .Add acPRBNLargeCapacity, "Large Capacity"
                .Add acPRBNLargeFmt, "Large Format"
                .Add acPRBNLower, "Lower"
                .Add acPRBNManual, "Manual"
                .Add acPRBNMiddle, "Middle"
                .Add acPRBNSmallFmt, "Small Format"
                .Add acPRBNTractor, "Tractor"
                .Add acPRBNUpper, "Upper"
                
            Case epePaperSize
                .Add acPRPS10x14, "10x14"
                .Add acPRPS11x17, "11x17"
                .Add acPRPSA3, "A3"
                .Add acPRPSA4, "A4"
                .Add acPRPSA4Small, "A4 Small"
                .Add acPRPSA5, "A5"
                .Add acPRPSB4, "B4"
                .Add acPRPSB5, "B5"
                .Add acPRPSCSheet, "C Size Sheet"
                .Add acPRPSDSheet, "D Size Sheet"
                .Add acPRPSEnv10, "Envelope #10"
                .Add acPRPSEnv11, "Envelope #11"
                .Add acPRPSEnv12, "Envelope #12"
                .Add acPRPSEnv14, "Envelope #14"
                .Add acPRPSEnv9, "Envelope #9"
                .Add acPRPSEnvB4, "Envelope B4"
                .Add acPRPSEnvB5, "Envelope B5"
                .Add acPRPSEnvB6, "Envelope B6"
                .Add acPRPSEnvC3, "Envelope C3"
                .Add acPRPSEnvC4, "Envelope C4"
                .Add acPRPSEnvC5, "Envelope C5"
                .Add acPRPSEnvC6, "Envelope C6"
                .Add acPRPSEnvC65, "Envelope C65"
                .Add acPRPSEnvDL, "Envelope DL"
                .Add acPRPSEnvItaly, "Italian Envelope"
                .Add acPRPSEnvMonarch, "Monarch Envelope"
                .Add acPRPSEnvPersonal, "Envelope"
                .Add acPRPSESheet, "E Size Sheet"
                .Add acPRPSExecutive, "Executive"
                .Add acPRPSFanfoldLglGerman, "German Legal Fanfold"
                .Add acPRPSFanfoldStdGerman, "German Standard Fanfold"
                .Add acPRPSFanfoldUS, "U.S. Standard Fanfold"
                .Add acPRPSFolio, "Folio"
                .Add acPRPSLedger, "Ledger"
                .Add acPRPSLegal, "Legal"
                .Add acPRPSLetter, "Letter"
                .Add acPRPSLetterSmall, "Letter Small"
                .Add acPRPSNote, "Note"
                .Add acPRPSQuarto, "Quarto"
                .Add acPRPSStatement, "Statement"
                .Add acPRPSTabloid, "Tabloid"
                .Add acPRPSUser, "User-Defined"
            
            '--------------------
            ' API constants
            '--------------------
            
            Case epeTTOption
                .Add 1, "Bitmap"
                .Add 2, "Download"
                .Add 3, "Bitmap/Download"
                .Add 4, "Substitute"
            
            Case epeCollate
                .Add 0, "False"
                .Add 1, "True"
            
            Case epeDisplayFlags
                .Add 1, "Grayscale"
                .Add 2, "Interlaced"
                
            Case epeICMMethod
                .Add 1, "None"
                .Add 2, "System"
                .Add 3, "Driver"
                .Add 4, "Device"
            
            Case epeICMIntent
                .Add 1, "Saturate"
                .Add 2, "Contrast"
                .Add 3, "Colormetric"
                
            Case epeMediaType
                .Add 1, "Standard"
                .Add 2, "Glossy"
                .Add 3, "Transparency"
            
            Case epeDitherType
                .Add 1, "None"
                .Add 2, "Coarse"
                .Add 3, "Fine"
                .Add 4, "Line Art"
                .Add 5, "Grayscale"
            
        End Select
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : NTrim
' Author    : Adam Waller
' Date      : 5/18/2020
' Purpose   : Trim a string to a null character terminator.
'---------------------------------------------------------------------------------------
'
Private Function NTrim(strText) As String
    Dim lngPos As Long
    lngPos = InStr(1, strText, vbNullChar)
    If lngPos > 0 Then
        NTrim = Left$(strText, lngPos - 1)
    Else
        NTrim = strText
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNullTermStringByOffset
' Author    : Adam Waller
' Date      : 5/18/2020
' Purpose   : Returns the value of a null-terminated string by offset.
'---------------------------------------------------------------------------------------
'
Private Function GetNullTermStringByOffset(strData As String, lngHeaderLen As Long, intOffset As Integer) As String
    
    Dim lngNull As Long
    Dim lngStart As Long
    
    lngStart = intOffset - lngHeaderLen
    lngNull = InStr(lngStart, strData, vbNullChar)
    
    ' Return the string if we found a null terminator
    If lngNull > 0 Then GetNullTermStringByOffset = Mid$(strData, lngStart, lngNull - lngStart)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNullTermByteArray
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Convert a string to a null terminated byte array.
'---------------------------------------------------------------------------------------
'
Private Function GetNullTermByteArray(strValue As String, lngLen As Long) As Byte

    Dim strReturn As String
    Dim bteReturn() As Byte
    
    ' Build return string with buffer
    strReturn = strValue & vbNullChar & Space$(lngLen - (Len(strValue) + 1))

    bteReturn = strReturn
    GetNullTermByteArray = bteReturn
    
End Function