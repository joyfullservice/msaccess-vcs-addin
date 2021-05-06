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
Private Const ModuleName As String = "clsDevMode"

' See the following links for additional technical details regarding the DEVMODE strcture:
' https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode
' https://stackoverflow.com/questions/49560317/64-bit-word-vba-devmode-dmduplex-returns-4
' http://toddmcdermid.blogspot.com/2009/02/microsoft-access-2003-and-printer.html
' https://github.com/x-ware-ltd/access-scc-addin/blob/master/Modules/modExtendSaveAsText.ACM

' Constant to convert tenths of millimeters to inches for human readability
Private Const TEN_MIL As Double = 0.00393701

' API constants for reading printer properties
' These may not be needed any longer but are kept here for referencing.
'Private Const READ_CONTROL = &H20000
'Private Const PRINTER_ACCESS_USE = &H8
'Private Const GENERIC_READ = &H80000000
Private Const DM_OUT_BUFFER = 2

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

Private Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" _
    (ByVal pPrinterName As String, phPrinter As LongPtr, pDefault As Any) As Long
Private Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
Private Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" _
    (ByVal hwnd As Long, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, _
    ByVal pDevModeOutput As LongPtr, ByVal pDevModeInput As LongPtr, ByVal fMode As Long) As Long

' Enum for types that can be expanded to friendly
' values for storing in version control.
Public Enum ePrintEnum
    ' Access constants
    epeColor
    epeColumnLayout
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

' Constants to verify that the property is available
Private Enum edmFlags
    DM_ORIENTATION = &H1
    DM_PAPERSIZE = &H2
    DM_PAPERLENGTH = &H4
    DM_PAPERWIDTH = &H8
    DM_SCALE = &H10
    DM_COPIES = &H100
    DM_DEFAULTSOURCE = &H200
    DM_PRINTQUALITY = &H400
    DM_COLOR = &H800
    DM_DUPLEX = &H1000
    DM_YRESOLUTION = &H2000
    DM_TTOPTION = &H4000
    DM_COLLATE = &H8000
    DM_FORMNAME = &H10000
    'DM_LOGPIXELS = &H20000
    'DM_BITSPERPEL = &H40000
    'DM_PELSWIDTH = &H80000
    'DM_PELSHEIGHT = &H100000
    DM_DISPLAYFLAGS = &H200000
    DM_DISPLAYFREQUENCY = &H400000
    DM_ICMMETHOD = &H800000
    DM_ICMINTENT = &H1000000
    DM_MEDIATYPE = &H2000000
    DM_DITHERTYPE = &H4000000
    'DM_PANNINGWIDTH = &H20000000
    'DM_PANNINGHEIGHT = &H40000000
End Enum


' Enums of printer constants for translating between
' values and friendly names.
Private m_dEnum(0 To ePrintEnum.[_Last] - 1) As Dictionary

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


'---------------------------------------------------------------------------------------
' Procedure : LoadFromExportFile
' Author    : Adam Waller
' Date      : 5/20/2020
' Purpose   : Load sections from export file
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromExportFile(strFile As String)

    Dim varLines As Variant
    Dim lngLine As Long
    Dim cBlock(1 To 3) As clsConcat
    Dim cBuffer(1 To 3) As clsConcat
    Dim strHex As String
    Dim strChar As String
    Dim bteBuffer() As Byte
    Dim intBlock As Integer
    Dim strLine As String
    Dim lngChar As Long
    Dim lngPos As Long
    Dim udtMipBuffer As tMipBuffer
    Dim udtDevModeBuffer As tDevModeBuffer
    Dim udtDevNamesBuffer As tDevNamesBuffer
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Blocks: 1=Mip, 2=DevMode, 3=DevNames

    ' Clear existing structures and create block classes.
    ClearStructures

    If Not FSO.FileExists(strFile) Then Exit Sub
    
    ' Open the export file, checking to see if it is in UCS format
    If HasUcs2Bom(strFile) Then
        varLines = Split(ReadFile(strFile, "Unicode"), vbCrLf)
    Else
        varLines = Split(ReadFile(strFile), vbCrLf)
    End If
    
    ' Read the text file line by line, loading the block data
    Perf.OperationStart "Read File DevMode"
    For lngLine = 0 To UBound(varLines)
        strLine = Trim$(varLines(lngLine))
        ' Look for header if not inside block
        If intBlock = 0 Then
            ' Check for header
            Select Case strLine
                Case "PrtMip = Begin":      intBlock = 1
                Case "PrtDevMode = Begin":  intBlock = 2
                Case "PrtDevNames = Begin": intBlock = 3
            End Select
        Else
            ' Inside block
            If strLine = "End" Then
                intBlock = 0
                ' Exit loop after adding all the blocks
                If Not (cBlock(1) Is Nothing _
                    Or cBlock(2) Is Nothing _
                    Or cBlock(3) Is Nothing) Then Exit For
            ElseIf Left$(strLine, 2) = "0x" Then
                ' Create block class, if it doesn't exist
                If cBlock(intBlock) Is Nothing Then Set cBlock(intBlock) = New clsConcat
                ' Add bytes after the "0x" prefix, and before the " ,"
                ' at the end of the line.
                cBlock(intBlock).Add Mid$(strLine, 3, Len(strLine) - 4)
            ElseIf strLine = "Begin" Then
                ' Reached the end of the header section. We should
                ' have already exited the loop, but just in case...
                Exit For
            End If
        End If
    Next lngLine
    
    ' Convert hex block data to string
    strChar = "&h00"
    For intBlock = 1 To 3
        ' Block will not be created if not found in source file.
        ' (Such as a file that was already sanitized.)
        If Not cBlock(intBlock) Is Nothing Then
            strHex = cBlock(intBlock).GetStr
            Set cBuffer(intBlock) = New clsConcat
            ' Each two hex characters represent one bit
            ReDim bteBuffer(0 To (Len(strHex) / 2) + 1)
            ' Loop through each set of 2 characters to get bytes
            For lngChar = 1 To Len(strHex) Step 2
                ' Apply two characters to buffer. (Faster than concatenating strings)
                Mid$(strChar, 3, 2) = Mid$(strHex, lngChar, 2)
                lngPos = ((lngChar + 1) / 2) - 1
                bteBuffer(lngPos) = CLng(strChar)
            Next lngChar
            Select Case intBlock
                Case 1
                    udtMipBuffer.strBuffer = bteBuffer
                    LSet m_tMip = udtMipBuffer
                Case 2
                    udtDevModeBuffer.strBuffer = bteBuffer
                    LSet m_tDevMode = udtDevModeBuffer
                Case 3
                    udtDevNamesBuffer.strBuffer = bteBuffer
                    LSet m_tDevNames = udtDevNamesBuffer
            End Select
        End If
    Next intBlock
    Perf.OperationEnd

    CatchAny eelError, "Error loading printer settings from file: " & strFile, _
        ModuleName & ".LoadFromExportFile", True, True

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
Public Sub LoadFromPrinter(strPrinter As String)

    Dim hPrinter As LongPtr
    Dim udtDefaults As PRINTER_DEFAULTS
    Dim lngReturn As Long
    Dim strBuffer As String
    Dim udtBuffer As tDevModeBuffer
    Dim objPrinter As Access.Printer

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Clear our existing devmode structures
    ClearStructures
    
    ' Open a handle to read the default printer
    lngReturn = OpenPrinter(strPrinter, hPrinter, ByVal 0&)

    CatchAny eelError, "Error getting printer pointer " & strPrinter, ModuleName & ".LoadFromPrinter", True, True
    If lngReturn <> 0 And hPrinter <> 0 Then

        ' Check size of DevMode structure to make sure it fits in our buffer.
        lngReturn = DocumentProperties(0, hPrinter, strPrinter, 0, 0, 0)
        If lngReturn > 0 Then
            ' Read the devmode structure
            strBuffer = NullPad(lngReturn + 100)
            lngReturn = DocumentProperties(0, hPrinter, strPrinter, StrPtr(strBuffer), 0, DM_OUT_BUFFER)
            
            If lngReturn > 0 Then
                ' Load into DevMode type
                udtBuffer.strBuffer = strBuffer
                LSet m_tDevMode = udtBuffer
            
            End If
        Else
            Log.Error eelWarning, "There has been an error with loading DevMode structure. lngReturn:'" & lngReturn & "'", _
                ModuleName & ".LoadFromPrinter"
        End If
    End If

    CatchAny eelError, "Error getting printer devMode " & strPrinter, ModuleName & ".LoadFromPrinter", True, True
    ' Close printer handle
    If hPrinter <> 0 Then ClosePrinter hPrinter
    
    ' Attempt to load the printer object
    Set objPrinter = GetPrinterByName(strPrinter)

    If objPrinter Is Nothing Then
        Log.Error eelWarning, "Could not find printer '" & strPrinter & "' on this system.", _
            ModuleName & ".LoadFromPrinter"
    Else
        ' Load in the DevNames structure
        If Options.ShowDebug Then Log.Add "Loading Printer info for: '" & strPrinter & "'."
        
        SetDevNames objPrinter
        ' Load in the margin defaults
        SetMipFromPrinter objPrinter
    End If

    CatchAny eelError, "Error with printer devMode " & strPrinter, _
        ModuleName & ".LoadFromPrinter", True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadFromReport/Form
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Wrapper functions for loading objects by type.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromReport(rptReport As Access.Report)
    LoadFromObject rptReport
End Sub
Public Sub LoadFromForm(frmForm As Access.Form)
    LoadFromObject frmForm
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Return the loaded structures in a dictionary format. (For saving to
'           : Version Control.) Enums are translated to appropriate values.
'---------------------------------------------------------------------------------------
'
Public Function GetDictionary() As Dictionary
    Set GetDictionary = New Dictionary
    With GetDictionary
        ' Only add device information if not using the default printer.
        If DevNamesHasData And (m_tDevNames.intDefault = 0) Then .Add "Device", DevNamesToDictionary()
        If DevModeHasData Then .Add "Printer", DevModeToExport()
        If MipHasData Then .Add "Margins", MipToDictionary()
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : DevModeToExport
' Author    : Adam Waller
' Date      : 11/9/2020
' Purpose   : Return a dictionary of the DevMode settings that we have selected
'           : to export, based on the current options.
'---------------------------------------------------------------------------------------
'
Private Function DevModeToExport() As Dictionary

    Dim varKey As Variant
    Dim dDM As Dictionary
    Dim dOpt As Dictionary
    
    Set dDM = DevModeToDictionary
    Set dOpt = Options.ExportPrintSettings
    Set DevModeToExport = New Dictionary

    With DevModeToExport
        For Each varKey In dDM.Keys
            If dOpt.Exists(varKey) Then
                If CBool(dOpt(varKey)) Then
                    .Add varKey, dDM(varKey)
                End If
            End If
        Next varKey
    End With
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : DevModeHasData
' Author    : Adam Waller
' Date      : 11/4/2020
' Purpose   : Wrapper functions to ensure that we have loaded data into the following
'           : structures. Sometimes an export file may not contain all of these
'           : sections, in which case we should not attempt to map it to a dictionary.
'---------------------------------------------------------------------------------------
'
Private Function DevModeHasData() As Boolean
    ' Should have fields flag set
    DevModeHasData = (m_tDevMode.lngFields > 0)
End Function
Private Function DevNamesHasData() As Boolean
    ' Look for a driver offset. (Should always have this, if set.)
    DevNamesHasData = (m_tDevNames.intDriverOffset > 0)
End Function
Private Function MipHasData() As Boolean
    ' Item layout should either be 1953 or 1954
    MipHasData = (m_tMip.rItemLayout > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasData
' Author    : Adam Waller
' Date      : 1/14/2021
' Purpose   : Returns true if we have data in any of the three structures.
'---------------------------------------------------------------------------------------
'
Public Function HasData() As Boolean
    HasData = (DevModeHasData Or DevNamesHasData Or MipHasData)
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFromObject
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Load settings from a form or report object
'---------------------------------------------------------------------------------------
'
Private Sub LoadFromObject(objSource As Object)

    Dim udtDevModeBuffer As tDevModeBuffer
    Dim udtDevNamesBuffer As tDevNamesBuffer
    Dim udtMipBuffer As tMipBuffer

    ' Clear any existing structure data
    ClearStructures

    ' DevMode
    If Not IsNull(objSource.PrtDevMode) Then
        udtDevModeBuffer.strBuffer = objSource.PrtDevMode
        LSet m_tDevMode = udtDevModeBuffer
    End If
        
    ' DevNames
    If Not IsNull(objSource.PrtDevNames) Then
        With udtDevNamesBuffer
            ' Pad right side of buffer with nulls rather than spaces.
            .strBuffer = objSource.PrtDevNames & NullPad(Len(.strBuffer) - Len(objSource.PrtDevNames))
        End With
        LSet m_tDevNames = udtDevNamesBuffer
    End If
    
    ' Mip (Margins)
    If Not IsNull(objSource.PrtMip) Then
        udtMipBuffer.strBuffer = objSource.PrtMip
        LSet m_tMip = udtMipBuffer
    End If

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
    
    Dim strName As String
    Dim lngFld As Long
    Dim cDM As tDevMode

    LSet cDM = m_tDevMode
    lngFld = cDM.lngFields
    
    Set DevModeToDictionary = New Dictionary
    
    With DevModeToDictionary
        strName = NTrim(StrConv(cDM.strDeviceName, vbUnicode))
        ' Only save the printer name here if it is not the default printer.
        If strName <> vbNullString And m_tDevNames.intDefault = 0 Then .Add "DeviceName", strName
        '.Add "SpecVersion", cDM.intSpecVersion
        '.Add "DriverVersion", cDM.intDriverVersion
        '.Add "Size", cDM.intSize
        '.Add "DriverExtra", cDM.intDriverExtra
        '.Add "Fields", cDM.lngFields
        If BitSet(lngFld, DM_ORIENTATION) Then .Add "Orientation", GetEnum(epeOrientation, cDM.intOrientation)
        If BitSet(lngFld, DM_PAPERSIZE) Then .Add "PaperSize", GetEnum(epePaperSize, cDM.intPaperSize)
        If BitSet(lngFld, DM_PAPERLENGTH) Then .Add "PaperLength", Round(cDM.intPaperLength * TEN_MIL, 2)
        If BitSet(lngFld, DM_PAPERWIDTH) Then .Add "PaperWidth", Round(cDM.intPaperWidth * TEN_MIL, 2)
        If BitSet(lngFld, DM_SCALE) Then .Add "Scale", cDM.intScale
        If BitSet(lngFld, DM_COPIES) And cDM.intCopies > 1 Then .Add "Copies", cDM.intCopies    ' Only add for more than 1 copy
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
' Procedure : MipToDictionary
' Author    : Adam Waller
' Date      : 6/2/2020
' Purpose   : Convert the printer margins to a dictionary object. Use inches for sizes
'           : to match what the user sees in the margin dialogs. (1440 twips per inch)
'---------------------------------------------------------------------------------------
'
Private Function MipToDictionary() As Dictionary

    Dim cMip As tMip

    LSet cMip = m_tMip
    Set MipToDictionary = New Dictionary
    
    With MipToDictionary
        .Add "LeftMargin", GetInch(cMip.xLeftMargin)
        .Add "TopMargin", GetInch(cMip.yTopMargin)
        .Add "RightMargin", GetInch(cMip.xRightMargin)
        .Add "BotMargin", GetInch(cMip.yBotMargin)
        .Add "DataOnly", CBool(cMip.fDataOnly)
        .Add "Width", GetInch(cMip.xWidth)
        .Add "Height", GetInch(cMip.yHeight)
        .Add "DefaultSize", CBool(cMip.fDefaultSize)
        .Add "Columns", cMip.cxColumns
        .Add "ColumnSpacing", GetInch(cMip.yColumnSpacing)
        .Add "RowSpacing", GetInch(cMip.xRowSpacing)
        .Add "ItemLayout", GetEnum(epeColumnLayout, cMip.rItemLayout)
        .Add "FastPrint", cMip.fFastPrint  ' Reserved
        .Add "Datasheet", cMip.fDatasheet  ' Reserved
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : DevNamesToDictionary
' Author    : Adam Waller
' Date      : 10/30/2020
' Purpose   : Return a dictionary object with the PrtDevNames values.
'---------------------------------------------------------------------------------------
'
Private Function DevNamesToDictionary() As Dictionary

    Dim cDN As tDevNames

    LSet cDN = m_tDevNames
    Set DevNamesToDictionary = New Dictionary
    
    With DevNamesToDictionary
        .Add "DriverName", NTrim(Mid$(StrConv(cDN.strData, vbUnicode), cDN.intDriverOffset - 7))
        .Add "DeviceName", NTrim(Mid$(StrConv(cDN.strData, vbUnicode), cDN.intDeviceOffset - 7))
        .Add "Port", NTrim(Mid$(StrConv(cDN.strData, vbUnicode), cDN.intOutputOffset - 7))
        .Add "Default", (cDN.intDefault = 1)
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetMargins
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Sets the printer margins based on dictionary values, as read from
'           : source file or MIP structure. Thankfully we can set all of these
'           : using the Access object model.  :-)
'           : Note that this does not use type checking to verify that values
'           : have not been messed up. If you put a string value in a margin property
'           : for example, it will throw an error.
'           : Reference: http://etutorials.org/Microsoft+Products/access/Chapter+5.+Printers/Recipe+5.3+Programmatically+Change+Margin+and+Column+Settings+for+Reports/
'---------------------------------------------------------------------------------------
'
Public Sub SetPrinterMargins(oPrinter As Access.Printer, dMargins As Dictionary)

    Dim varKey As Variant
    
    ' Loop through properties.
    With oPrinter
        For Each varKey In dMargins.Keys
            Select Case varKey
            
                ' Set margins from dictionary values
                Case "LeftMargin": .LeftMargin = GetTwips(dMargins(varKey))
                Case "TopMargin": .TopMargin = GetTwips(dMargins(varKey))
                Case "RightMargin": .RightMargin = GetTwips(dMargins(varKey))
                Case "BotMargin": .BottomMargin = GetTwips(dMargins(varKey))
                Case "DataOnly": .DataOnly = dMargins(varKey)
                Case "Columns": .ItemsAcross = dMargins(varKey)
                Case "ColumnSpacing": .ColumnSpacing = GetTwips(dMargins(varKey))
                Case "RowSpacing": .RowSpacing = GetTwips(dMargins(varKey))
                Case "ItemLayout": .ItemLayout = GetEnum(epeColumnLayout, dMargins(varKey))
                
                ' Special handling for paper size
                Case "DefaultSize": .DefaultSize = dMargins(varKey)
                Case "Width":
                    If .ItemSizeWidth <> GetTwips(dMargins(varKey)) Then
                        If .DefaultSize Then .DefaultSize = False
                        .ItemSizeWidth = GetTwips(dMargins(varKey))
                    End If
                Case "Height":
                    If .ItemSizeHeight <> GetTwips(dMargins(varKey)) Then
                        If .DefaultSize Then .DefaultSize = False
                        .ItemSizeHeight = GetTwips(dMargins(varKey))
                    End If
            
                Case Else
                    ' Could not find that property.
                    MsgBox "Margin property " & CStr(varKey) & " not found.", vbExclamation
            End Select
        Next varKey
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetPrinterOptions
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Applies the dictionary of options to a printer. For the properties that
'           : can be set through the object model, they are applied that way. For those
'           : that can only be set through the DevMode API structure, they are applied
'           : through updating the DevMode.
'           : Reference: http://etutorials.org/Microsoft+Products/access/Chapter+5.+Printers/Recipe+5.4+Programmatically+Change+Printer+Options/
'---------------------------------------------------------------------------------------
'
Public Sub SetPrinterOptions(objFormOrReport As Object, dSettings As Dictionary)

    Dim oPrinter As Access.Printer
    Dim intType As Integer
    Dim intCnt As Integer
    Dim strForm As String
    Dim bteForm() As Byte
    Dim varKey As Variant
    Dim blnSetDevMode As Boolean
    Dim strDevModeExtra As String
    Dim tBuffer As tDevModeBuffer
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure we are using the correct object type
    If TypeOf objFormOrReport Is Access.Report Then
        intType = acReport
    ElseIf TypeOf objFormOrReport Is Access.Form Then
        intType = acForm
    Else
        Log.Error eelCritical, "Can only set printer options for a form or report object: " & _
            objFormOrReport.Name, ModuleName & ".SetPrinterOptions"
        Exit Sub
    End If
    
    ' Check printer device to see if we are using a specific printer
    If dSettings.Exists("DeviceName") Then
        Set oPrinter = GetPrinterByName(dSettings("DeviceName"))
        If oPrinter Is Nothing Then
            Log.Error eelWarning, "Printer " & dSettings("DeviceName") & " not found for " & objFormOrReport.Name, _
                ModuleName & ":SetPrinterOptions"
            Exit Sub
        End If
        ' Set as printer for this report or form.
        With objFormOrReport
            Set .Printer = oPrinter
            .UseDefaultPrinter = False
        End With
    Else
        ' Use default printer (If not already set)
        objFormOrReport.UseDefaultPrinter = True
    End If
    
    ' Apply regular printer options
    Set oPrinter = objFormOrReport.Printer
    With oPrinter
        For Each varKey In dSettings.Keys
            Select Case varKey
                Case "Orientation": .Orientation = GetEnum(epeOrientation, dSettings(varKey))
                Case "PaperSize": .PaperSize = GetEnum(epePaperSize, dSettings(varKey))
                Case "Copies": .Copies = dSettings(varKey)
                Case "PrintQuality": .PrintQuality = GetEnum(epePrintQuality, dSettings(varKey))
                Case "Color": .ColorMode = GetEnum(epeColor, dSettings(varKey))
                Case "Duplex": .Duplex = GetEnum(epeDuplex, dSettings(varKey))
                Case "DefaultSource": .PaperBin = GetEnum(epePaperBin, dSettings(varKey))
            End Select
        Next varKey
    End With
    
    ' Other properties will require some more work, since we need to interact
    ' with the DevMode structure.
    LoadFromObject objFormOrReport
    strDevModeExtra = objFormOrReport.PrtDevMode
    
    ' Loop through properties again, this time applying change to DevMode structure.
    With m_tDevMode
        For Each varKey In dSettings.Keys
            Select Case varKey
                Case "PaperLength": SetDmProp .intPaperLength, DM_PAPERLENGTH, Round(dSettings(varKey) / TEN_MIL, 0), .lngFields, blnSetDevMode
                Case "PaperWidth":  SetDmProp .intPaperWidth, DM_PAPERWIDTH, Round(dSettings(varKey) / TEN_MIL, 0), .lngFields, blnSetDevMode
                Case "Scale":       SetDmProp .intScale, DM_SCALE, dSettings(varKey), .lngFields, blnSetDevMode
                Case "Resolution":  SetDmProp .intResolution, DM_YRESOLUTION, dSettings(varKey), .lngFields, blnSetDevMode
                Case "TTOption":    SetDmProp .intTTOption, DM_TTOPTION, GetEnum(epeTTOption, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "Collate":     SetDmProp .intCollate, DM_COLLATE, GetEnum(epeCollate, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "DisplayFlags":        SetDmProp .lngDisplayFlags, DM_DISPLAYFLAGS, GetEnum(epeDisplayFlags, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "DisplayFrequency":    SetDmProp .lngDisplayFrequency, DM_DISPLAYFREQUENCY, dSettings(varKey), .lngFields, blnSetDevMode
                Case "ICMMethod":   SetDmProp .lngICMMethod, DM_ICMMETHOD, GetEnum(epeICMMethod, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "ICMIntent":   SetDmProp .lngICMIntent, DM_ICMINTENT, GetEnum(epeICMIntent, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "MediaType":   SetDmProp .lngMediaType, DM_MEDIATYPE, GetEnum(epeMediaType, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "DitherType":  SetDmProp .lngDitherType, DM_DITHERTYPE, GetEnum(epeDitherType, dSettings(varKey)), .lngFields, blnSetDevMode
                Case "FormName"
                    ' This one is a little more fun...
                    If (Not BitSet(.lngFields, DM_FORMNAME)) _
                        Or (dSettings(varKey) <> NTrim(StrConv(.strFormName, vbUnicode))) Then
                        ' Assign byte arrays for string values
                        strForm = StrConv(dSettings(varKey) & vbNullChar, vbFromUnicode)
                        bteForm = strForm & NullPad(32 - Len(strForm))
                        For intCnt = 1 To 32
                            .strFormName(intCnt) = bteForm(intCnt - 1)
                        Next intCnt
                        blnSetDevMode = True
                        ' Update fields flag
                        If Not BitSet(.lngFields, DM_FORMNAME) Then
                            .lngFields = .lngFields Or DM_FORMNAME
                        End If
                    End If
            End Select
        Next varKey
    End With
    
    ' Check flag to see if we have changed anything
    If blnSetDevMode Then
        tBuffer.strBuffer = Replace(tBuffer.strBuffer, " ", vbNullChar)
        LSet tBuffer = m_tDevMode
        ' Overwrite first part of structure while preserving possible
        ' extra data used by print driver.
        Mid(strDevModeExtra, 1, 94) = tBuffer.strBuffer
        objFormOrReport.PrtDevMode = strDevModeExtra
    End If
    
    ' Tweak a property so the report knows it needs to be saved.
    With objFormOrReport
        .Caption = .Caption
    End With
    CatchAny eelError, "Error setting print settings for: " & objFormOrReport.Name, _
        ModuleName & ".SetPrinterOptions", True, True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplySettings
' Author    : Adam Waller
' Date      : 10/28/2020
' Purpose   : Applies the dictionary object of printer settings to the current
'           : DevMode, MIP and DevNames structures. Expects a dictionary structure
'           : like the "items" collection that is created when saving print settings.
'---------------------------------------------------------------------------------------
'
Public Sub ApplySettings(dSettings As Dictionary)

    Dim intCnt As Integer
    Dim strForm As String
    Dim bteForm() As Byte
    Dim varKey As Variant
    Dim blnSetDevMode As Boolean
    Dim dItems As Dictionary
    Dim strPrinter As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Set the properties in the DevNames structure.
    ' Note that this simply sets the printer to one with a matching name. It doesn't try to reconstruct
    ' an identical share and port name or install a missing printer.
    strPrinter = dNZ(dSettings, "Device\DeviceName")
    If strPrinter = vbNullString Then
        ' Use default printer
        LoadFromDefaultPrinter
        ' Clear the device name, since we are not binding this
        ' form/report to a specific printer.
        For intCnt = 1 To 32
            m_tDevMode.strDeviceName(intCnt) = 0
        Next intCnt
    Else
        ' Load defaults from specific printer
        LoadFromPrinter strPrinter
    End If

    ' Set the properties in the DevMode structure.
    If IsObject(dSettings("Printer")) Then
        With m_tDevMode
            Set dItems = dSettings("Printer")
            For Each varKey In dItems.Keys
                Select Case varKey
                    ' Note that any specified DeviceName in m_tDevMode would have already been set through
                    ' the intial call that loaded the DevMode structure directly from the printer using the Windows API.
                
                    ' These properties can be set on the report/form object, or through PrtDevMode
                    Case "Orientation": SetDmProp .intOrientation, DM_ORIENTATION, GetEnum(epeOrientation, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "PaperSize":   SetDmProp .intPaperSize, DM_PAPERSIZE, GetEnum(epePaperSize, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "Copies":      SetDmProp .intCopies, DM_COPIES, dItems(varKey), .lngFields, blnSetDevMode
                    Case "PrintQuality":    SetDmProp .intPrintQuality, DM_PRINTQUALITY, GetEnum(epePrintQuality, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "Color":       SetDmProp .intColor, DM_COLOR, GetEnum(epeColor, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "Duplex":      SetDmProp .intDuplex, DM_DUPLEX, GetEnum(epeDuplex, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "DefaultSource":   SetDmProp .intDefaultSource, DM_DEFAULTSOURCE, GetEnum(epePaperBin, dItems(varKey)), .lngFields, blnSetDevMode
                
                    ' These can only be set through PrtDevMode
                    Case "PaperLength": SetDmProp .intPaperLength, DM_PAPERLENGTH, Round(dItems(varKey) / TEN_MIL, 0), .lngFields, blnSetDevMode
                    Case "PaperWidth":  SetDmProp .intPaperWidth, DM_PAPERWIDTH, Round(dItems(varKey) / TEN_MIL, 0), .lngFields, blnSetDevMode
                    Case "Scale":       SetDmProp .intScale, DM_SCALE, dItems(varKey), .lngFields, blnSetDevMode
                    Case "Resolution":  SetDmProp .intResolution, DM_YRESOLUTION, dItems(varKey), .lngFields, blnSetDevMode
                    Case "TTOption":    SetDmProp .intTTOption, DM_TTOPTION, GetEnum(epeTTOption, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "Collate":     SetDmProp .intCollate, DM_COLLATE, GetEnum(epeCollate, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "DisplayFlags":        SetDmProp .lngDisplayFlags, DM_DISPLAYFLAGS, GetEnum(epeDisplayFlags, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "DisplayFrequency":    SetDmProp .lngDisplayFrequency, DM_DISPLAYFREQUENCY, dItems(varKey), .lngFields, blnSetDevMode
                    Case "ICMMethod":   SetDmProp .lngICMMethod, DM_ICMMETHOD, GetEnum(epeICMMethod, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "ICMIntent":   SetDmProp .lngICMIntent, DM_ICMINTENT, GetEnum(epeICMIntent, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "MediaType":   SetDmProp .lngMediaType, DM_MEDIATYPE, GetEnum(epeMediaType, dItems(varKey)), .lngFields, blnSetDevMode
                    Case "DitherType":  SetDmProp .lngDitherType, DM_DITHERTYPE, GetEnum(epeDitherType, dItems(varKey)), .lngFields, blnSetDevMode
                    
                    ' String values are a little more fun...
                    Case "FormName"
                        If (Not BitSet(.lngFields, DM_FORMNAME)) _
                            Or (dItems(varKey) <> NTrim(StrConv(.strFormName, vbUnicode))) Then
                            ' Assign byte arrays for string values
                            strForm = StrConv(dItems(varKey) & vbNullChar, vbFromUnicode)
                            bteForm = strForm & NullPad(32 - Len(strForm))
                            For intCnt = 1 To 32
                                .strFormName(intCnt) = bteForm(intCnt - 1)
                            Next intCnt
                            blnSetDevMode = True
                            ' Update fields flag
                            If Not BitSet(.lngFields, DM_FORMNAME) Then
                                .lngFields = .lngFields Or DM_FORMNAME
                            End If
                        End If
                End Select
            Next varKey
        End With
    End If
    
    ' Set the printer margins in the MIP structure
    If IsObject(dSettings("Margins")) Then
        With m_tMip
            Set dItems = dSettings("Margins")
            For Each varKey In dItems.Keys
                Select Case varKey
                
                    ' Set margins from dictionary values
                    Case "LeftMargin": .xLeftMargin = GetTwips(dItems(varKey))
                    Case "TopMargin": .yTopMargin = GetTwips(dItems(varKey))
                    Case "RightMargin": .xRightMargin = GetTwips(dItems(varKey))
                    Case "BotMargin": .yBotMargin = GetTwips(dItems(varKey))
                    Case "DataOnly": .fDataOnly = dItems(varKey)
                    Case "Columns": .cxColumns = dItems(varKey)
                    Case "ColumnSpacing": .yColumnSpacing = GetTwips(dItems(varKey))
                    Case "RowSpacing": .xRowSpacing = GetTwips(dItems(varKey))
                    Case "ItemLayout": .rItemLayout = GetEnum(epeColumnLayout, dItems(varKey))
                    Case "FastPrint": .fFastPrint = Abs(CBool(dItems(varKey))) 'These are quite likely unneded; they do not appear to have an effect on the file creation/export.
                    Case "Datasheet": .fDatasheet = Abs(CBool(dItems(varKey))) 'These are quite likely unneded; they do not appear to have an effect on the file creation/export.

                    ' Special handling for paper size
                    Case "DefaultSize": .fDefaultSize = Abs(dItems(varKey))
                    Case "Width":
                        If .xWidth <> GetTwips(dItems(varKey)) Then
                            If CBool(.fDefaultSize) Then .fDefaultSize = Abs(False)
                            .xWidth = GetTwips(dItems(varKey))
                        End If
                    Case "Height":
                        If .yHeight <> GetTwips(dItems(varKey)) Then
                            If CBool(.fDefaultSize) Then .fDefaultSize = Abs(False)
                            .yHeight = GetTwips(dItems(varKey))
                        End If
                
                    Case Else
                        ' Could not find that property.
                        Log.Error eelWarning, "Margin property " & CStr(varKey) & " not found.", _
                            ModuleName & ":ApplySettings"
                End Select
            Next varKey
        End With
    End If
    CatchAny eelError, "Error applying print settings for: " & strPrinter, _
        ModuleName & ".ApplySettings", True, True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetPrintSettingsFileName
' Author    : Adam Waller
' Date      : 1/14/2021
' Purpose   : Return the file name for the print vars json file.
'---------------------------------------------------------------------------------------
'
Public Function GetPrintSettingsFileName(cDbObject As IDbComponent) As String
    GetPrintSettingsFileName = cDbObject.BaseFolder & GetSafeFileName(cDbObject.Name) & ".json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddToExportFile
' Author    : Adam Waller
' Date      : 11/2/2020
' Purpose   : Creates a temporary file from the contents of strFile, inserting the
'           : DevMode, DevNames and MIP blocks into the file header. This prepares the
'           : file for import into the database using the loaded print settings.
'---------------------------------------------------------------------------------------
'
Public Function AddToExportFile(strFile As String) As String

    Dim strTempFile As String
    Dim strLine As String
    Dim varLines As Variant
    Dim strData As String
    Dim lngLine As Long
    Dim blnFound As Boolean
    Dim blnInBlock As Boolean
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Load data from export file
    strData = ReadFile(strFile)
    varLines = Split(strData, vbCrLf)

    ' Use concatenation class for performance reasons.
    With New clsConcat
        .AppendOnAdd = vbCrLf
        
        ' Loop through lines in file, searching for location to insert blocks.
        For lngLine = LBound(varLines) To UBound(varLines)
            
            ' Get single line
            strLine = varLines(lngLine)
            
            ' Check line contents till we reach the insertion point.
            If Not blnFound Then
                Select Case Trim$(strLine)
                    Case "PrtMip = Begin", "PrtDevMode = Begin", "PrtDevNames = Begin", _
                        "PrtDevModeW = Begin", "PrtDevNamesW = Begin"
                        ' If we find any of these blocks in the file, we should remove
                        ' them since they are being replaced with the inserted ones.
                        blnInBlock = True
                    Case "End"
                        ' End of a block section.
                        If Not blnInBlock Then .Add strLine
                        blnInBlock = False
                    Case "Begin"
                        'Verify indent level
                        If strLine <> "    Begin" Then
                            .Add strLine
                        Else
                            ' Insert our blocks before this line.
                            .Add GetPrtMipBlock
                            .Add GetPrtDevModeBlock
                            .Add GetPrtDevNamesBlock
                            .Add strLine
                            blnFound = True
                        End If
                    Case Else
                        ' Continue building file contents
                        .Add strLine
                End Select
            Else
                ' Already inserted block content.
                .Add strLine
            End If
        Next lngLine
    
        ' Write to new file
        strTempFile = GetTempFile
        WriteFile .GetStr, strTempFile
        
    End With

    ' Return path to temp file
    AddToExportFile = strTempFile
    CatchAny eelError, "Error adding to export file: " & strFile, _
        ModuleName & ".AddToExportFile", True, True
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetInch
' Author    : Adam Waller
' Date      : 6/2/2020
' Purpose   : Convert a twips value to inches, rounded to 4 decimal places.
'---------------------------------------------------------------------------------------
'
Private Function GetInch(lngTwips As Long) As Single
    GetInch = Round(lngTwips / 1440, 4)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTwips
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Return twips from inches
'---------------------------------------------------------------------------------------
'
Private Function GetTwips(sngInches As Single) As Long
    GetTwips = Round(sngInches * 1440, 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : BitSet
' Author    : Adam Waller
' Date      : 5/19/2020
' Purpose   : Returns true if the flag is set.
'---------------------------------------------------------------------------------------
'
Private Function BitSet(lngFlags As Long, lngValue As edmFlags) As Boolean
    BitSet = CBool((lngFlags And lngValue) = lngValue)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDmProp
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Set a DevMode property, including the fields flag, if the value has
'           : changed from the existing value. Sets the blnChanged flag to true if
'           : the property was set or changed.
'---------------------------------------------------------------------------------------
'
Private Sub SetDmProp(ByRef cDMProp As Variant, lngFlag As edmFlags, varValue As Variant, ByRef lngFields As Long, ByRef blnChanged As Boolean)
    
    ' Check existing flag
    If Not BitSet(lngFields, lngFlag) Then
        blnChanged = True
    Else
        ' Check existing value
        If cDMProp <> varValue Then
            ' Set property by name
            cDMProp = varValue
            blnChanged = True
        End If
    End If
    
    ' Check fields flag, and update flag if we have made a change.
    If blnChanged And Not BitSet(lngFields, lngFlag) Then lngFields = lngFields Or lngFlag
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDevNames
' Author    : Adam Waller
' Date      : 10/30/2020
' Purpose   : Wrapper to encode PrtDevNames values from passed printer object.
'---------------------------------------------------------------------------------------
'
Private Sub SetDevNames(objPrinter As Access.Printer)

    Dim strDriver As String
    Dim strDevice As String
    Dim strPort As String
    Dim strData As String
    Dim bteData() As Byte
    Dim intCnt As Integer
    Dim blnDefault As Boolean

    ' Get device properties
    With objPrinter
        ' Default printer is stored differently.
        blnDefault = (.DeviceName = Application.Printer.DeviceName)
        If blnDefault Then
            ' Determined the size by reviewing an exported report
            ' that uses the default printer and noting the offsets.
            strDriver = NullPad(23)
            strDevice = NullPad(23)
        Else
            strDriver = .DriverName & vbNullChar
            strDevice = Left$(.DeviceName & vbNullChar, 32)
        End If
        strPort = .Port & vbNullChar
    End With
    
    ' Fill in DevNames structure
    With m_tDevNames
        strData = strDriver & strDevice & strPort
        .intDriverOffset = 8  ' This seems to match what I typically see...
        .intDeviceOffset = .intDriverOffset + Len(strDriver)
        .intOutputOffset = .intDeviceOffset + Len(strDevice)
        .intDefault = Abs(blnDefault)
        ' Convert string data to byte array
        strData = StrConv(strData, vbFromUnicode)
        bteData = strData & NullPad(255 - Len(strData))
        For intCnt = 1 To 255
            .strData(intCnt) = bteData(intCnt - 1)
        Next intCnt
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetMipFromPrinter
' Author    : Adam Waller
' Date      : 11/2/2020
' Purpose   : Sets the margins (binary) structure from the default values of a printer
'           : object. This is used when building the MIP blob section before importing
'           : a report from export files. This function primarily sets the defaults,
'           : then the actual report margins are loaded from the JSON file.
'---------------------------------------------------------------------------------------
'
Private Sub SetMipFromPrinter(objPrinter As Access.Printer)

    ' Set margins from printer object
    With objPrinter
        m_tMip.xLeftMargin = .LeftMargin
        m_tMip.yTopMargin = .TopMargin
        m_tMip.xRightMargin = .RightMargin
        m_tMip.yBotMargin = .BottomMargin
        m_tMip.fDataOnly = Abs(.DataOnly)
        m_tMip.cxColumns = .ItemsAcross
        m_tMip.yColumnSpacing = .ColumnSpacing
        m_tMip.xRowSpacing = .RowSpacing
        m_tMip.rItemLayout = .ItemLayout
        ' Paper size should just map across...
        m_tMip.fDefaultSize = Abs(.DefaultSize)
        m_tMip.xWidth = .ItemSizeWidth
        m_tMip.yHeight = .ItemSizeHeight
        ' Reserved properties
        ' (Maybe set these to default values?)
        'm_tMip.fFastPrint =
        'm_tmip.fDatasheet =
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetPrinterByName
' Author    : Adam Waller
' Date      : 10/22/2020
' Purpose   : Return a printer object matching a specific printer name. (Or nothing if
'           : the printer name is not found.)
'---------------------------------------------------------------------------------------
'
Private Function GetPrinterByName(strName As String) As Access.Printer
    Dim prt As Access.Printer
    For Each prt In Access.Printers
        If prt.DeviceName = strName Then
            Set GetPrinterByName = prt
            Exit For
        End If
    Next prt
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetEnum
' Author    : Adam Waller
' Date      : 5/15/2020
' Purpose   : Return an enum value, 0 or UNKNOWN if not found.
'---------------------------------------------------------------------------------------
'
Public Function GetEnum(eType As ePrintEnum, varValue As Variant, Optional Convert As eEnumConversion = eecAuto) As Variant

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
Private Sub BuildEnum(eType As ePrintEnum)

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
                
            Case epeColumnLayout
                .Add acPRHorizontalColumnLayout, "Horizontal Columns"
                .Add acPRVerticalColumnLayout, "Vertical Columns"
                
            '--------------------
            ' API constants
            '--------------------
            
            Case epeTTOption
                .Add 1, "Bitmap"
                .Add 2, "Download"
                .Add 3, "Substitute Device Font"
                .Add 4, "Download Outline"
            
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

End Sub


'---------------------------------------------------------------------------------------
' Procedure : NTrim
' Author    : Adam Waller
' Date      : 5/18/2020
' Purpose   : Trim a string to a null character terminator.
'---------------------------------------------------------------------------------------
'
Private Function NTrim(ByVal strText As String) As String
    Dim lngPos As Long
    lngPos = InStr(1, strText, vbNullChar)
    If lngPos > 0 Then
        NTrim = Left$(strText, lngPos - 1)
    Else
        NTrim = strText
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetHash
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Returns a hash of the devmode structure *without* the two reserved
'           : settings of FastPrint and Datasheet. This is useful in comparing to
'           : see if the print settings match the default printer settings.
'---------------------------------------------------------------------------------------
'
Public Function GetHash()

    Dim dSettings As Dictionary
    
    Set dSettings = Me.GetDictionary
    
    ' Remove the reserved settings, which are Access specific and may not match the
    ' settings retrieved from the default printer.
    If dSettings.Exists("Margins") Then
        With dSettings("Margins")
            If .Exists("FastPrint") Then .Remove "FastPrint"
            If .Exists("Datasheet") Then .Remove "Datasheet"
        End With
    End If
    
    'Debug.Print ConvertToJson(dSettings, "  ")
    GetHash = GetDictionaryHash(dSettings)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPrtDevModeBlock
' Author    : Adam Waller
' Date      : 10/27/2020
' Purpose   : Return a formatted PrtDevMode block.
'---------------------------------------------------------------------------------------
'
Public Function GetPrtDevModeBlock() As String
    Dim udtBuffer As tDevModeBuffer
    udtBuffer.strBuffer = Replace(udtBuffer.strBuffer, " ", vbNullChar)
    LSet udtBuffer = m_tDevMode
    GetPrtDevModeBlock = GetBlobFromString("PrtDevMode", udtBuffer.strBuffer)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPrtMipBlock
' Author    : Adam Waller
' Date      : 10/27/2020
' Purpose   : Return a formatted PrtMip block.
'---------------------------------------------------------------------------------------
'
Public Function GetPrtMipBlock() As String
    Dim udtBuffer As tMipBuffer
    udtBuffer.strBuffer = Replace(udtBuffer.strBuffer, " ", vbNullChar)
    LSet udtBuffer = m_tMip
    GetPrtMipBlock = GetBlobFromString("PrtMip", udtBuffer.strBuffer)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPrtDevNamesBlock
' Author    : Adam Waller
' Date      : 10/27/2020
' Purpose   : Return a formatted PrtDevNames block.
'---------------------------------------------------------------------------------------
'
Public Function GetPrtDevNamesBlock() As String
    Dim udtBuffer As tDevNamesBuffer
    udtBuffer.strBuffer = Replace(udtBuffer.strBuffer, " ", vbNullChar)
    LSet udtBuffer = m_tDevNames
    GetPrtDevNamesBlock = GetBlobFromString("PrtDevNames", RTrimNulls(udtBuffer.strBuffer, 1))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBlobFromString
' Author    : Adam Waller
' Date      : 10/27/2020
' Purpose   : Convert a string to a hexidecimal binary representation used in Access
'           : report sections like PrtDevMode.
'---------------------------------------------------------------------------------------
'
Private Function GetBlobFromString(strSection As String, strContent As String, Optional intIndent As Integer = 4) As String

    Dim intCol As Integer
    Dim lngPos As Long
    Dim bteContent() As Byte
    Dim lngLen As Long
    Dim strBte As String
    
    ' Use concatenation class to drastically improve string handling performance.
    With New clsConcat

        ' Start with section beginning
        .Add Space$(intIndent), strSection, " = Begin", vbCrLf
        
        ' Convert string to byte array
        bteContent = strContent
        lngLen = Len(strContent) * 2
        
        ' Build content lines
        Do While lngPos < lngLen
            .Add Space$(intIndent + 4), "0x"
            For intCol = 0 To 31
                If lngPos + intCol >= lngLen Then
                    lngPos = lngLen
                    Exit For
                Else
                    strBte = LCase(Hex$(bteContent(lngPos + intCol)))
                    If Len(strBte) = 1 Then .Add "0"
                    .Add strBte
                End If
            Next intCol
            lngPos = lngPos + 32
            If lngPos < lngLen Then .Add " ,", vbCrLf
        Loop

        ' Add section closing
        .Add vbCrLf
        .Add Space$(intIndent), "End"
        
        ' Return blob string
        GetBlobFromString = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : NullPad
' Author    : Adam Waller
' Date      : 10/30/2020
' Purpose   : Like the Space$() function, but uses vbNullChar instead.
'---------------------------------------------------------------------------------------
'
Private Function NullPad(lngNumber As Long) As String
    NullPad = Replace$(Space$(lngNumber), " ", vbNullChar)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RTrimNulls
' Author    : Adam Waller
' Date      : 10/30/2020
' Purpose   : Trims the null characters off the right end of a string, leaving the
'           : specified null characters at the end. (Used for variable length structures)
'---------------------------------------------------------------------------------------
'
Private Function RTrimNulls(strData As String, lngLeaveCount As Long) As String

    Dim lngCnt As Long
    Dim strTrimmed As String
    
    ' Walk backwards through the string, looking for the first non-null character.
    If InStr(1, strData, vbNullChar) > 0 Then
        For lngCnt = Len(strData) To 1 Step -1
            If Mid$(strData, lngCnt, 1) <> vbNullChar Then
                ' Found a non-null character.
                strTrimmed = Left$(strData, lngCnt) & NullPad(lngLeaveCount)
                Exit For
            End If
        Next lngCnt
    End If
    
    ' Return result
    If strTrimmed = vbNullString Then
        ' If no nulls found, then return original data.
        RTrimNulls = strData
    Else
        RTrimNulls = strTrimmed
    End If
    
End Function