Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_PrinterSupport"
'' 8/15/2014: KD Need to make sure this is imported to the Master Claim Admin

''' ######################### for the common dialog

Public Const GWL_HINSTANCE = (-6)
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const HCBT_ACTIVATE = 5
Public Const WH_CBT = 5
Public Const LF_FACESIZE = 32


Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const PD_PRINTSETUP = &H40
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113
Private Const OPAQUE = 0
Private Const TRANSPARENT = 1


Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Public Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Public Type CHOOSECOLORS
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hWndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type





Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256



'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY



Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

    'Public Type PRINTDLGS
Public Type PRINTDLG_TYPE
        lStructSize As Long
        hWndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type


Private Type DEVNAMES_TYPE
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 200
End Type

Private Type DEVMODE_TYPE
  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
  dmICMMethod As Long
  dmICMIntent As Long
  dmMediaType As Long
  dmDitherType As Long
  dmReserved1 As Long
  dmReserved2 As Long
  dmPanningWidth As Long
  dmPanningHeight As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
'Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
'Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Public Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

Public lParenthWnd As Long
Public hHook As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long




Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long


'' from here -= newly added
'Private Const CCHDEVICENAME = 32
'Private Const CCHFORMNAME = 32
'Private Const GMEM_MOVEABLE = &H2
'Private Const GMEM_ZEROINIT = &H40
'Private Const DM_DUPLEX = &H1000&
'Private Const DM_ORIENTATION = &H1&
'Private Const PD_PRINTSETUP = &H40
'Private Const PD_DISABLEPRINTTOFILE = &H80000
'Private Const PHYSICALOFFSETX As Long = 112
'Private Const PHYSICALOFFSETY As Long = 113
'Private Const OPAQUE = 0
'Private Const TRANSPARENT = 1

'Public Type JOB_INFO_2
'   JobId As Long
'   pPrinterName As Long
'   pMachineName As Long
'   pUserName As Long
'   pDocument As Long
'   pNotifyName As Long
'   pDatatype As Long
'   pPrintProcessor As Long
'   pParameters As Long
'   pDriverName As Long
'   pDevMode As Long
'   pStatus As Long
'   pSecurityDescriptor As Long
'   Status As Long
'   Priority As Long
'   Position As Long
'   StartTime As Long
'   UntilTime As Long
'   TotalPages As Long
'   Size As Long
'   Submitted As SYSTEMTIME
'   Time As Long
'   PagesPrinted As Long
'End Type

Public Type ACL
  AclRevision As Byte
  Sbz1 As Byte
  AclSize As Integer
  AceCount As Integer
  Sbz2 As Integer
End Type

Public Type SECURITY_DESCRIPTOR
  Revision As Byte
  Sbz1 As Byte
  Control As Long
  Owner As Long
  Group As Long
  Sacl As ACL
  Dacl As ACL
End Type

Public Type JOB_INFO_2
  JobId As Long
  pPrinterName As String
  pMachineName As String
  pUserName As String
  pDocument As String
  pNotifyName As String
  pDatatype As String
  pPrintProcessor As String
  pParameters As String
  pDriverName As String
  pDevMode As DevMode
  pStatus As String
  pSecurityDescriptor As SECURITY_DESCRIPTOR
  Status As Long
  Priority As Long
  position As Long
  StartTime As Long
  UntilTime As Long
  TotalPages As Long
  Size As Long
  Submitted As SYSTEMTIME
  Time As Long
  PagesPrinted As Long
End Type

Public Type PRINTER_INFO_1
  flags As Long
  pDescription As String
  pName As String
  pComment As String
End Type

Public Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_WARMING_UP = &H10000
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_STATUS_DELETED = &H100
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
Public Const JOB_STATUS_USER_INTERVENTION = &H400
Public Const JOB_STATUS_RESTART = &H800


Public Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Dim rectForm As RECT, rectMsg As RECT
'Dim x As Long, y As Long
'    If lMsg = HCBT_ACTIVATE Then
'        'Show the MsgBox at a fixed location (0,0)
'        GetWindowRect wParam, rectMsg
'
'
'        x = screen.Width / screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.left) / 2
'        y = screen.Height / screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.top) / 2
'        Debug.Print "Screen " & screen.Height / 2
'        Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.left) / 2
'        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
'        'Release the CBT hook
'        UnhookWindowsHookEx hHook
'    End If
'    WinProcCenterScreen = False
End Function

Public Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim rectForm As RECT, rectMsg As RECT
Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect lParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.left + (rectForm.Right - rectForm.left) / 2) - ((rectMsg.Right - rectMsg.left) / 2)
        Y = (rectForm.top + (rectForm.Bottom - rectForm.top) / 2) - ((rectMsg.Bottom - rectMsg.top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function


Public Function SelectPrinter(Optional PrintFlags As Long = PD_PRINTSETUP, _
    Optional sOrigPrinterName As String, Optional bSetVBPrinterObject As Boolean = False) As String
On Error GoTo Block_Err
Dim strProcName As String

Dim sLongPrinterName As String
Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE
Dim lpDevMode As Long, lpDevName As Long
Dim bReturn As Integer
Dim p1 As Printer, sNewPrinterName As String
Dim hWndOwner As Long

    strProcName = ClassName & ".SelectPrinter"

    hWndOwner = Application.hWndAccessApp

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = hWndOwner
    PrintDlg.flags = PrintFlags
    
    
    sOrigPrinterName = DefaultPrinterInfo()
    Debug.Print sOrigPrinterName
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 2)

'''    lRet = SetDefaultPrinter(sOrigPrinterName)
        
    
'    sOrigPrinterName = Printer.DeviceName
    If Len(sOrigPrinterName) > 0 Then
        If bSetVBPrinterObject = True Then  '' KD
            For Each p1 In Printers
                If InStr(1, p1.DeviceName, sOrigPrinterName, vbTextCompare) > 0 Then
                    Set Printer = p1
                    Exit For
                End If
            Next
        End If
    End If
    
    
    ' set up for the dialog box
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION
'    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    
    On Error GoTo Block_Err
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If
    
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    
    ' show the dialog box and wait:
    If PrintDialog(PrintDlg) <> 0 Then
        '
        ' Mike's amendment to handle long printer names
        CopyMemory DevName, ByVal lpDevName, Len(DevName)
        sLongPrinterName = Mid$(DevName.extra, DevName.wDeviceOffset - DevName.wDriverOffset + 1)
        sLongPrinterName = left$(sLongPrinterName, InStr(sLongPrinterName, Chr$(0)) - 1)
        
        DoEvents ' allow dialog to remove itself from display
        

        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames
        
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        
        sNewPrinterName = UCase$(left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        
        sNewPrinterName = UCase$(Mid(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        
        ' Code now handles long printer names properly
        Dim sTemp As String
'        sTemp = DevName.extra
        
     
        SelectPrinter = sLongPrinterName

        If bSetVBPrinterObject = True Then  '' KD
            If Printer.DeviceName <> sLongPrinterName Then
                For Each p1 In Printers
                    If p1.DeviceName = sLongPrinterName Then
                        Set Printer = p1
                    End If
                Next
            End If
            
            ' Transfer settings from the Devmode structure to the
            ' VB printer object (this example just transfers some
            ' of them but you can of course use transfer more)
            Printer.Copies = DevMode.dmCopies
            Printer.Duplex = DevMode.dmDuplex
            Printer.Orientation = DevMode.dmOrientation
            Printer.PaperSize = DevMode.dmPaperSize
            Printer.PrintQuality = DevMode.dmPrintQuality
            Printer.ColorMode = DevMode.dmColor
            Printer.PaperBin = DevMode.dmDefaultSource
            
                    '        Printer.TopMargin
                    '        Printer.BottomMargin
                    '        Printer.LeftMargin
                    '        Printer.RightMargin
                                    
                    '        Printer.ColumnSpacing
                    '
                    '        Printer.ItemLayout
                    '        Printer.ItemsAcross
                    '        Printer.ItemSizeHeight
                    '        Printer.ItemSizeWidth
                    '
                    '        Printer.RowSpacing
                    '
                    '        Printer.DataOnly
                    '        Printer.DefaultSize
                    '      SetBkMode Printer.hdc, TRANSPARENT

        End If
        
        
    Else
        SelectPrinter = "" ' user cancelled
        For Each p1 In Printers
            If p1.DeviceName = sOrigPrinterName Then
                Set Printer = p1
                Exit For
            End If
        Next
        GlobalFree PrintDlg.hDevNames
        GlobalFree PrintDlg.hDevMode
    End If
    

    
Block_Exit:
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function TestPrinterStuff()
Dim sOrigPrinterName As String


    
    Debug.Print SelectPrinter(, sOrigPrinterName, False)
    
    Debug.Print "Orig was: " & sOrigPrinterName
    
'    lRet = SetDefaultPrinter(sOrigPrinterName)

End Function