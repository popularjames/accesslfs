Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, _
'    phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function GetPrinterApi Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, _
         ByVal Level As Long, buffer As Long, ByVal pbSize As Long, pbSizeNeeded As Long) As Long

Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, _
        ByVal Level As Long, pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long

Private Declare Function FindFirstPrinterChangeNotificationLong Lib "winspool.drv" Alias "FindFirstPrinterChangeNotification" _
  (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal lpPrinterNotifyOptions As Long) As Long

Private Declare Function FindNextPrinterChangeNotificationByLong Lib "winspool.drv" Alias "FindNextPrinterChangeNotification" _
    (ByVal hChange As Long, pdwChange As Long, pPrinterOptions As PRINTER_NOTIFY_OPTIONS, ppPrinterNotifyInfo As Long) As Long

Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE = &HFFFF ' Infinite timeout

Private Declare Function FindClosePrinterChangeNotification Lib "winspool.drv" (ByVal hChange As Long) As Long

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

Private Type PRINTER_DEFAULTS
  pDatatype As String
  pDevMode As DevMode
  DesiredAccess As Long
End Type

'\\ Declarations
Private Type PRINTER_NOTIFY_INFO_DATA
  Type As Integer
  Field As Integer
  Reserved As Long
  Id As Long
  adwData(0 To 1) As Long
End Type

Private Type PRINTER_NOTIFY_INFO
  dwVersion As Long
  dwFlags As Long
  dwCount As Long
End Type

Public Enum Printer_Status
   PRINTER_STATUS_READY = &H0
   PRINTER_STATUS_PAUSED = &H1
   PRINTER_STATUS_ERROR = &H2
   PRINTER_STATUS_PENDING_DELETION = &H4
   PRINTER_STATUS_PAPER_JAM = &H8
   PRINTER_STATUS_PAPER_OUT = &H10
   PRINTER_STATUS_MANUAL_FEED = &H20
   PRINTER_STATUS_PAPER_PROBLEM = &H40
   PRINTER_STATUS_OFFLINE = &H80
   PRINTER_STATUS_IO_ACTIVE = &H100
   PRINTER_STATUS_BUSY = &H200
   PRINTER_STATUS_PRINTING = &H400
   PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
   PRINTER_STATUS_NOT_AVAILABLE = &H1000
   PRINTER_STATUS_WAITING = &H2000
   PRINTER_STATUS_PROCESSING = &H4000
   PRINTER_STATUS_INITIALIZING = &H8000
   PRINTER_STATUS_WARMING_UP = &H10000
   PRINTER_STATUS_TONER_LOW = &H20000
   PRINTER_STATUS_NO_TONER = &H40000
   PRINTER_STATUS_PAGE_PUNT = &H80000
   PRINTER_STATUS_USER_INTERVENTION = &H100000
   PRINTER_STATUS_OUT_OF_MEMORY = &H200000
   PRINTER_STATUS_DOOR_OPEN = &H400000
   PRINTER_STATUS_SERVER_UNKNOWN = &H800000
   PRINTER_STATUS_POWER_SAVE = &H1000000
End Enum

Private Type PRINTER_INFO_2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   JobsCount As Long
   AveragePPM As Long
End Type



Private Declare Function FreePrinterNotifyInfoByLong Lib "winspool.drv" Alias "FreePrinterNotifyInfo" (ByVal pInfo As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub CopyMemoryPRINTER_NOTIFY_INFO Lib "kernel32" Alias "RtlMoveMemory" (Destination As PRINTER_NOTIFY_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryPRINTER_NOTIFY_INFO_DATA Lib "kernel32" Alias "RtlMoveMemory" (Destination As PRINTER_NOTIFY_INFO_DATA, ByVal Source As Long, ByVal Length As Long)
Private aData() As PRINTER_NOTIFY_INFO_DATA



'''
'''Public Enum Printer_Change_Notification_General_Flags
'''    PRINTER_CHANGE_FORM = &H70000
'''    PRINTER_CHANGE_PORT = &H700000
'''    PRINTER_CHANGE_JOB = &HFF00
'''    PRINTER_CHANGE_PRINTER = &HFF
'''    PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
'''    PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Form_Flags
'''    PRINTER_CHANGE_ADD_FORM = &H10000
'''    PRINTER_CHANGE_SET_FORM = &H20000
'''    PRINTER_CHANGE_DELETE_FORM = &H40000
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Port_Flags
'''    PRINTER_CHANGE_ADD_PORT = &H100000
'''    PRINTER_CHANGE_CONFIGURE_PORT = &H200000
'''    PRINTER_CHANGE_DELETE_PORT = &H400000
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Job_Flags
'''    PRINTER_CHANGE_ADD_JOB = &H100
'''    PRINTER_CHANGE_SET_JOB = &H200
'''    PRINTER_CHANGE_DELETE_JOB = &H400
'''    PRINTER_CHANGE_WRITE_JOB = &H800
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Printer_Flags
'''    PRINTER_CHANGE_ADD_PRINTER = &H1
'''    PRINTER_CHANGE_SET_PRINTER = &H2
'''    PRINTER_CHANGE_DELETE_PRINTER = &H4
'''    PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Processor_Flags
'''    PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
'''    PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
'''End Enum
'''
'''Public Enum Printer_Change_Notification_Driver_Flags
'''    PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
'''    PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
'''    PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
'''End Enum
'''
'''
Public Enum Job_Notify_Field_Indexes
    JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
    JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
    JOB_NOTIFY_FIELD_PORT_NAME = &H2
    JOB_NOTIFY_FIELD_USER_NAME = &H3
    JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
    JOB_NOTIFY_FIELD_DATATYPE = &H5
    JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
    JOB_NOTIFY_FIELD_PARAMETERS = &H7
    JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
    JOB_NOTIFY_FIELD_DEVMODE = &H9
    JOB_NOTIFY_FIELD_STATUS = &HA
    JOB_NOTIFY_FIELD_STATUS_STRING = &HB
    JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
    JOB_NOTIFY_FIELD_DOCUMENT = &HD
    JOB_NOTIFY_FIELD_PRIORITY = &HE
    JOB_NOTIFY_FIELD_POSITION = &HF
    JOB_NOTIFY_FIELD_SUBMITTED = &H10
    JOB_NOTIFY_FIELD_START_TIME = &H11
    JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
    JOB_NOTIFY_FIELD_TIME = &H13
    JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
    JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
    JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
    JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
End Enum


''' ######################### for the common dialog

''''**************************************
''''Windows API/Global Declarations for :Sh
''''     ow Printer Document Properties setup dia
''''     log
''''**************************************
''''SHSTDAPI_(BOOL) SHInvokePrinterCommandA
''''     (HWND hwnd, UINT uAction, LPCSTR lpBuf1,
''''     LPCSTR lpBuf2, BOOL fModal);
'''
'''
'''Private Declare Function SHInvokePrinterCommand Lib "shell32.dll" Alias "SHInvokePrinterCommandA" (ByVal hWnd As Long, ByVal uAction As enPrinterActions, ByVal Buffer1 As String, ByVal Buffer2 As String, ByVal Modal As Long) As Long
'''
'''
'''Public Enum enPrinterActions
'''    PRINTACTION_OPEN = 0
'''    PRINTACTION_PROPERTIES = 1
'''    PRINTACTION_NETINSTALL = 2
'''    PRINTACTION_NETINSTALLLINK = 3
'''    PRINTACTION_TESTPAGE = 4
'''    PRINTACTION_OPENNETPRN = 5
'''    PRINTACTION_DOCUMENTDEFAULTS = 6
'''    PRINTACTION_SERVERPROPERTIES = 7
'''End Enum
''''**************************************
'''' Name: Show Printer Document Properties
''''     setup dialog
'''' Description:Shows the printer document
''''     properties dialog box from code.
'''' By: Duncan Jones
''''
''''
'''' Inputs:None
''''
'''' Returns:None
''''
''''Assumes:None
''''
''''Side Effects:This entry in Shell32.dll
''''     is only present in version 4.71 and abov
''''     e (Windows NT 4 and Internet Explorer 4.
''''     0 or above)
''''This code is copyrighted and has limite
''''     d warranties.
''''Please see http://www.Planet-Source-Cod
''''     e.com/xq/ASP/txtCodeId.22127/lngWId.1/qx
''''     /vb/scripts/ShowCode.htm
''''for details.
''''**************************************
'''
'''
'''
'''Public Sub DisplayDocumentDefaults(ByVal PrinterName As String, ByVal hWnd As Long)
'''    Dim lRet As Long
'''    '\\ Only version 4.71 and above have thi
'''    '     s :. jump over error
'''    On Error Resume Next
'''    lRet = SHInvokePrinterCommand(hWnd, PRINTACTION_DOCUMENTDEFAULTS, PrinterName, "", 0)
'''End Sub
'''

'
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
'Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
'Private Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

'
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'
'Private tFileDialog As OPENFILENAME
'Private tColorDialog As CHOOSECOLORS
'Private tFontDialog As CHOOSEFONTS
'Private tPrintDialog As PRINTDLGS



'''''' ###################### for inspecting the queue status and such:
'''Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'''
'''Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
'''
'''
'''Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'''
'''Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'''
Private Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" (ByVal hPrinter As Long, ByVal FirstJob As Long, ByVal NoJobs As Long, ByVal Level As Long, _
   pJob As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

' constants for PRINTER_DEFAULTS structure
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ACCESS_ADMINISTER = &H4
'''
'''' constants for DEVMODE structure
'''Private Const CCHDEVICENAME = 32
'''Private Const CCHFORMNAME = 32

'Private Type PRINTER_DEFAULTS
'   pDatatype As String
'   pDevMode As Long
'   DesiredAccess As Long
'End Type
'''
'''Private Type DEVMODE
'''   dmDeviceName As String * CCHDEVICENAME
'''   dmSpecVersion As Integer
'''   dmDriverVersion As Integer
'''   dmSize As Integer
'''   dmDriverExtra As Integer
'''   dmFields As Long
'''   dmOrientation As Integer
'''   dmPaperSize As Integer
'''   dmPaperLength As Integer
'''   dmPaperWidth As Integer
'''   dmScale As Integer
'''   dmCopies As Integer
'''   dmDefaultSource As Integer
'''   dmPrintQuality As Integer
'''   dmColor As Integer
'''   dmDuplex As Integer
'''   dmYResolution As Integer
'''   dmTTOption As Integer
'''   dmCollate As Integer
'''   dmFormName As String * CCHFORMNAME
'''   dmLogPixels As Integer
'''   dmBitsPerPel As Long
'''   dmPelsWidth As Long
'''   dmPelsHeight As Long
'''   dmDisplayFlags As Long
'''   dmDisplayFrequency As Long
'''End Type
'''
'''Private Type SYSTEMTIME
'''   wYear As Integer
'''   wMonth As Integer
'''   wDayOfWeek As Integer
'''   wDay As Integer
'''   wHour As Integer
'''   wMinute As Integer
'''   wSecond As Integer
'''   wMilliseconds As Integer
'''End Type
'''
'''
''''Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
''''Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'''Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
''''Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
''''Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
'''Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
'''''' ######################
'''''' ###################### Events
'''''' ######################
'''
'''
'''Private Const DM_IN_BUFFER = 8
'''Private Const DM_OUT_BUFFER = 2
'''Private Const DM_FORMNAME = &H10000
'''Private Const DM_COPIES = &H100&
'''Private Const DM_MODIFY = 8
'''Private Const DM_COPY = 2

Private Enum PrinterChangeNotifications
    PRINTER_CHANGE_ADD_PRINTER = &H1
    PRINTER_CHANGE_SET_PRINTER = &H2
    PRINTER_CHANGE_DELETE_PRINTER = &H4
    PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
    PRINTER_CHANGE_PRINTER = &HFF
    PRINTER_CHANGE_ADD_JOB = &H100
    PRINTER_CHANGE_SET_JOB = &H200
    PRINTER_CHANGE_DELETE_JOB = &H400
    PRINTER_CHANGE_WRITE_JOB = &H800
    PRINTER_CHANGE_JOB = &HFF00
    PRINTER_CHANGE_ADD_FORM = &H10000
    PRINTER_CHANGE_SET_FORM = &H20000
    PRINTER_CHANGE_DELETE_FORM = &H40000
    PRINTER_CHANGE_FORM = &H70000
    PRINTER_CHANGE_ADD_PORT = &H100000
    PRINTER_CHANGE_CONFIGURE_PORT = &H200000
    PRINTER_CHANGE_DELETE_PORT = &H400000
    PRINTER_CHANGE_PORT = &H700000
    PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
    PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
    PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
    PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
    PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
    PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
    PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
    PRINTER_CHANGE_TIMEOUT = &H80000000
End Enum

Private Enum JobChangeNotificationFields
    JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
    JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
    JOB_NOTIFY_FIELD_PORT_NAME = &H2
    JOB_NOTIFY_FIELD_USER_NAME = &H3
    JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
    JOB_NOTIFY_FIELD_DATATYPE = &H5
    JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
    JOB_NOTIFY_FIELD_PARAMETERS = &H7
    JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
    JOB_NOTIFY_FIELD_DEVMODE = &H9
    JOB_NOTIFY_FIELD_STATUS = &HA
    JOB_NOTIFY_FIELD_STATUS_STRING = &HB
    JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
    JOB_NOTIFY_FIELD_DOCUMENT = &HD
    JOB_NOTIFY_FIELD_PRIORITY = &HE
    JOB_NOTIFY_FIELD_POSITION = &HF
    JOB_NOTIFY_FIELD_SUBMITTED = &H10
    JOB_NOTIFY_FIELD_START_TIME = &H11
    JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
    JOB_NOTIFY_FIELD_TIME = &H13
    JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
    JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
    JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
    JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
End Enum

Private Type PRINTER_NOTIFY_OPTIONS
    Version As Long '\\should be set to 2
    flags As Long
    Count As Long
    lpPrintNotifyOptions As Long
End Type

Private Type PRINTER_NOTIFY_OPTIONS_TYPE
    Type As Integer
    Reserved_0 As Integer
    Reserved_1 As Long
    Reserved_2 As Long
    Count As Long
    pFields As Long
End Type

Private PrintOptions As PRINTER_NOTIFY_OPTIONS
Private PrinterNotifyOptions(0 To 1) As PRINTER_NOTIFY_OPTIONS_TYPE

Private mEventHandle As Long
Private mhPrinter As Long

Private FileDialog As OPENFILENAME
Private ColorDialog As CHOOSECOLORS
Private FontDialog As CHOOSEFONTS
Private PrintDialog As PRINTDLG_TYPE
Private ParenthWnd As Long


Public Event StatusChange(sNewStatus As String)


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property




''' ######################
''' ###################### Object properties
''' ######################
'''Private Function SelectPrinterByName(ByVal printer_name As String) As Boolean
'''    Dim i As Integer
'''     SelectPrinterByName = True
'''    For i = 0 To Printers.Count - 1
'''        If Printers(i).DeviceName = printer_name Then
'''            Set Printer = Printers(i)
'''            SelectPrinterByName = False
'''            Exit For
'''        End If
'''    Next i
'''End Function



Public Function KdMonitorPrinter() As Boolean
Dim lRet As Long
Dim SizeNeeded As Long
Dim buffer() As Long
Dim pDef As PRINTER_DEFAULTS
Dim index As Long
Dim mPinfo As PRINTER_INFO_2

'    pDef.DesiredAccess = PRINTER_ACCESS_USE
    pDef.DesiredAccess = PRINTER_ACCESS_ADMINISTER
    lRet = OpenPrinter(Printer.DeviceName, mhPrinter, pDef)
    

    Stop
    
'    lRet = GetPrinter(mhPrinter, 2, 0&, 0&, SizeNeeded)
    
    index = 2
    
    
    ReDim Preserve buffer(0 To 1) As Long
    lRet = GetPrinterApi(mhPrinter, index, buffer(0), UBound(buffer), SizeNeeded)
    ReDim Preserve buffer(0 To (SizeNeeded / 4) + 3) As Long
    lRet = GetPrinterApi(mhPrinter, index, buffer(0), UBound(buffer) * 4, SizeNeeded)

    Dim lIndex As Long
    
   
    With mPinfo '\\ This variable is of type PRINTER_INFO_2
       .pServerName = StringFromPointer(buffer(0), 1024)
       .pPrinterName = StringFromPointer(buffer(1), 1024)
       .pShareName = StringFromPointer(buffer(2), 1024)
       .pPortName = StringFromPointer(buffer(3), 1024)
       .pDriverName = StringFromPointer(buffer(4), 1024)
       .pComment = StringFromPointer(buffer(5), 1024)
       .pLocation = StringFromPointer(buffer(6), 1024)
       .pDevMode = buffer(7)
       .pSepFile = StringFromPointer(buffer(8), 1024)
       .pPrintProcessor = StringFromPointer(buffer(9), 1024)
       .pDatatype = StringFromPointer(buffer(10), 1024)
       .pParameters = StringFromPointer(buffer(11), 1024)
       .pSecurityDescriptor = buffer(12)
       .Attributes = buffer(13)
       .Priority = buffer(14)
       .DefaultPriority = buffer(15)
       .StartTime = buffer(16)
       .UntilTime = buffer(17)
       .Status = buffer(18)
       .JobsCount = buffer(19)
       .AveragePPM = buffer(20)
    End With



    Call ClosePrinter(mhPrinter)
Stop

    For lIndex = 0 To UBound(buffer)
        Debug.Print CStr(lIndex) & " = " & StringFromPointer(buffer(lIndex), 1024)
If lIndex Mod 20 = 0 Then
    Stop
End If
     Next
    

' so far the below crashes
'    Call InitializeNotifyOptions
'    Call StartWatching
End Function


Public Function TryToGetStatus()
Dim lRet As Long
Dim SizeNeeded As Long
Dim buffer() As Long
Dim jBuffer() As Long

Dim pDef As PRINTER_DEFAULTS
Dim index As Long
Dim mPinfo As PRINTER_INFO_2

'    pDef.DesiredAccess = PRINTER_ACCESS_USE
    pDef.DesiredAccess = PRINTER_ACCESS_ADMINISTER
    lRet = OpenPrinter(Printer.DeviceName, mhPrinter, pDef)
    

    Stop
    
    index = 2
    ReDim Preserve buffer(0 To 1) As Long
    lRet = GetPrinterApi(mhPrinter, index, buffer(0), UBound(buffer), SizeNeeded)
    ReDim Preserve buffer(0 To (SizeNeeded / 4) + 3) As Long
    lRet = GetPrinterApi(mhPrinter, index, buffer(0), UBound(buffer) * 4, SizeNeeded)

    Dim lIndex As Long
    
'    For lIndex = 0 To UBound(buffer)
'        Debug.Print CStr(lIndex) & " = " & StringFromPointer(buffer(lIndex), 1024)
'    Next
    
    With mPinfo '\\ This variable is of type PRINTER_INFO_2
       .pServerName = StringFromPointer(buffer(0), 1024)
       .pPrinterName = StringFromPointer(buffer(1), 1024)
       .pShareName = StringFromPointer(buffer(2), 1024)
       .pPortName = StringFromPointer(buffer(3), 1024)
       .pDriverName = StringFromPointer(buffer(4), 1024)
       .pComment = StringFromPointer(buffer(5), 1024)
       .pLocation = StringFromPointer(buffer(6), 1024)
       .pDevMode = buffer(7)
       .pSepFile = StringFromPointer(buffer(8), 1024)
       .pPrintProcessor = StringFromPointer(buffer(9), 1024)
       .pDatatype = StringFromPointer(buffer(10), 1024)
       .pParameters = StringFromPointer(buffer(11), 1024)
       .pSecurityDescriptor = buffer(12)
       .Attributes = buffer(13)
       .Priority = buffer(14)
       .DefaultPriority = buffer(15)
       .StartTime = buffer(16)
       .UntilTime = buffer(17)
       .Status = buffer(18)
       .JobsCount = buffer(19)
       .AveragePPM = buffer(20)
    End With

    If mPinfo.Status <> 0 Then
        Debug.Print CheckPrinterStatus(mPinfo.Status)
        Stop
    End If
Dim lJobStructCount As Long

'    If mPinfo.JobsCount > 0 Then
        ReDim Preserve jBuffer(0 To 1) As Long
        lRet = EnumJobs(mhPrinter, 0, 100, 2, ByVal 0, 0, SizeNeeded, lJobStructCount)
        ReDim Preserve jBuffer(0 To SizeNeeded / 4 - 1) As Long
        lRet = EnumJobs(mhPrinter, 0, mPinfo.JobsCount, 2, jBuffer(0), UBound(jBuffer), SizeNeeded, lJobStructCount)

If lJobStructCount < 1 Then
    Stop
Else

End If

Dim mJob As JOB_INFO_2
Dim lInde As Long
        Do While lIndex < UBound(jBuffer)
            With mJob
                .JobId = jBuffer(0 + lIndex)
                .pPrinterName = StringFromPointer(jBuffer(1 + lIndex), 1024)
                .pMachineName = StringFromPointer(jBuffer(2 + lIndex), 1024)
                .pUserName = StringFromPointer(jBuffer(3 + lIndex), 1024)
                .pDocument = StringFromPointer(jBuffer(4 + lIndex), 1024)
                .pNotifyName = StringFromPointer(jBuffer(5 + lIndex), 1024)
                .pDatatype = StringFromPointer(jBuffer(6 + lIndex), 1024)
                .pPrintProcessor = StringFromPointer(jBuffer(7 + lIndex), 1024)
                .pParameters = StringFromPointer(jBuffer(8 + lIndex), 1024)
                .pDriverName = StringFromPointer(jBuffer(9 + lIndex), 1024)
'                .pDevMode = jBuffer(10 + lIndex)
                .pStatus = StringFromPointer(jBuffer(11 + lIndex), 1024)
'                .pSecurityDescriptor = jBuffer(12 + lIndex)
                .Status = jBuffer(13 + lIndex)
                .Priority = jBuffer(14 + lIndex)
                .position = jBuffer(15 + lIndex)
                .StartTime = jBuffer(16 + lIndex)
                .UntilTime = jBuffer(17 + lIndex)
                .TotalPages = jBuffer(18 + lIndex)
                .Size = jBuffer(19 + lIndex)
                '.Submitted = jBuffer(20 + lIndex)
                .Time = jBuffer(21 + lIndex)
                .PagesPrinted = jBuffer(22 + lIndex)
Stop
            End With
            lIndex = lIndex + 23
        Loop
        Stop
'    End If

    Call ClosePrinter(mhPrinter)
Stop

' so far the below crashes
'    Call InitializeNotifyOptions
'    Call StartWatching
End Function

Public Function TestExample()
Dim pi1 As PRINTER_INFO_1  ' holds a little information about the printer
Dim hPrinter As Long  ' handle to the default printer once it is opened
Dim arraybuf() As Long  ' resizable array used as a buffer
Dim jobinfo As JOB_INFO_2  ' holds detailed info about a print job
Dim needed As Long  ' receives space needed in the buffer array
Dim numitems As Long  ' receives the number of items returned
Dim lendivfour As Long  ' the size in Long-type units of the jobinfo structure
Dim c As Long  ' counter variable
Dim retval As Long  ' return value

            ' -- Get the name of the default printer. --
            ' Determine how much space is needed to get the printer information.
        '    retval = EnumPrinters(PRINTER_ENUM_DEFAULT, "", 1, ByVal 0, 0, needed, numitems)
        '    ' Resize the array buffer to the needed size in bytes.
        '    ReDim arraybuf(0 To needed / 4 - 1)  ' remember each element is 4 bytes
        '    ' Retrieve the information about the default printer.
        '    retval = EnumPrinters(PRINTER_ENUM_DEFAULT, "", 1, arraybuf(0), needed, needed, numitems)
        '    ' Copy the printer name into the structure.  The rest is unnecessary.
        '    pi1.pName = Space(lstrlen(arraybuf(2)))
        '    retval = lstrcpy(pi1.pName, arraybuf(2))
        
        ' -- Obtain a handle to the default printer (using default configuration). --
retval = OpenPrinter(Printer.DeviceName, hPrinter, ByVal CLng(0))

    ' -- Enumerate the default printer's print jobs currently queued. --
    ' Determine how much space is needed to get the print jobs' information.
    retval = EnumJobs(hPrinter, 0, 100, 2, ByVal 0, 0, needed, numitems)
    ' Resize the array buffer to the needed size in bytes.
    If needed = 0 Then
        Debug.Print "No jobs found in the queue"
        GoTo ClosePtr
    End If
    ReDim arraybuf(0 To needed / 4 - 1)  ' remember each element is 4 bytes
        ' Retrieve the information about the print jobs.
    retval = EnumJobs(hPrinter, 0, 100, 2, arraybuf(0), needed, needed, numitems)
        ' Display the number of print jobs currently in the queue.
    If numitems > 0 Then
        Debug.Print "There are"; numitems; "print jobs currently in the queue."
    Else
        Debug.Print "No print jobs are currently in the queue."
    End If

' For each print job, copy its data into the structure.  Then display selected
' information from the structure.  For brevity, this example copies only a
' few of the data members into the structure.
    lendivfour = Len(jobinfo) / 4  ' this is the number of elements for each structure in the array
    For c = 0 To numitems - 1  ' loop through each item
        ' Copy selected information into the structure: the job ID number, the
        ' name of the user who printed it, the total number of pages, and
        ' the time it was added into the queue.
        jobinfo.JobId = arraybuf(lendivfour * c)  ' the first element of the array chunk
        jobinfo.pUserName = Space(lstrlen(arraybuf(lendivfour * c + 3)))  ' fourth element
        retval = lstrcpy(jobinfo.pUserName, arraybuf(lendivfour * c + 3))
        jobinfo.TotalPages = arraybuf(lendivfour * c + 18)  ' nineteenth element

        jobinfo.pStatus = StringFromPointer(arraybuf(lendivfour * c + 11), 1024)    ' 12th element
        CopyMemory jobinfo.Submitted, arraybuf(lendivfour * c + 20), Len(jobinfo.Submitted)  ' twenty-first element
        
        ' Display the copied information.
        Debug.Print "Job ID number:"; jobinfo.JobId
        Debug.Print "Printed by user: "; jobinfo.pUserName
        Debug.Print "Number of pages:"; jobinfo.TotalPages
        Debug.Print "Status: " & jobinfo.pStatus
        Debug.Print "Placed in queue on: ";
        ' (display the date and time stored in jobinfo.Submitted)
        Debug.Print jobinfo.Submitted.wMonth; "-"; jobinfo.Submitted.wDay; "-"; jobinfo.Submitted.wYear; " ";
        Debug.Print jobinfo.Submitted.wHour; ":"; jobinfo.Submitted.wMinute; ":"; jobinfo.Submitted.wSecond; " GMT"
    Next c

' Close the printer handle now that it is no longer needed.
ClosePtr:
    retval = ClosePrinter(hPrinter)
End Function


'\\ Use
'...
Public Function StartWatching()
Dim lpPrintInfoBuffer As Long
Dim pdwChange As Long
Dim mData As PRINTER_NOTIFY_INFO
Dim lEventsFound As Long
Dim lRet As Long
Dim pDef As PRINTER_DEFAULTS

    pDef.DesiredAccess = PRINTER_ACCESS_USE
'    pDef.DesiredAccess = PRINTER_ACCESS_ADMINISTER
    lRet = OpenPrinter(Printer.DeviceName, mhPrinter, pDef)

    


    mEventHandle = FindFirstPrinterChangeNotificationLong(mhPrinter, 0, 0, VarPtr(PrintOptions))
    

Again:
'    Call WaitForSingleObject(mEventHandle, INFINITE)
    Call WaitForSingleObject(mEventHandle, 1000 * 60)   ' a minute
    lEventsFound = lEventsFound + 1
    
    Call FindNextPrinterChangeNotificationByLong(mEventHandle, pdwChange, PrintOptions, lpPrintInfoBuffer)

    Call CopyMemoryPRINTER_NOTIFY_INFO(mData, lpPrintInfoBuffer, Len(mData))
    
    If mData.dwCount > 0 Then
        ReDim aData(1 To mData.dwCount) As PRINTER_NOTIFY_INFO_DATA
        
        Call CopyMemoryPRINTER_NOTIFY_INFO_DATA(aData(1), lpPrintInfoBuffer + Len(mData), Len(aData(1)) * mData.dwCount)
        
        
Stop
        
        
        Erase aData
        Call FreePrinterNotifyInfoByLong(lpPrintInfoBuffer)
    
    End If
    If lEventsFound < 5 Then
        GoTo Again
    End If
    
    Call FindClosePrinterChangeNotification(mEventHandle)
    Call ClosePrinter(mhPrinter)
    
End Function


Private Sub InitializeNotifyOptions()

'    With PrintOptions
'        .Version = 2 '\\ This must be set to 2
'        .Count = 2 '\\ There is job notification and printer notification
'        '\\ The type of printer events we are interested in...
'
'        With PrinterNotifyOptions(0)
'
'            .Type = PRINTER_CHANGE_ADD_JOB
'            ''    .Type = PRINTER_NOTIFY_TYPE
'
'            ReDim pFieldsPrinter(0 To 19) As Integer
'            '\\ Add the list of printer events you are interested in being notified about
'            '\\ to this list. Note that the fewer notifications you ask for the less of a
'            '\\ burden your app place upon the system.
'            pFieldsPrinter(0) = PRINTER_CHANGE_FAILED_CONNECTION_PRINTER
'            pFieldsPrinter(1) = PRINTER_CHANGE_ADD_JOB
'            pFieldsPrinter(2) = PRINTER_CHANGE_SET_JOB
'            pFieldsPrinter(3) = PRINTER_CHANGE_DELETE_JOB
'            pFieldsPrinter(4) = PRINTER_CHANGE_WRITE_JOB
'            pFieldsPrinter(5) = PRINTER_CHANGE_JOB
''            pFieldsPrinter(6) = PRINTER_CHANGE_TIMEOUT
'
'
'            .Count = (UBound(pFieldsPrinter) - LBound(pFieldsPrinter)) + 1 '\\ Add one as the array is zero based
'            .pFields = VarPtr(pFieldsPrinter(0))
'        End With
'        '\\ The type of print job events we are interested in...
'        With PrinterNotifyOptions(1)
'            .Type = JOB_NOTIFY_FIELD_STATUS
'            '' Origianally:    .Type = JOB_NOTIFY_TYPE
'            '\\ Add the list of print job events you are interested in being notified about
'            '\\ to this list. Note that the fewer notifications you ask for the less of a
'            '\\ burden your app place upon the system.
'            ReDim pFieldsJob(0 To 22) As Integer
'            pFieldsJob(0) = JOB_NOTIFY_FIELD_PRINTER_NAME
'            pFieldsJob(1) = JOB_NOTIFY_FIELD_MACHINE_NAME
'            pFieldsJob(2) = JOB_NOTIFY_FIELD_PORT_NAME
'            pFieldsJob(3) = JOB_NOTIFY_FIELD_USER_NAME
'            pFieldsJob(4) = JOB_NOTIFY_FIELD_NOTIFY_NAME
'            pFieldsJob(5) = JOB_NOTIFY_FIELD_DATATYPE
'            pFieldsJob(6) = JOB_NOTIFY_FIELD_PRINT_PROCESSOR
'            pFieldsJob(7) = JOB_NOTIFY_FIELD_PARAMETERS
'            pFieldsJob(8) = JOB_NOTIFY_FIELD_DRIVER_NAME
'            pFieldsJob(9) = JOB_NOTIFY_FIELD_DEVMODE
'            pFieldsJob(10) = JOB_NOTIFY_FIELD_STATUS
'            pFieldsJob(11) = JOB_NOTIFY_FIELD_STATUS_STRING
'            pFieldsJob(12) = JOB_NOTIFY_FIELD_DOCUMENT
'            pFieldsJob(13) = JOB_NOTIFY_FIELD_PRIORITY
'            pFieldsJob(14) = JOB_NOTIFY_FIELD_POSITION
'            pFieldsJob(15) = JOB_NOTIFY_FIELD_SUBMITTED
'            pFieldsJob(16) = JOB_NOTIFY_FIELD_START_TIME
'            pFieldsJob(17) = JOB_NOTIFY_FIELD_UNTIL_TIME
'            pFieldsJob(18) = JOB_NOTIFY_FIELD_TIME
'            pFieldsJob(19) = JOB_NOTIFY_FIELD_TOTAL_PAGES
'            pFieldsJob(20) = JOB_NOTIFY_FIELD_PAGES_PRINTED
'            pFieldsJob(21) = JOB_NOTIFY_FIELD_TOTAL_BYTES
'            .Count = (UBound(pFieldsJob) - LBound(pFieldsJob)) + 1 '\\ Add one as the array is zero based
'            .pFields = VarPtr(pFieldsJob(0))
'        End With
'        .lpPrintNotifyOptions = VarPtr(PrinterNotifyOptions(0))
'    End With

End Sub



Public Function SelectPrinter() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sPrinter As String
Dim sJob As String

    strProcName = ClassName & ".SelectPrinter"



'    Debug.Print oPDlg.PrinterName

    'DoCmd.RunCommand acCmdPrint
    ' will set Printer to the correct printer, then get the settings from that..


Dim hInst As Long
Dim Thread As Long
Dim hwnd As Long

    
    ParenthWnd = hwnd
    PrintDialog.hWndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
'    If centerForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
    sPrinter = Printer.DeviceName
    
    Debug.Print sPrinter
    PrintDialog.flags = &H100& + &H100000 + &H4&
    
    SelectPrinter = PrintDlg(PrintDialog)
    sPrinter = Printer.DeviceName
    
    Debug.Print sPrinter
    
    Debug.Print PrintDialog.hdc
    Debug.Print PrintDialog.hDevMode
    Debug.Print PrintDialog.hDevNames
    
    Debug.Print PrintDialog.hInstance
    Debug.Print PrintDialog.hPrintTemplate
    Debug.Print PrintDialog.hWndOwner
    Debug.Print PrintDialog.lpfnPrintHook
    Debug.Print PrintDialog.lStructSize
    
    
    Debug.Print sJob
'    Call Me.ShowPrinter(Application.hWndAccessApp)



Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



'''Public Sub SetPrinterForm(FormName As String, Optional copies As Long = 1)
'''Dim hPrinter As Long
'''Dim cbRequired  As Long
'''Dim pd As PRINTER_DEFAULTS
'''Dim pi2 As PRINTER_INFO_2
'''Dim buff() As Byte
'''Dim dm As DEVMODE
'''Dim bDevMode() As Byte
'''Dim lBytestNeeded As Long
'''
'''    'pd.DesiredAccess = standard_rights_required Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE
'''    'pd.DesiredAccess = standard_rights_required
'''    If (OpenPrinter(Printer.DeviceName, hPrinter, pd) <> 0) Then
'''        '    Result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
'''        If (GetPrinter(hPrinter, 2&, 0&, 0&, cbRequired) = 0) Then
'''            ReDim buff(1 To cbRequired) As Byte
'''
'''            If (GetPrinter(hPrinter, 2, buff(1), cbRequired, cbRequired) <> 0) Then
'''                Call CopyMemory(pi2, buff(1), Len(pi2))
'''
'''                ReDim bDevMode(1 To cbRequired)
'''                If (pi2.pDevMode) Then
'''                    Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
'''                Else
'''                    Call DocumentProperties(0&, hPrinter, Printer.DeviceName, bDevMode(1), 0&, DM_OUT_BUFFER)
'''                End If
'''
'''                Call CopyMemory(dm, bDevMode(1), Len(dm))
'''
'''                dm.dmFormName = FormName & Chr(0)
'''                dm.dmCopies = copies
'''                dm.dmFields = DM_FORMNAME Or DM_COPIES
'''
'''                Call CopyMemory(bDevMode(1), dm, Len(dm))
'''                Call DocumentProperties(0&, hPrinter, Printer.DeviceName, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
'''                pi2.pDevMode = VarPtr(bDevMode(1))
'''
'''                Call SetPrinter(hPrinter, 2, pi2, 0&)
'''            End If
'''        End If
'''        Call ClosePrinter(hPrinter)
'''    End If
'''
'''End Sub


''' ###################### for inspecting the queue status and such:
''' ###################### for inspecting the queue status and such:

'''Private Function GetString(ByVal PtrStr As Long) As String
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim StrBuff As String * 256
'''
'''    strProcName = ClassName & ".GetString"
'''
'''   'Check for zero address
'''   If PtrStr = 0 Then
'''      GetString = " "
'''      GoTo Block_Exit
'''   End If
'''
'''   'Copy data from PtrStr to buffer.
'''   CopyMemory ByVal StrBuff, ByVal PtrStr, 256
'''
'''   'Strip any trailing nulls from string.
'''   GetString = StripNulls(StrBuff)
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    GoTo Block_Exit
'''End Function
'''
'''Private Function StripNulls(OriginalStr As String) As String
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''
'''    strProcName = ClassName & ".StripNulls"
'''
'''   'Strip any trailing nulls from input string.
'''   If (InStr(OriginalStr, Chr(0)) > 0) Then
'''      OriginalStr = left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
'''   End If
'''
'''   'Return modified string.
'''   StripNulls = OriginalStr
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    GoTo Block_Exit
'''End Function
'''
'''Private Function PtrCtoVbString(Add As Long) As String
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim sTemp As String * 512
'''Dim x As Long
'''
'''    strProcName = ClassName & ".PtrCtoVbString"
'''
'''    x = lstrcpy(sTemp, Add)
'''    If (InStr(1, sTemp, Chr(0)) = 0) Then
'''         PtrCtoVbString = ""
'''    Else
'''         PtrCtoVbString = left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
'''    End If
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    GoTo Block_Exit
'''End Function

Private Function CheckPrinterStatus(PI2Status As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim tempStr As String

    strProcName = ClassName & ".CheckPrinterStatus"


   If PI2Status = 0 Then   ' Return "Ready"
      CheckPrinterStatus = "Printer Status = Ready" & vbCrLf
   Else
      tempStr = ""   ' Clear
      If (PI2Status And PRINTER_STATUS_BUSY) Then
         tempStr = tempStr & "Busy  "
      End If

      If (PI2Status And PRINTER_STATUS_DOOR_OPEN) Then
         tempStr = tempStr & "Printer Door Open  "
      End If

      If (PI2Status And PRINTER_STATUS_ERROR) Then
         tempStr = tempStr & "Printer Error  "
      End If

      If (PI2Status And PRINTER_STATUS_INITIALIZING) Then
         tempStr = tempStr & "Initializing  "
      End If

      If (PI2Status And PRINTER_STATUS_IO_ACTIVE) Then
         tempStr = tempStr & "I/O Active  "
      End If

      If (PI2Status And PRINTER_STATUS_MANUAL_FEED) Then
         tempStr = tempStr & "Manual Feed  "
      End If

      If (PI2Status And PRINTER_STATUS_NO_TONER) Then
         tempStr = tempStr & "No Toner  "
      End If

      If (PI2Status And PRINTER_STATUS_NOT_AVAILABLE) Then
         tempStr = tempStr & "Not Available  "
      End If

      If (PI2Status And PRINTER_STATUS_OFFLINE) Then
         tempStr = tempStr & "Off Line  "
      End If

      If (PI2Status And PRINTER_STATUS_OUT_OF_MEMORY) Then
         tempStr = tempStr & "Out of Memory  "
      End If

      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempStr = tempStr & "Output Bin Full  "
      End If

      If (PI2Status And PRINTER_STATUS_PAGE_PUNT) Then
         tempStr = tempStr & "Page Punt  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_JAM) Then
         tempStr = tempStr & "Paper Jam  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_OUT) Then
         tempStr = tempStr & "Paper Out  "
      End If

      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempStr = tempStr & "Output Bin Full  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_PROBLEM) Then
         tempStr = tempStr & "Page Problem  "
      End If

      If (PI2Status And PRINTER_STATUS_PAUSED) Then
         tempStr = tempStr & "Paused  "
      End If

      If (PI2Status And PRINTER_STATUS_PENDING_DELETION) Then
         tempStr = tempStr & "Pending Deletion  "
      End If

      If (PI2Status And PRINTER_STATUS_PRINTING) Then
         tempStr = tempStr & "Printing  "
      End If

      If (PI2Status And PRINTER_STATUS_PROCESSING) Then
         tempStr = tempStr & "Processing  "
      End If

      If (PI2Status And PRINTER_STATUS_TONER_LOW) Then
         tempStr = tempStr & "Toner Low  "
      End If

      If (PI2Status And PRINTER_STATUS_USER_INTERVENTION) Then
         tempStr = tempStr & "User Intervention  "
      End If

      If (PI2Status And PRINTER_STATUS_WAITING) Then
         tempStr = tempStr & "Waiting  "
      End If

      If (PI2Status And PRINTER_STATUS_WARMING_UP) Then
         tempStr = tempStr & "Warming Up  "
      End If

      'Did you find a known status?
      If Len(tempStr) = 0 Then
         tempStr = "Unknown Status of " & PI2Status
      End If

      'Return the Status
      CheckPrinterStatus = "Printer Status = " & tempStr & vbCrLf
   End If
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



'''Public Function CheckPrinter(PrinterStr As String, JobStr As String) As String
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim hPrinter As Long
'''Dim ByteBuf As Long
'''Dim BytesNeeded As Long
'''Dim pi2 As PRINTER_INFO_2
'''Dim JI2 As JOB_INFO_2
'''Dim PrinterInfo() As Byte
'''Dim JobInfo() As Byte
'''Dim Result As Long
'''Dim LastError As Long
'''Dim PrinterName As String
'''Dim tempStr As String
'''Dim NumJI2 As Long
'''Dim pDefaults As PRINTER_DEFAULTS
'''Dim i As Integer
'''
'''
'''    strProcName = ClassName & ".CheckPrinter"
'''
'''   'Set a default return value if no errors occur.
'''   CheckPrinter = "Printer info retrieved"
'''
'''   'NOTE: You can pick a printer from the Printers Collection
'''   'or use the EnumPrinters() API to select a printer name.
'''
'''   'Use the default printer of Printers collection.
'''   'This is typically, but not always, the system default printer.
'''   PrinterName = Printer.DeviceName
'''
'''   'Set desired access security setting.
''''   pDefaults.DesiredAccess = PRINTER_ACCESS_USE
'''' pDefaults.DesiredAccess = PRINTER_ACCESS_ADMINISTER
'''
'''   'Call API to get a handle to the printer.
'''   Result = OpenPrinter(PrinterName, hPrinter, pDefaults)
'''   If Result = 0 Then
'''      'If an error occurred, display an error and exit sub.
'''      CheckPrinter = "Cannot open printer " & PrinterName & ", Error: " & Err.LastDllError
'''      GoTo Block_Exit
'''   End If
'''
'''   'Init BytesNeeded
'''   BytesNeeded = 0
'''
'''   'Clear the error object of any errors.
'''   Err.Clear
'''
'''   'Determine the buffer size that is needed to get printer info.
'''   '       (GetPrinter(hPrinter, 2&, 0&, 0&, cbRequired) = 0) Then
'''   Result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
'''
'''   'Check for error calling GetPrinter.
'''   If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
'''      'Display an error message, close printer, and exit sub.
'''      CheckPrinter = " > GetPrinter Failed on initial call! <"
'''      ClosePrinter hPrinter
'''      GoTo Block_Exit
'''   End If
'''
'''   'Note that in Charles Petzold's book "Programming Windows 95," he
'''   'states that because of a problem with GetPrinter on Windows 95 only, you
'''   'must allocate a buffer as much as three times larger than the value
'''   'returned by the initial call to GetPrinter. This is not done here.
'''   ReDim PrinterInfo(1 To BytesNeeded)
'''
'''   ByteBuf = BytesNeeded
'''
'''   'Call GetPrinter to get the status.
'''   Result = GetPrinter(hPrinter, 2, PrinterInfo(1), ByteBuf, BytesNeeded)
'''
'''   'Check for errors.
'''   If Result = 0 Then
'''      'Determine the error that occurred.
'''      LastError = Err.LastDllError()
'''
'''      'Display error message, close printer, and exit sub.
'''      CheckPrinter = "Couldn't get Printer Status!  Error = " & LastError
'''      ClosePrinter hPrinter
'''      GoTo Block_Exit
'''   End If
'''
'''   'Copy contents of printer status byte array into a
'''   'PRINTER_INFO_2 structure to separate the individual elements.
'''   CopyMemory pi2, PrinterInfo(1), Len(pi2)
'''
'''   'Check if printer is in ready state.
'''   PrinterStr = CheckPrinterStatus(pi2.Status)
'''
'''   'Add printer name, driver, and port to list.
'''   PrinterStr = PrinterStr & "Printer Name = " & GetString(pi2.pPrinterName) & vbCrLf
'''   PrinterStr = PrinterStr & "Printer Driver Name = " & GetString(pi2.pDriverName) & vbCrLf
'''   PrinterStr = PrinterStr & "Printer Port Name = " & GetString(pi2.pPortName) & vbCrLf
'''
'''   'Call API to get size of buffer that is needed.
'''   Result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, ByVal 0&, 0&, BytesNeeded, NumJI2)
'''
'''   'Check if there are no current jobs, and then display appropriate message.
'''   If BytesNeeded = 0 Then
'''      JobStr = "No Print Jobs!"
'''   Else
'''      'Redim byte array to hold info about print job.
'''      ReDim JobInfo(0 To BytesNeeded)
'''
'''      'Call API to get print job info.
'''      Result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, JobInfo(0), BytesNeeded, ByteBuf, NumJI2)
'''
'''      'Check for errors.
'''      If Result = 0 Then
'''         'Get and display error, close printer, and exit sub.
'''         LastError = Err.LastDllError
'''         CheckPrinter = " > EnumJobs Failed on second call! <  Error = " & LastError
'''         ClosePrinter hPrinter
'''         GoTo Block_Exit
'''      End If
'''
'''      'Copy contents of print job info byte array into a
'''      'JOB_INFO_2 structure to separate the individual elements.
'''      For i = 0 To NumJI2 - 1   ' Loop through jobs and walk the buffer
'''          CopyMemory JI2, JobInfo(i * Len(JI2)), Len(JI2)
'''
'''          ' List info available on Jobs.
'''          Debug.Print "Job ID" & vbTab & JI2.JobId
'''          Debug.Print "Name Of Printer" & vbTab & GetString(JI2.pPrinterName)
'''          Debug.Print "Name Of Machine That Created Job" & vbTab & GetString(JI2.pMachineName)
'''          Debug.Print "Print Job Owner's Name" & vbTab & GetString(JI2.pUserName)
'''          Debug.Print "Name Of Document" & vbTab & GetString(JI2.pDocument)
'''          Debug.Print "Name Of User To Notify" & vbTab & GetString(JI2.pNotifyName)
'''          Debug.Print "Type Of Data" & vbTab & GetString(JI2.pDatatype)
'''          Debug.Print "Print Processor" & vbTab & GetString(JI2.pPrintProcessor)
'''          Debug.Print "Print Processor Parameters" & vbTab & GetString(JI2.pParameters)
'''          Debug.Print "Print Driver Name" & vbTab & GetString(JI2.pDriverName)
'''          Debug.Print "Print Job 'P' Status" & vbTab & GetString(JI2.pStatus)
'''          Debug.Print "Print Job Status" & vbTab & JI2.Status
'''          Debug.Print "Print Job Priority" & vbTab & JI2.Priority
'''          Debug.Print "Position in Queue" & vbTab & JI2.Position
'''          Debug.Print "Earliest Time Job Can Be Printed" & vbTab & JI2.StartTime
'''          Debug.Print "Latest Time Job Will Be Printed" & vbTab & JI2.UntilTime
'''          Debug.Print "Total Pages For Entire Job" & vbTab & JI2.TotalPages
'''          Debug.Print "Size of Job In Bytes" & vbTab & JI2.Size
'''          'Because of a bug in Windows NT 3.51, the time member is not set correctly.
'''          'Therefore, do not use the time member on Windows NT 3.51.
'''          Debug.Print "Elapsed Print Time" & vbTab & JI2.time
'''          Debug.Print "Pages Printed So Far" & vbTab & JI2.PagesPrinted
'''
'''          'Display basic job status info.
'''          JobStr = JobStr & "Job ID = " & JI2.JobId & vbCrLf & "Total Pages = " & JI2.TotalPages & vbCrLf
'''
'''          tempStr = ""   'Clear
'''          'Check for a ready state.
'''          If JI2.pStatus = 0& Then   ' If pStatus is Null, check Status.
'''            If JI2.Status = 0 Then
'''               tempStr = tempStr & "Ready!  " & vbCrLf
'''            Else  'Check for the various print job states.
'''               If (JI2.Status And JOB_STATUS_SPOOLING) Then
'''                  tempStr = tempStr & "Spooling  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_OFFLINE) Then
'''                  tempStr = tempStr & "Off line  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_PAUSED) Then
'''                  tempStr = tempStr & "Paused  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_ERROR) Then
'''                  tempStr = tempStr & "Error  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_PAPEROUT) Then
'''                  tempStr = tempStr & "Paper Out  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_PRINTING) Then
'''                  tempStr = tempStr & "Printing  "
'''               End If
'''
'''               If (JI2.Status And JOB_STATUS_USER_INTERVENTION) Then
'''                  tempStr = tempStr & "User Intervention Needed  "
'''               End If
'''
'''               If Len(tempStr) = 0 Then
'''                  tempStr = "Unknown Status of " & JI2.Status
'''               End If
'''            End If
'''        Else
'''            ' Dereference pStatus.
'''            tempStr = PtrCtoVbString(JI2.pStatus)
'''        End If
'''
'''          'Report the Job status.
'''          JobStr = JobStr & tempStr & vbCrLf
'''          Debug.Print JobStr & tempStr
'''      Next i
'''   End If
'''
'''   'Close the printer handle.
'''   ClosePrinter hPrinter
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    GoTo Block_Exit
'''End Function



'' ################### Common Dialog
'
'Public Function ShowOpenFilesDialog(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim ret As Long
'Dim Count As Integer
'Dim fileNameHolder As String
'Dim LastCharacter As Integer
'Dim NewCharacter As Integer
'Dim tempFiles(1 To 200) As String
'Dim hInst As Long
'Dim Thread As Long
'
'
'    strProcName = ClassName & ".ShowOpenFilesDialog"
'
'    strProcName = Me.ClassName
'
'    lParenthWnd = hWnd
'    tFileDialog.lStructSize = Len(tFileDialog)
'    tFileDialog.hwndOwner = hWnd
'    tFileDialog.lpstrFileTitle = Space$(2048)
'    tFileDialog.nMaxFileTitle = Len(tFileDialog.lpstrFileTitle)
'    tFileDialog.lpstrFile = tFileDialog.lpstrFile & Space$(2047) & Chr$(0)
'    tFileDialog.nMaxFile = Len(tFileDialog.lpstrFile)
'
'    'If tFileDialog.flags = 0 Then
'        tFileDialog.flags = OFS_FILE_OPEN_FLAGS
'    'End If
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
''    If centerForm = True Then
''        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
''    Else
''        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
''    End If
'
'    ret = GetOpenFileName(tFileDialog)
'
'    If ret Then
'        If Trim$(tFileDialog.lpstrFileTitle) = "" Then
'            LastCharacter = 0
'            Count = 0
'            While ShowOpenFilesDialog.nFilesSelected = 0
'                NewCharacter = InStr(LastCharacter + 1, tFileDialog.lpstrFile, Chr$(0), vbTextCompare)
'                If Count > 0 Then
'                    tempFiles(Count) = Mid(tFileDialog.lpstrFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
'                Else
'                    ShowOpenFilesDialog.sLastDirectory = Mid(tFileDialog.lpstrFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
'                End If
'                Count = Count + 1
'                If InStr(NewCharacter + 1, tFileDialog.lpstrFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, tFileDialog.lpstrFile, Chr$(0) & Chr$(0), vbTextCompare) Then
'                    tempFiles(Count) = Mid(tFileDialog.lpstrFile, NewCharacter + 1, InStr(NewCharacter + 1, tFileDialog.lpstrFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
'                    ShowOpenFilesDialog.nFilesSelected = Count
'                End If
'                LastCharacter = NewCharacter
'            Wend
'            ReDim ShowOpenFilesDialog.sFiles(1 To ShowOpenFilesDialog.nFilesSelected)
'            For Count = 1 To ShowOpenFilesDialog.nFilesSelected
'                ShowOpenFilesDialog.sFiles(Count) = tempFiles(Count)
'            Next
'        Else
'            ReDim ShowOpenFilesDialog.sFiles(1 To 1)
'            ShowOpenFilesDialog.sLastDirectory = left$(tFileDialog.lpstrFile, tFileDialog.nFileOffset)
'            ShowOpenFilesDialog.nFilesSelected = 1
'            ShowOpenFilesDialog.sFiles(1) = Mid(tFileDialog.lpstrFile, tFileDialog.nFileOffset + 1, InStr(1, tFileDialog.lpstrFile, Chr$(0), vbTextCompare) - tFileDialog.nFileOffset - 1)
'        End If
'        ShowOpenFilesDialog.bCanceled = False
'        GoTo Block_Exit
'    Else
'        ShowOpenFilesDialog.sLastDirectory = ""
'        ShowOpenFilesDialog.nFilesSelected = 0
'        ShowOpenFilesDialog.bCanceled = True
'        Erase ShowOpenFilesDialog.sFiles
'        GoTo Block_Exit
'    End If
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'
'
'Public Function ShowSave(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedFile
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim ret As Long
'Dim hInst As Long
'Dim Thread As Long
'
'    strProcName = ClassName & ".ShowSave"
'
'    lParenthWnd = hWnd
'    tFileDialog.lStructSize = Len(tFileDialog)
'    tFileDialog.hwndOwner = hWnd
'    tFileDialog.lpstrFileTitle = Space$(2048)
'    tFileDialog.nMaxFileTitle = Len(tFileDialog.lpstrFileTitle)
'    tFileDialog.lpstrFile = Space$(2047) & Chr$(0)
'    tFileDialog.nMaxFile = Len(tFileDialog.lpstrFile)
'
'    If tFileDialog.flags = 0 Then
'        tFileDialog.flags = OFS_FILE_SAVE_FLAGS
'    End If
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If centerForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ret = GetSaveFileName(tFileDialog)
'    ReDim ShowSave.sFiles(1)
'
'    If ret Then
'        ShowSave.sLastDirectory = left$(tFileDialog.lpstrFile, tFileDialog.nFileOffset)
'        ShowSave.nFilesSelected = 1
'        ShowSave.sFiles(1) = Mid(tFileDialog.lpstrFile, tFileDialog.nFileOffset + 1, InStr(1, tFileDialog.lpstrFile, Chr$(0), vbTextCompare) - tFileDialog.nFileOffset - 1)
'        ShowSave.bCanceled = False
'        GoTo Block_Exit
'    Else
'        ShowSave.sLastDirectory = ""
'        ShowSave.nFilesSelected = 0
'        ShowSave.bCanceled = True
'        Erase ShowSave.sFiles
'        GoTo Block_Exit
'    End If
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'
'
'Public Function ShowColor(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As SelectedColor
'On Error GoTo Block_Err
'Dim strProcName As String
'
'Dim customcolors() As Byte  ' dynamic (resizable) array
'Dim i As Integer
'Dim ret As Long
'Dim hInst As Long
'Dim Thread As Long
'
'
'    strProcName = ClassName & ".ShowColor"
'
'    lParenthWnd = hWnd
'    If tColorDialog.lpCustColors = "" Then
'        ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
'
'        For i = LBound(customcolors) To UBound(customcolors)
'          customcolors(i) = 254 ' sets all custom colors to white
'        Next i
'
'        tColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
'    End If
'
'    tColorDialog.hwndOwner = hWnd
'    tColorDialog.lStructSize = Len(tColorDialog)
'    tColorDialog.flags = COLOR_FLAGS
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If centerForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ret = ChooseColor(tColorDialog)
'    If ret Then
'        ShowColor.bCanceled = False
'        ShowColor.oSelectedColor = tColorDialog.rgbResult
'        GoTo Block_Exit
'    Else
'        ShowColor.bCanceled = True
'        ShowColor.oSelectedColor = &H0&
'        GoTo Block_Exit
'    End If
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'
'
'Public Function ShowFont(ByVal hWnd As Long, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As SelectedFont
'On Error GoTo Block_Err
'Dim strProcName As String
'
'Dim ret As Long
'Dim lfLogFont As LOGFONT
'Dim hInst As Long
'Dim Thread As Long
'Dim i As Integer
'
'    strProcName = ClassName & ".ShowFont"
'
'    lParenthWnd = hWnd
'    tFontDialog.nSizeMax = 0
'    tFontDialog.nSizeMin = 0
'    'tFontDialog.nFontType =  screen.font
'    tFontDialog.hwndOwner = hWnd
'    tFontDialog.hDC = 0
'    tFontDialog.lpfnHook = 0
'    tFontDialog.lCustData = 0
'    tFontDialog.lpLogFont = VarPtr(lfLogFont)
'    If tFontDialog.iPointSize = 0 Then
'        tFontDialog.iPointSize = 10 * 10
'    End If
'    tFontDialog.lpTemplateName = Space$(2048)
'    tFontDialog.rgbColors = RGB(0, 255, 255)
'    tFontDialog.lStructSize = Len(tFontDialog)
'
'    If tFontDialog.flags = 0 Then
'        tFontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
'    End If
'
'    For i = 0 To Len(startingFontName) - 1
'        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
'    Next
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If centerForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ret = ChooseFont(tFontDialog)
'
'    If ret Then
'        ShowFont.bCanceled = False
'        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
'        ShowFont.bItalic = lfLogFont.lfItalic
'        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
'        ShowFont.bUnderline = lfLogFont.lfUnderline
'        ShowFont.lColor = tFontDialog.rgbColors
'        ShowFont.nSize = tFontDialog.iPointSize / 10
'        For i = 0 To 31
'            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
'        Next
'
'        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
'        GoTo Block_Exit
'    Else
'        ShowFont.bCanceled = True
'        GoTo Block_Exit
'    End If
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'
'
'Public Function ShowPrinter(ByVal hWnd As Long, Optional ByVal centerForm As Boolean = True) As Long
'On Error GoTo Block_Err
'Dim strProcName As String
'
'Dim hInst As Long
'Dim Thread As Long
'
'    strProcName = ClassName & ".ShowPrinter"
'
'    lParenthWnd = hWnd
'    tPrintDialog.hwndOwner = hWnd
'    tPrintDialog.lStructSize = Len(tPrintDialog)
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If centerForm = True Then
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ShowPrinter = PrintDlg(tPrintDialog)
'    Debug.Print Printer.DeviceName
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function
'


Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String

  Dim sRet As String
  Dim lRet As Long

  If lpString = 0 Then
    StringFromPointer = ""
    Exit Function
  End If

  If IsBadStringPtrByLong(lpString, lMaxLength) Then
    '\\ An error has occured - do not attempt to use this pointer
      StringFromPointer = ""
    Exit Function
  End If

  '\\ Pre-initialise the return string...
  sRet = Space$(lMaxLength)
  CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
  If Err.LastDllError = 0 Then
    If InStr(sRet, Chr$(0)) > 0 Then
      sRet = left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
  End If

  StringFromPointer = sRet

End Function