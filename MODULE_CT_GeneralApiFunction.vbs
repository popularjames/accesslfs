Option Compare Database
Option Explicit
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
 
 '32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) _
As Long
 
Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
 
'DLC 02/09/2010 - Bug 2321 (Bring IE window to the foreground prior to download)
'HC  05/  /2010 - centralized most of the winapis to this class in preparation for the move to 2010, added all the types as public types
'DLC 06/09/2010 - Changed parameter lpTempFileName in GetTempFile to ByRef to return the retrieved value
'DLC 01/30/2012 - Added GetTimer() and declaration for GetSystemTime to allow for more accurate timing

Private Const CC_PREVENTFULLOPEN = &H4
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function CoCreateGuid_Alt Lib "ole32.dll" Alias "CoCreateGuid" (pGuid As Any) As Long
Private Declare PtrSafe Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
'DLC 02/09/10 - Used to bring a window to the foreground
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long

Private Declare PtrSafe Function StringFromGUID2_Alt Lib "ole32.dll" Alias "StringFromGUID2" (pGuid As Any, ByVal address As Long, ByVal max As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Boolean
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare PtrSafe Sub ReleaseCapture Lib "user32" ()

Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, rectangle As RECT) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare PtrSafe Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupDlg As PAGESETUPDLG) As Long
Private Declare PtrSafe Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

'Timer API:
Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As Long
Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'OLE-Stream functions :
Public Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As LongPtr, ByRef ppstm As Any) As Long
Public Declare PtrSafe Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As LongPtr) As Long

Public Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
Public Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Initialization GDIP:
Public Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
'Tear down GDIP:
Public Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
'Load GDIP-Image from file :
Public Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As LongPtr, BITMAP As LongPtr) As Long
'Create GDIP- graphical area from Windows-DeviceContext:
Public Declare PtrSafe Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As LongPtr, GpGraphics As LongPtr) As Long
'Delete GDIP graphical area :
Public Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
'Copy GDIP-Image to graphical area:
Public Declare PtrSafe Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
'Clear allocated bitmap memory from GDIP :
Public Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal Image As LongPtr) As Long
'Retrieve windows bitmap handle from GDIP-Image:
Public Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal Background As LongPtr) As Long
'Retrieve Windows-Icon-Handle from GDIP-Image:
Public Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr) As Long
'Scaling GDIP-Image size:
Public Declare PtrSafe Function GdipGetImageThumbnail Lib "gdiplus" (ByVal Image As LongPtr, ByVal thumbWidth As LongPtr, ByVal thumbHeight As LongPtr, thumbImage As LongPtr, Optional ByVal CallBack As LongPtr = 0, Optional ByVal callbackData As LongPtr = 0) As Long
'Retrieve GDIP-Image from Windows-Bitmap-Handle:
Public Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, BITMAP As LongPtr) As Long
'Retrieve GDIP-Image from Windows-Icon-Handle:
Public Declare PtrSafe Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As LongPtr, BITMAP As LongPtr) As Long
'Retrieve width of a GDIP-Image (Pixel):
Public Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As LongPtr, Width As LongPtr) As Long
'Retrieve height of a GDIP-Image (Pixel):
Public Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As LongPtr, Height As LongPtr) As Long
'Save GDIP-Image to file in seletable format:
Public Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, clsidEncoder As GUID, encoderParams As Any) As Long
'Save GDIP-Image in OLE-Stream with seletable format:
Public Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As LongPtr, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
'Retrieve GDIP-Image from OLE-Stream-Object:
Public Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IUnknown, Image As LongPtr) As Long
'Create a gdip image from scratch
Public Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As LongPtr, ByVal Height As LongPtr, ByVal stride As LongPtr, ByVal PixelFormat As LongPtr, scan0 As Any, BITMAP As LongPtr) As Long
'Get the DC of an gdip image
Public Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As LongPtr, graphics As LongPtr) As Long
'Blit the contents of an gdip image to another image DC using positioning
Public Declare PtrSafe Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal Image As LongPtr, ByVal dstx As LongPtr, ByVal dsty As LongPtr, ByVal dstwidth As LongPtr, ByVal dstheight As LongPtr, ByVal srcx As LongPtr, ByVal srcy As LongPtr, ByVal srcwidth As LongPtr, ByVal srcheight As LongPtr, ByVal srcUnit As LongPtr, Optional ByVal imageAttributes As LongPtr = 0, Optional ByVal CallBack As LongPtr = 0, Optional ByVal callbackData As LongPtr = 0) As Long

'Convert a windows bitmap to OLE-Picture :
'Windows-Bitmap in OLE-Picture konvertieren:
Public Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As LongPtr, IPic As Object) As Long
'GUID-Typ aus String erhalten:
Public Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long

Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'SA 03/22/2012 - Added function to get screen dimensions
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Const GUID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    'IPicture

Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type ChooseColor
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
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Public Type DevMode
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
End Type

Public Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Public Type PAGESETUPDLG
        lStructSize As Long
        hWndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Public Type PageSetupResults
    bCancelled As Boolean
    lStructSize As Long
    hHwndOwner As Long
    tDevMode As DevMode
    tDevNames As DEVNAMES
    lFlags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Public Type TSize
    X As Double
    Y As Double
End Type

Public Enum PicFileType
    pictypeBMP = 1
    pictypeGIF = 2
    pictypePNG = 3
    pictypeJPG = 4
End Enum

Public Type GDIPStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As LongPtr
    SuppressExternalCodecs As LongPtr
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type PICTDESC
    cbSizeOfStruct As Long
    PicType As Long
    hImage As LongPtr
    xExt As Long
    yExt As Long
End Type
Public Type EncoderParameter
    UUID As GUID
    NumberOfValues As LongPtr
    Type As LongPtr
    Value As LongPtr
End Type
Public Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Public Sub ShellEx(FileName As String)
    ShellExecute GetForegroundWindow, "Open", FileName, "", "", 1
End Sub

Public Function FreeMemory(ByVal hMem As Long) As Long
    FreeMemory = GlobalFree(hMem)
End Function

Public Function GetPageSetupDialog(pPagesetupDlg As PAGESETUPDLG) As Long
    GetPageSetupDialog = PAGESETUPDLG(pPagesetupDlg)
End Function

Public Function GetWindowPositionLong(ByVal hwnd As Long, ByVal nIndex As Long) As Long
    GetWindowPositionLong = GetWindowLong(hwnd, nIndex)
End Function
Public Function SetWindowPositionLong(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    SetWindowPositionLong = SetWindowLong(hwnd, nIndex, dwNewLong)
End Function
Public Function GetCursorPosition(lpPoint As POINTAPI) As Long
    GetCursorPosition = GetCursorPos(lpPoint)
End Function

Public Sub CaptureRelease()
    Call ReleaseCapture
End Sub
Public Function GetDevice(ByVal hwnd As Long) As Long
    GetDevice = GetDC(hwnd)
End Function

Public Function GetDeviceCapitals(ByVal hdc As Long, ByVal nIndex As Long) As Long
    GetDeviceCapitals = GetDeviceCaps(hdc, nIndex)
End Function
' someone had this as a long afor the last param when it's Any
Public Sub SendWindowMessage(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Call SendMessage(hwnd, wMsg, wParam, lParam)
End Sub
Public Function LockWindow(ByVal hwnd As Long) As Boolean
    LockWindow = LockWindowUpdate(hwnd)
End Function

Public Sub SleepWait(ByVal dwMilliseconds As Long)
    Sleep (dwMilliseconds)
End Sub
Public Function createGUID(ByRef pGuid As Variant) As Long
    createGUID = CoCreateGuid_Alt(pGuid)
End Function

Public Function StringFromGuidALT(ByRef pGuid As Variant, ByVal address As Long, ByVal max As Long) As Long
    StringFromGuidALT = StringFromGUID2_Alt(pGuid, address, max)
End Function

' remove from generalutilities to here
Public Function IsWindowIconic(ByRef ChkForm As Form) As Boolean
    IsWindowIconic = (IsIconic(ChkForm.hwnd) <> 0)
End Function

Public Function FileDialog(ShowType As Byte, Title As String, hwnd As Long, Optional InitialDir, Optional filter, Optional FileName, Optional DefExt) As String
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String, tmpStr As String, SnitialDir As String, strDefExt As String
    
    If IsMissing(filter) Then 'Default Text Filter
        sFilter = "TEXT FILE (*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Else
        sFilter = filter
    End If
    
    tmpStr = CurrentDb.Name
    If IsMissing(InitialDir) Then
        SnitialDir = GetDirectoryAndFilename(1, tmpStr)
    Else
        SnitialDir = InitialDir
    End If
    
    If IsMissing(DefExt) Then
        strDefExt = ""
    Else
        strDefExt = DefExt
    End If
    
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        If IsMissing(FileName) = False Then .lpstrFile = FileName Else .lpstrFile = ""
        .hWndOwner = hwnd
    'OpenFile.hInstance = hInstance
        .lpstrFilter = sFilter
        .nFilterIndex = 1
        .lpstrFile = .lpstrFile & String(257 - Len(.lpstrFile), 0)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = SnitialDir
        .lpstrTitle = Title
        .lpstrDefExt = strDefExt
        .flags = 0
    End With
    
    
    Select Case ShowType
    Case 0 'Show Open
        'lReturn = GetOpenFileName(OpenFile)
        ' HC 5/2010 changed to call common
        lReturn = GetFile(OpenFile)
    Case 1 'Show Save
        lReturn = GetSaveFileNameA(OpenFile)
    End Select
    
    
    If lReturn = 0 Then
        Exit Function
    Else
        sFilter = Trim(OpenFile.lpstrFile)
        tmpStr = ""
        'Trim The PHAT
        tmpStr = left(sFilter, InStr(1, sFilter, Chr(0)) - 1)
    End If
    FileDialog = tmpStr
End Function
Public Function GetFile(ByRef OpenFile As OPENFILENAME) As Long
    GetFile = GetOpenFileName(OpenFile)
End Function
Public Function GetTempDirectory(DisplayErrors As Boolean) As String
    On Error GoTo GetTempDirectoryError
    
    Dim nBufferLength As Long, lpBuffer As String, tmpStr As String
    Dim RetBufferLen As Long
    
    nBufferLength = 500
    lpBuffer = String(500, " ")
    
    RetBufferLen = GetTempPath(nBufferLength, lpBuffer)
    
    If RetBufferLen > 0 Then
        tmpStr = left(lpBuffer, RetBufferLen)
        If Right(tmpStr, 1) <> "/" And Right(tmpStr, 1) <> "\" Then
            tmpStr = tmpStr & IIf(left(tmpStr, 1) = "\" Or left(tmpStr, 1) = "/", left(tmpStr, 1), "\")
        End If
    Else
        If DisplayErrors = True Then
            MsgBox "No temp directory was returned!" & vbCrLf & vbCrLf & "'C:\TEMP\' will be used!", vbCritical, "Error getting Temp Directory"
        End If
        tmpStr = "C:\TEMP\"
    End If
    
    
    GetTempDirectory = tmpStr
    
    
GetTempDirectoryExit:
    On Error Resume Next
    Exit Function
    
GetTempDirectoryError:
    If DisplayErrors = True Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "'C:\TEMP\' will be used!", vbCritical, "Error getting Temp Directory"
    End If
    GetTempDirectory = "C:\TEMP\"
    Resume GetTempDirectoryExit
End Function
'DLC 06/09/2010 - Changed parameter lpTempFileName to ByRef to return the retrieved value
Public Function GetTempFile(ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByRef lpTempFileName As String) As Long
    GetTempFile = GetTempFileNameA(lpszPath, lpPrefixString, wUnique, lpTempFileName)
End Function

Function ChooseColor(ByVal DefaultColor As Long, Hwd As Long) As Long
    Dim CustomColors() As Byte
    Dim CC As ChooseColor
    Dim lReturn As Long
    CC.lStructSize = Len(CC)
    CC.hWndOwner = Hwd
    CC.hInstance = 0
    CC.lpCustColors = StrConv(CustomColors, vbUnicode)
    CC.flags = 0 ' CC_PREVENTFULLOPEN
    lReturn = ChooseColorAPI(CC)
    If lReturn <> 0 Then
        ChooseColor = CC.rgbResult
    Else
        ChooseColor = DefaultColor
    End If
End Function

Function OpenURL(sFile As String, Optional vArgs As Variant, Optional vShow As Variant, Optional vInitDir As Variant, Optional vVerb As Variant, _
                        Optional vhWnd As Variant) As Long
On Error Resume Next

    If IsMissing(vArgs) Then vArgs = vbNullString
    If IsMissing(vShow) Then vShow = vbNormalFocus
    If IsMissing(vInitDir) Then vInitDir = vbNullString
    If IsMissing(vVerb) Then vVerb = vbNullString
    If IsMissing(vhWnd) Then vhWnd = 0

    OpenURL = ShellExecute(vhWnd, vVerb, sFile, vArgs, vInitDir, vShow)
End Function

Function GetNetworkUserName() As String
' Returns the network login name
    Dim lngLen As Long, lngX As Long
    Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        GetNetworkUserName = left$(strUserName, lngLen - 1)
    Else
        GetNetworkUserName = vbNullString
    End If
End Function
' Clipboard functions added by Dave Brady for 1.6 Release
Function ClipBoard_GetData()
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim retval As Long

   On Error GoTo eTrap
    
   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
         
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If

   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      retval = lstrcpy(MyString, lpClipMemory)
      retval = GlobalUnlock(hClipMemory)
      
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If

OutOfHere:

   retval = CloseClipboard()
   ClipBoard_GetData = MyString
Exit Function

eTrap:
    MsgBox "Sorry, the clip board is unavailable"

End Function

Function ClipBoard_SetData(MyString As String)
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, X As Long
   
   On Error GoTo eTrap

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Function
   End If

   ' Clear the Clipboard.
   X = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If
   
eTrap:

End Function

'DLC 02/09/10 - Used to bring a window to the foreground
Public Sub SetWindowFocus(hwnd As Long)
    On Error Resume Next
    SetForegroundWindow hwnd
End Sub


Public Function NowAsUTC() As Date
    NowAsUTC = LocalTimeToUTC(Now)
End Function

Public Function LocalTimeToUTC(dteTime As Date) As Date
    Dim dteLocalFileTime As FILETIME
    Dim dteFileTime As FILETIME
    Dim dteLocalSystemTime As SYSTEMTIME
    Dim dteSystemTime As SYSTEMTIME
    With dteLocalSystemTime
        .wYear = CInt(Year(dteTime))
        .wMonth = CInt(Month(dteTime))
        .wDay = CInt(Day(dteTime))
        .wHour = CInt(Hour(dteTime))
        .wMinute = CInt(Minute(dteTime))
        .wSecond = CInt(Second(dteTime))
    End With
    Call SystemTimeToFileTime(dteLocalSystemTime, dteLocalFileTime)
    Call LocalFileTimeToFileTime(dteLocalFileTime, dteFileTime)
    Call FileTimeToSystemTime(dteFileTime, dteSystemTime)

    LocalTimeToUTC = CDate(dteSystemTime.wYear & "-" & dteSystemTime.wMonth & "-" & dteSystemTime.wDay & " " & _
                     dteSystemTime.wHour & ":" & dteSystemTime.wMinute & ":" & dteSystemTime.wSecond)
End Function

' DLC - 01/30/12
'    - Returns the number of seconds since the start of the current month
'    - Similar to Timer() but accurate to 1ms whereas Timer() is accurate to about 15ms
Public Function GetTimer() As Double
    Dim dteSystemTime As SYSTEMTIME
    Call GetSystemTime(dteSystemTime)
    GetTimer = (((CDbl((dteSystemTime.wDay * 24) + dteSystemTime.wHour) * 60) + CDbl(dteSystemTime.wMinute)) * 60) + CDbl(dteSystemTime.wSecond) + (dteSystemTime.wMilliseconds / 1000)
End Function

Public Function ScreenWidth() As Integer
    'SA 03/22/2012 - New function for screen width
    ScreenWidth = GetSystemMetrics(0)
End Function

Public Function ScreenHeight() As Integer
    'SA 03/22/2012 - New function for screen height
    ScreenHeight = GetSystemMetrics(1)
End Function




Function GetDirectory(Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim Path As String
    Dim r As Long, X As Long, pos As Integer
     
     '   Root folder = Desktop
    bInfo.pidlRoot = 0&
     
     '   Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
    Else
        bInfo.lpszTitle = Msg
    End If
     
     '   Type of directory to return
    bInfo.ulFlags = &H1
     
     '   Display the dialog
    X = SHBrowseForFolder(bInfo)
     
     '   Parse the result
    Path = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal Path)
    If r Then
        pos = InStr(Path, Chr$(0))
        GetDirectory = left(Path, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function