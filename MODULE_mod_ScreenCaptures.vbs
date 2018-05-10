Option Compare Database
Option Explicit


'' Last Modified: 05/26/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''   Capture the active window as a screen grab, then save it to a file
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/26/2015  - KD: Added

''
'' AUTHOR
''  =====================================
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################



Private Const ClassName As String = "mod_ScreenCaptures"

'Declare a UDT to store a GUID for the IPicture OLE Interface
 
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
 
 
'Declare a UDT to store the bitmap information
Private Type uPicDesc
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type
 
 
 
'Windows API Function Declarations
#If Win64 = 1 And VBA7 = 1 Then
    
    'Does the clipboard contain a bitmap/metafile?
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
 
    'Open the clipboard to read
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
 
    'Get a pointer to the bitmap/metafile
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
 
    'Close the clipboard
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
 
    'Convert the handle into an OLE IPicture interface.
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
 
    'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
    Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
 
    'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
    Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 
    'Uses the Keyboard simulation
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 
#Else
 
    'Does the clipboard contain a bitmap/metafile?
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
 
    'Open the clipboard to read
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
 
    'Get a pointer to the bitmap/metafile
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
 
    'Close the clipboard
    Private Declare Function CloseClipboard Lib "user32" () As Long
 
    'Convert the handle into an OLE IPicture interface.
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
     
    'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
    Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
 
    'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
    Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
 
    'Uses the Keyboard simulation
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 
#End If
   
 
'The API format types we're interested in
Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4
 
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
 
 
 
' Subroutine    : AltPrintScreen
' Purpose       : Capture the Active window, and places on the Clipboard.
Public Function AltPrintScreen() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".AltPrintScreen"
    
    keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
 
 
 
' Subroutine    : PastePicture
' Purpose       : Get a Picture object showing whatever's on the clipboard.
Public Function PastePicture() As IPicture
On Error GoTo Block_Err
Dim strProcName As String
Dim h As Long
Dim hPtr As Long
Dim hPal As Long
Dim lPicType As Long
Dim hCopy As Long
 
    strProcName = ClassName & ".PastePicture"
 
    'Check if the clipboard contains the required format
    If IsClipboardFormatAvailable(CF_BITMAP) Then
        'Get access to the clipboard
        h = OpenClipboard(0&)
        
        If h > 0 Then
            'Get a handle to the image data
            hPtr = GetClipboardData(CF_BITMAP)
            
            hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
            
            'Release the clipboard to other programs
            h = CloseClipboard
            
            'If we got a handle to the image, convert it into a Picture object and return it
            If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, CF_BITMAP)
        End If
    End If
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
 
 
 
' Subroutine    : CreatePicture
' Purpose       : Converts a image (and palette) handle into a Picture object.
' NOTE          : Requires a reference to the "OLE Automation" type library
Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
On Error GoTo Block_Err
Dim strProcName As String
Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
    'OLE Picture types
Const PICTYPE_BITMAP = 1
Const PICTYPE_ENHMETAFILE = 4


    strProcName = ClassName & ".CreatePicture"
 
    ' Create the Interface GUID (for the IPicture interface)
 
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
 
    ' Fill uPicInfo with necessary parts.
    With uPicInfo
        .Size = Len(uPicInfo) ' Length of structure.
        .Type = PICTYPE_BITMAP ' Type of Picture
        .hPic = hPic ' Handle to image.
        .hPal = hPal ' Handle to palette (if bitmap).
    End With
 
    
    ' Create the Picture object.
    r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)
 
    ' If an error occurred, show the description
    If r <> 0 Then Debug.Print "Create Picture: " & OLEError(r)
 
    ' Return the new Picture object.
    Set CreatePicture = IPic
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
 
 
' Subroutine    : OLEError
' Purpose       : Gets the message text for standard OLE errors
Private Function OLEError(lErrNum As Long) As String
    'OLECreatePictureIndirect return values
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".OLEError"

    Const E_ABORT = &H80004004
    Const E_ACCESSDENIED = &H80070005
    Const E_FAIL = &H80004005
    Const E_HANDLE = &H80070006
    Const E_INVALIDARG = &H80070057
    Const E_NOINTERFACE = &H80004002
    Const E_NOTIMPL = &H80004001
    Const E_OUTOFMEMORY = &H8007000E
    Const E_POINTER = &H80004003
    Const E_UNEXPECTED = &H8000FFFF
    Const S_OK = &H0
 
    Select Case lErrNum
        Case E_ABORT
            OLEError = " Aborted"
        Case E_ACCESSDENIED
            OLEError = " Access Denied"
        Case E_FAIL
            OLEError = " General Failure"
        Case E_HANDLE
            OLEError = " Bad/Missing Handle"
        Case E_INVALIDARG
            OLEError = " Invalid Argument"
        Case E_NOINTERFACE
            OLEError = " No Interface"
        Case E_NOTIMPL
            OLEError = " Not Implemented"
        Case E_OUTOFMEMORY
            OLEError = " Out of Memory"
        Case E_POINTER
            OLEError = " Invalid Pointer"
        Case E_UNEXPECTED
            OLEError = " Unknown Error"
        Case S_OK
            OLEError = " Success!"
    End Select
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
 
 
' Routine   : SaveClip2Bit
' Purpose   : Saves Picture object to desired location.
' Arguments : Path to save the file
Public Function SaveClip2Bit(sSavePath As String) As Boolean
Dim strProcName As String
On Error GoTo Block_Err:
 
    strProcName = ClassName & ".SaveClip2Bit"
 
    AltPrintScreen
    PauseSecs (2)
    SavePicture PastePicture, sSavePath

    PauseSecs 1
    
    SaveClip2Bit = FileExists(sSavePath)

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
 
 
 
' Routine   : Pause
' Purpose   : Gives a short interval for proper image capture.
' Arguments : Seconds to wait.
Public Function PauseSecs(iNumberOfSeconds As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sgPauseTime As Single
Dim sgStart As Single

    strProcName = ClassName & ".Pause"
 
    sgPauseTime = iNumberOfSeconds
 
    sgStart = Timer
 
    Do While Timer < sgStart + sgPauseTime
        DoEvents
    Loop
    PauseSecs = True
 
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function SaveActiveWindowToFile(sFilePath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".SaveActiveWindowToFile"
    
    Sleep 20000
    
    AltPrintScreen
    
    SaveActiveWindowToFile = SaveClip2Bit(sFilePath)
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function