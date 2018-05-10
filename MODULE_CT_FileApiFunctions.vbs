Option Compare Database
Option Explicit

'**********************************
'**  Function Declarations:
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare PtrSafe Function CreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long)
Private Declare PtrSafe Function GetFileInformationByHandle& Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION)
Private Declare PtrSafe Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Private Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Declare PtrSafe Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "Shell32" _
   (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare PtrSafe Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare PtrSafe Function GetSystemWindowsDirectory Lib "kernel32" Alias "GetSystemWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal uSize As Long) As Long

Private Const MAX_PATH = 260

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type BY_HANDLE_FILE_INFORMATION
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        dwVolumeSerialNumber As Long
        nFileSizeHigh As Long
        nFileSizeLow As Long
        nNumberOfLinks As Long
        nFileIndexHigh As Long
        nFileIndexLow As Long
End Type

Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const SCS_32BIT_BINARY& = 0
Private Const SCS_DOS_BINARY& = 1
Private Const SCS_OS216_BINARY& = 5
Private Const SCS_PIF_BINARY& = 3
Private Const SCS_POSIX_BINARY& = 4
Private Const SCS_WOW_BINARY& = 2
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
    
Public Sub ShellExe(FileName As String)
On Error GoTo ErrorHappened
    ShellExecute GetForegroundWindow, "Open", FileName, "", "", 1
    
ExitNow:
    On Error Resume Next
    Exit Sub
    
ErrorHappened:
    MsgBox Err.Description
    Resume ExitNow
    Resume
End Sub

Public Function getWinDir() As String
On Error Resume Next
Dim buffer$, windir$
Dim ret%

    windir$ = ""
    buffer$ = Space(255)
    ret% = GetWindowsDirectory(buffer$, 255)
    windir$ = left$(buffer$, ret%)
    
    getWinDir = windir$
End Function

Public Function GetSystemPath()
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath = left(strFolder, InStr(strFolder, _
    Chr(0)) - 1)
Else
    GetSystemPath = ""
End If
End Function

Public Function GetSpecialFolderLocation(CSIDL As Long) As String
   Dim sPath As String
   Dim pidl As Long
   
  'fill the idl structure with the specified folder item
   If SHGetSpecialFolderLocation(0, CSIDL, pidl) = 0 Then
     
     'if the pidl is returned, initialize
     'and get the path from the id list
      sPath = Space$(MAX_PATH)
      
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then

        'return the path
         GetSpecialFolderLocation = left(sPath, InStr(sPath, Chr$(0)) - 1)
         
      End If
     'free the pidl
      Call CoTaskMemFree(pidl)
    End If
End Function


Public Function CCA_GetFileSize(FileName As String) As Double
' DLC - 01/18/10
' New version to correctly return the size of the specified file.
On Error GoTo ErrorHandler
    Dim dl As Long, hFile As Long
    Dim myFileInfo As BY_HANDLE_FILE_INFORMATION
    Dim ReturnValue As Double
    Dim Power As Double
    Dim str As String
    Dim index As Integer
    'Get a handle to the file
    hFile = CreateFile(FileName, 0, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
    'Get the file info
    dl = GetFileInformationByHandle&(hFile, myFileInfo)
        With myFileInfo
            'If the File is over 4GB, the .nFileSizeHigh will contain a value
            ReturnValue = .nFileSizeHigh * (2 ^ 32)
            If .nFileSizeLow >= 0 Then
                ReturnValue = ReturnValue + .nFileSizeLow
            Else
               ' DLC
               ' If .nFileSizeLow is negative it is because an unsigned long is stored in a signed long
               ' variable. To convert the signed long to a double it must first be converted to Octal.
               ' It can then be decoded by multiplying each digit with the corresponding power of 8.
               ' This is to work around the fact that VBA does not support the unsigned long integers
               ' returned by the windows API call.
               str = Oct(.nFileSizeLow)
               Power = 1
               For index = Len(str) To 1 Step -1
                   ReturnValue = ReturnValue + CDbl(Mid(str, index, 1)) * Power
                   Power = Power * 8
               Next index
            End If
        End With
Error_Exit:
    On Error Resume Next
    dl = CloseHandle(hFile)
    CCA_GetFileSize = ReturnValue
    Exit Function
ErrorHandler:
    ReturnValue = 0
    Resume Error_Exit
End Function

'Public Function CCA_GetFileSize(FileName As String) As Double
'On Error Resume Next
'
'Dim dl As Long, hFile As Long
'Dim myFileInfo As BY_HANDLE_FILE_INFORMATION
'
''Give the Function a value just in case it bombs
'CCA_GetFileSize = 2147483647
'
''Get a handle to the file
'hFile = CreateFile(FileName, 0, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
'
''Get the file info
'dl = GetFileInformationByHandle&(hFile, myFileInfo)
'    With myFileInfo
'        If .nFileSizeLow < 0 Then 'Greater than 1 gig
'            CCA_GetFileSize = (2147483647 / 10) + (.nFileSizeLow / 10)
'        Else
'            CCA_GetFileSize = (.nFileSizeLow / 10)
'        End If
'    End With
''Close the Hande to the File
'dl = CloseHandle(hFile)
'End Function