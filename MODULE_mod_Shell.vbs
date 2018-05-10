
Option Compare Database
Option Explicit

'Public Declare Function GetEnvironmentVariable& Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long)
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 '                                     (ByVal hWnd As Long, ByVal lpOperation As String, _
 '                                      ByVal lpFile As String, ByVal lpParameters As String, _
 '                                      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
'Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" _
 '                                               (ByVal hWnd As Long, ByVal lpOperation As String, _
 '                                                ByVal lpFile As String, lpParameters As Any, _
 '                                                lpDirectory As Any, ByVal nShowCmd As Long) As Long
'
'Public Declare Function apiGetComputerName Lib "kernel32" Alias _
 '                                           "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'
'Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'
'
'Public Enum EShellShowConstants
'    essSW_HIDE = 0
'    essSW_MAXIMIZE = 3
'    essSW_MINIMIZE = 6
'    essSW_SHOWMAXIMIZED = 3
'    essSW_SHOWMINIMIZED = 2
'    essSW_SHOWNORMAL = 1
'    essSW_SHOWNOACTIVATE = 4
'    essSW_SHOWNA = 8
'    essSW_SHOWMINNOACTIVE = 7
'    essSW_SHOWDEFAULT = 10
'    essSW_RESTORE = 9
'    essSW_SHOW = 5
'End Enum
'
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'
'Private Const ERROR_FILE_NOT_FOUND = 2&
'Private Const ERROR_PATH_NOT_FOUND = 3&
'Private Const ERROR_BAD_FORMAT = 11&
'Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
'Private Const SE_ERR_ASSOCINCOMPLETE = 27
'Private Const SE_ERR_DDEBUSY = 30
'Private Const SE_ERR_DDEFAIL = 29
'Private Const SE_ERR_DDETIMEOUT = 28
'Private Const SE_ERR_DLLNOTFOUND = 32
'Private Const SE_ERR_FNF = 2                ' file not found
'Private Const SE_ERR_NOASSOC = 31
'Private Const SE_ERR_PNF = 3                ' path not found
'Private Const SE_ERR_OOM = 8                ' out of memory
'Private Const SE_ERR_SHARE = 26
Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&


'
'Public Function ShellEx( _
 '       ByVal sFile As String, _
 '       Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
 '       Optional ByVal sParameters As String = "", _
 '       Optional ByVal sDefaultDir As String = "", _
 '       Optional sOperation As String = "open", _
 '       Optional Owner As Long = 0 _
 '     ) As Boolean
'    Dim lR As Long
'    Dim lErr As Long, sErr As Long
'    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
'        eShowCmd = 0
'    End If
'    On Error Resume Next
'    If (sParameters = "") And (sDefaultDir = "") Then
'        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
'    Else
'        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
'    End If
'    If (lR < 0) Or (lR > 32) Then
'        ShellEx = True
'    Else
'        ' raise an appropriate error:
'        lErr = vbObjectError + 1048 + lR
'        Select Case lR
'        Case 0
'            lErr = 7: sErr = "Out of memory"
'        Case ERROR_FILE_NOT_FOUND
'            lErr = 53: sErr = "File not found"
'        Case ERROR_PATH_NOT_FOUND
'            lErr = 76: sErr = "Path not found"
'        Case ERROR_BAD_FORMAT
'            sErr = "The executable file is invalid or corrupt"
'        Case SE_ERR_ACCESSDENIED
'            lErr = 75: sErr = "Path/file access error"
'        Case SE_ERR_ASSOCINCOMPLETE
'            sErr = "This file type does not have a valid file association."
'        Case SE_ERR_DDEBUSY
'            lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
'        Case SE_ERR_DDEFAIL
'            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
'        Case SE_ERR_DDETIMEOUT
'            lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
'        Case SE_ERR_DLLNOTFOUND
'            lErr = 48: sErr = "The specified dynamic-link library was not found."
'        Case SE_ERR_FNF
'            lErr = 53: sErr = "File not found"
'        Case SE_ERR_NOASSOC
'            sErr = "No application is associated with this file type."
'        Case SE_ERR_OOM
'            lErr = 7: sErr = "Out of memory"
'        Case SE_ERR_PNF
'            lErr = 76: sErr = "Path not found"
'        Case SE_ERR_SHARE
'            lErr = 75: sErr = "A sharing violation occurred."
'        Case Else
'            sErr = "An error occurred occurred whilst trying to open or print the selected file."
'        End Select
'
'        Err.Raise lErr, , "Shell error", sErr
'        ShellEx = False
'    End If
'
'End Function


'
'
Public Function ShellWait(pathname As String, Optional WindowStyle As Long) As Long
    On Error GoTo Err_handler

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long

    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    ' Start the shelled application:
    ret& = CreateProcessA(0&, pathname, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)

EXIT_HERE:
    Exit Function
Err_handler:
    MsgBox Err.Description, vbExclamation, "E R R O R"
    Resume EXIT_HERE

End Function