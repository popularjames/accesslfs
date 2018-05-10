Option Compare Database
Option Explicit


'' Last Modified: 05/26/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''   Some basic windows API's that are used a lot..
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/26/2015  - KD: Added ability to copy the active window to the clipboard - but then I decided to move it
''                  to it's own module
''  - 05/23/2012  - KD: added SW_MAXIMIZE and made that the choice for activating windows
''  - 04/27/2012  - KD: Added find window and some others

''
'' AUTHOR
''  =====================================
''  Kevin Dearing (Well, not really, someone at Microsoft, I just
''  glue it together so I can use it in my own code!)
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################


Private Const ClassName As String = "mod_Windows_APIs"


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_MAXIMIZE  As Long = 3
Private Const SW_RESTORE As Long = 9

'' Following conctants are used in the dictionary which is set in SetupClassDictionary
Private Const gcClassnameMSWord  As String = "OpusApp"
Private Const gcClassnameMSExcel  As String = "XLMAIN"
Private Const gcClassnameMSIExplorer As String = "IEFrame"
Private Const gcClassnameMSVBasic As String = "wndclass_desked_gsk"
Private Const gcClassnameNotePad As String = "Notepad"
Private Const gcClassnameMyVBApp As String = "ThunderForm"
Private Const gcClassnameMSOutlook As String = "rctrl_renwnd32"
Private cdctClassesByAppName As Scripting.Dictionary


Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long


'' 20120416 KD Added
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS, sets dialog title
End Type


'**** jc copied these from JAC imported from CCA_FILE_API_Functions


'' 20120416 KD Added - extended.. Using File Copy / Delete, rename as it tends to be quicker than FSO
Private Const FO_COPY = &H2 ' Copy File/Folder
Private Const FO_DELETE = &H3 ' Delete File/Folder
Private Const FO_MOVE = &H1 ' Move File/Folder
Private Const FO_RENAME = &H4 ' Rename File/Folder
Private Const FOF_ALLOWUNDO = &H40 ' Allow to undo rename, delete ie sends to recycle bin
Private Const FOF_FILESONLY = &H80  ' Only allow files
Private Const FOF_NOCONFIRMATION = &H10  ' No File Delete or Overwrite Confirmation Dialog
Private Const FOF_SILENT = &H4 ' No copy/move dialog
Private Const FOF_SIMPLEPROGRESS = &H100 ' Does not display file names

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long


'*********

'***** For Function ProcessTerminate JS 07/01/2013
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'***** For Function ProcessTerminate JS 07/01/2013

'***** For function CheckForProcByExe JS 07/01/2013
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal strModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
'***** For function CheckForProcByExe JS 07/01/2013

Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
Alias "GetDiskFreeSpaceExA" _
(ByVal lpcurRootPathName As String, _
lpFreeBytesAvailableToCaller As Currency, _
lpTotalNumberOfBytes As Currency, _
lpTotalNumberOfFreeBytes As Currency) As Long






Public Function Wait(WaitTime As Integer) As Boolean
    ' routine to suspend application for x seconds
    ' WaitTime specifies number of 1/10 of seconds to wait
    Sleep WaitTime * 100
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function IsAppRunning(Optional sClassName As String, Optional sWindowTitle As String, Optional sAppName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim lRet As Long

    strProcName = ClassName & ".FunctionName"
    If sClassName <> "" Then
        lRet = FindWindow(vbNullString, sWindowTitle)
    ElseIf sWindowTitle <> "" Then
        lRet = FindWindow(sClassName, vbNullString)
    Else
        lRet = FindWindow(GetClassNameFromAppName(sAppName), vbNullString)
    End If

    If lRet <> 0 Then IsAppRunning = True

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetClassNameFromTitle(sWindowTitle As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim hwnd As Long, lpClassName As String
Dim nMaxCount As Long, lResult As Long

    strProcName = ClassName & ".GetClassNameFromTitle"

    nMaxCount = 256
    lpClassName = Space(nMaxCount)
   
    hwnd = FindWindow(vbNullString, sWindowTitle)
    
    If hwnd = 0 Then
        GetClassNameFromTitle = "Couldn't find the window."
    Else
        lResult = GetClassName(hwnd, lpClassName, nMaxCount)
        GetClassNameFromTitle = "Window: " + sWindowTitle + Chr$(13) + Chr$(10) + "Classname: " + left$(lpClassName, lResult)
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ActivateWindowClass(psClassname As String) As Boolean
Dim hwnd As Long
    hwnd = FindWindow(psClassname, vbNullString)
    If hwnd > 0 Then
        ShowWindow hwnd, SW_MAXIMIZE    '   SW_SHOWNORMAL
        ActivateWindowClass = True
    End If
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ActivateApplicationWindow(Optional sClassName As String, Optional sWindowTitle As String, Optional sAppName As String) As Boolean
Dim hwnd As Long

    If sClassName <> "" Then
        hwnd = FindWindow(sClassName, vbNullString)
    ElseIf sWindowTitle <> "" Then
        hwnd = FindWindow(vbNullString, sWindowTitle)
    ElseIf sAppName <> "" Then
        hwnd = FindWindow(GetClassNameFromAppName(sAppName), vbNullString)
    Else
    
    End If

    If hwnd > 0 Then
        ShowWindow hwnd, SW_MAXIMIZE
        ActivateApplicationWindow = True
    End If
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetClassNameFromAppName(ByVal strAppName As String) As String

    If cdctClassesByAppName Is Nothing Then
        Call SetupClassDictionary
    End If
    strAppName = UCase(strAppName)
    
    If cdctClassesByAppName.Exists(strAppName) = True Then
        GetClassNameFromAppName = cdctClassesByAppName.Item(strAppName)
    End If
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub SetupClassDictionary()
    Set cdctClassesByAppName = New Scripting.Dictionary
    With cdctClassesByAppName
        .Add "WORD", gcClassnameMSWord
        .Add "MICROSOFT WORD", gcClassnameMSWord
        .Add "MSWORD", gcClassnameMSWord
        .Add "EXCEL", gcClassnameMSExcel
        .Add "MSEXCEL", gcClassnameMSExcel
        .Add "EXPLORER", gcClassnameMSIExplorer
        .Add "IEXPLORER", gcClassnameMSIExplorer
        .Add "MSVBASIC", gcClassnameMSVBasic
        .Add "VBASIC", gcClassnameMSVBasic
        .Add "NOTEPAD", gcClassnameNotePad
        .Add "OUTLOOK", gcClassnameMSOutlook
        .Add "MSOUTLOOK", gcClassnameMSOutlook
 
    End With
End Sub





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' If the path passed is a valid UNC path, returns true
'* 12/6/12 JC moved from CCA_FILE_API_Functions

Public Function ParentFolderPath(strFullFilePath As String) As String
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".ParentFolderPath"
    Call PathInfoFromPath(strFullFilePath, , ParentFolderPath)
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Given a path (UNC or mapped) will break the details up into sections
''' Note, strFileName will NOT have the . & Extension
'''
'* 12/6/12 JC moved from CCA_FILE_API_Functions

Public Function PathInfoFromPath(ByVal strFullPath As String, Optional ByRef strFileName As String, _
    Optional ByRef strParentFolder As String, Optional ByRef strFileExtension As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFile As Scripting.file
Dim oRegExp As RegExp
Dim oMatches As MatchCollection
Dim oMatch As Match
    
    strProcName = ClassName & ".PathInforFromPath"
    
    strFullPath = TrimNull(strFullPath)
    Set oFso = New Scripting.FileSystemObject
    If oFso.FileExists(strFullPath) = True Then
        Set oFile = oFso.GetFile(strFullPath)
    
        strParentFolder = QualifyFldrPath(oFile.ParentFolder)
        strFileExtension = oFso.GetExtensionName(strFullPath)
    
        strFileName = Replace(oFile.Name, "." & strFileExtension, "", 1, 1, vbTextCompare)
    Else
        ' File doesn't exist yet, so let's parse the string ourselves!
        Set oRegExp = New RegExp
        With oRegExp
            .Global = False
            .IgnoreCase = True
            .Pattern = "^(.*?\\*)([^\\]+)\\*$"
        End With
        
        Set oMatches = oRegExp.Execute(strFullPath)
        If oMatches.Count > 0 Then
            Set oMatch = oMatches.Item(0)
            strParentFolder = QualifyFldrPath(oMatch.SubMatches(0))
            strFileName = oMatch.SubMatches(1)
            
            oRegExp.Pattern = "^.+?(\.[^\\]+)$"
            strFileExtension = oRegExp.Replace(strFullPath, "$1")
            ' to match, we get rid of the period if found:
            strFileExtension = Replace(strFileExtension, ".", "")
            If strFileExtension = strFullPath Then strFileExtension = ""

        End If
        
    End If

    PathInfoFromPath = Len(strParentFolder)

Block_Exit:
    Set oFile = Nothing
    Set oFso = Nothing
    Set oRegExp = Nothing
    Set oMatches = Nothing
    Set oMatch = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    PathInfoFromPath = 0
End Function

'* 12/6/12 JC moved from CCA_FILE_API_Functions

Public Function TrimNull(strStart As String) As String
    TrimNull = left$(strStart, lstrlenW(StrPtr(strStart)))
End Function


'* 12/6/12 JC moved from CCA_FILE_API_Functions

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Insures the path passed ends with a \
Public Function QualifyFldrPath(sPath As String) As String
    'add trailing slash if required
    If Right$(sPath, 1) <> "\" Then
        QualifyFldrPath = sPath & "\"
    Else
        QualifyFldrPath = sPath
    End If
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' 12/5/12 JAC imprted from CCA_FILE_API_Functions
Public Function CopyFile(strCurrentPath As String, strdestinationpath As String, Optional blnConfirmPrompt As Boolean = True, Optional ErrMsg As String = "") As Boolean
Dim op As SHFILEOPSTRUCT
Dim strProcName As String
Dim sWorkPath As String
Dim sWildCardReplacement As String
Dim sShortName As String
Dim oFso As Scripting.FileSystemObject
Dim sFldr As String
On Error GoTo Block_Err
Dim lRet As Long

    strProcName = ClassName & ".CopyFile"
    CopyFile = True
    
    Set oFso = New Scripting.FileSystemObject
    
    sFldr = oFso.GetParentFolderName(strdestinationpath)
    If oFso.FolderExists(sFldr) = False Then
        CreateFolders sFldr
    End If
    
        ' --- delete it if its there already
    DeleteFile strdestinationpath, blnConfirmPrompt
    
    CopyFile = True
    With op
        .wFunc = FO_COPY
        .pFrom = strCurrentPath
        .pTo = strdestinationpath
        If blnConfirmPrompt = False Then
            .fFlags = FOF_SILENT + FOF_NOCONFIRMATION    '   FOF_SILENT    '   FOF_NOCONFIRMATION & FOF_SILENT
        End If
    End With
    lRet = SHFileOperation(op)

    Select Case lRet
    Case 0
        CopyFile = True
    Case Else
        CopyFile = False
    End Select
    
Block_Exit:
    Set oFso = Nothing
    Exit Function

Block_Err:
    ReportError Err, strProcName, strCurrentPath & " to " & strdestinationpath
    CopyFile = False
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' 12/6/12 JAC imprted from CCA_FILE_API_Functions
Public Function DeleteFile(strCurrentPath As String, Optional blnConfirmPrompt As Boolean = True) As Boolean
Dim op As SHFILEOPSTRUCT
Dim strProcName As String
Dim lRet As Long
On Error GoTo Block_Err

    strProcName = ClassName & ".DeleteFile"
    
    DeleteFile = True
    If FileExists(strCurrentPath) = True Then
        With op
            .wFunc = FO_DELETE
            .pFrom = strCurrentPath
            If blnConfirmPrompt = False Then
                .fFlags = FOF_NOCONFIRMATION
            End If
        End With
        lRet = SHFileOperation(op)
        
        Select Case lRet
        Case 0
            DeleteFile = True
        Case Else
            DeleteFile = False
        End Select
    End If

Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName, strCurrentPath
    DeleteFile = False
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
''' 12/6/12 JAC imprted from CCA_FILE_API_Functions


Public Function RenameFile(strCurrentPath As String, strNewPath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim op As SHFILEOPSTRUCT
Dim lRet As Long

    strProcName = ClassName & ".RenameFile"
    
    RenameFile = False
    
    With op
        .wFunc = FO_RENAME ' Set function
        .pTo = strNewPath ' Set new path
        .pFrom = strCurrentPath ' Set current path
        .fFlags = FOF_SILENT
    End With
    ' Perform operation
    lRet = SHFileOperation(op)
    
    Select Case lRet
    Case 0
        RenameFile = True
    Case Else
        RenameFile = False
    End Select
    
Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName, strCurrentPath
    RenameFile = False
    GoTo Block_Exit
End Function


''' 12/6/12 JAC imprted from CCA_FILE_API_Functions

Public Function MoveFile(ByVal SourceFile As String, ByVal destinationFile As String, Optional ByVal Override As Boolean = False, Optional ErrMsg As String = "") As Boolean
    Dim fso As New FileSystemObject
    
    On Error GoTo Err_handler
    
    Call fso.CopyFile(SourceFile, destinationFile, Override)
    If fso.FileExists(destinationFile) Then
        fso.DeleteFile (SourceFile)
        If fso.FileExists(SourceFile) Then
            Err.Raise vbObjectError + 513, "Can not delete source file " & SourceFile
        End If
    Else
        Err.Raise vbObjectError + 513, "Can not copy file " & SourceFile & " to " & destinationFile
    End If
    
    MoveFile = True
    
Exit_Function:
    Set fso = Nothing
    Exit Function
    
Err_handler:
    MoveFile = False
    ErrMsg = "ERROR#: [" & CStr(Err.Number) & "].   ERROR DESCRIPTION: [" & Err.Description & "]"
    Resume Exit_Function
End Function



''' 12/6/12 JAC imprted from CCA_FILE_API_Functions

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' If the path passed is a valid UNC path, returns true
Public Function FileExtension(strFullFilePath As String) As String
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".FileExtension"
    Call PathInfoFromPath(strFullFilePath, , , FileExtension)
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


 

Function ProcessTerminate(Optional lprocessid As Long, Optional lHwndWindow As Long) As Boolean
    Dim lhwndProcess As Long
    Dim lExitCode As Long
    Dim lRetVal As Long
    Dim lhThisProc As Long
    Dim lhTokenHandle As Long
    Dim tLuid As LUID
    Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINATE = &H1
    Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Const SE_PRIVILEGE_ENABLED = &H2

    On Error Resume Next
    If lHwndWindow Then
        'Get the process ID from the window handle
        lRetVal = GetWindowThreadProcessId(lHwndWindow, lprocessid)
    End If
    
    If lprocessid Then
        'Give Kill permissions to this process
        lhThisProc = GetCurrentProcess
        
        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        'Set the number of privileges to be change
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        'Enable the kill privilege in the access token of this process
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

        'Open the process to kill
        lhwndProcess = OpenProcess(PROCESS_TERMINATE, 0, lprocessid)
    
        If lhwndProcess Then
            'Obtained process handle, kill the process
            ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
            Call CloseHandle(lhwndProcess)
        End If
    End If
    On Error GoTo 0
End Function

Public Function CheckForProcByExe(pEXEName As String) As Long
On Error Resume Next
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim lProcessIDs() As Long
Dim cbNeeded2 As Long
Dim lNumElements2 As Long
Dim lModules(1 To 200) As Long
Dim lRet As Long
Dim strModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim lProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(lProcessIDs(1), cb, cbNeeded)
Loop
NumElements = cbNeeded / 4
For i = 1 To NumElements
    'Get a handle to the Process
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcessIDs(i))
    'Got a Process handle
    If hProcess <> 0 Then
        'Get an array of the module handles for the specified
        'process
        lRet = EnumProcessModules(hProcess, lModules(1), 200, cbNeeded2)
        'If the Module Array is retrieved, Get the ModuleFileName
        If lRet <> 0 Then
            strModuleName = Space(MAX_PATH)
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, lModules(1), strModuleName, nSize)
            strModuleName = left(strModuleName, lRet)
            'Check for the client application running
            If InStr(UCase(strModuleName), UCase(pEXEName)) Then
                CheckForProcByExe = lProcessIDs(i)
                Exit Function
            Else
                CheckForProcByExe = 0
            End If
        End If
    End If
    'Close the handle to the process
    lRet = CloseHandle(hProcess)
Next
End Function



Function CloseByProcessHandle(ViewerProcessHandle As Long) As Boolean
    
    Dim NumberOfTries As Integer
    Dim ProcessViewer As Long
    Dim TerminateResult As Boolean
    
    Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINATE = &H1
    
    If ViewerProcessHandle = 0 Then
        CloseByProcessHandle = True
        Exit Function
    End If

    NumberOfTries = 0 'should not be that many IrfanView windows opened, right?
    TerminateResult = False
    
    Do While Not TerminateResult And NumberOfTries < 10
        TerminateResult = ProcessTerminate(, ViewerProcessHandle)
        
'        If TerminateResult = False Then
'            MsgBox "Could not close all instances of IrfanView automatically." & vbNewLine & vbNewLine & _
'                    "You must close them yourself manually, then please try again.", vbExclamation
'            CloseAllIrfanView = False
'            Exit Function
'        End If
        NumberOfTries = NumberOfTries + 1
    Loop
    
    If NumberOfTries < 10 Then
        CloseByProcessHandle = True
        Exit Function
    End If
    
ErrorClosing:
    CloseByProcessHandle = False


End Function





Function CloseByProcessTitle(Title As String) As Boolean
    Dim hWndThis As Long
    Dim Class As String
    Dim ProcessWasFound As Boolean
    
    If Nz(Title, "") = "" Then
        CloseByProcessTitle = True
        Exit Function
    End If
    
    Class = "*"
    Title = "*" & Title & "*"
    CloseByProcessTitle = False
    ProcessWasFound = False
    
    hWndThis = FindWindow(vbNullString, vbNullString)
    While hWndThis
        Dim sTitle As String, sClass As String
        sTitle = Space$(255)
        sTitle = left$(sTitle, GetWindowText(hWndThis, sTitle, Len(sTitle)))
        sClass = Space$(255)
        sClass = left$(sClass, GetClassName(hWndThis, sClass, Len(sClass)))
        If sTitle Like Title And sClass Like Class Then
            Debug.Print sTitle, sClass
            ProcessWasFound = True
            CloseByProcessTitle = CloseByProcessHandle(hWndThis)
        End If
        hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
    Wend

    If Not ProcessWasFound Then
        CloseByProcessTitle = True
    End If
    
End Function

' ?FindExecutableForFile("\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\NIRF_C0834_0001.pdf")
Public Function FindExecutableForFile(strDataFile As String, Optional strDir As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim lngApp As Long
Dim strApp As String

    strProcName = ClassName & ".FindExecutableForFile"

    strApp = Space(260)
    lngApp = FindExecutable(strDataFile, strDir, strApp)
    
    If lngApp > 32 Then
        FindExecutableForFile = strApp
    Else
        FindExecutableForFile = "No matching application."
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function