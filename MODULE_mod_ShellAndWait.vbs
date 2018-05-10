Option Compare Database
Option Explicit

''' *************************************************************************
''' Module Constant Declaractions Follow
''' *************************************************************************
''' Constant for the dwDesiredAccess parameter of the OpenProcess API function.
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
''' Constant for the lpExitCode parameter of the GetExitCodeProcess API function.
Private Const STILL_ACTIVE As Long = &H103


''' *************************************************************************
''' Module Variable Declaractions Follow
''' *************************************************************************
''' It's critical for the shell and wait procedure to trap for errors, but I
''' didn't want that to distract from the example, so I'm employing a very
''' rudimentary error handling scheme here. This variable is used to pass error
''' messages between procedures.
Public gszErrMsg As String


''' *************************************************************************
''' Module DLL Declaractions Follow
''' *************************************************************************
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long


Public Function bShellAndWait(ByVal szCommandLine As String, Optional ByVal iWindowState As Integer = vbHide) As Boolean

    Dim lTaskID As Long
    Dim lProcess As Long
    Dim lExitCode As Long
    Dim lResult As Long
    
    On Error GoTo ErrorHandler

    ''' Run the Shell function.
    lTaskID = Shell(szCommandLine, iWindowState)
    
    ''' Check for errors.
    If lTaskID = 0 Then Err.Raise 9999, , "Shell function error."
    
    ''' Get the process handle from the task ID returned by Shell.
    lProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0&, lTaskID)
    
    ''' Check for errors.
    If lProcess = 0 Then Err.Raise 9999, , "Unable to open Shell process handle."
    
    ''' Loop while the shelled process is still running.
    Do
        ''' lExitCode will be set to STILL_ACTIVE as long as the shelled process is running.
        lResult = GetExitCodeProcess(lProcess, lExitCode)
        DoEvents
    Loop While lExitCode = STILL_ACTIVE
    
    bShellAndWait = True
    Exit Function
    
ErrorHandler:
    gszErrMsg = Err.Description
    bShellAndWait = False
End Function