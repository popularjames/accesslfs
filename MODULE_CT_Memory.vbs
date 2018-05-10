'---------------------------------------------------------------------------------------
' Module    : CT_Memory
' Author    : SA
' Date      : 11/15/2012
' Purpose   : Provides information about memory being used by Access
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Type PROCESS_MEMORY_COUNTERS
   cb                         As Long
   PageFaultCount             As Long
   PeakWorkingSetSize         As Long
   WorkingSetSize             As Long
   QuotaPeakPagedPoolUsage    As Long
   QuotaPagedPoolUsage        As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage     As Long
   PagefileUsage              As Long
   PeakPagefileUsage          As Long
End Type

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const ToMB = 1048576

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" (ByVal handle As Long) As Long

Public Function GetCurrentProcessMemory() As Single
'Get a handle to the Process and Open
On Error GoTo ErrorHappened
    Dim lngCBSize2           As Long
    Dim lngModules(1 To 200) As Long
    Dim lngReturn            As Long
    Dim lngHwndProcess       As Long
    Dim pmc                  As PROCESS_MEMORY_COUNTERS
    
    lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, GetCurrentProcessId)
    If lngHwndProcess <> 0 Then
        'Get an array of the module handles for the specified process
        lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
        'If the Module Array is retrieved, Get the ModuleFileName
        If lngReturn <> 0 Then
            'Get the Site of the Memory Structure
            pmc.cb = LenB(pmc)
            lngReturn = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
        End If
    End If
    
    GetCurrentProcessMemory = pmc.PagefileUsage / (2 ^ 20) 'Convert to MB
    
ExitNow:
On Error Resume Next
    lngReturn = CloseHandle(lngHwndProcess)
Exit Function
ErrorHappened:
    GetCurrentProcessMemory = -1
    Resume ExitNow
    Resume
End Function

Private Function GetCurrentProcessCounter() As PROCESS_MEMORY_COUNTERS
'Get a handle to the Process and Open
On Error GoTo ErrorHappened
    Dim lngCBSize2           As Long
    Dim lngModules(1 To 200) As Long
    Dim lngReturn            As Long
    Dim lngHwndProcess       As Long
    Dim pmc                  As PROCESS_MEMORY_COUNTERS
    
    lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, GetCurrentProcessId)
    If lngHwndProcess <> 0 Then
        'Get an array of the module handles for the specified process
        lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
        'If the Module Array is retrieved, Get the ModuleFileName
        If lngReturn <> 0 Then
            'Get the Site of the Memory Structure
            pmc.cb = LenB(pmc)
            lngReturn = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
        End If
    End If
    
    GetCurrentProcessCounter = pmc
    
ExitNow:
On Error Resume Next
    lngReturn = CloseHandle(lngHwndProcess)
Exit Function
ErrorHappened:
    Resume ExitNow
    Resume
End Function

Public Function GetPagefileUsage() As Single
    GetPagefileUsage = GetCurrentProcessCounter.PagefileUsage / ToMB 'Convert to MB
End Function

Public Function GetPagefileUsagePeak() As Single
    GetPagefileUsagePeak = GetCurrentProcessCounter.PeakPagefileUsage / ToMB 'Convert to MB
End Function

Public Function GetWorkingSetSize() As Single
    GetWorkingSetSize = GetCurrentProcessCounter.WorkingSetSize / ToMB 'Convert to MB
End Function

Public Function GetWorkingSetSizePeak() As Single
    GetWorkingSetSizePeak = GetCurrentProcessCounter.PeakWorkingSetSize / ToMB 'Convert to MB
End Function