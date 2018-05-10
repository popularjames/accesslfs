Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'HC  5/2010 left in the ClsReg class
Private Declare PtrSafe Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long)

Private Declare PtrSafe Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long)

Private Declare PtrSafe Function RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbvaluename As Long, ByVal lpReserved As Long, lpType As Long, _
        ByVal lpData&, lpcbData As Long)

Private Declare PtrSafe Function RegEnumValueStr& Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbvaluename As Long, ByVal lpReserved As Long, lpType As Long, _
    ByVal lpData$, lpcbData As Long)

Private Declare PtrSafe Function RegQueryValue& Lib "advapi32.dll" Alias "RegQueryValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long)

Private Declare PtrSafe Function RegQueryValueExStr& Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long)

Private Declare PtrSafe Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long)
Private Declare PtrSafe Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long)         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare PtrSafe Function RegSetValueExStr& Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long)         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare PtrSafe Function RegQueryInfoKey& Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
    (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcbValues&, lpcbMaxValueNameLen&, _
        lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As FILETIME)

Private Declare PtrSafe Function RegEnumKeyEx& Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, lpcbName&, ByVal lpReserved&, ByVal lpClass$, lpcbClass&, lpftLastWriteTime As FILETIME)
Private Declare PtrSafe Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey As Long)

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const SYNCHRONIZE = &H100000

Public Enum RegistryHives
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
Public Enum RegAccessEnum
    READ_CONTROL = &H20000
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    STANDARD_RIGHTS_REQUIRED = &HF0000
    STANDARD_RIGHTS_ALL = &H1F0000
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREATE_LINK = &H20
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
End Enum
Public Enum RegDataTypeEnum
    REG_NONE = 0                       ' No value type
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_BINARY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
    REG_LINK = 6                       ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7                   ' Multiple Unicode strings
    REG_RESOURCE_LIST = 8              ' Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

Public Function GetRegValueStr$(ByVal ParHive As RegistryHives, ByVal ParKey$, Optional ByVal ParValue$ = "{Default}")
    Dim LocResult%, LochKey&, LocDataLen&
    Dim LocData$
    Dim LocEndString&
    Dim LocErrState As Boolean
    
    LocData = String(255, Chr(0))
    LocDataLen = Len(LocData)
    LocResult = RegOpenKeyEx(ParHive, ParKey, 0&, KEY_READ, LochKey)
    If ParValue = "{Default}" Then
        LocResult = RegQueryValue(LochKey, vbNullString, LocData, LocDataLen)
    Else
        LocResult = RegQueryValueExStr(LochKey, ParValue, 0&, REG_SZ, LocData, LocDataLen)
    End If
    If LocResult <> 0 Then
        Err.Raise vbObjectError + 512 + LocResult, "GetRegValueStr", "Unable to open key"
        LocErrState = True
        GetRegValueStr = vbNullString
    End If
    LocResult = RegCloseKey(LochKey)
    If LocErrState Then Exit Function
    LocEndString = InStr(1, LocData, Chr(0), vbBinaryCompare)
    If LocEndString > 0 Then
        LocData = left(LocData, LocEndString - 1)
    End If
    GetRegValueStr = LocData
End Function

Public Function WriteRegValueStr(ByVal ParHive As RegistryHives, ByVal ParKey$, ByVal ParData$, Optional ByVal ParValue$ = "{Default}", Optional ByVal ParCreate As Boolean = False) As Boolean
    Dim LocResult&, LochKey&, LocDataLen&
    Dim LocDisposition&
    Dim LocErrState As Boolean
    LocDataLen = Len(ParData)
    If ParCreate Then
        LocResult = RegCreateKeyEx(ParHive, ParKey, 0&, vbNullString, 0&, KEY_ALL_ACCESS, 0&, LochKey, LocDisposition)
    Else
        LocResult = RegOpenKeyEx(ParHive, ParKey, 0&, KEY_ALL_ACCESS, LochKey)
    End If
    If LocResult <> 0 Then
        Err.Raise vbObjectError + 512 + 1, "WriteRegValueStr", "Unable to open key"
        LocErrState = True
        WriteRegValueStr = False
        LocResult = RegCloseKey(LochKey)
        Exit Function
    End If
    If ParValue = "{Default}" Then
        LocResult = RegSetValue(LochKey, vbNullString, REG_SZ, ParData, LocDataLen)
    Else
        LocResult = RegSetValueExStr(LochKey, ParValue, 0&, REG_SZ, ParData, LocDataLen)
    End If
    If LocResult <> 0 Then
        Err.Raise vbObjectError + 512 + 1, "WriteRegValueStr", "Unable to set value"
        LocErrState = True
        WriteRegValueStr = False
        LocResult = RegCloseKey(LochKey)
        Exit Function
    End If
    LocResult = RegCloseKey(LochKey)
    WriteRegValueStr = True
End Function

'
'Public Function ListRegKeys(ByVal ParRoot As RegistryHives, ByVal ParKey$) As Variant
'    Dim LocResult&
'    Dim LocHandle&
'    Dim LocClass As String * 255
'    Dim LocClassLen&, LocSubKeyCount&, LocMaxSKLength&, LocMaxClassLen&, LocValueCount&, LocMaxValNameLen&
'    Dim LocMaxValLen&, LocSD&, LocRWTime As FILETIME
'    Dim LocName$
'    Dim z&
'    Dim LocNameLen&
'    Dim LocRegKeyArray() As String
'    Dim LocStrings As New ClsStrings
'
'    LocResult = RegOpenKeyEx(ParRoot, ParKey, 0&, KEY_READ, LocHandle)
'    LocClassLen = 256
'    LocMaxSKLength = 1024
'    LocMaxClassLen = 1024
'    LocMaxValNameLen = 1024
'    LocMaxValLen = 1024
'    LocResult = RegQueryInfoKey(LocHandle, LocClass, LocClassLen, 0&, LocSubKeyCount, LocMaxSKLength, LocMaxClassLen, LocValueCount, LocMaxValNameLen, LocMaxValLen, LocSD, LocRWTime)
'    ReDim LocRegKeyArray(LocSubKeyCount - 1)
'    For z = 0 To LocSubKeyCount - 1
'        LocName = String(LocMaxSKLength, Chr(0))
'        LocNameLen = Len(LocName) + 1
'        LocResult = RegEnumKeyEx(LocHandle, z, LocName, LocNameLen, 0&, LocClass, LocClassLen, LocRWTime)
'        LocRegKeyArray(z) = LocStrings.TrimNullStr(LocName)
'    Next z
'    ListRegKeys = LocRegKeyArray
'    Set LocStrings = Nothing
'    LocResult = RegCloseKey(LocHandle)
'End Function

'Public Function LisRegValuesLng(ByVal ParRoot As RegistryHives, ByVal ParKey$) As Variant
'    Dim LocResult&
'    Dim LocHandle&
'    Dim LocClass As String * 255
'    Dim LocClassLen&, LocSubKeyCount&, LocMaxSKLength&, LocMaxClassLen&, LocValueCount&, LocMaxValNameLen&
'    Dim LocMaxValLen&, LocSD&, LocRWTime As FILETIME
'    Dim LocName$
'    Dim z&
'    Dim LocNameLen&
'    Dim LocType&
'    Dim LocRegValArray() As String
'    Dim LocStrings As New ClsStrings
'    Dim LocData&
'    Dim LocDataLen&
'
'    LocClassLen = 256
'    LocMaxSKLength = 1024
'    LocMaxClassLen = 1024
'    LocMaxValNameLen = 1024
'    LocMaxValLen = 1024
'    LocDataLen = Len(LocData)
'    LocResult = RegOpenKeyEx(ParRoot, ParKey, 0&, KEY_READ, LocHandle)
'    LocResult = RegQueryInfoKey(LocHandle, LocClass, LocClassLen, 0&, LocSubKeyCount, LocMaxSKLength, LocMaxClassLen, LocValueCount, LocMaxValNameLen, LocMaxValLen, LocSD, LocRWTime)
'    ReDim LocRegValArray(LocValueCount - 1)
'    For z = 0 To LocValueCount - 1
'        LocName = String(LocMaxValLen, Chr(0))
'        LocNameLen = Len(LocName) + 1
'        LocResult = RegEnumValue(LocHandle, z, LocName, LocNameLen, 0&, LocType, LocData, LocDataLen)
'        LocRegValArray(z) = LocStrings.TrimNullStr(LocName)
'    Next z
'
'    LocResult = RegCloseKey(LocHandle)
'    LisRegValuesLng = LocRegValArray
'End Function