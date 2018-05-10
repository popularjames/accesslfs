Option Compare Database
Option Explicit

'HC since this is a CnlyFolderBrowse class left the declares here
Public Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal Folder As Long, ByRef idl As Long) As Long
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal idl As Long, ByVal Path As String) As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (ByRef bi As BROWSEINFO) As Long

'Get Directory Windows API
Private selection As String

Private Const conNoError = 0&
Private Const conMaxPath = 260
Private Const conErrorExtendedError = 1208&

' Messages sent to callback function.
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2

' Messages to browser from callback function.
Private Const WM_USER = &H400
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Type BROWSEINFO
    hWndOwner As Long           ' Owner
    pidlRoot As Long            ' Can be null
    strDisplayName As String    ' Recieves display name of folder
    strTitle As String          ' title/instructions for user
    ulFlags As Long             ' 0 or BIF constants
    lpfn As Long                ' Address for callback
    lParam As Long              ' Passes to callback
    iImage As Long              ' index to the system image list
End Type

'HC 5/2010 changed to use the centralized winapi functions
Private Function SetSelection(hwnd As Long, varSel As Variant) As Long
    ' Set the selection in the Shell browser dialog. If
    ' varSel is numeric, the code assumes that the value is a
    ' PDIL, and calls SendMessage with that information. If
    ' not, the code converts the value to a string, and
    ' tells the browser that it's sending a string value instead.
    If IsNumeric(varSel) Then
        Call SendWindowMessage(hwnd, BFFM_SETSELECTION, 0, CLng(varSel))
    Else
        Call SendWindowMessage(hwnd, BFFM_SETSELECTION, 1, CStr(varSel & ""))
    End If
End Function
'HC 5/2010 changed to use the centralized winapi functions
Private Sub SetStatus(hwnd As Long, strText As String)
    Call SendWindowMessage(hwnd, BFFM_SETSTATUSTEXT, 0, CStr(strText & ""))
End Sub

Public Function BrowseForFolder(InitDir As String, szTitle As String, hwnd As Long) As String
'Opens a Treeview control that displays the directories in a computer
On Error GoTo GetDirectoryFromTreeError

Dim lpIDList As Long
Dim sBuffer As String
Dim tBrowseInfo As BROWSEINFO

With tBrowseInfo

    .hWndOwner = hwnd
    .pidlRoot = &H0
    .strDisplayName = Space(conMaxPath)
    .strTitle = szTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
    If InitDir <> "" Then
        .lpfn = FnPtrToLong(AddressOf CallBackFunction)
        selection = InitDir
    End If

    
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(conMaxPath)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = IIf(sBuffer = "", InitDir, sBuffer)
Else
    BrowseForFolder = ""
End If

GetDirectoryFromTreeExit:
    On Error Resume Next
    Exit Function
GetDirectoryFromTreeError:
    MsgBox Err.Description
    BrowseForFolder = InitDir
    Resume GetDirectoryFromTreeExit
End Function


Private Function CallBackFunction(ByVal hwnd As Long, ByVal lngMsg As Long, ByVal lngParam As Long, ByVal lngData As Long) As Long
On Error Resume Next
    ' This is an example callback function, simply demonstrating
    ' the kinds of things this function can do.
    
    ' The formal declaration of this function is specified by the
    ' SHBrowseForFolder API call, and cannot be changed.
    
    Dim strFolder As String
        
    Select Case lngMsg
        Case BFFM_INITIALIZED
            Call EnableOK(hwnd, False)
            Call SetSelection(hwnd, selection)
            
        Case BFFM_SELCHANGED
            strFolder = Space(conMaxPath)
            
            ' Try to resolve the lngIDL into a real path
            If CBool(SHGetPathFromIDList(lngParam, strFolder)) Then
                ' Trim the null characters
                strFolder = TrimNull(strFolder)
                Call SetStatus(hwnd, strFolder)
                Call EnableOK(hwnd, True)
            Else
                ' Not a file system path, so
                ' clear that status, and disable
                ' the OK button.
                Call SetStatus(hwnd, vbNullString)
                Call EnableOK(hwnd, False)
            End If
    End Select
    CallBackFunction = 0
End Function
Private Sub EnableOK(hwnd As Long, fEnable As Boolean)
    ' Enable the OK button on the Shell browser dialog.
    Dim lngEnable As Long
    
    If fEnable Then
        lngEnable = 1
    Else
        lngEnable = 0
    End If
    Call SendWindowMessage(hwnd, BFFM_ENABLEOK, 0, lngEnable)
End Sub
Private Function TrimNull(ByVal strValue As String) As String
    ' Find the first vbNullChar in a string, and return
    ' everything prior to that character. Extremely
    ' useful when combined with the Windows API function calls.
    
    ' In:
    '   strValue:
    '       Input text, possibly containing a null character
    '       (chr$(0), or vbNullChar)
    ' Out:
    '   Return Value:
    '       strValue trimmed on the right, at the location
    '       of the null character, if there was one.
    
    Dim intPos As Integer
    
    intPos = InStr(strValue, vbNullChar)
    Select Case intPos
        ' It's best to put the most likely case first.
        Case Is > 1
            ' Found in the string, so return the portion
            ' up to the null character.
            TrimNull = left$(strValue, intPos - 1)
        Case 0
            ' Not found at all, so just
            ' return the original value.
            TrimNull = strValue
        Case 1
            ' Found at the first position, so return
            ' an empty string.
            TrimNull = ""
    End Select
End Function
Private Function FnPtrToLong(ByVal lngFnPtr As Long) As Long
    ' Given a function pointer as a Long, return a Long.
    ' Sure looks like this function isn't doing anything,
    ' and in reality, it's not.
    FnPtrToLong = lngFnPtr
End Function