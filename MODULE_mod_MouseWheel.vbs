Option Compare Database
Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
     ByVal hwnd As Long, _
     ByVal Msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const WM_MouseWheel = &H20A
Public lpPrevWndProc As Long
Public CMouse As clsMouseWheel

Public Function WindowProc(ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    'Look at the message passed to the window. If it is
    'a mouse wheel message, call the FireMouseWheel procedure
    'in the CMouseWheel class, which in turn raises the MouseWheel
    'event. If the Cancel argument in the form event procedure is
    'set to False, then we process the message normally, otherwise
    'we ignore it.  If the message is something other than the mouse
    'wheel, then process it normally
    Select Case uMsg
        Case WM_MouseWheel
            CMouse.FireMouseWheel
            If CMouse.MouseWheelCancel = False Then
                WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
            End If


        Case Else
           WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function