Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private frm As Access.Form
Private intCancel As Integer
Public Event MouseWheel(Cancel As Integer)

Public Property Set Form(frmIn As Access.Form)
    'Define Property procedure for the class which
    'allows us to set the Form object we are
    'using with it. This property is set from the
    'form class module.
    Set frm = frmIn
End Property

Public Property Get MouseWheelCancel() As Integer
    'Define Property procedure for the class which
    'allows us to retrieve whether or not the Form
    'event procedure canceled the MouseWheel event.
    'This property is retrieved by the WindowProc
    'function in the standard basSubClassWindow
    'module.

    MouseWheelCancel = intCancel
End Property

Public Sub HookForm()
    'Called from the form's OnOpen or OnLoad
    'event. This procedure is what "hooks" or
    'subclasses the form window. If you hook the
    'the form window, you must unhook it when completed
    'or Access will crash.

    lpPrevWndProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, _
                                    AddressOf WindowProc)
      Set CMouse = Me
   End Sub

Public Sub UnHookForm()
    'Called from the form's OnClose event.
    'This procedure must be called to unhook the
    'form window if the SubClassHookForm procedure
    'has previously been called. Otherwise, Access will
    'crash.

    Call SetWindowLong(frm.hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Sub FireMouseWheel()

    'Called from the WindowProc function in the
    'basSubClassWindow module. Used to raise the
    'MouseWheel event when the WindowProc function
    'intercepts a mouse wheel message.
    RaiseEvent MouseWheel(intCancel)
End Sub