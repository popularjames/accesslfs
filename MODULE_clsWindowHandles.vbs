Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private miWindowHandle As Long
Private mstrWindowName As String

Public Property Let WindowHandle(ByVal vData As Long)
    miWindowHandle = vData
End Property

Public Property Get WindowHandle() As Long
    WindowHandle = miWindowHandle
End Property

Public Property Let WindowName(ByVal vData As String)
    mstrWindowName = vData
End Property

Public Property Get WindowName() As String
    WindowName = mstrWindowName
End Property