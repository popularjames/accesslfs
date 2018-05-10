Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_printerName As String
Private m_devMode As Object
Private m_pDevMode As Long
Private m_dirty As Boolean
''''''''''''''''''''''
' Com property
''''''''''''''''''''''
Public Property Get Com() As Object
    Set Com = m_devMode
End Property
''''''''''''''''''''''
' ID property
''''''''''''''''''''''
Public Property Get ID() As Long
    ID = m_pDevMode
End Property
''''''''''''''''''''''
' PrinterName property
''''''''''''''''''''''
Public Property Get PrinterName() As String
  PrinterName = m_printerName
End Property
''''''''''''''''''''''
' Dirty property
''''''''''''''''''''''
Public Property Get Dirty() As Boolean
    Dirty = m_dirty
End Property
Public Property Let Dirty(ByVal Value As Boolean)
    m_dirty = Value
End Property
''''''''''''''''''''''
' RequestSave method
''''''''''''''''''''''
Public Sub RequestSave()
    m_dirty = True
End Sub
''''''''''''''''''''''
' Save method
''''''''''''''''''''''
Public Sub Save()
    If Not m_devMode.SaveBlackIceDEVMODE(m_printerName, m_pDevMode) Then Err.Raise vbObjectError + 1, "Cannot save devmode"
    m_dirty = False
End Sub
Private Sub Class_Initialize()
    m_printerName = "Connolly Fax"

    Set m_devMode = CreateObject("BLACKICEDEVMODE.BlackIceDEVMODECtrl.1")
    If m_devMode Is Nothing Then Err.Raise vbObjectError + 1, "Black Ice DevMode OCX not installed"

    m_pDevMode = m_devMode.LoadBlackIceDEVMODE(m_printerName)
    If m_pDevMode = 0 Then Err.Raise vbObjectError + 1, "Black Ice Printer Driver not working"
End Sub
Private Sub Class_Terminate()
    ' If nothing to do, do nothing.
    If m_devMode Is Nothing Or m_pDevMode = 0 Then Exit Sub
    On Error GoTo Terminate_Fail

    ' Save changes.
    If m_dirty Then Save

Terminate_Fail:
    Dim errNumber As Integer
    Dim ErrSource As String
    Dim errDesc As String
    errNumber = Err.Number
    ErrSource = Err.Source
    If errNumber <> 0 Then errDesc = Err.Description
    On Error GoTo 0

    ' Release dev mode structure, regardless.
    m_devMode.ReleaseBlackIceDEVMODE m_pDevMode
    Set m_devMode = Nothing
    m_pDevMode = 0
    m_dirty = False

    If errNumber > 0 Then Err.Raise errNumber, ErrSource, errDesc
End Sub