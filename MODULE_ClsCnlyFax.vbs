Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_sender As Object ' Connolly.FaxSender
''''''''''''''''''''''
' Init method
''''''''''''''''''''''
Public Sub Class_Initialize()
    Set m_sender = CreateObject("Connolly.FaxSender")
End Sub
''''''''''''''''''''''
' Terminate method
''''''''''''''''''''''
Private Sub Class_Terminate()
    Set m_sender = Nothing
End Sub
''''''''''''''''''''''
' SendFax method
'
' On success, returns empty string.  On failure, returns reason.
''''''''''''''''''''''
Public Function SendFax(ByVal tiffFile As String, ByVal FaxNumber As String, _
        ByVal senderEmailAddress As String, ByVal uniqueIdentifier As String) As String
    SendFax = m_sender.SendFax(tiffFile, FaxNumber, senderEmailAddress, uniqueIdentifier)
End Function
''''''''''''''''''''''
' WaitForFile method
'
' Waits for up to the specified number of milliseconds or until the file is found.
' Returns whether it was found.
''''''''''''''''''''''
Public Function WaitForFile(ByVal pathMask As String, ByVal milliseconds As Integer, _
        ByVal recurse As Boolean)
    WaitForFile = m_sender.WaitForFile(pathMask, milliseconds, recurse)
End Function
''''''''''''''''''''''
' GenerateId method
'
' Generates a unique ID.
''''''''''''''''''''''
Public Function GenerateId() As String
    GenerateId = m_sender.GenerateId()
End Function
''''''''''''''''''''''
' TempDirectory property
'
' Gets raw temp directory.
''''''''''''''''''''''
Public Property Get TempDirectory() As String
    TempDirectory = m_sender.TempDirectory
End Property

''''''''''''''''''''''''''''''''''
' Suffix Property
' Added 4/23/2012 Curlan Johnson
''''''''''''''''''''''''''''''''''
Public Property Get Suffix() As String
    FaxSuffix = m_sender.FaxSuffix
End Property

''''''''''''''''''''''''''''''''''
' Suffix Property
' Added 4/23/2012 Curlan Johnson
''''''''''''''''''''''''''''''''''
Public Property Let Suffix(ByVal Value As String)
    m_sender.FaxSuffix = Value
End Property

''''''''''''''''''''''
' SMTP Host property
''''''''''''''''''''''
Public Property Get Host() As String
    Host = m_sender.Host
End Property
Public Property Let Host(ByVal Value As String)
    m_sender.Host = Value
End Property
''''''''''''''''''''''
' SMTP Port property
''''''''''''''''''''''
Public Property Get Port() As String
    Port = m_sender.Port
End Property
Public Property Let Port(ByVal Value As String)
    m_sender.Port = Value
End Property
''''''''''''''''''''''
' Timeout property
''''''''''''''''''''''
Public Property Get timeout() As Integer
    timeout = m_sender.timeout
End Property
Public Property Let timeout(ByVal Value As Integer)
    m_sender.timeout = Value
End Property
''''''''''''''''''''''
' GetDestinationTag method
''''''''''''''''''''''
Public Function GetDestinationTag(ByVal destinationPath As String) As String
    GetDestinationTag = m_sender.GetDestinationTag(destinationPath)
End Function
''''''''''''''''''''''
' FaxSuffix property
''''''''''''''''''''''
Public Property Get FaxSuffix() As String
    FaxSuffix = m_sender.FaxSuffix
End Property
Public Property Let FaxSuffix(ByVal Value As String)
    m_sender.FaxSuffix = Value
End Property
''''''''''''''''''''''
' FaxPrefix property
''''''''''''''''''''''
Public Property Get FaxPrefix() As String
    FaxPrefix = m_sender.FaxPrefix
End Property
Public Property Let FaxPrefix(ByVal Value As String)
    m_sender.FaxPrefix = Value
End Property