Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property





Public Property Get UserMessage() As String
    UserMessage = Me.lblMessage.Caption
End Property
Public Property Let UserMessage(sMessage As String)
    Me.lblMessage.Caption = sMessage
End Property





Public Property Get Title() As String
    Title = Me.Caption
End Property
Public Property Let Title(sMessage As String)
    Me.Caption = sMessage
End Property
