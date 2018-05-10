Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvText As String
Private mvResults As Integer
Private mvTitle As String


Public Property Let Text(data As String)
    MvText = data
    Me.Txt = "" & MvText
End Property


Public Property Get Text() As String
    Text = MvText
End Property

Public Property Get Results() As Integer
    Results = mvResults
End Property

Public Property Get Title() As String
    Title = mvTitle
End Property
Public Property Let Title(val As String)
    mvTitle = val
    Me.lblTitle.Caption = mvTitle
    Me.Caption = mvTitle
End Property



Private Sub CmdCancel_Click()
    mvResults = False
    Me.visible = False
End Sub

Private Sub cmdOk_Click()
    MvText = "" & Txt
    mvResults = True
    Me.visible = False
End Sub

Private Sub Form_Open(Cancel As Integer)
mvResults = 1
End Sub
