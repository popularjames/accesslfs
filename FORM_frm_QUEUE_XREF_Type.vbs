Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "QueueType"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Queue Type Setup"
    
    iAppPermission = UserAccess_Check(Me)
End Sub
