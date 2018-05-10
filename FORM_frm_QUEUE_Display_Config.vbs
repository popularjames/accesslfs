Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "QueueDispCnfg"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    Me.RecordSource = "SELECT * FROM QUEUE_Display_Config WHERE AccountID = " & gintAccountID & " ORDER BY groupname, comboboxorder"
    Me.Requery
End Sub
