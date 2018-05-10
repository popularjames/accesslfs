Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AuditNum"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_BeforeInsert(Cancel As Integer)
    Me.AccountID = gintAccountID
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Audit Number"
    
    Call Account_Check(Me, "ADMIN_Audit_Number")
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    Me.AccountID.RowSource = "SELECT * FROM ADMIN_Client_Account WHERE AccountID=" & gintAccountID
    Me.AccountID.Requery
    
End Sub
