Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "UserAcct"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Account User Setup"
    
    Call Account_Check(Me, "ADMIN_User_Account")
    iAppPermission = UserAccess_Check(Me)
End Sub

Private Sub UserID_AfterUpdate()
    Me.AccountID = gintAccountID
End Sub

Private Sub UserID_Enter()
    Me.UserID.RowSource = "select * from ADMIN_User u where not exists (select 1 from ADMIN_User_Account ua where ua.UserID = u.UserID)"
    Me.UserID.Requery
End Sub
