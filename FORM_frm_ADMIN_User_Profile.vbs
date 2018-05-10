Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "UserProfile"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Current()
    If IsSubForm(Me) Then
        Me.Parent.UserProfID = Me.UserProfID
        
        If Me.NewRecord Then
            Me.Parent.Exceptions.Enabled = False
        Else
            Me.Parent.Exceptions.Enabled = True
        End If
    End If

End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "User Profile Maintenance"
    
    Call Account_Check(Me)
    
    If IsSubForm(Me) = False Then
        iAppPermission = UserAccess_Check(Me)
    End If
    
    Me.RecordSource = "select * from ADMIN_User_Profile where AccountID = " & gintAccountID
    Me.UserProfID.ColumnHidden = True
    
    Me.UserID.RowSource = "select * from ADMIN_User_Account where AccountID = " & gintAccountID
    Me.UserID.Requery
    
    Me.ProfileID.RowSource = "select * from ADMIN_Profile where AccountID = " & gintAccountID
    Me.ProfileID.Requery
    

End Sub


Private Sub ProfileID_AfterUpdate()
    Me.AccountID = gintAccountID
End Sub

Private Sub UserID_Enter()
    Me.UserID.RowSource = "select * from ADMIN_User_Account ua where not exists (select 1 from ADMIN_User_Profile up where ua.UserID = up.UserID)"
    Me.UserID.Requery
End Sub
