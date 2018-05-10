Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "Profile"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = CheckForNull(Me.Controls, True)
End Sub


Private Sub Form_Current()
    If IsSubForm(Me) Then
        Me.Parent.ProfileID = Me.ProfileID
        
        If Me.NewRecord Then
            Me.Parent.ProfileDetails.Enabled = False
            Me.Parent.lblProfileDetail.Caption = "Profile Details"
        Else
            Me.Parent.lblProfileDetail.Caption = Chr(34) & Me.ProfileID & Chr(34) & " Profile Details"
            Me.Parent.ProfileDetails.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Profile Maintenance"
    
    Call Account_Check(Me, "ADMIN_Profile")
    
    If IsSubForm(Me) = False Then
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
    
    ' set default account ID
    Me.AccountID.RowSource = "select * from ADMIN_Client_Account where AccountID = " & gintAccountID
    Me.AccountID.DefaultValue = gintAccountID
    Me.AccountID.ColumnHidden = True

End Sub
