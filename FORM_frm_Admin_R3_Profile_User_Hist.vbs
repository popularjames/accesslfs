Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()

    Dim iAppPermission As Integer
    
    Me.Caption = "Setup User Exception"
    Call Account_Check(Me)
    
    If IsSubForm(Me) Then
        Me.UserProfID.ColumnHidden = True
    Else
        Me.UserProfID.ColumnHidden = False
        iAppPermission = UserAccess_Check(Me)
    End If
    
'    Me.RecordSource = "SELECT auph.* " & _
'                        "From cms_auditors_claims.dbo.ADMIN_User_Profile AUP" & _
'                        "inner join cms_auditors_claims.dbo.ADMIN_User_Profile_Audit_Hist AUPH on AUP.UserProfID = auPH.UserProfID" & _
'                        "where aup.userprofid = Form.frm_admin_R3_User_profile.UserprofID"
'
'
    
'    Me.RecordSource = "select ue.* " & _
'                        " from (ADMIN_User_Profile AS up INNER JOIN ADMIN_User_Exception AS ue ON up.UserProfID = ue.UserProfID) " & _
'                        " INNER JOIN ADMIN_Profile AS ap ON up.ProfileID = ap.ProfileID" & _
'                        " where ap.AccountID = " & gintAccountID
End Sub
