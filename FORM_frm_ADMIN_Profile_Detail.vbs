Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "ProfileDtl"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub ActionID_Enter()
    ActionID.RowSource = "SELECT ActionID, ActionDesc FROM ADMIN_Action WHERE ActionID not in (select ActionID from ADMIN_Profile_Detail where ProfileID = '" & ProfileID & "' and AppID = '" & AppID & "')"
    ActionID.Requery
End Sub

Private Sub AppID_AfterUpdate()
    AppID.DefaultValue = Chr(34) & AppID & Chr(34)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = CheckForNull(Me.Controls, True)
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Profile Details"
    
    Call Account_Check(Me)
    
    If IsSubForm(Me) Then
        Me.ProfileID.ColumnHidden = True
    Else
        Me.ProfileID.ColumnHidden = False
        Me.RecordSource = "select pd.* from ADMIN_Profile_Detail pd inner join ADMIN_Profile p on pd.ProfileID = p.ProfileID where p.AccountID = " & gintAccountID
    
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
End Sub
