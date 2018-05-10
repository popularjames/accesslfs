Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "UserException"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub ActionID_Enter()
    ActionID.RowSource = "SELECT ActionID, ActionDesc FROM ADMIN_Action WHERE ActionID not in (select ActionID from ADMIN_User_Exception where UserProfID = " & UserProfID & " and AppID = '" & AppID & "')"
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
    
    Me.Caption = "Setup User Exception"
    Call Account_Check(Me)
    
    If IsSubForm(Me) Then
        Me.UserProfID.ColumnHidden = True
    Else
        Me.UserProfID.ColumnHidden = False
        iAppPermission = UserAccess_Check(Me)
    End If
    Me.RecordSource = "select ue.* " & _
                        " from (ADMIN_User_Profile AS up INNER JOIN ADMIN_User_Exception AS ue ON up.UserProfID = ue.UserProfID) " & _
                        " INNER JOIN ADMIN_Profile AS ap ON up.ProfileID = ap.ProfileID" & _
                        " where ap.AccountID = " & gintAccountID
End Sub
