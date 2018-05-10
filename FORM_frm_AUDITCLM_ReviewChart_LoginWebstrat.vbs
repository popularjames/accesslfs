Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim frmParent As Form

Private Sub btnCancel_Click()
    'MsgBox self.frmParent.Name
    'Forms("frm_AUDITCLM_ReviewChart").mbFinished = False
End Sub

Private Sub btnDone_Click()
On Error GoTo Err_Command3_Click
    
    'Forms("frm_AUDITCLM_ReviewChart").mbFinished = True
    DoCmd.Close

Exit_Command3_Click:
    Exit Sub

Err_Command3_Click:
    MsgBox Err.Description
    Resume Exit_Command3_Click
End Sub

Private Sub Form_Close()
    'Forms("frm_AUDITCLM_ReviewChart").strUserName = Me.UserName
    'Forms("frm_AUDITCLM_ReviewChart").strPassWord = Me.password
End Sub
