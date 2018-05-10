Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub btnDone_Click()
On Error GoTo Err_Command3_Click

    DoCmd.Close

Exit_Command3_Click:
    Exit Sub

Err_Command3_Click:
    MsgBox Err.Description
    Resume Exit_Command3_Click
End Sub

Private Sub Form_Close()
    Forms("frm_APPEAL_main").mbFinished = True
    Forms("frm_APPEAL_main").strMemo = Me.strMemo
    Forms("frm_APPEAL_main").strMailTo = Me.MailTo
    Forms("frm_APPEAL_main").strMailSubject = Me.MailSubject
End Sub

Private Sub Form_Current()
    Me.MailTo = "gautam.malhotra@connolly.com;Jason.Kyle@Connolly.com" 'joseph.casella@connolly.com;jim.spause@connolly.com
    Me.MailSubject = "New Appeal Packages Waiting To Be Sent"
End Sub
