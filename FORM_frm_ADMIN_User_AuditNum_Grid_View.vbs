Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AuditNum_AfterUpdate()
    Dim rs As DAO.RecordSet
    Set rs = CurrentDb.OpenRecordSet("select * from ADMIN_Audit_Number where AuditNum = " & CStr(Me.AuditNum))
    Me.EffDt = rs("EffDt")
    Me.TermDt = rs("TermDt")
End Sub

Private Sub AuditNum_Enter()
    Me.AuditNum.RowSource = "select * from ADMIN_Audit_Number aan where not exists " & _
                      "  ( select 1 from ADMIN_User_AuditNum aau " & _
                      "    where aau.AuditNum = aan.AuditNum " & _
                      "    and aau.UserID = '" & Me.UserID.Value & "')"
End Sub
