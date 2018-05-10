Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
        Me.RecordSource = ""
        Call Get_Data
End Sub


Public Sub Get_Data()
      Dim SQL As String
       
       SQL = "select * from v_Fax_Status_History where CnlyClaimNum = '" & [Forms]![frm_QUEUE_Incomplete_MR]![frm_QUEUE_MR_Request_Sub].Form![txtCnlyClaimNum] & "'" & _
       " and Client_ext_Ref_ID = '" & gbl_INC_Client_Id & "'"
       Me.RecordSource = SQL
       Me.Refresh
         
End Sub
