Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const Msg As String = "The following medical records are necessary in order to process claim "
Const cust_service As String = "Connolly Customer Service"
Public id_set As Boolean


Private Sub Form_Load()
'Need this line otherwise RecordSource gets messed up
Me.RecordSource = ""
Call SetRecordSource

End Sub

Private Sub Form_Current()
 
 If id_set = True Then
    
    If Me.RecordSet.recordCount = 0 Then
       
       Call Form_frm_QUEUE_Incomplete_MR_Review_Claim_Detail.Clean_Fields
       Call Form_frm_Fax_Status_History_MR.Get_Data
       Me.txtMRequested = ""
       Me.txtNotes = ""
       Me.txtFrom = ""
     
    Else
     
       If (Me.txtCnlyClaimNum <> "") Then
       Call SetMRRequestedField(Me.txtCnlyClaimNum, Me.txtICN)
       Else
         Me.txtNotes = ""
         Me.txtMRequested = ""
         Call SetMRRequestedField("", "")
       End If
       Call Form_frm_QUEUE_Incomplete_MR_Review_Claim_Detail.Get_Data
       Call Form_frm_Fax_Status_History_MR.Get_Data
    End If
    
  End If

End Sub


 Public Sub SetRecordSource(Optional strSQL As String = "select * FROM QUEUE_MR_Request_Fax order by ICN")
        
         id_set = True
         Me.RecordSource = strSQL
         Me.Refresh
        
End Sub


Public Function SetMRRequestedField(CnlyClaimNum As String, Icn As String)
    
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim requestText As String
    Dim index As Integer
    index = 1
    requestText = Msg & Icn & ":"
    Dim Notes As Variant
        
        Notes = GetRequestInfoAndNotes(CnlyClaimNum)
        Me.txtMRequested = requestText & Nz(Notes(0), "")
        Me.txtNotes = Notes(1)
        Me.txtFrom = gbl_FromFieldForMR

        Exit Function
    
SetMRRequestedField = True
End Function


Private Sub tglICN_Click()
Dim strSQL As String

If Me.tglICN.Value = 0 Then
    'Default ordering will do
    Call SetRecordSource

Else
    'Want descending
    strSQL = "select * FROM QUEUE_MR_Request_Fax F order by ICN desc"
    Call SetRecordSource(strSQL)
End If
End Sub

Public Sub MoveToNext()
 DoCmd.GoToRecord acForm, Me.Name, acNext
 End Sub
 
