Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adj_ReasonText_AfterUpdate()
     Dim MyAdo As clsADO
     Dim rs As ADODB.RecordSet
     
     Set MyAdo = New clsADO
     MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Select Case Me.adj_ReasonText
    Case 1
        Me.txtInfo.Value = "This claim will be resent for adjustment with the New ICN. Change Claim status to WAITING TO SEND TO PAYER above." & _
        "The actual savings might be different than the projected savings in the claim."
    Case 2
        Me.txtInfo.Value = "This claim will be sent to the MAC asking for an explaination."
    Case 3, 6
        Me.txtInfo.Value = "Fetching MR Request Date for this claim..."
        screen.MousePointer = 11
        Set rs = MyAdo.OpenRecordSet("select cast(MIN(lh.LetterReqDt) as date) as MRDate from CMS_AUDITORS_Claims.dbo.LETTER_Detail ld " & _
                                        "inner join CMS_AUDITORS_Claims.dbo.LETTER_Header lh " & _
                                        "on lh.InstanceID = ld.InstanceID " & _
                                        "where ld.CnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "' and lh.LetterType LIKE 'VADRA%'")
        screen.MousePointer = 0
        If rs.recordCount > 0 Then
            Me.adj_ConnollyAdjDate = rs.Fields(0).Value
        Else
            MsgBox "No MR Request letter found for this claim"
        End If
        Me.txtInfo.Value = "This claim will be manually invoiced to CMS, if the Provider adjusted the claim after our MR request date shown above. Change claim status to CLAIM READY FOR MANUAL INVOICE above."
    Case Else
        Me.txtInfo.Value = "Choose an appropriate Reason"
    End Select
End Sub

Private Sub GetFields()


End Sub





Private Sub btn_AdjICN_Click()
     Dim MyAdo As clsADO
     Dim rs As ADODB.RecordSet
     
     Set MyAdo = New clsADO
     MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
     Set rs = MyAdo.OpenRecordSet("select ICN, DRG from CMS_Data_NCH.dbo.INT_HDR where CnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "'")
    
    If rs.recordCount > 0 Then
        Me.adj_ICN = rs.Fields(0).Value
        MsgBox "Adjusted DRG on the new claim (" & rs.Fields(0).Value & ") is " & rs.Fields(1).Value
    Else
        MsgBox "No Claim information loaded yet"
     End If
     
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord = True Then
        Me.CnlyClaimNum = Me.Parent.Form.CnlyClaimNum
    End If
End Sub




Private Sub adj_ReasonText_group_Change()
Me.adj_ReasonText = Null
 Me.adj_ReasonText.Requery
 Me.adj_ReasonText = Me.adj_ReasonText.ItemData(0)
End Sub
