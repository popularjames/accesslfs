Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const sendStatus As String = "Sent"
Const timeout As Long = 600000
Const interval As Integer = 30000
Dim timePassed As Long
Dim strDocID As String
Dim strCnlyClaimNum As String
Dim strLockUser As String
Dim strICN As String
Private myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private closedOnce As Integer

Property Let DocID(doc_id As String)
     strDocID = doc_id
End Property

Property Let ClaimNum(claim_num As String)
     strCnlyClaimNum = claim_num
End Property

Property Let Icn(ICN_num As String)
     strICN = ICN_num
End Property

Sub Form_Load()
  
    Me.TimerInterval = interval
    timePassed = 0
    closedOnce = False
End Sub

    Sub Form_Timer()

    Dim sqlStatus As String
    timePassed = timePassed + interval
    sqlStatus = "select Status, DocID, CnlyClaimNum from FAX_WORK_Queue where Client_ext_Ref_ID = '" & gbl_INC_Client_Id & _
        "' and DocID = '" & strDocID & "' and CnlyClaimNum = '" & strCnlyClaimNum & "'"
    
    If timePassed <= timeout Then
        Me.RecordSource = sqlStatus
        Me.Refresh

        If Me.txtStatus = sendStatus Then
           Me.TimerInterval = 0
           Call ChangeClaimStatus
           
                If gbl_TriggerFormTotal = gbl_TriggerFormCurrent Then
                    If closedOnce = False Then
                    DoCmd.Close acForm, Me.Name, acSaveNo
                    closedOnce = True
                    End If
                Else
                    If closedOnce = False Then
                    gbl_TriggerFormCurrent = gbl_TriggerFormCurrent + 1
                    End If
                End If
        End If
    Else
        
        If gbl_TriggerFormTotal = gbl_TriggerFormCurrent Then
                If closedOnce = False Then
                MsgBox ("Timed out while waiting for Fax Status to be updated to " & Chr(34) & "Send" & Chr(34) & " for claim " & strICN & ".")
                DoCmd.Close acForm, Me.Name, acSaveNo
                closedOnce = True
                End If
        Else
            If closedOnce = False Then
            gbl_TriggerFormCurrent = gbl_TriggerFormCurrent + 1
            End If
        End If
           
    End If

    End Sub

Private Function ChangeClaimStatus() As Boolean
    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As String
    Dim ErrMsg As String
    Dim Msg As String
    Msg = "Changed claim status to " & Chr(34) & "Provider Contacted for Incomplete Medical Records" & Chr(34) & " for claim " & strICN & "."

    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")

                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                    cmd.commandType = adCmdStoredProc

                    If ExecuteChangeStatusProc(cmd) <> 0 Then
                        
                            'Maybe this record is locked, unlock it (but only if it's locked by the same user)
                            cmd.CommandText = "usp_AUDITCLM_Hdr_UnLock"
                            cmd.Parameters.Refresh
                            cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
                            cmd.Parameters("@pLockUserID") = GetUserName()
                            cmd.Execute
                            spReturnVal = Nz(cmd.Parameters("@pErrMsg").Value, "")
                            
                            If spReturnVal = "" Then
                                'Change status if claim was unlocked succesfully
                                
                                If ExecuteChangeStatusProc(cmd) <> 0 Then
                                     ChangeClaimStatus = False
                                Else
                                     MsgBox (Msg)
                                     ChangeClaimStatus = True
                                End If
                                                        
                                'Lock record again
                                cmd.CommandText = "usp_AUDITCLM_Hdr_Lock"
                                cmd.Parameters.Refresh
                                cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
                                cmd.Parameters("@pLockUser") = GetUserName()
                                cmd.Parameters("@pLockTime") = Now()
                                cmd.Execute
                                spReturnVal = Nz(cmd.Parameters("@pErrMsg").Value, "")
                                
                            Else
                            MsgBox ("Failed to change claim status after fax had been sent: " & spReturnVal)
                            End If
                    
                    Else
                            ChangeClaimStatus = True
                            MsgBox (Msg)
                    End If

    Set MyCodeAdo = Nothing
    Set cmd = Nothing


End Function

Private Function ExecuteChangeStatusProc(cmd As ADODB.Command) As Integer
                    cmd.CommandText = "usp_AUDITCLM_Incomplete_MR_Change_Status"
                    cmd.Parameters.Refresh
                    cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
                    cmd.Parameters("@DocID") = strDocID
                    cmd.Parameters("@Client_ext_Ref_ID") = gbl_INC_Client_Id
                    cmd.Execute
                    ExecuteChangeStatusProc = cmd.Parameters("@Return_Value")

End Function
