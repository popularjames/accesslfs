Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub CnlyClaimNum_DblClick(Cancel As Integer)
'Opens the Audit Claim main screen when the user double clicks within the frm_QA_Claims_List results
'Unit test 9/10/2012 by KCF
    If Me.CnlyClaimNum & "" <> "" Then
        DisplayAuditClmMainScreen Me.CnlyClaimNum
    End If
End Sub


Private Sub Form_Click()
'Unit test 9/10/2012 by KCF
    If IsSubForm(Me) Then
        If Me.SelHeight > 1 Then 'If more than 1 row selected in the List of Claims
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = Me.SelHeight
            Me.Parent.Form.cmdSubmitQA.Enabled = False 'KCF only submit a single record
            Me.Parent.Form.cmdRemoveReviews.Enabled = True 'kcf
        Else 'If 0 or 1 rows selected in the list of claims
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = 1
        End If
    End If
End Sub

Private Sub Form_Current()
'For current record selected in the Claim List form
'Unit test 9/10/2012
    If IsSubForm(Me) Then 'Only if the Claim List is a subform
        If Not (Me.RecordSet Is Nothing) Then
            If Me.RecordSet.recordCount > 0 Then
                If Me.Parent.DetailFormLoaded Then
                    'Populate the Results form with the currently selected Claim in the List
                     If Me.Parent.Name = "frm_QA_Review_Main" Then
                        Me.Parent.Form.frm_QA_Claim_Review_Result.Form.RecordSource = "select * from QA_Review_Worktable where SeqNo = " & Me.SeqNo & " and CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
                        Me.Parent.Form.frm_QA_Claim_Review_Result.Form.txtAHClmStatus = Me.txtAHClmStatus 'KCF 2/15/2013
                    ElseIf Me.Parent.Name = "frm_qa_review_Main_submitted" Then
                        Me.Parent.Form.frm_QA_Claim_Review_Result.Form.RecordSource = "select * from QA_Review_Worktable where SeqNo = " & Me.SeqNo & " and CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
                        Me.Parent.Form.frm_QA_Claim_Review_Result.Form.txtAHClmStatus = Me.txtAHClmStatus 'KCF 2/15/2013
                   'MsgBox (Me.txtAHClmStatus)
                    End If
                End If
            End If
        End If

        If Me.SelHeight > 1 Then
        'If more than 1 row selected in the datasheet
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = Me.SelHeight
            Me.Parent.Form.cmdSubmitQA.Enabled = False 'KCF only submit a single record
            Me.Parent.Form.cmdRemoveReviews.Enabled = True 'kcf
        Else
            Me.Parent.StartRow = Me.SelTop
            Me.Parent.RowsSelected = 1
        End If
        
    
    End If
    


End Sub

Private Sub Form_Load()
'When the form is loaded will also call stored procedure to populate the QA Worktable with appropriate Claims
'When the form is loaded, will populate the Claim List with records that are qualified for Review (SubmitDate is null)
'Unit test 9/10/2012 by kcf

'CHANGE HISTORY: Updated Thursday 8/29/2013 to use TRY \ CATCH in the usp_QA_Review_Worktable)_Insert; Changed the ADO connection & added the Error Handling

'BEGIN 8/29/2013: KCF
    
'    Dim myCode_ADO As clsADO
'    Dim cmd As ADODB.Command
     Dim strErrMsg As String
    
     On Error GoTo Err_handler
    
'    Set myCode_ADO = New clsADO
'    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")

'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = myCode_ADO.CurrentConnection
'    cmd.commandType = adCmdStoredProc
'    cmd.CommandText = "dbo.usp_QA_Review_Worktable_Insert"
'    cmd.Parameters.Refresh
'    cmd.Execute

'    If cmd.Parameters("@RETURN_VALUE") <> 0 Then
'        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_Worktable_Insert"
'        Err.Raise 50001, "usp_QA_Review_Worktable_Update", strErrMsg
'    End If
'END 8/29/2013: KCF
'VS 10/29/2014: usp_QA_Review_Worktable_Insert is scheduled to run every hour now, so it should not be executed from VBA code anymore.

    
    '10/11/2012 use view instead of table
    If Me.Parent.Name = "frm_QA_Review_Main" Then
        Me.RecordSource = "select Top 1000 * from v_QA_Review_WorkTable_Unsubmitted order by Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN"
    ElseIf Me.Parent.Name = "frm_QA_Review_Main_submitted" Then
        Me.RecordSource = "Select Top 1000 * from v_QA_Review_WorkTable_Submitted where QAStatus in ('A', 'C', 'R') and submitdate > Now() - 5 order by SubmitDate desc, Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN"
    End If
    
    
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    Me.AllowEdits = False
    
    If IsSubForm(Me) Then
        Me.Parent.GridFormLoaded = True
    End If
    
    'Set myCode_ADO = Nothing
 
'BEGIN 8/29/2013: KCF
Exit_Sub:
'    Set myCode_ADO = Nothing
'    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
'END 8/29/2013: KCF
    
End Sub

Private Sub ICN_DblClick(Cancel As Integer)
'User can double click the Claim Num in the List to bring up the Main Claim form
'Unit test 9/10/2012 by kcf
    If Me.CnlyClaimNum & "" <> "" Then
        DisplayAuditClmMainScreen Me.CnlyClaimNum
    End If
End Sub
