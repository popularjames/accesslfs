Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private WithEvents frm_QA_Review_Result_Collection As Form_frm_QA_Review_Result_Collection
'Private WithEvents frmReAssignSelect As Form_frm_QUEUE_ReAssign_Select

Private mstrFilter As String
Private mstrSort As String
Private mbDetailFormLoaded As Boolean
Private mbGridFormLoaded As Boolean
Private miRowsSelected As Long
Private miStartRow As Long
Private miAppPermission As Integer
Private mstrUserProfile As String
Private mstrUserName As String
Public mbRemove As Boolean

Public strListAuditors As String
Public strListDRGs As String  '3/21/2013 KCF - to allow search for multiple DRGs

Const CstrFrmAppID As String = "QAMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get DetailFormLoaded() As Boolean
    DetailFormLoaded = mbDetailFormLoaded
End Property

Public Property Let DetailFormLoaded(ByVal vData As Boolean)
    mbDetailFormLoaded = vData
End Property

Public Property Get GridFormLoaded() As Boolean
    GridFormLoaded = mbGridFormLoaded
End Property

Public Property Let GridFormLoaded(ByVal vData As Boolean)
    mbGridFormLoaded = vData
End Property

Public Property Get RowsSelected() As Long
    RowsSelected = miRowsSelected
End Property

Public Property Let RowsSelected(ByVal vData As Long)
    miRowsSelected = vData
    Me.txtRecordSelected = vData
End Property

Public Property Get StartRow() As Long
    StartRow = miStartRow
End Property

Public Property Let StartRow(ByVal vData As Long)
    miStartRow = vData
End Property

Sub cbAuditor_Click()
'*****************************************************************
'Created by Kathleen C Flanagan Tuesday 2/11/2013
'Will create a string that will be used in the WHERE clause of the Select SQL when the 'Apply Filters' is clicked
'*****************************************************************
If strListAuditors = "" Then
    strListAuditors = "'" & Me.cbAuditor & "'"
    MsgBox (strListAuditors)
Else: strListAuditors = strListAuditors + ", '" + Me.cbAuditor + "'"
    MsgBox (strListAuditors)
End If

End Sub

Sub cbDRG_Click()
'*****************************************************************
'Created by Kathleen C Flanagan Tuesday 3/21/2013
'Will create a string that will be used in the WHERE clause of the Select SQL when the 'Apply Filters' is clicked
'*****************************************************************
If strListDRGs = "" Then
    strListDRGs = "'" & Me.cbDRG & "'"
    MsgBox (strListDRGs)
Else: strListDRGs = strListDRGs + ", '" + Me.cbDRG + "'"
    MsgBox (strListDRGs)
End If

End Sub

Private Sub cmdClearFilter_Click()
'Clear out the filters & re-query the recordset
'KCF 11/25/2013 - Include the new 320.2 claim status
    
    'Claim Filters
    mstrFilter = ""
    strListAuditors = ""
    Me.cbAuditor = ""
    Me.cbAdj_Reviewtype = "" 'KCF 10/23/2012 to allow sort by Review Type (Complex or Prepay)
    Me.cbConcept = ""
    'BEGIN 3/21/2013 KCF: Replace DRG combo box with the string of DRGs
    Me.cbDRG = ""
    strListDRGs = ""
    'END 3/21/2013 KCF: Replace DRG combo box with the string of DRGs
    Me.cbTeam = ""
    Me.cbProvName = ""
    Me.cbCurPayerName = ""
    Me.cbStatus = "" 'KCF 11/7/2012 to allow filter by queuetype description
    Me.cbPredict = "" 'KCF 3/10/2014 to allow filter by Predictive Model
    Me.cbICDcode = "" 'Tivya to add ICD version filter
    'Sort Filters
    mstrSort = ""
    Me.cbSort1 = ""
    Me.cbSort2 = ""
    Me.cbSort3 = ""
    
    'ICN Search Filter
    Me.txtICN = "" 'KCF 2/22/2013 to clear the ICN search box
    
    '10/11/2012 use view to select records, 10/23/2012 set default sort to start with Review Type (Complex or Prepay)
    Me.frm_QA_Claims_List.Form.RecordSource = "select Top 1000 * from v_QA_Review_WorkTable_Unsubmitted WHERE AHclmstatus in ('314', '320', '320.2', '321', '322') order by Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN"
    Me.frm_QA_Claims_List.Form.Refresh
    
    Set_Filters_ComboBoxes
    
    If Me.frm_QA_Claims_List.Form.RecordSet.recordCount > 0 Then
        Me.RowsSelected = 1
    Else
        Me.RowsSelected = 0
    End If

End Sub

Private Sub cmdApplyFilters_Click()
'Query the recordset with the values selected in the filter drop downs
'Replace the value from cbAuditor with the strListAuditors value
'KCF 11/25/2013 - Update claim status filter to include the 320.2 claim status
    
    'Set the WHERE selections via the UI 'Select Claims' combo box
    mstrFilter = "AHclmstatus in ('314', '320', '320.2', '321', '322') "
    'BEGIN 2/11/2013 KCF: Replace cbAuditor with strListauditors
    If strListAuditors & "" <> "" Then
        mstrFilter = mstrFilter & "and Auditor in (" & strListAuditors & ")"
    End If
'    If Me.cbAuditor & "" <> "" Then
'        mstrFilter = mstrFilter & "and Auditor = '" & Me.cbAuditor & "' "
'    End If
    'END  2/11/2013 KCF: Replace cbAuditor with strListauditors
    'Add combo box to allow filter based upon Review Type (Complex or Prepay)
    If Me.cbAdj_Reviewtype & "" <> "" Then
        mstrFilter = mstrFilter & "and Adj_ReviewType = '" & Me.cbAdj_Reviewtype & " ' "
    End If
    If Me.cbConcept & "" <> "" Then
        mstrFilter = mstrFilter & "and ConceptID = '" & Me.cbConcept & "' "
    End If
    'BEGIN 3/21/2013 KCF: Replace DRG combobox with string of DRGs
    If strListDRGs & "" <> "" Then
        mstrFilter = mstrFilter & "and DRG in (" & strListDRGs & ")"
    End If
    'If Me.cbDRG & "" <> "" Then
    '    mstrFilter = mstrFilter & "and DRG = '" & Me.cbDRG & "' "
    'End If
    'END 3/21/2013 KCF: Replace DRG combobox with string of DRGs
    If Me.cbTeam & "" <> "" Then
        mstrFilter = mstrFilter & "and AuditTeam = '" & Me.cbTeam & "' "
    End If
    If Me.cbProvName & "" <> "" Then
        mstrFilter = mstrFilter & "and ProvName = '" & Me.cbProvName & "' "
    End If
    If Me.cbCurPayerName & "" <> "" Then
        mstrFilter = mstrFilter & "and CurPayerName = '" & Me.cbCurPayerName & "' "
    End If
    'BEGIN KCF 11/7/2012 to include filter on QueueType
    If Me.cbStatus & "" <> "" Then
        mstrFilter = mstrFilter & "and QueueTypeDescription = '" & Me.cbStatus & "' "
    End If
    'END KCF 11/7/2012 to include filter on QueueType
    'BEGIN KCF 3/10/2014 to include filter on Predictive Percentile
    If Me.cbPredict & "" <> "" Then
        mstrFilter = mstrFilter & "and Report_Stats_Score_Group >= " & Me.cbPredict & " "
    End If
    'END KCF 3/10/2014 to include filter on Predictive Percentile
    ' Tivya ICD version
    If Me.cbICDcode & "" <> "" Then
        mstrFilter = mstrFilter & "and IcdversionCDflag = " & Me.cbICDcode & " "
    End If
    ' end ICD version
    'Set the SORT options via the UI 'Sort Claims' combo boxes
    mstrSort = ""
    If Me.cbSort1 & "" <> "" Then
        mstrSort = mstrSort & Me.cbSort1
    End If
    If Me.cbSort2 & "" <> "" Then
        If mstrSort & "" <> "" Then
            mstrSort = mstrSort & ", " & Me.cbSort2
        Else
            mstrSort = mstrSort & Me.cbSort2
        End If
    End If
    If Me.cbSort3 & "" <> "" Then
        If mstrSort & "" <> "" Then
            mstrSort = mstrSort & ", " & Me.cbSort3
        Else
            mstrSort = mstrSort & Me.cbSort3
        End If
    End If
    
    If mstrSort & "" <> "" Then
        mstrSort = "ORDER BY " & mstrSort & ", ICN"
    Else
        mstrSort = "ORDER BY  Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN" 'kcf 10/23/2012 set the default to begin with Review Type (complex or Prepay)
    End If

    
    'Construct the SQL for the Claim List
    If mstrFilter <> "" Then
        'mstrFilter = Right(mstrFilter, Len(mstrFilter) - 4)
        '10/11/2012 change to use view instead of table as record source
        Me.frm_QA_Claims_List.Form.RecordSource = "select Top 1000 * from v_QA_Review_WorkTable_Unsubmitted where " & mstrFilter & mstrSort
        
        Me.frm_QA_Claims_List.Form.Refresh
        Set_Filters_ComboBoxes
    ElseIf mstrFilter & "" = "" Then
        '10/11/2012 change to use view instead of record source
        'BEGIN 3/21/2013 KCF: change to fix the WHERE clause - unsubmitted should not have a QAStatus
        'Me.frm_QA_Claims_List.Form.RecordSource = "Select Top 1000 * from v_QA_Review_Worktable_unsubmitted where QAStatus in ('C', 'A', 'R') " & mstrSort
        Me.frm_QA_Claims_List.Form.RecordSource = "Select Top 1000 * from v_QA_Review_Worktable_unsubmitted WHERE AHClmSTatus in ('314', '320', '320.2', '321', '322') " & mstrSort
        'END 3/21/2013 KCF: change to fix the WHERE clause - unsubmitted should not have a QAStatus
    End If
    
    If Me.frm_QA_Claims_List.Form.RecordSet.recordCount > 0 Then
        Me.RowsSelected = 1
    Else
        Me.RowsSelected = 0
        Me.frm_QA_Claim_Review_Result.Form.RecordSource = "Select * from v_QA_Review_Worktable_Unsubmitted where 1 = 2" 'kcf empty subform if no records returned in the list
    End If
    
End Sub


Private Sub cmdExit_Click()
'Update Friday 5/31/2013 to clear the lock fields on the qa table when the user simply exits the form.

    If Me.frm_QA_Claim_Review_Result.Form.Dirty = True Then
        Me.frm_QA_Claim_Review_Result.Form.Undo
    End If
    
      Me.frm_QA_Claim_Review_Result.Form.LockDt = Null
    Me.frm_QA_Claim_Review_Result.Form.LockUser = ""
    
    DoCmd.Close acForm, Me.Name
 
    
End Sub

Private Sub cmdFindICN_Click()
'*****************************************************************
'Created by Kathleen C Flanagan Tuesday 2/11/2013
'Will allows QA to find specific ICN
'*****************************************************************

If Me.txtICN & "" = "" Then
    MsgBox ("Please provide an ICN to search")
End If

Me.frm_QA_Claims_List.Form.RecordSource = "Select * from v_QA_Review_Worktable_unsubmitted where ICN = '" & Me.txtICN & "'"

If Me.frm_QA_Claims_List.Form.RecordSet.recordCount = 0 Then
    MsgBox ("Requested ICN was not found.")
End If

End Sub

Sub cmdRemoveReviews_Click()
'Form will allow user to process muliple claims at one time.

    Dim rs As RecordSet
    Dim i As Long
    Dim RowToStart As Long
    Dim RowsToProcess As Long
    Dim strCnlyClaimNum As String
    Dim intCnlyClaimSeqNo As Integer
    Dim strQAComment As String
    Dim strQACommentCmd As String
    Dim strQACommentDRGReassign As String
    Dim strQACommentUserID As String
    Dim strCnlyClaimNumToFind As String
    
    'Use SQL to check the status, etc.
    Dim MyAdo As New clsADO
    Dim QAStatCkRS As ADODB.RecordSet
    
    Dim strSQL As String
    
    mstrUserName = GetUserName()

    Set rs = Me.frm_QA_Claims_List.Form.RecordSet
    
    'BEGIN 9/28 - update to go back to list position after submit
    RowToStart = miStartRow
    RowsToProcess = miRowsSelected
    
    'If user starts multi-row selection from the bottom, will re-orient from the top
    If Me.frm_QA_Claims_List.Form.CurrentRecord > miStartRow Then
        rs.MoveFirst
        rs.Move RowToStart - 1
    End If
    
    'IF not on the first row, will identify the claimnum in the row above - will use this value to find that record after update
    If Me.frm_QA_Claims_List.Form.CurrentRecord > 1 Then
        rs.Move -1
        strCnlyClaimNumToFind = Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum
        rs.Move 1
    End If
    'END 9/28 - update to go back to list position after submit
    
    Me.cmdClearFilter.SetFocus
    
    'Will open up form to enter comments that can be submitted with the claims
    DoCmd.OpenForm "frm_QA_Comment", , , , , acDialog, "Waive"
    strQACommentCmd = Forms("frm_QA_Comment").txtCommentCmd
    strQAComment = Forms("frm_QA_Comment").txtQAComment
    DoCmd.Close acForm, "frm_QA_Comment", acSaveNo
    
    
    'BEGIN: If to handle whether user responds with submit or cancel
    If strQACommentCmd = "QACommentSubmit" Then
        
        'Begin processing the claims
        For i = RowToStart To RowToStart + (RowsToProcess - 1)
        
        Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum.SetFocus
    
        strCnlyClaimNum = Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum
        intCnlyClaimSeqNo = Me.frm_QA_Claim_Review_Result.Form.SeqNo
        
        Me.cmdClearFilter.SetFocus
      
        Me.cmdSubmitQA.Caption = "Submit to QA"
    
            If Me.frm_QA_Claim_Review_Result.Form.Dirty Then
            Me.frm_QA_Claim_Review_Result.Form.Dirty = False
            End If
            
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strSQL = "Select CnlyClaimNum, SeqNo, QAStatusCheck, DRGReassign from cms_auditors_code.dbo.v_QA_Review_Worktable_STatusCheck where CnlyClaimNum ='" & strCnlyClaimNum & "'"
        Set QAStatCkRS = MyAdo.OpenRecordSet(strSQL)
        
        If QAStatCkRS("DRGReassign") = "Y" Then
            DoCmd.OpenForm "frm_QA_Comment", , , , , acDialog, "DRGReassign"
            strQACommentCmd = Forms("frm_QA_Comment").txtCommentCmd
            strQACommentDRGReassign = Forms("frm_QA_Comment").txtQAComment
            strQACommentUserID = Forms("frm_QA_Comment").cboUserID
            DoCmd.Close acForm, "frm_QA_Comment", acSaveNo
        
            If strQACommentCmd = "QACommentSubmit" Then
                DRGReassign strCnlyClaimNum, mstrUserName, strQACommentUserID, strQACommentDRGReassign
                
                SubmitQA strCnlyClaimNum, intCnlyClaimSeqNo, strQAComment, mstrUserName
            Else
                'Do Nothing
            End If
        
        Else
            SubmitQA strCnlyClaimNum, intCnlyClaimSeqNo, strQAComment, mstrUserName
        End If
        
        rs.MoveNext
            
        Next
        
        MsgBox ("Records Saved")
        
        Set rs = Nothing
        
        If mstrFilter <> "" Then
            cmdApplyFilters_Click
        Else
            cmdClearFilter_Click
        End If
    
    'BEGIN 9/28 - update to go back to list position after submit
        Me.frm_QA_Claims_List.SetFocus
        If Me.frm_QA_Claims_List.Form.CurrentRecord > 0 Then
            Me.frm_QA_Claims_List.Form.CnlyClaimNum.SetFocus
            If strCnlyClaimNumToFind & "" <> "" Then
                DoCmd.FindRecord strCnlyClaimNumToFind
            End If
        End If
        Me.frm_QA_Claims_List.Form.AuditTeam.SetFocus
    'END 9/28 - update to go back to list position after submit
    
        Me.cmdClearFilter.SetFocus
        Me.cmdRemoveReviews.Enabled = False

    ElseIf strQACommentCmd = "QACommentCancel" Then
  
    End If
    'END: If to handle whether users responds with submit or cancel

        
 End Sub
        

Private Sub cmdSubmitQA_Click()
    Dim strCnlyClaimNum As String
    Dim intCnlyClaimSeqNo As Integer
    Dim strQAWaiveComment As String
    Dim strQACommentDRGReassign As String
    Dim strQACommentCmd As String
    Dim strQACommentUserID As String
    Dim RowToStart As Long
    
    'Use SQL to check the status, etc.
    Dim MyAdo As New clsADO
    Dim QAStatCkRS As ADODB.RecordSet
    
    Dim strSQL As String
    
    Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum.SetFocus
  
    'BEGIN 9/28 - update to go back to list position after submit
    RowToStart = Me.frm_QA_Claims_List.Form.CurrentRecord
    'END 9/28 - update to go back to list position after submit
    
    
    mstrUserName = GetUserName()
    strCnlyClaimNum = Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum
    intCnlyClaimSeqNo = Me.frm_QA_Claim_Review_Result.Form.SeqNo
    strQAWaiveComment = ""
     
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    'BEGIN 10/16/2012 include field for AHClmStatus in recordset
    strSQL = "Select CnlyClaimNum, SeqNo, QAStatusCheck, DRGReassign, AHClmStatus, QAStatusCheck from cms_auditors_code.dbo.v_QA_Review_Worktable_STatusCheck where CnlyClaimNum ='" & strCnlyClaimNum & "' and SeqNo = " & intCnlyClaimSeqNo
    'END 10/16/2012 include field for AHClmStatus in recordset
    Set QAStatCkRS = MyAdo.OpenRecordSet(strSQL)
     

    Me.cmdClearFilter.SetFocus
      
    Me.cmdSubmitQA.Caption = "Submit to QA"
            
    If QAStatCkRS.BOF = True And QAStatCkRS.EOF = True Then
        MsgBox ("There is an issue submitting this claim, so please try again.  If the problem persists, please contact a system administrator.")
        Exit Sub
    End If
            
    If QAStatCkRS("DRGReassign") = "Y" Then
        DoCmd.OpenForm "frm_QA_Comment", , , , , acDialog, "DRGReassign"
        strQACommentCmd = Forms("frm_QA_Comment").txtCommentCmd
        strQACommentDRGReassign = Forms("frm_QA_Comment").txtQAComment
        strQACommentUserID = Forms("frm_QA_Comment").cboUserID
        DoCmd.Close acForm, "frm_QA_Comment", acSaveNo
        
        If strQACommentCmd = "QACommentSubmit" Then
            DRGReassign strCnlyClaimNum, mstrUserName, strQACommentUserID, strQACommentDRGReassign
        Else
            Exit Sub
        End If
        
    End If
    
    
    If QAStatCkRS.recordCount <> 1 Then
        MsgBox ("There is an issue submitting this claim.  Please try again.  If the problem persists, please contact Claim Admin support.")
    'BEGIN 10/16/2012 include check that the QA staff have updated the Claim from 314 to appropriate Review status
    ElseIf QAStatCkRS("AHClmSTatus") = "314" And QAStatCkRS("QAStatusCheck") <> "A" Then
        MsgBox ("The current Claim is still 'Under Review by Medical Director'.  Please review the Claim in the main Claim Form and update as appropriate.  Then try again to submit.")
    'END 10/16/2012 include check that the QA staff have updated the Claim from 314 to appropriate Review status
    ElseIf QAStatCkRS("QAStatusCheck") = "I" Then
        MsgBox ("The current Claim QA record cannot be submitted. All Claim QA questions must be answered and a comment must be provided for any questions that have a 'No' response.  Please complete the Claim QA Review and re-submit.")
    ElseIf QAStatCkRS("QAStatusCheck") = "W" Then
        MsgBox ("The current Claim QA record cannnot be submitted. All Claim QA questions must be answered and a comment must be provided for any questions that have a 'No' response.  Please complete the Claim QA Review and re-submit.")
    ElseIf QAStatCkRS("QAStatusCheck") = "A" Then
         'Manually set the dirty value to False because of conflict with stored procedure after getting the ClaimNum and SeqNo to pass to update stored procedure
            If Me.frm_QA_Claim_Review_Result.Form.Dirty Then
                Me.frm_QA_Claim_Review_Result.Form.Dirty = False
            End If
         SubmitQA strCnlyClaimNum, intCnlyClaimSeqNo, strQAWaiveComment, mstrUserName
         
            If mstrFilter <> "" Then
                cmdApplyFilters_Click
            Else
                cmdClearFilter_Click
            End If
         
         MsgBox ("Record Saved")
         
    ElseIf QAStatCkRS("QAStatusCheck") = "C" Then
        'Manually set the dirty value to False because of conflict with stored procedure after getting the ClaimNum and SeqNo to pass to update stored procedure
            If Me.frm_QA_Claim_Review_Result.Form.Dirty Then
                   Me.frm_QA_Claim_Review_Result.Form.Dirty = False
            End If
        SubmitQA strCnlyClaimNum, intCnlyClaimSeqNo, strQAWaiveComment, mstrUserName
         
            If mstrFilter <> "" Then
            
                cmdApplyFilters_Click
            Else
                cmdClearFilter_Click
            End If
            
        MsgBox ("Record Saved")
    Else
        MsgBox ("Exception")
    End If
        Set QAStatCkRS = Nothing
      
    'BEGIN 9/28 - update to go back to list position after submit
    If miRowsSelected > 0 Then
        Me.frm_QA_Claims_List.Form.RecordSet.Move RowToStart - 1
    End If
    'END 9/28 - update to go back to list position after submit
    

  
    
End Sub

Public Sub DRGReassign(CnlyClaimNum As String, mstrUserName As String, strQACommentUserID As String, strQACommentDRGReassign As String)
'Calls stp that will send email alert when the MN claim should be DRG
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_QA_Review_ClaimReassign"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = CnlyClaimNum
    cmd.Parameters("@pReviewer") = mstrUserName
    cmd.Parameters("@pAssignedTo") = strQACommentUserID
    cmd.Parameters("@pReassignReason") = strQACommentDRGReassign
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_ClaimReassign"
        Err.Raise 50001, "usp_QA_Review_Worktable_Update", strErrMsg
    End If
    
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub


End Sub


Public Sub SubmitQA(CnlyClaimNum As String, SeqNo As Integer, strQAWaiveComment As String, mstrUserName As String)

    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_QA_Review_Worktable_Update"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = CnlyClaimNum
    cmd.Parameters("@pSeqNo") = SeqNo
    cmd.Parameters("@pStrQAWaiveComment") = strQAWaiveComment
    cmd.Parameters("@pmstrUserName") = mstrUserName
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_Worktable_Update"
        Err.Raise 50001, "usp_QA_Review_Worktable_Update", strErrMsg
    End If
    
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub




Private Sub Form_Load()
    Dim strCnlyClaimNum As String
    Dim intSeqNo As Integer
   
    Me.Caption = "Claim QA"
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
   
    mstrFilter = ""
    Set_Filters_ComboBoxes
        
    If mbDetailFormLoaded And mbGridFormLoaded Then
        strCnlyClaimNum = Me.frm_QA_Claims_List.Form.CnlyClaimNum
        intSeqNo = Me.frm_QA_Claims_List.Form.SeqNo
        Me.frm_QA_Claim_Review_Result.Form.RecordSource = "select * from QA_Review_Worktable where CnlyClaimNum = '" & strCnlyClaimNum & "' and SeqNo = " & intSeqNo
    End If
    
    Me.RowsSelected = 1
        
End Sub

Private Sub Set_Filters_ComboBoxes()
    If mstrFilter <> "" Then
    '10/11/2012 update to use the view instead of the table
        Me.cbTeam.RowSource = "select distinct AuditTeam from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "order by AuditTeam"
        Me.cbAuditor.RowSource = "select distinct Auditor from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "order by Auditor"
        Me.cbConcept.RowSource = "select distinct ConceptID from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by ConceptID"
        Me.cbDRG.RowSource = "select distinct DRG from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by DRG"
        Me.cbProvName.RowSource = "Select distinct ProvName from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by ProvName"
        'KCF 11/7/2012 - include ability to filter on QueueType
        Me.cbStatus.RowSource = "Select distinct QueueTypeDescription from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by QueueTypeDescription"
        Me.cbPredict.RowSource = "Select distinct Report_Stats_Score_Group from v_qa_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by Report_Stats_Score_Group desc"
        'Tivya ICDversion
        Me.cbICDcode.RowSource = "Select distinct IcdversionCDflag from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by IcdversionCDflag"
    Else
    '10/11/2012 update to use view instead of table reference
        Me.cbTeam.RowSource = "select distinct AuditTeam from v_QA_Review_Worktable_Unsubmitted order by AuditTeam"
        Me.cbAuditor.RowSource = "select distinct Auditor from v_QA_Review_Worktable_Unsubmitted order by Auditor"
        Me.cbConcept.RowSource = "select distinct ConceptID from v_QA_Review_Worktable_Unsubmitted order by ConceptID"
        Me.cbDRG.RowSource = "select distinct DRG from v_QA_Review_Worktable_Unsubmitted order by DRG"
        Me.cbProvName.RowSource = "Select distinct ProvName from v_QA_Review_Worktable_Unsubmitted order by ProvName"
        'KCF 11/7/2012 - include ability to filter onQueueType
        Me.cbStatus.RowSource = "Select distinct QueueTypeDescription from v_qa_Review_Worktable_Unsubmitted order by QueueTypeDescription"
        Me.cbPredict.RowSource = "Select distinct Report_Stats_Score_Group from v_qa_Review_Worktable_Unsubmitted Order by Report_Stats_Score_Group desc"
       Me.cbICDcode.RowSource = "Select distinct IcdversionCDflag from v_qa_Review_Worktable_Unsubmitted order by IcdversionCDflag"
    End If
    
End Sub

Private Sub Command87_Click()
On Error GoTo Err_Command87_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_QA_Review_Main_submitted"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command87_Click:
    Exit Sub

Err_Command87_Click:
    MsgBox Err.Description
    Resume Exit_Command87_Click
    
End Sub
