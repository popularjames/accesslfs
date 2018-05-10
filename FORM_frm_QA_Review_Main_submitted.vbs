Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'******************************************************************************************************************
'Claim QA version 2.0 Update; Created by Kathleen C Flanagan; Wednesday 2/27/2013
'
'Description: new form created to provide functionality to allow QA Team to update previously submitted QA Scores
'and\or (re-)send claims back to Auditors
'
'Form is access from a button on the frm_QA_Review_Main
'Can call the stored procedures usp_QA_Review_Worktable_QAScore_Update and usp_QA_Review_ReturnAuditor_Alert
'Relies on v_QA_Review_WorkTable_Submitted
'******************************************************************************************************************

'Private WithEvents frm_QA_Review_Result_Collection As Form_frm_QA_Review_Result_Collection KCF 2/27/2013: never implemented
'Private WithEvents frmReAssignSelect As Form_frm_QUEUE_ReAssign_Select KCF 2/27/2013: never implemented

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
Public strQAsubmitDtTo As String

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

Private Sub cmdExit_Click()
Dim strErrMsg
'Update Friday 5/31/2013 to clear the lock fields in the qa table when the user exits.

    If Me.frm_QA_Claim_Review_Result.Form.Dirty = True Then
        Me.frm_QA_Claim_Review_Result.Form.Undo
    End If
    
    'Me.frm_QA_Claim_Review_Result.Form.LockDt = Null
    'Me.frm_QA_Claim_Review_Result.Form.LockUser = ""
    
    DoCmd.Close acForm, Me.Name
    
Exit_Sub:
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
   
    Me.Caption = "Claim QA - Completed Claims"
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
   
    mstrFilter = ""
    Set_Filters_ComboBoxes
       
'3/21/2013 FIX
If Not (Me.RecordSet Is Nothing) Then
     If mbDetailFormLoaded And mbGridFormLoaded Then
         strCnlyClaimNum = Nz(Me.frm_QA_Claims_List.Form.CnlyClaimNum, "")
         intSeqNo = Me.frm_QA_Claims_List.Form.SeqNo
       
         Me.frm_QA_Claim_Review_Result.Form.RecordSource = "select * from QA_Review_Worktable where CnlyClaimNum = '" & strCnlyClaimNum & "' and SeqNo = " & intSeqNo
         Me.frm_QA_Claim_Review_Result.Form.txtAHClmStatus = Me.frm_QA_Claims_List.Form.AHClmstatus 'KCF 2/15/2013
     End If
End If
     
    'Disabled the command buttons until the page is dirty and/or the criteria to allow return (clmstatus is still Recovery)
    Me.cmdReturnAuditor.Enabled = False
    Me.cmdUpdateQA.Enabled = False
     
    Me.RowsSelected = 1
    Me.Refresh
        
End Sub


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
    Me.txtQASubmitDtFrom = "" 'KCF 2/22/2013 to allow search by submit date
    Me.txtQASubmitDtTo = "" 'KCF 2/22/2013 to allow search by submit date
    
    'Sort Filters
    mstrSort = ""
    Me.cbSort1 = ""
    Me.cbSort2 = ""
    Me.cbSort3 = ""
    
    'ICN Search Filter
    Me.txtICN = "" 'KCF 2/22/2013 to clear the ICN search box
    
    '2/22/2013 use view to select records, 10/23/2012 set default sort to start with SubmitDate
    Me.frm_QA_Claims_List.Form.RecordSource = "Select Top 1000 * from v_QA_Review_WorkTable_Submitted where QAStatus in ('A', 'C', 'R') and submitdate > Now() - 5 order by SubmitDate desc, Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN"
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
    
    'Set the WHERE selections via the UI 'Select Claims' combo box
    mstrFilter = ""
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
    'BEGIN KCF 2/22/2013 to include date search on SubmitDate
    If (Nz(Me.txtQASubmitDtFrom, "") <> "" And Nz(Me.txtQASubmitDtTo, "") = "") Or (Nz(Me.txtQASubmitDtFrom, "") = "" And Nz(Me.txtQASubmitDtTo, "") <> "") Then
        MsgBox ("Please provide both a Date From and a Date To.")
    ElseIf (Nz(Me.txtQASubmitDtFrom, "") <> "" And Nz(Me.txtQASubmitDtTo, "") <> "") Then
            strQAsubmitDtTo = CDate(Me.txtQASubmitDtTo) + 1 'KCF 2/22/2013 increment the to date up by 1 day; will handle the datetime when users enters the same date for To & From
            mstrFilter = mstrFilter & "and (SubmitDate > #" & Me.txtQASubmitDtFrom & "# and SubmitDate < #" & strQAsubmitDtTo & "#) "
    End If
    'END KCF 2/22/2013 to include date search on SubmitDate
    
    
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
        mstrSort = "ORDER BY  SubmitDate desc, Adj_ReviewType desc, MRReceivedDt, AuditTeam, Auditor, ICN" 'kcf 2/22/2013 set the default to begin with SubmitDate
    End If
    
    'Construct the SQL for the Claim List
    If mstrFilter <> "" Then
        mstrFilter = Right(mstrFilter, Len(mstrFilter) - 4)
        '10/11/2012 change to use view instead of table as record source
        Me.frm_QA_Claims_List.Form.RecordSource = "select Top 1000 * from v_QA_Review_WorkTable_submitted where " & mstrFilter & mstrSort
        If Me.frm_QA_Claims_List.Form.RecordSet.recordCount > 0 Then
            Me.frm_QA_Claims_List.Form.Refresh
        Else
            MsgBox ("Your search criteria found no records. Please try again.")
        End If
        Set_Filters_ComboBoxes
    ElseIf mstrFilter & "" = "" Then
        '10/11/2012 change to use view instead of record source
        Me.frm_QA_Claims_List.Form.RecordSource = "Select Top 1000 * from v_QA_Review_Worktable_submitted " & mstrSort
    End If
    
    If Me.frm_QA_Claims_List.Form.RecordSet.recordCount > 0 Then
        Me.RowsSelected = 1
    Else
        Me.RowsSelected = 0
        Me.frm_QA_Claim_Review_Result.Form.RecordSource = "Select * from v_QA_Review_Worktable_submitted where 1 = 2" 'kcf empty subform if no records returned in the list"
    End If
    
End Sub

Private Sub cmdFindICN_Click()
'*****************************************************************
'Created by Kathleen C Flanagan Tuesday 2/11/2013
'Will allows QA to find specific ICN
'*****************************************************************

If Me.txtICN & "" = "" Then
    MsgBox ("Please provide an ICN to search")
End If

Me.frm_QA_Claims_List.Form.RecordSource = "Select * from v_QA_Review_Worktable_submitted where ICN = '" & Me.txtICN & "'"

If Me.frm_QA_Claims_List.Form.RecordSet.recordCount = 0 Then
    MsgBox ("Requested ICN was not found.")
End If

End Sub


Private Sub Set_Filters_ComboBoxes()
    If mstrFilter <> "" Then
    '10/11/2012 update to use the view instead of the table
        Me.cbTeam.RowSource = "select distinct AuditTeam from v_QA_Review_Worktable_submitted where " & mstrFilter & "order by AuditTeam"
        Me.cbAuditor.RowSource = "select distinct Auditor from v_QA_Review_Worktable_submitted where " & mstrFilter & "order by Auditor"
        Me.cbConcept.RowSource = "select distinct ConceptID from v_QA_Review_Worktable_submitted where " & mstrFilter & "Order by ConceptID"
        Me.cbDRG.RowSource = "select distinct DRG from v_QA_Review_Worktable_submitted where " & mstrFilter & "Order by DRG"
        Me.cbProvName.RowSource = "Select distinct ProvName from v_QA_Review_Worktable_submitted where " & mstrFilter & "Order by ProvName"
        'KCF 11/7/2012 - include ability to filter on QueueType
        Me.cbStatus.RowSource = "Select distinct QueueTypeDescription from v_QA_Review_Worktable_Unsubmitted where " & mstrFilter & "Order by QueueTypeDescription"
    Else
    '10/11/2012 update to use view instead of table reference
        Me.cbTeam.RowSource = "select distinct AuditTeam from v_QA_Review_Worktable_submitted order by AuditTeam"
        Me.cbAuditor.RowSource = "select distinct Auditor from v_QA_Review_Worktable_submitted order by Auditor"
        Me.cbConcept.RowSource = "select distinct ConceptID from v_QA_Review_Worktable_submitted order by ConceptID"
        Me.cbDRG.RowSource = "select distinct DRG from v_QA_Review_Worktable_submitted order by DRG"
        Me.cbProvName.RowSource = "Select distinct ProvName from v_QA_Review_Worktable_submitted order by ProvName"
        'KCF 11/7/2012 - include ability to filter onQueueType
        Me.cbStatus.RowSource = "Select distinct QueueTypeDescription from v_qa_Review_Worktable_Unsubmitted order by QueueTypeDescription"
    End If
    
End Sub


Private Sub cmdReturnAuditor_Click()
    Dim strCnlyClaimNum As String
    Dim intCnlyClaimSeqNo As Integer
    Dim strQACommentCmd As String
    Dim strQACommentUserID As String
    
    'Use SQL to check the status, etc.
    Dim MyAdo As New clsADO
    Dim QAStatCkRS As ADODB.RecordSet
    
    Dim strSQL As String
    
    Dim strErrMsg As String
    
    'On Error GoTo Err_Handler
    
    mstrUserName = GetUserName()
    strCnlyClaimNum = Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum
    intCnlyClaimSeqNo = Me.frm_QA_Claim_Review_Result.Form.SeqNo
    
    Me.frm_QA_Claim_Review_Result.Form.txtQAStatus = "R"
    Me.frm_QA_Claim_Review_Result.Form.Refresh
    'Me.frm_QA_Claim_Review_Result.Form.mbFormDirty = False
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    'BEGIN 10/16/2012 include field for AHClmStatus in recordset
    strSQL = "Select CnlyClaimNum, SeqNo, QAStatus, AHClmStatus from cms_auditors_code.dbo.v_QA_Review_Worktable_Submitted where CnlyClaimNum ='" & strCnlyClaimNum & "' and SeqNo = " & intCnlyClaimSeqNo
    'END 10/16/2012 include field for AHClmStatus in recordset
    Set QAStatCkRS = MyAdo.OpenRecordSet(strSQL)

    Me.cmdClearFilter.SetFocus

    SubmitQA strCnlyClaimNum, intCnlyClaimSeqNo, mstrUserName
             
    If mstrFilter <> "" Then
        cmdApplyFilters_Click
    Else
        cmdClearFilter_Click
    End If
             
    MsgBox ("Claim Returned to Auditor")

    Me.cmdReturnAuditor.Enabled = False

Exit_Sub:
    Set MyAdo = Nothing
    'Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

    
End Sub

Public Sub SubmitQA(CnlyClaimNum As String, SeqNo As Integer, mstrUserName As String)

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
    cmd.Parameters("@pmstrUserName") = mstrUserName
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_Worktable_Update"
        Err.Raise 50001, "usp_QA_Review_Worktable_Update", strErrMsg
    End If
    
    Me.frm_QA_Claim_Review_Result.Form.LockDt = Null
    Me.frm_QA_Claim_Review_Result.Form.LockUser = ""
    
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


Sub cmdUpdateQA_Click()
    Call UpdateQARecord
    Me.cmdClearFilter.SetFocus
    Me.cmdUpdateQA.Enabled = False
End Sub

Public Sub UpdateQARecord()
 
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    
    Dim strCnlyClaimNum As String
    Dim intCnlyClaimSeqNo As Integer
    Dim strQAComment As String
    Dim strQACommentCmd As String
    Dim strQACommentUserID As String
    
    Dim strErrMsg As String
    
    'On Error GoTo Err_Handler

    mstrUserName = GetUserName()
    strCnlyClaimNum = Me.frm_QA_Claim_Review_Result.Form.CnlyClaimNum
    intCnlyClaimSeqNo = Me.frm_QA_Claim_Review_Result.Form.SeqNo
    
     'Will open up form to enter comments that can be submitted with the claims
    DoCmd.OpenForm "frm_QA_Comment", , , , , acDialog, "UpdateQA"
    strQACommentCmd = Forms("frm_QA_Comment").txtCommentCmd
    strQAComment = Forms("frm_QA_Comment").txtQAComment
    DoCmd.Close acForm, "frm_QA_Comment", acSaveNo
    
    'Handle cancel
    If strQACommentCmd = "QACommentCancel" Then
        'Me.frm_QA_Claim_Review_Result.Form.Form_BeforeUpdate (True)
        Me.frm_QA_Claim_Review_Result.Form.Undo
        Exit Sub
    End If
    
    'Me.frm_QA_Claim_Review_Result.Dirty = False
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_QA_Review_Worktable_QAScore_Update"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
    cmd.Parameters("@pSeqNo") = intCnlyClaimSeqNo
    cmd.Parameters("@pmstrUserName") = mstrUserName
    cmd.Parameters("@pMNPertLab") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNPertLab, "")
    cmd.Parameters("@pMNPhysOrder") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNPhysOrder, "")
    cmd.Parameters("@pMNCompleteMR") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNGrammar, "")
    cmd.Parameters("@pMNCorrectDecision") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNCorrectDecision, "")
    cmd.Parameters("@pMNCodingCorrect") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNCodingCorrect, "")
    cmd.Parameters("@pRationaleCorrect") = Nz(Me.frm_QA_Claim_Review_Result.Form.RationaleCorrect, "")
    cmd.Parameters("@pMNGrammar") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbMNGrammar, "")
    cmd.Parameters("@pDRGCorrectDecision") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbDRGCorrectDecision, "")
    cmd.Parameters("@pDRGCorrect") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbDRGCorrect, "")
    cmd.Parameters("@pDRGCorrectDischarge") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbDRGCorrectDischarge, "")
    cmd.Parameters("@pDRGCodingChange") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbDRGCodingChange, "")
    cmd.Parameters("@pDRGClaimReferMN") = Nz(Me.frm_QA_Claim_Review_Result.Form.cbDRGClaimReferMN, "")
    cmd.Parameters("@pQAUpdate_Comment") = strQAComment
    cmd.Parameters("@pPrevQAScore") = Me.frm_QA_Claim_Review_Result.Form.QAScore
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_QA_Review_Worktable_QAScore_Update"
        Err.Raise 50001, "usp_QA_Review_Worktable_QAScore_Update", strErrMsg
    End If
    
    MsgBox ("QA Record updated.")
 
    
    
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
