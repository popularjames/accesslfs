Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim frmResult As Form_frm_GENERAL_Generic_StatusPopup
Const CstrFrmAppID As String = "QueueMgt"
Private miAppPermission As Integer

Private mstrUserName As String
Option Explicit

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Sub RefreshData()
    Dim strSQL As String
    strSQL = " SELECT SupervisorID, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims GROUP BY SupervisorID"
    RefreshComboBox strSQL, Me.cboTeam
End Sub

Private Sub cboTeam_Click()
    Dim strSQL As String
    strSQL = " SELECT FromAuditor, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where supervisorID = '" & Nz(Me.cboTeam, "") & "' GROUP BY FromAuditor"
    RefreshListBox strSQL, Me.lstAuditors


    'removed 7/12 strSQL = " SELECT UserID, SupervisorID from Admin_User where SupervisorID LIKE '" & Me.cboTeam & "*'"
    
    If Me.cboTeam = "CNLY_MN" Then
      strSQL = " SELECT UserID, SupervisorID from Admin_User where SupervisorID = 'CNLY MN 1' or Supervisorid = 'CNLY MN 2' or supervisorid = 'CNLY MN 3' or supervisorid = 'CNLY MN 4'"
    Else
      strSQL = " SELECT UserID, SupervisorID from Admin_User where SupervisorID LIKE '" & Me.cboTeam & "*'"
    End If
    
    RefreshListBox strSQL, Me.lstAssignTo


    strSQL = " SELECT ProvStCd, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where supervisorID = '" & Nz(Me.cboTeam, "") & "' GROUP BY ProvStCd"
    RefreshListBox strSQL, Me.lstState
End Sub
Private Sub chkAllAge_Click()
If Me.chkAllAge.Value <> 0 Then
    Me.lstAge.Enabled = False
Else
    Me.lstAge.Enabled = True
End If
    Me.txtCount = GetCountOfSelection
End Sub

Private Sub chkAllDRG_Click()
    If Me.chkAllDRG.Value <> 0 Then
        Me.lstDRG.Enabled = False
    Else
        Me.lstDRG.Enabled = True
    End If
    
    RefreshAgeList

    Me.txtCount = GetCountOfSelection

End Sub

Private Sub cmdAging_Click()
    DoCmd.OpenForm "frm_QUEUE_AuditorAssign_Aging", acFormDS
End Sub

Private Sub cmdAssign_Click()
    Dim fmrStatus As Form_ScrStatus
    Dim lngProgressCount As Long
    Dim sMsg As String
    Dim strErrMsg As String
    Dim intStatus As Integer
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    Dim varItemSelected As Variant
    Dim strAuditor As String
    Dim i As Integer
    Dim j As Integer
    Dim strResults As String
    Dim lngClaimCap As Long
    Dim bResult As Boolean
    Dim lngCount As Long
    Dim lngtotal As Long
    Dim lngRecordTracker As Long
    Dim intClaimCap As Long
    Dim strAuditorList As String
    Dim strDRGList As String
    Dim strAgeList As String
        
    On Error GoTo ErrHandler

    Dim mstrAuditorsArray() As String
    ReDim mstrAuditorsArray(0)

    Set frmResult = Nothing
    intClaimCap = Nz(Me.txtCap, 0)
        
    For Each varItemSelected In Me.lstAssignTo.ItemsSelected
        ReDim Preserve mstrAuditorsArray(UBound(mstrAuditorsArray) + 1)
        mstrAuditorsArray(UBound(mstrAuditorsArray)) = Me.lstAssignTo.Column(0, varItemSelected)
    Next varItemSelected

    lngCount = 0
    strAuditorList = BuildMultiSelectStringFromList(Me.lstAuditors, True)
    If strAuditorList = "" Then
        Exit Sub
    End If
    
    strSQL = " SELECT  * FROM v_QUEUE_AuditorAssign_Claims where FromAuditor IN " & strAuditorList
    
    Dim strStateList As String
    strStateList = BuildMultiSelectStringFromList(Me.lstState, True)
    If Not strStateList = "" Then
        strSQL = strSQL & " and ProvStCd IN " & strStateList
    End If
   
   
   
    If Me.chkAllDRG.Value = 0 Then
        strDRGList = BuildMultiSelectStringFromList(Me.lstDRG, True)
        If strDRGList <> "" Then
            strSQL = strSQL & " AND DRG IN " & strDRGList
        End If
    Else
        strDRGList = ""
    End If
    If Me.chkAllAge.Value = 0 Then
        strAgeList = BuildMultiSelectStringFromList(Me.lstAge, False)
        If strAgeList <> "" Then
            strSQL = strSQL & " AND MRReceived_Age IN " & strAgeList
        End If
    Else
        strAgeList = ""
    End If
    
    strSQL = strSQL & " ORDER BY  MRReceived_Age  desc "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL)
    rst.MoveLast
    rst.MoveFirst
    lngtotal = rst.recordCount
    
    'DPR - top command in msaccess is busted doing it manually - FML
    If lngtotal > intClaimCap Then
        lngtotal = intClaimCap
    End If
       
    
    If MsgBox("You are about to reassign " & CStr(lngtotal) & " claims to " & CStr(UBound(mstrAuditorsArray)) & " auditors. Are you sure?", vbYesNo + vbQuestion) = vbNo Then
        Err.Raise 65000, , "User Canceled"
    End If
    
    intStatus = 1
    Set fmrStatus = New Form_ScrStatus
    With fmrStatus
         .ShowCancel = True
         .ShowMessage = False
         .ShowMessage = True
         .ProgVal = 0
         .ProgMax = lngtotal
         .TimerInterval = 50
         .show
         .visible = True
     End With
    
    
    strErrMsg = ""
    strResults = ""
    j = 1
    lngRecordTracker = 0
    
    While Not rst.EOF
        If j > UBound(mstrAuditorsArray) Then
            j = 1
        End If
            
        If lngRecordTracker >= lngtotal Then
            GoTo WeAreDone
        End If
        
        strAuditor = mstrAuditorsArray(j)
        sMsg = "Processing Claim Num :: " & rst!CnlyClaimNum & " - " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        DoEvents
        'just to test
'        Wait 10
        
        bResult = AssignClaim(rst!CnlyClaimNum, Identity.UserName, strAuditor, strErrMsg)
        If bResult Then
            lngCount = lngCount + 1
        Else: End If
        strResults = strResults & strErrMsg & vbCrLf
        intStatus = intStatus + 1
        j = j + 1
        lngRecordTracker = lngRecordTracker + 1
        rst.MoveNext
        
        
        If fmrStatus.EvalStatus(2) = True Then
                sMsg = "Assignment Canceled!"
                fmrStatus.StatusMessage sMsg
                DoEvents
                strErrMsg = sMsg
                GoTo WeAreDone
        End If
        
        
        
        
    Wend
    
WeAreDone:     strResults = strResults & vbCrLf & "*******" & CStr(lngCount) & " claims assigned." & CStr(lngtotal - lngCount) & " claims failed due to errors.*******"
    
    
    Set frmResult = New Form_frm_GENERAL_Generic_StatusPopup
    
    frmResult.TextData = strResults
    frmResult.TextLabel = "Assignment Results"
    frmResult.RefreshData
    frmResult.visible = True
    Me.RefreshData
    
Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub cmdRefresh_Click()
    Dim strSQL As String
    strSQL = " SELECT FromAuditor, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where supervisorID = '" & Nz(Me.cboTeam, "") & "' GROUP BY FromAuditor"
    RefreshListBox strSQL, Me.lstAuditors


    strSQL = " SELECT UserID, SupervisorID from Admin_User where SupervisorID LIKE '" & Me.cboTeam & "*'"
    RefreshListBox strSQL, Me.lstAssignTo


    strSQL = " SELECT ProvStCd, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where supervisorID = '" & Nz(Me.cboTeam, "") & "' GROUP BY ProvStCd"
    RefreshListBox strSQL, Me.lstState
    
    
    
    If Me.Frame28.Value = 1 Then
        RefreshStateList
        'DRG
        RefreshDRGList
        RefreshAgeList
        Me.txtCount = GetCountOfSelection
    Else
    
        RefreshStateList
        'AGE
        RefreshAgeList
        RefreshDRGList
        Me.txtCount = GetCountOfSelection
    End If
End Sub

Private Sub Form_Load()
    
Dim MyAdo As clsADO

    Call Account_Check(Me)

    Dim rsPermission As ADODB.RecordSet

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    mstrUserName = Identity.UserName

    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub

    
    
    
    
    
    
    Me.cboTeam = ""
    Me.chkAllAge.Value = 0
    Me.chkAllDRG.Value = 0
    Me.lstAge.Enabled = True
    Me.lstDRG.Enabled = True
    Me.txtCount = 0
    Me.Frame28.Value = 2
    RefreshData
End Sub

Private Sub Frame28_Click()
    Dim intAgeLeft As Integer
    Dim intDRGLeft As Integer
    Dim intDRGRight As Integer
    Dim intAgeRight As Integer
    
    intAgeLeft = Me.lstAge.left
    intDRGLeft = Me.lstDRG.left
    
    If Me.Frame28.Value = 1 Then
        'DRG
        Me.lstDRG.left = Me.bxAuditor.left + Me.bxAuditor.Width + 500
        Me.chkAllDRG.left = Me.lstDRG.left
        Me.lstAge.left = lstDRG.left + lstDRG.Width + 500
        Me.chkAllAge.left = Me.lstAge.left
        Me.lblAllAge.left = Me.chkAllAge.left + 500
        Me.lblAllDRG.left = Me.chkAllDRG.left + 500
        
    Else
        Me.lstAge.left = Me.bxAuditor.left + Me.bxAuditor.Width + 500
        Me.lstDRG.left = lstAge.left + lstAge.Width + 500
        Me.chkAllAge.left = Me.lstAge.left
        Me.chkAllDRG.left = Me.lstDRG.left
        
        Me.lblAllAge.left = Me.chkAllAge.left + 500
        Me.lblAllDRG.left = Me.chkAllDRG.left + 500
        
        
    End If
    
      

    If Me.Frame28.Value = 1 Then
        'DRG
        RefreshDRGList
        RefreshAgeList
        Me.txtCount = GetCountOfSelection
    Else
    
        'AGE
        RefreshAgeList
        RefreshDRGList
        Me.txtCount = GetCountOfSelection
    End If

End Sub

Private Sub lstAge_Click()
    If Me.Frame28.Value = 2 Then
    'AGE IS PRIMARY
        RefreshDRGList
    End If
         Me.txtCount = GetCountOfSelection
End Sub

Private Sub lstAuditors_Click()
    

    If Me.Frame28.Value = 1 Then
        RefreshStateList
        'DRG
        RefreshDRGList
        RefreshAgeList
        Me.txtCount = GetCountOfSelection
    Else
    
        RefreshStateList
        'AGE
        RefreshAgeList
        RefreshDRGList
        Me.txtCount = GetCountOfSelection
    End If

End Sub
Public Function BuildMultiSelectStringFromList(lstBox As listBox, bString As Boolean) As String
     
    On Error GoTo ErrHandler
     
    Dim intI As Long
    Dim bFirstTime As Boolean
    Dim IntCntr As Long
    Dim strString As String
    Dim intCtr As Long
     
    bFirstTime = True
    IntCntr = 0
     
    For intI = 0 To lstBox.ListCount
        If lstBox.Selected(intI) = True Then
            
            If bFirstTime Then
                If bString Then
                  strString = "( '" & lstBox.ItemData(intI) & "'"
                Else
                  strString = "(" & lstBox.ItemData(intI)
                End If
                bFirstTime = False
                IntCntr = IntCntr + 1
            Else
                If bString Then
                    strString = strString & ", '" & lstBox.ItemData(intI) & "'"
                Else
                    strString = strString & "," & lstBox.ItemData(intI)
                End If
                IntCntr = intCtr + 1
            End If
        End If
    Next intI
     
    If IntCntr <> 0 Then
      strString = strString & ")"
    Else
      strString = ""
    End If
    BuildMultiSelectStringFromList = strString
     
    Exit Function
ErrHandler:
        BuildMultiSelectStringFromList = ""
End Function

Private Sub lstDRG_Click()
    If Me.Frame28.Value = 1 Then
    'DRG IS PRIMARY
        RefreshAgeList
   
    End If
     Me.txtCount = GetCountOfSelection
End Sub
Private Sub RefreshStateList()
    
    Dim strSQL As String
    Dim strAgeList As String
    Dim strStateList As String
    
    
    
    If Not BuildMultiSelectStringFromList(Me.lstAuditors, True) = "" Then
    
    
        strSQL = " SELECT ProvStCd, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where supervisorID = '" & Nz(Me.cboTeam, "") & "' and FromAuditor IN " & BuildMultiSelectStringFromList(Me.lstAuditors, True) & " GROUP BY ProvStCd"
    '    strSQL = strSQL & " where FromAuditor IN " & BuildMultiSelectStringFromList(Me.lstAuditors, True)
        RefreshListBox strSQL, Me.lstState
    End If
    
    
    
    
End Sub


Private Sub RefreshDRGList()
    
    Dim strSQL As String
    Dim strAgeList As String
    Dim strStateList As String
    
    
    
    
    strSQL = " SELECT DRG ,MSDRGDesc , SUm(1) as Count "
    strSQL = strSQL & " FROM v_QUEUE_AuditorAssign_Claims "
    strSQL = strSQL & " where FromAuditor IN " & BuildMultiSelectStringFromList(Me.lstAuditors, True)
    
    strStateList = BuildMultiSelectStringFromList(Me.lstState, True)
    If Not strStateList = "" Then
        strSQL = strSQL & " and ProvStCd IN " & strStateList
    End If
    
    
    
    If Me.Frame28.Value = 2 Then
       'AGE IS PRIMARY
       strAgeList = BuildMultiSelectStringFromList(Me.lstAge, False)
        If strAgeList <> "" Then
            strSQL = strSQL & " AND MRReceived_Age IN " & strAgeList
        End If
    End If
    strSQL = strSQL & " GROUP BY DRG ,MSDRGDesc order by drg  "
    Me.lstDRG.RowSource = strSQL
    Me.txtCount = GetCountOfSelection

End Sub
Private Sub RefreshAgeList()
    Dim strDRGList As String
    Dim strAuditorList As String
    Dim strSQL As String
    Dim strStateList As String
    
    strAuditorList = BuildMultiSelectStringFromList(Me.lstAuditors, True)
    strSQL = " SELECT MRReceived_Age, SUm(1) as Count FROM v_QUEUE_AuditorAssign_Claims where FromAuditor IN " & strAuditorList
   
    strStateList = BuildMultiSelectStringFromList(Me.lstState, True)
    If Not strStateList = "" Then
        strSQL = strSQL & " and ProvStCd IN " & strStateList
    End If
  
   
    If Me.Frame28.Value = 1 Then
    'DRG IS PRIMARY
        If Me.chkAllDRG.Value = 0 Then
            strDRGList = BuildMultiSelectStringFromList(Me.lstDRG, True)
            If strDRGList <> "" Then
                strSQL = strSQL & " AND DRG IN " & strDRGList
            End If
        Else
            strDRGList = ""
        End If
    End If
     
        strSQL = strSQL & " GROUP BY MRReceived_Age order by MRReceived_Age"
    
    'RefreshListBox strSQL, Me.lstDRG
    Me.lstAge.RowSource = strSQL
    Me.txtCount = GetCountOfSelection
End Sub
Private Function GetCountOfSelection() As Long

    Dim strDRGList As String
    Dim strAuditorList As String
    Dim strAgeList As String
    Dim rst As DAO.RecordSet
    
    Dim strSQL As String
    
    strAuditorList = BuildMultiSelectStringFromList(Me.lstAuditors, True)
    
    If strAuditorList = "" Then
        GetCountOfSelection = 0
        Exit Function
    End If
    
    strSQL = " SELECT SUm(1) as Cnt FROM v_QUEUE_AuditorAssign_Claims where FromAuditor IN " & strAuditorList
   
    Dim strStateList As String
    strStateList = BuildMultiSelectStringFromList(Me.lstState, True)
    If Not strStateList = "" Then
        strSQL = strSQL & " and ProvStCd IN " & strStateList
    End If
   
   
    If Me.chkAllDRG.Value = 0 Then
        strDRGList = BuildMultiSelectStringFromList(Me.lstDRG, True)
        If strDRGList <> "" Then
            strSQL = strSQL & " AND DRG IN " & strDRGList
        End If
    Else
        strDRGList = ""
    End If
    
    
    If Me.chkAllAge.Value = 0 Then
        strAgeList = BuildMultiSelectStringFromList(Me.lstAge, False)
        If strAgeList <> "" Then
            strSQL = strSQL & " AND MRReceived_Age IN " & strAgeList
        End If
    Else
        strAgeList = ""
    End If
    
    
    
    Set rst = CurrentDb.OpenRecordSet(strSQL)
    If Not rst.EOF Then
        GetCountOfSelection = Nz(rst!Cnt, 0)
    Else
        GetCountOfSelection = 0
    End If
    



Set rst = Nothing




End Function
Private Function AssignClaim(strCnlyClaimNum As String, _
                             strAssignedFrom As String, _
                             strAssignedTo As String, _
                             ByRef strErrMsg As String) As Boolean


    Dim rsQueueHdr As ADODB.RecordSet
    Dim rsNote As ADODB.RecordSet
    Dim iNoteID As Long
    Dim bResult As Boolean
    Dim strLockUserID As String
    Dim strAction As String
    Dim strQueueStatus As String
    Dim strRationale As String
    Dim rsReassignCheck As ADODB.RecordSet
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim dChkDate As Date
    Dim strStatus As String
    
    On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_data_Database")
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
      
    MyAdo.sqlString = "select * from QUEUE_Hdr where CnlyClaimNum = '" & strCnlyClaimNum & "'"
    Set rsQueueHdr = MyAdo.OpenRecordSet
        
    myCode_ADO.BeginTrans

    If rsQueueHdr.BOF = True And rsQueueHdr.EOF = True Then
        strErrMsg = "Item " & strCnlyClaimNum & " is no longer in queue!"
        AssignClaim = False
        GoTo Rollback
    Else
        ' check if record has been update by someone else
        dChkDate = rsQueueHdr("LastUpdate")
        strQueueStatus = rsQueueHdr("QueueStatus")
        If strStatus = "FORWARD" Then
            strErrMsg = "Can not reassign claim " & strCnlyClaimNum & "." & vbCrLf & "Status must be Open to reassign!!"
            AssignClaim = False
            GoTo Rollback
        End If
    
        MyAdo.sqlString = "select LockUserID, Adj_Rationale from dbo.AUDITCLM_Hdr where CnlyClaimNum = '" & strCnlyClaimNum & "'"
        Set rsReassignCheck = MyAdo.OpenRecordSet
        strRationale = rsReassignCheck("Adj_Rationale") & ""
        strLockUserID = Trim(rsReassignCheck("LockUserID") & "")
        rsReassignCheck.Close
        Set rsReassignCheck = Nothing
            'If it's locked, don't reassign
            If Len(strLockUserID) > 0 Then
              strErrMsg = "Claim " & strCnlyClaimNum & " cannot be reassigned because it is locked for editing by " & strLockUserID & "."
              AssignClaim = False
              GoTo Rollback
            'Something was entered in Rationale - give user choice
            ElseIf Len(strRationale) > 0 Then
                'Show rationale and let user choose to continue, don't assign, or assign all
                strErrMsg = "Rationale exists for claim " & strCnlyClaimNum & "."
                AssignClaim = False
                GoTo Rollback
            End If
    
            ' update Queue header
            rsQueueHdr("LastUpdate") = Now()
            rsQueueHdr("UpdateUser") = Identity.UserName()
            rsQueueHdr("AssignedDate") = Date
            rsQueueHdr("AssignedFrom") = strAssignedFrom
            rsQueueHdr("AssignedTo") = strAssignedTo
            bResult = myCode_ADO.Update(rsQueueHdr, "usp_QUEUE_Hdr_Update")
            If bResult = False Then
                strErrMsg = "Error updating queue"
                AssignClaim = False
               GoTo Rollback
            End If
        End If
    
    myCode_ADO.CommitTrans
    strErrMsg = "Claim :: " & strCnlyClaimNum & " Assigned to :: " & strAssignedTo
    AssignClaim = True

Exit_Function:
    Set MyAdo = Nothing
    Set MyAdo = Nothing
    Exit Function

Rollback:
    'MsgBox strErrMsg, vbInformation
    AssignClaim = False
    MyAdo.RollbackTrans
    GoTo Exit_Function
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    AssignClaim = False
    MyAdo.RollbackTrans
    Resume Exit_Function
End Function




Private Sub lstState_Click()
    If Me.Frame28.Value = 1 Then
        'DRG
        RefreshDRGList
        RefreshAgeList
        Me.txtCount = GetCountOfSelection
    Else
    
        'AGE
        RefreshAgeList
        RefreshDRGList
        Me.txtCount = GetCountOfSelection
    End If
End Sub
