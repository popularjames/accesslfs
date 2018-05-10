Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DPR 4/23/2012 - Modified the Procedure to manually unlock a claim
'DPR 8/17/2012 - Moved the refresh of the collection of recordset objects out of the initial load.
'                These are only set when the GET method is called to reduce the overhead of the form's load


Option Compare Database

Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Public Event AuditClmError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private mrsAuditClmHdr As ADODB.RecordSet
Private mrsAuditClmDtl As ADODB.RecordSet
Private mrsAuditClmDiag As ADODB.RecordSet
Private mrsAuditClmProc As ADODB.RecordSet
Private mrsAuditClmProcRev As ADODB.RecordSet
Private mrsAuditClmDiagRev As ADODB.RecordSet
Private mrsAuditClmClaimsPlus As ADODB.RecordSet
Private mrsAuditClmHdrAdditionalInfo As ADODB.RecordSet
Private mrsNotes As ADODB.RecordSet

Private mAppID As String
Private mCnlyClaimNum As String
Private mbClaimExists As Boolean
Private mNoteID As Long
Private mbLockedForEdit As Boolean
Private mstrLockedUser As String
Private mdLockedDt As Date
Private mstrCurrentUser As String
Private mstrLoadClaimStatus As String
Private mstrClaimAuditor As String
Private mstrOverrideAuditor As String
Private mstrCreditOverrideReason As String

Public Property Let CnlyClaimNum(ByVal vData As String)
    mCnlyClaimNum = vData
End Property
Public Property Get CnlyClaimNum() As String
    CnlyClaimNum = mCnlyClaimNum
End Property
Public Property Get rsAuditClmHdr() As ADODB.RecordSet
    Set rsAuditClmHdr = mrsAuditClmHdr
End Property
Public Property Get rsAuditClmDtl() As ADODB.RecordSet
    '    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed
    If mrsAuditClmDtl Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
        Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
        Set rsAuditClmDtl = mrsAuditClmDtl
    Else
        Set rsAuditClmDtl = mrsAuditClmDtl
    End If
    
    
End Property

Public Property Get rsAuditClmDiag() As ADODB.RecordSet
    
'    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed
    
    If mrsAuditClmDiag Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_Diag WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'  ORDER BY LineNum"
        Set mrsAuditClmDiag = MyAdo.OpenRecordSet()
        Set rsAuditClmDiag = mrsAuditClmDiag
    Else
        Set rsAuditClmDiag = mrsAuditClmDiag
    End If
    
    
End Property

Public Property Get rsAuditClmProc() As ADODB.RecordSet
    
'    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed

    If mrsAuditClmProc Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_Proc WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'  ORDER BY LineNum"
        Set mrsAuditClmProc = MyAdo.OpenRecordSet()
        Set rsAuditClmProc = mrsAuditClmProc
    Else
        Set rsAuditClmProc = mrsAuditClmProc
    End If
    


End Property

Public Property Get rsAuditClmClaimsPlus() As ADODB.RecordSet
    
    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed

    If mrsAuditClmClaimsPlus Is Nothing Then
       MyAdo.sqlString = " SELECT * from AUDITCLM_ClaimsPlus WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
       Set mrsAuditClmClaimsPlus = MyAdo.OpenRecordSet()
       Set rsAuditClmClaimsPlus = mrsAuditClmClaimsPlus
    Else
        Set rsAuditClmClaimsPlus = mrsAuditClmClaimsPlus
    End If
    

End Property

Public Property Get rsAuditClmProcRev() As ADODB.RecordSet
    
    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed
    
    If mrsAuditClmProcRev Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_REVISED_Proc WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'  ORDER BY LineNum"
        Set mrsAuditClmProcRev = MyAdo.OpenRecordSet()
        Set rsAuditClmProcRev = mrsAuditClmProcRev
    Else
        Set rsAuditClmProcRev = mrsAuditClmProcRev
    End If
    

End Property

Public Property Get rsAuditClmDiagRev() As ADODB.RecordSet
    
    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed

    If mrsAuditClmDiagRev Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_REVISED_Diag WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "' ORDER BY LineNum"
        Set mrsAuditClmDiagRev = MyAdo.OpenRecordSet()
        Set rsAuditClmDiagRev = mrsAuditClmDiagRev
    Else
        Set rsAuditClmDiagRev = mrsAuditClmDiagRev
    End If
    

End Property

Public Property Get rsNotes() As ADODB.RecordSet
    
    '8/17/2012
    '************DPR Moved the refresh of these objects out of the initial load.  Only setting them when needed
    
    If mrsNotes Is Nothing Then
        MyAdo.sqlString = " SELECT * from NOTE_Detail where NoteID = " & mNoteID
        Set mrsNotes = MyAdo.OpenRecordSet
        Set rsNotes = mrsNotes
    Else
        Set rsNotes = mrsNotes
    End If
    


End Property

Public Property Get rsAuditClmHdrAdditionalInfo() As ADODB.RecordSet
    

'    5/13/2013
    '************TK: adding auditclm_hdr_additional recordset call
    
'    If mrsAuditClmHdrAdditionalInfo Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_hdr_AdditionalInfo WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
        Set mrsAuditClmHdrAdditionalInfo = MyAdo.OpenRecordSet()
        Set rsAuditClmHdrAdditionalInfo = mrsAuditClmHdrAdditionalInfo
'    Else
'        mrsAuditClmHdrAdditionalInfo.Requery
'       Set rsAuditClmHdrAdditionalInfo = mrsAuditClmHdrAdditionalInfo
'    End If
    
    
End Property

Public Property Get LockedForEdit() As Boolean
     LockedForEdit = mbLockedForEdit
End Property

Public Property Get LockedUser() As String
     LockedUser = mstrLockedUser
End Property

Public Property Get LockedDate() As Date
     LockedDate = mdLockedDt
End Property

Public Property Get ClaimExists() As Boolean
    ClaimExists = mbClaimExists
End Property

'This is a property of the class that will allow us to build processing logic around status code changes
Public Property Get LoadClaimStatus() As String
    LoadClaimStatus = mstrLoadClaimStatus
End Property


' TKL 3/2/2011: auditor credit override
Public Property Get ClaimAuditor() As String
    ClaimAuditor = mstrClaimAuditor     ' this value should only be set when loading claim
End Property

' TKL 3/2/2011: auditor credit override
Public Property Let OverrideAuditor(data As String)
    mstrOverrideAuditor = data
End Property

' TKL 3/2/2011: auditor credit override
Public Property Let CreditOverrideReason(data As String)
    mstrCreditOverrideReason = data
End Property
Public Function LoadClaim(strCnlyClaimNum As String, Optional bAllowChange As Boolean = True) As Boolean
'Damon 7/15/08
'Loaads a claim based on what is passed to the method
    On Error GoTo ErrHandler
        
    Dim strErrSource As String
    strErrSource = "clsAuditClm_LoadClaim"
    
   
    mbLockedForEdit = False
    mbClaimExists = False
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = strCnlyClaimNum
    
    MyAdo.sqlString = " SELECT * from AUDITCLM_Hdr WHERE cnlyClaimNum = '" & strCnlyClaimNum & "' and AccountID = " & gintAccountID
    'open the audit claims header and disconnect
    Set mrsAuditClmHdr = MyAdo.OpenRecordSet()
    
    If Not (mrsAuditClmHdr.BOF And mrsAuditClmHdr.EOF) Then
        'Assigning the current Claim Status at the time the claim was loaded
        'This is a property of the class that will allow us to build processing logic around status code changes
        mstrLoadClaimStatus = mrsAuditClmHdr("ClmStatus")
        
        
        ' TKL 3/2/2011: auditor credit override
        mstrClaimAuditor = mrsAuditClmHdr("Adj_Auditor") & ""
        mstrOverrideAuditor = ""
        
        'Check if the claim has an existing lock on it
        'if so, lock it down
        mstrLockedUser = mrsAuditClmHdr("LockUserID") & ""
        If Not IsNull(mrsAuditClmHdr("LockDt")) Then
            mdLockedDt = mrsAuditClmHdr("LockDt")
        End If
        
        If (mstrLockedUser = "" Or mstrLockedUser = mstrCurrentUser) And bAllowChange Then
            If LockClaim Then
                mbLockedForEdit = True
                mstrLockedUser = mstrCurrentUser
            End If
        End If
        
        
        'get the noteid associated with the claim
        mNoteID = Nz(mrsAuditClmHdr("NoteID"), -1)
        
        
        '8/17/2012
        '************DPR - Commenting out the initial load of these objects.  Will move the refreshes to the get event of the class and only retrieve them when needed
                '        'open the audit claims detail and disconnect
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'"
                '        Set mrsAuditClmDtl = MYADO.OpenRecordSet()
                '
                '        'Load other audit claim objects
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_Diag WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'  ORDER BY LineNum"
                '        Set mrsAuditClmDiag = MYADO.OpenRecordSet()
                '
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_Proc WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'  ORDER BY LineNum"
                '        Set mrsAuditClmProc = MYADO.OpenRecordSet()
                '
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_REVISED_Diag WHERE cnlyClaimNum = '" & strCnlyClaimNum & "' ORDER BY LineNum"
                '        Set mrsAuditClmDiagRev = MYADO.OpenRecordSet()
                '
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_REVISED_Proc WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'  ORDER BY LineNum"
                '        Set mrsAuditClmProcRev = MYADO.OpenRecordSet()
                '
                '        MYADO.SQLstring = " SELECT * from AUDITCLM_ClaimsPlus WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'"
                '        Set mrsAuditClmClaimsPlus = MYADO.OpenRecordSet()
                '
                '        MYADO.SQLstring = " SELECT * from NOTE_Detail where NoteID = " & mNoteID
                '        Set mrsNotes = MYADO.OpenRecordSet
        
        '8/17/2012
        '************DPR - Commenting out the initial load of these objects.  Will move the refreshes to the get event of the class and only retrieve them when needed
       
        LoadClaim = True
        mbClaimExists = True
    Else
        mbClaimExists = False
        LoadClaim = False
    End If
    
Exit_Function:
    Exit Function
    
ErrHandler:
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function

Public Function SaveClaim() As Boolean
    
'8/17/2012
'DPR change the save events to only save things that are loaded.  Do not call savs for objects that are not loaded
    
    Dim bResult As Boolean
    Dim strErrSource As String
    Dim iResult As Integer
    Dim strErrMsg As String

    On Error GoTo ErrHandler
    
    
    strErrSource = "clsAuditClm_SaveClaim"
    
    myCode_ADO.BeginTrans
    bResult = False

    ' check claim data to make sure it is OK to save
    If Me.ValidateClaim() = False Then
        bResult = False
        Err.Raise 65000, strErrSource, "Claim could not be validated.  Record not saved."
    Else
        bResult = True
    End If
    
    If bResult Then
        ' save the notes
        If Not mrsNotes Is Nothing Then
            bResult = SaveData_Notes
            If bResult = False Then
                Err.Raise 65000, strErrSource, "Error saving claim notes.  Record not saved"
            End If
        Else
            bResult = True
        End If
    End If
    
    If bResult Then
        ' save revised diagnosis codes
        If Not mrsAuditClmDiagRev Is Nothing Then
            bResult = SaveData_RevDiag
            If bResult = False Then
                Err.Raise 65000, strErrSource, "Error saving revised diag codes.  Record not saved"
            End If
        Else
            bResult = True
        End If
    End If
    
    If bResult Then
        ' save revised proc codes
        If Not mrsAuditClmProcRev Is Nothing Then
            bResult = SaveData_RevProc
            If bResult = False Then
                Err.Raise 65000, strErrSource, "Error saving revised Proc codes.  Record not saved"
            End If
        Else
            bResult = True
        End If
    End If
    
'DPR - This is not in use, do not load it
'    If bResult Then
'        ' save ClaimPlus data
'        bResult = SaveData_ClaimsPlus
'        If bResult = False Then
'            Err.Raise 65000, strErrSource, "Error saving ClaimPlus data.  Record not saved"
'        End If
'    End If
'
    If bResult Then
        If Not mrsAuditClmHdrAdditionalInfo Is Nothing Then
            ' save additional claim data
            bResult = SaveData_AdditionalHdrInfo
            If bResult = False Then
                Err.Raise 65000, strErrSource, "Error saving additional claim header data.  Record not saved"
            End If
        Else
            bResult = True
        End If
    End If
    
    If bResult Then
        ' save claim header
        mrsAuditClmHdr("LastUpDt") = Now
        mrsAuditClmHdr("LastUpUser") = mstrCurrentUser
        bResult = myCode_ADO.Update(mrsAuditClmHdr, "usp_AuditClm_Hdr_Update")
        
        If bResult = False Then
            Err.Raise 65000, strErrSource, "Error saving claim header data.  Record not saved"
        End If
        
        'We just saved the recordset, so let's move to the beginning so we can access its values
        If mrsAuditClmHdr.EOF Then
            mrsAuditClmHdr.MoveFirst
        End If
    End If
    
    
    If bResult Then
        ' save claim detail
        If Not mrsAuditClmDtl Is Nothing Then
            If mrsAuditClmDtl.recordCount > 0 Then 'for PRP reviewtype claims that dont have claim detail record JS 10/22/2012
                bResult = myCode_ADO.Update(mrsAuditClmDtl, "usp_AUDITCLM_Dtl_Update")
                If bResult = False Then
                    Err.Raise 65000, strErrSource, "Error saving claim detail data.  Record not saved"
                End If
            End If
        Else
            bResult = True
        End If
    End If
   
    
    ' TKL 3/2/2011: auditor credit override
    If bResult Then
        bResult = SaveData_CreditOverride
        If bResult = False Then
            Err.Raise 65000, strErrSource, "Error updating credit override.  Record not saved"
        End If
    End If
    
    '501 is the designated Claims plus movement
    'Call export procedure if this is the new status
    'If mrsAuditClmHdr.Fields("ClmStatus") = "501" Then
    '    If bResult Then
    '        bResult = SaveData_ExportClaimsPlus
    '        If bResult = False Then
    '            Err.Raise 65000, strErrSource, "Error Moving Claim to claims plus"
    '        End If
    '    End If
    'End If
        
    If bResult Then
        ' claim is saved.  Now move claim to next queue
        bResult = SaveData_ApplyQueue
        If bResult = False Then
            Err.Raise 65000, strErrSource, "Error updating queue.  Record not saved"
        End If
    End If
    
    myCode_ADO.CommitTrans
    SaveClaim = bResult
    
Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    myCode_ADO.RollbackTrans
    SaveClaim = bResult
    GoTo Exit_Sub
End Function


Private Function SaveData_ApplyQueue() As Boolean
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.RecordSet
    Dim rsQueue As New ADODB.RecordSet
    
    Dim strNextQueue As String
    Dim dtMaxThresholdDt As Date
    Dim iResult As Integer
    Dim strErrMsg As String
    Dim strSQL As String
    Dim strErrSource As String
    
    On Error GoTo ErrHandler
    
    strErrSource = "SaveData_ApplyQueue"
    
    'if there is no status update we don't need to update queue
    If Me.LoadClaimStatus = mrsAuditClmHdr("ClmStatus") Then
      SaveData_ApplyQueue = True
      GoTo Exit_Function
    End If
    
    
    'Determine what the status change is.  This will determine what the next logical queue is
    strSQL = "select * from AUDITCLM_Process_Logics where CurrStatus = '" & Me.LoadClaimStatus & _
            "' and DataType = '" & mrsAuditClmHdr("DataType") & "' and NextStatus = '" & _
            mrsAuditClmHdr("ClmStatus") & "' and ProcessModule = 'AuditClm' and ProcessType = 'Manual' " & _
            " and AccountID = " & gintAccountID
    
    MyAdo.sqlString = strSQL
    Set rs = MyAdo.OpenRecordSet
    
    ' get the current queue information.  This will determine the threshold.
    strSQL = "select * from QUEUE_Hdr where CnlyClaimNum = '" & mrsAuditClmHdr("CnlyClaimNum")
    Set rsQueue = MyAdo.OpenRecordSet
    
    'Check to make sure that there is a valid queue for the claim to move to.
    If rs.EOF = True And rs.BOF = True Then
        SaveData_ApplyQueue = False
        strErrMsg = "SaveData_ApplyQueue - Error updating Work Queue.  Status change not allowed by rules table."
        Err.Raise 65000, strErrSource, strErrMsg
    Else
        strNextQueue = rs("NextQueue")
    End If
    
    'Stored procedure to update the queue and move the claim along to the next status
    myCode_ADO.sqlString = "usp_QUEUE_Hdr_Fill_Update"
    myCode_ADO.SQLTextType = StoredProc
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_QUEUE_Hdr_Fill_Update"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pCnlyClaimNum") = mrsAuditClmHdr("CnlyClaimNum")
    cmd.Parameters("@pDataType") = mrsAuditClmHdr("DataType")
    cmd.Parameters("@pAuditNum") = mrsAuditClmHdr("Adj_AuditNum")
    cmd.Parameters("@pQueueType") = strNextQueue
    
    
    Select Case Nz(rs("GracePeriod"))
        Case Is > 0
            dtMaxThresholdDt = rs("GracePeriod") + Now
            cmd.Parameters("@pMaxThresholdDt") = dtMaxThresholdDt
        Case Is = -1
            dtMaxThresholdDt = rsQueue("MaxThresholdDt")
            cmd.Parameters("@pMaxThresholdDt") = dtMaxThresholdDt
    End Select
    
    cmd.Parameters("@pLastUpdate") = Now
    cmd.Parameters("@pUpdateUser") = Identity.UserName
    
    iResult = myCode_ADO.Execute(cmd.Parameters)
    strErrMsg = Nz(cmd.Parameters("@pErrMsg").Value, "")

    'Make sure there are no errors
    If strErrMsg <> "" Then
        SaveData_ApplyQueue = False
        strErrMsg = "Error updating Work Queue"
        Err.Raise 65000, strErrSource, strErrMsg
    Else
        SaveData_ApplyQueue = True
    End If
    
Exit_Function:
    Set rs = Nothing
    Set rsQueue = Nothing
    Set cmd = Nothing
    Exit Function

ErrHandler:
    'Return false so the whole save event can rollback.
    SaveData_ApplyQueue = False
    Resume Exit_Function

End Function


Private Function SaveData_RevProc() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String


    myCode_ADO.sqlString = "usp_AUDITCLM_REVISED_Proc_Forced_Delete"
    myCode_ADO.SQLTextType = StoredProc
        
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
        
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    LocCmd.Parameters("@pLineNum") = 0                      ' delete all rows
        
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
        
    bResult = True
    If iResult <> 0 Then
        Err.Raise 65000, "", "Error updating claim proc recordset"
        bResult = False
    Else
        If mrsAuditClmProcRev.recordCount > 0 Then
            bResult = myCode_ADO.Update(mrsAuditClmProcRev, "usp_AUDITCLM_REVISED_Proc_Insert")
            If bResult = False Then
                Err.Raise 65000, "", "Error updating claim proc recordset"
            End If
        End If
    End If
    
    SaveData_RevProc = bResult

Exit_Sub:
    Set LocCmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_RevProc = False
    GoTo Exit_Sub
End Function

Private Function SaveData_ClaimsPlus() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String

    If mrsAuditClmClaimsPlus.recordCount > 0 Then
        bResult = myCode_ADO.Update(mrsAuditClmClaimsPlus, "usp_AUDITCLM_CLaimsPlus_Apply")
        If bResult = False Then
            Err.Raise 65000, "", "Error updating claim plus recordset"
        End If
    Else
        bResult = True
    End If
    
    Set LocCmd = Nothing
    
    If bResult Then
        SaveData_ClaimsPlus = True
    Else
        SaveData_ClaimsPlus = False
    End If

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_ClaimsPlus = False
    GoTo Exit_Sub
End Function


Private Function SaveData_RevDiag() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
    
    
    ''Diag Revised
    myCode_ADO.sqlString = "usp_AUDITCLM_REVISED_Diag_Forced_Delete"
    myCode_ADO.SQLTextType = StoredProc
        
    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    LocCmd.Parameters("@pLineNum") = 0                      ' delete all rows
        
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    bResult = True
    If iResult <> 0 Then
            Err.Raise 65000, "", "Error updating claim diag recordset"
            bResult = False
    Else
        If mrsAuditClmDiagRev.recordCount > 0 Then
            bResult = myCode_ADO.Update(mrsAuditClmDiagRev, "usp_AUDITCLM_REVISED_Diag_Insert")
            If bResult = False Then
                Err.Raise 65000, "", "Error updating claim diag recordset"
            End If
        End If
    End If
    
        
    SaveData_RevDiag = bResult

Exit_Sub:
    Set LocCmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_RevDiag = False
    GoTo Exit_Sub
End Function


Private Function SaveData_Notes() As Boolean
    Dim bResult As Boolean
    
    On Error GoTo ErrHandler
    
    If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
        'If the noteID is -1 then we need to create a new ID
        If mNoteID = -1 Then
            'This is a public function that gets a unique ID based on the app being passed to the method
            mNoteID = GetAppKey("NOTE")
            'Set the recordset of the header to contain the new note ID
            Me.UpdateField "NoteID", mNoteID
            'Apply this new noteID to all of the records in the note recordset
            If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
                mrsNotes.MoveFirst
                While Not mrsNotes.EOF
                    mrsNotes.Update
                    mrsNotes("NoteID") = mNoteID
                    mrsNotes.MoveNext
                Wend
            End If
        End If
        'Pass the recordset back to SQL synching the results
        bResult = myCode_ADO.Update(mrsNotes, "usp_NOTE_Detail_Apply")
        
        If bResult = False Then
            Err.Raise 65000, "", "Error updating claim note"
        End If
   Else
        bResult = True
    End If
    
    SaveData_Notes = bResult

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_Notes = False
    GoTo Exit_Sub
End Function

Private Function SaveData_AdditionalHdrInfo() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String

    If Not (mrsAuditClmHdrAdditionalInfo Is Nothing) Then
        If mrsAuditClmHdrAdditionalInfo.recordCount > 0 Then
            Select Case gintAccountID
                Case 1, 2, 3
                    bResult = True
                Case 4
                    bResult = myCode_ADO.Update(mrsAuditClmHdrAdditionalInfo, "usp_AMERIGROUP_AUDITCLM_Hdr_AdditionalInfo_Apply")
                    If bResult = False Then
                        Err.Raise 65000, "", "Error updating additional claim header info"
                    End If
            End Select
        Else
            Select Case gintAccountID
                Case 1, 2, 3
                    bResult = True
                Case 4
                    Set LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
                    LocCmd.commandType = adCmdStoredProc
                    LocCmd.CommandText = "usp_AMERIGROUP_AUDITCLM_Hdr_AdditionalInfo_Delete"
                    LocCmd.Parameters.Refresh
                    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
                    myCode_ADO.SQLTextType = StoredProc
                    myCode_ADO.sqlString = "usp_AMERIGROUP_AUDITCLM_Hdr_AdditionalInfo_Delete"
            
                    iResult = myCode_ADO.Execute(LocCmd.Parameters)
                    iResult = LocCmd.Parameters("@RETURN_VALUE")
                    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
                    bResult = False
                    If iResult <> 0 Then
                        Err.Raise 65000, "", "Error updating additional claim header info"
                    Else
                        bResult = True
                    End If
            End Select
        End If
    Else
        bResult = True
    End If
    
    Set LocCmd = Nothing
    
    If bResult Then
        SaveData_AdditionalHdrInfo = True
    Else
        SaveData_AdditionalHdrInfo = False
    End If

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_AdditionalHdrInfo = False
    GoTo Exit_Sub
End Function


Public Function UpdateField(FieldName As String, FieldValue As Variant)
    Dim strErrSource As String
    
    On Error GoTo ErrHandler
    strErrSource = "clsAuditClm_UpdateField"
    mrsAuditClmHdr(FieldName).Value = FieldValue
    mrsAuditClmHdr.UpdateBatch adAffectAllChapters 'must do this to avoid recordset not syncing when displaying on forms
Exit Function

ErrHandler:
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
End Function

Private Sub Class_Initialize()
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    mAppID = "AUDITCLM"
    mstrCurrentUser = Identity.UserName
End Sub
Private Sub Class_Terminate()
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set mrsAuditClmHdr = Nothing
    Set mrsAuditClmDtl = Nothing
    Set mrsAuditClmDiag = Nothing
    Set mrsAuditClmProc = Nothing
    Set mrsAuditClmProcRev = Nothing
    Set mrsAuditClmDiagRev = Nothing
    Set mrsAuditClmClaimsPlus = Nothing
    Set mrsAuditClmHdrAdditionalInfo = Nothing
    Set mrsNotes = Nothing
End Sub


Public Function ValidateClaim() As Boolean
'Tues 2/5/2013 by KCF - Update so that doesn't error off if the ConceptID is at the header & detail level
    Dim strErrSource As String
    Dim bDetail As Boolean
    
    strErrSource = "clsAuditClm_ValidateClaim"
    
    If Nz(mrsAuditClmHdr.Fields("Adj_ConceptID"), "") = "" Then
    
        '8/17/2012 - Since we do not load the detail any longer unless it is needed, we need it for this check;
            
        If mrsAuditClmDtl Is Nothing Then
            MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
            Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
        End If
    
    
        If mrsAuditClmDtl.recordCount > 0 Then 'PRP reviewtype claims have no detail records
    
            ' No concept ID on header record.  Check if the detail has something specified
            mrsAuditClmDtl.MoveFirst
            bDetail = False
            
            'Go through the recordset and look for a non-blank idea code
            Do While Not mrsAuditClmDtl.EOF
                If Nz(mrsAuditClmDtl.Fields("Adj_ConceptID"), "") <> "" Or Nz(mrsAuditClmDtl.Fields("RecoveryReason"), "") <> "" Then
                    'Code found, set the check equal to true
                    bDetail = True
                    Exit Do
                End If
                mrsAuditClmDtl.MoveNext
            Loop
                
        End If
                
        If bDetail = False Then
            'If bdetail is false then raise an error telling the user why this failed
            RaiseEvent AuditClmError("No idea code is associated with this claim.", 65000, strErrSource)
        End If
        ValidateClaim = bDetail
    End If
    
    
    
    If Nz(mrsAuditClmHdr.Fields("Adj_ConceptID"), "") <> "" Then
        ' Header concept ID specified. Loop through detail concept to make sure they're they are the same
        ' as the header concept ID
        
        
        '8/17/2012 - Since we do not load the detail any longer unless it is needed, we need it for this check;
        If mrsAuditClmDtl Is Nothing Then
            MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
            Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
        End If
        
        bDetail = True
        If mrsAuditClmDtl.recordCount > 0 Then 'PRP reviewtype claims have no detail records JS  10/22/2012
            mrsAuditClmDtl.MoveFirst
            
            'Go through the recordset and compare the detail codes to the header codes
            Do While Not mrsAuditClmDtl.EOF
                If Nz(mrsAuditClmDtl.Fields("Adj_ConceptID"), "") <> "" Then
                    If mrsAuditClmHdr.Fields("Adj_ConceptID") <> mrsAuditClmDtl.Fields("Adj_ConceptID") Then 'BEGIN: Tues 2/5/2013 by KCF - Check that the Adj_ConceptID is not different on the header & detail rows
                        'Codes differ, set the check equal to false
                        bDetail = False
                        Exit Do
                    End If 'END: Tues 2/5/2013 by KCF - Check that the Adj_ConceptID is not different on the header & detail rows
                End If
                mrsAuditClmDtl.MoveNext
            Loop
            
            If bDetail = False Then
                'If bdetail is false then raise an error telling the user why this failed
                RaiseEvent AuditClmError("Header idea code with differing detail idea codes.", 65000, strErrSource)
            End If
        End If
        ValidateClaim = bDetail
    End If
    
    
    '************************
    '** BEGIN JS 04/22/2013 Check for Adj_Ind = 'Y' and detail adj_ConceptID should not be blank
    '************************
       
    'Since we do not load the detail any longer unless it is needed, we need it for this check;
    If mrsAuditClmDtl Is Nothing Then
        MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
        Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
    End If
    
    bDetail = True
    If mrsAuditClmDtl.recordCount > 0 Then 'PRP reviewtype claims have no detail records JS  10/22/2012
        mrsAuditClmDtl.MoveFirst
        
        'Go through the recordset and look for adj_ind = 'Y' with a missing adj_conceptid
        Do While Not mrsAuditClmDtl.EOF
            If Nz(mrsAuditClmDtl.Fields("Adj_Ind"), "") = "Y" Then
                'If there is a line with adj_ind set to Y and the conceptid field is empty then set the flag to false
                'VS 1/7/2016 Prompt the user for Recovery Reason when setting status to Recovery - Complex and Semi-Automated Claims only
                If (Nz(mrsAuditClmDtl.Fields("Adj_ConceptID"), "") = "" Or Nz(mrsAuditClmDtl.Fields("RecoveryReason"), "") = "") _
                And (mrsAuditClmHdr.Fields("ClmStatus") = "320" Or mrsAuditClmHdr.Fields("ClmStatus") = "320.2" Or mrsAuditClmHdr.Fields("ClmStatus") = "322") _
                And (mrsAuditClmHdr.Fields("Adj_ReviewType") = "C" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "CV" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "CVDRG" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "CVMU" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "PRP" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "S" _
                Or mrsAuditClmHdr.Fields("Adj_ReviewType") = "SV") Then
                    bDetail = False
                    Exit Do
                End If
            End If
            mrsAuditClmDtl.MoveNext
        Loop
        
        If bDetail = False Then
            'If bdetail is false then raise an error telling the user why this failed
            RaiseEvent AuditClmError("Detail line Idea Code (ConceptID) and/or Recovery Reason cannot be blank if indicator is set to 'Y'.", 65000, strErrSource)
        End If
    End If
    ValidateClaim = bDetail

    
    '************************
    '** END JS 04/22/2013 Check for Adj_Ind = 'Y' and detail adj_ConceptID should not be blank
    '************************
    
    
  
    '************************
    '** BEGIN   JS 05/07/2013 Check for Adj_Ind = 'Y' and lnReimbAmt should not be 0
    '**         JS 05/08/2013 Exclude HH claims from this check
    '**         JS 06/10/2013 Exclude HH claims from this check
    '************************

    If Nz(mrsAuditClmHdr.Fields("Datatype"), "") <> "HH" And Nz(mrsAuditClmHdr.Fields("Datatype"), "") <> "SNF" Then
        'Since we do not load the detail any longer unless it is needed, we need it for this check;
        If mrsAuditClmDtl Is Nothing Then
            MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
            Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
        End If
        
        bDetail = True
        If mrsAuditClmDtl.recordCount > 0 Then 'PRP reviewtype claims have no detail records JS  10/22/2012
            mrsAuditClmDtl.MoveFirst
            
            'Go through the recordset and look for adj_ind = 'Y' with a lnreimbamt = 0
            Do While Not mrsAuditClmDtl.EOF
                If Nz(mrsAuditClmDtl.Fields("Adj_Ind"), "") = "Y" Then
                    'If there is a line with adj_ind set to Y and lnreimbamt = 0 then set the flag to false
                    If Nz(mrsAuditClmDtl.Fields("LnReimbAmt"), 0) = 0 Then
                        bDetail = False
                        Exit Do
                    End If
                End If
                mrsAuditClmDtl.MoveNext
            Loop
            
            If bDetail = False Then
                'If bdetail is false then raise an error telling the user why this failed
                RaiseEvent AuditClmError("Claim Detail Edit line indicator cannot be 'Y' when LnReimbAmt = $0.00", 65000, strErrSource)
            End If
        End If
        ValidateClaim = bDetail
    End If
    
    '************************
    '** END JS 05/07/2013 Check for Adj_Ind = 'Y' and lnReimbAmt should not be 0
    '************************
    
    '****************IDEA CODE CHECK DETAIL*********************
    
'    '************************
'    '** BEGIN JS 07/05/2015 Do not allow recoveries where adjustments are less than $10 ($1 for overpayment)
'    '**       JS 07/14/2015 Remove per Gautam request, there might be recoveries that need to be less than $10
'    '************************
'
'    'check the clmstatus is a recovery
'    Dim ClmStatusRs As ADODB.RecordSet
'    With MyAdo
'        .SQLTextType = SQLtext
'        .sqlString = " SELECT * from xref_ClaimStatus WHERE clmstatus = '" & Nz(mrsAuditClmHdr.Fields("ClmStatus"), "") & "' and ClmStatusGroup = 'RC'" 'recovery status
'
'        Set ClmStatusRs = .ExecuteRS
'        If .GotData Then
'
''            'check projected savings is the difference between ReimbAmt and Adj_ReimbAmt
''            Dim ShouldBeProjectedSavings As Currency
''            ShouldBeProjectedSavings = Nz(mrsAuditClmHdr.Fields("ReimbAmt"), 0) - Nz(mrsAuditClmHdr.Fields("Adj_ReimbAmt"), 0)
''            If left(Nz(mrsAuditClmHdr.Fields("ClmStatus"), ""), 3) = "322" Then
''                if not ShouldBeProjectedSavings < 0 and abs(ShouldBeProjectedSavings) <> Nz(mrsAuditClmHdr.Fields("Adj_ReimbAmt"), 0)
''            Else
''
''            End If
'
'            'check projected savings meet thresold $
'            If left(Nz(mrsAuditClmHdr.Fields("ClmStatus"), ""), 3) = "322" Then
'                If Abs(Nz(mrsAuditClmHdr.Fields("Adj_ProjectedSavings"), 0)) < 1 Then
'                    RaiseEvent AuditClmError("Claim ProjSavings cannot be less than $1 for Underpayment", 65000, strErrSource)
'                    ValidateClaim = False
'                End If
'            Else
'                If Abs(Nz(mrsAuditClmHdr.Fields("Adj_ProjectedSavings"), 0)) < 10 Then
'                    RaiseEvent AuditClmError("Claim ProjSavings cannot be less than $10 for Overpayment", 65000, strErrSource)
'                    ValidateClaim = False
'                End If
'            End If
'
'        End If
'    End With
'    '************************
'    '** END JS 07/05/2015
'    '************************

Exit Function

ErrHandler:
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    ValidateClaim = False
End Function

Public Function LockClaim() As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsAuditClm_LockClaim"
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = Me.CnlyClaimNum
    
    ' lock claim for edit
    myCode_ADO.sqlString = "usp_AUDITCLM_Hdr_Lock"
    myCode_ADO.SQLTextType = StoredProc

    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    LocCmd.Parameters("@pLockUser") = Identity.UserName()
    LocCmd.Parameters("@pLockTime") = Now()
   

    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    If iResult <> 0 Then
        MsgBox strErrMsg
        LockClaim = False
        GoTo ErrHandler
    End If
    
    LockClaim = True
Exit Function

ErrHandler:
    LockClaim = False
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    
End Function
Public Function UnlockClaim(Optional MsgSuppress As Boolean = False) As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsAuditClm_UnLockClaim"
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = Me.CnlyClaimNum
    
    'lock claim for edit
    myCode_ADO.Connect
    myCode_ADO.sqlString = "usp_AUDITCLM_Hdr_UnLock"
    myCode_ADO.SQLTextType = StoredProc

    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    LocCmd.Parameters("@pLockUserID") = Identity.UserName()
    
    
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    Select Case iResult
        Case Is <> 0
            If MsgSuppress = False Then
                ' this is to suppress message when use exit claim screen
                MsgBox "Claim is locked by " & strErrMsg
            End If
            UnlockClaim = False
        Case Else
            UnlockClaim = True
    End Select
Exit Function

ErrHandler:
    UnlockClaim = False
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    
End Function

Public Function SyncConceptCodes(ClmType As CnlyClaimLevel)

    If ClmType = ClmDetail Then
    
        mrsAuditClmDtl.MoveFirst
        
        Do While Not mrsAuditClmDtl.EOF
            mrsAuditClmDtl.Fields("Adj_ConceptID").Value = ""
            mrsAuditClmDtl.Fields("Adj_VulnerabilityCd").Value = ""
            mrsAuditClmDtl.Fields("Adj_Ind").Value = ""
            mrsAuditClmDtl.Fields("Adj_ProjectedSavings").Value = Null  '*JAC 10/1/08 Added this
            mrsAuditClmDtl.Fields("RecoveryReason").Value = Null 'VS 12/8/2015 RVC Updates
            mrsAuditClmDtl.MoveNext
        Loop
        
        mrsAuditClmDtl.MoveFirst
        
    Else
            mrsAuditClmHdr.MoveFirst
            mrsAuditClmHdr.Fields("Adj_ConceptID").Value = ""
            mrsAuditClmHdr.Fields("Adj_VulnerabilityCd").Value = ""
            mrsAuditClmDtl.Fields("Adj_ProjectedSavings").Value = Null
        
    End If
    
        mrsAuditClmDtl.UpdateBatch adAffectAllChapters 'must do this to avoid recordset not syncing
        mrsAuditClmHdr.UpdateBatch adAffectAllChapters 'must do this to avoid recordset not syncing

End Function

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub

Private Function SaveData_ExportClaimsPlus() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String


    Set LocCmd = Nothing
    
    
    myCode_ADO.sqlString = "usp_Export_CLaimsPlusQueue_Insert"
    myCode_ADO.SQLTextType = StoredProc
    
    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    LocCmd.Parameters("@pAccountID") = mrsAuditClmHdr.Fields("AccountID")
    
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    If iResult <> 0 Then
          Err.Raise 65000, "", "Error exporting to claims plus queue " & strErrMsg
          bResult = False
    Else
          bResult = True
    End If
    
    SaveData_ExportClaimsPlus = bResult

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    RaiseEvent AuditClmError(Err.Description, Err.Number, "SaveData_ExportClaimsPlus")
    SaveData_ExportClaimsPlus = False
    GoTo Exit_Sub
End Function


'TKL 3/2/2011: auditor credit override
Private Function SaveData_CreditOverride() As Boolean
    Dim bResult As Boolean
    Dim LocCmd As ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String

    On Error GoTo ErrHandler
    
    If mstrOverrideAuditor <> "" Then
        myCode_ADO.sqlString = "usp_AUDITCLM_Auditor_Credit_Override_Insert"
        myCode_ADO.SQLTextType = StoredProc
            
        Set LocCmd = New ADODB.Command
        LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
        LocCmd.CommandText = myCode_ADO.sqlString
        LocCmd.commandType = adCmdStoredProc
        LocCmd.Parameters.Refresh
            
        LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
        LocCmd.Parameters("@pOldAuditor") = mstrClaimAuditor
        LocCmd.Parameters("@pNewAuditor") = mstrOverrideAuditor
        LocCmd.Parameters("@pComment") = mstrCreditOverrideReason
            
        LocCmd.Execute
        iResult = LocCmd.Parameters("@RETURN_VALUE")
        strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
            
        bResult = True
        If iResult <> 0 Then
            Err.Raise 65000, "", "Error inserting into AUDITCLM_Auditor_Credit_Override "
            bResult = False
        End If
    Else
        bResult = True
    End If
    
    SaveData_CreditOverride = bResult

Exit_Sub:
    Set LocCmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_CreditOverride = False
    GoTo Exit_Sub
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function RollbackStatus(Optional sRollbackUserID As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = TypeName(Me) & ".RollbackStatus"

    If sRollbackUserID = "" Then
        sRollbackUserID = Identity.UserName()
    End If

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_AUDITCLM_Status_Rollback_ClmAdmin"
        .Parameters.Refresh
        .Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
        .Parameters("@pRollbackNumber") = 1
        .Parameters("@pRollbackReason") = sRollbackUserID & " opted to manually rollback in Claim Admin"
        .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            RollbackStatus = False
            LogMessage strProcName, "ERROR", "Rollback status failed!", .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        End If
    End With

    RollbackStatus = True

Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function





Public Function UnLockClaimForce(Optional MsgSuppress As Boolean = False) As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsAuditClm_UnLockClaimForce"
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = Me.CnlyClaimNum
    
    'lock claim for edit
    myCode_ADO.Connect
    myCode_ADO.sqlString = "usp_AUDITCLM_Hdr_Force_UnLock"
    myCode_ADO.SQLTextType = StoredProc

    Set LocCmd = New ADODB.Command
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.CommandText = myCode_ADO.sqlString
    LocCmd.commandType = adCmdStoredProc
    LocCmd.Parameters.Refresh
            
    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    iResult = myCode_ADO.Execute(LocCmd.Parameters)
    iResult = LocCmd.Parameters("@RETURN_VALUE")
    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
    Select Case iResult
        Case Is <> 0
            If MsgSuppress = False Then
                ' this is to suppress message when use exit claim screen
                MsgBox "Claim is locked by " & strErrMsg
            End If
            UnLockClaimForce = False
        Case Else
            UnLockClaimForce = True
    End Select
Exit Function

ErrHandler:
    UnLockClaimForce = False
    RaiseEvent AuditClmError(Err.Description, Err.Number, strErrSource)
    
End Function