Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130515 KD: How does anything work around here? How do people get work done!?!?!
Dim previousCnlyClaimNum As String
Dim controlGotFocusCount As Integer
Private Const frmAppID As String = "AuditClm"
Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1
Dim rsNotes As ADODB.RecordSet
Dim NoteID As Long
Dim myCode_ADO As clsADO
Dim MyAdo As clsADO
Dim strUserFullName As String

Function getDocID()
    getDocID = left(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10) & Right(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10)
End Function

Private Sub cboOutcome_AfterUpdate()

Dim strOutcome As String
Dim sFilter As String
Dim sLOB As String
Dim sQADate As String

Set MyAdo = New clsADO
strOutcome = cboOutcome.Text
sFilter = ""
    
    If (strOutcome = "Approved: Clinical argument/evidence sufficient to support Discussion approval" _
    Or strOutcome = "Approved: New information/documentation received during discussion" _
    Or strOutcome = "Approved: Concept Updated/changed" _
    Or strOutcome = "Approved: Inpatient only procedure" _
    Or strOutcome = "Approved: Incorrect review criteria") _
    Then
             
             MsgBox ("Please enter Discussion Claim Note while Approving this Discussion!")
             
             Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
            
             frmGeneralNotes.frmAppID = frmAppID
             
             NoteID = Nz(DLookup("[NoteID]", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'"), -1)
             
             MyAdo.ConnectionString = GetConnectString("v_Data_Database")
             MyAdo.sqlString = "SELECT * from NOTE_Detail where NoteID = " & NoteID
             Set rsNotes = MyAdo.OpenRecordSet
             Set frmGeneralNotes.NoteRecordSource = rsNotes
             frmGeneralNotes.RefreshData
             frmGeneralNotes.NoteTypeID.Value = "DISCUSSION"
             ShowFormAndWait frmGeneralNotes
             Set frmGeneralNotes = Nothing
    
    Else
    strUserFullName = (DLookup("[UserName]", "ADMIN_User", "[UserID] ='" & Identity.UserName & "'"))
            
    'Based on my conversation With Marcia QAs will be reviewing only their types of claims, so there is no need to check!!!
        sLOB = ""
        sQADate = ""
        
        sLOB = Nz(DLookup("[LOB]", "Recon_QA_Contact", "[QAManager] ='" & strUserFullName & "'"), "")
        sQADate = Nz(DLookup("QADate", "QA_RECONS_Auditor", "cnlyclaimnum = '" & Me.CnlyClaimNum & "' and LOB = '" & sLOB & "'"), "")
        
        'This is QA manager and this claim had not been QAed yet!
        If sLOB <> "" And sQADate <> "" Then
        MsgBox ("Please enter Discussion Claim Note while Denying Discussion that's been Approved by the Auditor!")
             
             Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
            
             frmGeneralNotes.frmAppID = frmAppID
             
             NoteID = Nz(DLookup("[NoteID]", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'"), -1)
             
             MyAdo.ConnectionString = GetConnectString("v_Data_Database")
             MyAdo.sqlString = "SELECT * from NOTE_Detail where NoteID = " & NoteID
             Set rsNotes = MyAdo.OpenRecordSet
             Set frmGeneralNotes.NoteRecordSource = rsNotes
             frmGeneralNotes.RefreshData
             frmGeneralNotes.NoteTypeID.Value = "DISCUSSION"
             ShowFormAndWait frmGeneralNotes
             Set frmGeneralNotes = Nothing
        End If
        
    End If
                    
    
End Sub


Private Sub cboOutcome_Click()
    'MG 3/31/2014 only auto check attached box when user is on Post TD screen
    If Me.Parent.Controls("frmRECONSelection").Value = "4" Then
    
        If cboOutcome.Value = "Post TD" Or cboOutcome.Value = "Post RRL" Then
            Me.GenerateLetter.Value = True
        Else
            Me.GenerateLetter.Value = False
        End If
    End If
End Sub

Private Sub cboStandardDenialLetterType_AfterUpdate()

    Dim strUser As String
    Dim strUserFullName As String
    Dim resp As VbMsgBoxResult
    Dim therWording As String
        
    On Error GoTo Cleanup
    
    If Me.Parent.CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If UserRights = "user" Then
        Exit Sub
    End If
    
    therWording = ""
    
    strUser = Identity.UserName
    strUserFullName = DLookup("[UserName]", "ADMIN_User", "[UserID] ='" & strUser & "'")
    
    resp = MsgBox("You are about to load the standard denial letter into the rationale. This will overwrite all rationale data." & vbCrLf & "Would  you like to continue?", vbYesNo + vbQuestion, "Load Denial Letter")
            Select Case resp
                    Case vbYes
                        DoCmd.Hourglass True
                        
                        '12/11/2013 MG create language based on combo box selection
                        If left(cboStandardDenialLetterType.Value, 3) = "SEL" Then
                            Me.Rationale = ""
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "The" Then
                            Me.Parent.DocToRationale ("StandardDenial_THER")
                            therWording = ", on behalf of James Lee D.O., R.Ph, Chief Medical Officer, Cotiviti, LLC." & vbCrLf & "866-360-2507 (Press 4)"
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "MN " Then
                            Me.Parent.DocToRationale ("StandardDenial_MN")
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "IRF" Then
                            Me.Parent.DocToRationale ("StandardDenial_IRF")
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "CON" Then
            
                            If Right(cboStandardDenialLetterType.Value, 6) = "Denial" Then
                              Me.Parent.DocToRationale ("StandardDenial_CON")
                            Else
                              Me.Parent.DocToRationale ("StandardPartial_CON")
                            End If
                            
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "Add" Then
                            Me.Parent.DocToRationale ("StandardDenial")
                        ElseIf left(cboStandardDenialLetterType.Value, 3) = "2nd" Then
                            Me.Parent.DocToRationale ("StandardDenial_SEC")
                            
                        End If
                        
                        Me.Rationale = Me.Rationale & vbCrLf & "Submitted by:" & vbCrLf & Replace(strUserFullName, ".", " ") & therWording
            End Select
    
    DoCmd.Hourglass False
    Exit Sub
    
Cleanup:
    DoCmd.Hourglass False
    If Err.Number > 0 Then
            MsgBox Err.Number & " " & Err.Description
    End If
      
End Sub

Private Sub Form_Click()

    'MG 4/25/2014 TODO: add function here to clear exception when user click on the row of claim
    'If UserRights = "user" Then
    '    MsgBox "Mail Method: " & Me.CnlyClaimNum
    'Else
    '    MsgBox "No Response Method: " & Me.CnlyClaimNum
    'End If
    
End Sub

Private Sub Form_Load()
    
    Dim strRecSelect As String
    
    'MsgBox "Worktable form load"
    
    If UserRights = "user" Then
        strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_User & "' order by ReconAge DESC"
    Else
        strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_sysUser & "' order by ReconAge DESC"
    End If
    'gbl_frmLoad = 1
    Me.RecordSource = strRecSelect

End Sub

Private Sub Rationale_Click()

    Dim outcomeDisplay As String
    outcomeDisplay = Nz(Outcome.Value, "")
    
    'Disable rational if recon is approved
    If outcomeDisplay = "Approved" Or outcomeDisplay = "Post RRL" Then
        Rationale.Locked = True
        MsgBox "You do not need to write a rational for an approved recon. The system will automatically create a standard letter."
    Else
    
        Rationale.Locked = False
        'Prompt user to select denied outcome to write rational
        If outcomeDisplay = "" Then
            Rationale.Locked = True
            MsgBox "You need to select Denied or Partial Outcome before writing/editing a rational."
        End If
        
    End If
        
End Sub

Sub updateCtl(StrLock As String)
Dim ctl As Control

    For Each ctl In Me.Controls
        If (ctl.ControlType <> acLabel) And ctl.Name <> "cmdDenial" Then
        ctl.Locked = StrLock
    End If
    Next ctl
    If UserRights = "user" Then
            Me.Controls("Rationale").Locked = True
    End If
              
      Me.Icn.Locked = True
      
End Sub

Sub ControlGotFocus()

Dim strDocID As String
Dim StrSetUpdateUser As String
Dim StrClaimUnlockWrkTbl As String
Dim StrClaimUnlockRlts As String
Dim strSetDocID As String
Dim strLockUser As String
Dim strRationale As Variant

Dim LockCount As Integer


'strDocID = Format(Now(), "yyyymmddhhmmssms") & Rnd()

    strDocID = getDocID
    
    controlGotFocusCount = controlGotFocusCount + 1
    
    If gbl_frmLoad = 1 Or controlGotFocusCount <= 1 Then
        Exit Sub
    Else
        controlGotFocusCount = 2 'MG 9/24/2013 Because this function gets called 3 times during loading...1st time call not have the control available
    End If
    
    'MG If option value is not SAVED APPEAL OR POST TD THEN CONTINUE
    If Me.Parent.Controls("frmRECONSelection").Value = "1" Or Me.Parent.Controls("frmRECONSelection").Value = "2" Then
        Me.Parent.Controls("cmbCnlyClaimNum") = Nz(Me.CnlyClaimNum, "")
        Me.Parent.Controls("ClientNum") = 1
                
        gbl_CnlyClmNum = Nz(Me.CnlyClaimNum, "")
        
        If gbl_CnlyClmNum = "" Then
            Exit Sub
        End If
        
        'Set up query strings
        StrClaimUnlockWrkTbl = "DELETE LK.*, WT.Rationale" & _
                            " FROM QUEUE_RECON_Review_Locks AS LK" & _
                            " INNER JOIN QUEUE_RECON_Review_Result_WorkTable AS WT" & _
                            " ON LK.CnlyClaimNum = WT.CnlyClaimNum" & _
                            " Where (Len(WT.Rationale) = 0  OR ISNULL(WT.Rationale) = True)" & _
                            " AND LK.UpdateUser = '" & Identity.UserName & "'"
                            
        'VS 3/19/2015 Added DocID as PK to QUEUE_RECON_Review_Locks. Made DocID a part of PK in QUEUE_RECON_Review_Results. Joined on both CnlyClaimNum and DocID.
        StrClaimUnlockRlts = "DELETE LK.*, WT.Rationale" & _
                            " FROM QUEUE_RECON_Review_Locks AS LK" & _
                            " INNER JOIN QUEUE_RECON_Review_Results AS WT" & _
                            " ON LK.CnlyClaimNum = WT.CnlyClaimNum AND LK.DocID = WT.DocID" & _
                            " Where LK.UpdateUser = '" & Identity.UserName & "'"
         
        'insert the current claim into the lock table
        'VS 3/19/2015 insert DocID as well
        StrSetUpdateUser = "INSERT INTO  QUEUE_RECON_Review_Locks (CnlyClaimNum, DocID, UpdateUser) Select '" & gbl_CnlyClmNum & "','" & strDocID & "','" & Identity.UserName & "'"
        
        'update the DocID
        strSetDocID = "Update QUEUE_RECON_Review_Result_WorkTable set DocID = '" & strDocID & "' Where CnlyClaimNum ='" & gbl_CnlyClmNum & "'"
        
        'Get the Rationale from the worktable for the current claim
        'If Me.Parent.Controls("frmRECONSelection") = 1 Then
        '   strRationale = DLookup("[Rationale]", "QUEUE_RECON_Review_Result_WorkTable", "[CnlyClaimNum] ='" & gbl_CnlyClmNum & "'")
           
        'Get the Rationale from the Results table for the current claim
        'Else
        '   strRationale = DLookup("[Rationale]", "QUEUE_RECON_Review_Results", "[CnlyClaimNum] ='" & gbl_CnlyClmNum & "'")
         
        'End If
        
        If Identity.UserName = "" Then
            GoTo FormUpdate
        Else
            DoCmd.SetWarnings (False)
            DoCmd.RunSQL (StrClaimUnlockWrkTbl)
            DoCmd.RunSQL (StrClaimUnlockRlts)
            'Me.Parent.Controls("txtRationale") = strRationale
            'Unlock controls just incase they're locked
                        updateCtl False
                       ' Me.UpdateUser.Locked = True
                       ' Me.CnlyClaimNum.Locked = True
                    
            'Check if the claim is locked by another user
            LockCount = DCount("[CnlyClaimNum]", "QUEUE_RECON_Review_Locks", _
                     "[UpdateUser] <> '" & Identity.UserName & _
                     "' AND [CnlyClaimNum] = '" & gbl_CnlyClmNum & "'")
                
                If LockCount = 0 Then
                        DoCmd.RunSQL (StrSetUpdateUser)
                            If Len(Me.DocID) = 0 Or IsNull(Me.DocID) Then
                                DoCmd.SetWarnings (False)
                                DoCmd.RunSQL (strSetDocID)
                                DoCmd.SetWarnings (True)
                            End If
                Else
                        strLockUser = DLookup("[UpdateUser]", "QUEUE_RECON_Review_Locks", "[cnlyClaimNum] ='" & gbl_CnlyClmNum & "'")
                        
                        If Me.Parent.Controls("cmbCnlyClaimNum") <> previousCnlyClaimNum Then
                            MsgBox "Record is locked for editing by " & strLockUser, vbCritical, "Record locked"
                        End If
                        
                        updateCtl True
                End If
            DoCmd.SetWarnings (True)
        End If
    Else
       
        Me.Parent.Controls("cmbCnlyClaimNum") = Nz(Me.CnlyClaimNum, "")
        
        
        '12/20/2013 MG This is needed to show fax history for each tab
        Select Case Me.Parent.Controls("frmRECONSelection").Value
           Case 3
              Me.Parent.Controls("ClientNum") = 4 'this clientnum for faxing is appeal
           Case 4
              Me.Parent.Controls("ClientNum") = 5 'this clientnum for faxing is td post
           Case Else
              Me.Parent.Controls("ClientNum") = 1
        End Select

        gbl_CnlyClmNum = Me.CnlyClaimNum
        
        If gbl_CnlyClmNum = "" Then
            Exit Sub
        End If
     
    End If
    
    'MG 9/6/2013 one way to not display Editing is Lock by X person twice when loading the form
    previousCnlyClaimNum = Me.Parent.Controls("cmbCnlyClaimNum")
    
FormUpdate:
    'If Me.Parent.Controls("frmRECONSelection").Value <> "3" Then
    'End If
    'MsgBox "Testing focused on " & Me.ICN
    Me.Parent.Controls("txtFaxAttempts") = DLookup("FaxAttempts", "v_QUEUE_RECON_Fax_Attempts", "CnlyClaimNum ='" & CnlyClaimNum & "'")
    Me.Parent.Controls("DocID") = Me.DocID
    Me.Refresh
    
    '08/16/2013 MG refresh detail automatically
    Me.Parent.frm_QUEUE_RECON_Review_Claim_Detail.SourceObject = "frm_QUEUE_RECON_Review_Claim_Detail"
    
End Sub



Private Sub cmdDenial_Click()
Dim strUser As String
Dim strUserFullName As String
Dim resp As VbMsgBoxResult
    
    On Error GoTo Cleanup
    
    If Me.Parent.CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If UserRights = "user" Then
        Exit Sub
    End If
    
    
    
    strUser = Identity.UserName
    strUserFullName = DLookup("[UserName]", "ADMIN_User", "[UserID] ='" & strUser & "'")
    
     resp = MsgBox("You are about to load the standard letter into the rationale. This will overwrite all rationale data." & vbCrLf & "Would  you like to continue?", vbYesNo + vbQuestion, "Load Denial Letter")
            Select Case resp
                    Case vbYes
                        DoCmd.Hourglass True
                        Me.Parent.DocToRationale ("StandardDenial")
                        Me.Rationale = Me.Rationale & vbCrLf & "Submitted by:" & vbCrLf & strUserFullName
            End Select
    
    DoCmd.Hourglass False
    Exit Sub
    
Cleanup:
    DoCmd.Hourglass False
    If Err.Number > 0 Then
            MsgBox Err.Number & " " & Err.Description
    End If
      
End Sub

Private Sub Form_Current()
    
    If CurrentProject.AllForms("frm_QUEUE_RECON_main").IsLoaded = True Then
        ControlGotFocus
    End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If screen.ActiveControl.Name = "Rationale" Then
        Exit Sub
    End If
    
    Select Case KeyCode
    
     Case vbKeyUp
         KeyCode = 0
         On Error Resume Next
         DoCmd.GoToRecord acActiveDataObject, , acPrevious
        
     Case vbKeyDown
         KeyCode = 0
         On Error Resume Next
         DoCmd.GoToRecord acActiveDataObject, , acNext
        
     End Select

End Sub

'4/14/2015 VS Open Main Claims Screen by clicking ICN
Private Sub ICN_DblClick(Cancel As Integer)

     If Me.CnlyClaimNum & "" <> "" Then
        DisplayAuditClmMainScreen Me.CnlyClaimNum
    End If

End Sub

Private Sub frmGeneralNotes_NoteAdded()
    If SaveData_Notes Then
        MsgBox "Note added"
    End If
End Sub
'VS Request Auditor to enter notes if Discussion is being Approved!
Private Function SaveData_Notes() As Boolean
    Dim bResult As Boolean
    Dim updated As Boolean
    On Error GoTo ErrHandler
    
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    
    updated = False
    
    myCode_ADO.ConnectionString = GetConnectString("v_Code_Database")
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    
    If Not (rsNotes.BOF = True And rsNotes.EOF = True) Then
        'If the noteID is -1 then we need to create a new ID
        If NoteID = -1 Then
            'This is a public function that gets a unique ID based on the app being passed to the method
            NoteID = GetAppKey("NOTE")
            updated = True
        
        End If
        
        'Set the recordset of the header to contain the new note ID
        'Apply this new noteID to all of the records in the note recordset
        If Not (rsNotes.BOF = True And rsNotes.EOF = True) Then
            rsNotes.MoveFirst
            While Not rsNotes.EOF
                rsNotes.Update
                rsNotes("NoteID") = NoteID
                rsNotes.MoveNext
            Wend
            
            If updated Then
                MyAdo.sqlString = "UPDATE AUDITCLM_Hdr SET NoteID = " & NoteID & " WHERE " & _
                        " CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
                MyAdo.SQLTextType = sqltext
                MyAdo.Execute
            End If
                  
        End If
        
        'Pass the recordset back to SQL synching the results
        bResult = myCode_ADO.Update(rsNotes, "usp_NOTE_Detail_Apply")
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
