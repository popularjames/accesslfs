Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130515 KD: How does anything work around here? How do people get work done!?!?!

'MG 4/24/2013 change the below network path is needed
'Const strDenialLetter_MN = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\denial_MN.docx"
'Const strDenialLetter_IRF = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\L21402102387807LAA21402102387807LAAETTER_REPOSITORY\_TEMPLATES\RECON\denial_IRF.docx"
'Const strDenialLetter_CON = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\denial_CON.docx"
'Const strDenialLetter_SEC = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\denial_sec.docx"


'Const strApprovalLetter = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\approval.docx"
'Const strAppealLetter = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\appeal.docx"
'Const strPostRRLLetter = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON\post_RRL.docx"
Const rptReconPostRRL = "ReconPostRRL"
Const rptReconReviewResults = "ReconReviewResults"
Const rptAppeal = "rptAppeal"
Const rptPostTD = "PostTD"

Dim currentClientNum As Integer 'MG values are 1=recon and 4=recon with appeals
Dim rptName As String
Dim ContractId As Integer

Public Function CheckFormRecord()
    CheckFormRecord = Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSet.recordCount
End Function

Function ValidateLetter(FaxNumber As Variant, Recipient As Variant, Regading As Variant, Sender As Variant, Rationale As Variant, Outcome As Variant)

Dim ReturnVal As Integer


    If IsNull(FaxNumber) Then
        GoTo PromptUser
    End If
    
    If IsNull(Recipient) Then
        GoTo PromptUser
    End If
       
    If IsNull(Regading) Then
        GoTo PromptUser
    End If
        
    If IsNull(Rationale) Then
        GoTo PromptUser
    End If
    
    
    If IsNull(Sender) Then
        GoTo PromptUser
    End If
    
    If IsNull(Outcome) Then
        GoTo PromptUser
    End If
        
    ValidateLetter = 1
Exit Function

PromptUser:
    MsgBox "Fax Number, Recipient, Regading, From, Outcome or Rationale cannot be blank. Please review and try again.", vbCritical, "Missing Data"

End Function

Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function

Sub SaveRecords()
    
    '08/16/2013 MG this code will prevent recon exception from clearing when no claim detail are loaded. Users should not see the Expression error anymore
    If Len(Nz(cmbCnlyClaimNum.Value, "")) > 5 Then
    
        Dim MyCodeAdo As New clsADO
        Dim cmd As ADODB.Command
        Dim strRecSelect As String
    
        strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_sysUser & "' order by CnlyClaimNum"
            
        If Identity.UserName = "" Then
            Exit Sub
        End If
            
        MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
            
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = MyCodeAdo.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "usp_Queue_Recon_Review_Update" 'MG 6/26/2013 this is where exceptions are clear because this usp includes multiple claims save
        cmd.Parameters.Refresh
        cmd.Parameters("@VarUser") = Identity.UserName
        cmd.Parameters("@AssignedUser") = Me.frm_QUEUE_RECON_Review_Claim_Detail.Controls("AssignedTo").Value
        cmd.Execute
        
        Set MyCodeAdo = Nothing
        Set cmd = Nothing
        
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strRecSelect
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Refresh
    
    Else
        MsgBox "Claim detail information needs to be loaded prior to saving. Please send a screen shot to Michael Guan at michael.guan@connolly.com"
    End If
    
End Sub

Sub RecordsetFilter()

'MsgBox "test value = " & Me.frmRECONSelection.Value

Dim strSQL As String
Dim strVisible As Variant
    
    'By default show all claims
    If IsNull(Me.txtCnlyClaimNumLkUp) Then
        
        If Me.frmRECONSelection.Value = 1 Then 'MG New Recon
            strSQL = "select * from QUEUE_RECON_Review_Result_WorkTable where ClientNum='1' AND assignedTo Like '" & gbl_sysUser & "' order by reconAge DESC"
            Me.QUEUE_RECON_Review_Result_WorkTable.Controls("txtStartDt").Enabled = True
        End If
        
        If Me.frmRECONSelection.Value = 2 Then 'MG Saved Recon
            'MG applies for Saved Recon
            If UserRights = "user" Then 'customer service
                'MG 6/12/2013 Per R'Lay, CS should only be concern with failed FAX or document that needs to be faxed for USER access. By default, most people have auditor access and some have ADMIN aka most managers and DS team
                'Not sure why these are setup as pass thru SQL? I can't use view (freezes), so this is a workaround
                strSQL = " SELECT * FROM QUEUE_RECON_Review_Results WHERE ClientNum='1' AND docID IN (SELECT docID FROM QUEUE_RECON_READY_AND_NOT_FAX) ORDER BY reconAge DESC"
                
            Else
                'admin and auditors look at this
                strSQL = "select * from QUEUE_RECON_Review_Results where ClientNum='1' AND assignedTo Like '" & gbl_sysUser & "' order by reconAge DESC"
            End If
            Me.QUEUE_RECON_Review_Result_WorkTable.Controls("txtStartDt").Enabled = False
            
        End If
        
        If Me.frmRECONSelection.Value = 3 Then 'MG Saved Appeal
        
            If UserRights = "user" Or UserRights = "admin" Then
                'MG 6/12/2013 Per R'Lay, CS should only be concern with failed FAX or document that needs to be faxed for USER access. By default, most people have auditor access and some have ADMIN aka most managers and DS team
                'Not sure why these are setup as pass thru SQL? I can't use view (freezes), so this is a workaround
                strSQL = " SELECT * FROM QUEUE_RECON_Review_Results WHERE ClientNum='4' and GenerateLetter = false AND docID IN (SELECT docID FROM QUEUE_RECON_READY_AND_NOT_FAX) ORDER BY reconAge DESC"
            End If
                       
            'MG no need for auditors to review the recon/appeal letter since they are not reviewing it anyway
        End If
        
        If Me.frmRECONSelection.Value = 4 Then 'MG Post TD Letters
        
            If UserRights = "user" Or UserRights = "admin" Then
                'MG 11/1/2013 CS will need to fax Post TD Letters and don't show claims showing outcome='MR Review' as these should be review normally
                strSQL = " SELECT * FROM QUEUE_RECON_Review_Results WHERE ClientNum='5' and outcome NOT IN ('MR Review','None') and docID IN (SELECT docID FROM QUEUE_RECON_READY_AND_NOT_FAX) ORDER BY reconAge DESC"
            End If
            
            'MG no need for auditors to write rational letter since it contains standard language
        End If
        
        
    Else
        Dim ClientNum As Integer
        If Me.frmRECONSelection.Value = "1" Or Me.frmRECONSelection.Value = "2" Then
            ClientNum = 1 'recon
            'MsgBox "recon"
        ElseIf Me.frmRECONSelection.Value = "3" Then
            ClientNum = 4 'appeal
            'MsgBox "appeal"
        ElseIf Me.frmRECONSelection.Value = "4" Then
            ClientNum = 5 'Post TD
            'MsgBox "post td"
        End If
    
        'show search claim based on ICN or CnlyClaimNum
        If Me.frmRECONSelection.Value = 1 Then
            strSQL = "select * from QUEUE_RECON_Review_Result_WorkTable where (ICN Like '" & Me.txtCnlyClaimNumLkUp & "%' OR CnlyClaimNum like '" & Me.txtCnlyClaimNumLkUp & "%') AND assignedTo Like'" & gbl_sysUser & "' and ClientNum='1'"
        Else
            strSQL = "select * from QUEUE_RECON_Review_Results where (ICN Like '" & Me.txtCnlyClaimNumLkUp & "%' OR CnlyClaimNum like '" & Me.txtCnlyClaimNumLkUp & "%') AND assignedTo Like'" & gbl_sysUser & "' and ClientNum='" & ClientNum & "'"
        End If

    End If
    
    If Not Me.QUEUE_RECON_Review_Result_WorkTable.Form Is Nothing Then
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strSQL
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Refresh
    End If
    
End Sub

Private Sub cmdClear_Click()
    Me.txtCnlyClaimNumLkUp = Null
    RecordsetFilter
End Sub



Private Sub ClearExceptQueue(strClaimNum As String, strExceptType As String)

Dim MyCodeAdo As clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String
    
    Set MyCodeAdo = New clsADO
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_QUEUE_Exception_Delete"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = strClaimNum
    cmd.Parameters("@pExceptionType") = strExceptType
    cmd.Parameters("@pLastUpdate") = Now()
    cmd.Parameters("@pUpdateUser") = Identity.UserName
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    
    If spReturnVal <> 0 Then
        ErrMsg = cmd.Parameters("@pErrMsg")
        'MsgBox ErrMsg, vbCritical, "Error Clearing Queue"
    'Else
    '    MsgBox "Claim cleared from the exception queue.", vbInformation, "Queue Cleared"
    End If

    Set MyCodeAdo = Nothing
    Set cmd = Nothing

End Sub

Private Sub cmdDocToRationale_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If

    DocToRationale (screen.ActiveControl.Name)

End Sub

Private Sub cmdEHR_Click()

Dim strType As String
Dim strFaxNum As String
Dim strRecpt As String
Dim strICN As String

    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    strICN = Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("ICN").Value
    
    If MsgBox("The document with claim number '" & strICN & "' will be tagged as E.H.R. " & vbCrLf & vbCrLf & "Would you like to continue?", vbYesNo + vbQuestion, "E.H.R Claim") = vbNo Then
        Exit Sub
    End If
    
    strType = "EHR"
    
    strFaxNum = DLookup("[FaxNum]", "FAX_PopFaxNum", "[FaxType] ='" & strType & "'")
    strRecpt = DLookup("[Recipient]", "FAX_PopFaxNum", "[FaxType] ='" & strType & "'")
    
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("FaxNum") = strFaxNum
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("Recipient") = strRecpt
    Me.Refresh



End Sub

Private Sub cmdFailedFaxReport_Click()
    
    DoCmd.OpenForm "frm_QUEUE_RECON_Failed_Fax", acFormDS

End Sub

Private Sub cmdFaxStat_Click()

    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
        DoCmd.Close acForm, "frm_Fax_Selection"
    End If
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "RECON"
    'Forms!frm_Fax_Selection.Controls("cmdSendFax").visible = False

End Sub


Private Sub GenerateLetter(ICNToGen As String)
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    
    Dim db As Database
    
    Dim faxImage As ClsCnlyFaxImage
    'Dim rsado As clsADO
    
    Dim strFilePath As String
    Dim strNewFilePath As String
    Dim strICN As String
    Dim strLTTRID As String
    Dim strGerLTTR As String
    Dim strInsertLetter As String
    Dim strInsertQueue As String
    Dim strFileLoction As Variant
    Dim strID As String
    Dim strDeleteWktb As String
    Dim strDeleteQueue As String
    Dim strInsertHist As String
    Dim prtDefault As Printer
    'Dim myCodeADO As New clsADO
    
    Dim intCnt As Integer
    Dim strOutputPath As String
    
    Dim DocIDRs As DAO.RecordSet

    
    If Me.frmRECONSelection.Value = 1 Then
        MsgBox "You cannot generate letters from this view. Please switch to the Saved Discussions and try again", vbInformation, gbl_MsgBoxTitleLTTR
        Exit Sub
    End If
    
    Select Case ICNToGen
        Case "All"
    
            If MsgBox("You are about to attach all the checked letters to the claim and send to the fax queue. Would you like to continue?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbNo Then
            Exit Sub
            End If
            
        Case Else
              
            strICN = Me.QUEUE_RECON_Review_Result_WorkTable.Controls("ICN").Value
            
            If MsgBox("The document with claim number '" & strICN & "' will be attached and sent to the fax queue. Please ensure that the document is checked as attached." & vbCrLf & vbCrLf & "Would you like to continue?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbNo Then
            Exit Sub
            End If
    End Select
    
    
    Set db = CurrentDb
    strID = "PRODUAT"
    'strID = "DEVUAT"
    
    DoCmd.Hourglass True
    
    Dim ClientNum As Integer
    If Me.frmRECONSelection.Value = 1 Or Me.frmRECONSelection.Value = 2 Then
        ClientNum = 1
    ElseIf Me.frmRECONSelection.Value = 3 Then
        ClientNum = 4
    ElseIf Me.frmRECONSelection.Value = 4 Then
        ClientNum = 5
    End If
    
    Select Case ICNToGen
    
        Case "All"
            strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, Template, CnlyClaimNum, FaxNum, Recipient, FromName, Outcome, Regading, Rationale  from QUEUE_RECON_Review_Results" & _
                           " where GenerateLetter <> 0" & _
                           " AND ClientNum='" & ClientNum & "'" & _
                           " AND len(FaxNum)>9 AND len(recipient)>1 and len(regading)>1 and len(FromName)>1"
        Case Else
            strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, Template, CnlyClaimNum, FaxNum, Recipient, FromName, Outcome, Regading, Rationale  from QUEUE_RECON_Review_Results" & _
                           " where GenerateLetter <> 0" & _
                           " AND cnlyClaimNum = '" & ICNToGen & "'" & _
                           " AND ClientNum='" & ClientNum & "'" & _
                           " AND len(FaxNum)>9 AND len(recipient)>1 and len(regading)>1 and len(FromName)>1"
     End Select
                    
    Set DocIDRs = db.OpenRecordSet(strGerLTTR)
    'MG get total record count. Need to use moveLast or else it will not capture total record count
    DocIDRs.MoveLast
     
    Dim totalDocumentToFax As Integer
    totalDocumentToFax = DocIDRs.recordCount
    
    DocIDRs.MoveFirst
    

    'MsgBox "test: record count = " & totalDocumentToFax
        
        
    If (totalDocumentToFax < 1) Then
       MsgBox "You Have no letters to generate", vbInformation, gbl_MsgBoxTitleLTTR
       GoTo Cleanup
    End If
    
    'MG disabled below b/c it was really annoying for CS to not be able to fax documents if 1 document was missing certain element
    'While Not DocIDRs.EOF
    '    With DocIDRs
                
    '        If ValidateLetter(!FaxNum, !Recipient, !Regading, !FromName, !Rationale, !Outcome) <> 1 Then
    '        GoTo Cleanup
    '        End If
    '      .MoveNext
    '    End With
    'Wend
    
    'Set faxImage = New ClsCnlyFaxImage
    
    intCnt = 0
    
    strFileLoction = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & strID & "'")
    
    With DocIDRs
        .MoveLast
        .MoveFirst
    End With
    
    
    While Not DocIDRs.EOF
        With DocIDRs
           
            Set faxImage = New ClsCnlyFaxImage
            
            gbl_DocID = !DocID
            gbl_CnlyClmNum = !CnlyClaimNum
            strLTTRID = Format(Now(), "yyyymmddhhmmssms")
            strOutputPath = strFileLoction
            strNewFilePath = "FAX_" & gbl_DocID & "_" & strLTTRID
          
            faxImage.OutputPath = strOutputPath
            faxImage.ID = strNewFilePath
            
            Set Application.Printer = Application.Printers("Connolly Fax")
            Set prtDefault = Application.Printer
    
            If Me.frmRECONSelection.Value = 3 Then 'saved appealed
                rptName = getReportOrDocName(rptAppeal)
                DoCmd.OpenReport rptName, , , , acHidden
                
            ElseIf Me.frmRECONSelection.Value = 4 Then 'Post TD Letters
                If !Outcome = "Post TD" Then 'Some claims could be MR Review having client num 4 and we don't want to fax these out
                    rptName = getReportOrDocName(rptPostTD)
                    DoCmd.OpenReport rptName, , , , acHidden
                    
                ElseIf !Outcome = "Post RRL" Then
                
                    rptName = getReportOrDocName(rptReconPostRRL)
                    DoCmd.OpenReport rptName, , , , acHidden
                End If
            Else
                'Need to add logic to generate POST RRL when auditor select this
               ' If !Outcome = "Post RRL" Then
                    
                    'MsgBox Me.frm_QUEUE_RECON_Review_Claim_Detail.Form.Controls("Adj_ReviewType").Value
                    'MG 5/12/2014 POST RRL dynamic language content populate during open report will not work. So logic is added here to check adjReviewType before opening report
               '     getPostRRL Me.frm_QUEUE_RECON_Review_Claim_Detail.Form.Controls("Adj_ReviewType").Value, "print"
                    
                    'DoCmd.OpenReport "rpt_QUEUE_RECON_Post_RRL", , , , acHidden 'Post RRL
                    
              '  Else
                    rptName = getReportOrDocName(rptReconReviewResults)
                    DoCmd.OpenReport rptName, , , , acHidden 'recon response
                    'DoCmd.OpenReport "rpt_QUEUE_RECON_Review_Results", , , , acHidden 'recon response
                'End If
                
            End If
            
            faxImage.killClass = -1
                
            .Edit
            !GenerateLetter = 0
            !GenerateLetterDate = Now()
            .Update
            strInsertLetter = "INSERT INTO AUDITCLM_References ( CnlyClaimNum, CreateDt, RefType, RefSubType, RefLink, LastUpdateUser )" & _
                            " Select '" & !CnlyClaimNum & "','" & !GenerateLetterDate & "', ""ATTACH"", ""ProvCorres"" ,'" & strOutputPath & strNewFilePath & ".TIF' ,'" & Identity.UserName & "'"

            strDeleteWktb = "Delete * from FAX_Review_Worktable Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"
            
            strDeleteQueue = "Delete * from FAX_Work_Queue Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"
            
            strInsertQueue = "Insert into FAX_Review_Worktable(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,SenderPhoneNum,UpdateUser,DocImage, CnlyClaimNum) " & _
                            " Select DocID, " & currentClientNum & ", FaxNum, Recipient, Regading, FromName, PhoneNum, '" & Identity.UserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                            " From Queue_RECON_Review_Results" & _
                            " Where cnlyClaimNum =  '" & !CnlyClaimNum & "' AND ClientNum='" & ClientNum & "'"
            strInsertHist = "Insert into FAX_Review_Hist(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,SenderPhoneNum,UpdateUser,DocImage, CnlyClaimNum) " & _
                            " Select DocID, " & currentClientNum & ", FaxNum, Recipient, Regading, FromName, PhoneNum, '" & Identity.UserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                            " From Queue_RECON_Review_Results" & _
                            " Where cnlyClaimNum =  '" & !CnlyClaimNum & "' AND ClientNum='" & ClientNum & "'"
            
            db.Execute (strInsertLetter)
            db.Execute (strDeleteWktb)
            db.Execute (strDeleteQueue)
            db.Execute (strInsertQueue)
            db.Execute (strInsertHist)
            Call updateFaxTables(gbl_DocID, "EFAX", !CnlyClaimNum, 1, "", "")
            'Call ClearExceptQueue(!CnlyClaimNum, "EX014") 'MG 6/13/2012 will need to remove this since exception will be deleted when saved.
            'Call ClearExceptQueue(!CnlyClaimNum, "AOR") 'MG 6/13/2012 will need to remove this since exception will be deleted when saved.
            Set faxImage = Nothing
            .MoveNext
               intCnt = intCnt + 1
        End With
    Wend
    
    Me.Refresh
    
    DoCmd.Hourglass False
    
    'Adding sleep so the files have enough time to build.
    Sleep (3000)
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "RECON"

Cleanup:
    DoCmd.Hourglass False
    'Set faxImage = Nothing
    Set DocIDRs = Nothing
    'Set rsAdo = Nothing
    Set db = Nothing

End Sub

Private Sub cmdFaxSummary_Click()

    DoCmd.OpenForm "frm_QUEUE_RECON_Fax_Summary", acFormDS
    
End Sub

Private Sub cmdGenerateLetter_Click()

    GenerateLetter ("All") 'MG this is for all documents includeing recon and appeal

End Sub

Private Sub cmdGenerateLetterSingle_Click()

    'MG language are stored in the Report form itself. No need to get it from the word document because this method is the much faster
    'If Me.frmRECONSelection.Value = 3 Then
    '    DocToRationale ("StandardAppeal")
    'End If
    
    GenerateLetter (Me.cmbCnlyClaimNum)

End Sub

Private Sub cmdPreview_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If Me.frmRECONSelection.Value = 1 Then
        MsgBox "You cannot preview a letters from this view. Please switch to the Saved Discussions and try again", vbInformation, "Preview Fax Document"
        Exit Sub
    End If
    
    gbl_DocID = Me.DocID
    
    If Me.frmRECONSelection.Value = 3 Then 'saved appealed
                rptName = getReportOrDocName(rptAppeal)
                DoCmd.OpenReport rptName, acViewPreview
    
    ElseIf Me.frmRECONSelection.Value = 4 And Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").Value = "Post TD" Then
        rptName = getReportOrDocName(rptPostTD)
        DoCmd.OpenReport rptName, acViewPreview
    ElseIf Me.frmRECONSelection.Value = 4 And Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").Value = "Post RRL" Then
        rptName = getReportOrDocName(rptReconPostRRL)
        DoCmd.OpenReport rptName, acViewPreview
    Else
    
        'MG 3/31/2014 add preview for POST RRL
      '  If Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").Value = "Post RRL" Then
            'DoCmd.OpenReport "rpt_QUEUE_RECON_Post_RRL", acViewPreview
      '      getPostRRL Me.frm_QUEUE_RECON_Review_Claim_Detail.Form.Controls("Adj_ReviewType").Value, "preview"

      '  Else
            rptName = getReportOrDocName(rptReconReviewResults)
            DoCmd.OpenReport rptName, acViewPreview
        'End If
        
        
    End If

End Sub

Private Sub cmdReload_Click()
    'MG 9/30/2013 Disabled below code because auditor can have 0 record that the time, but more recon may have been assigned to him/her.
    'They should always be able to click on the refresh button.
    'If CheckFormRecord = 0 Then
    '    Exit Sub
    'End If

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
       
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    
    If UserRights = "user" Then
        'MsgBox "usp_QUEUE_Recon_Ready_And_Not_Fax ran"
        cmd.CommandText = "usp_QUEUE_Recon_Ready_And_Not_Fax" 'CS push refresh to see recons not faxed
    Else
        'MsgBox "usp_Queue_Recon_Review_Results_Worktable_Load ran"
        cmd.CommandText = "usp_Queue_Recon_Review_Results_Worktable_Load" 'Auditors and admins push refresh to see new recons added
        cmd.Parameters.Refresh
        cmd.Parameters("@pErrMsg") = ""
    End If
    
    cmd.Execute

    'MsgBox "test"
    
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Requery


On Error GoTo Cleanup

    
Cleanup:
    If Err.Number > 0 Then
            MsgBox Err.Number & " " & Err.Description
    End If
   
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing

End Sub

'MG 4/18/2013
Private Sub cmdSave_Click()

    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    '12/12/2013 MG Below is commented out. We should allow auditors to select approved and auto populate it with approval language in SAVED Recon tab.
    'Sometimes auditors may change a DENIED TO APPROVED
    'If Me.frmRECONSelection.Value = 2 Then
    '    Exit Sub
    'End If
    
    'Mike Guan 4/08/2013 Ensure user select recon outcome, so that system can auto update claim status
    Dim strOutcome As String
    Dim strAdllDoc As String
    
    strOutcome = Nz(Me.QUEUE_RECON_Review_Result_WorkTable.Controls("OutCome").Value, "")
    strAdllDoc = Nz(Me.QUEUE_RECON_Review_Result_WorkTable.Controls("ReceivedAddlDocFlag").Value, "")
    
    If strOutcome <> "" And strOutcome <> "N/A" Then

            If (strOutcome = "Partial: New information/documentation received during discussion" _
            Or strOutcome = "Partial: Clinical argument/evidence sufficient to support decision") _
        Then
            MsgBox "Since the Discussion is partial approved, please remember to adjust the claim information."
        End If
        
            If (strOutcome = "Approved: Clinical argument/evidence sufficient to support Discussion approval" _
            Or strOutcome = "Approved: New information/documentation received during discussion" _
            Or strOutcome = "Approved: Concept Updated/changed" _
            Or strOutcome = "Approved: Inpatient only procedure" _
            Or strOutcome = "Approved: Incorrect review criteria") _
            Then
                DocToRationale ("StandardApproval")
            End If
        
        If strOutcome = "Post RRL" Then
            DocToRationale ("StandardPostRRL") 'MG 4/25/2014 this is to display it to the auditor in the rational box. Really, the final version is coming from the report for CS to fax because it reference claims information.
        End If
        
        If (strOutcome = "Partial: New information/documentation received during discussion" _
            Or strOutcome = "Partial: Clinical argument/evidence sufficient to support decision" _
            Or strOutcome = "Denied") And Nz(Me.QUEUE_RECON_Review_Result_WorkTable.Controls("Rationale").Value, "") = "" Then
            MsgBox ("Your changes can not be saved unless you complete Rationale!")
            Exit Sub
        End If
            
        If (strOutcome = "Approved") And Nz(Me.QUEUE_RECON_Review_Result_WorkTable.Controls("Rationale").Value, "") = "" Then
            MsgBox ("Your changes can not be saved unless you complete Rationale!")
            Exit Sub
        End If
              
        'VS 08/04/2015 Reminder Message per Marcia's request.
        If (strAdllDoc = "Y") Then
             MsgBox ("Please remember to change Additional Documentation Flag to " & Chr(34) & "N" & Chr(34) & " if no Additional Documentation was received for this Discussion!")
        End If
         
        If MsgBox("You are about to save all your changes?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbYes Then
                SaveRecords
        End If
              
    Else
        MsgBox "You need to select APPROVED, DENIED or PARTIAL in Outcome field.", vbExclamation, "Before saving..."
    End If
    

End Sub


Private Sub cmdSearch_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    
    RecordsetFilter

End Sub


Private Sub cmdSeeClaimInQueue_Click()

Dim strRecSource As String

    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    'VS 3/19/2015 Added DocID to Queue_RECON_Review_Results tables. Joined on DocID in addition to CnlyClaimNum. Added 5 to client ext ref id list.
    strRecSource = "SELECT DV.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                            " INNER JOIN Queue_RECON_Review_Results AS DV" & _
                            " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum AND FWQ.DocID = DV.DocID" & _
                            " WHERE FWQ.Client_ext_Ref_ID IN ('1','4', '5')" & _
                            " AND FWQ.cnlyClaimNum = '" & Me.cmbCnlyClaimNum & "'" & _
                            " AND DV.AssignedTo Like '" & gbl_sysUser & "' Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
    
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "RECON"
    
    Forms!frm_Fax_Selection.RecordSource = strRecSource
    
    Forms!frm_Fax_Selection.Requery


End Sub

Private Sub Command64_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If


Dim strFilter As String
    
    DoCmd.OpenForm "frm_AUDITCLM_References_Grid_View"
    Forms!frm_AUDITCLM_References_Grid_View.Controls("cmdAttach").visible = False
    Forms!frm_AUDITCLM_References_Grid_View.Controls("btn_comment").visible = False
    Forms!frm_AUDITCLM_References_Grid_View.Controls("cmdFaxStatCS").visible = False
    Forms!frm_AUDITCLM_References_Grid_View.Controls("Line50").visible = False
    Forms!frm_AUDITCLM_References_Grid_View.Controls("cmdOpenFax").visible = False
    
    Forms!frm_AUDITCLM_References_Grid_View.RecordSource = "SELECT * FROM v_AUDITCLM_References WHERE cnlyClaimNum = '" & Me.cmbCnlyClaimNum & "'"
    Forms!frm_AUDITCLM_References_Grid_View.Caption = "Claim Number" & " - " & Me.cmbCnlyClaimNum
    
    Forms!frm_AUDITCLM_References_Grid_View.Requery
     
End Sub


Public Function FindSomething()
    Dim oForm As Form
    Dim oCtl As Control
    Dim iCnt As Integer
    
        Set oForm = Application.Forms("frm_QUEUE_RECON_Main")
        For Each oCtl In oForm.Controls
            
            If InStr(1, oCtl.Name, "fax") > 0 Then
                Debug.Print oCtl.Name
            End If
        Next
End Function


Public Sub DocToRationale(strCallingObject As String)
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If

    Dim oDoc As Word.Document
    Dim objApp As Word.Application
    Set objApp = New Word.Application
    Dim dlg As clsDialogs
    Set dlg = New clsDialogs
    Dim sFilePath As String
    Dim ErrMsg As String
    Dim strDenialLTTR As String
    Dim Path As String
        
    On Error GoTo ErrHandler

    'MG original logic
    'If strCallingObject = "cmdDocToRationale" Then
    '    With dlg
    '          sFilePath = .OpenPath("C:\", docf, , "Pick a word document to load!")
    '           If sFilePath = "" Then
    '            Exit Sub
    '        End If
    '    End With
    
    'Else
    '    sFilePath = strDenialLTTR
    'End If
        
    If strCallingObject = "cmdDocToRationale" Then
    
        With dlg
              sFilePath = .OpenPath("C:\", docf, , "Pick a word document to load!")
               If sFilePath = "" Then
                Exit Sub
            End If
        End With
        
    'VS 3/6/2015 Let's make this table driven!
    Else
        sFilePath = getReportOrDocName(strCallingObject)
        
    End If
    
    'Open an exisiting document
    Open sFilePath For Binary Access Read Write Lock Read Write As #1
    Close #1
    
    Set oDoc = objApp.Documents.Open(sFilePath, , False)
    oDoc.Activate
    
    'Multi-Line Textbox or RichTextBox control
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("Rationale").Value = Replace(oDoc.Content, vbCr, vbNewLine)
    
    Me.Refresh
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("Rationale").SetFocus
    
     
Cleanup:
    objApp.Application.Quit SaveChanges:=False
    Set objApp = Nothing
    Set oDoc = Nothing
Exit Sub


ErrHandler:
    If Err.Number = 70 Then
            ErrMsg = "File Locked for editing by another user. Please close the file and try again."
            Close #1
    Else
            ErrMsg = Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source
    End If
    
            MsgBox "Error: " & ErrMsg, vbCritical, "Error Loading File"
     Me.Refresh
     
    GoTo Cleanup

End Sub




Private Sub Form_Load()

'mg 3/31/2014 this will populate outcome option dynamically depending on the screen user select
populateOutcomeOption ("")

'Created: Curlan Johnson
'Date: 5/8/12
'This will set a few business rules and permissions for use of the RECON Screen
'FindSomething

Dim db As Database
Dim strInsertSql As String
Dim strRecSelect As String
Dim strUserChk As String

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim LoadRs As ADODB.RecordSet
Dim CheckUserRS As ADODB.RecordSet

Dim file1 As String
Dim file2 As String

'Test code. Please don't build it!
Dim fso As New FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

file1 = "\\ccaintranet.com\dfs-cms-fld\Imaging\Client\Work\CMS\Faxing\Pending Receipts\ALJ\1c1ddee601cfb0c200000168.eml"
file2 = "\\ccaintranet.com\dfs-cms-fld\Imaging\Client\Work\CMS\Faxing\Pending Receipts\ALJ\1c1ddee601cfb0c200000168_Renamed.eml"

If fso.FileExists(file1) Then
    fso.MoveFile file1, file2
End If
    
    'MG By default client number = 1
    currentClientNum = 1
 
    If UserRights = "user" Then 'MG CS access
        Me.frmRECONSelection.Value = 2 'MG By default, check the SAVED RECON tab for them
        Me.QUEUE_RECON_Review_Result_WorkTable.Controls("txtStartDt").Enabled = False
        'lblViewReconReadyToFax.visible = True
        'cboViewReconReadyToFax.visible = True
        cmdFailedFaxReport.visible = True
        cmdFaxSummary.visible = True
        
        cmdGenerateLetter.visible = True
        cmdGenerateLetterSingle.visible = True
        optSavedAppeal.visible = True
        optSavedRecon.visible = True
        optPostTD.visible = True
        optNewRecon.visible = False

        Call checkFilters
        
    ElseIf UserRights = "admin" Then 'MG Admin are data service and managers
    
        Me.frmRECONSelection.Value = 1 'MG check new RECON tab
        optSavedAppeal.visible = True
        optSavedRecon.visible = True
        optNewRecon.visible = True
        optPostTD.visible = True
        cmdFailedFaxReport.visible = True
        cmdFaxSummary.visible = True
        
        'lblViewReconReadyToFax.visible = True
        'cboViewReconReadyToFax.visible = True
        
        
    Else 'auditor access

        Me.frmRECONSelection.Value = 1 'MG check new RECON tab
        
        optNewRecon.visible = True
        optSavedRecon.visible = True
        
        cmdGenerateLetter.visible = False
        cmdGenerateLetterSingle.visible = False
        optSavedAppeal.visible = False
        'lblViewReconReadyToFax.visible = False
        'cboViewReconReadyToFax.visible = False
        cmdFailedFaxReport.visible = False
        cmdFaxSummary.visible = False
        optPostTD.visible = False
        
    End If
    
    
        


    
    If gbl_frmLoad = 1 Then
        GoTo Cleanup
    Else
        gbl_frmLoad = 1
    End If
    
On Error GoTo Cleanup

    'Set db = CurrentDb

                   
    'If UserRights = "user" Then
    '    MsgBox "load test"
    '    strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_User & "' order by ICN"
    'Else
    '    strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_sysUser & "' AND cnlyClaimNum NOT in (select cnlyClaimNum from v_QUEUE_RECON_Appeal) order by ICN"
    'End If

    'myCodeADO.ConnectionString = GetConnectString("v_CODE_Database")

    'Set cmd = New ADODB.Command
    'cmd.ActiveConnection = myCodeADO.CurrentConnection
    'cmd.commandType = adCmdStoredProc
    'cmd.CommandText = "usp_Queue_Recon_Review_Results_Worktable_Load"
    'cmd.Execute
    
    'Me.QUEUE_RECON_Review_Result_WorkTable.SourceObject = "frm_QUEUE_RECON_Review_Result_WorkTable"
    'Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strRecSelect
    'Me.Fax_Status_Queue__Reconsideration_.SourceObject = "frm_Fax_Status_History"
    
    'Me.frm_QUEUE_RECON_Review_Claim_Detail.SourceObject = "frm_QUEUE_RECON_Review_Claim_Detail"
    
    'Me.Refresh
    'gbl_frmLoad = 0
    'Me.txtCnlyClaimNumLkUp.SetFocus
    
'    RecordsetFilter

Cleanup:
    'MsgBox "clean up"
    
    If Err.Number > 0 Then
            MsgBox Err.Number & " " & Err.Description
    End If
   
    If gbl_frmLoad = 1 Then gbl_frmLoad = 0
    Set MyCodeAdo = Nothing
    Set LoadRs = Nothing
    Set CheckUserRS = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
'Created: Curlan Johnson
'Date: 5/8/12

Dim db As Database
Dim UnloadRs As ADODB.RecordSet
Dim strDeleteSql As String
Dim strUnloadSql As String
Dim strLogout As String
Dim intCount As Integer
Dim msgResp As String
Dim StrClearRationale As String
Dim rsAdo As clsADO
On Error GoTo Cleanup
    
    If Identity.UserName = "" Then
        Exit Sub
    End If

    strDeleteSql = "Delete * from QUEUE_RECON_Review_Locks Where UpdateUser =  '" & Identity.UserName & "'"
                        

    'VS 3/19/2015 Added DocID to QUEUE_RECON_Review_Locks tables. Joined on DocID in addition to CnlyClaimNum.
    StrClearRationale = " UPDATE QUEUE_RECON_Review_Result_WorkTable AS WT" & _
                        " INNER JOIN QUEUE_RECON_Review_Locks AS LK" & _
                        " ON WT.CnlyClaimNum = LK.CnlyClaimNum AND LK.DocID = WT.DocID" & _
                        " Set WT.Rationale = ''" & _
                        " Where Len(WT.Rationale) > 0" & _
                        " AND LK.UpdateUser = '" & Identity.UserName & "'"
    
    'VS 3/19/2015 Added DocID to QUEUE_RECON_Review_Locks tables. Joined on DocID in addition to CnlyClaimNum.
    strUnloadSql = "Select LK.*, WT.Rationale" & _
                        " FROM QUEUE_RECON_Review_Locks AS LK" & _
                        " INNER JOIN QUEUE_RECON_Review_Result_WorkTable AS WT" & _
                        " ON LK.CnlyClaimNum = WT.CnlyClaimNum AND LK.DocID = WT.DocID" & _
                        " Where Len(WT.Rationale) > 0" & _
                        " AND LK.UpdateUser = '" & Identity.UserName & "'"
                    
    Set db = CurrentDb
    Set rsAdo = New clsADO
    rsAdo.ConnectionString = GetConnectString("v_DATA_Database")
    rsAdo.SQLTextType = sqltext
    rsAdo.sqlString = strUnloadSql
    
    Set UnloadRs = rsAdo.OpenRecordSet
    If Not (UnloadRs.EOF And UnloadRs.BOF) Then
        msgResp = MsgBox("You Have not saved all your changes. Would you like to save your changes before exiting?", vbYesNoCancel + vbQuestion, gbl_MsgBoxTitleLTTR)
        Select Case msgResp
            Case vbYes
                SaveRecords
            Case vbNo
                DoCmd.SetWarnings (False)
                DoCmd.RunSQL (StrClearRationale)
                DoCmd.SetWarnings (True)
            Case Else
                Cancel = True
                GoTo Cleanup
         End Select
    End If
    
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL (strDeleteSql)
    DoCmd.SetWarnings (True)
    
Cleanup:
    If Err.Number > 0 Then
        MsgBox Err.Number & " " & Err.Description
    End If

    Set UnloadRs = Nothing
    Set rsAdo = Nothing
    Set db = Nothing

End Sub



Private Sub frmRECONSelection_AfterUpdate()
    'MsgBox "Recon selection radio option change"
    
    'capture the client number based on radio button changes
    If Me.frmRECONSelection.Value = 1 Or Me.frmRECONSelection.Value = 2 Then
        currentClientNum = 1
        populateOutcomeOption ("")
    ElseIf Me.frmRECONSelection.Value = 3 Then
        currentClientNum = 4
        populateOutcomeOption ("")
    ElseIf Me.frmRECONSelection.Value = 4 Then
        currentClientNum = 5
        populateOutcomeOption ("POSTTD")
    End If
        
    'RecordsetFilter
    Call checkFilters
    
End Sub




Private Sub optAttachSend_AfterUpdate()

    Call checkFilters
        
End Sub

'MG this sub procedure can be used in multiple cases
'MG this sub procedure can be used in multiple cases
Private Sub checkFilters()

    Dim strSQL As String
    
    'OPTION selection on upper left side
    If Me.frmRECONSelection.Value = 1 Then 'new recon
        
    ElseIf Me.frmRECONSelection.Value = 2 Then 'saved recon
        strSQL = "select * from QUEUE_RECON_Review_Results where assignedTo Like '" & gbl_sysUser & "' AND GenerateLetter = True and clientNum='1' order by reconAge DESC"
        
    ElseIf Me.frmRECONSelection.Value = 3 Then 'saved appeal
        strSQL = "select * from QUEUE_RECON_Review_Results where GenerateLetter = True and clientNum='4' order by reconAge DESC"
        
    ElseIf Me.frmRECONSelection.Value = 4 Then 'post TD
        strSQL = "select * from QUEUE_RECON_Review_Results where GenerateLetter = True and clientNum='5' order by reconAge DESC"
        
    Else
        MsgBox "No attach option selected."
    End If
    
    'MG If option is showing ALL for user access, only display documents that needs faxing. For auditors, it should display all recon assigned to them
    '1 = All
    '2 = Checked
    If Me.optAttachSend.Value = 1 Then
        RecordsetFilter
    Else
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strSQL
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Refresh
    End If

End Sub

Private Sub txtCnlyClaimNumLkUp_AfterUpdate()

    RecordsetFilter

End Sub

'mg 3/31/2014 This function drives the options available on each screen
Public Function populateOutcomeOption(screen As String)
    
    'MsgBox "PopulateOutcomeOption Function Called"
    
    'clear all items in combo box
    Dim i As Integer
    For i = 1 To Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").ListCount
        'Remove an item from the ListBox.
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").RemoveItem 0
    Next i
       
    'display POST TD Letter generation option or update status to Late MR
    If screen = "POSTTD" Then
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "", "0" 'MG 4/14 show blank result so CS can tell quickly if this claim has been reviewed or not
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Post TD", "1"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "MR Review", "2"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Post RRL", "3"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "None", "4"
    Else
        'Always display the below options if user is not in POST TD Screen
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Approved: Clinical argument/evidence sufficient to support Discussion approval", "0"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Denied", "1"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Partial: New information/documentation received during discussion", "2"
        'Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Post RRL", "3"
        
        'VS Per Brian's Request expanding on Approval Reasons
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Approved: New information/documentation received during discussion", "3"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Approved: Concept Updated/changed", "4"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Approved: Inpatient only procedure", "5"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Approved: Incorrect review criteria", "6"
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("cboOutcome").AddItem "Partial: Clinical argument/evidence sufficient to support decision", "7"
        
    End If
    
End Function



'mg 5/12/2014 this function will detemine what type of post rrl to use
Public Function getPostRRL(adjReviewType As String, openType As String)
    
    'Have just Postpay claims right now.
     rptName = getReportOrDocName(rptReconPostRRL)
    
     If openType = "preview" Then
        DoCmd.OpenReport rptName, acViewPreview
     End If
     
     If openType = "print" Then
        DoCmd.OpenReport rptName, , , , acHidden 'Post RRL
     End If
    
    
'    If openType = "preview" Then
'        If adjReviewType = "PRP" Then
'            DoCmd.OpenReport "rpt_QUEUE_RECON_Post_RRL_PrePay", acViewPreview
'        Else
'            DoCmd.OpenReport rpt_QUEUE_RECON_Post_RRL_PostPay, acViewPreview
'        End If
'    End If
'
'    If openType = "print" Then
'        If adjReviewType = "PRP" Then
'            DoCmd.OpenReport "rpt_QUEUE_RECON_Post_RRL_PrePay", , , , acHidden 'PrePay RRL
'        Else
'            DoCmd.OpenReport "rpt_QUEUE_RECON_Post_RRL_PostPay", , , , acHidden 'Post RRL
'        End If
'    End If
    
End Function

Public Function getReportOrDocName(DocType As String) As String
    
    ContractId = DLookup("ContractID", "AUDITCLM_Hdr", "CnlyClaimNum = '" & Nz(cmbCnlyClaimNum.Value, "") & "'")
    getReportOrDocName = DLookup("DocName", "RECON_Templates_Per_Contract", "ContractID = " & ContractId & " and DocType = '" & DocType & "'")

End Function
