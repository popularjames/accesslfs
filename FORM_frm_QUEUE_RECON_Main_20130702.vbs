Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130515 KD: How does anything work around here? How do people get work done!?!?!

'MG 4/24/2013 change the below network path is needed
Const strDenialLetter = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON_DENIAL\denial.docx"
Const strApprovalLetter = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_TEMPLATES\RECON_DENIAL\approval.docx"


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


Private Sub cmbCnlyClaimNum_DblClick(Cancel As Integer)

'ViewImage

End Sub

Sub SaveRecords()

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
    cmd.CommandText = "usp_Queue_Recon_Review_Update"
    cmd.Parameters.Refresh
    cmd.Parameters("@VarUser") = Identity.UserName
    cmd.Parameters("@AssignedUser") = Me.frm_QUEUE_RECON_Review_Claim_Detail.Controls("AssignedTo").Value
    cmd.Execute
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strRecSelect
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Refresh
    
End Sub

Sub RecordsetFilter()

Dim strSQL As String
Dim strVisible As Variant
        
    If UserRights = "user" Then
        Me.frmRECONSelection.Value = 2 'MG 2=Saved Recon. I'm adding in bunch of comments based on Curlan's code
    End If
    
    'By default show all claims
    If IsNull(Me.txtCnlyClaimNumLkUp) Then
        
        If Me.frmRECONSelection.Value = 1 Then 'MG 1=New Recon
            strSQL = "select * from QUEUE_RECON_Review_Result_WorkTable where assignedTo Like '" & gbl_sysUser & "' order by ICN"
        Else 'MG 1=Saved recon
            'strSQL = "select * from QUEUE_RECON_Review_Results where assignedTo Like '" & gbl_sysUser & "' order by ICN"
            
            'MG applies for Saved Recon
            If UserRights = "user" Then
                'MG 6/12/2013 Per R'Lay, CS should only be concern with failed FAX or document that needs to be faxed for USER access. By default, most people have auditor access and some have ADMIN aka most managers and DS team
                'Not sure why these are setup as pass thru SQL? I can't use view (freezes), so this is a workaround
                strSQL = " SELECT * FROM QUEUE_RECON_Review_Results WHERE docID IN (SELECT docID FROM QUEUE_RECON_READY_AND_NOT_FAX) ORDER BY CreateDt"
            Else
                strSQL = "select * from QUEUE_RECON_Review_Results where assignedTo Like '" & gbl_sysUser & "' order by ICN"
            End If
            
        End If
        
    Else
        'show search claim
        Select Case Me.OptSearch.Value
        
            Case 1 'mg search by icn
                If Me.frmRECONSelection.Value = 1 Then
                    strSQL = "select * from QUEUE_RECON_Review_Result_WorkTable where ICN Like '" & Me.txtCnlyClaimNumLkUp & "%' AND assignedTo Like'" & gbl_sysUser & "'"
                Else
                    strSQL = "select * from QUEUE_RECON_Review_Results where ICN Like '" & Me.txtCnlyClaimNumLkUp & "%' AND assignedTo Like'" & gbl_sysUser & "'"
                End If
           Case 2 'mg search by cnlyclaimnum
                If Me.frmRECONSelection.Value = 1 Then
                    strSQL = "select * from QUEUE_RECON_Review_Result_WorkTable where cnlyClaimNum = '" & Me.txtCnlyClaimNumLkUp & "' AND assignedTo Like'" & gbl_sysUser & "'"
                Else
                    strSQL = "select * from QUEUE_RECON_Review_Results where cnlyClaimNum = '" & Me.txtCnlyClaimNumLkUp & "' AND assignedTo Like '" & gbl_sysUser & "'"
                End If
        End Select
    End If
    
    If Not Me.QUEUE_RECON_Review_Result_WorkTable.Form Is Nothing Then
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strSQL
        Me.QUEUE_RECON_Review_Result_WorkTable.Form.Refresh
    End If
    
End Sub

'MG For testing pupose for adhoc fax request. Probably don't need it anymore since Viktoria has a process setup.
Private Sub cmdAdhocFax_Click()
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
        DoCmd.Close acForm, "frm_Fax_Selection"
    End If
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "Adhoc"
    'Forms!frm_Fax_Selection.Controls("cmdSendFax").visible = False

End Sub

Private Sub cmdClear_Click()

'If CheckFormRecord = 0 Then
'    Exit Sub
'End If

    
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

'Private Sub test()
'Dim oAdo As clsADO
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("Whatever")
'        .SqlString = "usp_BlahBLah"
'        .SQLTextType = StoredProc
'        .Parameters.Refresh
'        .Parameters("@pCnlyClaimNum") = WHatever
'        .Execute
'        If .GotData = True Then
'
'        End If
'    End With
'End Sub


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
        MsgBox "You cannot generate letters from this view. Please switch to the Saved RECON's and try again", vbInformation, gbl_MsgBoxTitleLTTR
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
    
    Select Case ICNToGen
        Case "All"
                    strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, Template, CnlyClaimNum, FaxNum, Recipient, FromName, Outcome, Regading, Rationale  from QUEUE_RECON_Review_Results" & _
                                   " where GenerateLetter <> 0" & _
                                   " AND AssignedTo Like '" & gbl_sysUser & "'"
        Case Else
                    strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, Template, CnlyClaimNum, FaxNum, Recipient, FromName, Outcome, Regading, Rationale  from QUEUE_RECON_Review_Results" & _
                                   " where GenerateLetter <> 0" & _
                                   " AND cnlyClaimNum = '" & ICNToGen & "'" & _
                                   " AND AssignedTo Like '" & gbl_sysUser & "'"
     End Select
                    
    Set DocIDRs = db.OpenRecordSet(strGerLTTR)
    If (DocIDRs.EOF And DocIDRs.BOF) Then
       MsgBox "You Have no letters to generate", vbInformation, gbl_MsgBoxTitleLTTR
       GoTo Cleanup
    End If
    
    While Not DocIDRs.EOF
        With DocIDRs
            If ValidateLetter(!FaxNum, !Recipient, !Regading, !FromName, !Rationale, !Outcome) <> 1 Then
            GoTo Cleanup
            End If
          .MoveNext
        End With
    Wend
    
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
            strLTTRID = Format(Now(), "yyyymmddhhmmssms")
            strOutputPath = strFileLoction
            strNewFilePath = "FAX_" & gbl_DocID & "_" & strLTTRID
          
            faxImage.OutputPath = strOutputPath
            faxImage.ID = strNewFilePath
            
                    
            Set Application.Printer = Application.Printers("Connolly Fax")
            Set prtDefault = Application.Printer
    
            DoCmd.OpenReport "rpt_QUEUE_RECON_Review_Results", , , , acHidden
            
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
                            " Select DocID, 1, FaxNum, Recipient, Regading, FromName, PhoneNum, '" & Identity.UserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                            " From Queue_RECON_Review_Results" & _
                            " Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"
            strInsertHist = "Insert into FAX_Review_Hist(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,SenderPhoneNum,UpdateUser,DocImage, CnlyClaimNum) " & _
                            " Select DocID, 1, FaxNum, Recipient, Regading, FromName, PhoneNum, '" & Identity.UserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                            " From Queue_RECON_Review_Results" & _
                            " Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"
            
            db.Execute (strInsertLetter)
            db.Execute (strDeleteWktb)
            db.Execute (strDeleteQueue)
            db.Execute (strInsertQueue)
            db.Execute (strInsertHist)
            Call updateFaxTables(gbl_DocID, "EFAX", !CnlyClaimNum, 1, "", "")
            Call ClearExceptQueue(!CnlyClaimNum, "EX014") 'MG 6/13/2012 will need to remove this since exception will be deleted when saved.
            Call ClearExceptQueue(!CnlyClaimNum, "AOR") 'MG 6/13/2012 will need to remove this since exception will be deleted when saved.
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

Private Sub cmdGenerateLetter_Click()

    GenerateLetter ("All")

End Sub

Private Sub cmdGenerateLetterSingle_Click()

    GenerateLetter (Me.cmbCnlyClaimNum)

End Sub

Private Sub cmdPreview_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    If Me.frmRECONSelection.Value = 1 Then
        MsgBox "You cannot preview a letters from this view. Please switch to the Saved RECON's and try again", vbInformation, "Preview Fax Document"
        Exit Sub
    End If
    
    gbl_DocID = Me.DocID
    
    DoCmd.OpenReport "rpt_QUEUE_RECON_Review_Results", acViewPreview
    

End Sub

Private Sub cmdReload_Click()
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If



Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command

    
    On Error GoTo Cleanup
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
       
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_Queue_Recon_Review_Results_Worktable_Load"
    cmd.Execute

    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Requery



    
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
    
    If Me.frmRECONSelection.Value = 2 Then
        Exit Sub
    End If
    
    'Mike Guan 4/08/2013 Ensure user select recon outcome, so that system can auto update claim status
    Dim strOutcome As String
    
    strOutcome = Nz(Me.QUEUE_RECON_Review_Result_WorkTable.Controls("OutCome").Value, "")
    
    If strOutcome <> "" And strOutcome <> "N/A" Then

        If strOutcome = "Partial" Then
                MsgBox "Since the recon is partial approved, please remember to adjust the claim information."
        End If


        If MsgBox("You are about to save all your changes?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbYes Then
            
            If strOutcome = "Approved" Then
                DocToRationale ("StandardApproval")
            End If
            
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
    
    strRecSource = "SELECT DV.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                            " INNER JOIN Queue_RECON_Review_Results AS DV" & _
                            " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum" & _
                            " WHERE FWQ.Client_ext_Ref_ID = ""1""" & _
                            " AND FWQ.cnlyClaimNum = '" & Me.cmbCnlyClaimNum & "'" & _
                            " AND DV.AssignedTo Like '" & gbl_sysUser & "' Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
    
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "RECON"
    
    Forms!frm_Fax_Selection.RecordSource = strRecSource
    
    Forms!frm_Fax_Selection.Requery


End Sub

Private Sub Command107_Click()

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

Private Sub cmdFax_Click()


Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
        
    If Identity.UserName = "" Then
        Exit Sub
    End If
        
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_QUEUE_RECON_Load_FAX_Review_Worktable"
    cmd.Parameters.Refresh
    cmd.Parameters("@varUpdateUser") = Identity.UserName
    cmd.Parameters("@varSysUser") = gbl_sysUser
    cmd.Execute

    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal

End Sub

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
    
    ElseIf strCallingObject = "StandardDenial" Then
    
        sFilePath = strDenialLetter
        
    ElseIf strCallingObject = "StandardApproval" Then
    
        sFilePath = strApprovalLetter
        
    End If



    
    ''Open an exisiting document
    
    Open sFilePath For Binary Access Read Write Lock Read Write As #1
    Close #1
    
    Set oDoc = objApp.Documents.Open(sFilePath, , False)
    oDoc.Activate
    ''Multi-Line Textbox or RichTextBox control
    Me.QUEUE_RECON_Review_Result_WorkTable.Form.Controls("Rationale").Value = Replace(oDoc.Content, vbCr, vbNewLine + vbNewLine)
    
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

'Created: Curlan Johnson
'Date: 5/8/12
'This will set a few business rules and permissions for use of the RECON Screen

Dim db As Database
Dim strInsertSql As String
Dim strRecSelect As String
Dim strUserChk As String

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim LoadRs As ADODB.RecordSet
Dim CheckUserRS As ADODB.RecordSet

On Error GoTo Cleanup
    
    If gbl_frmLoad = 1 Then
        GoTo Cleanup
    Else
        gbl_frmLoad = 1
    End If
    
    Set db = CurrentDb
    
    
    Me.frmRECONSelection.Value = 1
    
                   
    If UserRights = "user" Then
        strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_User & "' order by ICN"
    Else
        strRecSelect = "select * from QUEUE_RECON_Review_Result_Worktable where AssignedTo Like '" & gbl_sysUser & "' order by ICN"
    End If

    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_Queue_Recon_Review_Results_Worktable_Load"
    cmd.Execute
    
    Me.QUEUE_RECON_Review_Result_WorkTable.SourceObject = "frm_QUEUE_RECON_Review_Result_WorkTable"
    'On Error Resume Next
    'Me.QUEUE_RECON_Review_Result_WorkTable.Form.RecordSource = strRecSelect
    'On Error GoTo Cleanup
    Me.Fax_Status_Queue__Reconsideration_.SourceObject = "frm_Fax_Status_History"
    
    Me.frm_QUEUE_RECON_Review_Claim_Detail.SourceObject = "frm_QUEUE_RECON_Review_Claim_Detail"
    
    Me.Refresh
    gbl_frmLoad = 0
    Me.txtCnlyClaimNumLkUp.SetFocus
    
'    RecordsetFilter

Cleanup:
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
                        

                        
    StrClearRationale = " UPDATE QUEUE_RECON_Review_Result_WorkTable AS WT" & _
                        " INNER JOIN QUEUE_RECON_Review_Locks AS LK" & _
                        " ON WT.CnlyClaimNum = LK.CnlyClaimNum" & _
                        " Set WT.Rationale = ''" & _
                        " Where Len(WT.Rationale) > 0" & _
                        " AND LK.UpdateUser = '" & Identity.UserName & "'"
    
    strUnloadSql = "Select LK.*, WT.Rationale" & _
                        " FROM QUEUE_RECON_Review_Locks AS LK" & _
                        " INNER JOIN QUEUE_RECON_Review_Result_WorkTable AS WT" & _
                        " ON LK.CnlyClaimNum = WT.CnlyClaimNum" & _
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
    
    RecordsetFilter

End Sub


Private Sub optAttachSend_AfterUpdate()

Dim strSQL As String
    
    strSQL = "select * from QUEUE_RECON_Review_Results where assignedTo Like '" & gbl_sysUser & "' AND GenerateLetter = True order by CnlyClaimNum"
    
    'MG If this is showing New Recon, then exit sub
    If Me.frmRECONSelection.Value = 1 Then
        Exit Sub
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




Private Sub optAttachSend_BeforeUpdate(Cancel As Integer)

    'MG 6/12/2013 refresh recon ready and not faxed yet
    
    Dim db As Database
    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    On Error GoTo Cleanup
       
        Set db = CurrentDb
    
        MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = MyCodeAdo.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "usp_QUEUE_Recon_Ready_And_Not_Fax"
        cmd.Execute

        'MsgBox "recon not faxed refresh test"
        
Cleanup:
        If Err.Number > 0 Then
            MsgBox Err.Number & " " & Err.Description
        End If
       
        Set MyCodeAdo = Nothing
        Set cmd = Nothing
        Set db = Nothing
        

End Sub

Private Sub txtCnlyClaimNumLkUp_AfterUpdate()

    RecordsetFilter

End Sub
