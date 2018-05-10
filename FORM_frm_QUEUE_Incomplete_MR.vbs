Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database

Public Function CheckFormRecord()

CheckFormRecord = Me.frm_QUEUE_MR_Request_Sub.Form.RecordSet.recordCount

End Function


Function ValidateLetter(FaxNumber As Variant, Recipient As Variant)

Dim ReturnVal As Integer


    If IsNull(FaxNumber) Then
      GoTo PromptUser
    End If
    
    If IsNull(Recipient) Then
      GoTo PromptUser
    End If


ValidateLetter = 1
Exit Function

PromptUser:
    MsgBox "Fax Number or Recipient cannot be blank. Please review and try again.", vbCritical, "Missing Data"

End Function

Sub RecordsetFilter()
Dim strSQL As String
 
    If Me.txtCnlyClaimNumLkUp <> "" Then
    
        Select Case Me.OptSearch.Value
            'gbl_sysUser_Inc is % for users who send faxes
            Case 1 'search by ICN
            strSQL = "select * FROM QUEUE_MR_Request_Fax F WHERE ICN = '" & Me.txtCnlyClaimNumLkUp & "'"
       
            Case 2 'search by Claim Id
            strSQL = "select * FROM QUEUE_MR_Request_Fax WHERE CnlyClaimNum = '" & Me.txtCnlyClaimNumLkUp & "'"
       
        End Select
        Call Form_frm_QUEUE_MR_Request_Sub.SetRecordSource(strSQL)
        
    Else
       Call Form_frm_QUEUE_MR_Request_Sub.SetRecordSource
          
    End If

End Sub

Private Sub cmdClear_Click()

Call Form_frm_QUEUE_MR_Request_Sub.SetRecordSource
Me.txtCnlyClaimNumLkUp = ""

End Sub


Private Sub cmdFaxStat_Click()

If CheckFormRecord = 0 Then
    Exit Sub
End If

If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
    DoCmd.Close acForm, "frm_Fax_Selection"
End If

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "INC_MR"

End Sub


Private Sub GenerateLetter(ICNToGen As String)

If CheckFormRecord = 0 Then
    Exit Sub
End If


Dim db As Database

Dim faxImage As ClsCnlyFaxImage

Dim strFilePath As String
Dim strNewFilePath As String
Dim strICN As String
Dim strLTTRID As String
Dim strGerLTTR As String
Dim strInsertLetter As String
Dim strInsertQueue As String
Dim strFileLoction As Variant
Dim strDeleteWktb As String
Dim strDeleteQueue As String
Dim strInsertHist As String
Dim strInsertRef As String
'Dim testme As String
Dim deleted As Boolean
Dim strSetDocID As String


Dim intCnt As Integer

Dim DocIDRs As DAO.RecordSet

Select Case ICNToGen
    Case "All"

        If MsgBox("You are about to attach all the checked letters to the claim and send to the fax queue. Would you like to continue?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbNo Then
        Exit Sub
        End If
        
    Case Else
          
        strICN = Me.frm_QUEUE_MR_Request_Sub.Controls("txtICN").Value
        
        If MsgBox("The document with claim number '" & strICN & "' will be attached and sent to the fax queue. Please ensure that the document is checked as attached." & vbCrLf & vbCrLf & "Would you like to continue?", vbYesNo + vbQuestion, gbl_MsgBoxTitleLTTR) = vbNo Then
        Exit Sub
        End If
End Select


Set db = CurrentDb

DoCmd.Hourglass True

'1: letter had not been generated
'0: it's been generated
Select Case ICNToGen
    Case "All"
                strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, CnlyClaimNum, ICN, FaxNum, Recipient from QUEUE_MR_Request_Fax" & _
                               " where GenerateLetter <> 0"
                                                     
    Case Else
                strGerLTTR = "select DocId, GenerateLetter, GenerateLetterDate, CnlyClaimNum, ICN, FaxNum, Recipient from QUEUE_MR_Request_Fax" & _
                               " where GenerateLetter <> 0" & _
                               " AND cnlyClaimNum = '" & ICNToGen & "'"


                            
 End Select
                
Set DocIDRs = db.OpenRecordSet(strGerLTTR)
If (DocIDRs.EOF And DocIDRs.BOF) Then
   MsgBox "You Have no letters to generate", vbInformation, gbl_MsgBoxTitleLTTR
   GoTo Cleanup
End If

While Not DocIDRs.EOF
    With DocIDRs
        If ValidateLetter(!FaxNum, !Recipient) <> 1 Then
        GoTo Cleanup
        End If
      .MoveNext
    End With
Wend

'Set faxImage = New ClsCnlyFaxImage

intCnt = 0

strFileLoction = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & strMRID & "'")

With DocIDRs
    .MoveLast
    .MoveFirst
End With

While Not DocIDRs.EOF
    With DocIDRs
       
        Set faxImage = New ClsCnlyFaxImage
        
        gbl_DocID = generateDocID()
        
        strLTTRID = Format(Now(), "yyyymmddhhmmssms")
        strOutputPath = strFileLoction
        strNewFilePath = "FAX_" & gbl_DocID & "_" & strLTTRID
      
        faxImage.OutputPath = strOutputPath
        faxImage.ID = strNewFilePath
        
        Call Form_frm_QUEUE_MR_Request_Sub.SetMRRequestedField(!CnlyClaimNum, !Icn)
        
        'testme = strOutputPath & strNewFilePath & ".pdf"
        
        Form_frm_QUEUE_MR_Request_Sub.id_set = True
          
        .Edit
        !GenerateLetter = 0
        !GenerateLetterDate = Now()
        .Update
                
                
                strSetDocID = "Update QUEUE_MR_Request_Fax set DocID = '" & gbl_DocID & "' Where CnlyClaimNum ='" & !CnlyClaimNum & "'"
                strInsertLetter = "INSERT INTO INC_MR_Insert_Ref_Info values ('" & !CnlyClaimNum & "', '" & gbl_DocID & "', '" & !Icn & "', '" & strLTTRID & "')"
                strDeleteWktb = "Delete * from FAX_Review_Worktable Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"

                strDeleteQueue = "Delete * from FAX_Work_Queue Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"

                strInsertQueue = "Insert into FAX_Review_Worktable (DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail, SenderPhoneNum, UpdateUser, DocImage,  CnlyClaimNum)" & _
                " Select '" & gbl_DocID & "', 3, FaxNum, Recipient, ICN, " & "'" & gbl_FromFieldForMR & "', " & "PhoneNum, '" & GetUserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                " From QUEUE_MR_Request_Fax" & _
                                " Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"

                strInsertHist = "Insert into FAX_Review_Hist(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,SenderPhoneNum,UpdateUser,DocImage, CnlyClaimNum) " & _
                               " Select '" & gbl_DocID & "', 3, FaxNum, Recipient, ICN, " & "'" & gbl_FromFieldForMR & "', " & "PhoneNum, '" & GetUserName & "','" & strOutputPath & strNewFilePath & ".TIF', CnlyClaimNum" & _
                " From QUEUE_MR_Request_Fax" & _
                                " Where cnlyClaimNum =  '" & !CnlyClaimNum & "'"

                db.Execute (strSetDocID)
                db.Execute (strInsertLetter)
                db.Execute (strDeleteWktb)
                db.Execute (strDeleteQueue)
                db.Execute (strInsertQueue)
                db.Execute (strInsertHist)
                Call updateFaxTables(gbl_DocID, "EFAX", !CnlyClaimNum, 1, "", "")
                
                Set Application.Printer = Application.Printers("Connolly Fax")
                Set prtDefault = Application.Printer
                
                DoCmd.OpenReport "rpt_QUEUE_MR_Request_Fax", , , , acHidden
                'DoCmd.OutputTo acReport, "rpt_CUST_SERV_Cover", acFormatRTF, strCover, False
                 '                   strConCat = FileConcat(strFileList, strFileLoction, strFileName)
                  '                  sFile = DocToPdf(strConCat, True)
                
                'DoCmd.OutputTo acReport, "rpt_QUEUE_MR_Request_Fax", acFormatPDF, testme, False
                'PdfToTiff (testme)
                'deleted = DeleteFile(testme, False)
         MsgBox (faxImage.OutputPath)
         MsgBox (faxImage.ID)
         faxImage.killClass = -1
         Set faxImage = Nothing
        .MoveNext
           intCnt = intCnt + 1
    End With
Wend

Me.Refresh

DoCmd.Hourglass False

'Adding sleep so the files have enough time to build.
Sleep (3000)

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "INC_MR"

Cleanup:
    DoCmd.Hourglass False
    'Set faxImage = Nothing
    Set DocIDRs = Nothing
    Set rsAdo = Nothing
    Set db = Nothing

End Sub

Private Sub cmdGenerateLetter_Click()

GenerateLetter ("All")

End Sub

Private Sub cmdGenerateLetterSingle_Click()

GenerateLetter (Me.frm_QUEUE_MR_Request_Sub.Controls("txtCnlyClaimNum").Value)

End Sub

Private Sub cmdPreview_Click()
Dim strSetDocID As String
Dim db As Database

If CheckFormRecord = 0 Then
    Exit Sub
End If

Set db = CurrentDb
gbl_DocID = generateDocID()
strSetDocID = "Update QUEUE_MR_Request_Fax set DocID = '" & gbl_DocID & "' Where CnlyClaimNum ='" & Form_frm_QUEUE_MR_Request_Sub.txtCnlyClaimNum & "'"
db.Execute (strSetDocID)

DoCmd.OpenReport "rpt_QUEUE_MR_Request_Fax", acViewPreview, , "cnlyClaimNum = '" & Form_frm_QUEUE_MR_Request_Sub.txtCnlyClaimNum & "'"
Form_frm_QUEUE_MR_Request_Sub.id_set = True

End Sub

Private Sub cmdReload_Click()
Call CleanFormLeavingIncompletesOnly

If CheckFormRecord = 0 Then
    Form_frm_QUEUE_MR_Request_Sub.txtMRequested = ""

End If

RecordsetFilter

End Sub

Public Sub CleanSubFields()
        Form_frm_QUEUE_MR_Request_Sub.txtMRequested = ""
        Form_frm_QUEUE_MR_Request_Sub.txtFrom = ""
End Sub


Private Sub cmdSearch_Click()

RecordsetFilter

End Sub


Private Sub cmdSeeClaimInQueue_Click()

Dim strRecSource As String

If CheckFormRecord = 0 Then
    Exit Sub
End If

'VS 3/19/2015 Added DocID to Queue_RECON_Review_Results tables. Joined on DocID in addition to CnlyClaimNum. Added 5 to client ext ref id list.
strRecSource = "SELECT DV.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                        " INNER JOIN QUEUE_MR_Request_Fax AS DV" & _
                        " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum AND FWQ.DocID = DV.DocID" & _
                        " WHERE FWQ.Client_ext_Ref_ID = ""3""" & _
                        " AND FWQ.cnlyClaimNum = '" & Me.frm_QUEUE_MR_Request_Sub.Controls("txtCnlyClaimNum").Value & "'" & _
                        " Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "INC_MR"

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



Forms!frm_AUDITCLM_References_Grid_View.RecordSource = "SELECT * FROM v_AUDITCLM_References WHERE cnlyClaimNum = '" & Me.frm_QUEUE_MR_Request_Sub.Controls("txtCnlyClaimNum").Value & "'"
Forms!frm_AUDITCLM_References_Grid_View.Caption = "Claim Number" & " - " & Me.frm_QUEUE_MR_Request_Sub.Controls("txtCnlyClaimNum").Value

Forms!frm_AUDITCLM_References_Grid_View.Requery
 
End Sub

Private Sub Form_Load()

If UserRights_Inc = "user" Or UserRights_Inc = "Admin" Then

Call CleanFormLeavingIncompletesOnly

On Error GoTo Cleanup

Me.frmFaxStatusHistoryMR.SourceObject = "frm_Fax_Status_History_MR"

Me.frm_QUEUE_Incomplete_MR_Review_Claim_Detail.SourceObject = "frm_QUEUE_Incomplete_MR_Review_Claim_Detail"

Me.Refresh

Me.txtCnlyClaimNumLkUp.SetFocus

Cleanup:
If Err.Number > 0 Then
        MsgBox Err.Number & " " & Err.Description
End If
   
    If gbl_frmLoad = 1 Then gbl_frmLoad = 0

Else
        MsgBox ("You don't have permissions to fax letters requesting additional Medical Records.")
        DoCmd.Close acForm, Me.Name
End If
    
End Sub


Private Sub optAttachSend_AfterUpdate()

Dim strSQL As String

If Me.optAttachSend.Value = 2 Then
    strSQL = "select * FROM QUEUE_MR_Request_Fax WHERE GenerateLetter = True order by ICN"
Else
    strSQL = "select * FROM QUEUE_MR_Request_Fax order by ICN"
   
End If

        Call Form_frm_QUEUE_MR_Request_Sub.SetRecordSource(strSQL)

End Sub


Private Function generateDocID()
    generateDocID = left(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10) & Right(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10)
End Function
