Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Function CheckFormRecord()

CheckFormRecord = Me.RecordSet.recordCount

End Function


Function getDocID()
    getDocID = left(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10) & Right(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 10)
End Function

Private Sub cboProv_AfterUpdate()

Dim db As Database
Dim rsProvAddr As DAO.RecordSet
Dim strAddr As String

strAddr = " select Fax, Firstname & ' ' & LastName AS Recipient from PROV_Address" & _
        " where CnlyProvID = '" & Me.ProvNum & "'" & _
        " AND AddrType = '" & Me.cboProv & "'"

Set db = CurrentDb
Set rsProvAddr = db.OpenRecordSet(strAddr)



With rsProvAddr
    If .recordCount > 0 Then
            Me.FaxNum = Nz(!Fax, "")
            Me.Recipient = !Recipient
    
    Else
            Me.FaxNum = ""
            Me.Recipient = ""
    
    End If
End With

Me.Refresh

Cleanup:
    Set rsProvAddr = Nothing
    Set db = Nothing
        

End Sub

Private Sub cmdAddNotes_Click()

DoCmd.OpenForm "frm_CUST_SERV_Review_Results_AddNotes"
Forms("frm_CUST_SERV_Review_Results_AddNotes").txtCnlyClaimNum = Me.CnlyClaimNum

End Sub

Function ValidateLetter(FaxNumber As Variant, Recipient As Variant, Regading As Variant, From As Variant)

Dim ReturnVal As Integer


    If IsNull(FaxNumber) Or Len(FaxNumber) = 0 Then
      GoTo PromptUser
    End If
    
    If IsNull(Recipient) Or Len(Recipient) = 0 Then
      GoTo PromptUser
    End If
       
    If IsNull(Regading) Or Len(Regading) = 0 Then
     GoTo PromptUser
    End If
      
    If IsNull(From) Or Len(From) = 0 Then
     GoTo PromptUser
    End If
      

ValidateLetter = 1
Exit Function

PromptUser:
    MsgBox "Fax Number, Recipient, Regading and From cannot be blank. Please review and try again.", vbCritical, "Missing Data"

End Function



Private Sub cmdFaxStat_Click()

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
Dim strFileLoction As String
Dim strID As String
Dim strDeleteWktb As String
Dim strDeleteQueue As String
Dim strInsertHist As String
Dim strUpdateQueue As String
Dim strOutputPath As String
Dim strCover As String
Dim strFileLIst As String
Dim strFileName As String
Dim strConCat As String
Dim sFile As String

'Dim myCodeADO As New clsADO

Dim intCnt As Integer

Dim DocIDRs As DAO.RecordSet

If MsgBox("You are about to send the selected letters to the fax queue. Would you like to continue?", vbYesNo + vbQuestion, "Customer Service") = vbNo Then
    GoTo Cleanup
End If


If ValidateLetter(Nz(Me.FaxNum, Null), Nz(Me.Recipient, Null), Nz(Me.Regading, Null), Nz(Me.FromName, Null)) <> 1 Then
        GoTo Cleanup
End If

DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Set db = CurrentDb
'strID = "PRODUAT"
'strID = "DEVUAT"

'DoCmd.Hourglass True

'strGerLTTR = "Select ref.DocID,"

strGerLTTR = "select Ref.DocId, ref.FaxIND, Ref.CnlyClaimNum, Ref.RefLink, RR.FaxNum, RR.Recipient, RR.FromName, RR.Regading from CUST_SERV_FAX_Review_Results_Ref Ref" & _
            " inner Join CUST_SERV_FAX_Review_Results RR" & _
            " On Ref.CnlyClaimNum = RR.CnlyClaimNum" & _
            " where Ref.FaxIND <> 0" & _
            " And Ref.CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
            '" AND AssignedTo Like '" & gbl_sysUser & "'"

Set DocIDRs = db.OpenRecordSet(strGerLTTR)
If (DocIDRs.EOF And DocIDRs.BOF) Then
   MsgBox "You Have no letters to Fax", vbInformation, "Customer Service"
   GoTo Cleanup
End If

intCnt = 0
strID = "PRODUAT"

strFileLoction = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & strID & "'")

While Not DocIDRs.EOF
    With DocIDRs
       
        gbl_DocID = !DocID
        strLTTRID = Format(Now(), "yyyymmddhhmmssms")
        strOutputPath = !RefLink
                
        strFileName = "FAX_" & gbl_DocID & "_" & strLTTRID
                If Right(strOutputPath, 3) <> "Tif" Then
                        sFile = left(strOutputPath, Len(strOutputPath) - 3) & "Tif"
                Else
                        sFile = !DocImage
                End If
            
            If FileExists(sFile) = False Then
                  
                        Select Case Right(strOutputPath, 3)
                                Case "PDF"
                                    sFile = PdfToTiff(strOutputPath)
                                Case "doc"
                                    strCover = strFileLoction & strFileName & ".doc"
                                    strFileLIst = strCover & "|" & strOutputPath
                                    DoCmd.OutputTo acReport, "rpt_CUST_SERV_Cover", acFormatRTF, strCover, False
                                    strConCat = FileConcat(strFileLIst, strFileLoction, strFileName)
                                    sFile = DocToPdf(strConCat, True)
                                Case Is = "Tif"
                                    sFile = !DocImage
                                    
                        End Select
                End If
                
                strDeleteQueue = "Delete FWQ.* from FAX_Work_Queue FWQ" & _
                                " INNER JOIN CUST_SERV_FAX_Review_Results_Ref AS RR" & _
                                " ON FWQ.cnlyClaimNum = RR.cnlyClaimNum" & _
                                " AND FWQ.DocID = RR.DocID" & _
                                " Where FWQ.cnlyClaimNum =  '" & Me.CnlyClaimNum & "'" & _
                                " AND FWQ.DocID = '" & gbl_DocID & "'"
                                
                strDeleteWktb = "Delete FWQ.* from FAX_Review_Worktable FWQ" & _
                                " INNER JOIN CUST_SERV_FAX_Review_Results_Ref AS RR" & _
                                " ON FWQ.cnlyClaimNum = RR.cnlyClaimNum" & _
                                " AND FWQ.DocID = RR.DocID" & _
                                " Where FWQ.cnlyClaimNum =  '" & Me.CnlyClaimNum & "'" & _
                                " AND FWQ.DocID = '" & gbl_DocID & "'"
                                                                                
                strInsertQueue = "Insert into FAX_Review_Worktable(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,UpdateUser,DocImage, CnlyClaimNum) " & _
                                " Select Ref.DocID, 2, RR.FaxNum, RR.Recipient, RR.Regading, RR.FromName, '" & Identity.UserName & "','" & sFile & "', RR.CnlyClaimNum" & _
                                " from CUST_SERV_FAX_Review_Results_Ref Ref" & _
                                " Inner Join CUST_SERV_FAX_Review_Results RR" & _
                                " On Ref.CnlyClaimNum = RR.CnlyClaimNum" & _
                                " where Ref.FaxIND <> 0" & _
                                " And Ref.CnlyClaimNum = '" & Me.CnlyClaimNum & "'" & _
                                " AND Ref.DocID = '" & gbl_DocID & "'"
                                                 
                strInsertHist = "Insert into FAX_Review_Hist(DocID,Client_ext_Ref_ID,RCPTfaxNum,RCPTname,TransmitText,SenderEmail,UpdateUser,DocImage, CnlyClaimNum) " & _
                                " Select Ref.DocID, 2, RR.FaxNum, RR.Recipient, RR.Regading, RR.FromName, '" & Identity.UserName & "','" & sFile & "', RR.CnlyClaimNum" & _
                                " from CUST_SERV_FAX_Review_Results_Ref Ref" & _
                                " Inner Join CUST_SERV_FAX_Review_Results RR" & _
                                " On Ref.CnlyClaimNum = RR.CnlyClaimNum" & _
                                " where Ref.FaxIND <> 0" & _
                                " And Ref.CnlyClaimNum = '" & Me.CnlyClaimNum & "'" & _
                                " AND Ref.DocID = '" & gbl_DocID & "'"
                
                strUpdateQueue = " Update CUST_SERV_FAX_Review_Results_Ref Ref" & _
                                " Set FaxIND = 0" & _
                                " Where Ref.CnlyClaimNum = '" & Me.CnlyClaimNum & "'" & _
                                " AND Ref.DocID = '" & gbl_DocID & "'"
                                                
                db.Execute (strDeleteWktb)
                db.Execute (strDeleteQueue)
                db.Execute (strInsertQueue)
                db.Execute (strInsertHist)
                db.Execute (strUpdateQueue)
                Call updateFaxTables(gbl_DocID, "EFAX", !CnlyClaimNum, 1, "", "")
         
        .MoveNext
           intCnt = intCnt + 1
    End With
Wend

Me.Refresh

'DoCmd.Hourglass False

'Adding sleep so the files have enough time to build.
'Sleep (3000)

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"

Cleanup:
   ' DoCmd.Hourglass False
    Set faxImage = Nothing
    Set DocIDRs = Nothing
    'Set rsAdo = Nothing
    Set db = Nothing

End Sub

Private Sub cmdFaxStatCS_Click()

If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
    DoCmd.Close acForm, "frm_Fax_Selection"
End If

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"

End Sub

Private Sub cmdSeeClaimInQueue_Click()

Dim db As Database
Dim strRecSource As String

'If CheckFormRecord = 0 Then
'    Exit Sub
'End If

Set db = CurrentDb
Dim DocIDRs As DAO.RecordSet

strRecSource = "SELECT RR.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                        " INNER JOIN CUST_SERV_FAX_Review_Results AS RR " & _
                        " ON FWQ.CnlyClaimNum =  RR.CnlyClaimNum" & _
                        " WHERE FWQ.Client_ext_Ref_ID = ""2""" & _
                        " AND FWQ.cnlyClaimNum = '" & Me.CnlyClaimNum & "'" & _
                        " Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"


Set DocIDRs = db.OpenRecordSet(strRecSource)
If (DocIDRs.EOF And DocIDRs.BOF) Then
   MsgBox "This claim cannot be found in the Fax Queue", vbInformation, "Customer Service"
   GoTo Cleanup
End If

If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
    DoCmd.Close acForm, "frm_Fax_Selection"
End If

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"

Forms!frm_Fax_Selection.RecordSource = strRecSource

Forms!frm_Fax_Selection.Requery

Cleanup:
    Set db = Nothing
    Set DocIDRs = Nothing

End Sub



Private Sub Command67_Click()
Me.Refresh
End Sub

Private Sub Form_Load()

Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Dim strSQL As String
Dim OpenArg As String


OpenArg = Me.OpenArgs
'OpenArg = "21129100012102NCA"

DoCmd.ApplyFilter , "cnlyclaimnum = '" & OpenArg & "'"

Me.Caption = "CMS: Claim Number: " & Me.CnlyClaimNum

'Set myAdo = New clsADO
'myAdo.ConnectionString = GetConnectString("v_DATA_Database")
'strsql = "Select * from CUST_SERV_FAX_Review_Results where ICN = '" & OpenArg & "'"
           

'myAdo.SQLstring = strsql
'Set Me.Recordset = myAdo.OpenRecordSet
'myAdo.DisConnect
    
'Set myAdo = Nothing

End Sub
Private Sub cmdSave_Click()
On Error GoTo Err_cmdSave_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSave_Click:
    Exit Sub

Err_cmdSave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSave_Click
    
End Sub
