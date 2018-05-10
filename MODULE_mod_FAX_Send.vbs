Option Compare Database
Option Explicit
Dim aryForms() As Form_frm_INC_MR_Change_Status_Trigger

Function DocToPdf(SFileName As String, Optional bConvertToTif As Boolean)

Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim sFolder As String
Dim sDocFile As String
    
'sDocFile = "\\ccaintranet.com\dfs-cms-ds\data\CMS\AnalystFolders\Curlan\FAX_REPOSITORY\Template\21125700458402MSAEHR.docx"
       

Set oWordApp = New Word.Application
        

Set oWordDoc = oWordApp.Documents.Open(SFileName)
oWordDoc.Activate
With oWordDoc
       .ExportAsFixedFormat OutputFileName:=oWordApp.ActiveDocument.Path & "\" & left(oWordApp.ActiveDocument.Name, InStr(1, oWordApp.ActiveDocument.Name, ".doc") - 1) & ".pdf", _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False
End With

If bConvertToTif = True Then
   sDocFile = PdfToTiff(left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".pdf")
End If


DocToPdf = sDocFile

oWordDoc.Close
Set oWordDoc = Nothing
oWordApp.Quit
Set oWordApp = Nothing

If FileExists(left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".pdf") Then
    Kill left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".pdf"
End If

If FileExists(left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".doc") Then
        Kill left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".doc"
End If

If FileExists(left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".docx") Then
        Kill left(SFileName, InStr(1, SFileName, ".doc") - 1) & ".docx"
End If

End Function


Sub updateFaxTables(sDocID As String, sDocType As String, sCnlyClaimNum As String, sFaxInQueue As Integer, sFaxID As String, sError As String)

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim strRecSelect As String
    
MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_FAX_Update"
                cmd.Parameters.Refresh
                cmd.Parameters("@varDocID") = sDocID
                cmd.Parameters("@varDocType") = sDocType
                cmd.Parameters("@CnlyClaimNum") = sCnlyClaimNum
                cmd.Parameters("@FaxInQueue") = sFaxInQueue
                cmd.Parameters("@FaxID") = sFaxID
                cmd.Parameters("@FaxError") = sError
                cmd.Execute

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Sub


Sub faxDocuments(ClientRef As String, IcnToSend As String, LOB As String, Optional DocID As String)

Dim sFile As String
Dim sFaxNum As String
Dim sLog As String
Dim sEmail As String
Dim sID As String
Dim strFaxID As String
Dim strConCat As String
Dim strCover As String
Dim strFileLIst As String
Dim strFileName As String

Dim strQueuedDocument As String
Dim ofax As New ClsCnlyFax
Dim ofaxImage As New ClsCnlyFaxImage

Dim db As Database
Dim QueuedDocRs As DAO.RecordSet

Dim oWord As New Word.Application
Dim oDoc As New Word.Document
Dim i As Integer

ofax.Host = "cmsfax.smtp.ccaintranet.net"
ofax.Suffix = "@cmsconnollyfax.com"

Set db = CurrentDb
''strQueuedDocument = "Select * from FAX_Work_Queue where Client_ext_Ref_ID = ""1"" and FaxInQueue = True"

Select Case LOB
    Case "RECON"
                'MG 7-2-2013 Need to read recon and recon with appeal from RECON Form
                Select Case IcnToSend
                    Case "All"
                            strQueuedDocument = "SELECT FWQ.* from FAX_Work_Queue AS FWQ" & _
                                        " INNER JOIN v_QUEUE_RECON_Detail_View AS DV" & _
                                        " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum" & _
                                        " WHERE FWQ.Client_ext_Ref_ID IN ('1','4','5','6')" & _
                                        " AND FaxInQueue <> 0" & _
                                        " Order by UpdateDate desc"
                    Case Else
                            strQueuedDocument = "SELECT FWQ.* from FAX_Work_Queue AS FWQ" & _
                                        " INNER JOIN v_QUEUE_RECON_Detail_View AS DV" & _
                                        " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum" & _
                                        " WHERE FWQ.Client_ext_Ref_ID IN ('1','4','5','6')" & _
                                        " AND FaxInQueue <> 0" & _
                                        " AND DV.ICN = '" & IcnToSend & " '" & _
                                        " Order by UpdateDate desc"
                End Select
    Case "CUSTSERV"
             Select Case IcnToSend
                    Case "All"
                            strQueuedDocument = " SELECT * FROM v_FAX_Work_Queue_CustServ FWQ" & _
                                        " WHERE FaxInQueue <> 0" & _
                                        " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by UpdateDate desc"
                    Case Else
                            strQueuedDocument = " SELECT * FROM v_FAX_Work_Queue_CustServ FWQ" & _
                                        " WHERE FaxInQueue <> 0" & _
                                        " AND FWQ.UpdateUser Like '" & gbl_sysUser & "'" & _
                                        " AND FaxInQueue <> 0" & _
                                        " AND FWQ.ICN = '" & IcnToSend & " '" & _
                                        " AND FWQ.DocID = '" & DocID & " '" & _
                                        " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by UpdateDate desc"
            End Select
            
    Case "INC_MR", "ADHOC"
             Select Case IcnToSend
                    Case "All"
                            strQueuedDocument = "SELECT Hdr.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                                        " INNER JOIN AUDITCLM_Hdr AS Hdr" & _
                                        " ON FWQ.CnlyClaimNum = Hdr.CnlyClaimNum" & _
                                        " WHERE FWQ.Client_ext_Ref_ID = '" & ClientRef & "'" & _
                                        " AND FaxInQueue <> 0" & _
                                        " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by UpdateDate desc"
                    Case Else
                            strQueuedDocument = "SELECT Hdr.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                                        " INNER JOIN AUDITCLM_Hdr AS Hdr" & _
                                        " ON FWQ.CnlyClaimNum = Hdr.CnlyClaimNum" & _
                                        " WHERE FWQ.Client_ext_Ref_ID = '" & ClientRef & "'" & _
                                        " AND FaxInQueue <> 0" & _
                                        " AND Hdr.ICN = '" & IcnToSend & " '" & _
                                        " AND FWQ.DocID = '" & DocID & " '" & _
                                        " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by UpdateDate desc"
            End Select

 End Select
 
Set QueuedDocRs = db.OpenRecordSet(strQueuedDocument)

If QueuedDocRs.recordCount = 0 Then
    MsgBox "You have no fax to send at this time.", vbInformation, gbl_MsgBoxTitleLTTR
    GoTo Cleanup
End If

gbl_TriggerFormTotal = 0
gbl_TriggerFormCurrent = 0

With QueuedDocRs
   .MoveFirst
    While Not QueuedDocRs.EOF
    
           If LOB = "INC_MR" Then
              Call CleanFormLeavingIncompletesOnly
              If !Status = gbl_Fax_Status_Send Or !Status = gbl_Fax_Status_Waiting Then
                If CanSendAgain_INC(!CnlyClaimNum) = False Then
                    MsgBox ("You will not be able to send a fax requesting medical records at this point, because the status of claim " & !Icn & " is no longer " & Chr(34) & "Incomplete medical record received" & Chr(34) & ".")
                    GoTo NextRec
                End If
              End If
            End If
    
            gbl_DocID = !DocID
            strFaxID = Format(Now(), "yyyymmddhhmmssms")
            sFile = !DocImage
            sFaxNum = "1" & !RCPTfaxNum
            sEmail = "cms@cms.fax.ccaintranet.net"
            sID = !CnlyClaimNum & "_" & strFaxID
            sLog = ofax.SendFax(sFile, sFaxNum, sEmail, sID)
            Call updateFaxTables(!DocID, "EFAX", !CnlyClaimNum, 1, sID, sLog)
            Sleep (2000)
            
            If LOB = "INC_MR" Then
                Call AddMRToRefTable(!CnlyClaimNum, !DocID)
                gbl_TriggerFormTotal = gbl_TriggerFormTotal + 1
                ReDim Preserve aryForms(i)
               
                Set aryForms(i) = New Form_frm_INC_MR_Change_Status_Trigger
                aryForms(i).DocID = !DocID
                aryForms(i).ClaimNum = !CnlyClaimNum
                aryForms(i).Icn = !Icn
                i = i + 1

            End If
            
NextRec:
                   
        .MoveNext
    Wend
End With

If LOB = "INC_MR" Then
MsgBox "Faxing Complete", vbInformation, gbl_MsgBoxTitleMRLTR
Else
MsgBox "Faxing Complete", vbInformation, gbl_MsgBoxTitleLTTR
End If

Cleanup:
    Set ofax = Nothing
    Set QueuedDocRs = Nothing
    Set db = Nothing

End Sub


Function CanSendAgain_INC(ClaimNum As String) As Boolean

Dim strINC_MRClaims As String
Dim db As Database
Dim incRs As DAO.RecordSet

CanSendAgain_INC = False
Set db = CurrentDb

    strINC_MRClaims = "select CnlyClaimNum from QUEUE_MR_Request_Fax"
    Set incRs = db.OpenRecordSet(strINC_MRClaims)
    
    If incRs.recordCount <> 0 Then
    
        With incRs
           .MoveFirst
            While Not incRs.EOF
            
            If ClaimNum = !CnlyClaimNum Then
            CanSendAgain_INC = True
            End If
                    
            .MoveNext
            Wend
        End With
    
    End If
    
    Set incRs = Nothing
    Set db = Nothing
    
End Function


Function FileConcat(strFileLIst As String, strSavePath As String, strSaveName As String)

Dim oWordApp As New Word.Application
Dim oWordDoc As Word.Document
Dim oWordRange As Word.Range
Dim arrFileList As Variant
Dim intCnt As Integer

'split my string into an array
arrFileList = Split(strFileLIst, "|")
'Create a blank word document
Set oWordDoc = oWordApp.Documents.Add
'get the paragraph count to know where to put the first page break
Set oWordRange = oWordDoc.Paragraphs(oWordDoc.Paragraphs.Count).Range
        

For intCnt = LBound(arrFileList) To UBound(arrFileList)
    'adhoc check for word documents
    If Right(arrFileList(intCnt), 4) = "docx" Or Right(arrFileList(intCnt), 4) = ".doc" Then
            'Check the word count. If it's greater than 1 put in a pagebreak
            If oWordDoc.Words.Count > 1 Then oWordRange.InsertBreak Type:=wdPageBreak
            'Insert the file at the page break
            oWordRange.InsertFile arrFileList(intCnt)
            
            'get the paragraph count to know where to put the next page break
            Set oWordRange = oWordDoc.Paragraphs(oWordDoc.Paragraphs.Count).Range
            
    End If
Next intCnt
  
'MG 9/16/2013 take care of the double space issue since the .doc file was created in 2003 and we are merging the report->.rtf content with .doc and saving it with 2010 MS Word
oWordDoc.Content.ParagraphFormat.SpaceAfter = 0
 

'save the document
'oWordDoc.SaveAs (strSavePath & strSaveName)

'MG 9/16/2013 save it with the original format if possible.
oWordDoc.SaveAs (strSavePath & strSaveName), WdOriginalFormat.wdWordDocument

 
FileConcat = oWordDoc.FullName
 
oWordDoc.Close
Set oWordDoc = Nothing
Set oWordRange = Nothing

oWordApp.Quit
Set oWordApp = Nothing

End Function