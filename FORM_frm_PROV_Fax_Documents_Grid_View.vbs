Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim CnlyProvNum As String

'*******************************************************************************************************************
'MG 9/12/2013 Change variables below if needed
Const CONST_coversheetFileOutputPath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\FAX\CUSTOMER_SERVICE\original_with_cover_sheet\"
Const CONST_convertedFileOutputPath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\FAX\CUSTOMER_SERVICE\converted\"

Const CONST_convertedFormat = "TIF"
'*******************************************************************************************************************



Private Sub clearScreen()
    'Dim lstBoxCount As Integer
    'lstBoxCount = lstSelectedClaims.ListCount
    
    'Dim i As Integer
    'For i = 0 To lstBoxCount - 1
    'For i = 1 To lstSelectedClaims.ListCount - 1
        'Remove an item from the ListBox.
    '    lstSelectedClaims.RemoveItem i
    'Next i
    lstSelectedClaims.RowSource = ""
    createHeaderInListBox
    lblSaveConfirmation.Caption = ""

    
End Sub

Private Sub cboContactType_AfterUpdate()

On Error GoTo ErrHandler
        
    Dim strSQL As String
    Dim recordCount As Integer
    Dim index As Integer
    
    Dim db As Database
    Dim rs As DAO.RecordSet
    
    strSQL = " select fax,firstname,MiddleInit,lastName FROM v_PROV_Address where cnlyProvID='" & CnlyProvNum & "' AND addrType='" & cboContactType.Value & "'"
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(strSQL)
        
    'MG 9/12/2013 From field will always be the user id without the period
    txtFrom.Value = Replace(Identity.UserName, ".", " ")
        
    If rs.recordCount > 0 Then
        txtFaxNumber.Value = Nz(rs.Fields(0), "")
        txtTo.Value = Nz(rs.Fields(1), "") & " " & Nz(rs.Fields(2), "") & " " & Nz(rs.Fields(3), "")
    Else
        txtFaxNumber.Value = ""
        txtTo.Value = ""
    End If
    
    Set rs = Nothing
    Set db = Nothing
     
    'rs.MoveLast
    'recordCount = rs.recordCount 'get record count
    'rs.MoveFirst
    
    'For index = 1 To recordCount
    'For index = 1 To recordCount
    
    '    rs.MoveNext
    'Next index
   
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical
    
End Sub

Function fileTimeStamp() As String
    Dim timeStamp As String
    
    timeStamp = Replace(Now, ":", "")
    timeStamp = Replace(timeStamp, "/", "")
    timeStamp = Replace(timeStamp, " ", "")
    
    fileTimeStamp = timeStamp
    
End Function

Private Sub cmdCheckFaxQueue_Click()

    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"
End Sub


Private Sub cmdClearAllClaims_Click()
    clearScreen
End Sub

Private Sub cmdSave_Click()
    Dim currentRequestNumber As String
    Dim currentLetterType As String
    Dim currentletterDate As String
    Dim currentFilePathLink As String
    

    'MsgBox lstSelectedClaims.ListCount
    If lstSelectedClaims.ListCount > 1 Then
        If Len(Trim(txtRegarding.Value)) > 0 And Len(Trim(txtFaxNumber.Value)) > 0 And Len(Trim(txtTo.Value)) > 0 Then
            'MsgBox lstSelectedClaims.Column(0, 1) 'get first row request number
            'MsgBox lstSelectedClaims.Column(1, 1) 'get first row cnly claim num
            
    
            'MG VBA listbox index starts at 0
            'index starts at 1 because 0 is the header row
            Dim i As Integer
            For i = 1 To lstSelectedClaims.ListCount - 1

                currentRequestNumber = lstSelectedClaims.Column(0, i)
                currentLetterType = lstSelectedClaims.Column(1, i)
                currentletterDate = lstSelectedClaims.Column(2, i)
                currentFilePathLink = lstSelectedClaims.Column(3, i)
                
                'cmdSaveClaims lstSelectedClaims.Column(0, i), lstSelectedClaims.Column(1, i), Me.cboReason.Value, Me.txtOtherReason, DaysExtend
                addDocumentToFax currentRequestNumber, currentLetterType, currentletterDate, currentFilePathLink
                
            Next i
            
            lblSaveConfirmation.Caption = "Saved on " & Now & Chr(13) & Chr(10) & "The system will convert all documents you added to .TIF files. When conversion is completed, it will be added to FAX QUEUE."
        Else
            MsgBox "All fields such as Regarding, Fax Number, Recipient Name and From needs to be filled prior to faxing.", vbCritical, "System Error"
        End If
    End If

End Sub

Function addDocumentToFax(RequestNumber As String, LetterType As String, letterdate As String, FilePathLink As String)

    'converted file in TIF
    Dim backSlashCharPosition As Integer
    Dim newConvertedFileName As String
    Dim newConvertedFileNamePath As String
    
    backSlashCharPosition = InStrRev(FilePathLink, "\", -1)
    newConvertedFileName = Mid(FilePathLink, backSlashCharPosition + 1, Len(FilePathLink)) 'get original file name
    newConvertedFileName = Replace(newConvertedFileName, ".", "-") 'replace . with - This way we will know the original file name if there are any conversion error
    newConvertedFileName = newConvertedFileName & "-" & fileTimeStamp
    
    newConvertedFileNamePath = CONST_convertedFileOutputPath & newConvertedFileName & "." & CONST_convertedFormat
    
    
    'original file with coversheet add on
    Dim strFileLIst As String
    Dim strConCat As String
    Dim CoversheetFileNamePath As String
    Dim coverSheetOriginalContentFileName As String
    
    'MG 9/13/2013 add cover sheet
    'MG store value into global variable, so the access report will read it. This will speed up the creating coversheet process as it bypass MS SQL recordset loop
    gbl_fax_To = txtTo.Value
    gbl_fax_From = txtFrom.Value
    gbl_fax_FaxNumber = txtFaxNumber.Value
    gbl_fax_DocID = createGUID
    gbl_fax_Regarding = txtRegarding.Value
    gbl_fax_Comment = Nz(txtComment.Value, "") 'mg comment field is optional. All the other fields are required.
    
    
    CoversheetFileNamePath = CONST_coversheetFileOutputPath & newConvertedFileName & "-cs.doc" 'MG full path+filename for coversheet
    coverSheetOriginalContentFileName = newConvertedFileName & ".doc" 'MG filename for coversheet with original content (before conversion)
    
    DoCmd.OutputTo acReport, "rpt_CUST_SERV_Cover", acFormatRTF, CoversheetFileNamePath, False 'open report
    strFileLIst = CoversheetFileNamePath & "|" & FilePathLink
    
    'MG 9/19/2013 if it goes too fast, it will not work
    Sleep (3000)

    strConCat = FileConcat(strFileLIst, CONST_coversheetFileOutputPath, coverSheetOriginalContentFileName) 'merge coversheet and report
    
                
    Dim sourceFileToConvert As String
    sourceFileToConvert = CONST_coversheetFileOutputPath & coverSheetOriginalContentFileName
    
    If AddConverterQueueJob(sourceFileToConvert, CONST_convertedFormat, CONST_convertedFileOutputPath, newConvertedFileName, False, False, False, , False) = True Then
        'only add to sql table if file exists
        'mg store original path in fax worktable, so we can use this field to link to the job queue table
        Dim convertedFileFullpath As String
        convertedFileFullpath = CONST_convertedFileOutputPath & newConvertedFileName & "." & CONST_convertedFormat
        AddToCustFaxTable gbl_fax_DocID, txtFaxNumber.Value, txtTo.Value, txtRegarding.Value, txtFrom.Value, gbl_fax_Comment, RequestNumber, convertedFileFullpath
    Else
        LogMessage "addDocumentToFax", "ERROR", "This job failed to be added to the converter queue for some reason - check the logs!"
    End If
    
End Function


Private Sub createHeaderInListBox()
    'List box was created for users to confirm that selected claims are what they want to grant extension to
    Me.lstSelectedClaims.AddItem "RequestNumber,LetterType,LetterReqDt,Link"
End Sub

Private Sub lstSelectedClaims_DblClick(Cancel As Integer)
    'MG get value from selected row
    'MsgBox "row index = " & lstSelectedClaims.ItemsSelected.Item(0)
    'MsgBox lstSelectedClaims.Column(0, lstSelectedClaims.ItemsSelected.Item(0))
    
    'Dim instanceIDHighlighted As String
    'Dim cnlyClaimNumHighlighted As String
    
    'instanceIDHighlighted = lstSelectedClaims.Column(0, lstSelectedClaims.ItemsSelected.Item(0))
    'cnlyClaimNumHighlighted = lstSelectedClaims.Column(1, lstSelectedClaims.ItemsSelected.Item(0))
    
    'MsgBox instanceIDHighlighted
    'MsgBox cnlyClaimNumHighlighted
    If lstSelectedClaims.ItemsSelected.Count > 0 Then
        lstSelectedClaims.RemoveItem (lstSelectedClaims.ItemsSelected.Item(0))
    End If
    
End Sub

Private Sub Form_Load()

    'Dim testGUID
    'testGUID = createGUID
        
    'MG add header to list box
    clearScreen
    
     'Store data on temporary table for ms access to pick it up
    CnlyProvNum = Me.Parent!txtCnlyProvID.Value
    
    'documentLookup "040010" 'testing
    
    documentLookup CnlyProvNum
    
    'MG 09-10-2013 filter based on session ID
    Dim sqlString As String
    sqlString = " SessionID = " & Chr(34) & Identity.UserName & Chr(34)
               
    'MG refresh data sheet
    frm_PROV_Fax_Documents_Lookup.Form.filter = sqlString
    frm_PROV_Fax_Documents_Lookup.Form.FilterOn = True
    frm_PROV_Fax_Documents_Lookup.Form.Requery
    frm_PROV_Fax_Documents_Lookup.Form.Refresh
    
    Me.frm_PROV_Fax_Documents_Lookup.SetFocus
    
End Sub

Function documentLookup(cnlyProvID As String)

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_PROV_FAX_Document_Lookup"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyProvID") = cnlyProvID
    cmd.Execute
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
End Function

Function AddToCustFaxTable(thisDocID As String, thisFaxNum As String, thisRecipient As String, thisRegarding As String, thisFrom As String, thisComment As String, thisInstanceID As String, thisImageFilePath As String)
    
    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As Variant
    Dim ErrMsg As String
    Dim sIcn As String
    Dim sClaimNum As String
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                    cmd.commandType = adCmdStoredProc
                    cmd.CommandText = "usp_CUST_SERV_Load_FaxTbl_Prov_V2"
                    cmd.Parameters.Refresh
                    cmd.Parameters("@pDocID") = thisDocID
                    cmd.Parameters("@pFaxNum") = thisFaxNum
                    cmd.Parameters("@pRecipient") = thisRecipient
                    cmd.Parameters("@pFromName") = thisFrom
                    cmd.Parameters("@pRegarding") = thisRegarding
                    cmd.Parameters("@pComment") = thisComment
                    cmd.Parameters("@pInstanceID") = thisInstanceID
                    cmd.Parameters("@pImageFilePath") = thisImageFilePath
                    cmd.Execute
                    spReturnVal = cmd.Parameters("@Return_Value")
                    
    'MsgBox RefLink
    
    If spReturnVal <> 0 Then
        ErrMsg = cmd.Parameters("@ErrMsg")
        MsgBox ErrMsg, vbCritical, "Customer Service"
    End If
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
End Function

'MG 9/16/2013 use VBA to create GUID so to pass onto Report and store in SQL table
Function createGUID() As String
    'Dim TypeLib As Object
    'Dim Guid As String
    'Set TypeLib = CreateObject("Scriptlet.TypeLib")
    'Guid = TypeLib.Guid
    ' format is {24DD18D4-C902-497F-A64B-28B2FA741661}
    'Guid = Replace(Guid, "{", "X")
    'Guid = Replace(Guid, "}", "X")
    'Guid = Replace(Guid, "-", "X")
    'createGUID = left(Guid, 20) 'MG 9/16/2013 DocID is setup as varchar 20 in all the fax tables, so this need to be 20 max length.
                      'Don't ask me bc I wasn't the one who created the table structure...I'm just doing my job so I won't get fired

    'MG 9/19/2013 above is disable because the guid could create characters that VBA doesn't like during EML Receipt File, which will give a query error expression. The same SQL string will work fine in MS SQL directly though
    Dim randomNum As Integer
    randomNum = Int((100 - 1 + 1) * Rnd + 1) 'create random number from 1 to 100
    
    createGUID = fileTimeStamp + CStr(randomNum)
    
End Function
