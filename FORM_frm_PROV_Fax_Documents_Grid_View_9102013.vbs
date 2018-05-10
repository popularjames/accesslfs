Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrFilePath As String
Private strRowSource As String
Private strAppID As String
Private mstrFieldReference As String
Private mstrFieldValue As String
Private mstrAttachmentType As String

Const CstrFrmAppID As String = "ProvRef" 'mg need to fix this later for permisison issue. I was intially gonna copy the code exactly from customer service form, but alot of codes aren't really needed. The interface may look the same, but the VBA and SQL a bit different.

Public Property Get FieldReference() As String
    FieldReference = mstrFieldReference
End Property
Property Let FieldReference(data As String)
     mstrFieldReference = data
End Property
Public Property Get FieldValue() As String
    FieldValue = mstrFieldValue
End Property
Property Let FieldValue(data As String)
     mstrFieldValue = data
End Property
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CnlyRowSource(data As String)
     strRowSource = data
End Property
Property Get CnlyRowSource() As String
     CnlyRowSource = strRowSource
End Property

Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
    'Refresh the grid based on the rowsource passed into the form
    Me.RecordSource = CnlyRowSource
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub


Private Sub cboProv_AfterUpdate()

On Error GoTo ErrHandler
        
    Dim strSQL As String
    Dim recordCount As Integer
    Dim index As Integer
    
    Dim db As Database
    Dim rs As DAO.RecordSet
    
    strSQL = " select fax,firstname,MiddleInit,lastName FROM v_PROV_Address where cnlyProvID='" & Me.Parent!txtCnlyProvID.Value & "' AND addrType='" & cboProv.Value & "'"
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(strSQL)
        
    If rs.recordCount > 0 Then
        FaxNum.Value = Nz(rs.Fields(0), "")
        Recipient.Value = Nz(rs.Fields(1), "") & " " & Nz(rs.Fields(2), "") & " " & Nz(rs.Fields(3), "")
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



Private Sub cmdFaxStat_Click()
    'MG this is a modified version of the original. The original codes have too many hardcoded SQL in VBA, so now most of the SQL is in usp
        
    'step 1 insert or update data into CMS_AUDITORS_CLAIMS.dbo.CUST_SERV_FAX_Review_Results_Ref
    Dim validateStatus As String
    If ValidateLetter(FaxNum, Recipient, Regading, FromName) = 1 Then
        lblMessages.Caption = "Converting file(s) to .TIF Please wait . . ."
        validateStatus = "good"
    End If
        
    If validateStatus = "good" Then
        prepareCustServTables FaxNum, Recipient, Regading, FromName, Nz(Comment, "")
        MsgBox "The process may take a while because .doc files will need to be converted to .docx, .pdf and then .tif", vbExclamation, "The screen will appear unresponsive. Please do not exist the screen."
    End If
    
    'step 2 conversion of files
    'MG will clean below code and transform it into sp
    Dim db As Database
    
    Dim faxImage As ClsCnlyFaxImage
    
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
    Dim newDocImage As String
    Dim strNewOutputPath As String
    Dim intCnt As Integer
    Dim DocIDRs As DAO.RecordSet
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    Set db = CurrentDb
    
    Sleep (3000) 'mg allow time to ensure all data are updated
    
    strGerLTTR = " select * from v_CUST_SERV_FAX_DocID"
    
    Set DocIDRs = db.OpenRecordSet(strGerLTTR)
    
    If DocIDRs.recordCount > 0 Then
        DocIDRs.MoveLast
        'MsgBox "recordCount = " & DocIDRs.recordCount
        DocIDRs.MoveFirst
        'MG not sure but it seems that by adding this part...access get all rows
    End If
    
    If (DocIDRs.EOF And DocIDRs.BOF) Then
       MsgBox "You have no letters to fax.", vbInformation, "Customer Service"
       GoTo Cleanup
    End If
    
    intCnt = 0
    strID = "PRODUAT"
    
    strFileLoction = DLookup("[Location]", "FAX_FileLocation", "[PathId] ='" & strID & "'")
    Sleep (3000)

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
                    
                    'MG update ref to uncheck box
                    strUpdateQueue = " Update CUST_SERV_FAX_Review_Results_Ref Ref" & _
                                     " Set FaxIND = 0" & _
                                     " Where Ref.DocID = '" & gbl_DocID & "'"
                     
                    'MG update worktable with new file name
                    strNewFilePath = strFileLoction & strFileName & ".tif"
                    strNewOutputPath = " Update FAX_Review_Worktable wtb" & _
                                     " Set docImage = '" & strNewFilePath & "'" & _
                                     " Where wtb.DocID = '" & gbl_DocID & "'"
                                     
                    db.Execute (strUpdateQueue)
                    db.Execute (strNewOutputPath)
                    
                    
                    Call updateFaxTables(gbl_DocID, "EFAX", !CnlyClaimNum, 1, "", "")
             
            .MoveNext
               intCnt = intCnt + 1
        End With
    Wend
    
    Me.Refresh
    
    'DoCmd.Hourglass False
    
    'Adding sleep so the files have enough time to build.
    'Sleep (3000)
    lblMessages.Caption = "File(s) converted. Completed."
    
    DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"
    
Cleanup:
       ' DoCmd.Hourglass False
        Set faxImage = Nothing
        Set DocIDRs = Nothing
        Set db = Nothing
    
    
    
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



Private Sub processCustServReference()

    On Error GoTo ErrHandler
    
        
        
        Dim strSQL As String
        Dim recordCount As Integer
        Dim index As Integer
        
        Dim db As Database
        Dim rs As DAO.RecordSet
        
        strSQL = "SELECT maxRefLink,instanceID,ProvNum FROM PROV_FAX_Document WHERE Include<>0 AND SessionID=" & Chr(34) & Identity.UserName & Chr(34)
        Set db = CurrentDb
        Set rs = db.OpenRecordSet(strSQL)
        
        Dim thisMaxRefLink As String
        Dim thisInstanceID As String
        Dim thisProvNum As String
        
        rs.MoveLast
        recordCount = rs.recordCount 'get record count
        rs.MoveFirst
        
        If rs.recordCount > 0 Then
        
        
                                
            'MG Prepare CUST_SERV_FAX_Review_Results and CUST_SERV_FAX_Review_Results_Ref
            For index = 1 To recordCount
            
                thisMaxRefLink = Nz(rs.Fields(0), "")
                thisInstanceID = Nz(rs.Fields(1), "")
                thisProvNum = Nz(rs.Fields(2), "") 'MG you can also get this from the form
                
                prepareCustServTables thisMaxRefLink, thisInstanceID, thisProvNum, FromName, Comment
                
                rs.MoveNext
            Next index
        End If
        
        Set rs = Nothing
        Set db = Nothing
       

       
    Exit Sub
    
ErrHandler:
        MsgBox Err.Description, vbOKOnly + vbCritical
        
End Sub


Private Sub cmdView_Click()
    Dim strFileName As String
    strFileName = Me.RecordSet("MaxRefLink")
    SetFileReadOnly (strFileName)
    If UCase(Right(strFileName, 3)) = "TIF" Then
    
        Shell "explorer.exe " & strFileName, vbNormalFocus
        ' TK: Removed to work with new TS-CMS server
'        If UCase(left(GetPCName(), 9)) = "TS-FLD-03" Then
'            Shell "explorer.exe " & strFileName, vbNormalFocus
'        Else
'            Shell "c:\program files\Common Files\Microsoft Shared\MODI\11.0\mspview.exe " & strFileName, vbNormalFocus
'        End If
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If
End Sub


Private Sub Form_Load()
        
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
 
     'Store data on temporary table for msaccess to pick it up
    documentLookup Me.Parent!txtCnlyProvID.Value
    Me.RecordSource = "SELECT * FROM PROV_FAX_Document WHERE SessionID=" & Chr(34) & Identity.UserName & Chr(34)
    
    
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

Function prepareCustServTables(thisFaxNum As String, thisRecipient As String, thisRegarding As String, thisFrom As String, thisComment As String)
    
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
                    cmd.CommandText = "usp_CUST_SERV_Load_FaxTbl_Prov"
                    cmd.Parameters.Refresh
                    cmd.Parameters("@pFaxNum") = thisFaxNum
                    cmd.Parameters("@pRecipient") = thisRecipient
                    cmd.Parameters("@pFromName") = thisFrom
                    cmd.Parameters("@pRegarding") = thisRegarding
                    cmd.Parameters("@pComment") = thisComment
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


Private Sub cmdFaxStatCS_Click()

If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
    DoCmd.Close acForm, "frm_Fax_Selection"
End If

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"

End Sub
