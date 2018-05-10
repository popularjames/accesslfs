Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database
Private bCancel As Boolean

Private Sub cmdBrowse_Click()
Me.txtArchivePath = GetDirectory("Select Location")
End Sub

Private Sub CmdCancel_Click()
    bCancel = True
End Sub

Private Sub cmdImageArchive_Click()
    If IsDate(Me.txtMaxDate) Then
        If DateDiff("m", Me.txtMaxDate, Date) < 6 Then
            If MsgBox("archive date is less than 6 months from current.  Is this OK?", vbQuestion + vbYesNo) = vbYes Then
                Archive_Image Me.txtMaxDate
                MsgBox "Done"
            Else
                MsgBox "Cancelled"
            End If
        Else
            Archive_Image Me.txtMaxDate
            MsgBox "Done"
        End If
    End If
End Sub

Private Sub Archive_Image(strArchiveDate As String)
    Dim MyAdo As New clsADO
    Dim rs As ADODB.RecordSet
    Dim rsCheck As ADODB.RecordSet
    
    Dim fso As New FileSystemObject
    Dim f As file
    
    Dim strArchivePath As String
    Dim strArchiveFolder As String
    Dim strFinalPath As String
    Dim strArchiveImage As String
    Dim strArchiveProvFolder As String
    Dim strOrigImage As String
    Dim strOrigFilePath As String
    Dim strCnlyProvID As String
    Dim strSQL As String
    Dim strScannedDt As String
    Dim strLogFile As String
    
    Dim iFileNum
    Dim iArchivedSoFar As Long
    Dim iResult As Long
    Dim bResult As Boolean
    Dim bErrFlag As Boolean
    
    Dim strRecordAsRefSubType As String
                        
    
    On Error GoTo Err_handler
    
    bErrFlag = False
    strLogFile = ""
    Close
    
    bCancel = False
    
    strArchivePath = Me.txtArchivePath
    strArchiveFolder = Me.txtArchiveFolder
    
'    If Me.chkOCROnly.Value = 0 And Me.chkPreOCRonly <> 0 Then
'        MsgBox "You must select OCR images only if order to select PRE_OCR images only!"
'        Exit Sub
'    End If
    
    
    If Not IsDate(strArchiveDate) Then
    
        MsgBox "The archive date you entered is invalid!", vbExclamation, "Error in Date"
        Exit Sub
    End If
    
    If Me.Frame25.Value = 1 Then
         'Recovery
         If MsgBox("You have chosen to Archive RECOVERY images.  Are you sure?", vbQuestion + vbYesNo) = vbNo Then
             Exit Sub
         End If
    End If
    
     Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_SCANNING_Archive_GetImagesToArchive"
        .Parameters.Refresh
        .Parameters("@pArchiveOlderThan") = Format(strArchiveDate, "yyyy-mm-dd")
        .Parameters("@pArchiveRecovery") = Me.Frame25.Value
        Set oRs = .ExecuteRS

        ErrorReturned = Nz(.Parameters("@pErrMsg").Value, "")
        If ErrorReturned <> "" Then
            'LogMessage strProcName, "ERROR", "Problem searching for a concept", "Keyword: " & Nz(Me.txtSearchBox, "") & " Expand Search: " & IIf(Me.ckExpandSearch, 1, 0) & " Include Codes: " & IIf(Me.ckIncludeCodes, 1, 0)
            MsgBox ErrorReturned, vbExclamation
            Exit Sub
        End If
        If .GotData = False Then
            MsgBox "No Images to Archive Found.", vbExclamation
            Exit Sub
        End If
        Set rs = oRs
        
        'Me.cmbIssueType.Recordset.Requery
    End With
    
        
    
    
    
    
    
    
    
'
'        strSQL = "select distinct ah.CnlyProvID, ar.* " & _
'                 " from AUDITCLM_References ar " & _
'                 " join AUDITCLM_Hdr ah " & _
'                 "    on ar.CnlyClaimNum = ah.CnlyClaimNum " & _
'                 "   JOIN CMS_Auditors_Claims.dbo.XREF_ClaimStatus SS ON ss.ClmStatus = ah.ClmStatus " & _
'                 " where cast(ar.CreateDt as date) <= '" & strArchiveDate & "'" & _
'                 " and ar.RefType = 'IMAGE' " & _
'                 " and RefSubType IN ( 'MR', 'OCR', 'PDFNOOCR', 'esMDSource', 'esMDComb', 'BKMRSRC', 'BKMR', 'OCRDuplicate', 'WrongImage' )  " & _
'                 " and relatedclaimmatch = 0  "
'
'                 'aa.cnlyClaimNum = ar.cnlyClaimNum    and
'
'       If Me.chkOCROnly.Value <> 0 And Me.chkPreOCRonly = 0 Then
'            strSQL = strSQL & " and RefSubType IN ('OCR', 'PDFNOOCR')  "
'       ElseIf Me.chkOCROnly.Value <> 0 And Me.chkPreOCRonly <> 0 Then
'            strSQL = strSQL & " and RefSubType IN ('OCR')  "
'       Else
'            strSQL = strSQL & " and not exists ( select 1 from  CMS_AUDITORS_CLAIMS.dbo.Scanning_Image_Archive aa where ( aa.OriginalImagePath = ar.RefLink or aa.newimagepath = ar.reflink) )  "
'       End If
'
'       If Me.Frame25.Value = 1 Then
'            'Recovery
'            If MsgBox("You have chosen to Archive RECOVERY images.  Are you sure?", vbQuestion + vbYesNo) = vbNo Then
'                Exit Sub
'            End If
'            strSQL = strSQL & " and ss.ClmStatusGroup IN " & Option28.Tag
'
'       ElseIf Me.Frame25.Value = 2 Then
'            'Non Recovery
'            strSQL = strSQL & " and ss.ClmStatusGroup IN " & Option30.Tag & _
'                        " and not exists (select 1 from cms_auditors_claims.dbo.auditclm_status ss where ss.cnlyclaimnum = ar.cnlyclaimnum and ss.clmstatus in ('320','330','322','402','353','354') ) " & _
'                        " and not exists (select 1 from CMS_AUDITORS_Reports.dbo.rpt_r0045d ap where ap.cnlyclaimnum = ar.cnlyclaimnum)  " & _
'                        " and ss.Lifecyclegroup not in ('op','ar','hd','pd','py') "
'       Else
'            Err.Raise 6500, , "No image type chosen"
'       End If
'
'        MyAdo.sqlString = strSQL & " ORDER BY CreateDt ASC OPTION (MAXDOP 1)"
'
'        Set rs = MyAdo.OpenRecordSet
        
        If rs.recordCount > 0 Then
            ' create log folder
            If Right(Trim(strArchivePath), 1) <> "\" Then strArchivePath = strArchivePath & "\"
            strFinalPath = strArchivePath & strArchiveFolder
            
            If Not FolderExists(strFinalPath) Then
                If Not CreateFolder(strFinalPath) Then
                    MsgBox "Can not create folder " & strFinalPath
                    Exit Sub
                End If
            End If
            
            ' create log file
            iFileNum = FreeFile()
            strLogFile = strFinalPath & "\Image_Archive_Log " & Format(Now(), "yyyy-mm-dd hhmmss") & ".log"
            Open strLogFile For Output As #iFileNum
            Print #iFileNum, "Archive to ::" & strFinalPath
            
            'archive file
            iArchivedSoFar = 0
            
            Me.lblTotalArchive.Caption = "Total Records: " & CStr(rs.recordCount)
            Me.lblArchivedSoFar.Caption = "Archived So Far: 0"
            
            rs.MoveFirst
            
            While Not rs.EOF
                strOrigImage = rs("RefLink")
                
'                    If InStr(1, strOrigImage, "imagearchive", vbTextCompare) > 0 Then
'                        Stop
'                    End If
                    

                    
                    Print #iFileNum, "archiving claim :: " & rs("CnlyClaimNum") & " - file ::" & strOrigImage
                    
                    strCnlyProvID = rs("CnlyProvID")
                    'Debug.Print "Orig image   " & strOrigImage
                    strOrigFilePath = fso.GetParentFolderName(strOrigImage)
                    'Debug.Print "Orig path    " & strOrigFilePath
                    strArchiveProvFolder = strFinalPath & "\" & strCnlyProvID
                    'Debug.Print "Archive Path  " & strArchiveProvFolder
                    strArchiveImage = strArchiveProvFolder & "\" & fso.GetFileName(strOrigImage)
                    'Debug.Print "Archive image " & strArchiveImage
                    
                    ' update image path in AUDITCLM_References
                    strScannedDt = ConvertTimeToString(rs("CreateDt"))
                    

                    strRecordAsRefSubType = rs("RefSubType")
                    
                    'If Me.chkOCROnly.Value <> 0 And Me.chkPreOCRonly <> 0 And Len(rs("RefLink")) > 4 Then
                        Dim strJustTheFileName As String
                        Dim strNewPREOCRFilename As String
                        Dim strPreOCROrigImage As String
                        Dim strPreOCRArchiveImage As String
                        strJustTheFileName = GetFileName(strOrigImage)
                        
                        
                        
                        strNewPREOCRFilename = left(strJustTheFileName, Len(strJustTheFileName) - 4) & "PRE_OCR.pdf"
                        strPreOCROrigImage = "\\ccaintranet.com\DFS-CMS-FLD\Imaging\Client\Out\CMS\MedicalRecords\" & strCnlyProvID & "\" & strNewPREOCRFilename
                        If Not fso.FileExists(strPreOCROrigImage) Then
                            strNewPREOCRFilename = left(strJustTheFileName, Len(strJustTheFileName) - 4) & "PRE_OCR.tif"
                            strPreOCROrigImage = "\\ccaintranet.com\DFS-CMS-FLD\Imaging\Client\Out\CMS\MedicalRecords\" & strCnlyProvID & "\" & strNewPREOCRFilename
                        End If
                        If Not fso.FileExists(strPreOCROrigImage) Then
                            strNewPREOCRFilename = left(strJustTheFileName, Len(strJustTheFileName) - 4) & " PRE_OCR.pdf"
                            strPreOCROrigImage = "\\ccaintranet.com\DFS-CMS-FLD\Imaging\Client\Out\CMS\MedicalRecords\" & strCnlyProvID & "\" & strNewPREOCRFilename
                        End If
                        If Not fso.FileExists(strPreOCROrigImage) Then
                            strNewPREOCRFilename = left(strJustTheFileName, Len(strJustTheFileName) - 4) & " PRE_OCR.tif"
                            strPreOCROrigImage = "\\ccaintranet.com\DFS-CMS-FLD\Imaging\Client\Out\CMS\MedicalRecords\" & strCnlyProvID & "\" & strNewPREOCRFilename
                        End If
                        If Not fso.FileExists(strPreOCROrigImage) Then
                            'GoTo Next_Image
                            strPreOCROrigImage = ""
                        Else
                            strPreOCRArchiveImage = Replace(strArchiveImage, strJustTheFileName, strNewPREOCRFilename)
                        End If
                        
                    'End If
                    
                
                If fso.FileExists(strOrigImage) And strArchiveImage <> strOrigImage Then

                    
                    MyAdo.BeginTrans
                    
                    'If Not (Me.chkOCROnly.Value <> 0 And Me.chkPreOCRonly <> 0) Then
                        strSQL = " update AUDITCLM_References " & _
                                 " set RefLink = '" & strArchiveImage & "' " & _
                                 " where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' and " & _
                                 " RefLink = '" & strOrigImage & "' " & _
                                 " and RefType = 'IMAGE' "
                                 
                                 
                                 ' "where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' " & _
                                 ' "and CreateDt = '" & strScannedDt & "' " & _
                                 ' "and RefType = 'IMAGE'"

                        MyAdo.sqlString = strSQL
                        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                        MyAdo.SQLTextType = sqltext
                        
                        iResult = MyAdo.Execute
                        If iResult < 1 Then
                            Print #iFileNum, Space(5) & "ERROR: Can not update AUDITCLM_References"
                            MyAdo.RollbackTrans
                            bErrFlag = True
                            GoTo Next_Image
                        End If
                    
                    'End If
                    
                    ' update image path in SCANNING_Image_Log & SCANNING_Image_Error_Log
                    If rs("ExistsInImageLog") = 1 Then
                        'update SCANNING_Image_Log
'                        strSQL = "select * from SCANNING_Image_Log " & _
'                                 "where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' " & _
'                                 "and ScannedDt = '" & strScannedDt & "' "
'
'                        MyAdo.sqlString = strSQL
'                        Set rsCheck = MyAdo.OpenRecordSet
'                        If rsCheck.recordCount > 0 Then
                            strSQL = "update SCANNING_Image_Log " & _
                                     "set ImagePath = '" & strArchiveImage & "' " & _
                                     "where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' " & _
                                     "and imagepath = '" & rs("reflink") & "' "
                    
                            MyAdo.sqlString = strSQL
                            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                            MyAdo.SQLTextType = sqltext

                            iResult = MyAdo.Execute
                            If iResult < 1 Then
                                Print #iFileNum, Space(5) & "ERROR: Can not update SCANNING_Image_Log"
                                MyAdo.RollbackTrans
                                bErrFlag = True
                                GoTo Next_Image
                            End If
                        End If
                    
                        ' update SCANNING_Image_Error_Log
'                        strSQL = "select * from SCANNING_Image_Error_Log " & _
'                                 "where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' " & _
'                                 "and ScannedDt = '" & strScannedDt & "' "
'
'                        MyAdo.sqlString = strSQL
'                        Set rsCheck = MyAdo.OpenRecordSet
'                        If rsCheck.recordCount > 0 Then
                            strSQL = "update SCANNING_Image_Error_Log " & _
                                     "set ImagePath = '" & strArchiveImage & "' " & _
                                     "where CnlyClaimNum = '" & rs("CnlyClaimNum") & "' " & _
                                     "and imagepath = '" & rs("reflink") & "' "
                    
                            MyAdo.sqlString = strSQL
                            iResult = MyAdo.Execute
'                            If iResult < 1 Then
'                                Print #iFileNum, Space(5) & "ERROR: Can not update SCANNING_Image_Error_Log"
'                                MyAdo.RollbackTrans
'                                bErrFlag = True
'                                GoTo Next_Image
'                            End If
                        'End If
                    
                    'End If
                    
                    ' move image
                    If CreateFolder(strArchiveProvFolder) Then
                        Set f = fso.GetFile(strOrigImage)
                        
                        If f.Attributes And ReadOnly Then
                            f.Attributes = f.Attributes - ReadOnly      ' read only file
                        End If
                        Set f = Nothing
                        
                        DoEvents
                        DoEvents
                        DoEvents
                        DoEvents
                        DoEvents
                        DoEvents
                        
                        fso.CopyFile strOrigImage, strArchiveImage, True
                        
                        If fso.FileExists(strArchiveImage) Then
                            fso.DeleteFile (strOrigImage)
                            If fso.FileExists(strOrigImage) Then
                                Print #iFileNum, "ERROR: Can not delete original image"
                                MyAdo.RollbackTrans
                                bErrFlag = True
                                GoTo Next_Image
                            End If
                            
                            If strPreOCROrigImage <> "" Then
                                If fso.FileExists(strPreOCROrigImage) Then
                                    fso.CopyFile strPreOCROrigImage, strPreOCRArchiveImage, True
                                    If fso.FileExists(strPreOCRArchiveImage) Then
                                        Call LogArchive(rs("cnlyClaimNum"), strOrigImage, strArchiveImage, rs("RefType"), "PRE_OCR")
                                        Call fso.DeleteFile(strPreOCROrigImage, True)
                                    End If
                                End If
                            End If
                        Else
                            Print #iFileNum, "ERROR: Can not copy image to destination"
                            MyAdo.RollbackTrans
                            bErrFlag = True
                            GoTo Next_Image
                        End If
                    Else
                        Print #iFileNum, "ERROR: can not create provider achive folder " & strArchiveProvFolder
                        MyAdo.RollbackTrans
                        bErrFlag = True
                        GoTo Next_Image
                    End If
                    
                    If Not LogArchive(rs("cnlyClaimNum"), strOrigImage, strArchiveImage, rs("RefType"), strRecordAsRefSubType) Then
                        Print #iFileNum, "ERROR: Could not log the archive " & strArchiveProvFolder
                        MyAdo.RollbackTrans
                        bErrFlag = True
                        GoTo Next_Image
                    End If
                    MyAdo.CommitTrans
                    iArchivedSoFar = iArchivedSoFar + 1
                    lblImageArchive.Caption = strOrigImage
                    lblArchivedSoFar.Caption = "Archived So Far: " & CStr(iArchivedSoFar)
                    DoEvents
                    
                End If
                
Next_Image:
                If bCancel Then
                    GoTo Exit_Sub
                End If
                rs.MoveNext
            Wend
            
        End If
    
      
    
Exit_Sub:

    lblImageArchive.Caption = strOrigImage
    lblArchivedSoFar.Caption = "Archived So Far: " & CStr(iArchivedSoFar)
    
    If iFileNum <> "" Then
        Print #iFileNum, "Total archived: " & CStr(iArchivedSoFar)
    End If
    Set MyAdo = Nothing
    Set rs = Nothing
    Set fso = Nothing
    Close
    
    If bErrFlag = True Then MsgBox "Error encountered.  Check your log"
    
    Exit Sub

Err_handler:
    bErrFlag = True
    If strLogFile <> "" Then
        Print #iFileNum, "ERROR NUMBER [" & CStr(Err.Number) & "].  ERROR DESCRIPTION [" & Err.Description & "]"
    Else
        MsgBox Err.Description
    End If
    
    MyAdo.RollbackTrans
    
    Resume Exit_Sub
End Sub


Private Sub Form_Load()
    Me.lblImageArchive.Caption = ""
    Me.lblTotalArchive.Caption = ""
    Me.lblArchivedSoFar.Caption = ""
    Me.txtArchiveFolder = Format(Date, "yyyy-mm-dd")
End Sub

Private Function LogArchive(strCnlyClaimNum As String, _
                        stroriginalPath As String, _
                        strNewImage As String, _
                        strRefType As String, _
                        strRefSubType As String) As Boolean

    
    On Error GoTo ErrHandler


    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command

    ' get data
    Set myCode_ADO = New clsADO
    Set cmd = New ADODB.Command
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandText = "dbo.usp_SCANNING_IMAGE_Archive_Insert"
    
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters("@cnlyClaimNum") = strCnlyClaimNum
    cmd.Parameters("@ArchiveDate") = Now
    cmd.Parameters("@OriginalImagePath") = stroriginalPath
    cmd.Parameters("@NewImagePath") = strNewImage
    cmd.Parameters("@RefType") = strRefType
    cmd.Parameters("@RefSubType") = strRefSubType
    cmd.Parameters("@pERRMsg") = ""
    
    cmd.Execute
    
    If cmd.Parameters("@pErrMsg") <> "" Then
    
        LogArchive = False
        Exit Function
    End If

    LogArchive = True

Exit Function

ErrHandler:
    LogArchive = False
End Function
