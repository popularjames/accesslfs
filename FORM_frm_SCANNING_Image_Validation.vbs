Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean

Dim mstrCalledFrom As String




Const CstrFrmAppID As String = "ImageVal"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub cmdValidate_Click()
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rsScanImages As ADODB.RecordSet
    Dim rsImageConfig As ADODB.RecordSet
    Dim cmd As ADODB.Command
      
    Dim strValidatedPath As String
    Dim strDailyScansPath As String
    Dim strImageOutPath As String

    
    Dim strDailyScanFile As String
    Dim strValidatedFile As String
    Dim strImageOutFile As String
    Dim strImageError As String
    Dim strByPassMsg As String
    Dim strFileExt As String
    
    Dim bImageError As Boolean
    
    Dim iPageCnt As Integer
    Dim iTotalImage As Long
    Dim iErrCnt As Long
    
    
    Dim strErrMsg As String
    Dim strErrSource As String
    
    Dim strSQLcmd As String
    
    Dim fso As FileSystemObject
    
    Dim bResult As Boolean
    Dim bFileExists As Boolean
    Dim bMoveFile As Boolean
    Dim iResult As Long
    Dim iSearchCnt As Integer
    
    
    Dim bDelImage As Boolean
    Dim dbCount As Database
    Dim rsCount As RecordSet
    Dim strSQL As String
    
    
    On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandText = "usp_SCANNING_Update_Single_Image"
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    
    strErrSource = "cmdValdation"
    
    ' reset display info
    lblStatus.Caption = ""
    lblError.Caption = ""
    lblTotalCnt.Caption = ""
    iErrCnt = 0
    iTotalImage = 0
    
    
    '------------------------------------------------------------
    ' VALIDATE IMAGE PATHS
    '------------------------------------------------------------
    ' get image paths
    Set rsImageConfig = MyAdo.OpenRecordSet("select * from SCANNING_Config where AccountID = " & gintAccountID)
    If rsImageConfig.EOF = True Then
        MsgBox "Scanning configuration for account " & gstrAcctDesc & " has not been set up.  Please set it up and re-try", vbCritical
        GoTo Exit_Sub
    End If
    
    
    strValidatedPath = rsImageConfig("LocalPath") & ""
    strDailyScansPath = rsImageConfig("LocalHoldPath") & ""
    strImageOutPath = rsImageConfig("RemotePath") & ""
    
    
    
    ' check image paths
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(strValidatedPath) = False Then
        strErrMsg = "Temporary image hold path '" & strValidatedPath & "' does not exists or in accessible.  Please check."
        GoTo Err_handler
    End If
    
    If fso.FolderExists(strDailyScansPath) = False Then
        strErrMsg = "Temporary image hold path '" & strValidatedPath & "' does not exists or in accessible.  Please check."
        GoTo Err_handler
    End If
    
    If fso.FolderExists(strImageOutPath) = False Then
        strErrMsg = "Remote image path '" & strImageOutPath & "' does not exists or in accessible.  Please check."
        GoTo Err_handler
    End If
    
    If Right$(strValidatedPath, 1) <> "\" Then strValidatedPath = strValidatedPath & "\"
    If Right$(strDailyScansPath, 1) <> "\" Then strDailyScansPath = strDailyScansPath & "\"
    If Right$(strImageOutPath, 1) <> "\" Then strImageOutPath = strImageOutPath & "\"
    
    
    
    '------------------------------------------------------------
    ' MAIN BODY: VALIDATE IMAGES
    '------------------------------------------------------------
    strSQLcmd = "SELECT ScannedDt, CnlyClaimNum, ErrMsg, PageCnt, ImageType, '' as ImagePath, " & _
                "       ImageName , cnlyProvID, ReceivedDt, ReceivedMeth, ScanOperator " & _
                "FROM SCANNING_Image_Log_Tmp " & _
                "WHERE 1=2"
        
    Me.subSCANNING_Image_Error.Form.RecordSource = strSQLcmd
   
    
    
    strSQLcmd = "select * from SCANNING_Image_Log_Tmp where LocalPath like '" & strValidatedPath & "%' and AccountID = " & gintAccountID
    
    Set rsScanImages = MyAdo.OpenRecordSet(strSQLcmd, False)
    
    mbContinue = True
    
    If Not (rsScanImages.BOF = True And rsScanImages.EOF = True) Then
        With rsScanImages
            .MoveFirst
            Do While Not .EOF
                ' init variable before processing earch file
                bImageError = False
                bFileExists = False
                bMoveFile = True
                strByPassMsg = ""
                strImageError = ""
                strFileExt = ""
                
                On Error GoTo Err_handler
                
                iTotalImage = iTotalImage + 1
                Me.lblTotalCnt.Caption = "Total scanned: " & CStr(iTotalImage)
                Me.lblStatus.Caption = "Processing " & !ImageName
                
                ' TKL 9/21/2011 modify for HP process
                If InStr(1, ".PDF/.TIF", UCase(Right(!ImageName, 4))) = 0 Then strFileExt = ".TIF"
                strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
                
                iSearchCnt = 0
                bFileExists = fso.FileExists(strDailyScanFile)
                Do While bFileExists = False And iSearchCnt <= 6
                    iSearchCnt = iSearchCnt + 1
                    
                    Select Case iSearchCnt
                        Case 1
                            If strFileExt <> "" Then strFileExt = ".TIF"
                            strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
                        Case 2
                            If strFileExt <> "" Then strFileExt = ".PDF"
                            strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
                        Case 3
                            ' check validation directory. Files may have been moved there from previous run
                            If strFileExt <> "" Then strFileExt = ".TIF"
                            strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
                        Case 4
                            ' check validation directory. Files may have been moved there from previous run
                            If strFileExt <> "" Then strFileExt = ".PDF"
                            strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
                        Case 5
                            ' check image out directory. Files may have been moved there by SuperFlex
                            If strFileExt <> "" Then strFileExt = ".TIF"
                            strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
                        Case 6
                            ' check image out directory. Files may have been moved there by SuperFlex
                            If strFileExt <> "" Then strFileExt = ".PDF"
                            strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
                    End Select
                    
                    bFileExists = fso.FileExists(strDailyScanFile)
                    
                Loop
                
                If bFileExists = True And iSearchCnt >= 3 Then
                    ' do not move file if it's already in validated/out folder
                    bMoveFile = False
                End If
                

                    
                    
                If bFileExists Then
                    '------------------------------
                    'image exists do error checking
                    '------------------------------
                    strValidatedFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
                    strImageOutFile = strImageOutPath & !cnlyProvID & "\" & !ImageName & strFileExt

                    ' check image and log entry for error
                    If FileLocked(strDailyScanFile) Or fso.GetFile(strDailyScanFile).DateCreated > DateAdd("n", -5, Now()) Then
                        ' check page count
                        strImageError = "Error: file being copied"
                        bImageError = True
                    ' check image and log entry for error
                    ElseIf UCase(!ReceivedMeth) <> "HP" And UCase(!ReceivedMeth) <> "FileMatch" And !PageCnt = 0 Then
                        ' check page count
                        strImageError = "Error: page count is zero"
                        bImageError = True
'                    ElseIf UCase(!ReceivedMeth) <> "HP" And UCase(!ReceivedMeth) <> "FileMatch" And !PageCnt > 10 And !ImageType = "INV" And !ErrMsg <> "BYPASS - Alert: Invoice is more than 10 pages." Then
'                        ' check for invoice greater than 10 pages
'                        strImageError = "Alert: Invoice is more than 10 pages."
'                        bImageError = True
'                    ElseIf UCase(!ReceivedMeth) <> "HP" And UCase(!ReceivedMeth) <> "FileMatch" And !PageCnt <= 10 And !ImageType = "MR" And !ErrMsg <> "BYPASS - Alert: Medical record is less than 15 pages." Then
'                        ' check for invoice greater than 10 pages
'                        strImageError = "Alert: Medical record is less than 15 pages."
'                        bImageError = True
                    ElseIf UCase(!ReceivedMeth) <> "HP" And UCase(!ReceivedMeth) <> "FileMatch" And UCase(!ReceivedMeth) <> "ESMD" And left(!ScanStation, 4) <> "FAST" And left(!ImageName, Len(!cnlyProvID)) <> !cnlyProvID And !ErrMsg <> "BYPASS - Error: CnlyProvID does not match image name" Then
                        ' check for mis-match in claim number and image name
                        strImageError = "Error: CnlyProvID does not match image name"
                        bImageError = True
                    ElseIf UCase(!ReceivedMeth) <> "HP" And UCase(!ReceivedMeth) <> "FileMatch" And UCase(!ReceivedMeth) <> "ESMD" And (Mid(!ImageName, Len(!cnlyProvID) + 1, Len(!ImageType)) <> !ImageType And Mid(!ImageName, Len(!cnlyProvID) + 1, Len(Nz(!OrigImageType, ""))) <> Nz(!OrigImageType, "")) And !ErrMsg <> "BYPASS - Error: Image type does not match image name" Then
                        ' check for mis-match in image type and image name
                        strImageError = "Error: Image type does not match image name"
                        bImageError = True
                    Else
                        'no error yet.  Check page count
                        If UCase(strFileExt) = ".TIF" Then
                            'JS 20130308 Using now Damons new tiff page count function
                            iPageCnt = TifPageCount(strDailyScanFile) 'Count_TIF_Pages(strDailyScanFile)
                        Else
                            iPageCnt = Count_PDF_Pages(strDailyScanFile)
                        End If
                    
                        If iPageCnt < !PageCnt And left(!ErrMsg, 37) <> "BYPASS - Error: Page count mismatched" Then
                            bImageError = True
                            strImageError = "Error: Page count mismatched " & CStr(iPageCnt) & "/" & CStr(!PageCnt)
                        End If
                    End If   ' end of error checking
                    
                Else
                    '-------------------------------------------
                    'image does not exists.  set error
                    '-------------------------------------------
                    strValidatedFile = strValidatedPath & !cnlyProvID & "\" & !ImageName
                    strImageOutFile = strImageOutPath & !cnlyProvID & "\" & !ImageName
                    
                    strImageError = "Image does not exists" '"Image " & !ImageName & " does not exists"
                    bImageError = True
                End If
                

                
  
                                
                
                
                If Not (bImageError) Then
                    '--------------------------------------------------------
                    'no error up to this point.   Ready to process image
                    '--------------------------------------------------------
                    
                    ' set bypass message
                    If left(!ErrMsg, 9) = "BYPASS - " Then
                        strByPassMsg = !ErrMsg
                    Else
                        strByPassMsg = ""
                    End If
                    
                    'JS 12/12/2013 delete image only if there is no other row in image_log_tmp that needs it.
                    bDelImage = True
                    strSQL = "SELECT COUNT(ImageName) AS RwCnt FROM SCANNING_IMAGE_LOG_TMP where imagename = '" & !ImageName & "'"
                    
                    Set dbCount = CurrentDb
                    Set rsCount = dbCount.OpenRecordSet(strSQL, dbOpenSnapshot, dbForwardOnly)
                    
                    If rsCount.recordCount = 1 Then
                        If Nz(rsCount!RwCnt, 0) > 1 Then
                            bDelImage = False
                        End If
                    End If
                    'JS 12/12/2013
                    
                    myCode_ADO.BeginTrans

                    ' update database and move claim to next queue
                    On Error GoTo Err_handler
                    
                    cmd.Parameters("@pScannedDt") = ConvertTimeToString(!ScannedDt)
                    cmd.Parameters("@pCnlyClaimNum") = !CnlyClaimNum
                    cmd.Parameters("@pLocalImageName") = strValidatedFile
                    cmd.Parameters("@pRemoteImageName") = strImageOutFile
                    cmd.Parameters("@pPageCount") = iPageCnt
                    cmd.Parameters("@pUserMsg") = ""
                    cmd.Parameters("@pByPassMsg") = strByPassMsg
                    cmd.Parameters("@pAccountID") = gintAccountID
                    cmd.Execute
                    
                    iResult = cmd.Parameters("@RETURN_VALUE")
                    strErrMsg = Trim(cmd.Parameters("@pErrMsg").Value)
                    If iResult <> 0 Or strErrMsg <> "" Then
                        bImageError = True
                        strImageError = "Error advancing image to next queue. " & strErrMsg
                    End If
                    
                    
                    On Error Resume Next        ' from this point forward we want to capture the error via code
                    
                    ' check validated folder and create if not exists
                    If Not (bImageError) Then
                        If fso.FolderExists(strValidatedPath & !cnlyProvID) = False Then
                            bResult = CreateFolder(strValidatedPath & !cnlyProvID)
                            If bResult = False Then
                                bImageError = True
                                strImageError = "Error creating folder: " & strValidatedPath & !cnlyProvID
                            End If
                        End If
                    End If
                    
                    
                    ' delete existing file before moving file
                    'If Not (bImageError) Then
                    '    If bMoveFile Then
                    '        If fso.FileExists(strValidatedFile) Then
                    '            Call fso.DeleteFile(strValidatedFile, True)
                    '            If fso.FileExists(strValidatedFile) Then
                    '                bImageError = True
                    '                strImageError = "Error deleting image " & strValidatedFile
                    '            End If
                    '        End If
                    '    End If
                    'End If
                    
                    
                    
                    ' move image
                    If Not (bImageError) Then
                        If bMoveFile Then
                            Call fso.CopyFile(strDailyScanFile, strValidatedFile, True)
                            If fso.FileExists(strValidatedFile) = False Then
                                'image move is not successful. Rollback
                                myCode_ADO.RollbackTrans
                                bImageError = True
                                strImageError = "Error moving image " & strDailyScanFile
                            Else
                                If bDelImage Then
                                    Call fso.DeleteFile(strDailyScanFile)
                                End If
                            End If
                        End If
                    End If
                    


                    If bImageError Then
                        myCode_ADO.RollbackTrans
                    Else
                        myCode_ADO.CommitTrans
                    End If
                End If
                
                
                If bImageError Then
                    '------------------------------------------------------------------
                    ' there is some error with the image.  Update image with error
                    '------------------------------------------------------------------
                    
                    On Error GoTo Err_handler
                    
                    iErrCnt = iErrCnt + 1
                    lblError.Caption = "Err Count: " & iErrCnt
                    
                    myCode_ADO.BeginTrans
                    
                    cmd.Parameters("@pScannedDt") = ConvertTimeToString(!ScannedDt)
                    cmd.Parameters("@pCnlyClaimNum") = !CnlyClaimNum
                    cmd.Parameters("@pLocalImageName") = strValidatedFile
                    cmd.Parameters("@pRemoteImageName") = strImageOutFile
                    cmd.Parameters("@pPageCount") = iPageCnt
                    cmd.Parameters("@pUserMsg") = strImageError
                    cmd.Parameters("@pByPassMsg") = ""
                    cmd.Parameters("@pAccountID") = gintAccountID
                    cmd.Execute
                    
                    iResult = cmd.Parameters("@RETURN_VALUE")
                    strErrMsg = Trim(cmd.Parameters("@pErrMsg").Value)
                    If iResult <> 0 Or strErrMsg <> "" Then
                        myCode_ADO.RollbackTrans
                        GoTo Err_handler
                    Else
                        myCode_ADO.CommitTrans
                    End If
                    
                End If


                DoEvents
                DoEvents
                DoEvents
                DoEvents
                DoEvents
                
                If mbContinue = False Then Exit Do
                    
                .MoveNext
            Loop
        End With
    End If
    
    
    ' display errors
    strSQLcmd = "SELECT ScannedDt, CnlyClaimNum, ErrMsg, PageCnt, ImageType, " & _
                " IIf(UCase(Right([LocalPath],4)) Not In ('.TIF','.PDF'),'','#' & Replace([LocalPath],'\Validated','\DailyScans')) AS ImagePath, " & _
                "       ImageName , cnlyProvID, ReceivedDt, ReceivedMeth, ScanOperator " & _
                " FROM SCANNING_Image_Log_Tmp " & _
                " WHERE ErrMsg <> '' and AccountID = " & gintAccountID & _
                " order by ScannedDt"
    
    Me.subSCANNING_Image_Error.Form.RecordSource = strSQLcmd
    Me.subSCANNING_Image_Error.Form.Requery
    Me.subSCANNING_Image_Error.visible = True
    
    mbContinue = True
   
    
    
    If mbContinue = True Then
        lblStatus.Caption = "Scanning completed"
    Else
        lblStatus.Caption = "Scanned stopped"
    End If
    
    
    MsgBox lblStatus.Caption
    
Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set fso = Nothing
    Set cmd = Nothing
    Set rsImageConfig = Nothing
    Set rsScanImages = Nothing
    Set dbCount = Nothing
    Set rsCount = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Number & " - " & Err.Description
    MsgBox "Error in module " & strErrSource & vbCrLf & vbCrLf & strErrMsg
    
    Resume Exit_Sub
End Sub

Private Sub cmsStop_Click()
    mbContinue = False
End Sub



Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    Dim strSQLcmd As String
    
    Me.Caption = "Image Validation"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    lblStatus.Caption = ""
    lblError.Caption = ""
    lblTotalCnt.Caption = ""
        
    strSQLcmd = "SELECT ScannedDt, CnlyClaimNum, ErrMsg, PageCnt, ImageType, '' as ImagePath, " & _
                "       ImageName , cnlyProvID, ReceivedDt, ReceivedMeth, ScanOperator " & _
                "FROM SCANNING_Image_Log_Tmp " & _
                "WHERE 1=2"
        
    Me.subSCANNING_Image_Error.Form.RecordSource = strSQLcmd
    Me.subSCANNING_Image_Error.Form.Requery
    
    Me.subSCANNING_Image_Error.visible = True
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
End Sub
