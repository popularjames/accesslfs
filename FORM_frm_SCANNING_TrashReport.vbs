Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "ImageVal"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub

Private Sub Form_Load()
    lblImageName.Caption = ""
    lblTrashMessage.Caption = ""
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
End Sub
Private Sub txtTrashBarCode_AfterUpdate()

    Call RunValidation
    Call RecordValidationLog(Identity.UserName, Me.txtTrashBarCode, Me.lblTrashMessage.ForeColor, Me.lblTrashMessage.Caption)
    
End Sub

Sub RunValidation()

    Dim rsImageValidation As New ADODB.RecordSet
    Dim rsClaimStopTrash As New ADODB.RecordSet
    Dim oRs As New ADODB.RecordSet
    Dim strSQL As String
    Dim bFileExists As Boolean

On Error GoTo Err_handler
 
    Me.lblTrashMessage.Caption = "searching..."
    Me.lblTrashMessage.ForeColor = vbBlack
    Dim AddtlMessage As String
    Dim strValidatedPath As String
    Dim strDailyScansPath As String
    Dim strImageOutPath As String
    Dim strErrMsg As String
    Dim strFileExt As String
    Dim strDailyScanFile As String
    Dim iSearchCnt As Integer
    Dim dtScannedDt As Date
    
    Dim oAdo As clsADO
    
    StatusBar "searching... (check StopTrash table, Stop Trash)"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select top 1 CAMessage from SCANNING_Image_Log il " & _
                        " join SCANNING_Claim_StopTrash st on st.cnlyclaimnum = il.cnlyclaimnum " & _
                        " Where il.imagetype ='MR' and il.ImageName Like '" & Mid(Me.txtTrashBarCode, 6) & "%'"
        
        Set rsClaimStopTrash = oAdo.OpenRecordSet()
        If Not rsClaimStopTrash Is Nothing Then
            If Not rsClaimStopTrash.BOF And Not rsClaimStopTrash.EOF Then
                rsClaimStopTrash.MoveFirst
                Me.lblTrashMessage.Caption = rsClaimStopTrash("CAMessage")
                Me.lblTrashMessage.ForeColor = vbRed
                Beep
                GoTo Exit_Sub
            End If
        End If
    End With
    
    StatusBar "searching... (check StopTrash table, Image Lost)"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select 1 from SCANNING_Image_Log il " & _
                        " join SCANNING_ImageName_StopTrash st on st.ImageName = replace(replace(il.imagename,'.PDF',''),'.TIF','') " & _
                        " Where il.ImageName Like '" & Mid(Me.txtTrashBarCode, 6) & "%'"
        
        Set oRs = .ExecuteRS
        If .GotData = True Then
            Me.lblTrashMessage.Caption = "IMAGE WAS LOST! MUST be sent back to be rescanned."
            Me.lblTrashMessage.ForeColor = vbRed
            Beep
            GoTo Exit_Sub
        End If
    End With
 
    StatusBar "searching... (check SCANNING_Image_Log table)"
 
'If KeyCode = 9 Then
    Set MyAdo = New clsADO
    'MsgBox Me.txtTrashBarCode
    strSQL = " select * from SCANNING_Image_Log where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'" '& " and isnull(ValidationDt,'1/1/1900') <> '1/1/1900'"
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    MyAdo.sqlString = strSQL
    
    'open the audit claims header and disconnect
    Set rsImageValidation = MyAdo.OpenRecordSet()
    
    lblImageName.Caption = Mid(Me.txtTrashBarCode, 6)
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Me.lblTrashMessage.Caption = "Not Validated"
    Me.lblTrashMessage.ForeColor = vbRed
    
    bFileExists = False
    

    
    'This is for image that are in Scanning_Image_log, that means it should be on its way to be attached to a claim already, we just need to check the validation stage/status
    If Not rsImageValidation.EOF Then
        
        bFileExists = fso.FileExists(rsImageValidation("ImagePath"))
        dtScannedDt = rsImageValidation("ScannedDt")
        
        If Not bFileExists Then
            Me.lblTrashMessage.Caption = "Not Validated" & vbNewLine & "Image is NOT in Wilton." & vbNewLine & "SCANNED since " & rsImageValidation("ScannedDt") & "."
            Me.lblTrashMessage.ForeColor = vbRed
            Beep
        ElseIf Not IsNull(rsImageValidation("ValidationDt")) And rsImageValidation("ValidationDt") <> "1/1/1900" And bFileExists Then
            Me.lblTrashMessage.Caption = "VALIDATED"
            Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
            Beep
            'GoTo Exit_sub
        ElseIf bFileExists And (IsNull(rsImageValidation("ValidationDt")) Or rsImageValidation("ValidationDt") = "1/1/1900") Then
            Me.lblTrashMessage.Caption = "Not Validated" & vbNewLine & "Image is waiting for 2nd validation process."
            Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
            Beep
            'GoTo Exit_sub
        End If

    End If
        
    'continues the analysis, we might find more data or add more data
    Dim bScanned As Boolean
    bScanned = False
    
    'can we find it in scanning_image_log_tmp?, if so it might just be waiting there or stuck with an error
    strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    MyAdo.sqlString = strSQL
    Dim rsImageLogTmp As ADODB.RecordSet
    Set rsImageLogTmp = MyAdo.OpenRecordSet()
            
    If Not rsImageLogTmp.EOF Then
        dtScannedDt = rsImageLogTmp("ScannedDt")
        Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "SCANNED since " & dtScannedDt & "."
        bScanned = True
        If Nz(rsImageLogTmp("ErrMsg"), "") <> "" Then
            Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "NOT validating because: " & rsImageLogTmp("ErrMsg")
        End If
    End If
    
    'if it is a fastscan image and has not been matched yet, or it is a no match / finish just sitting there
    If InStr(1, Mid(Me.txtTrashBarCode, 4), "FAST") Then
    
        StatusBar "searching... (FastScan route)"

        Set MyAdo = New clsADO
        strSQL = " select * from SCANNING_fastscan_log where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
        MyAdo.ConnectionString = GetConnectString("v_Data_Database")
        MyAdo.sqlString = strSQL
        
        Dim rsFastScan As ADODB.RecordSet
        Set rsFastScan = MyAdo.OpenRecordSet()
        
        'has not even been linked yet
        If rsFastScan.EOF Then
            Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "FastScan Image NOT LINKED"
            Me.lblTrashMessage.ForeColor = vbRed
            Beep
            GoTo Exit_Sub
        End If
        
        'deleted fastscan image, let it go
        If left(rsFastScan("ProcStatusCd"), 7) = "DELETED" Then
            Me.lblTrashMessage.Caption = "VALIDATED - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
            Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
            Beep
            GoTo Exit_Sub
        End If
        
'        If Not bScanned Then
'            Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "SCANNED since " & rsFastScan("ScannedDt") & "."
'        End If
        
        Dim rsFastScanImageConfig As ADODB.RecordSet
        Set rsFastScanImageConfig = MyAdo.OpenRecordSet("select * from SCANNING_FastScan_Config where AccountID = " & rsFastScan("AccountID"))
        If rsFastScanImageConfig.EOF = True Then
            MsgBox "FastScan configuration for account " & gstrAcctDesc & " has not been set up.  Please set it up and re-try", vbCritical
            GoTo Exit_Sub
        End If

        Dim strFastScanStatus As String
        strFastScanStatus = rsFastScan("ProcStatusCd")

        Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "Image was linked to: " & Nz(DLookup("AcctDesc", "Admin_Client_Account", "accountID = " & rsFastScan("AccountID") & ""), "")
        Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "FastScan status is --> " & strFastScanStatus & " since: " & rsFastScan("ProcStatusLastUpDt")
        Me.lblTrashMessage.ForeColor = vbRed

        If strFastScanStatus <> "ATTACHED" Then
        
            Dim strFastScanMatchInPath As String
    
            strFastScanMatchInPath = rsFastScanImageConfig("MatchInPath") & ""
            If Right$(strFastScanMatchInPath, 1) <> "\" Then strFastScanMatchInPath = strFastScanMatchInPath & "\"
            
            
            Dim strFastScanMatchInFile As String

            
            strFileExt = Nz(rsFastScan("FileExt"), "")
            
            If Not bFileExists Then
                If strFileExt = "" Then
                    strFileExt = "TIF"
                    strFastScanMatchInFile = strFastScanMatchInPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & "." & strFileExt
                    bFileExists = fso.FileExists(strFastScanMatchInFile)
                    If Not bFileExists Then
                        strFileExt = "PDF"
                        strFastScanMatchInFile = strFastScanMatchInPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & "." & strFileExt
                        bFileExists = fso.FileExists(strFastScanMatchInFile)
                    End If
                Else
                    strFastScanMatchInFile = strFastScanMatchInPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & "." & strFileExt
                    bFileExists = fso.FileExists(strFastScanMatchInFile)
                End If
            End If
        
        Else
        
            Dim strClaimAttachPath As String
            Dim strProviderAttachPath As String
            Dim strAttachCnlyClaimNum As String
            Dim strAttachCnlyProvID As String
            Dim strFastScanClaimAttachFile As String
            Dim strFastScanProviderAttachFile As String
        
            strClaimAttachPath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = 'AuditClmRef' AND AccountID = " & gintAccountID), "")
            'strClaimAttachPath = rsFastScanImageConfig("ClaimAttachPath") & ""
            If Right$(strClaimAttachPath, 1) <> "\" Then strClaimAttachPath = strClaimAttachPath & "\"
            
            'strProviderAttachPath = rsFastScanImageConfig("ProviderAttachPath") & ""
            strProviderAttachPath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = 'ProvRef' AND AccountID = " & gintAccountID), "")
            If Right$(strProviderAttachPath, 1) <> "\" Then strProviderAttachPath = strProviderAttachPath & "\"
            
            strAttachCnlyClaimNum = Nz(rsFastScan("CnlyClaimNum"), "")
            strAttachCnlyProvID = DLookup("CnlyProvID", "auditclm_hdr", "CnlyClaimnum = '" & Nz(rsFastScan("CnlyClaimNum"), "") & "' and AccountID = " & gintAccountID)
        
            strClaimAttachPath = strClaimAttachPath & "CnlyClaimNum\FastScan\"
            strProviderAttachPath = strProviderAttachPath & "CnlyProvID\" & strAttachCnlyProvID & "\"
            
            If Not bFileExists Then
                strFileExt = "TIF"
                'strFastScanClaimAttachFile = strClaimAttachPath & strAttachCnlyClaimNum & "\" & rsFastScan("ImageName") & "." & strFileExt
                strFastScanClaimAttachFile = strClaimAttachPath & rsFastScan("ImageName") & "." & strFileExt
                bFileExists = fso.FileExists(strFastScanClaimAttachFile)
                If Not bFileExists Then
                    strFileExt = "PDF"
                    'strFastScanClaimAttachFile = strClaimAttachPath & strAttachCnlyClaimNum & "\" & rsFastScan("ImageName") & "." & strFileExt
                    strFastScanClaimAttachFile = strClaimAttachPath & rsFastScan("ImageName") & "." & strFileExt
                    bFileExists = fso.FileExists(strFastScanClaimAttachFile)
                End If
            End If
            
            If Not bFileExists Then
                strFileExt = "TIF"
                'strFastScanProviderAttachFile = strProviderAttachPath & strAttachCnlyProvID & "\" & rsFastScan("ImageName") & "." & strFileExt
                strFastScanProviderAttachFile = strProviderAttachPath & rsFastScan("ImageName") & "." & strFileExt
                bFileExists = fso.FileExists(strFastScanProviderAttachFile)
                If Not bFileExists Then
                    strFileExt = "PDF"
                    'strFastScanProviderAttachFile = strAttachCnlyProvID & strAttachCnlyProvID & "\" & rsFastScan("ImageName") & "." & strFileExt
                    strFastScanProviderAttachFile = strProviderAttachPath & rsFastScan("ImageName") & "." & strFileExt
                    bFileExists = fso.FileExists(strFastScanProviderAttachFile)
                End If
            End If
            
        End If
    End If
    
    If bFileExists Then
        Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "Image IS in Wilton"
        Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
    Else
        Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & IIf(dtScannedDt = "12:00:00 AM", "", "Image IS NOT in Wilton for " & DateDiff("d", dtScannedDt, Now) & " days")
        Me.lblTrashMessage.ForeColor = vbRed
    End If
        
'        StatusBar "searching... (check Config Tables)"
'
'        AddtlMessage = ""
'
'        Dim rsImageConfig As ADODB.Recordset
'        Set rsImageConfig = MyAdo.OpenRecordSet("select * from SCANNING_Config where AccountID = " & gintAccountID)
'        If rsImageConfig.EOF = True Then
'            MsgBox "Scanning configuration for account " & gstrAcctDesc & " has not been set up.  Please set it up and re-try", vbCritical
'            GoTo Exit_sub
'        End If
'
'        strValidatedPath = rsImageConfig("LocalPath") & ""
'        strDailyScansPath = rsImageConfig("LocalHoldPath") & ""
'        strImageOutPath = rsImageConfig("RemotePath") & ""
'
'        Set rsImageConfig = Nothing
'
'        Set rsImageConfig = MyAdo.OpenRecordSet("select * from SCANNING_FastScan_Config where AccountID = " & gintAccountID)
'        If rsImageConfig.EOF = True Then
'            MsgBox "FastScan configuration for account " & gstrAcctDesc & " has not been set up.  Please set it up and re-try", vbCritical
'            GoTo Exit_sub
'        End If
'
'        Dim strFastScanMatchInPath As String
'
'        strFastScanMatchInPath = rsImageConfig("MatchInPath") & ""
'
'        If fso.FolderExists(strValidatedPath) = False And AccountID = 1 Then 'only check for CMS, MCR operators cannot see this folder
'            strErrMsg = "Validated image hold path '" & strValidatedPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strDailyScansPath) = False And AccountID = 1 Then 'only check for CMS, MCR operators cannot see this folder
'            strErrMsg = "Temporary image hold path '" & strDailyScansPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strImageOutPath) = False Then
'            strErrMsg = "Remote image path '" & strImageOutPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strFastScanMatchInPath) = False Then
'            strErrMsg = "FastScan Match In Path '" & strFastScanMatchInPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'
'        If Right$(strValidatedPath, 1) <> "\" Then strValidatedPath = strValidatedPath & "\"
'        If Right$(strDailyScansPath, 1) <> "\" Then strDailyScansPath = strDailyScansPath & "\"
'        If Right$(strImageOutPath, 1) <> "\" Then strImageOutPath = strImageOutPath & "\"
'        If Right$(strFastScanMatchInPath, 1) <> "\" Then strFastScanMatchInPath = strFastScanMatchInPath & "\"
'
'        Me.lblTrashMessage.Caption = "NOT VALIDATED"
'        'Me.lblTrashMessage.ForeColor = vbRed
'        Beep
'
'        'if it is a fastscan image
'        If InStr(1, Mid(Me.txtTrashBarCode, 4), "FAST") Then
'
'            StatusBar "searching... (FastScan route)"
'
'            Set MyAdo = New clsADO
'            strSQL = " select * from SCANNING_fastscan_log where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'
'            Dim rsFastScan As ADODB.Recordset
'            Set rsFastScan = MyAdo.OpenRecordSet()
'
'            'has not even been linked yet
'            If rsFastScan.EOF Then
'                Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption + " - FASTSCAN IMAGE NOT LINKED"
'                Me.lblTrashMessage.ForeColor = vbRed
'                Beep
'                GoTo Exit_sub
'            End If
'
'            'deleted fastscan image, let it go
'            If left(rsFastScan("ProcStatusCd"), 7) = "DELETED" Then
'                Me.lblTrashMessage.Caption = "VALIDATED - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
'                Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
'                Beep
'                GoTo Exit_sub
'            End If
'
'            Dim strFastScanMatchInFile As String
'
'            strFastScanMatchInFile = strFastScanMatchInPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & "." & rsFastScan("FileExt")
'
'            bFileExists = fso.FileExists(strFastScanMatchInFile)
'
'            'other fastscan statuses that are not Matched, they should be sitting in the fastscan folder
'            If rsFastScan("ProcStatusCd") <> "MATCHED" Then
'                Me.lblTrashMessage.Caption = "Not Validated - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
'                Me.lblTrashMessage.ForeColor = vbRed
'
'                If bFileExists Then
'                    Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "File IS in Wilton"
'                    Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
'                Else
'                    Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & "Image IS NOT in Wilton for " & DateDiff("d", rsFastScan("ScannedDt"), Now) & " days."
'                    Me.lblTrashMessage.ForeColor = vbRed
'                End If
'
'                Beep
'                GoTo Exit_sub
'            End If
'
'
'            'continues only for MATCHED FastScan coversheets
'            Me.lblTrashMessage.Caption = rsFastScan("procstatuscd") & " since " & rsFastScan("scanneddt") & "." & vbNewLine & Me.lblTrashMessage.Caption
'
'            strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
'            Set MyAdo = New clsADO
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'            Dim rsFastScanImages As ADODB.Recordset
'            Set rsFastScanImages = MyAdo.OpenRecordSet()
'
'
'            If strFileExt <> "" Then strFileExt = ".TIF"
'            strDailyScanFile = strImageOutPath & !cnlyProvID & "\" & !ImageName & strFileExt
'
'            ' check image out directory. Files may have been moved there by SuperFlex
'            If strFileExt <> "" Then strFileExt = ".PDF"
'            strDailyScanFile = strImageOutPath & !cnlyProvID & "\" & !ImageName & strFileExt
'
'
'            If Not (rsFastScanImages.BOF = True And rsFastScanImages.EOF = True) Then
'
'                StatusBar "searching... (Checking FastScan matched images)"
'
'                rsFastScanImages.MoveFirst
'
'                With rsFastScanImages
'                    ' TKL 9/21/2011 modify for HP process
'                    If InStr(1, ".PDF/.TIF", UCase(Right(!ImageName, 4))) = 0 Then strFileExt = ".TIF"
'                    strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'
'                    iSearchCnt = 0
'
'                    bFileExists = fso.FileExists(strDailyScanFile)
'                    Do While bFileExists = False And iSearchCnt <= 6
'                        iSearchCnt = iSearchCnt + 1
'
'                        Select Case iSearchCnt
'                            Case 1
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 2
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 3
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 4
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 5
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strImageOutPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 6
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strImageOutPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                        End Select
'
'                        bFileExists = fso.FileExists(strDailyScanFile)
'
'                    Loop
'
'                    If Not bFileExists Then
'                        Me.lblTrashMessage.Caption = "FASTSCAN FILE NOT IN WILTON FOR " & DateDiff("d", !ScannedDt, Now) & " DAYS" & vbNewLine & Me.lblTrashMessage.Caption
'                        Me.lblTrashMessage.ForeColor = vbRed
'                    Else
'                        Me.lblTrashMessage.Caption = "FASTSCAN File in Wilton" & vbNewLine & Me.lblTrashMessage.Caption
'                        Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
'                        Beep
'                    End If
'
'                End With
'
'            Else
'
'                StatusBar "searching... (Checking FastScan Linked images)"
'
'                If Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".TIF") _
'                    And Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".TIFF") _
'                    And Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".PDF") Then
'
'                    Me.lblTrashMessage.Caption = "Linked FASTSCAN FILE NOT IN WILTON FOR " & DateDiff("d", rsFastScan("ScannedDt"), Now) & " DAYS" & vbNewLine & Me.lblTrashMessage.Caption
'                    Me.lblTrashMessage.ForeColor = vbRed
'                    Beep
'                    GoTo Exit_sub
'                Else
'                    Me.lblTrashMessage.Caption = "Linked File in Wilton" & vbNewLine & Me.lblTrashMessage.Caption
'                    Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
'                    Beep
'                End If
'
'            End If
'
'
'
'        Else
'
'            StatusBar "searching... (NOT FastScan route) Matched"
'
'            strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
'            Set MyAdo = New clsADO
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'            Dim rsScanImages As ADODB.Recordset
'            Set rsScanImages = MyAdo.OpenRecordSet()
'
'            If Not (rsScanImages.BOF = True And rsScanImages.EOF = True) Then
'
'                With rsScanImages
'
'                    AddtlMessage = "SCANNED since " & rsScanImages("ScannedDt") & "."
'
'                    ' TKL 9/21/2011 modify for HP process
'                    If InStr(1, ".PDF/.TIF", UCase(Right(!ImageName, 4))) = 0 Then strFileExt = ".TIF"
'                    strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'
'                    iSearchCnt = 0
'                    bFileExists = fso.FileExists(strDailyScanFile)
'                    Do While bFileExists = False And iSearchCnt <= 6
'                        iSearchCnt = iSearchCnt + 1
'
'                        Select Case iSearchCnt
'                            Case 1
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 2
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strDailyScansPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 3
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 4
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 5
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                            Case 6
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !cnlyProvID & "\" & !ImageName & strFileExt
'                        End Select
'
'                        bFileExists = fso.FileExists(strDailyScanFile)
'
'                    Loop
'
'                    If Not bFileExists Then
'                        Me.lblTrashMessage.Caption = "REGULAR FILE NOT IN WILTON FOR " & DateDiff("d", !ScannedDt, Now) & " DAYS" & vbNewLine & Me.lblTrashMessage.Caption
'                        Me.lblTrashMessage.ForeColor = vbRed
'                    Else
'                        Me.lblTrashMessage.Caption = "File IS in Wilton" & vbNewLine & Me.lblTrashMessage.Caption
'                        Me.lblTrashMessage.ForeColor = RGB(0, 153, 0)
'                    End If
'
'                End With
'
'            Else
'
'                StatusBar "searching... (NOT FastScan route) Not Linked"
'
'                Me.lblTrashMessage.Caption = "REGULAR IMAGE NOT LINKED" & vbNewLine & Me.lblTrashMessage.Caption
'                Me.lblTrashMessage.ForeColor = vbRed
'                Beep
'                GoTo Exit_sub
'
'
'            End If
'
'
'
'        End If
'
'
'
'
'
'    End If
'
'
'    Me.lblTrashMessage.Caption = Me.lblTrashMessage.Caption & vbNewLine & AddtlMessage
    
Exit_Sub:

    StatusBar "searching... (Exit_Sub)"

    Set MyAdo = Nothing
    Set fso = Nothing
    Set rsClaimStopTrash = Nothing
    Set rsImageValidation = Nothing
    Set rsFastScanImageConfig = Nothing
    'Set rsFastScanImages = Nothing
    Set rsFastScan = Nothing
    'Set rsScanImages = Nothing
    Set oRs = Nothing
    
    
    StatusBar
    
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    Me.lblTrashMessage.Caption = "Error in Trash module: " & strErrMsg
    Me.lblTrashMessage.ForeColor = vbRed
    MsgBox "Error in module " & Err.Source & vbCrLf & vbCrLf & strErrMsg
    
    GoTo Exit_Sub
'End If

End Sub
'  Dim rsImageValidation As New ADODB.Recordset
'  Dim rsImageValidation1 As New ADODB.Recordset
'  Dim strSQL As String
'  Dim strsql1 As String
'
'
'
''If KeyCode = 9 Then
'    Set MyAdo = New clsADO
'    'LL 2/16/2012: Added label9 to help scanners with why the image is not trashing
'
'    strSQL = "Select * from SCANNING_Image_Log where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
'    strsql1 = "Select * from SCANNING_Image_Log_Tmp where ImageName = '" & Mid(Me.txtTrashBarCode, 6) & "'"
'
'
'    Set MyAdo = New clsADO
'    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'    MyAdo.sqlString = strSQL
'
'    'open the audit claims header and disconnect
'    Set rsImageValidation = MyAdo.OpenRecordSet()
'
'    lblImageName.Caption = Mid(Me.txtTrashBarCode, 6)
'
'        'Me.lblTrashMessage.Caption = "VALIDATED"
'        'Me.lblTrashMessage.ForeColor = vbGreen
'
'    If Not rsImageValidation.EOF Then
'        If rsImageValidation("PDFCnt") = 0 Then
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Validation Process has not been run"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        ElseIf rsImageValidation("PDFCnt") = -1 Then
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Image not Transfered or Validation Process has not been run"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        ElseIf Not IsNull(rsImageValidation("ValidationDt")) And rsImageValidation("ValidationDt") <> "1/1/1900" Then
'            Me.lblTrashMessage.Caption = "VALIDATED"
'            Me.lblTrashMessage.ForeColor = vbGreen
'            Me.Label9.Caption = ""
'            Beep
'        Else
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Check Error Report"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        End If
'    Else
'        Set myADO1 = New clsADO
'        myADO1.ConnectionString = GetConnectString("v_Data_Database")
'        myADO1.sqlString = strsql1
'        Set rsImageValidation1 = myADO1.OpenRecordSet()
'
'        If rsImageValidation1.EOF Then
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Can't find Image. New CoverPage needed"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        ElseIf rsImageValidation1("ImageType") = "INV" And rsImageValidation1("PageCnt") > 10 Then
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Bad Image Type. New CoverPage needed"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        Else
'            Me.lblTrashMessage.Caption = "Not Validated"
'            Me.lblTrashMessage.ForeColor = vbRed
'            Me.Label9.Caption = "Image not scanned in or first validation hasn't been run yet"
'            Me.Label9.ForeColor = vbBlack
'            Beep
'        End If
'    End If
'
'    'Me.txtTrashBarCode.SelStart = 0
'    'Me.txtTrashBarCode.SelLength = Len(Me.txtTrashBarCode)
'    'Me.txtTrashBarCode.SetFocus
''End If
'End Sub


Private Sub txtTrashBarCode_Enter()
    Me.txtTrashBarCode = ""
End Sub


Private Sub txtTrashBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        'Me.Label4.Caption = ""
        Me.lblTrashMessage.Caption = ""
    End If
End Sub


Sub StatusBar(Optional Msg As Variant)
Dim temp As Variant

' if the Msg variable is omitted or is empty, return the control of the status bar to Access

If Not IsMissing(Msg) Then
 If Msg <> "" Then
  temp = SysCmd(acSysCmdSetStatus, Msg)
 Else
  temp = SysCmd(acSysCmdClearStatus)
 End If
Else
  temp = SysCmd(acSysCmdClearStatus)
End If
End Sub


Function RecordValidationLog(strUserID As String, strTrashBarCode As String, intTrashMessageColor As Long, strTrashMessageText As String) As Boolean
    
On Error GoTo Err_handler

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As Integer
    Dim ErrMsg As String
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_SCANNING_Trash_Log"
    cmd.Parameters.Refresh
    cmd.Parameters("@pUserID") = strUserID
    cmd.Parameters("@pTrashBarCode") = strTrashBarCode
    cmd.Parameters("@pTrashMsgColor") = intTrashMessageColor
    cmd.Parameters("@pTrashMsgText") = strTrashMessageText
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
    If spReturnVal <> 0 Or ErrMsg <> "" Then
        MsgBox ErrMsg, vbExclamation, "Error Trash Log"
        GoTo Err_handler
    End If
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
Exit_Function:
    Exit Function
    
Err_handler:
    Me.lblTrashMessage.Caption = "Error recording trash log"
    Me.lblTrashMessage.ForeColor = vbRed
    GoTo Exit_Function
    
End Function
