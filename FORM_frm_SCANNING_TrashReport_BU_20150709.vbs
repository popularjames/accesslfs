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
    Label5.Caption = ""
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
End Sub

Private Sub Text3_AfterUpdate()

    Dim rsImageValidation As New ADODB.RecordSet
    Dim rsClaimStopTrash As New ADODB.RecordSet
    Dim oRs As New ADODB.RecordSet
    Dim strSQL As String
    Dim bFileExists As Boolean

On Error GoTo Err_handler
 
    Me.Label5.Caption = "searching..."
    Me.Label5.ForeColor = vbBlack
    Dim AddtlMessage As String
    Dim strValidatedPath As String
    Dim strDailyScansPath As String
    Dim strImageOutPath As String
    Dim strErrMsg As String
    Dim strFileExt As String
    Dim strDailyScanFile As String
    Dim iSearchCnt As Integer
    
    Dim oAdo As clsADO
    
    StatusBar "searching... (check StopTrash table, Stop Trash)"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select top 1 CAMessage from SCANNING_Image_Log il " & _
                        " join SCANNING_Claim_StopTrash st on st.cnlyclaimnum = il.cnlyclaimnum " & _
                        " Where il.imagetype ='MR' and il.ImageName Like '" & Mid(Me.Text3, 6) & "%'"
        
        Set rsClaimStopTrash = oAdo.OpenRecordSet()
        If Not rsClaimStopTrash Is Nothing Then
            If Not rsClaimStopTrash.BOF And Not rsClaimStopTrash.EOF Then
                rsClaimStopTrash.MoveFirst
                Me.Label5.Caption = rsClaimStopTrash("CAMessage")
                Me.Label5.ForeColor = vbRed
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
                        " Where il.ImageName Like '" & Mid(Me.Text3, 6) & "%'"
        
        Set oRs = .ExecuteRS
        If .GotData = True Then
            Me.Label5.Caption = "IMAGE WAS LOST! MUST be sent back to be rescanned."
            Me.Label5.ForeColor = vbRed
            Beep
            GoTo Exit_Sub
        End If
    End With
 
    StatusBar "searching... (check SCANNING_Image_Log table)"
 
'If KeyCode = 9 Then
    Set MyAdo = New clsADO
    'MsgBox Me.Text3
    strSQL = " select * from SCANNING_Image_Log where ImageName = '" & Mid(Me.Text3, 6) & "'" '& " and isnull(ValidationDt,'1/1/1900') <> '1/1/1900'"
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    MyAdo.sqlString = strSQL
    
    'open the audit claims header and disconnect
    Set rsImageValidation = MyAdo.OpenRecordSet()
    
    lblImageName.Caption = Mid(Me.Text3, 6)
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Me.Label5.Caption = "Not Validated"
    Me.Label5.ForeColor = vbRed
    
    'This is for image that are in Scanning_Image_log, that means it should be on its way to be attached to a claim already, we just need to check the validation stage/status
    If Not rsImageValidation.EOF Then
        
        If rsImageValidation("PDFCnt") = 0 And gintAccountID <> 1 Then 'this is for MCR only
            Me.Label5.Caption = "Not Validated" & vbNewLine & "Validation Process has not been run" & vbNewLine & "SCANNED since " & rsImageValidation("ScannedDt") & "."
            Me.Label5.ForeColor = vbRed
            Beep
        ElseIf rsImageValidation("PDFCnt") = -1 And gintAccountID <> 1 Then 'this is for MCR only
            Me.Label5.Caption = "Not Validated" & vbNewLine & "Image not Transfered or Validation Process has not been run" & vbNewLine & "SCANNED since " & rsImageValidation("ScannedDt") & "."
            Me.Label5.ForeColor = vbRed
            Beep
        ElseIf Not IsNull(rsImageValidation("ValidationDt")) And rsImageValidation("ValidationDt") <> "1/1/1900" Then
            Me.Label5.Caption = "VALIDATED"
            Me.Label5.ForeColor = RGB(0, 153, 0)
            Beep
            GoTo Exit_Sub
        End If

    End If
        
    'continues the analysis, we might find more data or add more data
    Dim bScanned As Boolean
    bScanned = False
    
    'can we find it in scanning_image_log_tmp?, if so it might just be waiting there or stuck with an error
    strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.Text3, 6) & "'"
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    MyAdo.sqlString = strSQL
    Dim rsImageLogTmp As ADODB.RecordSet
    Set rsImageLogTmp = MyAdo.OpenRecordSet()
            
    If Not rsImageLogTmp.EOF Then
        Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "SCANNED since " & rsImageLogTmp("ScannedDt") & "."
        bScanned = True
        If Nz(rsImageLogTmp("ErrMsg"), "") <> "" Then
            Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "Is NOT validating because: " & rsImageLogTmp("ErrMsg")
        End If
    End If
    
    'if it is a fastscan image and has not been matched yet, or it is a no match / finish just sitting there
    If InStr(1, Mid(Me.Text3, 4), "FAST") Then
    
        StatusBar "searching... (FastScan route)"

        Set MyAdo = New clsADO
        strSQL = " select * from SCANNING_fastscan_log where ImageName = '" & Mid(Me.Text3, 6) & "'"
        MyAdo.ConnectionString = GetConnectString("v_Data_Database")
        MyAdo.sqlString = strSQL
        
        Dim rsFastScan As ADODB.RecordSet
        Set rsFastScan = MyAdo.OpenRecordSet()
        
        'has not even been linked yet
        If rsFastScan.EOF Then
            Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "FastScan Image NOT LINKED"
            Me.Label5.ForeColor = vbRed
            Beep
            GoTo Exit_Sub
        End If
        
        'deleted fastscan image, let it go
        If left(rsFastScan("ProcStatusCd"), 7) = "DELETED" Then
            Me.Label5.Caption = "VALIDATED - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
            Me.Label5.ForeColor = RGB(0, 153, 0)
            Beep
            GoTo Exit_Sub
        End If
        
        If Not bScanned Then
            Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "SCANNED since " & rsFastScan("ScannedDt") & "."
        End If
        
        Dim rsFastScanImageConfig As ADODB.RecordSet
        Set rsFastScanImageConfig = MyAdo.OpenRecordSet("select * from SCANNING_FastScan_Config where AccountID = " & rsFastScan("AccountID"))
        If rsFastScanImageConfig.EOF = True Then
            MsgBox "FastScan configuration for account " & gstrAcctDesc & " has not been set up.  Please set it up and re-try", vbCritical
            GoTo Exit_Sub
        End If

        Dim strFastScanMatchInPath As String

        strFastScanMatchInPath = rsFastScanImageConfig("MatchInPath") & ""
        If Right$(strFastScanMatchInPath, 1) <> "\" Then strFastScanMatchInPath = strFastScanMatchInPath & "\"
        
        Dim strFastScanMatchInFile As String
        
        strFileExt = Nz(rsFastScan("FileExt"), "")
        
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
        
        
        'other fastscan statuses that are not Matched, they should be sitting in the fastscan folder
        If rsFastScan("ProcStatusCd") <> "MATCHED" Then
            Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "FastScan status is --> " & rsFastScan("ProcStatusCd") & " since: " & rsFastScan("ProcStatusLastUpDt")
            Me.Label5.ForeColor = vbRed
            
            If bFileExists Then
                Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "Image IS in Wilton"
                Me.Label5.ForeColor = RGB(0, 153, 0)
            Else
                Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "Image IS NOT in Wilton for " & DateDiff("d", rsFastScan("ScannedDt"), Now) & " days"
                Me.Label5.ForeColor = vbRed
            End If
            
            Beep
            GoTo Exit_Sub
        End If
        
        
    End If
        
    
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
    MsgBox "Error in module " & Err.Source & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
'End If

'    Dim rsImageValidation As New ADODB.RecordSet
'    Dim rsClaimStopTrash As New ADODB.RecordSet
'    Dim oRS As New ADODB.RecordSet
'    Dim strSQL As String
'    Dim bFileExists As Boolean
'
'
'    Me.Label5.Caption = "searching..."
'    Me.Label5.ForeColor = vbBlack
'    Dim AddtlMessage As String
'    Dim strValidatedPath As String
'    Dim strDailyScansPath As String
'    Dim strImageOutPath As String
'    Dim strErrMsg As String
'    Dim strFileExt As String
'    Dim strDailyScanFile As String
'    Dim iSearchCnt As Integer
'
'    Dim oAdo As clsADO
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_Data_Database")
'        .SQLTextType = SQLText
'        .sqlString = "select top 1 CAMessage from SCANNING_Image_Log il " & _
'                        " join SCANNING_Claim_StopTrash st on st.cnlyclaimnum = il.cnlyclaimnum " & _
'                        " Where il.imagetype ='MR' and il.ImageName Like '%" & Mid(Me.Text3, 6) & "%'"
'
'        Set rsClaimStopTrash = oAdo.OpenRecordSet()
'        If Not rsClaimStopTrash.BOF And Not rsClaimStopTrash.EOF Then
'            rsClaimStopTrash.MoveFirst
'            Me.Label5.Caption = rsClaimStopTrash("CAMessage")
'            Me.Label5.ForeColor = vbRed
'            Beep
'            GoTo Exit_sub
'        End If
'    End With
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_Data_Database")
'        .SQLTextType = SQLText
'        .sqlString = "select 1 from SCANNING_Image_Log il " & _
'                        " join SCANNING_ImageName_StopTrash st on st.ImageName = replace(replace(il.imagename,'.PDF',''),'.TIF','') " & _
'                        " Where il.ImageName Like '%" & Mid(Me.Text3, 6) & "%'"
'
'        Set oRS = .ExecuteRS
'        If .GotData = True Then
'            Me.Label5.Caption = "IMAGE WAS LOST! MUST be sent back to be rescanned."
'            Me.Label5.ForeColor = vbRed
'            Beep
'            GoTo Exit_sub
'        End If
'    End With
'
''If KeyCode = 9 Then
'    Set MyAdo = New clsADO
'    'MsgBox Me.Text3
'    strSQL = " select * from SCANNING_Image_Log where ImageName = '" & Mid(Me.Text3, 6) & "'" '& " and isnull(ValidationDt,'1/1/1900') <> '1/1/1900'"
'    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'    MyAdo.sqlString = strSQL
'
'    'open the audit claims header and disconnect
'    Set rsImageValidation = MyAdo.OpenRecordSet()
'
'    lblImageName.Caption = Mid(Me.Text3, 6)
'
'    Dim fso As FileSystemObject
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    If Not rsImageValidation.EOF Then
'
'        If rsImageValidation("PDFCnt") = 0 And gintAccountID <> 1 Then
'            Me.Label5.Caption = "Not Validated" & vbNewLine & "Validation Process has not been run"
'            Me.Label5.ForeColor = vbRed
'            Beep
'        ElseIf rsImageValidation("PDFCnt") = -1 And gintAccountID <> 1 Then
'            Me.Label5.Caption = "Not Validated" & vbNewLine & "Image not Transfered or Validation Process has not been run"
'            Me.Label5.ForeColor = vbRed
'            Beep
'        ElseIf Not IsNull(rsImageValidation("ValidationDt")) And rsImageValidation("ValidationDt") <> "1/1/1900" Then
'            Me.Label5.Caption = "VALIDATED"
'            Me.Label5.ForeColor = RGB(0, 153, 0)
'            Beep
'        Else
'            Me.Label5.Caption = "Not Validated" & vbNewLine & "Check Error Report"
'            Me.Label5.ForeColor = vbRed
'            Beep
'        End If
'
'
'    Else
'
'        AddtlMessage = ""
'
'        Dim rsImageConfig As ADODB.RecordSet
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
'        If fso.FolderExists(strValidatedPath) = False Then
'            strErrMsg = "Temporary image hold path '" & strValidatedPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strDailyScansPath) = False Then
'            strErrMsg = "Temporary image hold path '" & strValidatedPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strImageOutPath) = False Then
'            strErrMsg = "Remote image path '" & strImageOutPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'        If fso.FolderExists(strFastScanMatchInPath) = False Then
'            strErrMsg = "FastScan Match In Path '" & strImageOutPath & "' does not exists or in accessible.  Please check."
'            GoTo Err_handler
'        End If
'
'
'        If Right$(strValidatedPath, 1) <> "\" Then strValidatedPath = strValidatedPath & "\"
'        If Right$(strDailyScansPath, 1) <> "\" Then strDailyScansPath = strDailyScansPath & "\"
'        If Right$(strImageOutPath, 1) <> "\" Then strImageOutPath = strImageOutPath & "\"
'        If Right$(strFastScanMatchInPath, 1) <> "\" Then strFastScanMatchInPath = strFastScanMatchInPath & "\"
'
'        Me.Label5.Caption = "NOT VALIDATED"
'        'Me.Label5.ForeColor = vbRed
'        Beep
'
'        If InStr(1, Mid(Me.Text3, 4), "FAST") Then
'
'            Set MyAdo = New clsADO
'            strSQL = " select * from SCANNING_fastscan_log where ImageName = '" & Mid(Me.Text3, 6) & "'"
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'
'            Dim rsFastScan As ADODB.RecordSet
'            Set rsFastScan = MyAdo.OpenRecordSet()
'
'            If rsFastScan.EOF Then
'                Me.Label5.Caption = Me.Label5.Caption + " - FASTSCAN IMAGE NOT LINKED"
'                Me.Label5.ForeColor = vbRed
'                Beep
'                GoTo Exit_sub
'            End If
'
'            If left(rsFastScan("ProcStatusCd"), 7) = "DELETED" Then
'                Me.Label5.Caption = "VALIDATED - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
'                Me.Label5.ForeColor = RGB(0, 153, 0)
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
'            If (left(rsFastScan("ProcStatusCd"), 7) = "FINISH" Or left(rsFastScan("ProcStatusCd"), 7) = "REJECT" Or left(rsFastScan("ProcStatusCd"), 7) = "NOMATCH") Then
'                Me.Label5.Caption = "Not Validated - FASTCAN IMAGE STATUS IS --> " & rsFastScan("ProcStatusCd")
'                Me.Label5.ForeColor = vbRed
'
'                If bFileExists Then
'                    Me.Label5.Caption = Me.Label5.Caption & vbNewLine & "File IS in Wilton"
'                    Me.Label5.ForeColor = RGB(0, 153, 0)
'                End If
'
'                Beep
'                GoTo Exit_sub
'            End If
'
'            Me.Label5.Caption = rsFastScan("procstatuscd") & " since " & rsFastScan("scanneddt") & "." & vbNewLine & Me.Label5.Caption
'
'            strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.Text3, 6) & "'"
'            Set MyAdo = New clsADO
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'            Dim rsFastScanImages As ADODB.RecordSet
'            Set rsFastScanImages = MyAdo.OpenRecordSet()
'
'            If Not (rsFastScanImages.BOF = True And rsFastScanImages.EOF = True) Then
'
'                rsFastScanImages.MoveFirst
'
'                With rsFastScanImages
'                    ' TKL 9/21/2011 modify for HP process
'                    If InStr(1, ".PDF/.TIF", UCase(Right(!ImageName, 4))) = 0 Then strFileExt = ".TIF"
'                    strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
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
'                                strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 2
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 3
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 4
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 5
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 6
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                        End Select
'
'                        bFileExists = fso.FileExists(strDailyScanFile)
'
'                    Loop
'
'                    If Not bFileExists Then
'                        Me.Label5.Caption = "FASTSCAN FILE NOT IN WILTON FOR " & DateDiff("d", !ScannedDt, now) & " DAYS" & vbNewLine & Me.Label5.Caption
'                        Me.Label5.ForeColor = vbRed
'                    Else
'                        Me.Label5.Caption = "FASTSCAN File in Wilton" & vbNewLine & Me.Label5.Caption
'                        Me.Label5.ForeColor = RGB(0, 153, 0)
'                        Beep
'                    End If
'
'                End With
'
'            Else
'
'                If Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".TIF") _
'                    And Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".TIFF") _
'                    And Not fso.FileExists(strDailyScansPath & rsFastScan("ProviderFolder") & "\" & rsFastScan("ImageName") & ".PDF") Then
'
'                    Me.Label5.Caption = "FASTSCAN FILE NOT IN WILTON FOR " & DateDiff("d", rsFastScan("ScannedDt"), now) & " DAYS" & vbNewLine & Me.Label5.Caption
'                    Me.Label5.ForeColor = vbRed
'                    Beep
'                    GoTo Exit_sub
'                Else
'                    Me.Label5.Caption = "File in Wilton" & vbNewLine & Me.Label5.Caption
'                    Me.Label5.ForeColor = RGB(0, 153, 0)
'                    Beep
'                End If
'
'            End If
'
'
'
'        Else
'
'            strSQL = " select * from SCANNING_Image_Log_tmp where ImageName = '" & Mid(Me.Text3, 6) & "'"
'            Set MyAdo = New clsADO
'            MyAdo.ConnectionString = GetConnectString("v_Data_Database")
'            MyAdo.sqlString = strSQL
'            Dim rsScanImages As ADODB.RecordSet
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
'                    strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
'
'                    iSearchCnt = 0
'                    bFileExists = fso.FileExists(strDailyScanFile)
'                    Do While bFileExists = False And iSearchCnt <= 6
'                        iSearchCnt = iSearchCnt + 1
'
'                        Select Case iSearchCnt
'                            Case 1
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 2
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strDailyScansPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 3
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 4
'                                ' check validation directory. Files may have been moved there from previous run
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 5
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".TIF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                            Case 6
'                                ' check image out directory. Files may have been moved there by SuperFlex
'                                If strFileExt <> "" Then strFileExt = ".PDF"
'                                strDailyScanFile = strValidatedPath & !CnlyProvID & "\" & !ImageName & strFileExt
'                        End Select
'
'                        bFileExists = fso.FileExists(strDailyScanFile)
'
'                    Loop
'
'                    If Not bFileExists Then
'                        Me.Label5.Caption = "REGULAR FILE NOT IN WILTON FOR " & DateDiff("d", !ScannedDt, now) & " DAYS" & vbNewLine & Me.Label5.Caption
'                        Me.Label5.ForeColor = vbRed
'                    Else
'                        Me.Label5.Caption = "File IS in Wilton" & vbNewLine & Me.Label5.Caption
'                        Me.Label5.ForeColor = RGB(0, 153, 0)
'                    End If
'
'                End With
'
'            Else
'
'
'                Me.Label5.Caption = "REGULAR IMAGE NOT LINKED" & vbNewLine & Me.Label5.Caption
'                Me.Label5.ForeColor = vbRed
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
'    Me.Label5.Caption = Me.Label5.Caption & vbNewLine & AddtlMessage
'
'Exit_sub:
'    Exit Sub
'
'Err_handler:
'
'    MsgBox "Error in validation process!"
'    'Me.Text3.SelStart = 0
'    'Me.Text3.SelLength = Len(Me.Text3)
'    'Me.Text3.SetFocus
''End If

End Sub

Private Sub Text3_Enter()
    Me.Text3 = ""
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        Me.Label4.Caption = ""
        Me.Label5.Caption = ""
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
