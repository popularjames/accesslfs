Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrCalledFrom As String

Dim mstrMatchINPath As String
Dim mstrMatchOUTPath As String
Dim mstrSplitINPath As String
Dim mstrSplitOUTPath As String
Dim mstrTIFViewerPath As String
Dim mstrAcrobatPath As String
Dim mstrMatchedCnlyClaimNum As String

Private Sub chk2dBarCode_AfterUpdate()
    Select Case chk2dBarCode
        Case vbTrue
            txt2dBarCode.Enabled = True
            txt2dBarCode = ""
            If Nz(txtReceivedDt, "") = "" Then
                txtReceivedDt.SetFocus
            ElseIf Nz(txtImageName, "") = "" Then
                txtImageName.SetFocus
            Else
                txt2dBarCode.SetFocus
                Me.lblError.Caption = ""
            End If
            tabADR.visible = True
            ClearTabADRFields
        Case Else
            txt2dBarCode = ""
            txt2dBarCode.Enabled = False
            If Nz(txtReceivedDt, "") = "" Then
                txtReceivedDt.SetFocus
            ElseIf Nz(txtImageName, "") = "" Then
                txtImageName.SetFocus
            Else
                txtTrackingNum.SetFocus
                txtTrackingNum = ""
            End If
            ClearTabADRFields
            tabADR.visible = False
    End Select
End Sub



Private Sub chk2dBarCode_Enter()
    Me.lblError.Caption = ""
End Sub

Private Sub cmdClear_Click()

'    Me.lblError.Caption = "-"
'    Me.lblError.ForeColor = vbBlack
    Me.txtReceivedDt.Enabled = True
    Me.txtImageName.Enabled = True
    Me.txtTrackingNum.Enabled = True
    Me.cmdSave.Enabled = True
    If Me.chk2dBarCode Then Me.txt2dBarCode.Enabled = True Else Me.txt2dBarCode.Enabled = False
    
    Me.Undo
    Me.txtImageName = ""
    Me.txtTrackingNum = ""
    'Me.chk2dBarCode = vbFalse
    Me.txt2dBarCode = ""
    If ValidateReceivedDt Then
        Me.txtImageName.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()

On Error GoTo Error_Handler

    If Not ValidateReceivedDt Then
        Me.txtReceivedDt = Null
        Me.txtReceivedDt.SetFocus
        Exit Sub
    End If
    
    If Not ValidateImageName Then
        Me.txtTrackingNum = Null
        Me.txtImageName = Null
        If Me.txt2dBarCode.Enabled Then Me.txt2dBarCode = Null
        Exit Sub
    End If

    If Not ValidateTrackingNum Then
        Me.txtTrackingNum = Null
        Me.txtImageName = Null
        If Me.txt2dBarCode.Enabled Then Me.txt2dBarCode = Null
        Exit Sub
    End If

    If Me.chk2dBarCode = vbTrue And Not Validate2DBarCode Then
        Me.txtTrackingNum = Null
        Me.txtImageName = Null
        If Me.txt2dBarCode.Enabled Then Me.txt2dBarCode = Null
        Me.txtImageName.SetFocus
        Exit Sub
    End If
    
'    Dim bol2dIssue As Boolean
'    bol2dIssue = False
'    If Me.txt2dBarCode = "CNCMSVADRA##00120140220141922870###211094005858040034737783110101" Or Me.txt2dBarCode = "0000000000" Then
'        Me.txt2dBarCode = ""
'        bol2dIssue = True
'    End If


    'Save coversheet to FastScan Table
    
    Dim ErrMsgTxt As String
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    Dim ErrCode As Integer
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_ScanCoverSheet_v2"
    cmd.Parameters.Refresh
    cmd.Parameters("@pAccountID").Value = gintAccountID
    cmd.Parameters("@pReceivedDt").Value = Me.txtReceivedDt
    cmd.Parameters("@pImageName").Value = Me.txtImageName
    cmd.Parameters("@pReceivedMeth").Value = "Mail"
    cmd.Parameters("@pCarrier").Value = "Any"
    cmd.Parameters("@pTrackingNum").Value = Me.txtTrackingNum
    cmd.Parameters("@p2dBarCode").Value = Me.txt2dBarCode
    cmd.Parameters("@pUserID").Value = Identity.UserName
    
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    ErrCode = cmd.Parameters("@Return_Value")
            
    lblError.ForeColor = vbGreen
    lblError.Caption = Replace(Me.txtImageName, "02CCA", "") & vbNewLine & "Coversheet Linked SUCCESSFULLY"
    
    lblError.Caption = lblError.Caption + vbNewLine + ErrMsgTxt
    
    'warning error code not recognized
    If ErrCode <> 4 And ErrCode <> 0 Then
        GoTo Error_Handler
    End If
    
    
            
    GoTo Clean_And_Exit
    
Error_Handler:

    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    ErrCode = cmd.Parameters("@Return_Value")
    
    If ErrMsgTxt <> "" Then
        If ErrCode = 2 Then 'recently scanned
            lblError.ForeColor = vbMagenta
            lblError.Caption = Replace(Me.txtImageName, "02CCA", "") & vbNewLine & Replace(ErrMsgTxt, "FastScan.usp_FastScan_ScanCoverSheet Error: ", "")
        ElseIf ErrCode = 3 Then 'image already exists
            lblError.ForeColor = vbMagenta
            lblError.Tag = cmd.Parameters("@pTrackingNum") & " [[" & cmd.Parameters("@pCoverSheetNum") & "]]"
            lblError.Caption = Replace(Me.txtImageName, "02CCA", "") & vbNewLine & "TrackingNum (Double Click to Copy):" & vbNewLine & cmd.Parameters("@pTrackingNum") & vbNewLine & Replace(ErrMsgTxt, "FastScan.usp_FastScan_ScanCoverSheet Error: ", "")
        Else
            lblError.ForeColor = vbRed
            lblError.Caption = Replace(Me.txtImageName, "02CCA", "") & vbNewLine & "CoverSheet Scan FAILED. You MUST Try Again." & vbNewLine & Replace(ErrMsgTxt, "FastScan.usp_FastScan_ScanCoverSheet Error: ", "")
        End If
    End If
    
    If ErrMsgTxt <> "" Then
        MsgBox ErrMsgTxt, vbExclamation, "Error during Scan process. CoverSheet was NOT processed!"
    End If
    
Clean_And_Exit:
    
    
    cmdClear_Click
    Me.txtImageName.SetFocus
    
    
    
    Set myCode_ADO = Nothing
    Set cmd = Nothing

End Sub

Private Sub Form_Load()

    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs
    End If
    
    If Nz(gintAccountID, 0) = 0 Then
        MsgBox "There is not a currently selected Account ID! Cannot continue.", vbInformation, "Error: Account not selected"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    mstrMatchINPath = "" & DLookup("MatchINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrMatchOUTPath = "" & DLookup("MatchOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrSplitINPath = "" & DLookup("SplitINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrSplitOUTPath = "" & DLookup("SplitOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrTIFViewerPath = "" & DLookup("TIFViewerPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)
    mstrAcrobatPath = "" & DLookup("AcrobatPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID)

    If Nz(mstrMatchINPath, "") = "" Or Nz(mstrMatchOUTPath, "") = "" Or Nz(mstrTIFViewerPath, "") = "" Or Nz(mstrAcrobatPath, "") = "" Then
        MsgBox "One or more of the FastScan config values are not setup for this Account. Please check the FastScan_Config table.", vbInformation, "Error with FastScan config values"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchOUTPath, 1) <> "\" Then mstrMatchOUTPath = mstrMatchOUTPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct FolderName from FastScanMaint.v_CA_Scanning_FastScan_Folders where accountid = " & gintAccountID & " order by FolderName"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            MsgBox "There are no FastScan folders setup for this account!" & vbNewLine & vbNewLine & "Cannot continue.", vbInformation, "Error: FastScan folders missing"
            DoCmd.Close acForm, Me.Name
            GoTo Cleanup
            Exit Sub
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
    

    
    
    AuditName.Caption = UCase(Nz(DLookup("ClientName", "Admin_Account_config", "accountid = " & gintAccountID), "ERROR"))

    Me.txtReceivedDt.Enabled = True
    Me.txtImageName.Enabled = True
    Me.txtTrackingNum.Enabled = True
    Me.cmdSave.Enabled = True
    If Me.chk2dBarCode Then Me.txt2dBarCode.Enabled = True Else Me.txt2dBarCode.Enabled = False
    
    Me.Undo
    Me.txtImageName = ""
    Me.txtTrackingNum = ""
    Me.chk2dBarCode = vbFalse
    Me.txt2dBarCode = ""
    
    Me.chk2dBarCode.Value = False

    Me.txtReceivedDt.SetFocus
    
Cleanup:
    oAdo.DisConnect
    Set oAdo = Nothing
    Set oRs = Nothing
    
End Sub



Private Sub lblError_DblClick(Cancel As Integer)
    If lblError.Tag <> "" Then
        txtTrackingNum = lblError.Tag
    End If
End Sub







Private Sub txt2dBarCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        lblError.Caption = "-"
        lblError.ForeColor = vbBlack
    End If
End Sub

Private Sub txtReceivedDt_Enter()
    'Me.lblError.Caption = ""
    'lblError.ForeColor = vbBlack
    'lblError.Caption = "-"
    With Me.txtReceivedDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Me.txtImageName = Null
    Me.txt2dBarCode = Null
    Me.txtTrackingNum = Null
    ClearTabADRFields
End Sub
Private Sub txtImageName_Enter()
    'Me.lblError.Caption = ""
    If Not ValidateReceivedDt() Then
        Me.txtImageName = Null
        txtReceivedDt = Null
        txtReceivedDt.SetFocus
        Exit Sub
    End If
    Me.txtImageName = Null
    Me.txt2dBarCode = Null
    Me.txtTrackingNum = Null
    ClearTabADRFields
End Sub
Private Sub txt2dBarCode_Enter()
    Me.lblError.Caption = ""
    If Not ValidateImageName() Then
        Me.txtImageName.SetFocus
        Exit Sub
    End If
    Me.txt2dBarCode = Null
    Me.txtTrackingNum = Null
    ClearTabADRFields
End Sub
Private Sub txtAutoSave_Enter()
    If Not ValidateTrackingNum Then
        Me.txtTrackingNum = Null
        Me.txtTrackingNum.SetFocus
    Else
        Call cmdSave_Click
    End If
End Sub

Private Sub txtImageName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        lblError.Caption = "-"
        lblError.ForeColor = vbBlack
    End If
End Sub








Function Validate2DBarCode() As Boolean

    Validate2DBarCode = False
    mstrMatchedCnlyClaimNum = ""
    
    If Me.chk2dBarCode = vbTrue And Nz(Me.txt2dBarCode, "") = "0000000000" Then
        MakeTabADRFieldsOK
        Validate2DBarCode = True
        Exit Function
    End If
    
    If Me.chk2dBarCode = vbTrue And left(Nz(Me.txt2dBarCode, ""), 2) <> "CN" Then
        ClearTabADRFields
        lblError.ForeColor = vbRed
        lblError.Caption = "Error: The CoverSheet Scan FAILED." & vbNewLine & "2D barcode does not start with CN"
        Exit Function
    End If
    
    Dim ErrMsgTxt As String
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    Dim ErrCode As Integer
    
    ErrCode = 0
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_Validate2DBarCode"
    cmd.Parameters.Refresh
    cmd.Parameters("@pAccountID").Value = gintAccountID
    cmd.Parameters("@p2dBarCode").Value = Me.txt2dBarCode
    cmd.Parameters("@pUserID").Value = Identity.UserName
    
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    ErrCode = cmd.Parameters("@Return_Value")
    mstrMatchedCnlyClaimNum = cmd.Parameters("@pMatchedCnlyClaimNum")
    
    If ErrCode = 1 Then 'warning message
        Validate2DBarCode = True
        mstrMatchedCnlyClaimNum = ""
        MakeTabADRFieldsOK
        Exit Function
    ElseIf ErrCode = 2 Then 'real error
        mstrMatchedCnlyClaimNum = ""
        MsgBox ErrMsgTxt, vbExclamation, "FastScan ScanCoverSheet"
        Validate2DBarCode = False
    End If
   
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_DATA_Database")
        .SQLTextType = sqltext
        .sqlString = "select ICN, ClmFromDt, ClmThruDt, BeneBirthDt, PatSurname from AuditClm_Hdr where accountid = " & gintAccountID & " and CnlyClaimNum = '" & mstrMatchedCnlyClaimNum & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Validate2DBarCode = False
            mstrMatchedCnlyClaimNum = ""
            Exit Function
        Else
            oRs.MoveFirst
            Me.txtADRClmFromDt = oRs("ClmFromDt")
            Me.txtADRClmThruDt = oRs("ClmThruDt")
            Me.txtADRDOB = oRs("BeneBirthDt")
            Me.txtADRLastName = oRs("PatSurName")
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
   
    Validate2DBarCode = True
    
End Function

Function ValidateImageName() As Boolean

    ValidateImageName = False
    
    If Nz(Me.txtImageName, "") = "" Then
        lblError.ForeColor = vbRed
        lblError.Caption = "Error: The CoverSheet Scan FAILED." & vbNewLine & "ImageName cannot be empty"
        Exit Function
    End If
    
    If left(Me.txtImageName, 5) <> "02CCA" Then
        lblError.ForeColor = vbRed
        lblError.Caption = "Error: The CoverSheet Scan FAILED." & vbNewLine & "ImageName barcode does not start with 02CCA"
        Exit Function
    End If
    
    ValidateImageName = True
    
End Function

Function ValidateReceivedDt() As Boolean

    ValidateReceivedDt = False
    
    If Nz(Me.txtReceivedDt, "") = "" Then
        MsgBox "ReceivedDt cannot be empty.", vbExclamation, "Error scanning FastScan CoverSheet"
        Exit Function
    End If
    
    If IsDate(Me.txtReceivedDt) = False Then
        MsgBox "ReceivedDt must be a valid date.", vbExclamation, "Error scanning FastScan CoverSheet"
        Exit Function
    ElseIf CDate(Me.txtReceivedDt) > Date Then
        MsgBox "ReceivedDt cannot be a future date.", vbExclamation, "Error scanning FastScan CoverSheet"
        Exit Function
    ElseIf CDate(Me.txtReceivedDt) <= DateAdd("m", -1, Date) Then
        MsgBox "ReceivedDt is Over 1 MONTH AGO!" & vbNewLine & vbNewLine & "Please make sure date is correct!", vbExclamation, "Error scanning FastScan CoverSheet"
    End If

    
    ValidateReceivedDt = True
    
End Function


Function ValidateTrackingNum() As Boolean

    ValidateTrackingNum = False
    
    Me.lblError.Tag = ""
    
    Me.lblError.Caption = ""
    Me.Repaint
    
    If Len(Nz(Me.txtTrackingNum, "")) < 10 Then
        lblError.ForeColor = vbRed
        lblError.Caption = "Error: The CoverSheet Scan FAILED." & vbNewLine & "Tracking Number cannot be less than 10 characters."
        Exit Function
    End If
    
    ValidateTrackingNum = True
    
End Function


Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub txtTrackingNum_Enter()
    Me.lblError.Caption = ""
    If Not ValidateReceivedDt Then
        Me.txtReceivedDt.SetFocus
        Exit Sub
    End If
    
    
    If Not ValidateImageName() Then
        Me.txtTrackingNum = Null
        Me.txt2dBarCode = Null
        Me.txtImageName.SetFocus
        Exit Sub
    End If
    
    If Me.chk2dBarCode Then
        If Not Validate2DBarCode() Then
            Me.txt2dBarCode = Null
            Me.txtTrackingNum = Null
            txt2dBarCode.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub ClearTabADRFields()
    Me.txtADRClmFromDt = ""
    Me.txtADRClmThruDt = ""
    Me.txtADRDOB = ""
    Me.txtADRLastName = ""
End Sub

Private Sub MakeTabADRFieldsOK()
    Me.txtADRClmFromDt = "--/--/--"
    Me.txtADRClmThruDt = "--/--/--"
    Me.txtADRDOB = "--/--/--"
    Me.txtADRLastName = "-----"
End Sub
