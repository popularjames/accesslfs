Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private CDLoad As Form

Private mstrLocalHoldPath As String
Private mstrLocalPath As String
Private mstrCalledFrom As String
Private mstrLastRequestNum As String
Private mstrUserName As String
Private mstrPCName As String
Private mstrSessionID As String
Private mstrRequestNum As String
Private mstrTrackingNumber As String
Private mstrCarrier As String
Private mstrReceivedDate As String
Private mstrReceivedMethod As String

Private MyAdo As clsADO
Private myCode_ADO As clsADO

Private Enum FormType
    Master = 1
    PhillyOffice = 2
    HumanaOffice = 3
End Enum

Const CstrFrmAppID As String = "ImageLog"



Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub btnLoadMRCD_Click()

'    DoCmd.Hourglass True
'
'    Set CDLoad = New Form_frm_SCANNING_CDLoad_Main
'
'    CDLoad.txtBatchID = Me.txtBatchID
'    CDLoad.txtCnlyProvID = Me.txtCnlyProvID
'    CDLoad.visible = True
'    CDLoad.Caption = "MR CD Load"
'    CDLoad.RefreshData
'
'    DoCmd.Hourglass False
    
   
    Me.Parent.Form("tabAuto").visible = True
    Me.Parent.Form("btnBrowseCDSourceFolder").SetFocus
    Me.Parent.Form("subQuickImageLog").Enabled = False
    
End Sub

Private Sub btnMarkAll_Click()
    If Me.SCANNING_Image_Log_WorkTable.Form.RecordSet Is Nothing Then
        MsgBox "You must enter a request number first", vbInformation, "Error: No request number"
        Exit Sub
    Else
        If Me.SCANNING_Image_Log_WorkTable.Form.RecordSet.recordCount = 0 Then
            MsgBox "You must enter a request number first", vbInformation, "Error: No request number"
            Exit Sub
        End If
    End If
    MarkAllMR
End Sub

Private Sub cboReceivedMeth_Change()
    If UCase(Me.cboReceivedMeth & "") = "FAX" Then
        Me.cboCarrier = ""
        Me.cboCarrier.visible = False
    Else
        Me.cboCarrier.visible = True
    End If
End Sub

Private Sub cboReceivedMeth_Exit(Cancel As Integer)
    If UCase(Me.cboReceivedMeth & "") = "FAX" Then
        Me.cboCarrier = ""
        Me.cboCarrier.visible = False
    Else
        Me.cboCarrier.visible = True
    End If
End Sub

Private Sub cboRequestNumber_AfterUpdate()
    Me.btnLoadMRCD.visible = False
    Me.SCANNING_Image_Log_WorkTable.Form.AllowEdits = True
    Refresh_Screen
End Sub





Private Sub cmdExit_Click()
'GOOD
    DoCmd.Close
End Sub



Private Sub cmdGenerateCoverPage_Click()
    Dim bErrFlag As Boolean
    Dim bEditFlag As Boolean
    Dim iAnswer As Integer
    
    Dim strMemberName As String
    Dim strMemberDOB As String
    Dim strClmFromDt As String
    Dim strClmThruDt As String
    Dim strParams As String
    Dim strErrMsg As String
    Dim cmd As ADODB.Command
    
    mstrReceivedMethod = Me.cboReceivedMeth & ""
    mstrCarrier = Me.cboCarrier & ""
    mstrReceivedDate = Me.txtReceivedDt & ""
    mstrTrackingNumber = Me.txtTrackingNumber & ""
    
    bErrFlag = False
    bEditFlag = False
    
    If mstrReceivedMethod & "" = "" Then
        MsgBox "Please select a received method"
        bErrFlag = True
        Exit Sub
    End If
    
    If mstrCarrier & "" = "" Then
        MsgBox "Please select a carrier"
        bErrFlag = True
        Exit Sub
    End If
    
    If mstrTrackingNumber & "" = "" Then
        MsgBox "Please enter a tracking number"
        bErrFlag = True
        Exit Sub
    End If
    
    If mstrReceivedDate & "" = "" Then
        MsgBox "Please enter a received date"
        bErrFlag = True
        Exit Sub
    End If
    
    
    ' first loop to print cover pages
    bEditFlag = False
    With Me.SCANNING_Image_Log_WorkTable.Form.RecordSet
        .MoveFirst
        While Not .EOF
            If UCase(!ImportFlag) = "Y" Then
                bEditFlag = True
                If !ImageType & "" = "" Then
                    MsgBox "ERROR: SeqNo " & !SeqNo & " -- missing image type"
                    bErrFlag = True
                    Exit Sub
                End If
            End If
            .MoveNext
        Wend
    End With
    
    
    'due to WPS not being one of our payer anymore I need to log the claims and not continue with the generation of cover pages
    ' JS 10/11/2012
    
    Dim myCode_ADO As clsADO
    'Dim cmd As ADODB.Command

    ' get data
    Set myCode_ADO = New clsADO
    Set cmd = New ADODB.Command
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandText = "dbo.usp_SCANNING_CheckandLog_InactivePayer_2"
    
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters("@pUserID") = Identity.UserName()
    cmd.Parameters("@pSessionID") = Me.txtBatchID
    cmd.Parameters("@pReceivedMeth") = mstrReceivedMethod
    cmd.Parameters("@pReceivedDt") = mstrReceivedDate
    cmd.Parameters("@pCarrier") = mstrCarrier
    cmd.Parameters("@pTrackingNum") = mstrTrackingNumber
    cmd.Parameters("@pRequestNum") = cboRequestNumber
    cmd.Parameters("@pERRMsg") = ""
    
    cmd.Execute
    
    If cmd.Parameters("@pErrMsg") <> "" Then
    
        MsgBox cmd.Parameters("@pErrMsg")
        Me.cboRequestNumber = ""
        Refresh_Screen
        bErrFlag = True
        
    End If
    
    If bErrFlag Then Exit Sub
    
    If bEditFlag Then
        ' execute code to insert to log and tracking table
        Set cmd = New ADODB.Command
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
        cmd.ActiveConnection = myCode_ADO.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "usp_SCANNING_Quick_Image_Log_Process"
        cmd.Parameters.Refresh
        cmd.Parameters("@pSessionID") = mstrSessionID
        cmd.Parameters("@pReceivedMeth") = mstrReceivedMethod
        cmd.Parameters("@pReceivedDt") = mstrReceivedDate
        cmd.Parameters("@pCarrier") = mstrCarrier
        cmd.Parameters("@pTrackingNum") = mstrTrackingNumber
        cmd.Parameters("@pLocalPath") = mstrLocalPath
        
        cmd.Execute
    
        strErrMsg = cmd.Parameters("@pErrMsg")
        If strErrMsg <> "" Then
            MsgBox strErrMsg
            bErrFlag = True
        End If
    
    
        Me.SCANNING_Image_Log_WorkTable.Form.Requery
        
        If bErrFlag Then Exit Sub
    
        ' second loop to print cover pages
        
        If cboReceivedMeth <> "CD" Or Nz(CDSubForm, 0) = 0 Then
        
            With Me.SCANNING_Image_Log_WorkTable.Form.RecordSet
                .MoveFirst
                While Not .EOF
                    If UCase(!ImportFlag) = "Y" Then
                        strParams = CStr(.ScannedDt) & ";" & CStr(.ReceivedDt) & ";" & .cnlyProvID & ";" & .Icn & ";" & _
                                    .CnlyClaimNum & ";" & .ImageName & ";" & .ImageType & ";" & .MemberName & ";" & .DOB & _
                                    ";" & CStr(!ClmFromDt) & ";" & CStr(!ClmThruDt) & ";" & !UserID & ";QuickEntry" & ";" & _
                                    .RequestNum & ";" & .TrackingNum & ";" & CStr(.SeqNo) & ";" & .SessionID
                        DoCmd.OpenReport "rpt_Scanning_Cover_page", acViewNormal, , , acWindowNormal, strParams
                        'DoCmd.OpenReport "rpt_Scanning_Cover_page", acViewPreview, , , acWindowNormal, strParams
                        'MsgBox "Click OK to continue"
                        'DoCmd.Close acReport, "rpt_Scanning_Cover_page"
                        'Debug.Print strParams
                    End If
                    .MoveNext
                Wend
            End With
            
        End If
        
        Set cmd = Nothing
        myCode_ADO.DisConnect
        Set myCode_ADO = Nothing
        
        If Nz(Me.CDSubForm, 0) = 0 Then
            iAnswer = MsgBox("Process complete.  Please collect your cover sheets." & vbCrLf & vbCrLf & "Select YES if everything is correct.  Select NO to go back", vbYesNo)
            If iAnswer = vbYes Then
                Purge_Record
                Me.cboRequestNumber = ""
                Refresh_Screen
            End If
        Else
            
            Me.SCANNING_Image_Log_WorkTable.Form.AllowEdits = False
            Me.btnLoadMRCD.visible = True
            MsgBox "Process Complete, you can click on the Load MR CD button to continue.", vbInformation, "MR info was generated."
            
        End If
    End If
End Sub




Private Sub Form_Close()
' GOOD
    ' purge record for current session
    Purge_Record
    
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub Form_Load()
' GOOD
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim iAppPermission As String
    Dim iTempAccountID As String
    Dim iFormType As Integer
    
    
    On Error GoTo Err_handler
      
    Me.Caption = "Quick Image Log"
        
    Call Account_Check(Me)
    
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub

    
    strErrSource = "Quick_SCANNING_Data_Entry.Load"
    
    
    ' init variables
    mstrUserName = Identity.UserName()
    mstrPCName = GetPCName()
    
    ' set scanning office location
    ' Master = 1; Philly Office = 2; Humana office = 3
    iFormType = FormType.Master
    
    ' set folder paths
    Call ScannedOffice_AfterUpdate
    
    ' init screen display
    Me.cboRequestNumber = ""
    mstrLastRequestNum = ""
    Refresh_Screen
    
    Me.SCANNING_Image_Log_WorkTable.Form.RecordSource = "SELECT * FROM SCANNING_Quick_Image_Log_WorkTable WHERE 1=2"
    
    If Me.CDSubForm = 1 Then
        Me.cmdGenerateCoverPage.Caption = "Generate CD Reqs"
    Else
        Me.cmdGenerateCoverPage.Caption = "Generate Cover Page"
    End If
    
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs
    End If

Exit_Sub:
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then
        Err.Raise vbObjectError + 513, strErrSource, strErrMsg
    Else
        Err.Raise Err.Number, strErrSource, Err.Description
    End If
End Sub


Private Sub ScannedOffice_AfterUpdate()
    Dim iLocalAccountID As Integer
    
    Select Case ScannedOffice
        Case 1      ' Philly office
            iLocalAccountID = gintAccountID
        Case 2      ' Humana office
            iLocalAccountID = 5
    End Select
    
    mstrLocalHoldPath = "" & DLookup("LocalHoldPath", "SCANNING_Config", "AccountID = " & iLocalAccountID)
    If Right$(mstrLocalHoldPath, 1) <> "\" Then mstrLocalHoldPath = mstrLocalHoldPath & "\"

    mstrLocalPath = "" & DLookup("LocalPath", "SCANNING_Config", "AccountID = " & iLocalAccountID)
    If Right$(mstrLocalPath, 1) <> "\" Then mstrLocalPath = mstrLocalPath & "\"

    If mstrLocalHoldPath = "" Then
        MsgBox "Temporary image hold path is not defined.  Please set the hold path first", vbCritical
    End If
    
    If FolderExist(mstrLocalHoldPath) = False Then
        MsgBox "Image path '" & mstrLocalHoldPath & "' does not exists or in accessible.  Please check.", vbCritical
    End If
    
End Sub




Public Sub Refresh_Screen()
    Dim cmd As ADODB.Command
    Dim rs As ADODB.RecordSet
    Dim strErrMsg As String
    Dim iRecordSelected As Integer
    
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
    End If
    
    If Me.cboRequestNumber & "" <> "" Then
        If Me.cboRequestNumber & "" <> mstrLastRequestNum Then
        
            
        
            If Nz(Me.CDSubForm, 0) = 0 Then Me.cboReceivedMeth = ""
            
            Me.cboCarrier = ""
            Me.txtReceivedDt = ""
            Me.txtTrackingNumber = ""
            mstrLastRequestNum = Me.cboRequestNumber
            
    
            ' get data
            Set myCode_ADO = New clsADO
            Set cmd = New ADODB.Command
            myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
            
            cmd.ActiveConnection = myCode_ADO.CurrentConnection
            cmd.CommandText = "usp_SCANNING_Quick_Image_Log_WorkTable_Insert"
            
            cmd.commandType = adCmdStoredProc
            cmd.Parameters.Refresh
            cmd.Parameters("@pRequestNum") = Me.cboRequestNumber
            cmd.Parameters("@pUserID") = mstrUserName
            If mstrSessionID <> "" Then
                cmd.Parameters("@pSessionID") = mstrSessionID
            End If
            
            Set rs = cmd.Execute
            
            
            If cmd.Parameters("@pErrMsg") <> "" Then
                MsgBox cmd.Parameters("@pErrMsg")
                Me.cboRequestNumber.Value = ""
                GoTo InsertError
            Else
                iRecordSelected = cmd.Parameters("@pRecordInserted")
                mstrSessionID = cmd.Parameters("@pSessionID")
                If iRecordSelected > 0 Then
                    Me.SCANNING_Image_Log_WorkTable.Form.RecordSource = "select a.*, (b.clmstatus + ""-"" + c.clmstatusdesc) as ClmStatus from (SCANNING_Quick_Image_Log_WorkTable as a " & _
                                                                            " left join auditclm_hdr as b on a.cnlyclaimnum = b.cnlyclaimnum) " & _
                                                                            " left join xref_ClaimStatus as c on b.clmstatus = c.clmstatus where SessionID = '" & mstrSessionID & "'"
                    Me.lblProviderName1.visible = True
                    Me.lblProviderName2.visible = True
                    Me.lblProviderName2.Caption = Me.SCANNING_Image_Log_WorkTable.Form.RecordSet("ProvName")
                    
                    Me.lblLetterSentDt1.visible = True
                    Me.lblLetterSentDt2.visible = True
                    Me.lblLetterSentDt2.Caption = Me.SCANNING_Image_Log_WorkTable.Form.RecordSet("LetterReqDt")
                    
                    If Nz(Me.CDSubForm, 0) = 1 Then
                        Me.btnMarkAll.visible = True
                        Me.cboReceivedMeth.Locked = True
                        Me.lblCnlyProvID.visible = True
                        Me.txtCnlyProvID.visible = True
                        Me.txtCnlyProvID = Me.SCANNING_Image_Log_WorkTable.Form.RecordSet("CnlyProvID")
                        Me.Parent.Form("txtCnlyProvID") = Me.SCANNING_Image_Log_WorkTable.Form.RecordSet("CnlyProvID")
                    End If
                    
                    Me.lblRecordRequested1.visible = True
                    Me.lblRecordRequested2.visible = True
                    Me.lblRecordRequested2.Caption = iRecordSelected & Space(3)
                    
                End If
            End If
            
            Set cmd = Nothing
            Set myCode_ADO = Nothing
            
            Set rs = Nothing
            
        End If
        
        Me.cboReceivedMeth.visible = True
        Me.cboCarrier.visible = True
        Me.txtReceivedDt.visible = True
        Me.txtTrackingNumber.visible = True
        Me.txtBatchID = mstrSessionID
        If Nz(Me.CDSubForm, 0) = 1 Then
            Me.Parent.Form("txtbatchid") = mstrSessionID
        End If
        
    Else
    
InsertError:

        ' reset screen and variables
        mstrSessionID = ""
        txtBatchID = ""
        mstrLastRequestNum = ""
        mstrRequestNum = ""
        If Nz(Me.CDSubForm, 0) = 0 Then mstrReceivedMethod = ""
        mstrCarrier = ""
        mstrReceivedDate = ""
        mstrTrackingNumber = ""
        
        If Nz(Me.CDSubForm, 0) = 0 Then Me.cboReceivedMeth = ""
        Me.cboCarrier = ""
        Me.txtReceivedDt = ""
        Me.txtTrackingNumber = ""
        
        Me.cboReceivedMeth.visible = False
        Me.cboCarrier.visible = False
        Me.txtReceivedDt.visible = False
        Me.txtTrackingNumber.visible = False
        Me.lblProviderName1.visible = False
        Me.lblProviderName2.visible = False
        Me.lblCnlyProvID.visible = False
        Me.txtCnlyProvID.visible = False
        Me.lblLetterSentDt1.visible = False
        Me.lblLetterSentDt2.visible = False
        Me.lblRecordRequested1.visible = False
        Me.lblRecordRequested2.visible = False
        
        Me.RecordSource = ""
        Me.Requery
            
        Me.SCANNING_Image_Log_WorkTable.Form.RecordSource = "select * from SCANNING_Quick_Image_Log_WorkTable where 1=2"
        
    End If

End Sub




Private Sub txtReceivedDt_Exit(Cancel As Integer)

    If Me.txtReceivedDt.Value <> "" Then
        If IsDate(Me.txtReceivedDt.Value) Then
        
            If Me.txtReceivedDt > Date Then
                MsgBox Chr(34) & Me.txtReceivedDt.Value & Chr(34) & " is a future date." & vbCrLf & "Please enter a valid received date"
                Cancel = True
                GoTo OutOfProc
            End If
            
            If DateDiff("d", Me.txtReceivedDt, Date) > 10 Then
                MsgBox Chr(34) & Me.txtReceivedDt.Value & Chr(34) & " is older than 10 days." & vbCrLf & "Please make sure date is valid"
'                Cancel = True
'                GoTo OutOfProc
            End If
            
            
            Me.txtReceivedDt = Format(Me.txtReceivedDt, "mm-dd-yyyy")
            Me.txtReceivedDt.DefaultValue = "#" & Me.txtReceivedDt & "#"
            
        Else
            MsgBox Chr(34) & Me.txtReceivedDt.Value & Chr(34) & " is not a valid date." & vbCrLf & "Please enter a valid received date"
            Cancel = True
        End If
    End If
    
OutOfProc:

    
    
End Sub


Public Sub Purge_Record()
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.sqlString = "exec usp_SCANNING_Quick_Image_Log_WorkTable_Purge '" & mstrSessionID & "'"
    myCode_ADO.Execute
    Set myCode_ADO = Nothing
    
    mstrSessionID = ""
End Sub


Sub MarkAllMR()
    Dim rs As DAO.RecordSet
    Set rs = Me.SCANNING_Image_Log_WorkTable.Form.RecordSet
    rs.MoveFirst
    While Not rs.EOF
        rs.Edit
            rs("ImportFlag") = "Y"
            rs("ImageType") = "MR"
        rs.Update
        rs.MoveNext
    Wend
End Sub
