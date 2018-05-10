Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim MyFastScan As clsFastScan
Dim mstrMatchINPath As String
Dim mstrMatchOUTPath As String
Dim mstrSplitINPath As String
Dim mstrSplitOUTPath As String
Dim mstrTIFViewerPath As String
Dim mstrAcrobatPath As String
Dim mstrLocalPath As String
Dim mstrClaimAttachPath As String
Dim mstrProvAttachPath As String


Dim mstrCurrentUser As String
Dim mstrSplitFlag As String
Dim mstrDataEntry_ICN_Flag As String
Dim mstrDataEntry_PayerName_Flag As String
Dim mstrDataEntry_CnlyProvID_Flag As String
Dim mstrCalledFrom As String
Dim mbolRejectUser As Boolean
Dim mbolFuzzyMatchUser As Boolean
Dim mbolPriorityUser As Boolean
Dim mstrOnlyThisCoverSheet As String
Dim mstrImageFileExt As String
Dim mstrImageFullName As String
Dim mstrDefaultImageType As String
Dim mstrRequiredImageType As String
Dim mstrReasonType As String
Dim mintSessionID As Long
Dim mintpageCnt As Integer
Dim mstrSaveSplit As String
Dim mstrImageViewer As String
Const CstrFrmAppID As String = "FastScanMain"
Private miAppPermission As Integer
Public mbAllowChange As Boolean
Public mbAllowAdd As Boolean
Public mbAllowView  As Boolean
Public mbAllowDelete As Boolean

Private WithEvents frmScanningFastScanSplit As Form_frm_FastScan_Split
Attribute frmScanningFastScanSplit.VB_VarHelpID = -1
Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1


Public Property Get SaveSplit() As String
    SaveSplit = mstrSaveSplit
End Property
Property Let SaveSplit(data As String)
     mstrSaveSplit = data
End Property

Private Sub cmdClearAllCriteria_Click()
    Me.txtCnlyProvID = ""
    Me.txtICN = ""
    Me.cmbPayerName = ""
    Me.txtPatFirstInit = ""
    Me.txtPatLastName = ""
    Me.txtclmFromDt = ""
    Me.txtPatDOB = ""
    Me.txtInstanceID = ""
    Me.txtPatCAN = ""
    Me.txtPatBIC = ""
    Me.txtMRNumber = ""
    Me.txtALJNumber = ""
    Me.txtQICNumber = ""
    Me.cmdPropagateAddAllReq.Tag = 0
End Sub

Private Sub cmdClearAllResults_Click()

    'Clear all main and related selections
    Dim sqlUpdate As String
    sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = false WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType in ('B','M','R')"
    CurrentDb.Execute (sqlUpdate)
    Me.Refresh


    Me.subfrm_Results.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where 1=2"
    
    If Me.TogRelated.Value = -1 Then
        Me.subfrm_Results_Other.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where 1=2"
    End If
    
    
    If Me.TogShowBarCodeResults.Enabled Then
        Me.TogShowBarCodeResults.Value = 0
        TogShowBarCodeResults.Locked = False
    End If
    
    DecisionSwitch
  
End Sub

Private Sub cmdNoMatch_Click()

Dim ErrMsgTxt As String

On Error GoTo Error_Handler

    If Me.cmdNoMatch.Caption = "SPLIT" Then
        Call ProcessNoMatchSplit
    Else
        Call ProcessNoMatchReject(Me.cmdNoMatch.Caption)
    End If
    
    If mstrOnlyThisCoverSheet <> "" And MyFastScan.CoverSheetNum = "NA" Then
        DoCmd.Close acForm, Me.Name
    End If
    
Exit Sub
    
Error_Handler:

If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
MsgBox ErrMsgTxt, vbExclamation, "Error during NoMatch / Split process"

    
End Sub

Private Sub ProcessNoMatchSplit()
    
Dim ErrMsgTxt As String
    
    
    'Make sure an ImageType has been selected
'    If Me.lstImageType.ListIndex = -1 Then
'        ErrMsgTxt = "You must select an Image Type before Split."
'        GoTo Validation_Error
'    End If
    
    If mintpageCnt = 0 Then
    
        Call GetImagePageCnt
    
        If mintpageCnt < 1 Then
            ErrMsgTxt = "Page count for image file is less than 1. Cannot Split"
            GoTo Error_Handler
        End If

    End If
    
    If Nz(mstrSplitINPath, "") = "" Or Nz(mstrSplitOUTPath, "") = "" Then
        MsgBox "One or more of the FastScan Split paths are not configured for this Account. Please check the FastScan_Config table.", vbInformation, "Error with FastScan Split folder paths"
        Exit Sub
    End If
    
    Dim strSplitPath As String
    strSplitPath = ""
    
    If mstrImageFileExt = "TIF" Then
        strSplitPath = mstrSplitINPath
    ElseIf mstrImageFileExt = "PDF" Then
        strSplitPath = mstrSplitOUTPath
    End If
    

    
    mstrSaveSplit = "N"

    If frmScanningFastScanSplit Is Nothing Then
        Set frmScanningFastScanSplit = New Form_frm_FastScan_Split
        If SysCmd(acSysCmdGetObjectState, acForm, "frm_FastScan_Split") Then
            ColObjectInstances.Add frmScanningFastScanSplit.hwnd & ""
            frmScanningFastScanSplit.txtCoverSheetNum = MyFastScan.CoverSheetNum
            frmScanningFastScanSplit.TxtFileName = MyFastScan.rsCoverSheet("ImageName")
            frmScanningFastScanSplit.txtPageCnt = mintpageCnt
            frmScanningFastScanSplit.SplitPath = strSplitPath
            frmScanningFastScanSplit.FileExt = mstrImageFileExt
            frmScanningFastScanSplit.ImageFullName = mstrImageFullName
            frmScanningFastScanSplit.ReasonCd = Me.lstNoMatchReason
            frmScanningFastScanSplit.RefreshScreen
            ShowFormAndWait frmScanningFastScanSplit
        End If
        Set frmScanningFastScanSplit = Nothing
    Else
        frmScanningFastScanSplit.SetFocus
    End If
    
    If mstrSaveSplit = "N" Then
        ErrMsgTxt = "User cancelled Split"
        GoTo Validation_Error
    End If
    
    GoTo Clean_And_Exit

Validation_Error:
    DoCmd.Hourglass False
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "Split validation error"
    Exit Sub
    
Error_Handler:
    DoCmd.Hourglass False
    If ErrMsgTxt <> "" Then
        MsgBox ErrMsgTxt, vbExclamation, "Error during Split process."
    End If
    
Clean_And_Exit:

    If Not IsNull(MyFastScan.CoverSheetNum) Then Call FastScan_UnLock(MyFastScan.CoverSheetNum)
    
    StartAllOver
    
    DoCmd.Hourglass False

End Sub

Private Sub cmdClearNoMatch_Click()
    Me.lstNoMatchReason = ""
    lstNoMatchReason_Change
End Sub

Private Sub cmdMatch_Click()

On Error GoTo Error_Handler

    Dim ErrMsgTxt As String
    Dim ErrLevel As Integer
    Dim UserAnswer As Integer
    
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    Dim rs As ADODB.RecordSet
    
    'Dim strCoverSheet As String
    'Dim strCnlyClaimNum As String
    Dim strCnlyProvID As String
    Dim strSQL As String

    Dim strImageFolderDest As String

    'Dim strImageFileNameDest As String
    'Dim strReceivedDt As String
    Dim strImageType As String
    Dim strAttachPath As String
    

    
    strCnlyProvID = ""
    strImageFolderDest = ""
    strImageType = Me.lstImageType
    

    DoCmd.Hourglass True

    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

'    Make sure the whole row has been selected.
'    If Not Me.subfrm_Results.Form.WholeRowSelected Then
'        ErrMsgTxt = "You must select one whole row in order to proceed."
'        GoTo Error_Handler
'    End If

    'Make sure a match reason has been selected
    If Me.lstNoMatchReason.ListIndex = -1 Then
        ErrMsgTxt = "You must select a Match reason first."
        GoTo Validation_Error
    End If
    
    'validate that a the correct Image Type was selected for the NoMatch / Match reason if it is enforced
    If mstrRequiredImageType = "Y" And lstImageType <> mstrDefaultImageType Then
        ErrMsgTxt = "The selected reason only works for Image Type = " & mstrDefaultImageType & " , please fix your selections and try again."
        GoTo Validation_Error
    End If

    'Make sure one claim has been selected from the results
    If Me.TogPropagate.Value = 0 And ResultsSelected() = 0 Then
        ErrMsgTxt = "A claim has not been selected from the results. Cannot Match"
        GoTo Validation_Error
    End If
    
    'If working with propagate there should not be claims marked from the main results
    If ResultsSelected() > 0 And Me.TogPropagate.Value = -1 Then
        ErrMsgTxt = "When you work with Propagate you cannot have claims with a check in the search results, only claims added to the Propagate list will be matched."
        GoTo Validation_Error
    End If
    
    'If working with propagate there should be at least one claim in the propagate list
    If ResultsOtherSelected("P") = 0 And Me.TogPropagate.Value = -1 Then
        ErrMsgTxt = "When you work with Propagate you should have at least one claim in the Propagate list."
        GoTo Validation_Error
    End If
   
    'Make sure an ImageType has been selected
    If Me.lstImageType.ListIndex = -1 Then
        ErrMsgTxt = "You must select an Image Type before matching."
        GoTo Validation_Error
    End If
       
    'Make sure we have a coversheet and claim to match
    If MyFastScan.CoverSheetNum = "" Then
        ErrMsgTxt = "CoverSheet number is missing. Cannot Match."
        GoTo Validation_Error
    End If
    
   
    If Not CloseAllViewerWindows(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\"))) Then
        ErrMsgTxt = "Cannot automatically close the viewer!, please close the Image window manually and try again."
        GoTo Validation_Error
    End If
    
    'Make sure the imagefile still exists
    If Not FileExists(mstrImageFullName) Then
        ErrMsgTxt = "Image File " & MyFastScan.rsCoverSheet("ImageName") & ".* does not exist anymore. Cannot Match"
        GoTo Validation_Error
    End If
    
    'do some validation on the SQL side
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_ProcessMatch_Validate"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheetNum").Value = MyFastScan.CoverSheetNum
    cmd.Parameters("@pReasonType").Value = mstrReasonType
    cmd.Parameters("@pImageType").Value = strImageType
    cmd.Parameters("@pUserID").Value = mstrCurrentUser
    cmd.Parameters("@pSessionID").Value = mintSessionID
    
    cmd.Execute

    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    ErrLevel = Trim(cmd.Parameters("@pErrorLevel").Value) & ""
    If ErrMsgTxt <> "" Then
        If ErrLevel = 1 Then 'if error we can't continue
            GoTo Error_Handler
        ElseIf ErrLevel = 0 Then 'if warning ask the question
            UserAnswer = MsgBox(ErrMsgTxt, vbQuestion + vbYesNo + vbDefaultButton2, "Match Validation")
            If UserAnswer <> vbYes Then
                ErrMsgTxt = "Match was cancelled by user."
                GoTo Validation_Error
            End If
        ElseIf ErrLevel = 2 Then 'warning message
            GoTo Validation_Error
        Else
            ErrMsgTxt = "Invalid ErrorLevel returned by usp_FastScan_ProcessMatch_Validate"
            GoTo Error_Handler
        End If
    End If
        

    If mintpageCnt = 0 Then
    
        Call GetImagePageCnt
    
        If mintpageCnt < 1 Then
            ErrMsgTxt = "Page count for image file is less than 1. Cannot Match"
            GoTo Error_Handler
        End If

    End If

    Dim MyAdo As clsADO
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    strSQL = "SELECT DISTINCT CnlyProvID FROM FastScan.v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ProcessInd = 1"
    Set rs = MyAdo.OpenRecordSet(strSQL)
    
    rs.MoveFirst
    While Not rs.EOF
        With rs
        
            strCnlyProvID = UCase(Replace(!cnlyProvID, "_", "-"))
            strCnlyProvID = Replace(Trim(strCnlyProvID), " ", "")
            
            
            If mstrReasonType = "MA" Then
                strImageFolderDest = mstrMatchOUTPath & strCnlyProvID & "\"
                strAttachPath = ""
            ElseIf mstrReasonType = "ATC" Then
                strImageFolderDest = mstrClaimAttachPath & "CnlyClaimNum\FastScan\"
            ElseIf mstrReasonType = "ATP" Then
                strImageFolderDest = mstrProvAttachPath & "CnlyProvID\" & strCnlyProvID & "\"
            End If
            
            'Make sure destination folder exists
            If Not FolderExists(strImageFolderDest) Then
                If Not CreateFolder(strImageFolderDest) Then
                    ErrMsgTxt = "Provider Image Folder " & strImageFolderDest & " does not exist. Cannot Match"
                    GoTo Validation_Error
                End If
            End If
            
            'lets copy the file to its destination before starting the transaction
            'this one has the extension that we care about
            If Not CopyFile(mstrImageFullName, strImageFolderDest & MyFastScan.rsCoverSheet("ImageName") & "." & mstrImageFileExt, False) Then
                ErrMsgTxt = "Could not move image file to its destination" & strImageFolderDest & MyFastScan.rsCoverSheet("ImageName") & ". Cannot Match"
                GoTo Error_Handler
            End If
            
            .MoveNext
            
        End With
    Wend

    
    'Start transaction
    myCode_ADO.BeginTrans
    
    
    'create row in scanning_image_log_tmp table and mark image as matched in FastScan table
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_ProcessMatch_v5"
    cmd.Parameters.Refresh
    'cmd.Parameters("@pCoverSheetNum").Value = strCoverSheet
    cmd.Parameters("@pReasonType").Value = mstrReasonType
    cmd.Parameters("@pReasonCd").Value = Me.lstNoMatchReason
    cmd.Parameters("@pImageType").Value = strImageType
    cmd.Parameters("@pFileExt").Value = mstrImageFileExt
    cmd.Parameters("@pLocalPath").Value = mstrLocalPath
    cmd.Parameters("@pAttachPath").Value = strImageFolderDest
    cmd.Parameters("@pPageCnt").Value = mintpageCnt
    'cmd.Parameters("@pAccountID").Value = gintAccountID
    'cmd.Parameters("@pAuditNum").Value = mintCurrentAuditNum
    cmd.Parameters("@pUserID").Value = mstrCurrentUser
    cmd.Parameters("@pSessionID").Value = mintSessionID
    
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    If ErrMsgTxt <> "" Then
        GoTo Error_Handler
    End If
                            
                                                            
    'delete image from FastScan folder
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile (mstrImageFullName)
    
    If FileExists(mstrImageFullName) Then
        ErrMsgTxt = "Could not delete original image file " & mstrImageFullName & ". Cannot Match"
        GoTo Error_Handler
    End If
    
    'commit transaction
    myCode_ADO.CommitTrans
    
    'here we copy all file formats for the same image
    If Not CopyFile(mstrImageFullName, strImageFolderDest & MyFastScan.rsCoverSheet("ImageName") & ".*", False) Then
    End If

    GoTo Clean_And_Exit
    
Validation_Error:
    DoCmd.Hourglass False
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Set MyAdo = Nothing
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "Match Validation Error"
    Exit Sub
    
    
Error_Handler:
    DoCmd.Hourglass False
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "Error during Match process"
    On Error Resume Next
    myCode_ADO.RollbackTrans
    
Clean_And_Exit:


    
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Set MyAdo = Nothing
    
    If Not IsNull(MyFastScan.CoverSheetNum) Then Call FastScan_UnLock(MyFastScan.CoverSheetNum)
    
    StartAllOver
    
    If Not CloseAllViewerWindows(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\"))) Then
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If

    
    If mstrOnlyThisCoverSheet <> "" And ErrMsgTxt = "" Then
        DoCmd.Close acForm, Me.Name
    End If
    
End Sub



Private Sub ProcessNoMatchReject(ProcessType As String)
    Dim ErrMsgTxt As String
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    
    DoCmd.Hourglass True
    
    If ProcessType = "No Match" Then ProcessType = "NOMATCH" 'in case someone fixes the spelling of the button's caption :)
    
    ProcessType = UCase(ProcessType)
    
    If ProcessType <> "NOMATCH" And ProcessType <> "REJECT" And ProcessType <> "FINISH" Then
        ErrMsgTxt = "Invalid Process Type: " & ProcessType & ". Alert Data Services"
        GoTo Error_Handler
    End If
    
    'Make sure an ImageType has been selected
    If Me.lstImageType.ListIndex = -1 Then
        ErrMsgTxt = "You must select an Image Type before No Matching."
        GoTo Validation_Error
    End If
    
    'validate that a the correct Image Type was selected for the NoMatch / Match reason if it is enforced
    If mstrRequiredImageType = "Y" And lstImageType <> mstrDefaultImageType Then
        ErrMsgTxt = "The selected reason only works for Image Type = " & mstrDefaultImageType & " , please fix your selections and try again."
        GoTo Validation_Error
    End If
    
    'Make sure ICN is long enough for Data Entry
    If mstrDataEntry_ICN_Flag = "Y" And (Len(Me.txtICN) < 12 Or InStr(1, Me.txtICN, "%") > 0 Or InStr(1, Me.txtICN, "_") > 0) Then
        ErrMsgTxt = "You must enter the ICN listed on the image first."
        GoTo Validation_Error
    End If
    
    'Make sure PayerName was selected for Data Entry
    If mstrDataEntry_PayerName_Flag = "Y" And Me.cmbPayerName.ListIndex = -1 Then
        ErrMsgTxt = "You must enter the Payer Name listed on the image first."
        GoTo Validation_Error
    End If
    
    'Make sure PayerName was selected for Data Entry
    If mstrDataEntry_CnlyProvID_Flag = "Y" And Len(Trim(Me.txtCnlyProvID)) < 3 Then
        ErrMsgTxt = "You must enter the Provider Number listed on the Address Change request."
        GoTo Validation_Error
    End If
    
    'Make sure CnlyProvID is long enough for Data Entry
    If mstrDataEntry_CnlyProvID_Flag = "Y" And (Len(Me.txtCnlyProvID) < 4 Or InStr(1, Me.txtCnlyProvID, "%") > 0 Or InStr(1, Me.txtCnlyProvID, "_") > 0) Then
        ErrMsgTxt = "You must enter a valid CnlyProvID."
        GoTo Validation_Error
    End If
    
    
    'Make sure an no match reason has been selected
    If Me.lstNoMatchReason.ListIndex = -1 Then
        ErrMsgTxt = "You must select a No Match reason first."
        GoTo Validation_Error
    End If
    
    If Not CloseAllViewerWindows(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\"))) Then
        ErrMsgTxt = "Cannot automatically close the viewer! please close the Image window manually and try again."
        GoTo Error_Handler
    End If
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    'Start transaction
    myCode_ADO.BeginTrans
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_ProcessNoMatchReject_v5"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheetNum").Value = MyFastScan.CoverSheetNum
    cmd.Parameters("@pActionType").Value = ProcessType
    cmd.Parameters("@pNoMatchReasonCd").Value = Me.lstNoMatchReason
    cmd.Parameters("@pImageType").Value = Me.lstImageType
    cmd.Parameters("@pFileExt").Value = mstrImageFileExt
    cmd.Parameters("@pDE_ICN").Value = Me.txtICN
    cmd.Parameters("@pDE_PayerName").Value = Me.cmbPayerName
    cmd.Parameters("@pDE_CnlyProvID").Value = Me.txtCnlyProvID
    cmd.Parameters("@pUserID").Value = mstrCurrentUser
    
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    If ErrMsgTxt <> "" Then
        GoTo Error_Handler
    End If
    
    myCode_ADO.CommitTrans
    
    GoTo Clean_And_Exit
    
Validation_Error:
    DoCmd.Hourglass False
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "No Match validation error"
    Exit Sub
    
Error_Handler:
    DoCmd.Hourglass False
    If ErrMsgTxt <> "" Then
        MsgBox ErrMsgTxt, vbExclamation, "Error during No Match / Finish process."
    End If
    On Error Resume Next
    myCode_ADO.RollbackTrans
    
Clean_And_Exit:

    If Not IsNull(MyFastScan.CoverSheetNum) Then Call FastScan_UnLock(MyFastScan.CoverSheetNum)
    
    StartAllOver
    
    DoCmd.Hourglass False
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    
End Sub





Private Sub cmdOpenFile_Click()

    Dim bolReadyToProcess As Boolean
    Dim strThisCoverSheet As String
 
    
    Dim strImageFileCheckResult As String
    
    strThisCoverSheet = ""
    bolReadyToProcess = False
    
    If mstrMatchINPath = "" Then
        MsgBox "Variable mstrMatchINPath is empty, cannot continue", vbExclamation, "Cannot get next CoverSheetNum"
        GoTo ExitSub
    End If


    DoCmd.Hourglass True
    Do Until bolReadyToProcess
    
                    
    
        If MyFastScan.CoverSheetNum = "NA" Then GetNextCoverSheet
        
        Select Case MyFastScan.CoverSheetNum
        
            Case "NA"
                MsgBox "There are no more " & IIf(Me.TogWorkMode.Caption = "Only", Me.cmbProviderFolderWork, "") & " coversheet records ready to be processed", vbExclamation
                Me.cmbProviderFolderWork.Enabled = True
                Exit Do
            Case "ERROR"
                'MsgBox "There was an error trying to fetch the next available quick scan coversheet", vbExclamation
                Exit Do
            Case Else
                If Not FastScan_Lock(MyFastScan.CoverSheetNum) Then
                    MyFastScan.CoverSheetNum = "NA"
                    If mstrOnlyThisCoverSheet <> "" Then
                        MsgBox "ERROR: Cannot open the image file because of the following error: Could Not Lock the Coversheet", vbExclamation, "Error FastScan_Main cmdOpenFile_Click"
                        GoTo ExitSub
                    End If
                End If
                
        End Select
        
        'start with OK
        strImageFileCheckResult = "OK"
        

        Call GetImageFileExt(mstrMatchINPath & MyFastScan.rsCoverSheet("ProviderFolder") & "\" & MyFastScan.rsCoverSheet("ImageName"))
        
        If mstrImageFileExt = "" Then
            strImageFileCheckResult = "FILENOTFOUND"
        End If
        
        If strImageFileCheckResult = "OK" Then
            mstrImageFullName = mstrMatchINPath & MyFastScan.rsCoverSheet("ProviderFolder") & "\" & MyFastScan.rsCoverSheet("ImageName") & "." & mstrImageFileExt
            If Not FileLocked(mstrImageFullName) Then
                strImageFileCheckResult = "OK"
            Else
                strImageFileCheckResult = "FILELOCKED"
            End If
        End If
        
        If strImageFileCheckResult = "OK" Then
            bolReadyToProcess = True
        Else
            Call MarkFileIssue(MyFastScan.CoverSheetNum, strImageFileCheckResult)
            Call FastScan_UnLock(MyFastScan.CoverSheetNum)
            MyFastScan.CoverSheetNum = "NA"
            If mstrOnlyThisCoverSheet <> "" Then
                MsgBox "ERROR: Cannot open the image file because of the following error: " & strImageFileCheckResult, vbExclamation, "Error FastScan_Main cmdOpenFile_Click"
                GoTo ExitSub
            End If
        End If
        
        'if there was a specific Coversheet to open and it is not ready to open by now then it means there was an issue and we must exit now
        If mstrOnlyThisCoverSheet <> "" And Not bolReadyToProcess Then
            Call FastScan_UnLock(mstrOnlyThisCoverSheet)
            GoTo ExitSub
        End If
        
    Loop
    

    
    If bolReadyToProcess Then
    
        If OpenImageFile(mstrImageFullName) = True Then
            
            DoCmd.Hourglass False
            
            'MyFastScan.LoadCoverSheet(
            
            If MyFastScan.rsCoverSheet("ProcStatusCd") = "SPLITCOMPLETE" Then 'should a coversheet that was already split be available for processing again? let's start with no
                'MsgBox "This Coversheet was the source of a split. You cannot process it. It will open in Read-Only mode!", vbInformation, "FastScan - Main"
            Else
                ActivateFields
                Me.txtICN.SetFocus
            End If
            Me.cmbProviderFolderWork.Enabled = False
            
            RunSearch iInitialLoad:=1
        Else
        
            DoCmd.Hourglass False
            
            MsgBox "There was an error trying to open the image file: " & mstrImageFullName, vbExclamation
            StartAllOver
        End If
    Else
        DoCmd.Hourglass False
        StartAllOver
    End If
    
GoTo ExitSub


ExitSub:
DoCmd.Hourglass False

End Sub




Private Sub cmdPropagateAddAllReq_Click()
    Dim sqlUpdate As String
    Dim ErrMsgTxt As String
    Dim SearchResult As String
    
    If Me.cmdPropagateAddAllReq.Tag = 0 Then
        ErrMsgTxt = "Search only by the full request number in order to use this feature!"
        MsgBox ErrMsgTxt, vbExclamation, "Error"
        Exit Sub
    End If
    
    If ResultsOtherSelected("P") > 0 Then 'if there are already propagate claims
        ErrMsgTxt = "You cannot use this option when you already have claims in the propagate list. Clear all Propagate first then try again."
        MsgBox ErrMsgTxt, vbExclamation, "Error"
        Exit Sub
    End If
    
    If ResultsSelected() <> 1 Then
        ErrMsgTxt = "One results row must be marked with a check first"
        MsgBox ErrMsgTxt, vbExclamation, "Error"
        Exit Sub
    End If
   
    'insert all claims from request into the propagate list
    sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ResultType = 'P', ProcessInd = 1 WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'M' and RequestNum = (Select RequestNum from v_CA_SCANNING_FastScan_Search where AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'M' and ProcessInd = True)"
    CurrentDb.Execute (sqlUpdate)
    
    'delete all not propagete results
    sqlUpdate = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType <> 'P'"
    CurrentDb.Execute (sqlUpdate)
    
    'cmdClearAllResults_Click
    Me.subfrm_Results_Other.Form.Requery
    Me.Refresh
End Sub

Private Sub cmdRelatedSelectAll_Click()
    Dim sqlUpdate As String
    sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = True WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'R' and cnlyclaimnum = '" & IIf(Me.subfrm_Results.Form("ProcessInd"), Me.subfrm_Results.Form("CnlyClaimNum"), "") & "'"
    CurrentDb.Execute (sqlUpdate)
    Me.Refresh
End Sub

Private Sub cmdRelatedUnselectAll_Click()
    Dim sqlUpdate As String
    sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ProcessInd = false WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'R'"
    CurrentDb.Execute (sqlUpdate)
    Me.Refresh
End Sub



Private Sub Form_Load()

    mintSessionID = 0
    DoCmd.Hourglass False
    Me.frmAppID = CstrFrmAppID
    
    If left(Nz(Me.OpenArgs, ""), 3) = "frm" Then mstrCalledFrom = Me.OpenArgs
    
    mstrCurrentUser = Identity.UserName
    'mintCurrentAuditNum = 3055 'Identity.AuditNum

    mbolRejectUser = False
    
    Set MyFastScan = New clsFastScan

        
    mstrMatchINPath = Nz(DLookup("MatchINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrMatchOUTPath = Nz(DLookup("MatchOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrSplitINPath = Nz(DLookup("SplitINPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrSplitOUTPath = Nz(DLookup("SplitOUTPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrTIFViewerPath = Nz(DLookup("TIFViewerPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrAcrobatPath = Nz(DLookup("AcrobatPath", "v_CA_SCANNING_FastScan_Config", "AccountID = " & gintAccountID), "")
    mstrLocalPath = Nz(DLookup("LocalPath", "SCANNING_Config", "AccountID = " & gintAccountID), "")
    mstrClaimAttachPath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = 'AuditClmRef' AND AccountID = " & gintAccountID), "")
    mstrProvAttachPath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = 'ProvRef' AND AccountID = " & gintAccountID), "")
    
    
    
    If Nz(mstrMatchINPath, "") = "" _
        Or Nz(mstrMatchOUTPath, "") = "" _
        Or Nz(mstrTIFViewerPath, "") = "" _
        Or Nz(mstrAcrobatPath, "") = "" _
        Or Nz(mstrLocalPath, "") = "" _
        Or Nz(mstrClaimAttachPath, "") = "" _
        Or Nz(mstrProvAttachPath, "") = "" _
        Then
        MsgBox "One or more of the FastScan config values are not setup for this Account. Please check the FastScan_Config table.", vbInformation, "Error with FastScan config values"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    If Right$(mstrMatchINPath, 1) <> "\" Then mstrMatchINPath = mstrMatchINPath & "\"
    If Right$(mstrMatchOUTPath, 1) <> "\" Then mstrMatchOUTPath = mstrMatchOUTPath & "\"
    If Right$(mstrSplitINPath, 1) <> "\" Then mstrSplitINPath = mstrSplitINPath & "\"
    If Right$(mstrSplitOUTPath, 1) <> "\" Then mstrSplitOUTPath = mstrSplitOUTPath & "\"
    If Right$(mstrTIFViewerPath, 1) = "\" Then mstrTIFViewerPath = left(mstrTIFViewerPath, Len(mstrTIFViewerPath) - 1)
    If Right$(mstrAcrobatPath, 1) = "\" Then mstrAcrobatPath = left(mstrAcrobatPath, Len(mstrAcrobatPath) - 1)
    If Right$(mstrLocalPath, 1) <> "\" Then mstrLocalPath = mstrLocalPath & "\"
    If Right$(mstrClaimAttachPath, 1) <> "\" Then mstrClaimAttachPath = mstrClaimAttachPath & "\"
    If Right$(mstrProvAttachPath, 1) <> "\" Then mstrProvAttachPath = mstrProvAttachPath & "\"
    
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct FolderName from FastScanMaint.v_FastScan_UserAuthFolders where accountid = " & gintAccountID & " order by FolderName"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            MsgBox "There are no FastScan folders setup for this account or user does not have access to any folder!" & vbNewLine & vbNewLine & "Cannot continue.", vbInformation, "Error: FastScan folders missing"
            DoCmd.Close acForm, Me.Name
            GoTo Cleanup
            Exit Sub
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
    
    
    'SELECT ImageType, ImageTypeDisplay, ImageTypeDesc FROM SCANNING_XREF_ImageType WHERE (((Active)="Y") AND ((FastScanVisible)="Y")) ORDER BY IIf(imagetype="MR" Or imagetype="RECON",1,2), ImageType;
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT ImageType, ImageTypeDisplay, ImageTypeDesc FROM FastScanmaint.v_SCANNING_XREF_ImageType_Account WHERE Active='Y' AND FastScanVisible='Y' AND AccountID = " & str(gintAccountID) & " ORDER BY Case imagetype when 'MR' then 1 when 'RECON' then 2 else 3 end, ImageType"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            MsgBox "There are no Image Types setup for this account!" & vbNewLine & vbNewLine & "Cannot continue.", vbInformation, "Error: Image Types missing"
            DoCmd.Close acForm, Me.Name
            GoTo Cleanup
            Exit Sub
        Else
            Set Me.lstImageType.RecordSet = oRs
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
        
    
    
    mstrOnlyThisCoverSheet = ""
    
   
    Call Account_Check(Me)
   
    'miAppPermission = UserAccess_Check(Me)

    
    miAppPermission = GetAppPermission(Me.frmAppID)
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowView = (miAppPermission And gcAllowView)
    mbAllowDelete = (miAppPermission And gcAllowDelete)
    
    If miAppPermission = 0 Or mbAllowView = False Then
        MsgBox "You do not have permission to view this form"
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    'This authorizes the users to do Fuzzy Matches
    '*** NOTE this one uses the FastScanMain appid ***
    If mbAllowChange Then
        mbolFuzzyMatchUser = True
    End If
    
    'This authorizes the users to process Priority images
    '*** NOTE this one uses the FastScanMain appid ***
    If mbAllowAdd Then
        mbolPriorityUser = True
    End If
    
    
        If left(Me.OpenArgs, 3) = "frm" Or Me.OpenArgs() & "" = "" Then
            If Me.OpenArgs() & "" <> "" Then mstrCalledFrom = Me.OpenArgs
            Me.cmbProviderFolderWork.Enabled = True
            StartAllOver
        Else
        
            Me.frmAppID = "FastScanIssues"

            'miAppPermission = UserAccess_Check(Me)

            miAppPermission = GetAppPermission(Me.frmAppID)
            mbAllowChange = (miAppPermission And gcAllowChange)
            mbAllowAdd = (miAppPermission And gcAllowAdd)
            mbAllowView = (miAppPermission And gcAllowView)
            mbAllowDelete = (miAppPermission And gcAllowDelete)
            
            If miAppPermission = 0 Or mbAllowView = False Then
                MsgBox "You do not have permission to process No Matches"
                DoCmd.Close acForm, Me.Name
                Exit Sub
            End If

            'This authorizes the users that can reject images that has been no matches
            '*** NOTE this one uses the FastScanIssues appid ***
            If mbAllowChange Then
                mbolRejectUser = True
            End If
        
            If Not mbolRejectUser Then
                MsgBox "You are not authorized to process the Issues queue.", vbExclamation, "Unauthorized"
                DoCmd.Close acForm, Me.Name
                Exit Sub
            End If
            StartAllOver
            Me.cmbProviderFolderWork.Enabled = False
            mstrOnlyThisCoverSheet = Me.OpenArgs
            cmdOpenFile_Click
            If MyFastScan.CoverSheetNum = "NA" Or MyFastScan.CoverSheetNum = "ERROR" Then
                DoCmd.Close acForm, Me.Name
                Exit Sub
            End If

        End If
    
    'Call FillReasonList
    
    Call FillFolderCombo
    
Cleanup:
    oAdo.DisConnect
    Set oAdo = Nothing
    Set oRs = Nothing
    
End Sub

Private Sub GetNextCoverSheet()

Dim spReturnVal As Integer
Dim ErrMsg As String
Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command

    MyFastScan.CoverSheetNum = "NA"

    If Identity.UserName = "" Then
        Exit Sub
    End If

    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_GetNextCoverSheet_v5"
    cmd.Parameters.Refresh
    cmd.Parameters("@pAccountID") = gintAccountID
    cmd.Parameters("@pUserID") = mstrCurrentUser
    cmd.Parameters("@pPriorityUser") = IIf(mbolPriorityUser, 1, 0)
    cmd.Parameters("@pOnlyThisCoverSheet") = mstrOnlyThisCoverSheet
    cmd.Parameters("@pProviderFolder") = Me.cmbProviderFolderWork
    cmd.Parameters("@pWorkMode") = Me.TogWorkMode.Caption
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
    If spReturnVal <> 0 Or ErrMsg <> "" Then
        MyFastScan.CoverSheetNum = "NA" 'Nz(cmd.Parameters("@pNextCS_Number"), "")
        MsgBox ErrMsg, vbExclamation, "Error GetNextCoverSheet"
        GoTo Exit_Sub
    End If
    mintSessionID = 0
    mintSessionID = GetAppKey("SCANNING")
    
    MyFastScan.LoadCoverSheet Nz(cmd.Parameters("@pNextCS_Number"), "NA")
    
    Set Me.RecordSet = MyFastScan.rsCoverSheet
    
    LoadTXTFields
        
'    Me.txtCoverSheetNum =
'    MyFastScan.rsCoverSheet("ImageName") = Nz(cmd.Parameters("@pNextCS_ImageName"), "NA")
'    MyFastScan.rsCoverSheet("ProviderFolder") = Nz(cmd.Parameters("@pNextCS_ProviderFolder"), "")
'    Me.txtReceivedDt = Nz(cmd.Parameters("@pNextCS_ReceivedDt"), "")
'    Me.txtLinkedDt = Nz(cmd.Parameters("@pNextCS_LinkedDt"), "")
'    Me.txtTrackingNum = Nz(cmd.Parameters("@pNextCS_TrackingNum"), "")
'    Me.txtReceivedMeth = Nz(cmd.Parameters("@pNextCS_ReceivedMeth"), "")
'    Me.lstImageType = Nz(cmd.Parameters("@pNextCS_ImageType"), "")
'    MyFastScan.rsCoverSheet("ProcStatusCd") = Nz(cmd.Parameters("@pNextCS_ProcStatusCd"), "")
'    MyFastScan.rsCoverSheet("NoMatchReasonCd") = Nz(cmd.Parameters("@pNextCS_NoMatchReasonCd"), "")
'    Me.txtCnlyProvID = Nz(cmd.Parameters("@pNextCS_DE_CnlyProvID"), "")
'    Me.txtICN = Nz(cmd.Parameters("@pNextCS_DE_ICN"), "")
'    Me.cmbPayerName = Nz(cmd.Parameters("@pNextCS_DE_PayerName"), "")

    
Exit_Sub:

    Set MyCodeAdo = Nothing
    Set cmd = Nothing


End Sub

Sub LoadTXTFields()

    With MyFastScan
        
        If .CoverSheetNum <> "NA" Then
            'Me.txtCoverSheetNum = .CoverSheetNum
            'Me.txtImageName = Nz(.rsCoverSheet("ImageName"), "")
            'Me.txtProviderFolder = Nz(.rsCoverSheet("ProviderFolder"), "")
            'Me.txtReceivedDt = Nz(.rsCoverSheet("ReceivedDt"), "")
            'Me.txtLinkedDt = Nz(.rsCoverSheet("ScannedDt"), "")
            'Me.txtTrackingNum = Nz(.rsCoverSheet("TrackingNum"), "")
            'Me.txtReceivedMeth = Nz(.rsCoverSheet("ReceivedMeth"), "")
            Me.lstImageType = Nz(.rsCoverSheet("ImageType"), "")
            'Me.txtProcStatusCd = Nz(.rsCoverSheet("ProcStatusCd"), "")
            'Me.txtNoMatchReasonCd = Nz(.rsCoverSheet("NoMatchReasonCd"), "")
            Me.txtCnlyProvID = Nz(.rsCoverSheet("DataEntry_CnlyProvID"), "")
            Me.txtICN = Nz(.rsCoverSheet("DataEntry_ICN"), "")
            Me.cmbPayerName = Nz(.rsCoverSheet("DataEntry_PayerName"), "")
        Else
            'Me.txtCoverSheetNum = .CoverSheetNum
            'Me.txtCoverSheetNum = ""
            'Me.txtImageName = ""
            'Me.txtProviderFolder = ""
            'Me.txtReceivedDt = ""
            'Me.txtLinkedDt = ""
            'Me.txtTrackingNum = ""
            'Me.txtReceivedMeth = ""
            Me.lstImageType = ""
            'Me.txtProcStatusCd = ""
            'Me.txtNoMatchReasonCd = ""
            Me.txtCnlyProvID = ""
            Me.txtICN = ""
            Me.cmbPayerName = ""
        End If
    End With

End Sub

'Function ImageFileExists(strImageFileName As String) As String
'
'On Error GoTo ErrorHappened
'
'    Dim Result As String
'
'    'Check for existing
'    If Len(Dir(strImageFileName)) > 0 Then
'        If Not FileLocked(strImageFileName) Then
'            Result = "OK"
'        Else
'            Result = "FILELOCKED"
'        End If
'    Else
'        Result = "FILENOTFOUND"
'    End If
'
'ExitNow:
'    On Error Resume Next
'    ImageFileExists = Result
'    Exit Function
'ErrorHappened:
'    Result = "FILEERROR"
'    Resume ExitNow
'End Function


Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      'MsgBox "Error #" & str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
   End If
End Function
                
Function FastScan_Lock(strCoverSheet As String) As Boolean
Dim spReturnVal As Integer
Dim ErrMsg As String
Dim Result As Boolean
Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command

    If Identity.UserName = "" Then
        Exit Function
    End If
    
    Result = False

    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_CoverSheetLock"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheet") = strCoverSheet
    cmd.Parameters("@pUserID") = mstrCurrentUser
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")

    If spReturnVal <> 0 Or ErrMsg <> "" Then
        Result = False
    Else
        Result = True
    End If

    Set MyCodeAdo = Nothing
    Set cmd = Nothing

    FastScan_Lock = Result

End Function
Function FastScan_UnLock(strCoverSheet As String) As Boolean
Dim spReturnVal As Integer
Dim ErrMsg As String
Dim Result As Boolean
Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command

    If Identity.UserName = "" Then
        Exit Function
    End If
    
    If Nz(strCoverSheet, "") = "" Or strCoverSheet = "NA" Or strCoverSheet = "ERROR" Then
        Exit Function
    End If
    
    Result = False

    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_CoverSheetUnLock"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheet") = strCoverSheet
    cmd.Parameters("@pUserID") = mstrCurrentUser
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")

    If spReturnVal <> 0 Or ErrMsg <> "" Then
        Result = False
    Else
        Result = True
    End If

    Set MyCodeAdo = Nothing
    Set cmd = Nothing

    FastScan_UnLock = Result

End Function


Function MarkFileIssue(strCoverSheet As String, strIssueCode As String) As Boolean
Dim spReturnVal As Integer
Dim ErrMsg As String
Dim Result As Boolean
Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command

    If Identity.UserName = "" Then
        Exit Function
    End If
    
    Result = False

    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_CoverSheetFileIssue"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheet") = strCoverSheet
    cmd.Parameters("@pUserID") = mstrCurrentUser
    cmd.Parameters("@pFileIssue") = strIssueCode
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")

    If spReturnVal <> 0 Or ErrMsg <> "" Then
        Result = False
    Else
        Result = True
    End If

    Set MyCodeAdo = Nothing
    Set cmd = Nothing

End Function

Function OpenImageFile(strImageFileName As String) As Boolean

Dim fso As New FileSystemObject


    OpenImageFile = False
    
    If Not CloseAllViewerWindows(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\"))) Then
        DoCmd.Close acForm, Me.Name
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")

 
    If Not fso.FileExists(strImageFileName) Then
            MsgBox "The File you are looking for was renamed or moved. Check the file name and try again", vbCritical, "File Does Not Exists"
            GoTo Cleanup
    End If
    
    'SetFileReadOnly (strFileName)
    If UCase(Right(strImageFileName, 3)) = "TIF" Then
    
        Shell mstrTIFViewerPath & " " & strImageFileName, vbNormalFocus
    
        'mstrImageViewer = Right(mstrTIFViewerPath, Len(mstrTIFViewerPath) - InStrRev(mstrTIFViewerPath, "\"))
        
        OpenImageFile = True
        
    ElseIf UCase(Right(strImageFileName, 3)) = "PDF" Then
    
        Shell mstrAcrobatPath & " " & strImageFileName, vbNormalFocus
        
        'mstrImageViewer = Right(mstrAcrobatViewerPath, Len(mstrAcrobatViewerPath) - InStrRev(mstrAcrobatViewerPath, "\"))
        
        OpenImageFile = True
        
    Else
        MsgBox "This is not a PDF or TIF image, code is not ready to open this file. Contact DS.", vbExclamation
        mstrImageViewer = ""
    End If

Cleanup:

    Set fso = Nothing

End Function




Private Sub cmdSearch_Click()
    RunSearch iInitialLoad:=0
End Sub
Private Sub RunSearch(iInitialLoad As Integer, Optional strCnlyClaimNum As String)
'    Dim oAdo As clsADO
'    Dim oRs As AdoDb.Recordset

On Error GoTo Error_Handler

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim ErrorReturned As String
    Dim CriteriaPoints As Integer
    Dim spReturnVal As Integer
    Dim strResultTypesFound As String
    
    Dim strICN As String
    Dim strCnlyProvID As String
    Dim strPatFirstInit As String
    Dim strInstanceID As String
    Dim strPatLastName As String
    Dim strPatDOB As String
    Dim strClmFromDt As String
    Dim strPatCan As String
    Dim strPatBic As String
    Dim strMRNumber As String
    Dim strALJNumber As String
    Dim strQICNumber As String
    
    
    Dim strAllStringCriteria As String
    Dim bolUsingWildcards As Boolean
    
    strICN = Nz(Me.txtICN, "")
    strCnlyProvID = Nz(Me.txtCnlyProvID, "")
    strPatFirstInit = Nz(Me.txtPatFirstInit, "")
    strInstanceID = Nz(Me.txtInstanceID, "")
    strPatLastName = Nz(Me.txtPatLastName, "")
    strPatDOB = IIf(Nz(Me.txtPatDOB, "") <> "", Format(Me.txtPatDOB, "yyyy-mm-dd"), "")
    strClmFromDt = IIf(Nz(Me.txtclmFromDt, "") <> "", Format(Me.txtclmFromDt, "yyyy-mm-dd"), "")
    strPatCan = Nz(Me.txtPatCAN, "")
    strPatBic = Nz(Me.txtPatBIC, "")
    strMRNumber = Nz(Me.txtMRNumber, "")
    strALJNumber = Nz(Me.txtALJNumber, "")
    strQICNumber = Nz(Me.txtQICNumber, "")
    strCnlyClaimNum = Nz(strCnlyClaimNum, "")
    
    If iInitialLoad = 0 Then
    
        If Me.txtPatFirstInit = "%" Or Me.txtPatFirstInit = "_" Then
            Me.txtPatFirstInit = ""
        End If
    
        If Me.txtPatCAN = "%" Or Me.txtPatCAN = "_" Then
            Me.txtPatCAN = ""
        End If
    
        If Me.txtPatBIC = "%" Or Me.txtPatBIC = "_" Then
            Me.txtPatBIC = ""
        End If
    
    
        If strCnlyClaimNum = "" Then
            strAllStringCriteria = Nz(Me.txtCnlyProvID, "") & _
                                    Nz(Me.txtInstanceID, "") & _
                                    Nz(Me.txtICN, "") & _
                                    Nz(Me.txtPatFirstInit, "") & _
                                    Nz(Me.txtPatLastName, "") & _
                                    Nz(Me.txtPatCAN, "") & _
                                    Nz(Me.txtPatBIC, "") & _
                                    Nz(Me.txtMRNumber, "") & _
                                    Nz(Me.txtALJNumber, "") & _
                                    Nz(Me.txtQICNumber, "")
            
            'If the only search criteria is the InstanceID and it is at least 17 characters then we can say the results will be all claims in that request
            'therefore we set the flag in the button to 1
            If strAllStringCriteria = Me.txtInstanceID And Len(Me.txtInstanceID) >= 17 Then
                Me.cmdPropagateAddAllReq.Tag = 1
            Else
                Me.cmdPropagateAddAllReq.Tag = 0
            End If
            
            bolUsingWildcards = InStr(1, strAllStringCriteria, "%") > 0 Or InStr(1, strAllStringCriteria, "_") > 0
            If bolUsingWildcards And Not mbolFuzzyMatchUser Then
                MsgBox "You are not authorized to do Fuzzy Searches"
                Exit Sub
            End If
        
            If Me.txtCnlyProvID <> "" And IsNull(Me.txtCnlyProvID) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtCnlyProvID) < 4 Then
                    MsgBox "Provider Number must be at least 4 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtInstanceID <> "" And IsNull(Me.txtInstanceID) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtInstanceID) <> 17 Then
                    MsgBox "Request Number must be 17 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtICN <> "" And IsNull(Me.txtICN) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtICN) < 12 Then
                    MsgBox "ICN must be at least 12 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtPatFirstInit <> "" And IsNull(Me.txtPatFirstInit) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtPatFirstInit) > 1 Then
                    Me.txtPatFirstInit = left(Me.txtPatFirstInit, 1)
                End If
            End If
            
            If Me.txtPatLastName <> "" And IsNull(Me.txtPatLastName) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtPatLastName) < 2 Then
                    MsgBox "Patient Last Name must be at least 2 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtPatDOB <> "" And IsNull(Me.txtPatDOB) = False Then
                If IsDate(Me.txtPatDOB) = False Then
                    MsgBox "Patient DOB must be a valid date."
                    Exit Sub
                ElseIf CDate(Me.txtPatDOB) >= Date Then
                    MsgBox "Patient DBO cannot be a future date."
                    Exit Sub
                End If
            End If
            
            If Me.txtclmFromDt <> "" And IsNull(Me.txtclmFromDt) = False Then
                If IsDate(Me.txtclmFromDt) = False Then
                    MsgBox "ClmFromDt Date must be a valid date."
                    Exit Sub
                ElseIf CDate(Me.txtclmFromDt) >= Date Then
                    MsgBox "ClmFromDt cannot be a future date."
                    Exit Sub
                End If
            End If
            
            If Me.txtPatCAN <> "" And IsNull(Me.txtPatCAN) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtPatCAN) < 5 Then
                    MsgBox "Patient CAN must be at least 5 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtMRNumber <> "" And IsNull(Me.txtMRNumber) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtMRNumber) < 4 Then
                    MsgBox "MR Number must be at least 4 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtALJNumber <> "" And IsNull(Me.txtALJNumber) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtALJNumber) < 5 Then
                    MsgBox "ALJ Appeal Number must be at least 5 characters"
                    Exit Sub
                End If
            End If
            
            If Me.txtQICNumber <> "" And IsNull(Me.txtQICNumber) = False And Not mbolFuzzyMatchUser Then
                If Len(Me.txtQICNumber) < 5 Then
                    MsgBox "QIC Appeal Number must be at least 5 characters"
                    Exit Sub
                End If
            End If
        
        End If

    End If

    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .commandType = adCmdStoredProc
        .CommandText = "FastScan.usp_FastScan_Search_v7"
        .Parameters.Refresh
        .Parameters("@pInitialLoad") = iInitialLoad
        .Parameters("@pFuzzyMatch") = IIf(mbolFuzzyMatchUser And bolUsingWildcards, 1, 0)
        .Parameters("@pAccountID") = gintAccountID
        .Parameters("@pCnlyClaimNum") = strCnlyClaimNum
        .Parameters("@pICN") = strICN
        .Parameters("@pCnlyProvID") = strCnlyProvID
        .Parameters("@pPatFirstInit") = strPatFirstInit
        .Parameters("@pInstanceID") = strInstanceID
        .Parameters("@pPatLastName") = strPatLastName
        .Parameters("@pPatDOB") = strPatDOB
        .Parameters("@pClmFromDt") = strClmFromDt
        .Parameters("@pPatCan") = strPatCan
        .Parameters("@pPatBic") = strPatBic
        .Parameters("@pMRNumber") = strMRNumber
        .Parameters("@pALJNumber") = strALJNumber
        .Parameters("@pQICNumber") = strQICNumber
        .Parameters("@pCoverSheetNum") = MyFastScan.CoverSheetNum
        .Parameters("@pUserID") = mstrCurrentUser
        .Parameters("@pSessionID") = mintSessionID
        DoCmd.Hourglass True
        .Execute
        DoCmd.Hourglass False
        strResultTypesFound = .Parameters("@pResultTypesFound")
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = .Parameters("@pErrMsg")
        'mintSessionID = .Parameters("@pSessionID")
    End With
    
    If spReturnVal <> 0 Or ErrorReturned <> "" Then
        MsgBox ErrorReturned & vbNewLine & vbNewLine & "Error, check the search parameters and try again.", vbExclamation, "Error"
        'Exit Sub
    End If

    'If this is an initial load and there were barcode results (no errors)
    If iInitialLoad = 1 And InStr(1, strResultTypesFound, "B", vbTextCompare) > 0 Then
        Me.subfrm_Results.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search as a where AccountID = " & gintAccountID & " and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'B' order by CnlyProvID, ClmStatus, ICN"
                                                '& _ " and cnlyclaimnum not in (select cnlyclaimnum from v_CA_SCANNING_FastScan_Search where UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and cnlyclaimnum = a.cnlyclaimnum and ResultType = 'P' )"
        Me.TogShowBarCodeResults.Enabled = True
        Me.TogShowBarCodeResults.Caption = "Barcode Results"
        Me.TogShowBarCodeResults.Value = -1
        Me.TogShowBarCodeResults.Locked = False
    'if this is an initial load and there were no barcode results
    ElseIf iInitialLoad = 1 And InStr(1, strResultTypesFound, "B", vbTextCompare) = 0 Then
        Me.TogShowBarCodeResults.Value = 0
        Me.TogShowBarCodeResults.Caption = "No Barcode Results"
        Me.TogShowBarCodeResults.Enabled = False
    'is not an initial load but a regular search and there were results
    ElseIf iInitialLoad = 0 And InStr(1, strResultTypesFound, "M", vbTextCompare) > 0 Then
        Me.subfrm_Results.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search as a where AccountID = " & gintAccountID & " and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'M'" & _
                                                " and cnlyclaimnum not in (select cnlyclaimnum from v_CA_SCANNING_FastScan_Search where AccountID = " & gintAccountID & " and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and cnlyclaimnum = a.cnlyclaimnum and ResultType = 'P' ) order by CnlyProvID, ClmStatus, ICN "
        'if there were barcode results but the user chose to do a search let's give him the option to click the button and show the barcode results
        If Me.TogShowBarCodeResults.Enabled Then
            Me.TogShowBarCodeResults.Caption = "Show Barcode Results"
            Me.TogShowBarCodeResults.Value = 0
            TogShowBarCodeResults.Locked = False
        End If
    'is not an initial load but a regular search and there were no results
    ElseIf ErrorReturned = "" Then
        MsgBox "No matches found, check the search parameters and try again.", vbExclamation, "Search Claims"
    End If
    Me.subfrm_Results.Form.Requery

        
    If Me.TogRelated.Value = -1 Then
        Me.subfrm_Results_Other.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where 1=2"
            
        Me.subfrm_Results_Other.Form.Requery
    End If

    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
'    If Me.subfrm_Results.Form.RecordSet.recordCount > 0 Or Me.subfrm_Results_Other.Form.RecordSet.recordCount > 0 Then
'        Me.lstNoMatchReason.Enabled = True
'        Me.cmdClearNoMatch.Enabled = True
'        Me.cmdMatch.Enabled = True
'    Else
'        Me.cmdMatch.Enabled = False
'    End If
'
'    If mbolRejectUser Then
'        Me.cmdNoMatch.Caption = "Finish"
'    Else
'        Me.cmdNoMatch.Caption = "NoMatch"
'    End If
'
'    Me.lstNoMatchReason = ""
'    lstNoMatchReason_Change
'
GoTo Clean_And_Exit
    
Error_Handler:
    DoCmd.Hourglass False
    If Nz(ErrorReturned, "") = "" Then ErrorReturned = Err.Description
    MsgBox ErrorReturned, vbExclamation, "Error during Search process"
    
Clean_And_Exit:
    DecisionSwitch
    
End Sub


Private Sub cmdClearImageType_Click()
    'txtReceivedDt = ""
    If Me.lstImageType <> "" Then
        lstImageType = ""
        lstImageType_Change
    End If
    
End Sub

Private Sub StartAllOver()

    CreatePK "v_CA_SCANNING_FastScan_Search", "AccountID, UserID,SessionID,CoverSheetNum,ResultType,CnlyClaimNum,RelatedClaimNum"

    Me.subfrm_Results.SourceObject = "frm_FastScan_SearchResults"
    Me.subfrm_Results.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search as a where 1=2"
    
    Me.subfrm_Results_Other.SourceObject = "frm_FastScan_SearchResults_Related"
    Me.subfrm_Results_Other.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search where 1=2"
    
    cmdClearAllCriteria_Click
    DeactivateFields
    DoCmd.Hourglass False
    
End Sub

Private Sub ActivateFields()

    'Me.txtReceivedDt.Enabled = True
    Me.cmbPayerName.Enabled = True
    Me.txtICN.Enabled = True
    Me.txtInstanceID.Enabled = True
    Me.txtCnlyProvID.Enabled = True
    Me.txtPatFirstInit.Enabled = True
    Me.txtPatLastName.Enabled = True
    Me.txtPatDOB.Enabled = True
    Me.txtclmFromDt.Enabled = True
    Me.txtPatCAN.Enabled = True
    Me.txtPatBIC.Enabled = True
    Me.txtMRNumber.Enabled = True
    Me.txtALJNumber.Enabled = True
    Me.txtQICNumber.Enabled = True
    Me.cmdClearImageType.Enabled = True
    Me.cmdClearNoMatch.Enabled = True
    Me.cmdSearch.Enabled = True
    Me.cmdNotes.Enabled = True
    Me.cmdNotes.Caption = "Notes"
    If MyFastScan.rsNotes.recordCount > 0 Then
        Me.cmdNotes.Caption = "Notes " & ChrW(&H2713)
    End If
    Me.lstImageType.Enabled = True
    Me.lstImageType.SetFocus
'    If mbolRejectUser Then
'        Me.cmdNoMatch.Caption = "Finish"
'    Else
'        Me.cmdNoMatch.Caption = "NoMatch"
'    End If
    Me.lstNoMatchReason.Enabled = True
    Me.TogPropagate.Enabled = True
    Me.TogRelated.Enabled = True
    Me.TogPropagate.Value = 0
    Me.TogRelated.Value = -1
    Me.cmdRelatedSelectAll.Enabled = True
    Me.cmdRelatedUnselectAll.Enabled = True
    
    DecisionSwitch
    
End Sub


Private Sub DeactivateFields()

    Dim sqlDelete As String
    sqlDelete = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID
    CurrentDb.Execute (sqlDelete)
    
    OtherResultsToggle "Related"


    mintpageCnt = 0
    mstrImageFileExt = ""
    mstrImageFullName = ""
    mstrSplitFlag = "N"
    mstrDataEntry_ICN_Flag = "N"
    mstrDataEntry_PayerName_Flag = "N"
    mstrDataEntry_CnlyProvID_Flag = "N"
    mintSessionID = 0
    Me.cmdOpenFile.SetFocus
    Call cmdClearImageType_Click
    Me.cmdClearImageType.Enabled = False
    Call cmdClearNoMatch_Click
    Me.cmdClearNoMatch.Enabled = False
    Set Me.RecordSet = Nothing
    Call MyFastScan.UnLoadCoverSheet
    Set Me.RecordSet = MyFastScan.rsCoverSheet
'    Me.RecordSource = "Select " & Chr(34) & "NA" & Chr(34) & " as CoverSheetNum, " & _
'                        Chr(34) & Chr(34) & " as ImageName, " & _
'                        Chr(34) & Chr(34) & " as ScannedDt, " & _
'                        Chr(34) & Chr(34) & " as ReceivedDt, " & _
'                        Chr(34) & Chr(34) & " as TrackingNum, " & _
'                        Chr(34) & Chr(34) & " as ReceivedMeth, " & _
'                        Chr(34) & Chr(34) & " as NoMatchReasonCd, " & _
'                        Chr(34) & Chr(34) & " as ProcStatusCd "
    'Me.txtCoverSheetNum = "NA"
    'Me.txtImageName = ""
    'Me.txtLinkedDt = ""
    'Me.txtReceivedDt = ""
    'Me.txtTrackingNum = ""
    'Me.txtReceivedMeth = ""
    'Me.txtReceivedMeth = ""
    'Me.txtCarrier = ""
    'Me.txtReceivedDt.Enabled = False
    Me.cmbPayerName.Enabled = False
    Me.txtICN.Enabled = False
    Me.txtInstanceID.Enabled = False
    Me.txtCnlyProvID.Enabled = False
    Me.txtPatFirstInit.Enabled = False
    Me.txtPatLastName.Enabled = False
    Me.txtPatDOB.Enabled = False
    Me.txtclmFromDt.Enabled = False
    Me.txtPatCAN.Enabled = False
    Me.txtPatBIC.Enabled = False
    Me.txtMRNumber.Enabled = False
    Me.txtALJNumber.Enabled = False
    Me.txtQICNumber.Enabled = False
    Me.cmdSearch.Enabled = False
    Me.cmdNotes.Enabled = False
    Me.cmdNotes.Caption = "Notes"
    Me.cmdMatch.Enabled = False
    Me.cmdNoMatch.Enabled = False
    Me.cmdNoMatch.Caption = "NoMatch"
    Me.lstImageType.Enabled = False
    Me.lstNoMatchReason.Enabled = False
    'Me.cmdOpenFile.Enabled = True
    Me.TogPropagate.Enabled = True
    Me.TogRelated.Enabled = True
    Me.TogPropagate.Value = 0
    Me.TogRelated.Value = -1
    Me.TogPropagate.Enabled = False
    Me.TogRelated.Enabled = False
    Me.cmdRelatedSelectAll.Enabled = False
    Me.cmdRelatedUnselectAll.Enabled = False
    Me.cmdPropagateAdd.Enabled = False
    Me.cmdPropagateAddAllReq.Enabled = False
    Me.cmdPropagateRemove.Enabled = False
    Me.cmdPropagateRemoveAll.Enabled = False
    
End Sub









'
'Function CloseAllImageViewer(strViewer As String) As Boolean
'
'    Dim NumberOfTries As Integer
'    Dim ProcessViewer As Long
'    Dim TerminateResult As Boolean
'
'    If strViewer = "" Then
'        CloseAllImageViewer = True
'        Exit Function
'    End If
'
'    NumberOfTries = 0 'should not be that many IrfanView windows opened, right?
'
'    'Close all instances of IrfanView
'    ProcessViewer = CheckForProcByExe(strViewer)
'    Do While ProcessViewer > 0 Or NumberOfTries > 20
'        TerminateResult = ProcessTerminate(ProcessViewer)
'
''        If TerminateResult = False Then
''            MsgBox "Could not close all instances of IrfanView automatically." & vbNewLine & vbNewLine & _
''                    "You must close them yourself manually, then please try again.", vbExclamation
''            CloseAllIrfanView = False
''            Exit Function
''        End If
'        ProcessViewer = CheckForProcByExe(strViewer)
'        NumberOfTries = NumberOfTries + 1
'    Loop
'
'    If NumberOfTries <= 20 Then
'        CloseAllImageViewer = True
'        Exit Function
'    End If
'
'ErrorClosing:
'    CloseAllImageViewer = False
'    MsgBox "Could not close all instances of " & strViewer & " automatically." & vbNewLine & vbNewLine & _
'            "You must close them yourself manually, then please try again.", vbExclamation
'
'End Function
'
'
'Function CloseAllIrfanView() As Boolean
'
'    Dim NumberOfTries As Integer
'    Dim ProcessIrfanView As Long
'    Dim TerminateResult As Boolean
'
'    NumberOfTries = 0 'should not be that many IrfanView windows opened, right?
'
'    'Close all instances of IrfanView
'    ProcessIrfanView = CheckForProcByExe("i_view32.exe")
'    Do While ProcessIrfanView > 0 Or NumberOfTries > 20
'        TerminateResult = ProcessTerminate(ProcessIrfanView)
'
''        If TerminateResult = False Then
''            MsgBox "Could not close all instances of IrfanView automatically." & vbNewLine & vbNewLine & _
''                    "You must close them yourself manually, then please try again.", vbExclamation
''            CloseAllIrfanView = False
''            Exit Function
''        End If
'        ProcessIrfanView = CheckForProcByExe("i_view32.exe")
'        NumberOfTries = NumberOfTries + 1
'    Loop
'
'    If NumberOfTries <= 20 Then
'        CloseAllIrfanView = True
'        Exit Function
'    End If
'
'ErrorClosing:
'    CloseAllIrfanView = False
'    MsgBox "Could not close all instances of IrfanView automatically." & vbNewLine & vbNewLine & _
'            "You must close them yourself manually, then please try again.", vbExclamation
'
'End Function
'
'Function CloseAllAcrobat() As Boolean
'
'    Dim NumberOfTries As Integer
'    Dim ProcessAcrobat As Long
'    Dim TerminateResult As Boolean
'
'    NumberOfTries = 0 'should not be that many IrfanView windows opened, right?
'
'    'Close all instances of IrfanView
'    ProcessAcrobat = CheckForProcByExe("acrobat.exe")
'    Do While ProcessAcrobat > 0 Or NumberOfTries > 20
'        TerminateResult = ProcessTerminate(ProcessAcrobat)
'
''        If TerminateResult = False Then
''            MsgBox "Could not close all instances of IrfanView automatically." & vbNewLine & vbNewLine & _
''                    "You must close them yourself manually, then please try again.", vbExclamation
''            CloseAllIrfanView = False
''            Exit Function
''        End If
'        ProcessAcrobat = CheckForProcByExe("acrobat.exe")
'        NumberOfTries = NumberOfTries + 1
'    Loop
'
'    If NumberOfTries <= 20 Then
'        CloseAllAcrobat = True
'        Exit Function
'    End If
'
'ErrorClosing:
'    CloseAllAcrobat = False
'    MsgBox "Could not close all instances of Acrobat automatically." & vbNewLine & vbNewLine & _
'            "You must close them yourself manually, then please try again.", vbExclamation
'
'End Function


Private Sub Form_Close()
    DoCmd.Hourglass False
    If left(mstrCalledFrom, 3) = "frm" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
    If Nz(MyFastScan.CoverSheetNum, "NA") <> "NA" Then Call FastScan_UnLock(MyFastScan.CoverSheetNum)
    Call CloseAllViewerWindows(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\")))
    
    If CurrentProject.AllForms("frm_FastScan_IssuesMain").IsLoaded Then
        Forms!frm_FastScan_IssuesMain.TimerInterval = 100
    End If
    
End Sub


Private Sub lstImageType_Change()
    Call FillReasonList
    Call AutoSelectMatchReason(Nz(MyFastScan.rsCoverSheet("NoMatchReasonCd"), ""))
End Sub

Private Sub lstNoMatchReason_Change()

    If lstNoMatchReason.ListIndex <> -1 Then
    
        Dim strSplitIndicator As String
        Dim strDE_ICNIndicator As String
        Dim strDE_PAYERNAMEIndicator As String
        Dim strDE_PROVNUMIndicator As String
        
        'setting flags defaults
        mstrSplitFlag = "N"
        mstrDataEntry_ICN_Flag = "N"
        mstrDataEntry_PayerName_Flag = "N"
        mstrDataEntry_CnlyProvID_Flag = "N"
        mstrDefaultImageType = ""
        mstrRequiredImageType = "N"
        
        
        'Setting defaults NoMatch button caption
        Call AutoSelectNoMatchCaption(Nz(Me.lstNoMatchReason, ""))
        
        'Find the current reason in the list recordset
        Dim bolReasonFound As Boolean
        bolReasonFound = False
        If Not Me.lstNoMatchReason.RecordSet Is Nothing Then
            With Me.lstNoMatchReason.RecordSet
                .MoveFirst
                If Not .EOF Then
                    While Not .EOF And Not bolReasonFound
                        If !ReasonCode = Me.lstNoMatchReason Then
                            mstrDefaultImageType = Nz(!DefaultImageType, "")
                            mstrRequiredImageType = Nz(!RequireDefaultImageType, "N")
                            mstrReasonType = Nz(!ReasonType, "")
                            strSplitIndicator = Nz(!SplitReason, "N")
                            strDE_ICNIndicator = Nz(!DataEntry_ICN, "N")
                            strDE_PAYERNAMEIndicator = Nz(!DataEntry_PayerName, "N")
                            strDE_PROVNUMIndicator = Nz(!DataEntry_PROVNUM, "N")
                            bolReasonFound = True
                        End If
                        .MoveNext
                    Wend
                End If
            End With
        End If
        
        
        
        'Default ImageType indicator column if the user has not selected one already
'        If Nz(lstNoMatchReason.Column(6, lstNoMatchReason.ListIndex), "") <> "" Then
'            mstrDefaultImageType = lstNoMatchReason.Column(6, lstNoMatchReason.ListIndex)
            If lstImageType = "" Then
                lstImageType = mstrDefaultImageType
            End If
'        End If
        
        'Required ImageType indicator column
        'If Nz(lstNoMatchReason.Column(7, lstNoMatchReason.ListIndex), "") <> "" Then
        '    mstrRequiredImageType = Nz(lstNoMatchReason.Column(7, lstNoMatchReason.ListIndex), "N")
            If mstrRequiredImageType = "Y" And Nz(mstrDefaultImageType, "") <> "" Then
                lstImageType = mstrDefaultImageType
            End If
        'End If
        
        'Reason Type column
        'mstrReasonType = Nz(lstNoMatchReason.Column(8, lstNoMatchReason.ListIndex), "")
        
        'Split indicator column
        'If Nz(lstNoMatchReason.Column(2, lstNoMatchReason.ListIndex), "N") = "Y" Then
        If strSplitIndicator = "Y" Then
            cmdNoMatch.Caption = "SPLIT"
            mstrSplitFlag = "Y"
        End If
        
        'DataEntry ICN indicator column
        'If Nz(lstNoMatchReason.Column(3, lstNoMatchReason.ListIndex), "N") = "Y" Then
        If strDE_ICNIndicator = "Y" Then
            mstrDataEntry_ICN_Flag = "Y"
        End If
        
        'DataEntry PayerName indicator column
        'If Nz(lstNoMatchReason.Column(4, lstNoMatchReason.ListIndex), "N") = "Y" Then
        If strDE_PAYERNAMEIndicator = "Y" Then
            mstrDataEntry_PayerName_Flag = "Y"
        End If
        
        'DataEntry ProvNum indicator column
        'If Nz(lstNoMatchReason.Column(5, lstNoMatchReason.ListIndex), "N") = "Y" Then
        If strDE_PROVNUMIndicator = "Y" Then
            mstrDataEntry_CnlyProvID_Flag = "Y"
        End If
        
    End If
    
End Sub

Private Sub FillReasonList()

Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim strDecision As String

    If Me.cmdMatch.Enabled Then
        strDecision = "MATCH"
    Else
        strDecision = "NOMATCH"
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "exec fastscan.usp_FastScan_GetValidMatchCodesForImageType_V2 " & gintAccountID & ", '" & MyFastScan.rsCoverSheet("NoMatchReasonCd") & "', '" & strDecision & "', " & IIf(mbolRejectUser, 1, 0) & ", '" & Nz(Me.lstImageType, "") & "'"
        Set oRs = .ExecuteRS
        Set Me.lstNoMatchReason.RecordSet = oRs
    End With
    
    If Not oRs Is Nothing Then
        If Not (oRs.EOF And oRs.BOF) Then
            oRs.MoveFirst
            While Not oRs.EOF
                oRs.MoveNext
            Wend
        End If
    End If
'Dim oAdo As clsADO
''Dim ors As ADODB.Recordset
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
'        .SQLTextType = SQLtext
'        .sqlString = "select ReasonCode, ReasonDesc, SplitReason, DataEntry_ICN, DataEntry_PayerName, DataEntry_ProvNum, DefaultImageType, RequireDefaultImageType from v_CA_SCANNING_FastScan_NoMatch_Reasons where AccountID = " & gintAccountID & " and ReasonType IN (" & IIf(strDecision = "MATCH", "'MA'", IIf(mbolRejectUser, "'NM','RJ'", "'NM'")) & ") order by reasoncode"
'        .Parameters.Refresh
'        Set Me.lstNoMatchReason.RecordSet = .ExecuteRS
'    End With
    
    Me.lstNoMatchReason = ""
    Me.lstNoMatchReason.Requery
    
    
End Sub

Function ResultsSelected() As Integer
    
    Dim bSelectedCount As Integer
   
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "select Total = isnull(count(1),0) from FastScan.v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ProcessInd = 1 and ResultType in ('M','B')"
        Set oRs = .ExecuteRS
        If .GotData = True Then
            bSelectedCount = oRs("Total")
        Else
            bSelectedCount = 0
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
   
   
'    With Me.subfrm_Results.Form.RecordSet.Clone
'        If .BOF And .EOF Then
'            ResultsSelected = 0
'            Exit Function
'        End If
'        .MoveFirst
'        While Not .EOF
'            If !ProcessInd Then
'                If bSelectedCount = 0 Then 'if this is the first time
'                    bSelectedCount = bSelectedCount + 1
'                Else
'                    'more than one result record selected
'                    bSelectedCount = bSelectedCount + 1
'                    .MoveLast
'                End If
'            End If
'            .MoveNext
'        Wend
'    End With
    
    ResultsSelected = bSelectedCount
    
End Function

Function ResultsOtherSelected(strResultType As String) As Integer
    
    Dim bSelectedCount As Integer
   
    With Me.subfrm_Results_Other.Form.RecordSet.Clone
        If .BOF And .EOF Then
            ResultsOtherSelected = 0
            Exit Function
        End If
        .MoveFirst
        While Not .EOF
            If !ProcessInd And !ResultType = strResultType Then
                If bSelectedCount = 0 Then 'if this is the first time
                    bSelectedCount = bSelectedCount + 1
                Else
                    'more than one result record selected
                    bSelectedCount = bSelectedCount + 1
                    .MoveLast
                End If
            End If
            .MoveNext
        Wend
    End With
    
    ResultsOtherSelected = bSelectedCount
    
End Function

Function ResultsOtherThanSelected(strResultType As String) As Integer
    
    'returns True if there are additional results from type other than the one passed via parameter. This excludes main result types M and B
    
    Dim bSelectedCount As Integer
   
    With Me.subfrm_Results_Other.Form.RecordSet.Clone
        If .BOF And .EOF Then
            ResultsOtherThanSelected = 0
            Exit Function
        End If
        .MoveFirst
        While Not .EOF
            If !ProcessInd And !ResultType <> strResultType And (!ResultType <> "M" And !ResultType <> "B") Then
                If bSelectedCount = 0 Then 'if this is the first time
                    bSelectedCount = bSelectedCount + 1
                Else
                    'more than one result record selected
                    bSelectedCount = bSelectedCount + 1
                    .MoveLast
                End If
            End If
            .MoveNext
        Wend
    End With
    
    ResultsOtherThanSelected = bSelectedCount
    
End Function


Sub GetImageFileExt(ImageFullPath As String)
    If FileExists(ImageFullPath & ".PDF") Then
        mstrImageFileExt = "PDF"
    ElseIf FileExists(ImageFullPath & ".TIF") Then
        mstrImageFileExt = "TIF"
    Else
        mstrImageFileExt = ""
    End If
End Sub

Sub GetImagePageCnt()
    If mstrImageFileExt = "TIF" Then
        'lets count the pages
        mintpageCnt = TifPageCount(mstrImageFullName)
    End If
    
    If mstrImageFileExt = "PDF" Then
        'lets count the pages
        mintpageCnt = Count_PDF_Pages(mstrImageFullName)
    End If
    
End Sub


Sub FillFolderCombo()
    Dim MyAdo As clsADO
    Dim strSQL As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    strSQL = "select FolderName from FastScanMaint.v_FastScan_UserAuthFolders where AccountID = " & gintAccountID & " order by FolderPriority, FolderName"
    
    MyAdo.sqlString = strSQL
    Set Me.cmbProviderFolderWork.RecordSet = MyAdo.OpenRecordSet
    
    MyAdo.DisConnect

    Set MyAdo = Nothing
    
    If Me.cmbProviderFolderWork.ListCount > 0 Then
        Me.cmbProviderFolderWork = Me.cmbProviderFolderWork.ItemData(0)
    End If
    
End Sub


Private Sub cmdPropagateAdd_Click()

    Dim sqlUpdate As String
    Dim ErrMsgTxt As String
    Dim SearchResult As String
    
    If ResultsSelected() <> 1 Then
        ErrMsgTxt = "One results row must be marked with a check first"
        MsgBox ErrMsgTxt, vbExclamation, "Error"
        Exit Sub
    End If
    
    sqlUpdate = "UPDATE v_CA_SCANNING_FastScan_Search SET ResultType = 'P' WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ProcessInd = True"
    CurrentDb.Execute (sqlUpdate)
    
    'when working with propagate we cannot have any related results
    sqlUpdate = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE ResultType = 'R' and AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID
    CurrentDb.Execute (sqlUpdate)
    
    'cmdClearAllResults_Click
    Me.subfrm_Results_Other.Form.Requery
    Me.Refresh
    
End Sub


Private Sub TogPropagate_Click()

On Error GoTo Error_Handler

If TogPropagate.Value = -1 Then
    
    If ResultsOtherThanSelected("P") > 0 Then
        If MsgBox("Are you sure you want to discard the selected claims and start working with Propagated claims?", vbQuestion + vbYesNo, "Discard Selected Claims") = vbNo Then
            Exit Sub
        End If
    End If

    Dim sqlDelete As String
    'delete all propagate results in case there is one
    sqlDelete = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'P'"
    CurrentDb.Execute (sqlDelete)
    'and unselect all related
    Call cmdRelatedUnselectAll_Click

    OtherResultsToggle "Propagate"

Else

    TogPropagate.Value = -1
                                                
End If

DecisionSwitch

Exit_Sub:
    Exit Sub

Error_Handler:

    TogPropagate.Value = 0

    GoTo Exit_Sub
                                                
End Sub


Private Sub TogRelated_Click()

On Error GoTo Error_Handler

If TogRelated.Value = -1 Then
    
    If ResultsOtherThanSelected("R") > 0 Then
        If MsgBox("Are you sure you want to discard the selected claims and start working with Related claims?", vbQuestion + vbYesNo, "Discard Selected Claims") = vbNo Then
            Exit Sub
        Else
            Dim sqlDelete As String
            'delete all propagate results
            sqlDelete = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ProcessInd = 1"
            CurrentDb.Execute (sqlDelete)
            
            'do the search again, we don't know what the user was doing in propagate, better to start off again
            Call cmdSearch_Click
        End If
    End If

    OtherResultsToggle "Related"

'   Call Me.subfrm_Results.Form.ProcessInd_Click

Else

    TogRelated.Value = -1
                                                
End If

DecisionSwitch


Exit_Sub:

 
    Exit Sub

Error_Handler:

    TogRelated.Value = 0

    GoTo Exit_Sub

End Sub


Sub OtherResultsToggle(strOtherType As String)
    If strOtherType = "Propagate" Then
        Me.TogRelated = 0
        Me.TogPropagate = -1
        Me.cmdPropagateAdd.Enabled = True
        Me.cmdPropagateAddAllReq.Enabled = True
        Me.cmdPropagateRemove.Enabled = True
        Me.cmdPropagateRemoveAll.Enabled = True
        Me.cmdRelatedSelectAll.Enabled = False
        Me.cmdRelatedUnselectAll.Enabled = False
        Me.subfrm_Results_Other.SourceObject = "frm_FastScan_SearchResults_Propagate"
        Me.subfrm_Results_Other.Form.RecordSource = "SELECT * " & _
            "From v_CA_SCANNING_FastScan_Search " & _
            "WHERE AccountID = " & gintAccountID & " and (((v_CA_SCANNING_FastScan_Search.[CoverSheetNum])='" & MyFastScan.CoverSheetNum & "') AND ((v_CA_SCANNING_FastScan_Search.[UserID])='" & mstrCurrentUser & "') AND ((v_CA_SCANNING_FastScan_Search.[SessionID])= " & mintSessionID & " ) AND ((v_CA_SCANNING_FastScan_Search.[ResultType])='P')) order by CnlyProvID, ClmStatus, ICN"
    End If
    If strOtherType = "Related" Then
        Me.TogPropagate = 0
        Me.TogRelated = -1
        Me.cmdPropagateAdd.Enabled = False
        Me.cmdPropagateAddAllReq.Enabled = False
        Me.cmdPropagateRemove.Enabled = False
        Me.cmdPropagateRemoveAll.Enabled = False
        Me.cmdRelatedSelectAll.Enabled = True
        Me.cmdRelatedUnselectAll.Enabled = True
        Me.subfrm_Results_Other.SourceObject = "frm_FastScan_SearchResults_Related"
        Me.subfrm_Results_Other.Form.RecordSource = "select * from v_CA_SCANNING_FastScan_search where 1=2"
    End If
End Sub

Private Sub cmdPropagateRemove_Click()
    Dim UserAnswer
    Dim sqlDelete As String
    
    If ResultsOtherSelected("P") > 0 Then
        UserAnswer = MsgBox("Are you sure you want to remove ICN " & Me.subfrm_Results_Other.Form.Icn & " from the Propagate list?", vbQuestion + vbYesNo, "Remove")
        If UserAnswer = vbNo Then
            Exit Sub
        End If
        sqlDelete = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'P' and CnlyClaimNum = '" & Me.subfrm_Results_Other.Form.CnlyClaimNum & "'"
        CurrentDb.Execute (sqlDelete)
    End If
    
    Me.subfrm_Results_Other.Form.Requery
    DecisionSwitch
End Sub

Private Sub cmdPropagateRemoveAll_Click()
    Dim UserAnswer
    Dim sqlDelete As String
    
    If ResultsOtherSelected("P") > 0 Then
        UserAnswer = MsgBox("Are you sure you want to remove All claims from the Propagate list?", vbQuestion + vbYesNo, "Remove All")
        If UserAnswer = vbNo Then
            Exit Sub
        End If
        
        sqlDelete = "DELETE FROM v_CA_SCANNING_FastScan_Search WHERE AccountID = " & gintAccountID & " and CoverSheetNum = '" & MyFastScan.CoverSheetNum & "' and UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'P'"
        CurrentDb.Execute (sqlDelete)
        
        'we need to do the search again
        Call cmdSearch_Click
    End If
    

    
    Me.subfrm_Results_Other.Form.Requery
    DecisionSwitch
End Sub



Private Sub TogShowBarCodeResults_Click()
    Dim UserAnswer As Integer
    If TogShowBarCodeResults.Value = -1 Then
        If Not Me.subfrm_Results.Form.RecordSet Is Nothing Then
            If Me.subfrm_Results.Form.RecordSet.recordCount > 0 Then
                UserAnswer = MsgBox("The search results will be cleared and the Barcode results will be shown instead. Continue?", vbYesNo + vbQuestion, "Show Barcode Results")
                If UserAnswer <> vbYes Then
                    Exit Sub
                End If
            End If
        End If
    End If
    Me.subfrm_Results.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Search as a where AccountID = " & gintAccountID & " and  UserID = '" & mstrCurrentUser & "' and SessionID = " & mintSessionID & " and ResultType = 'B' order by CnlyProvID, ClmStatus, ICN"
    TogShowBarCodeResults.Locked = True
End Sub

Private Sub TogWorkMode_Click()
    If Me.TogWorkMode Then
        Me.TogWorkMode.Caption = "First"
    Else
        Me.TogWorkMode.Caption = "Only"
    End If
End Sub

Private Sub txtICN_AfterUpdate()
Dim ErrMsgTxt As String
    Dim strCnlyClaimNum As String
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    
    If Len(Trim(Me.txtICN)) > 30 Then
       
       
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = myCode_ADO.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "FastScan.usp_FastScan_Process_BarCode"
        cmd.Parameters.Refresh
        cmd.Parameters("@pBarCodeText").Value = Me.txtICN
        
        cmd.Execute
                
        ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
        strCnlyClaimNum = Nz(cmd.Parameters("@pCnlyClaimNum").Value, "")
       
        Me.txtICN = ""
        
        If ErrMsgTxt <> "" Then
            GoTo Error_Handler
        End If
        
        If strCnlyClaimNum = "" Then
            ErrMsgTxt = "Barcode process did not find claim"
            GoTo Error_Handler
        End If
        
        RunSearch iInitialLoad:=0, strCnlyClaimNum:=strCnlyClaimNum
        
    End If

Exit_Sub:
    
Set cmd = Nothing
Set myCode_ADO = Nothing

Exit Sub
    
Error_Handler:
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "Error during Barcode process"
    GoTo Exit_Sub
End Sub

Private Sub CreatePK(ByVal TableName As String, ByVal Fields As String)
    TurnOffDeveloperErrorHandling True
On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
    TurnOffDeveloperErrorHandling False
End Sub

Sub DecisionSwitch()
    'if there are claims selected or in propagate list
    If ResultsSelected() > 0 Or ResultsOtherSelected("P") > 0 Then
        cmdMatch.Enabled = True
        cmdNoMatch.Enabled = False
        Call FillReasonList
        Call AutoSelectMatchReason(Nz(MyFastScan.rsCoverSheet("NoMatchReasonCd"), ""))
    Else
        cmdMatch.Enabled = False
        cmdNoMatch.Enabled = True
        Call FillReasonList
        Call AutoSelectNoMatchCaption(Nz(Me.lstNoMatchReason, ""))
    End If
    lstNoMatchReason_Change
    'lstImageType = ""
End Sub

Sub AutoSelectMatchReason(txtCurrentReasonCd As String)
    Dim strDefaultMatchCode As String
    If Nz(txtCurrentReasonCd, "") = "" Then
        strDefaultMatchCode = "M01"
    Else
        strDefaultMatchCode = "" & DLookup("DefaultMatchCode", "v_FastScan_NoMatch_Reasons", "AccountID = " & gintAccountID & " and ReasonCode = '" & txtCurrentReasonCd & "'")
    End If
    Dim i As Integer
    For i = 0 To Me.lstNoMatchReason.ListCount - 1
        If Me.lstNoMatchReason.Column(0, i) = strDefaultMatchCode Then
            'Me.lstNoMatchReason.Selected(i) = True
            Me.lstNoMatchReason = Me.lstNoMatchReason.ItemData(i)
            lstNoMatchReason_Change
            Exit For
        End If
    Next

End Sub

Sub AutoSelectNoMatchCaption(txtSelectedReasonCd As String)
    Dim strFinishable As String
    If Nz(txtSelectedReasonCd, "") = "" Then
        strFinishable = "N"
    Else
        strFinishable = "" & DLookup("Finishable", "v_FastScan_NoMatch_Reasons", "AccountID = " & gintAccountID & " and ReasonCode = '" & txtSelectedReasonCd & "'")
    End If
    If Nz(strFinishable, "N") = "Y" Then
        If mbolRejectUser Then
            Me.cmdNoMatch.Caption = "Finish"
        Else
            Me.cmdNoMatch.Caption = "NoMatch"
        End If
    Else
        Me.cmdNoMatch.Caption = "NoMatch"
    End If
End Sub

Private Sub myFastScan_FastScanError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub

Private Sub cmdNotes_Click()

Dim noClaimNotes As String
    
    On Error GoTo Err_cmdddNote_Click
    
    Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add

    
    frmGeneralNotes.frmAppID = Me.frmAppID
    Set frmGeneralNotes.NoteRecordSource = MyFastScan.rsNotes
    frmGeneralNotes.RefreshData
    ShowFormAndWait frmGeneralNotes
    'lstTabs_Click
    Set frmGeneralNotes = Nothing

    MyFastScan.SaveCoverSheet


Exit_cmdddNote_Click:
    Exit Sub

Err_cmdddNote_Click:
    MsgBox Err.Description
    Resume Exit_cmdddNote_Click

End Sub

Private Function CloseAllViewerWindows(strWindowTitle As String) As Boolean
    CloseAllViewerWindows = False
    If CloseByProcessTitle(strWindowTitle) Then
        CloseAllViewerWindows = True
    End If
    
    'JS:2015-08-17 doing this because there might be viewer windows with no image file inside. This could be caused by an attempt to open a corrupted image
    Call CloseByProcessTitle("Adobe Acrobat")
    Call CloseByProcessTitle("IrfanView")
    
End Function
