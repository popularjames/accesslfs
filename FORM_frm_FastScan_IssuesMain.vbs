Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Const CstrFrmAppID As String = "FastScanIssues"
'Private miAppPermission As Integer
'Public mbAllowChange, mbAllowAdd, mbAllowView, mbAllowDelete As Boolean
Dim mstrCalledFrom As String
Dim mstrLastChoice As String

Dim mstrMatchINPath As String
Dim mstrMatchOUTPath As String
Dim mstrSplitINPath As String
Dim mstrSplitOUTPath As String
Dim mstrTIFViewerPath As String
Dim mstrAcrobatPath As String

Private WithEvents frmScanningFastScanHistory As Form_frm_FastScan_History
Attribute frmScanningFastScanHistory.VB_VarHelpID = -1

Private Sub cmbFolders_AfterUpdate()
    FillIssueCombo
    FillIssueResults
    If Me.cmbFolders.ListIndex <> -1 Then
        Me.lblFolders.BackColor = vbWhite
        Me.lblFolders.ForeColor = vbRed
    Else
        Me.lblFolders.ForeColor = vbBlack
        Me.lblFolders.BackColor = vbWhite
    End If
    
End Sub

Private Sub cmdClearFolder_Click()
    Me.cmbFolders = Null
    FillIssueCombo
    FillIssueResults
    If Me.cmbFolders.ListIndex <> -1 Then
        Me.lblFolders.BackColor = vbWhite
        Me.lblFolders.ForeColor = vbRed
    Else
        Me.lblFolders.ForeColor = vbBlack
        Me.lblFolders.BackColor = vbWhite
    End If

End Sub


Public Sub cmbIssueType_Click()
    If Me.cmbIssueType <> "No Issues Found" Then
        FillIssueResults
    End If
    If Me.subfrm_IssueResults.Form.RecordSet.recordCount = 0 Then
        FillIssueCombo
    End If
   
End Sub


Private Sub cmdClearSearchFor_Click()
    Me.txtSearchText = ""
End Sub

Private Sub cmdOpen_Click()
    
    If Me.subfrm_IssueResults.Form.RecordSet.recordCount = 0 Then
        Exit Sub
    End If
    
    If Me.subfrm_IssueResults.Form.RecordSet("ProcStatusCd") = "MATCHED" Then
        MsgBox "You cannot open a Matched coversheet, as the image is now attached to the claim.", vbInformation, "Error"
        Exit Sub
    End If
    
    If CurrentProject.AllForms("frm_FastScan_Main").IsLoaded Then
        MsgBox "You can only open one coversheet at a time!" & _
                vbNewLine & vbNewLine & _
                "Close the currently open coversheet before trying to open a new one.", vbExclamation, "Fix No Match"
        Exit Sub
    End If
    
    If Me.subfrm_IssueResults.Form("ProcStatusCd") = "MATCHED" Then
        MsgBox "You cannot open a CoverSheet that was already MATCHED. The Image is now attached to a claim.", vbExclamation, "Fix No Match"
        Exit Sub
    End If
        
    If left(Me.subfrm_IssueResults.Form("ProcStatusCd"), 7) = "DELETED" Then
        MsgBox "You cannot open a CoverSheet that was DELETED.", vbExclamation, "Fix No Match"
        Exit Sub
    End If
    
    If Me.subfrm_IssueResults.Form("ProcStatusCd") = "SPLITCOMPLETE" Then
        MsgBox "This Coversheet was the source of a split. You cannot process it. It will open in Read-Only mode!", vbInformation, "FastScan - Main"
    End If
    

   
    DoCmd.OpenForm "frm_FastScan_Main", acNormal, , , , acNormal, Me.subfrm_IssueResults.Form("CoverSheetNum")
    
    DoEvents
    DoEvents
    

End Sub

Sub cmdRefresh_Click()
        
    Me.txtSearchText = ""
    Me.cmbSearchBy.Value = Me.cmbSearchBy.ItemData(1)
    
    
    mstrLastChoice = Me.cmbIssueType
    RefreshData
'    Me.cmbIssueType = LastChoice
'    If Me.cmbIssueType.ListIndex <> -1 Then
'        cmbIssueType_Click
'    End If
End Sub

Private Sub cmdSearch_Click()

    If Me.cmbSearchBy.ListIndex = -1 Then
        MsgBox "You must first select a criteria to search for", vbExclamation, "Must Select Search Criteria"
        Exit Sub
    End If

    If Nz(Me.txtSearchText, "") = "" Then
        MsgBox "You must enter the text you are searching for", vbExclamation, "Must Enter Search Text"
        Exit Sub
    End If
    
    If Len(Trim(Nz(Me.txtSearchText, ""))) < 4 Then
        MsgBox "Text being searched must be at least 4 characters long.", vbExclamation, "Must Enter Valid Search Text"
        Exit Sub
    End If

    Me.cmbIssueType = ""
    
    Me.subfrm_IssueResults.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Issues where 1=2"

    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = StoredProc
        .sqlString = "FastScan.usp_FastScan_CoverSheetIssuesSearch_v3"
        .Parameters.Refresh
        .Parameters("@pAccountID") = gintAccountID
        .Parameters("@pSearchBy") = Me.cmbSearchBy
        .Parameters("@pSearchForText") = Me.txtSearchText
        .Parameters("@pUserID") = Identity.UserName
       
        Set oRs = .ExecuteRS

        ErrorReturned = Nz(.Parameters("@pErrMsg").Value, "")
        If ErrorReturned <> "" Then
            MsgBox ErrorReturned, vbExclamation
            Exit Sub
        End If
'        If .GotData = False Then
'            MsgBox "No Coversheet Found.", vbExclamation
'        End If
        Me.subfrm_IssueResults.Form.RecordSource = "select * from v_CA_SCANNING_FastScan_issues where userid = '" & Identity.UserName & "'"
      End With
      
End Sub


Private Sub cmdViewHistory_Click()

    If Me.subfrm_IssueResults.Form.RecordSet.recordCount = 0 Then
        Exit Sub
    End If

   If frmScanningFastScanHistory Is Nothing Then
        Set frmScanningFastScanHistory = New Form_frm_FastScan_History
        If SysCmd(acSysCmdGetObjectState, acForm, "frm_FastScan_History") Then
            ColObjectInstances.Add frmScanningFastScanHistory.hwnd & ""
            frmScanningFastScanHistory.OpenCoverSheetNum = Me.subfrm_IssueResults.Form("CoverSheetNum")
            frmScanningFastScanHistory.RefreshScreen
            ShowFormAndWait frmScanningFastScanHistory
        End If
        Set frmScanningFastScanHistory = Nothing
    Else
        frmScanningFastScanHistory.SetFocus
    End If
End Sub

Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
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
        .sqlString = "select distinct FolderName from FastScanMaint.v_CA_SCANNING_FastScan_Folders where accountid = " & gintAccountID & " order by FolderName"
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
    
    If Me.cmbFolders.ListIndex <> -1 Then
        Me.lblFolders.BackColor = vbWhite
        Me.lblFolders.ForeColor = vbRed
    Else
        Me.lblFolders.ForeColor = vbBlack
        Me.lblFolders.BackColor = vbWhite
    End If
    
    mstrLastChoice = ""
    
'    Me.Caption = "FastScan Issues"
'    'Me.RecordSource = ""
'
'    Me.frmAppID = CstrFrmAppID
'
'    Call Account_Check(Me)
'    miAppPermission = UserAccess_Check(Me)
'    If miAppPermission = 0 Then Exit Sub
'
'    miAppPermission = GetAppPermission(Me.frmAppID)
'    mbAllowChange = (miAppPermission And gcAllowChange)
'    mbAllowAdd = (miAppPermission And gcAllowAdd)
'    mbAllowView = (miAppPermission And gcAllowView)
'    mbAllowDelete = (miAppPermission And gcAllowDelete)

    RefreshData
    
Cleanup:
    oAdo.DisConnect
    Set oAdo = Nothing
    Set oRs = Nothing
    
End Sub

Private Sub FillIssueCombo()
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = StoredProc
        .sqlString = "FastScan.usp_FastScan_CoverSheetIssues_v5"
        .Parameters.Refresh
        .Parameters("@pAccountID") = gintAccountID
        .Parameters("@pProviderFolder") = Me.cmbFolders
        .Parameters("@pIssueType") = "" 'Summary
        .Parameters("@pHideOldFinished") = Abs(Me.optHideOldFinished)
        .Parameters("@pUserID") = Identity.UserName
        Set oRs = .ExecuteRS

        ErrorReturned = Nz(.Parameters("@pErrMsg").Value, "")
        If ErrorReturned <> "" Then
            'LogMessage strProcName, "ERROR", "Problem searching for a concept", "Keyword: " & Nz(Me.txtSearchBox, "") & " Expand Search: " & IIf(Me.ckExpandSearch, 1, 0) & " Include Codes: " & IIf(Me.ckIncludeCodes, 1, 0)
            MsgBox ErrorReturned, vbExclamation
            Exit Sub
        End If
        If .GotData = False Then
            MsgBox "No Issues Found.", vbExclamation
            Me.cmbIssueType.RowSource = "Select ""No Issues Found"""
            Exit Sub
        End If
        Set Me.cmbIssueType.RecordSet = oRs
        Me.cmbIssueType = Me.cmbIssueType.ItemData(1)
        
        'Me.cmbIssueType.Recordset.Requery
    End With
End Sub

Private Sub FillIssueResults()

On Error GoTo Error_Handler

    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    DoCmd.Hourglass True
    
    Me.subfrm_IssueResults.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Issues where 1=2"
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = StoredProc
        .sqlString = "FastScan.usp_FastScan_CoverSheetIssues_v5"
        .Parameters.Refresh
        .Parameters("@pAccountID") = gintAccountID
        .Parameters("@pProviderFolder") = Me.cmbFolders  'Me.cmbIssueType.Column(0, Me.cmbIssueType.ListIndex + 1)
        .Parameters("@pIssueType") = Me.cmbIssueType
        .Parameters("@pHideOldFinished") = Abs(Me.optHideOldFinished)
        .Parameters("@pUserID") = Identity.UserName
        Set oRs = .ExecuteRS

        ErrorReturned = Nz(.Parameters("@pErrMsg").Value, "")
        If ErrorReturned <> "" Then
            MsgBox ErrorReturned, vbExclamation
            Exit Sub
        End If
'        If .GotData = False Then
'            MsgBox "No Issues Found.", vbExclamation
'            Me.cmbIssueType.RowSource = "Select ""No Issues Found"""
'            'GoTo CleanUpAndExit
'        End If
        Me.subfrm_IssueResults.Form.RecordSource = "Select * from v_CA_SCANNING_FastScan_Issues where UserID = '" & Identity.UserName & "' order by receiveddt, scanneddt"
       
        'Me.cmbIssueType.Recordset.Requery
    End With

GoTo CleanupAndExit

Error_Handler:


CleanupAndExit:
DoCmd.Hourglass False
Set oRs = Nothing
Set oAdo = Nothing


End Sub

Private Sub FillFoldersCombo()

    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct FolderName from FastScanMaint.v_CA_SCANNING_FastScan_Folders where accountid = " & gintAccountID & " order by FolderName"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            MsgBox "There are no FastScan folders setup for this account! Cannot continue.", vbInformation, "Error: FastScan folders missing"
            DoCmd.Close acForm, Me.Name
            Exit Sub
        End If
        'Me.cmbFoldersType = Me.cmbIssueType.ItemData(1)
    End With
    
    Set Me.cmbFolders.RecordSet = oRs
    
End Sub

Private Sub RefreshData()
    If mstrLastChoice = "" Then
        FillIssueCombo
        FillIssueResults
        FillFoldersCombo
    Else
        Me.cmbIssueType = mstrLastChoice
        cmbIssueType_Click
    End If
End Sub

Private Sub Form_Timer()

    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim ErrorReturned As String
    Dim iResult As Integer
    Dim lngPos As Long
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
        .SQLTextType = StoredProc
        .sqlString = "FastScan.usp_FastScan_CoverSheetIssuesResultsUpdate"
        .Parameters.Refresh
        .Parameters("@pUserID") = Identity.UserName
        Set oRs = .ExecuteRS
        iResult = .Parameters("@RETURN_VALUE")
        ErrorReturned = Nz(.Parameters("@pErrMsg").Value, "")
        If ErrorReturned <> "" Then
            MsgBox ErrorReturned, vbExclamation
            Exit Sub
        End If
      End With
      
    'Call Forms("frm_FastScan_IssuesMain").Refresh
    
    lngPos = Me.subfrm_IssueResults.Form.RecordSet.AbsolutePosition
    Me.subfrm_IssueResults.Form.Requery
    If lngPos < Me.subfrm_IssueResults.Form.RecordSet.recordCount Then
        Me.subfrm_IssueResults.Form.RecordSet.AbsolutePosition = lngPos
    End If
    If Me.subfrm_IssueResults.Form.RecordSet.recordCount = 0 Then
        cmdRefresh_Click
    End If
    
    Set oAdo = Nothing
    Set oRs = Nothing
    
    
    Me.TimerInterval = 0
End Sub

Private Sub optHideOldFinished_Click()
    RefreshData
End Sub
