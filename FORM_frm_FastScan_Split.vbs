Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrNumberOfSplits As Integer
Private mstrFileExt As String
Private mstrSplitPath As String
Private mstrImageFullName As String
Private mstrReasonCd As String

Public Event UpdateReferences(ErrorCode As String, NewImageType As String, NewPageCount As Integer, Comment As String)

Const CstrFrmAppID As String = "FastScanSplit"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let FileExt(data As String)
     mstrFileExt = data
End Property

Property Let SplitPath(data As String)
     mstrSplitPath = data
End Property

Property Let ImageFullName(data As String)
     mstrImageFullName = data
End Property

Property Let ReasonCd(data As String)
     mstrReasonCd = data
End Property

Property Get ReasonCd() As String
     ReasonCd = mstrReasonCd
End Property


Public Sub RefreshScreen()

    Dim strError As String
    Dim i As Integer
    On Error GoTo ErrHandler
    
    Me.subfrm_FastScan_Splits_Worktable.Form.RecordSource = "SELECT * FROM v_CA_Scanning_FastScan_Splits_Worktable where 1=2"
    Me.Requery
    
    'clear the split start page combo
    For i = 0 To Me.cmbSplitStartPg.ListCount - 1
        Me.cmbSplitStartPg.RemoveItem 0
    Next i
    
'    'fill repeat pages combo
'    For i = 0 To Me.txtPageCnt - 2
'        Me.cmbRepeatPages.AddItem i
'    Next
'
'    Me.cmbRepeatPages = 0

    If Not Split_Recalc("FIRST", 0, Me.txtPageCnt) Then
        Forms!frm_FastScan_Main.SaveSplit = "N"
        DoCmd.Close acForm, Me.Name
    End If

    Me.subfrm_FastScan_Splits_Worktable.Form.RecordSource = "SELECT * FROM v_CA_SCANNING_FastScan_Splits_Worktable where CoverSheetNum = '" & Me.txtCoverSheetNum & "' and UserID = '" & Identity.UserName & "' and SplitNumber > 0 order by SplitNumber "
    'Me.subfrm_FastScan_Splits_Worktable.Form.Requery

    Me.cmbSplitStartPg = "--"

    Me.cmbRepeatSplit.Enabled = False
    Me.cmbSplitStartPg.Enabled = True
    Me.cmdAddSplit.Enabled = True
    Me.cmdDelSplit.Enabled = False
    Me.chkRepeatSplit.Enabled = False
    
    
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
    
End Sub



Private Sub chkRepeatSplit_AfterUpdate()
    If Me.chkRepeatSplit = vbTrue Then
        Me.cmbRepeatSplit.Enabled = True
        Me.cmbRepeatSplit = ""
    Else
        Me.cmbRepeatSplit = ""
        Me.cmbRepeatSplit.Enabled = False
    End If
End Sub

Private Sub cmdRestart_Click()
    Dim UserAnswer As Integer
    UserAnswer = MsgBox("Are you sure you want to reset? All entered splits will be erased", vbQuestion + vbYesNo, "Split")
    If UserAnswer = vbYes Then
        RefreshScreen
    End If
End Sub

'Private Sub cmdSaveRepeat_Click()
'
'    If Not Split_Recalc("FIRST", Me.cmbRepeatPages.value, Me.txtPageCnt) Then
'        Forms!frm_FastScan_Main.SaveSplit = "N"
'        DoCmd.Close acForm, Me.name
'    End If
'
'    Me.subfrm_FastScan_Splits_Worktable.Form.RecordSource = "SELECT * FROM v_CA_SCANNING_FastScan_Splits_Worktable where CoverSheetNum = '" & Me.txtCoverSheetNum & "' and UserID = '" & Identity.username & "'"
'    'Me.subfrm_FastScan_Splits_Worktable.Form.Requery
'
'    Me.cmbRepeatPages.Enabled = False
'    Me.cmbSplitStartPg.Enabled = True
'    Me.cmdAddSplit.Enabled = True
'
'End Sub


Private Sub cmdAddSplit_Click()

    If Me.cmbSplitStartPg.Value = "--" Then
        MsgBox "You must select a starting page for the next split first!", vbInformation, "Split"
        Exit Sub
    End If

    If Not Split_Recalc("ADD", Me.cmbSplitStartPg.Value, Me.txtPageCnt) Then
        Forms!frm_FastScan_Main.SaveSplit = "N"
        DoCmd.Close acForm, Me.Name
    End If
    Me.subfrm_FastScan_Splits_Worktable.Form.Requery
    Me.cmdDelSplit.Enabled = True

    If Not Me.chkRepeatSplit.Enabled Then
        Me.chkRepeatSplit.Enabled = True
        Me.chkRepeatSplit = vbFalse
        Me.cmbRepeatSplit = ""
    End If
    
    
End Sub

Private Sub cmdDelSplit_Click()

    If Not Split_Recalc("DELETE", 0, Me.txtPageCnt) Then
        Forms!frm_FastScan_Main.SaveSplit = "N"
        DoCmd.Close acForm, Me.Name
    End If

    Me.subfrm_FastScan_Splits_Worktable.Form.Requery

    Me.cmbSplitStartPg = "--"

End Sub

Private Sub CmdCancel_Click()
    If MsgBox("Are you sure you want to cancel and leave the Split module?", vbYesNo, "Cancel Splits") = vbYes Then
        DoCmd.Close acForm, Me.Name
    End If
End Sub



Private Sub cmdOk_Click()
    
    'RaiseEvent UpdateReferences(Me.ErrorCode, Me.NewImageType, val(Me.NewPageCount), Me.Comment)
    
    If Me.chkRepeatSplit = vbTrue And Me.cmbRepeatSplit = "" Then
        MsgBox "You must choose a Split number to repeat if you selected the Insert Split checkbox.", vbInformation, "Split"
        Exit Sub
    End If
    
    If mstrNumberOfSplits < 2 Then
        MsgBox "There is not a valid number of splits for this image, cannot continue.", vbInformation, "Split"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to save the Split definitions?", vbYesNo, "Split") = vbYes Then
        Call ProcessSplit
        DoCmd.Close acForm, Me.Name
    End If
    

    
End Sub

Function Split_Recalc(pstrAction As String, pintPage As Integer, pintPageCnt As Integer) As Boolean

    Dim ErrMsgTxt As String
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    Dim i As Integer
    Dim intnextAvailSplitStart
    
    Split_Recalc = False
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    'save split data and submit for conversion if TIF
    myCode_ADO.BeginTrans
    
    'create row in scanning_image_log_tmp table and mark image as matched in FastScan table
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_Splits_Recalc"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheetNum").Value = Me.txtCoverSheetNum
    cmd.Parameters("@pAction").Value = pstrAction
    cmd.Parameters("@pPage").Value = pintPage
    cmd.Parameters("@pPageCnt").Value = pintPageCnt
    cmd.Parameters("@pUserID").Value = Identity.UserName
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    intnextAvailSplitStart = cmd.Parameters("@pNextAvailSplitStart").Value
    mstrNumberOfSplits = cmd.Parameters("@pNumberOfSplits").Value
    
    If ErrMsgTxt <> "" Then
        GoTo Error_Handler
    End If
    
'    If Not DeleteFile(mstrImageFullName, False) Then
'        ErrMsgTxt = "Could not delete original file to be split. Split Failed"
'        GoTo Error_Handler
'    End If
    
    myCode_ADO.CommitTrans
    
'    If CInt(Me.cmbRepeatPages) + 2 = intnextAvailSplitStart Then
'        Me.cmdDelSplit.Enabled = False
'    End If
    
    If mstrNumberOfSplits = 1 Then '2 = intnextAvailSplitStart
        Me.cmdDelSplit.Enabled = False
        Me.chkRepeatSplit = vbFalse
        Me.chkRepeatSplit.Enabled = False
        Me.cmbRepeatSplit = ""
        Me.cmbRepeatSplit.Enabled = False
    Else
        Dim tmpcmbRepeatSplit As String
        tmpcmbRepeatSplit = Me.cmbRepeatSplit
        For i = 1 To Me.cmbRepeatSplit.ListCount
            Me.cmbRepeatSplit.RemoveItem 0
        Next
        For i = 1 To mstrNumberOfSplits
            Me.cmbRepeatSplit.AddItem i
        Next
        If tmpcmbRepeatSplit <> "" Then
            If CInt(tmpcmbRepeatSplit) <= mstrNumberOfSplits Then
                Me.cmbRepeatSplit = tmpcmbRepeatSplit
            End If
        End If
    End If
    
    For i = 0 To Me.cmbSplitStartPg.ListCount - 1
        Me.cmbSplitStartPg.RemoveItem 0
    Next i

    If intnextAvailSplitStart > CInt(Me.txtPageCnt) Then
        Me.cmdAddSplit.Enabled = False
    Else
        Me.cmdAddSplit.Enabled = True
        For i = intnextAvailSplitStart To Me.txtPageCnt
            Me.cmbSplitStartPg.AddItem i
        Next
    End If
    
    Me.cmbSplitStartPg = "--"
    
    Split_Recalc = True

    GoTo Clean_And_Exit

Error_Handler:
    
    myCode_ADO.RollbackTrans
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "FastScan_Split: Error during Split process"
    Split_Recalc = False
    
Clean_And_Exit:

    DoCmd.Hourglass False

    Set cmd = Nothing
    Set myCode_ADO = Nothing

End Function

Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    
    'For testing only
'            Me.txtCoverSheetNum = "1234567890"
'            Me.txtFileName = "filename"
'            Me.txtPageCnt = 100
'            Me.SplitPath = "thepath"
'            Me.FileExt = "ext"
'            Me.ImageFullName = "imagename"
'            Me.ReasonCd = "05"
            
    
    Me.Caption = "FastScan Split"
    
    Dim iAppPermission As Integer

    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    Me.subfrm_FastScan_Splits_Worktable.Form.RecordSource = ""
   
    CreatePK "v_CA_SCANNING_FastScan_Splits_Worktable", "CoverSheetNum, UserID, SplitNumber"
    
    'For testing only
'           RefreshScreen
    
End Sub

Sub ProcessSplit()

    Dim ErrMsgTxt As String
    
On Error GoTo Error_Handler

    DoCmd.Hourglass True
    
'    If mstrFileExt = "PDF" Then
'        If Not Forms!frm_FastScan_Main.CloseAllAcrobat Then
'            ErrMsgTxt = "Could not close the image viewer. Cannot Continue." & mstrSplitPath & Me.txtFileName & ". Cannot Split"
'            GoTo Error_Handler
'        End If
'    ElseIf mstrFileExt = "TIF" Then
'        If Not Forms!frm_FastScan_Main.CloseAllIrfanView Then
'            ErrMsgTxt = "Could not close the instance of IrfanView. Cannot Continue." & mstrSplitPath & Me.txtFileName & ". Cannot Split"
'        End If
'    End If
    
    If Not CloseByProcessTitle(Right(mstrImageFullName, Len(mstrImageFullName) - InStrRev(mstrImageFullName, "\"))) Then
        ErrMsgTxt = "Could not close the image viewer. Cannot Continue." & mstrSplitPath & Me.TxtFileName & ". Cannot Split"
        GoTo Error_Handler
    End If

    If Not FolderExists(mstrSplitPath) Then
        CreateFolder (mstrSplitPath)
        If Not FolderExists(mstrSplitPath) Then
            ErrMsgTxt = "FastScan_Split ERROR: Cannot create split folder " & mstrSplitPath & "."
            GoTo Error_Handler
        End If
    End If
    
    'move file to split folder
    If Not CopyFile(mstrImageFullName, mstrSplitPath, False) Then
        ErrMsgTxt = "Could not copy image file to its destination" & mstrSplitPath & Me.TxtFileName & ". Cannot Split"
        GoTo Error_Handler
    End If
    
    DoCmd.Hourglass False
    
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    'save split data and submit for conversion if TIF
    myCode_ADO.BeginTrans
    
    'create row in scanning_image_log_tmp table and mark image as matched in FastScan table
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "FastScan.usp_FastScan_ProcessSplit_v2"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCoverSheetNum").Value = Me.txtCoverSheetNum
    cmd.Parameters("@pFileExt").Value = mstrFileExt
    cmd.Parameters("@pReasonCd").Value = Me.ReasonCd
    cmd.Parameters("@pRepeatSplit").Value = IIf(Me.cmbRepeatSplit = "", Null, Me.cmbRepeatSplit)
    cmd.Parameters("@pUserID").Value = Identity.UserName
    cmd.Execute
            
    ErrMsgTxt = Trim(cmd.Parameters("@pErrMsg").Value) & ""
    If ErrMsgTxt <> "" Then
        GoTo Error_Handler
    End If
    
'    If Not DeleteFile(mstrImageFullName, False) Then
'        ErrMsgTxt = "Could not delete original file to be split. Split Failed"
'        GoTo Error_Handler
'    End If
    
    myCode_ADO.CommitTrans
    
    Forms!frm_FastScan_Main.SaveSplit = "Y"
    
    GoTo Clean_And_Exit
                                
Error_Handler:
    
    myCode_ADO.RollbackTrans
    If Nz(ErrMsgTxt, "") = "" Then ErrMsgTxt = Err.Description
    MsgBox ErrMsgTxt, vbExclamation, "FastScan_Split: Error during Split process"
    
Clean_And_Exit:

    DoCmd.Hourglass False
    
    Set cmd = Nothing
    Set myCode_ADO = Nothing
    
End Sub

Private Sub CreatePK(ByVal TableName As String, ByVal Fields As String)
    TurnOffDeveloperErrorHandling True
On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
    TurnOffDeveloperErrorHandling False
End Sub
