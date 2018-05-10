Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'' 8/15/2014: KD Need to make sure this is imported to the Master Claim Admin
Private coSettings As clsSettings

Private coRightClickListItem As ListItem
Private csListViewNameClicked As String
Public gintAccountID As Integer
Private cstrAccountListSproc As String

Private cbSpecificDateSelected As Boolean

Private coLVColPos As clsLVColumnPositions
Private WithEvents oMainErrorGrid As Form_frm_GENERAL_Datasheet
Attribute oMainErrorGrid.VB_VarHelpID = -1
Private WithEvents oDetailErrorGrid As Form_frm_GENERAL_Datasheet
Attribute oDetailErrorGrid.VB_VarHelpID = -1
Private WithEvents oClaimsGrid As Form_frm_GENERAL_Datasheet_ADO
Attribute oClaimsGrid.VB_VarHelpID = -1
Private Const cstrFormAppId As String = "CLAIM"


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = cstrFormAppId
End Property

Public Property Get QueueColumns() As clsLVColumnPositions
    If coLVColPos Is Nothing Then
        Set coLVColPos = New clsLVColumnPositions
        Call coLVColPos.SetDetails("QUEUE", Me.lvQueue)
        Call coLVColPos.SetDetails("QUEUEERRORS", Me.lvQErrorDetails)
        Call coLVColPos.SetDetails("GENERATE", Me.lvGenerate)
        Call coLVColPos.SetDetails("GENERATEERRORS", Me.lvGenerateErrs)
        Call coLVColPos.SetDetails("OUTPUT", Me.lvOutput)
        Call coLVColPos.SetDetails("OUTPUTERRORS", Me.lvOutputErrors)
    End If
    Set QueueColumns = coLVColPos
End Property

Public Property Get AccountListSproc() As String
    AccountListSproc = cstrAccountListSproc
End Property
Public Property Let AccountListSproc(strAccountListSproc As String)
    cstrAccountListSproc = strAccountListSproc
End Property


Public Property Get SelectedAccountId() As Long
    SelectedAccountId = gintAccountID
    GlobalSelectedAccountId = gintAccountID
End Property
Public Property Let SelectedAccountId(intAccountId As Long)
    gintAccountID = intAccountId
    
    GlobalSelectedAccountId = gintAccountID
End Property

Public Property Get RightClickedListItem() As ListItem
    Set RightClickedListItem = coRightClickListItem
End Property
Public Property Let RightClickedListItem(oLI As ListItem)
    Set coRightClickListItem = oLI
End Property

Public Property Get ListViewNameClicked() As String
    ListViewNameClicked = csListViewNameClicked
End Property
Public Property Let ListViewNameClicked(sListViewNameClicked As String)
    csListViewNameClicked = sListViewNameClicked
End Property


Public Property Get DaysLeftToSendForHighlight() As Integer
    If coSettings Is Nothing Then Set coSettings = New clsSettings
    
    DaysLeftToSendForHighlight = coSettings.GetSetting("ClaimClockFlagNumDays")
End Property

Private Sub PopulateAccountList()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".PopulateAccountList"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AccountList"
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Could not retrieve Account list", .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        End If
    End With
    
    
    Me.cmbAccountId.ColumnCount = 2
    Me.cmbAccountId.BoundColumn = 1
    Me.cmbAccountId.ColumnWidths = "0;2"
    Set Me.cmbAccountId.RecordSet = oRs
    
    
Block_Exit:
    Set oAdo = Nothing
    Set oRs = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim dtLastRunTime As Date


    strProcName = ClassName & ".RefreshData"
    DoCmd.Hourglass True

    '' Populate the AccountId combo
    Call PopulateAccountList
    

    Call RefreshQueue
    
    
    
    Me.txtTtlClaimsInQ = TotalListViewRow(Me.lvQueue, "ClaimCount")
    Me.txtTtlLettersInQ = TotalListViewRow(Me.lvQueue, "LetterCount")

    dtLastRunTime = GetLastLoadQueueRunTime

    Me.txtLastQueueRunTime = dtLastRunTime
    Me.txtNextQueRunTime = DateAdd("h", 1, dtLastRunTime)
    Me.txtQueueStatus = GetProcessorState()
    
    txtTotalQueieIDCount = GetTotalLettersForMaxQueueID(Me.SelectedAccountId)
    
    
    
    '' Now get the queue error details:
    Call RefreshQueueErrors
    Set coLVColPos = New clsLVColumnPositions
    Call coLVColPos.SetDetails("Queue", Me.lvQueue)
    Call coLVColPos.SetDetails("QueueErrors", Me.lvQErrorDetails)
    
    
    Call RefreshGenerate
    
    Me.txtGenerationLetterCnt = TotalListViewRow(Me.lvGenerate, "LetterCount")
    Me.txtGenerationClaimCnt = TotalListViewRow(Me.lvGenerate, "ClaimCount")
    
    
    Call RefreshGenerateErrors
    Call coLVColPos.SetDetails("Generate", Me.lvGenerate)
    Call coLVColPos.SetDetails("GenerateErrors", Me.lvGenerateErrs)
    
    
    
    Call RefreshOutput
    
    Call RefreshOutputErrors
    Call coLVColPos.SetDetails("Output", Me.lvOutput)
    Call coLVColPos.SetDetails("OutputErrors", Me.lvOutputErrors)
    
        
    
    Me.lblQueueItemCount.Caption = CStr(Me.lvQueue.ListItems.Count)
    Me.lblErrorQueueItemCount.Caption = CStr(Me.lvQErrorDetails.ListItems.Count)
    lblGenerateItemCount.Caption = CStr(Me.lvGenerate.ListItems.Count)
    lblErrorGenerateItemCount.Caption = CStr(Me.lvGenerateErrs.ListItems.Count)
    lblOutputItemCount.Caption = CStr(Me.lvOutput.ListItems.Count)
    lblErrorOutputItemCount.Caption = CStr(Me.lvOutputErrors.ListItems.Count)

    ' Generation Tab:
    Call GetGenerationDetails42Day(Me.Form)
    
    
    ' Output Tab:
    Me.txtOutBatchCount = TotalListViewRow(Me.lvOutput, "LetterType")  ' Note: this should only ever be 1 or 2 for each letter type
    Me.txtOutLetterCount = TotalListViewRow(Me.lvOutput, "LetterCount")
    Me.txtOutClaimCount = TotalListViewRow(Me.lvOutput, "ClaimCount")
    Me.txtOutPageCount = TotalListViewRow(Me.lvOutput, "Pages")
    
    Call GetOutputTodayDetails42Day(Me.Form)

    Call RefreshLoadQueueErrors
    
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    DoCmd.Hourglass False
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Sub RefreshLoadQueueErrors()
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String

    strProcName = ClassName & ".RefreshLoadQueueErrors"
    If oMainErrorGrid Is Nothing Then
        Set oMainErrorGrid = Me.sfrmMainErrorGrid.Form
    End If
    
    ' Orig
'    sSql = "SELECT PrintQueueRunId, Count(RelatedCnlyClaimNum) As ClmCount, LetterType, LetterDesc, ErrorMsg, AccountId FROM v_LETTER_AUtomation_AddQueueErrorLog " & _
'        " WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# " & _
'        " GROUP BY PrintQueueRunId, AccountId, LetterType, LetterDesc, ErrorMsg, AccountId "
    If Me.cmbAccountId <> 0 Then
'        sSql = "SELECT PrintQueueRunId, Count(RelatedCnlyClaimNum) As ClmCount, LetterType, LetterDesc, ErrorMsg, AccountId FROM v_LETTER_AUtomation_AddQueueErrorLog " & _
'            " WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# AND AccountId = " & CStr(Me.cmbAccountId) & _
'            " GROUP BY PrintQueueRunId, AccountId, LetterType, LetterDesc, ErrorMsg, AccountId "
        sSql = "SELECT Count(RelatedCnlyClaimNum) As ClmCount, COUNT(CnlyProvId) AS LetterCount, LetterType, LetterDesc, ErrorMsg, AccountId, ErrorTypeId FROM v_LETTER_AUtomation_AddQueueErrorLog " & _
            " WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# AND AccountId = " & CStr(Me.cmbAccountId) & _
            " GROUP BY AccountId, LetterType, LetterDesc, ErrorMsg, AccountId, ErrorTypeId "
        
        sSql = "SELECT * FROM v_LETTER_AUtomation_AddQueueErrorLogLetterCnt v WHERE v.AccountId = " & CStr(Me.cmbAccountId) & " AND v.ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# "
        
    Else
'        sSql = "SELECT PrintQueueRunId, Count(RelatedCnlyClaimNum) As ClmCount, LetterType, LetterDesc, ErrorMsg, AccountId FROM v_LETTER_AUtomation_AddQueueErrorLog " & _
'            " WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# " & _
'            " GROUP BY PrintQueueRunId, AccountId, LetterType, LetterDesc, ErrorMsg, AccountId "
'
        sSql = "SELECT Count(RelatedCnlyClaimNum) As ClmCount, COUNT( CnlyProvId) AS LetterCount, LetterType, LetterDesc, ErrorMsg, AccountId, ErrorTypeId FROM v_LETTER_AUtomation_AddQueueErrorLog " & _
            " WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# " & _
            " GROUP BY AccountId, LetterType, LetterDesc, ErrorMsg, AccountId, ErrorTypeId "
    
        sSql = "SELECT * FROM v_LETTER_Automation_AddQueueErrorLogLetterCnt v WHERE v.ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# "
    

    
    End If



    '  usp_LETTER_Automation_RefreshLoadQueueErrors
    Set oDb = CurrentDb()
        ' THis guy is a LONG one.. We need to plop this in a stored proc sooner rather than later!
        ' that's not the entire problem - the query needs to be tuned - very badly written (shame on you Kev!)
    oMainErrorGrid.InitData sSql, 2
    
    '' Now I want to add a subdatasheet
    sSql = "SELECT * FROM v_LETTER_AUtomation_AddQueueErrorLog WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & "# "


    If oDetailErrorGrid Is Nothing Then
        Set oDetailErrorGrid = Me.sfrmDetailErrorGrid.Form
    End If

'    Call oSubFrm.SubDataSheetInit(sSql, 2, "PrintQueueRunId;LetterType", "PrintQueueRunId;LetterType")
    'oSubFrm. InitData sSql, 2
    
'    Call oSubFrm.LinkFieldsToSubDatasheet("PrintQueueRunId, LetterType", "PrintQueueRunId, LetterType")
    
    'me.sfrm_LoadPrintQueueErrors.subData
'    Me.frm_GENERAL_Datasheet.Form.InitData strRowSource, 2
'    Me.frm_GENERAL_Datasheet.Form.RecordSource = strRowSource
    
Block_Exit:

    DoCmd.Hourglass False
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub RefreshQueue()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command

    strProcName = ClassName & ".RefreshQueue"

    Set oAdo = New clsADO
    With oAdo
'        .ConnectionString = GetConnectString("CMS_AUDITORS_CODE")
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintQueue_CurrentQueue"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        Set oRs = .ExecuteRS
    End With
    
    
    Call PopulateListView(Me.lvQueue, oRs)
   
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_OPS_Filters"
        .Parameters.Refresh
        .Parameters("@pFilterName") = "Letter Type"
        Set oRs = .ExecuteRS
        If .GotData = True Then

            Call RefreshComboBoxFromRecordset(oRs, Me.cmbFltrLetterType)
        End If
    End With
    
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If

    Call HighlightTimeSensitiveQueue
   
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub RefreshQueueErrors()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshQueueErrors"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintQueue_Errors"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        Set oRs = .ExecuteRS
    End With

    Call PopulateListView(Me.lvQErrorDetails, oRs)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub RefreshGenerate()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshGenerate"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintQueue_Generate"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        Set oRs = .ExecuteRS
    End With

    Call PopulateListView(Me.lvGenerate, oRs)
   
   ' txtGenerationLetterCnt
   ' txtGenerationClaimCnt
   
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub RefreshGenerateErrors()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshGenerateErrors"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintGenerate_Errors"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        Set oRs = .ExecuteRS
    End With

    Call PopulateListView(Me.lvGenerateErrs, oRs)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub RefreshOutput()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim dtDummy As Date

    strProcName = ClassName & ".RefreshOutput"

    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintOutput"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        
        .Parameters("@pRange") = Me.cmbRange
        .Parameters("@pShowCompleted") = Nz(fraShowCompleted.Value, 0)
        If cbSpecificDateSelected = True Then
            .Parameters("@pSpecificDate") = CDate(Format(Me.dtSpecificDate, "mm/dd/yyyy"))
        End If
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With

    Call PopulateListView(Me.lvOutput, oRs)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub RefreshOutputErrors()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshOutputErrors"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintOutput_Errors"
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        Set oRs = .ExecuteRS
    End With

    Call PopulateListView(Me.lvOutputErrors, oRs)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Function HighlightTimeSensitiveQueue() As Integer
On Error GoTo Block_Err
Dim oListV As ListView
Dim strProcName As String
Dim iRet As Integer
Dim iOurCol As Integer
Dim oCHdr As ColumnHeader
Dim oLI As ListItem
Dim bFoundCol As Boolean
Const sCOLUMNNAME As String = "ClaimCountdownClock"
Dim iSubIdx As Integer
Dim oLetterType As clsLetterType
Dim iLetterTypeCol As Integer
Dim iColCnt As Integer

    strProcName = ClassName & ".HighlightTimeSensitiveQueue"
    '' I should switch this over to use QueueColumns.GetLiValue but it works as is..
    Set oLetterType = New clsLetterType
    

    For Each oCHdr In Me.lvQueue.ColumnHeaders
        
        If LCase(oCHdr.Text) = LCase(sCOLUMNNAME) Then
            bFoundCol = True
            iOurCol = iColCnt
        End If
        
        If LCase(oCHdr.Text) = "lettertype" Then
            iLetterTypeCol = iColCnt
        End If
        
        iColCnt = iColCnt + 1
    Next
    
    If bFoundCol = False Then GoTo Block_Exit
    
    For Each oLI In Me.lvQueue.ListItems
        
        If IsNumeric(oLI.SubItems(iOurCol)) Then
            If CInt(oLI.SubItems(iOurCol)) <= Me.DaysLeftToSendForHighlight Then
                Set oLetterType = New clsLetterType
                If iLetterTypeCol > 0 Then
                    oLetterType.LetterType = oLI.SubItems(iLetterTypeCol)
                Else
                    oLetterType.LetterType = oLI.Text
                End If

                If oLetterType.IsTimeSensitive = True Then
                    iRet = iRet + 1
                    oLI.Bold = True
                    oLI.ForeColor = RGB(255, 0, 0)
                    oLI.ToolTipText = "!!! This letter type has claims about to expire! It needs to be released soon !!!"
                End If
            End If
        End If
    Next
        
    
Block_Exit:
    HighlightTimeSensitiveQueue = iRet
    Set oCHdr = Nothing
    Set oLI = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function TotalListViewRow(oLV As CustomControl, ByVal sCOLUMNNAME As String) As Long
On Error GoTo Block_Err
Dim oListV As ListView
Dim strProcName As String
Dim lRet As Long
Dim iOurCol As Integer
Dim oCHdr As ColumnHeader
Dim oLI As ListItem

    strProcName = ClassName & ".TotalListViewRow"

    For Each oCHdr In oLV.ColumnHeaders
        
        If LCase(oCHdr.Text) = LCase(sCOLUMNNAME) Then
            Exit For
        End If
        iOurCol = iOurCol + 1
    Next
    
    For Each oLI In oLV.ListItems
        If iOurCol > 0 Then
            If IsNumeric(oLI.SubItems(iOurCol)) = True Then
                lRet = lRet + CLng(oLI.SubItems(iOurCol))
            Else
                If Nz(oLI.SubItems(iOurCol), "") = "" Then
                    lRet = 0
                Else
                    lRet = lRet + 1
                End If
            End If
        
        Else
            If IsNumeric(oLI.Text) = True Then
                lRet = lRet + CLng(oLI.Text)
            Else
                lRet = lRet + 1
            End If
            
        End If
    Next
        
    
Block_Exit:
    TotalListViewRow = lRet
    Set oCHdr = Nothing
    Set oLI = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub PopulateListView(oLV As CustomControl, oRs As ADODB.RecordSet)
On Error GoTo Block_Err
Dim oListV As ListView
Dim strProcName As String
Dim oLItem As ListItem
Dim oFld As ADODB.Field
Dim iCnt As Integer

    strProcName = ClassName & ".PopulateListView"
    
    
    oLV.ListItems.Clear
    oLV.Sorted = False
    
    oLV.ColumnHeaders.Clear
    
    For Each oFld In oRs.Fields
        oLV.ColumnHeaders.Add , , oFld.Name
    Next
    
    While Not oRs.EOF
        Set oLItem = oLV.ListItems.Add(, , oRs(0).Value)
        iCnt = 0
        For Each oFld In oRs.Fields
            If oFld.Name <> oRs(0).Name Then
                oLItem.SubItems(iCnt) = CStr(Nz(oRs(oFld.Name).Value, ""))
            End If
            iCnt = iCnt + 1
        Next

        oRs.MoveNext
    Wend
    
        
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

'Private Sub cmdFakeGenerateErrors_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'
'    strProcName = ClassName & ".cmdFakeGenerateErrors_Click"
'
'    sSql = "Update Q SET Error = 1, ErrorDesc = 'Sample error', Status = Status + 'E' " & _
'        " FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue Q WHERE Status = 'W' AND PrintQueueId IN ( " & _
'        "    SELECT TOP 2 PrintQueueId FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue WHERE Status = 'W'  " & _
'        "    ORDER BY ClaimClock DESC )   "
''Stop
'    Call ExecuteSQL(sSql)
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

'Private Sub cmdFakeOutput_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'
'    strProcName = ClassName & ".cmdFakeOutput_Click"
'
'    sSql = "Update Q SET STATUS = 'G'  " & _
'        " FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue Q WHERE Status = 'W' "
''Stop
'    Call ExecuteSQL(sSql)
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

'Private Sub cmdFakeOutputErrors_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'
'    strProcName = ClassName & ".cmdFakeOutputErrors_Click"
'
'    sSql = "Update Q SET Error = 1, ErrorDesc = 'Printer Jammed' " & _
'        " FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue Q WHERE Status = 'G' AND PrintQueueId IN ( " & _
'        "    SELECT TOP 2 PrintQueueId FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue WHERE Status = 'G'  " & _
'        "    ORDER BY ClaimClock " & _
'        ")"
'Stop
'    Call ExecuteSQL(sSql)
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub
'
'Private Sub cmdFakeQueueErrors_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'
'    strProcName = ClassName & ".cmdFakeQueueErrors_Click"
'
'    sSql = "Update Q SET Addr01 = NULL, City = Null, State = Null, Zip = Null " & _
'        " FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue Q WHERE Status = 'Q' AND PrintQueueId IN ( " & _
'        "    SELECT TOP 2 PrintQueueId FROM CMS_AUDITORS_CLAIMS.dbo.LETTER_Print_Queue  " & _
'        "    ORDER BY ClaimClock " & _
'        ")"
'    Call ExecuteSQL(sSql)
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

Private Function ExecuteSQL(sSql As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".ExecuteSQL"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        .Execute
    End With
    
    Call RefreshData
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub ckShowCompleted_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ckShowCompleted.Value = Not ckShowCompleted.Value
    If Nz(Me.fraShowCompleted, 0) = 0 Then
        Me.fraShowCompleted = 1
    Else
        Me.fraShowCompleted = 0
    End If

End Sub

Private Sub cmbAccountId_Change()
    SelectedAccountId = cmbAccountId.Value
    Call RefreshData
End Sub

Private Sub cmbFltrLetterType_Change()
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLView As ListView
Dim oLView As Object
Dim oLI As ListItem
Dim iLTCol As Integer

    strProcName = ClassName & ".cmbFltrLetterType_Change"
    
    iLTCol = QueueColumns.GetDetails("Queue", "LEtterType")
    
    
    Set oLView = Me.lvQueue
    For Each oLI In oLView.ListItems
        If iLTCol = 0 Then
            If oLI.Text = Me.cmbFltrLetterType Then
                oLI.Checked = True
            Else
                oLI.Checked = False
            End If
        
        Else
'            If oLI.SubItems(QueueColumns.GetDetails("Queue", "LetterType")) = Me.cmbFltrLetterType Then
            If UCase("" & QueueColumns.GetLiValue(oLI, "Queue", "LetterType")) = UCase("" & Me.cmbFltrLetterType) Then
                oLI.Checked = True
            Else
                oLI.Checked = False
            End If
        
        End If
    Next
    
    
Block_Exit:
    Set oLI = Nothing
    Set oLView = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdManOverrides_Click()
    DoCmd.OpenForm "frm_BOLD_Manual_Overrides"

End Sub

Private Sub cmdOpenMailRoomDash_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdOpenMailRoomDash_Click"
    
    DoCmd.OpenForm "frm_BOLD_Mail_Dashboard", acNormal, , , , acWindowNormal
    DoCmd.Close acForm, Me.Name, False
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshData
End Sub

Private Sub cmdRefreshLoadQErrors_Click()
    RefreshLoadQueueErrors
End Sub

Private Sub cmdReleaseErrors_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLV As ListView
Dim oListV As Object
Dim oLItem As ListItem
Dim saryList() As String
Dim iNumItems As Integer
Dim sIdList As String
Dim sSql As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdReleaseErrors_Click"
    Stop
    ' get the list of QueueId's to "release" as acceptable errors
    
    Set oListV = Me.lvQErrorDetails
    For Each oLItem In oListV.ListItems
        If oLItem.Checked = True Then
            
            ReDim Preserve saryList(iNumItems)
            saryList(iNumItems) = CStr(oLItem.Text)
            iNumItems = iNumItems + 1
        End If
    Next
    
    ' Now, convert that to an in clause:
    sIdList = MultipleValuesToXml("PrintQueueId", saryList)
    
    sSql = "usp_LETTER_Automation_ReleaseQueueErrors"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = sSql
        .Parameters.Refresh
        .Parameters("@pAccountId") = Me.SelectedAccountId
        .Parameters("@pIDList") = sIdList
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "There was a problmen releasing the errors!", .Parameters("@pErrMsg"), True
            GoTo Block_Exit
        End If
    End With
    
    Call RefreshData
    
Block_Exit:
    Set oLItem = Nothing
    Set oListV = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdReleaseQdLetters_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
'Dim oLV As ListView
Dim oLV As Object
Dim saryReleaseThese() As String
Dim iIdx As Integer
Dim sInstanceIdList As String
Dim iSubCol As Integer
Dim sCriteria As String
Dim sXmlStart As String
Dim sXmlEnd As String
Dim sRet As String
Dim sLetterType As String
Dim sLetterReqDt As String
Dim sHeld As String
Dim sManualOverRide As String
Dim sQueueStatusDt As String
Dim sQueueDt As String

    strProcName = ClassName & ".cmdReleaseQdLetters_Click"
    '' this needs to be the DynamicInstanceId
    '' not letter type - well, not anymore.. :D
    If coLVColPos Is Nothing Then
        Set coLVColPos = New clsLVColumnPositions
        Call coLVColPos.SetDetails("QUEUE", Me.lvQueue)
    End If
    iSubCol = coLVColPos.GetDetails("QUEUE", "DynamicInstanceId")
    

    Set oLV = Me.lvQueue
    
    For Each oLI In oLV.ListItems
        If oLI.Checked = True Then

            sLetterType = QueueColumns.GetLiValue(oLI, "QUEUE", "LetterType")    ' oLI.Text

            sLetterReqDt = QueueColumns.GetLiValue(oLI, "QUEUE", "LetterReqDt")
            sHeld = QueueColumns.GetLiValue(oLI, "QUEUE", "Held")
            sManualOverRide = QueueColumns.GetLiValue(oLI, "QUEUE", "ManualOverRide")
            sQueueStatusDt = QueueColumns.GetLiValue(oLI, "QUEUE", "StatusDate")
            sQueueDt = QueueColumns.GetLiValue(oLI, "QUEUE", "QueueDate")
            
            Call ReleaseLetterTypes(sLetterType, sLetterReqDt, sHeld, sManualOverRide, sQueueStatusDt, sQueueDt)

        End If
    Next

    Call Me.RefreshData
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdSelAllForRelease_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLV As ListView
Dim oListV As Object
Dim oListItem As ListItem
Static bSelected As Boolean
Dim bFirst As Boolean

    strProcName = ClassName & ".CmdSelAllForRelease_Click"
    
    bSelected = IIf(bSelected, False, True) ' just reverse the setting
    
    Set oListV = Me.lvQueue
    bFirst = True
    
    For Each oListItem In oListV.ListItems
        ' If they refreshed it, the check boxes will not be checked..
        ' so, if the last thing was to have them selected,
        ' and they aren't, then have them selected again (since they aren't)
        If bFirst = True Then
            If bSelected = False And oListItem.Checked = False Then
                bSelected = True
            End If
        End If
        oListItem.Checked = bSelected
        
    Next
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSelectAllErrors_Click()
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLV As ListView
Dim oListV As Object
Dim oLItem As ListItem
Dim bSelect As Boolean

    strProcName = ClassName & ".cmdSelectAllErrors_Click"
    
    Select Case LCase(Me.cmdSelectAllErrors.Caption)
    Case "SELECT ALL ERRORS"
        bSelect = True
        Me.cmdSelectAllErrors.Caption = "DE-Select All Errors"
    Case "DE-SELECT ALL ERRORS"
        bSelect = False
        Me.cmdSelectAllErrors.Caption = "Select All Errors"
    Case Else
        bSelect = True
        Me.cmdSelectAllErrors.Caption = "DE-Select All Errors"
    End Select
    
    
    
    Set oListV = Me.lvQErrorDetails
    For Each oLItem In oListV.ListItems
        oLItem.Checked = bSelect
    Next
    
Block_Exit:
    Set oLItem = Nothing
    Set oListV = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSpecificDay_Click()
'On Error GoTo Err_btnChkDt_Click
'
'    Set frmCalendar = New Form_frm_GENERAL_Calendar
'    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
'    frmCalendar.DatePassed = Nz(Me.txtFromDate, Date)
'    frmCalendar.RefreshData
'    ShowFormAndWait frmCalendar
'
'    'To avoid value 12:00:00 AM when closing calendar form
'    If mReturnDate = #12:00:00 AM# Then
'        Exit Sub
'    End If
'
'    'Prevent user from entering FROM that is greater than TODATE
'    If CDate(mReturnDate) > CDate(Me.txtToDate) Then
'        MsgBox "From Date cannot be greater than To Date", vbOKOnly + vbCritical
'        Exit Sub
'    End If
'
'    Me.txtFromDate = mReturnDate
'
'Exit_btnChkDt_Click:
'    Exit Sub
'
'Err_btnChkDt_Click:
'    MsgBox Err.Description
'    Resume Exit_btnChkDt_Click
End Sub



Private Sub dtSpecificDate_Change()
    ' doesn't make any sense if we aren't showing 'Completed' ones does it?
    Me.fraShowCompleted = 1
    cbSpecificDateSelected = True
    Call RefreshData
End Sub

Private Sub dtSpecificDate_Updated(Code As Integer)
    cbSpecificDateSelected = True
End Sub

Private Sub Form_Load()
Dim iAccount As Integer
    DoCmd.Hourglass True
    'Me.fraShowCompleted = 0
    Set coSettings = New clsSettings
    Me.AccountListSproc = GetSetting("ACCOUNT_LIST_SPROC")
    iAccount = GetUserSetting(GetUserName(), "AccountId")
    gintAccountID = 0
    
    Set oMainErrorGrid = Me.sfrmMainErrorGrid.Form
    cbSpecificDateSelected = False
    Call LoadAccountList
    ' Pre-select the global account
    SelectedAccountId = iAccount
    Me.cmbAccountId = iAccount
    
    Me.tabDisplay.Pages(2).SetFocus
    
    Call RefreshData

    DoCmd.Hourglass False
End Sub

Public Sub LoadAccountList()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sList As String

    strProcName = ClassName & ".LoadAccountList"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = Me.AccountListSproc
        .Parameters.Refresh
        Set oRs = .ExecuteRS
        If .GotData = False Or Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", "No list of accounts retrieved from sproc: '" & Me.AccountListSproc & "'", Nz(.Parameters("@pErrMsg"), ""), True
        End If
    End With
    
    While Not oRs.EOF
        sList = sList & CStr(oRs("AccountID").Value) & ";" & oRs("ClientName").Value & ";"
        oRs.MoveNext
    Wend
    
    Me.cmbAccountId.RowSource = left(sList, Len(sList) - 1) ' don't want the trailing ;
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Resize()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim oCtl2 As Control
Dim lSpace As Long

    strProcName = ClassName & ".Form_Resize"
    
    
    If Me.Width < 500 Then
        GoTo Block_Exit
    End If
    
    If Me.InsideWidth < 13500 Then
        Me.InsideWidth = 13500
    End If
    If Me.InsideHeight < 11895 Then
        Me.InsideHeight = 11895
    End If
    
    ' Make sure it's not TOO small (but don't forget about the whole minimize thing..
    ' something like state or whatever..
    
    
    ' How about something like this:
    ' Go through all of the controls on the form
    ' I guess I'm going to have to use the tag property..
    For Each oCtl In Me.Controls
    Debug.Assert oCtl.Name <> "fraShowCompleted " '& oCtl.left
        Select Case UCase(oCtl.Tag)
        Case "R=RIGHT"
Debug.Print oCtl.Name & " Left: " & oCtl.left
'            lSpace = (Me.InsideWidth - (oCtl.left + oCtl.width))
            lSpace = 400
            If Me.InsideWidth - oCtl.Width - lSpace > 0 Then
                oCtl.left = Me.InsideWidth - oCtl.Width - lSpace
            End If
Debug.Print oCtl.Name & " Left: " & oCtl.left
            If IsControl(Me, CStr(oCtl.Name) & "_LBL") = True Then
                Set oCtl2 = Me.Controls(oCtl.Name & "_LBL")
                If oCtl.left - oCtl2.Width > 200 Then
                    oCtl2.left = oCtl.left - oCtl2.Width
                End If
            End If
        Case "R=WHOLE"
            lSpace = oCtl.left * 2
            oCtl.Width = Me.InsideWidth - lSpace
        ' don't really have to do left's
        
        End Select
    Next

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub








Private Sub fraShowCompleted_AfterUpdate()
    cbSpecificDateSelected = True
End Sub

Private Sub lvGenerate_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvGenerate, ColumnHeader)
End Sub

Private Sub lvGenerate_DblClick()
'Stop    ' what should we do with this one?

End Sub

Private Sub lvGenerate_ItemClick(ByVal Item As Object)
'Stop
End Sub

Private Sub lvGenerate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvGenerate.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If
    
    Call SetUpContextMenu("Generate")
'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvGenerate.SelectedItem.Text
    
    Set oLI = Me.lvGenerate.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvGenerate"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:

End Sub

Private Sub lvGenerateErrs_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvGenerateErrs, ColumnHeader)
End Sub

Private Sub lvGenerateErrs_DblClick()

'Dim oLI As ListItem
Dim strParameterString As String
Dim oLI As Object
    ' gives us the PrintQueueId (unless we change the first column)
    Debug.Print lvQErrorDetails.SelectedItem
    Set oLI = lvQErrorDetails.SelectedItem
    ' going to launch an 'edit' screen here..
    ' the problem is:
    ' - shouldn't we fix the provider address in the system?
    ' - if so, well, how do we make that generic
    '   so CMS and MCR can use it with only configurations
    ' (and without a hell of a lot of work - since this is supposed to be the quick solution)
    
    ' After the edit we need to refresh the errors
'Stop
    strParameterString = QueueColumns.GetLiValue(oLI, "GENERATEERRORS", "CnlyClaimNum")
    Navigate Me.Name, "CLAIM", "DblClick", strParameterString
    
End Sub

Private Sub lvGenerateErrs_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvOutput.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If
    
    Call SetUpContextMenu("GenerateErrors")

'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvOutput.SelectedItem.Text
    
    Set oLI = Me.lvOutput.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvOutput"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:
End Sub


Private Sub lvOutput_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvOutput, ColumnHeader)
End Sub

Private Sub lvOutput_ItemClick(ByVal Item As Object)
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sDynamicInstanceId As String
Dim sCriteria As String

    strProcName = ClassName & ".lvOutput_ItemClick"
    ' here we are going to show them the sample with the watermark
    
    Set oLI = Me.lvOutput.SelectedItem
    
'    sDynamicInstanceId = oLI.SubItems(12)   ' KD Comeback - UN- hard code this..
'    If sDynamicInstanceId = "" Then
'        GoTo Block_Exit
'    End If
    If Me.SelectedAccountId = 0 Then
        Me.SelectedAccountId = 1
'        Stop
    End If
    
    
    
    sCriteria = "LetterType = '" & QueueColumns.GetLiValue(oLI, "OUTPUT", "LetterType") & "' "
    sCriteria = sCriteria & " AND AccountId = " & CStr(Me.SelectedAccountId)
    sCriteria = sCriteria & " AND LetterReqDt = '" & QueueColumns.GetLiValue(oLI, "OUTPUT", "LetterReqDt") & "'"
    sCriteria = sCriteria & " AND InstanceId = '" & QueueColumns.GetLiValue(oLI, "OUTPUT", "InstanceId") & "'"
    
    
    '' Load these in the below grid..
    

    DoCmd.Hourglass True


    sSql = "SELECT * FROM v_LETTER_AUTOMATION_OPS_ClaimDetailForInstances WHERE " & sCriteria
'Stop

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
    Call PopulateListView(Me.lvOutputErrors, oRs)
    
        ' now refresh our column object (in case something changed..)
    Call QueueColumns.SetDetails("OutputErrors", Me.lvOutputErrors)

    
    
Block_Exit:
    Set oAdo = Nothing
    Set oLI = Nothing
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub lvOutput_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvOutput.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If
    
    Call SetUpContextMenu("Output")

'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvOutput.SelectedItem.Text
    
    Set oLI = Me.lvOutput.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvOutput"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:
End Sub



Private Sub lvOutputErrors_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvOutputErrors, ColumnHeader)
End Sub

Private Sub lvOutputErrors_DblClick()
' stop:
Dim oLI As Object
Dim strParameterString As String

    Set oLI = Me.lvOutputErrors.SelectedItem
    
    strParameterString = QueueColumns.GetLiValue(oLI, "OutputErrors", "CnlyClaimNum")
    
    Navigate Me.Name, "CLAIM", "DblClick", strParameterString

    
'    NavigateNow "CnlyClaimNum"
End Sub

Private Sub lvOutputErrors_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvOutputErrors.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If
    
    Call SetUpContextMenu("OutputErrors")

'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvOutputErrors.SelectedItem.Text
    
    Set oLI = Me.lvOutputErrors.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvOutput"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:
End Sub

Private Sub lvQErrorDetails_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvQErrorDetails, ColumnHeader)
End Sub

Private Sub NavigateNow(sFieldKeyName As String, Optional oRs As DAO.RecordSet, Optional sSearchType As String)
    'Damon added to control navigation
On Error GoTo ErrHandler
Dim strProcName As String
Dim strParameter As String
Dim strParameterString As String

Dim strError As String
Dim strParent As String
Dim arrParameters() As String
Dim intI As Integer
Dim strAppID As String
    Stop     ' this doesn't seem to be working the way I planned.. :D
    
    strProcName = ClassName & ".NavigateNow"
    
    
    If IsMissing(oRs) Or oRs Is Nothing Then

        Set oRs = Me.RecordSet
    End If
    
    strParameterString = ""

    strParent = Me.Name
    If sSearchType <> "" Then
        strAppID = sSearchType
    Else
        If strAppID = "" Then

            strAppID = Me.frmAppID
        End If
    End If

    
    If sFieldKeyName <> "" Then
        strParameter = sFieldKeyName
    Else
        strParameter = Nz(DLookup("Parameter", "GENERAL_Navigate", "SearchType = '" & strAppID & "' and ActionName = 'DblClick' and ParentForm = '" & strParent & "'"), "")

    End If
    
    
    
        
    arrParameters = Split(strParameter, "|")
    
    
    'added by mike g 11-04-2011
    If strParameter <> "" Then
        If UBound(arrParameters) > 0 Then
            For intI = 0 To UBound(arrParameters)
               strParameterString = strParameterString & Me.RecordSet(arrParameters(intI)) & "|"
            Next intI
        Else
            If isDAOField(Me.RecordSet, arrParameters(0)) = True Then
                strParameterString = strParameterString & Me.RecordSet(arrParameters(0))
            End If
        End If
        
        If strParameter <> "" And strParameterString <> "" Then
            Navigate strParent, strAppID, "DblClick", strParameterString
        End If
            
    Else
        Exit Sub
    End If
    
Block_Exit:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"
    GoTo Block_Exit
End Sub

Private Sub lvQErrorDetails_DblClick()

'Dim oLI As ListItem
Dim strParameterString As String
Dim oLI As Object
    ' gives us the PrintQueueId (unless we change the first column)
    Debug.Print lvQErrorDetails.SelectedItem
    Set oLI = lvQErrorDetails.SelectedItem
    ' going to launch an 'edit' screen here..
    ' the problem is:
    ' - shouldn't we fix the provider address in the system?
    ' - if so, well, how do we make that generic
    '   so CMS and MCR can use it with only configurations
    ' (and without a hell of a lot of work - since this is supposed to be the quick solution)
    
    ' After the edit we need to refresh the errors
'Stop
    strParameterString = oLI.SubItems(18)
    Navigate Me.Name, "CLAIM", "DblClick", strParameterString
    
'    NavigateNow "CnlyClaimNum"
    
End Sub



Private Sub lvQErrorDetails_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvQErrorDetails.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If

    Call SetUpContextMenu("QueueErrors")

'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvQErrorDetails.SelectedItem.Text
    
    Set oLI = Me.lvQErrorDetails.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvQErrorDetails"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:

End Sub

Private Sub lvQueue_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvQueue, ColumnHeader)
End Sub

Private Sub SortListView(oLV As CustomControl, ColumnHeader As Object)
    oLV.Sorted = False
    oLV.SortKey = ColumnHeader.index - 1
    oLV.SortOrder = IIf(oLV.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    oLV.Sorted = True
End Sub

Private Sub lvQueue_DblClick()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oAdo As clsADO
Dim oTmpltRs As ADODB.RecordSet
Dim oRs As ADODB.RecordSet
Dim sTmpFolder As String
Dim dctLtrTemplates As Scripting.Dictionary
Dim sDynamicInstanceId As String
Dim lRowsAffected As Long
Dim sLtrPath As String
Dim oWordApp As Word.Application
Dim sCriteria As String

    strProcName = ClassName & ".lvQueue_DblClick"
    ' here we are going to show them the sample with the watermark
    
    Set oLI = Me.lvQueue.SelectedItem
    
'    sDynamicInstanceId = oLI.SubItems(12)   ' KD Comeback - UN- hard code this..
'
'    If sDynamicInstanceId = "" Then
'        GoTo Block_Exit
'    End If
'
 
    sCriteria = "LetterType = '" & oLI.Text & "' "
    sCriteria = sCriteria & " AND AccountId = " & CStr(Me.SelectedAccountId)
    sCriteria = sCriteria & " AND LetterReqDt = '" & oLI.SubItems(3) & "'"
    sCriteria = sCriteria & " AND Held = '" & oLI.SubItems(7) & "'"
    sCriteria = sCriteria & " AND InProgress = '" & oLI.SubItems(8) & "'"
    sCriteria = sCriteria & " AND ManualOverRide = '" & oLI.SubItems(11) & "'"
            
    
    
    DoCmd.Hourglass True


    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_Sample"
        .Parameters.Refresh
        .Parameters("@pAccountId") = 1
        .Parameters("@pLetterType") = oLI.Text
        .Parameters("@pLetterReqDt") = oLI.SubItems(3)
        .Parameters("@pHeld") = oLI.SubItems(7)
        .Parameters("@pInProgress") = oLI.SubItems(8)
        .Parameters("@pManualOverRide") = oLI.SubItems(11)
'        .Parameters("@pDynamicInstanceId") = sDynamicInstanceId
        Set oTmpltRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    If oTmpltRs.recordCount < 1 Then
        Stop
        
    End If
    
    '' Copy templates to a temp work folder:
    If CopyTemplatesToTempWorkFldr(oTmpltRs, sTmpFolder, dctLtrTemplates) = False Then
        Stop
    End If
    
    ' and advance the oRs to the next one
    Set oRs = oTmpltRs.NextRecordset

    If oRs.recordCount < 1 Then
        Stop
    End If
    '' Now do the individual mail merges:
    Set oWordApp = New Word.Application
'
'    If PerformIndividualMailMerges(oWordApp, oRs, dctLtrTemplates, sTmpFolder, lRowsAffected, True, sLtrPath) = False Then
    If PerformIndividualMailMerges(oWordApp, oRs, dctLtrTemplates, sTmpFolder, lRowsAffected, True, sLtrPath) = False Then
        Stop
    End If
    oWordApp.visible = True
    
    SleepEvents 1
    
    Call ActivateApplicationWindow(, , "WORD")
    
Block_Exit:
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub lvQueue_ItemClick(ByVal Item As Object)
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sDynamicInstanceId As String
Dim sCriteria As String

    strProcName = ClassName & ".lvQueue_ItemClick"
    ' here we are going to show them the sample with the watermark
    
    Set oLI = Me.lvQueue.SelectedItem
    
'    sDynamicInstanceId = oLI.SubItems(12)   ' KD Comeback - UN- hard code this..
'    If sDynamicInstanceId = "" Then
'        GoTo Block_Exit
'    End If
    If Me.SelectedAccountId = 0 Then
        Me.SelectedAccountId = 1
'        Stop
    End If
    
    sCriteria = "LetterType = '" & oLI.Text & "' "
    sCriteria = sCriteria & " AND AccountId = " & CStr(Me.SelectedAccountId)
    sCriteria = sCriteria & " AND LetterReqDt = '" & oLI.SubItems(3) & "'"
    sCriteria = sCriteria & " AND Held = '" & oLI.SubItems(7) & "'"
    sCriteria = sCriteria & " AND InProgress = '" & oLI.SubItems(8) & "'"
    sCriteria = sCriteria & " AND ManualOverRide = '" & oLI.SubItems(11) & "'"
        
    
    '' Load these in the below grid..
    

    DoCmd.Hourglass True

'    sSql = "SELECT * FROM v_LETTER_AUTOMATION_OPS_ClaimDetailForInstances WHERE DynamicInstanceId = '" & sDynamicInstanceId & "' ORDER BY ICN "

    sSql = "SELECT * FROM v_LETTER_AUTOMATION_OPS_ClaimDetailForInstances WHERE " & sCriteria
'Stop

    If oClaimsGrid Is Nothing Then
        Set oClaimsGrid = Me.gdsClaimDetailsGrid.Form
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
'    oClaimsGrid.InitData sSql, 2
 
    
    oClaimsGrid.InitDataADO oRs, "v_Code_Database"
    Set oClaimsGrid.RecordSet = oRs
    
Block_Exit:
    Set oAdo = Nothing
    Set oLI = Nothing
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub lvQueue_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim oLI As ListItem

    If lvQueue.SelectedItem Is Nothing Then
        GoTo Block_Exit
    End If
    
    Call SetUpContextMenu("Queue")

'Dim oLV As ListView
'    Stop
'    Debug.Print oLV.SelectedItem.Text
Debug.Print CStr(X) & " y: " & CStr(Y)
    Debug.Print Me.lvQueue.SelectedItem.Text
    
    Set oLI = Me.lvQueue.SelectedItem
    
    oLI.Selected = True
    RightClickedListItem = oLI
    ListViewNameClicked = "lvQueue"
    
    
    If Button = acRightButton Then
'        Call CommandBars("BOLD_RightClickmnu").ShowPopup(x, y)
        Call CommandBars("BOLD_RightClickmnu").ShowPopup
    End If
Block_Exit:
    
End Sub



Private Sub oClaimsGrid_Click()
'On Error GoTo Block_Exit
'Dim strProcName As String
'Dim oFld As ADODB.Field
'Dim oRs As ADODB.RecordSet
'Dim sCnlyClaimNum As String
'
'
'
'    strProcName = ClassName & ".oClaimsGrid_Click"
'
'    Set oRs = oClaimsGrid.RecordsetClone
'
'    For Each oFld In oRs.Fields
'        If InStr(1, oFld.Name, "cnlyclaimnum", vbTextCompare) > 0 Then
'            sCnlyClaimNum = Nz(oRs(oFld.Name).Value, "")
'            Exit For
'        End If
'    Next
'
'    NewMain sCnlyClaimNum, ""
'
'Block_Exit:
'    Set oFld = Nothing
'    Set oRs = Nothing
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
End Sub



' Update the Detail grid if necessary
Private Sub oMainErrorGrid_Current()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim lPrintQueueRunId As Long
Dim lAccountId As Long
Dim lErrorTypeId As Long

Dim sLetterType As String
Dim sSql As String

    strProcName = ClassName & ".oMainErrorGrid_Current"
    

    If oMainErrorGrid Is Nothing Then
        GoTo Block_Exit
    End If
    If oMainErrorGrid.RecordSource = "" Then
'        Stop
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    
'' Ok, so we need to get the master fields to link to
''
'    lPrintQueueRunId = Nz(oMainErrorGrid.Controls("PrintQueueRunId"), 0)
    lAccountId = Nz(oMainErrorGrid.Controls("AccountId"), 0)
    sLetterType = Nz(oMainErrorGrid.Controls("LetterType"), "")
    lErrorTypeId = Nz(oMainErrorGrid.Controls("ErrorTypeId"), 0)
    
'    sSql = "SELECT * FROM v_LETTER_AUtomation_AddQueueErrorLog WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & _
'            "# AND (AccountId = " & CStr(lAccountId) & " OR " & CStr(lAccountId) & " = 0) AND PrintQueueRunId = " & CStr(lPrintQueueRunId) & " AND LetterType = """ & _
'            sLetterType & """"

    ''' MCR:
'    sSql = " SELECT RelatedCnlyClaimNum, ErrorMsg, AdditionalDetails, LetterType, LetterDesc, LetterSource, LetterCode, CnlyProvId, " & _
'        " AccountID , AddQueueErrorId, PrintQueueRunId, ErrorSql, RelatedViewName, ErrorTypeId, NotFoundinTblName, QueueRunDt "

    sSql = " SELECT RelatedCnlyClaimNum, ErrorMsg, AdditionalDetails, LetterType, LetterDesc, LetterSource, CnlyProvId, " & _
        " AccountID , AddQueueErrorId, PrintQueueRunId, ErrorSql, RelatedViewName, ErrorTypeId, NotFoundinTblName, QueueRunDt "

    
    sSql = sSql & " FROM v_LETTER_Automation_AddQueueErrorLog WHERE ErrDateTime >= #" & CStr(Month(Now())) & "/" & CStr(Day(Now())) & "/" & CStr(Year(Now())) & _
            "# AND (AccountId = " & CStr(lAccountId) & " OR " & CStr(lAccountId) & " = 0) AND LetterType = """ & _
            sLetterType & """ AND ErrorTypeId = " & CStr(lErrorTypeId)
       

    If oDetailErrorGrid Is Nothing Then
        Set oDetailErrorGrid = Me.sfrmDetailErrorGrid.Form
    End If
    
    oDetailErrorGrid.RecordSource = sSql
    oDetailErrorGrid.InitData sSql, 2
    

'            '    Me.txtConceptID = Nz(oMainGrid.Controls("ConceptID"), "")
'            '    Me.txtNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'            '    mNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'    'Refresh the tabs to ensure the main form is in sync with the other forms.
'
'    oAdo.ConnectionString = DataConnString
'    oAdo.sqlString = sSql
'    Set oRs = oAdo.OpenRecordSet()
'
'    ' if it's the same concept, no need to "click" the tab
'    If Not frmConceptHdr Is Nothing Then
'        If Me.txtConceptID <> frmConceptHdr.FormConceptID Then
'            lstTabs_Click
'        Else
''            Stop
'            lstTabs_Click
'        End If
'    Else
'        lstTabs_Click
'    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub tabDisplay_Change()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".tabDisplay_Change"
    
    If Me.tabDisplay.Value = 4 Then
        Me.dtSpecificDate = Now()
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
