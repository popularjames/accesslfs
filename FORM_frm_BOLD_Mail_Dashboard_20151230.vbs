Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'' 6/12/2015: KD enabled selecting multiple print jobs
'' 8/15/2014: KD Need to make sure this is imported to the Master Claim Admin

Private coSettings As clsSettings

Private coRightClickListItem As ListItem
Private csListViewNameClicked As String
Private cdctLISubItems As Scripting.Dictionary
Private cbDateSelected As Boolean

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Private Const TwipsPerInch = 1440
Private Const MouseNormal = 0   '(Default) The shape is determined by Microsoft Access
Private Const MouseArrow = 1
Private Const MouseIBeam = 3
Private Const MouseVerticalResize = 7 ' (Size N, S)
Private Const MouseHorizontalResize = 9 '  Horizontal Resize (Size E, W)
Private Const MouseBusy = 111 ' Busy (Hourglass)

Private Type Size
        cx As Long
        cy As Long
End Type

Private Const LF_FACESIZE = 32

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function apiCreateFontIndirect Lib "gdi32" Alias _
        "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function apiSelectObject Lib "gdi32" _
 Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hwnd As Long) As Long

Private Declare Function apiReleaseDC Lib "user32" _
  Alias "ReleaseDC" (ByVal hwnd As Long, _
  ByVal hdc As Long) As Long

Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function apiGetTextExtentPoint32 Lib "gdi32" _
Alias "GetTextExtentPoint32A" _
(ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, _
lpSize As Size) As Long

' Create an Information Context
 Private Declare Function apiCreateIC Lib "gdi32" Alias "CreateICA" _
 (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
 ByVal lpOutput As String, lpInitData As Any) As Long
 
' Close an existing Device Context (or information context)
 Private Declare Function apiDeleteDC Lib "gdi32" Alias "DeleteDC" _
 (ByVal hdc As Long) As Long

 Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 
 Private Declare Function GetDeviceCaps Lib "gdi32" _
 (ByVal hdc As Long, ByVal nIndex As Long) As Long
 
 ' Constants
 Private Const SM_CXVSCROLL = 2
 Private Const LOGPIXELSX = 88

' Array of strings used to build the ColumnWidth property
Private strWidthArray() As String

' Array of Column Widths.
' The entries are cumulative in order to
' aid matching of the start of each column
Private sngWidthArray() As Single

' Amount of extra space to add to edge of each column
Private m_ColumnMargin As Long

' ListBox/Combo we are resizing
Private m_Control As Access.Control
'

Private coLVColPos As clsLVColumnPositions

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get SelectedAccountId() As Long
    SelectedAccountId = gintAccountID
    GlobalSelectedAccountId = gintAccountID
End Property
Public Property Let SelectedAccountId(intAccountId As Long)
    gintAccountID = intAccountId
    
    GlobalSelectedAccountId = gintAccountID
End Property


Public Property Get QueueColumns() As clsLVColumnPositions
    If coLVColPos Is Nothing Then
        Set coLVColPos = New clsLVColumnPositions
        Call coLVColPos.SetDetails("QUEUE", Me.lvQueue)
        Call coLVColPos.SetDetails("QUEUEERRORS", Me.lvQErrorDetails)
    End If
    Set QueueColumns = coLVColPos
End Property

Public Property Get SpecificDateSelected() As Boolean
    SpecificDateSelected = cbDateSelected
End Property
Public Property Let SpecificDateSelected(bDateSelected As Boolean)
    cbDateSelected = bDateSelected
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


Public Function GetPagesToday() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetPagesToday"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT SUM(TTlPageCount) As PagesToday FROM LETTER_Automation_MAILOPS_Batchs B WHERE B.AddDate >= '" & Format(Now(), "mm/dd/yyyy") & "' "
        Set oRs = .ExecuteRS
        GetPagesToday = CLng(Nz(oRs("PagesToday").Value, 0))
    End With
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function GetLettersToday() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetLettersToday"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT SUM(LetterCount) As LettersToday FROM LETTER_Automation_MAILOPS_Batchs B WHERE B.AddDate >= '" & Format(Now(), "mm/dd/yyyy") & "' "
        Set oRs = .ExecuteRS
        GetLettersToday = CLng(Nz(oRs("LettersToday").Value, 0))
    End With
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function GetBatchesToday() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetBatchesToday"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT COUNT(DISTINCT MailBatchId) As BatchesToday FROM LETTER_Automation_MAILOPS_Batchs B WHERE B.AddDate >= '" & Format(Now(), "mm/dd/yyyy") & "' "
        Set oRs = .ExecuteRS
        GetBatchesToday = CLng(Nz(oRs("BatchesToday").Value, 0))
    End With
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshData"
    DoCmd.Hourglass True


    Call RefreshQueue
    Set coLVColPos = New clsLVColumnPositions
    Call coLVColPos.SetDetails("QUEUE", Me.lvQueue)

    txtTotalBatchCountToday = GetBatchesToday()
    txtTtlLetterCountToday = GetLettersToday()
    txtTtlPageCountToday = GetPagesToday()


    Me.txtTtlBatches = TotalListViewRow(Me.lvQueue, "LetterType")
    Me.txtTtlLettersInQ = TotalListViewRow(Me.lvQueue, "LetterCount")
    Me.txtTtlPagesInQ = TotalListViewRow(Me.lvQueue, "PageCount")
    
    
    '' Now get the queue error details:
    Call RefreshQueueErrors
    Call coLVColPos.SetDetails("ERRORS", Me.lvQErrorDetails)
    
    
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


Public Sub RefreshQueue()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".RefreshQueue"

    sSql = "usp_LETTER_Automation_MAILOPS_DashQueue"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = sSql
        .Parameters.Refresh
        .Parameters("@pAccountId") = Nz(Me.cmbBusiness, 0)
        .Parameters("@pShowCompleted") = Nz(fraShowCompleted, 0)
        .Parameters("@pRange") = Nz(Me.cmbRange, "")
        .Parameters("@pSpecificDay") = Format(Me.dtSpecificDate, "m/d/yyyy")
        .Parameters("@pLetterType") = Nz(Me.cmbFltrLetterType, "")
        .Parameters("@pOpsBatchId") = Nz(Me.cmbOpsBatchId, "")
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value
            Stop
        End If
        
    End With

    If oRs.State <> adStateOpen Then
        Stop
        Me.lvQErrorDetails.ListItems.Clear
        GoTo Block_Exit
    End If
    
    Call PopulateListView(Me.lvQueue, oRs, "MailBatchId")
   
    Call HighlightTimeSensitiveQueue("AddDate", , RGB(100, 200, 100), RGB(255, 0, 0))
   
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
Dim sSql As String

Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet


    strProcName = ClassName & ".RefreshQueueErrors"

    sSql = "SELECT LetterType, ErrDesc, BatchId, Status, BatchType, isTimeSensitive, DeadlineDt, LetterCount, PageCount, Duplex FROM tbl_Mailroom_Queue " & _
        " WHERE Error <> 0 AND AuditID = " & CStr(Me.cmbBusiness) & _
        " ORDER BY Status, BatchId"

'    Set oDb = CurrentDb()
'    Set oRs = oDb.OpenRecordSet(sSql)
    

'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("CMS_AUDITORS_CODE")
'        .SQLTextType = StoredProc
'        .sqlString = "usp_LETTER_MAILROOM_PrintQueue_CurrentQueue"
'        .Parameters.Refresh
'        Set oRs = .OpenRecordSet
'    End With


'    Call PopulateListViewDAO(Me.lvQErrorDetails, oRs, "BatchId")
'
'    Call HighlightTimeSensitiveQueue("DeadlineDt", lvQErrorDetails)
   
Block_Exit:
    Set oRs = Nothing
    Set oDb = Nothing
'    If Not oRs Is Nothing Then
'        If oRs.State = adStateOpen Then oRs.Close
'        Set oRs = Nothing
'    End If
'    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Function HighlightTimeSensitiveQueue(Optional sColName As String, Optional oLV As Object, Optional lColor As Long, Optional lErrColor As Long) As Integer
On Error GoTo Block_Err
'Dim oListV As ListView
Dim oListV As Object
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
Dim iStatusCol As Integer

Dim iColCnt As Integer

    strProcName = ClassName & ".HighlightTimeSensitiveQueue"
    Set oLetterType = New clsLetterType
    
    If lColor = 0 Then
        lColor = RGB(255, 0, 0)
    End If
    If lErrColor = 0 Then
        lErrColor = RGB(255, 0, 0)
    End If
    
    If Not oLV Is Nothing Then
        Set oListV = oLV
    Else
        Set oListV = Me.lvQueue
    End If


    For Each oCHdr In oListV.ColumnHeaders
        If sColName <> "" Then
            If LCase(oCHdr.Text) = LCase(sColName) Then
                bFoundCol = True
                iOurCol = iColCnt
            End If
        Else
            If LCase(oCHdr.Text) = LCase(sCOLUMNNAME) Then
                bFoundCol = True
                iOurCol = iColCnt
            End If
        
        End If
        
        If LCase(oCHdr.Text) = "status" Then
            iStatusCol = iColCnt
        End If
        
        If LCase(oCHdr.Text) = "lettertype" Then
            iLetterTypeCol = iColCnt
        End If
        
        iColCnt = iColCnt + 1
    Next
    If iStatusCol = 0 Then
        iStatusCol = iOurCol
    End If
    If bFoundCol = False Then GoTo Block_Exit
    
    For Each oLI In oListV.ListItems
        
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
                    If InStr(1, oLI.SubItems(iStatusCol), "E", vbTextCompare) > 0 Then
                        oLI.ForeColor = lErrColor
                    Else
                        oLI.ForeColor = lColor
                    End If
                    oLI.ToolTipText = "!!! This letter type has claims about to expire! It needs to be released soon !!!"
                End If
            End If
        ElseIf IsDate(oLI.SubItems(iOurCol)) Then
            Debug.Print DateDiff("d", CDate(oLI.SubItems(iOurCol)), Now())
            If DateDiff("d", CDate(oLI.SubItems(iOurCol)), Now()) < 2 Then
                Set oLetterType = New clsLetterType
                If iLetterTypeCol > 0 Then
                    oLetterType.LetterType = oLI.SubItems(iLetterTypeCol)
                Else
                    oLetterType.LetterType = oLI.Text
                End If

                If oLetterType.IsTimeSensitive = True Then
                    iRet = iRet + 1
                    oLI.Bold = True
                    If InStr(1, oLI.SubItems(iStatusCol), "E", vbTextCompare) > 0 Then
                        oLI.ForeColor = lErrColor
                    Else
                        oLI.ForeColor = lColor
                        oLI.Ghosted = True
                    End If

                    oLI.ToolTipText = "!!! This letter type has claims about to expire! It needs to be released soon !!!"
                End If
            
            End If
        End If
    Next
        
    
Block_Exit:
    HighlightTimeSensitiveQueue = iRet
    Set oCHdr = Nothing
    Set oLI = Nothing
    Set oListV = Nothing
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
Dim bFound As Boolean


    strProcName = ClassName & ".TotalListViewRow"

    For Each oCHdr In oLV.ColumnHeaders
        
        If LCase(oCHdr.Text) = LCase(sCOLUMNNAME) Then
            bFound = True
            Exit For
        End If
        iOurCol = iOurCol + 1
    Next
    
    If bFound = False Then GoTo Block_Exit
    
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

Private Sub PopulateListView(oLV As CustomControl, oRs As ADODB.RecordSet, Optional sIdCol As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oLItem As ListItem
Dim oFld As ADODB.Field
Dim iCnt As Integer
Dim dctMaxWidths As Scripting.Dictionary
Dim iHeader As Integer
Dim lWidth As Long
Const lPerCharWidth As Long = 200

    strProcName = ClassName & ".PopulateListView"
    
    Set cdctLISubItems = New Scripting.Dictionary
    Set dctMaxWidths = New Scripting.Dictionary
    
    oLV.ListItems.Clear
    oLV.Sorted = False
    
    oLV.ColumnHeaders.Clear
    
    For Each oFld In oRs.Fields
        oLV.ColumnHeaders.Add , , oFld.Name
    Next
    
    While Not oRs.EOF
        iHeader = 1
        Set oLItem = oLV.ListItems.Add(, , oRs(0).Value)
        iCnt = 0
        
        lWidth = Len(CStr(Nz(oRs(0).Value, ""))) * lPerCharWidth
        
        If dctMaxWidths.Exists(1) Then
            If lWidth > dctMaxWidths.Item(1) Then
                dctMaxWidths.Item(1) = lWidth
            End If
        Else
            dctMaxWidths.Add 1, lWidth
        End If
        
        For Each oFld In oRs.Fields

            lWidth = Len(CStr(Nz(oRs(oFld.Name).Value, ""))) * lPerCharWidth
            
            If dctMaxWidths.Exists(iHeader) Then
                If lWidth > dctMaxWidths.Item(iHeader) Then
                    dctMaxWidths.Item(iHeader) = lWidth
                End If
            Else
                dctMaxWidths.Add iHeader, lWidth
            End If
            
            If oFld.Name <> oRs(0).Name Then
                oLItem.SubItems(iCnt) = CStr(Nz(oRs(oFld.Name).Value, ""))
                If cdctLISubItems.Exists(UCase(oFld.Name)) = False Then
                    cdctLISubItems.Add UCase(oFld.Name), iCnt
                End If
                If sIdCol <> "" And UCase(sIdCol) = UCase(oFld.Name) Then
                    oLItem.Tag = oRs(oFld.Name).Value
                End If
            End If
            iHeader = iHeader + 1
            iCnt = iCnt + 1
        Next

        oRs.MoveNext
    Wend

    ' Now resize the columns to the max:
'    For iHeader = 1 To oLV.ColumnHeaders.Count
'        lWidth = dctMaxWidths.Item(iHeader)
'
'        oLV.ColumnHeaders(iHeader).Width = lWidth
'
'    Next
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub PopulateListViewDAO(oLV As CustomControl, oRs As DAO.RecordSet, Optional sIdCol As String)
On Error GoTo Block_Err
Dim oListV As ListView
Dim strProcName As String
Dim oLItem As ListItem
Dim oFld As DAO.Field
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
                If sIdCol <> "" And UCase(sIdCol) = UCase(oFld.Name) Then
                    oLItem.Tag = oRs(oFld.Name).Value
                End If
                
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


Private Sub ckShowCompleted_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ckShowCompleted.Value = Not ckShowCompleted.Value
    If Nz(Me.fraShowCompleted, 0) = 0 Then
        Me.fraShowCompleted = 1
    Else
        Me.fraShowCompleted = 0
    End If
    Call RefreshData
End Sub

Private Sub cmbBusiness_Change()
    Call RefreshData
End Sub

Private Sub cmdClearFilters_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdClearFilters_Click"
    
    Me.cmbRange = ""
    Me.cmbFltrLetterType = ""
    Me.cmbOpsBatchId = ""
    SpecificDateSelected = False
    Me.dtSpecificDate = Now()
    Me.fraShowCompleted = 0
    
    
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

'Private Sub cmdFakeQueueErrors_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'Dim oDb As DAO.Database
'Dim lSelectedId As Long
'
'
'    strProcName = ClassName & ".cmdFakeQueueErrors_Click"
'
'    If Nz(Me.lvQueue.SelectedItem, "") = "" Then
'        MsgBox "Select one to simulate an error on!", vbOKOnly, "Which one?"
'        GoTo Block_Exit
'    End If
'    lSelectedId = Me.lvQueue.SelectedItem.Tag
'
'
'    sSql = "Update tbl_MailRoom_Queue SET Status = Status & 'E', Error = 1, ErrDesc = 'Printer Out of paper' WHERE QueueId = " & CStr(lSelectedId)
'
'    Set oDb = CurrentDb
'    oDb.Execute sSql
'    Call RefreshData
'
'
''    Call ExecuteSQL(sSql)
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
        .ConnectionString = GetConnectString("cms_auditors_claims")
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

Private Sub cmdOpenManualDash_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdOpenManualDash_Click"
    
    DoCmd.OpenForm "frm_LETTER_Main_NEW"
    DoCmd.Close acForm, Me.Name, acSaveNo
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdOpenOldForm_Click()
    DoCmd.OpenForm "frm_LETTER_Main", acNormal, , , , , "1"
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdOpenOpsDashboard_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdOpenMailRoomDash_Click"
    
    DoCmd.OpenForm "frm_BOLD_Ops_Dashboard", acNormal, , , , acWindowNormal
    DoCmd.Close acForm, Me.Name, False
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdPrintChecked_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oLItem As ListItem
Dim oLV As CustomControl
Dim oLView As ListView
Dim lSelected As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim lOpsBatchId As Long
Dim strPrinterBefore As String
Dim sSelectedPrinter As String
Dim lBatchRowIdToClose As Long
Dim lSelOpsBatchId As Long
Dim oFrmStatus As Form_ScrStatus
Dim sBatches As String
Dim varyList() As String
Dim bDuplex As Boolean
Dim sXml As String
Dim sBatchType As String
'Dim iBatchTypeCol As Integer
Dim sThisLine As String


    strProcName = ClassName & ".cmdPrintChecked_Click"
    
    ' validate that they want to do this?
    Set oLV = Me.lvQueue
    ReDim varyList(0)
    
'    iBatchTypeCol = QueueColumns.GetDetails("QUEUE", "BatchType")
    
    
    For Each oLItem In oLV.ListItems
        If oLItem.Checked = True Then
            lSelected = lSelected + 1

            ReDim Preserve varyList(lSelected - 1)
            sBatchType = QueueColumns.GetLiValue(oLItem, "QUEUE", "BatchType")

                ' now we want something like this:
                ' "BatchId=123;BatchType=Normal"
            sThisLine = MakeXmlString("MailBatchid=" & oLItem.Text & ";BatchType=" & sBatchType, "row", , , True)

            varyList(UBound(varyList)) = sThisLine
            
            sMsg = sMsg & " * " & CStr(oLItem.Text) & " - " & oLItem.SubItems(3) & vbCrLf
            lOpsBatchId = CLng(oLItem.Text) ' this was a bug because we could have the same MailBatchId for 2 rows, Regular Batch and Manual Batch
                                            ' now, each of those should get their own batchid
            If QueueColumns.GetLiValue(oLItem, "QUEUE", "JobIsDuplex") = "Y" Then
                bDuplex = True
            End If
                '            Exit For    ' only allowing 1 at a time for now
        ElseIf oLItem.Selected = True Then
                    '        Stop
                    '            lSelected = lSelected + 1
                    '            lSelOpsBatchId = CLng(oLItem.Text)
                    '
                    '            ReDim Preserve varyList(lSelected - 1)
                    '            varyList(UBound(varyList)) = oLItem.Text
                    '
                    '            If QueueColumns.GetLiValue(oLItem, "QUEUE", "JobIsDuplex") = "Y" Then
                    '                bDuplex = True
                    '            End If
        End If
    Next

    sXml = MultipleRowsAndColumnsToXmlForJoin("list", varyList)
    
    If bDuplex = True = True Then
        LogMessage strProcName, "CONFIRM!", "There is at least 1 duplex job that you selected. Please insure that you set the print settings correctly!", Join(varyList, ","), True
    End If
    
    If lSelected > 1 Or bDuplex = True Then
        If sMsg <> "" Then
            If MsgBox("Are you sure you wish to print the following to the same printer?" & vbCrLf & sMsg, vbYesNo, "Confirm") = vbNo Then
                GoTo Block_Exit
            End If
        End If
    End If

    sSelectedPrinter = SelectPrinter(, strPrinterBefore)
    DoCmd.Hourglass True
    
    Set oFrmStatus = New Form_ScrStatus
    With oFrmStatus
        .ShowCancel = False
        .ShowMessage = False
        .ShowMessage = True
        .ProgMax = 100
        .TimerInterval = 50
        .ShowProgressBar = True
        .show
    End With
    
    '' Make the cover page (we need the details - let's get them from the database again
    '' instead of reading from the cached list view which may not be up to date..
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        '.sqlString = "usp_LETTER_Automation_MAILOPS_StartPrinting"
        .sqlString = "usp_LETTER_Automation_MAILOPS_StartPrintingMultNew"
        .Parameters.Refresh
        .Parameters("@pIDList") = sXml
        .Parameters("@pPrintOperator") = GetUserName
        .Parameters("@pPrinterName") = sSelectedPrinter
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
'        lBatchRowIdToClose = .Parameters("@pRowId").Value
    End With
    
    
    
    ' should I do all of them here (Should really only be 1 per MailBatchId.. (right??))
    ' It's only been like 2 months since I last touched this code - whatdayawant? lol
    While Not oRs.EOF
        lBatchRowIdToClose = oRs("RowId").Value
        lOpsBatchId = oRs("OpsLetterBatchId").Value
        sBatchType = Nz(oRs("BatchType").Value, "Regular Batch")    ' just in case
        
'        If CreateCoverPageAndPrint(lOpsBatchId, sBatchType, oRs("CombinedFilePath").Value, lBatchRowIdToClose, sSelectedPrinter, oFrmStatus) = False Then
'            LogMessage strProcName, "ERROR", "There was a problem creating the Cover sheet!", sSelectedPrinter, True
'            GoTo Block_Exit
'            Stop
'        End If

        '' Finished printing that guy..
        Set oAdo = New clsADO
          With oAdo
              .ConnectionString = CodeConnString
              .SQLTextType = StoredProc
              .sqlString = "usp_LETTER_Automation_MAILOPS_FinishPrinting"
              .Parameters.Refresh
              .Parameters("@pRowId") = lBatchRowIdToClose
'              Set oRs = .ExecuteRS
                .Execute
              If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                    Stop
                    LogMessage strProcName, "ERROR", "There was a problem creating the Cover sheet!", sSelectedPrinter, True
                    GoTo Block_Exit
              End If
          End With

        oRs.MoveNext
    Wend
    LogMessage strProcName, "NOTICE", "Finished printing batchs!", vbCrLf & sMsg, True
    
Block_Exit:
    
    Set oFrmStatus = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Call RefreshData
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

'Private Sub cmdPrintChecked_Click_LEGACY()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sMsg As String
'Dim oLItem As ListItem
'Dim oLV As CustomControl
'Dim oLView As ListView
'Dim lSelected As Long
'Dim oAdo As clsADO
'Dim oRs As ADODB.RecordSet
'Dim lOpsBatchId As Long
'Dim strPrinterBefore As String
'Dim sSelectedPrinter As String
'Dim lBatchRowIdToClose As Long
'Dim lSelOpsBatchId As Long
'Dim varyBatchList() As Long
'Dim lIdx As Long
'
'
'    strProcName = ClassName & ".cmdPrintChecked_Click"
'
'    ' validate that they want to do this?
'    Set oLV = Me.lvQueue
'
'    For Each oLItem In oLV.ListItems
'        If oLItem.Checked = True Then
'
'            lSelected = lSelected + 1
'            sMsg = sMsg & " * " & CStr(oLItem.Text) & " - " & oLItem.SubItems(3) & vbCrLf
'            lOpsBatchId = CLng(oLItem.Text)
'            ReDim Preserve varyBatchList(lSelected - 1)
'            varyBatchList(lSelected - 1) = lOpsBatchId
''            Exit For    ' only allowing 1 at a time for now
''        ElseIf oLItem.Selected = True Then
''            lSelOpsBatchId = CLng(oLItem.Text)
'        End If
'    Next
'
'    If lOpsBatchId = 0 Then
'        Stop
'        lOpsBatchId = lSelOpsBatchId
'    End If
'
'    If lSelected > 1 Then
'        If sMsg <> "" Then
'            If MsgBox("Are you sure you wish to print the following to the same printer?" & vbCrLf & sMsg, vbYesNo, "Confirm") = vbNo Then
'                GoTo Block_Exit
'            End If
'        End If
'    End If
'
'    sSelectedPrinter = SelectPrinter(, strPrinterBefore)
'    DoCmd.Hourglass True
'
'
'    For lIdx = 0 To UBound(varyBatchList)
'        lOpsBatchId = varyBatchList(lIdx)
'
'
''Stop
'        '' Make the cover page (we need the details - let's get them from the database again
'        '' instead of reading from the cached list view which may not be up to date..
'        Set oAdo = New clsADO
'        With oAdo
'            .ConnectionString = CodeConnString
'            .SQLTextType = StoredProc
'            .sqlString = "usp_LETTER_Automation_MAILOPS_StartPrinting"
'            .Parameters.Refresh
'            .Parameters("@pMailBatchId") = lOpsBatchId
'            .Parameters("@pPrintOperator") = GetUserName
'            .Parameters("@pPrinterName") = sSelectedPrinter
'            Set oRs = .ExecuteRS
'            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'                Stop
'            End If
'            lBatchRowIdToClose = .Parameters("@pRowId").Value
'        End With
'
'        ' should I do all of them here (Should really only be 1 per MailBatchId.. (right??))
'        ' It's only been like 2 months since I last touched this code - whatdayawant? lol
'
'
'        If CreateCoverPageAndPrint(lOpsBatchId, oRs("CombinedFilePath").Value, lBatchRowIdToClose, sSelectedPrinter) = False Then
'            LogMessage strProcName, "ERROR", "There was a problem creating the Cover sheet!", sSelectedPrinter, True
'            GoTo Block_Exit
'            Stop
'        End If
'
'        '' Finished printing that guy..
'        Set oAdo = New clsADO
'          With oAdo
'              .ConnectionString = CodeConnString
'              .SQLTextType = StoredProc
'              .sqlString = "usp_LETTER_Automation_MAILOPS_FinishPrinting"
'              .Parameters.Refresh
'              .Parameters("@pRowId") = lBatchRowIdToClose
'              Set oRs = .ExecuteRS
'              If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'                  Stop
'LogMessage strProcName, "ERROR", "There was a problem creating the Cover sheet!", sSelectedPrinter, True
'            GoTo Block_Exit
'              End If
'              lBatchRowIdToClose = .Parameters("@pRowId").Value
'          End With
'
''        oRs.MoveNext
'
'    Next    ' next batch
'
'    LogMessage strProcName, "COMPLETE", "Finished printing the below batches!", vbCrLf & sMsg, True
'
'Block_Exit:
'
'    If Not oRs Is Nothing Then
'        If oRs.State = adStateOpen Then oRs.Close
'        Set oRs = Nothing
'    End If
'    Set oAdo = Nothing
'    Call RefreshData
'    DoCmd.Hourglass False
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

Private Sub cmdRefresh_Click()
    Call RefreshData
    Call LoadComboBoxes
End Sub


Private Sub dtSpecificDate_Updated(Code As Integer)
    SpecificDateSelected = True
End Sub

Private Sub Form_Load()
    DoCmd.Hourglass True
'    DoCmd.OpenForm "frm_ALERT_NOTE", acNormal, , , , acHidden
    
    Set coSettings = New clsSettings
    LoadMailRoomData
    Me.TimerInterval = 300000
    
    LoadComboBoxes
    Me.dtSpecificDate = Now()
    Call RefreshData
    DoCmd.Hourglass False
End Sub

Private Sub Form_Resize()
On Error GoTo Block_Exit
Dim strProcName As String

    strProcName = ClassName & ".Form_Resize"
    
    Me.lvQueue.Width = Me.InsideWidth - (Me.lvQueue.left * 2)
    
    Me.lvQErrorDetails.Width = Me.InsideWidth - (Me.lvQErrorDetails.left * 2)
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Timer()
Dim lOrigInterval As Long

    ' SHould check to see if there is a message from Business Operations
    ' to pull a letter or whatnot..
    

    lOrigInterval = Me.TimerInterval
    Me.TimerInterval = 0
    LoadMailRoomData
    Call RefreshData
    Me.TimerInterval = lOrigInterval
End Sub

Private Sub lvQErrorDetails_ColumnClick(ByVal ColumnHeader As Object)
    Call SortListView(Me.lvQErrorDetails, ColumnHeader)
End Sub

Private Sub lvQErrorDetails_DblClick()
On Error GoTo Block_Err
Dim strProcName As String
Dim sDocLoc As String
Dim sTempLoc As String
Dim oLI As ListItem
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document


    strProcName = ClassName & ".lvQErrorDetails_DblClick"
    
    Set oLI = lvQErrorDetails.SelectedItem

    ' We should open the affected letter in read only mode..
    sDocLoc = Me.QueueColumns.GetLiValue(oLI, "Errors", "PathForPrinting")

    sTempLoc = GetUniqueFilename(, , FileExtension(sDocLoc))
    If CopyFile(sDocLoc, sTempLoc, False) = False Then
        LogMessage strProcName, "ERROR", "Could not copy the sample document to your temp folder: " & sTempLoc, sDocLoc, True
        GoTo Block_Exit
    End If

    Set oWordApp = New Word.Application
        ' open it read only!!!!
    Set oWordDoc = oWordApp.Documents.Open(sTempLoc, , True)
    oWordApp.visible = True
    
    oWordApp.Activate
    
    Call AppActivate(oWordApp.ActiveDocument.Name, False)
    Call SendKeys("^a", False)
    Sleep 500
    Call SendKeys("%{F9}", False)
    oWordApp.Activate

Block_Exit:
    Set oLI = Nothing
    Set oWordDoc = Nothing
    Set oWordApp = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
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
Dim sOrigPath As String
Dim sTempPath As String
Dim sFileExt As String



    strProcName = ClassName & ".lvQueue_DblClick"

    
    ' find the one dbl clicked on
    Debug.Print Me.lvQueue.SelectedItem.Text
    
    Set oLI = Me.lvQueue.SelectedItem
    
    ' get the path of the document
'    If cdctLISubItems Is Nothing Then
'        Call RefreshQueue
'    End If
'    sOrigPath = oLI.SubItems(coLVCols.GetDetails("QUEUE", "COMBINEDFILEPATH"))
    sOrigPath = QueueColumns.GetLiValue(oLI, "QUEUE", "CombinedFilePath")
'    sOrigPath = oLI.SubItems(cdctLISubItems.Item("COMBINEDFILEPATH"))
    ' copy it to a temp directory
    sFileExt = FileExtension(sOrigPath)
    ' I forget if it has the period - but we don't want it
    If left(sFileExt, 1) = "." Then
        sFileExt = Right(sFileExt, Len(sFileExt) - 1)
    End If
    
    Shell "explorer.exe """ & sOrigPath & """", vbNormalFocus
    GoTo Block_Exit
    
'''    sTempPath = GetUniqueFilename(, , sFileExt)
'''    If CopyFile(sOrigPath, sTempPath, False) = False Then
'''        Stop
'''    End If
'''
'''    ' open it read only
'''    Set oWordApp = New Word.Application
'''    Set oWordDoc = oWordApp.Documents.Open(sTempPath, , True)
'''    oWordApp.visible = True
'''    ' bring it to the top window..
'''    Dim lHwnd As Long
'''
'''    lHwnd = FindWindow(vbNullString, oWordDoc.Name & " [Read-Only] - Microsoft Word")
'''    If lHwnd > 0 Then
'''        Call BringWindowToTop(lHwnd)
'''    Else
'''        lHwnd = FindWindow(vbNullString, oWordDoc.Name & " [Read-Only] [Compatibility Mode] - Microsoft Word")
'''        If lHwnd > 0 Then
'''            Call BringWindowToTop(lHwnd)
'''        Else
'''            Stop
'''        End If
'''    End If
''''    Call ActivateApplicationWindow(, oWordDoc.Name)
    
Block_Exit:
    ' not going to quit the app here because we want the user to do that..
'    Set oWordDoc = Nothing
'    Set oWordApp = Nothing
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
    
    If Me.SelectedAccountId = 0 Then
        Me.SelectedAccountId = 1
    End If
    
'    sCriteria = "LetterType = '" & oLI.SubItems(coLVCols.GetDetails("QUEUE", "LetterType")) & "' "
'    sCriteria = sCriteria & " AND AccountId = " & CStr(Me.SelectedAccountId)
'    sCriteria = sCriteria & " AND LetterReqDt = '" & oLI.SubItems(coLVCols.GetDetails("QUEUE", "LetterReqDt")) & "'"
'
'    sCriteria = sCriteria & " AND OpsLetterBatchId = '" & oLI.Text & "'"

    sCriteria = "LetterType = '" & QueueColumns.GetLiValue(oLI, "QUEUE", "LetterType") & "' "
    sCriteria = sCriteria & " AND AccountId = " & CStr(Me.SelectedAccountId)
    sCriteria = sCriteria & " AND LetterReqDt = '" & QueueColumns.GetLiValue(oLI, "QUEUE", "LetterReqDt") & "'"
    
    sCriteria = sCriteria & " AND OpsLetterBatchId = '" & QueueColumns.GetLiValue(oLI, "QUEUE", "OpsLetterBatchId") & "'"
        
    
    '' Load these in the below grid..
    DoCmd.Hourglass True
    sSql = "SELECT * FROM v_LETTER_Automation_MAILOPS_TodayQueueDetails WHERE " & sCriteria
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
'Stop
        Set oRs = .ExecuteRS
    End With
    
    Call PopulateListView(Me.lvQErrorDetails, oRs)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set oLI = Nothing
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub lvQueue_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem

    strProcName = ClassName & ".lvQueue_DblClick"


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
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub LoadComboBoxes()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".LoadComboBoxes"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_MAILOPS_Filters"
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
   
   
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_MAILOPS_Filters"
        .Parameters.Refresh
        .Parameters("@pFilterName") = "Ops BatchIds"
        Set oRs = .ExecuteRS
        If .GotData = True Then
            Call RefreshComboBoxFromRecordset(oRs, Me.cmbOpsBatchId)
        
        End If
    End With
    
    
   
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



'Private Function GetColumnMaxWidth(ctl As ListView, col As Long) As Long
'' Loop through passed Column and calculate the
'' width of the largest string for all rows of this column.
'
'    ' Junk var
'    Dim lngRet As Long
'
'    ' Create our Font
'    Dim myfont As LOGFONT
'    Dim lngscreenXdpi As Long
'    Dim fontsize As Long
'    Dim hfont As Long, prevhfont As Long
'    Dim hdc As Long
'    Dim hDC2 As Long
'
'    ' Calc size of the string
'    Dim strText As String
'    Dim lngLength As Long
'    Dim stfSize As Size
'
'    ' Loop through the rows of the ctl
'    Dim ctr As Long
'    Dim MaxWidth As Long
'
'    ' Get Desktop's Device Context
'    hDC2 = apiGetDC(0&)
'    ' Create a compatible DC
'    hdc = CreateCompatibleDC(hDC2)
'
'    ' Release the handle to the Desktop DC
'    lngRet = apiReleaseDC(0&, hDC2)
'
'    'Get Current Screen Twips per Pixel
'    lngscreenXdpi = GetDPI()
'
'    ' Build our LogFont structure.
'    ' This  is required to create a font matching
'    ' the font selected into the Control we are passed
'    ' to the main function.
'    'Copy font stuff from Control's property sheet
'
'
'    With myfont
'        .lfFaceName = ctl.Font.Name & Chr$(0)  'Terminate with Null
'        fontsize = ctl.Font.Size
'        .lfWeight = ctl.Font.Weight
'        .lfItalic = ctl.Font.Italic
'        .lfUnderline = ctl.Font.UnderLine
'
'        ' Must be a negative figure for height or system will return
'        ' closest match on character cell not glyph
'        .lfHeight = (fontsize / 72) * -lngscreenXdpi
'    End With
'
'    ' Create our Font
'    hfont = apiCreateFontIndirect(myfont)
'    ' Select our Font into the Device Context
'    prevhfont = apiSelectObject(hdc, hfont)
'
'    ' Loop through all of the rows in the ListBox
'    ' for the given Column(col) and row(ctr)
'
'    ' Reset our max width var
'    MaxWidth = 0
'
''setup to make this handle empty controls. ' KD Comeback and fix this!!! 20130325
'    Dim i As Long
'    If (ctl.ListItems.Count = 0) Then
'        i = 1
'    Else
'        i = ctl.ListItems.Count
'    End If
'
'
'    For ctr = 0 To i - 1
'    'For ctr = 0 To ctl.ListCount - 1
'
'        strText = ctl.Column(col, ctr)
'
'        ' Let's get the width of output string
'        lngLength = Len(strText)
'        lngRet = apiGetTextExtentPoint32(hdc, strText, lngLength, stfSize)
'
'        ' Now compare with last result and save larger value
'        If stfSize.cx > MaxWidth Then MaxWidth = stfSize.cx
'    Next ctr
'
'    ' Select original Font back into DC
'    hfont = apiSelectObject(hdc, prevhfont)
'
'    ' Delete Font we created
'    lngRet = apiDeleteObject(hfont)
'
'    ' Release the DC
'    lngRet = apiDeleteDC(hdc)
'
'    ' Return the Height of the String in Twips
'    GetColumnMaxWidth = MaxWidth * (1440 / GetDPI())
''    strText = ctl.column(col, 0)
''ctl.ColumnWidths = "123;123"
''    MsgBox (ctl.column(col, 0).Value)
'  'ctl.ColumnWidths = Nz(ctl.ColumnWidths, "") & GetColumnMaxWidth & ";"
'
'End Function
'
'
'
'Private Function GetDPI() As Integer
'
'    ' Determine how many Twips make up 1 Pixel
'    ' based on current screen resolution
'
'    Dim lngIC As Long
'    lngIC = apiCreateIC("DISPLAY", vbNullString, _
'     vbNullString, vbNullString)
'
'    ' If the call to CreateIC didn't fail, then get the info.
'    If lngIC <> 0 Then
'        GetDPI = GetDeviceCaps(lngIC, LOGPIXELSX)
'        ' Release the information context.
'        apiDeleteDC lngIC
'    Else
'        ' Something has gone wrong. Assume a standard value.
'        GetDPI = 96
'    End If
' End Function
'
'
'
