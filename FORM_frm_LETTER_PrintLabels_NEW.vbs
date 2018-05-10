Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




''' Last Modified: 08/27/2013
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This Print Letter Form is to take all the letters selected and generate letters.
'''     Status: 'W' means it's been queued up and ready to be generated
'''             'G' means it's been generated
'''             'P' means it's been printed
'''             'R' means it's ready to be re-generated (then re-printed)
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 08/05/2013 - KD: Added sorting to print the letters with the most pages last
'''     in order to make sure that the operator doesn't have to baby sit the envelope
'''     stuffer
'''     Also, added a mail merge for letter lables for the documents with > 10 (or was it 8) pages
'''     which have to be sent in manila envelopes so operator doesn't have to manually type them up
'''
'''
'''  - 04/16/2013 - added logging throughout and more tweaking to barcode stuff
'''         also minor cleanup of some of the code (ado stuff was early bound, badly
'''         indented, etc..)
'''  - 03/28/2013 - added some logging and very minor tweaking to the barcode stuff..
'''  - 02/19/2013 - Added AddSecPagesCode and some minor other adjustments for adding
'''     bar codes to letter's footer for the envelop stuffer
'''  - ?????? - Created...
'''
''' AUTHOR
'''  =====================================
''' ????? Thieu ??????
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################

'
Private mstrAuditor As String

Private Const cs_USER_TEMPLATE_PATH_ROOT As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\"
Private Const cs_LABEL_FOLDER_PATH As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\Label_datasets\"

Private cdctSelectedLetters As clsLetterInstanceDct

'for resizing
Private ColResize1 As clsAutoSizeColumns
Private ColReSize2 As clsAutoSizeColumns
Private lngQueryType As Long '* These are values from msysobjects 1/4/6 = table, 5 = query


Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents frmFilter As Form_frm_GENERAL_Filter
Attribute frmFilter.VB_VarHelpID = -1

Private mReturnDate As Date
Private msAdvancedFilter As String

Private clBatchId As Long
Private clBatchColumn As Long
Private cdctQueueColumns As Scripting.Dictionary
Private cdctBatches As Scripting.Dictionary

Const CstrFrmAppID As String = "LetterQueuePrint"

Private ciBatchToSelect As Integer



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get BatchToSelect() As Integer
    BatchToSelect = ciBatchToSelect
End Property
Public Property Let BatchToSelect(iBatchToSelect As Integer)
    ciBatchToSelect = iBatchToSelect
    Me.TimerInterval = 500
End Property




Public Property Get FilteredByUngenerated() As Boolean
    If UCase(Me.cboViewType) = "VIEW UN-PROCESSED LETTERS" Then
        FilteredByUngenerated = True
    Else    'If UCase(Me.cboViewType) = "VIEW PROCESSED LETTERS" Then
        FilteredByUngenerated = False
    End If
End Property

Public Property Get FilteredByGenerated() As Boolean
    If UCase(Me.cboViewType) = "VIEW PROCESSED LETTERS" Then
        FilteredByGenerated = True
    Else
        FilteredByGenerated = False
    End If
End Property

Public Property Let NumSelected(lNumberSelected As Long)
    Me.lblNumSelected.Caption = Format(Nz(lNumberSelected, 0), "###,##0") & " items selected"
End Property


Public Property Get MostRecentBatchId() As Long
    MostRecentBatchId = clBatchId
End Property
Public Property Let MostRecentBatchId(lBatchId As Long)
    clBatchId = lBatchId
    Me.lblRecentBatchId.Caption = "Most recent batchid: " & CStr(clBatchId)
    If clBatchId = 0 Then
        Me.lblRecentBatchId.visible = False
    Else
        Me.lblRecentBatchId.visible = True
    End If
End Property


Private Sub cboAuditor_AfterUpdate()
    If Me.cboAuditor <> "View All" Then
        Me.cboAuditor.ForeColor = 16711680
    Else
        Me.cboAuditor.ForeColor = 0
    End If
    RefreshMain
End Sub

Private Sub cboLetterType_AfterUpdate()
    If Me.cboLetterType <> "View All" Then
        Me.cboLetterType.ForeColor = 16711680
    Else
        Me.cboLetterType.ForeColor = 0
    End If
    RefreshMain
End Sub

Private Sub cboViewType_AfterUpdate()
    Me.cboViewType.ForeColor = 0


    ' 20130815 KD: Ok, so, if they choose Unprocessed then we are going to change the
    ' 'View Letter(s)' button to read: PRE-view Letter(s)'
    ' otherwise we'll say Print Letters (or something like that..)
    Select Case UCase(Me.cboViewType)
    Case "VIEW UN-PROCESSED LETTERS"
        RefreshMain
        Me.cmbBatches = ""
    
    Case "VIEW PROCESSED LETTERS"
        RefreshMain
        
        If Me.MostRecentBatchId <> 0 Then
            Me.cmbBatches = Me.MostRecentBatchId
            Call cmdSelectByBatch_Click
        Else
            Me.cmbBatches = ""
        End If

    Case "View Errors"
        Stop
    End Select

    
End Sub



Private Sub cmbBatches_Change()
    Call cmdSelectByBatch_Click
End Sub


Private Sub cmdEditThreshold_Click()
    Me.txtPageThreshold.Enabled = True
    Me.txtPageThreshold.Locked = False
    Me.txtPageThreshold.SetFocus
End Sub

Private Sub cmdPrintLabels_Click()
On Error GoTo Block_Err
Dim strProcName As String

    ' KD: Come back here and do the following:
    ' Get the selected items
    ' construct a filter for each of them
    ' open the report (rpt_LETTER_Static_Details in preview mode with that filter
    ' somehow mark these as printed???
'    DoCmd.OpenReport "rpt_LETTER_Static_Details", acViewPreview, , "[LetterBatchId] = 6", acWindowNormal
    
' Should probably save these in a table and join them for the report..
' Yeah, we'll use a local table since it's a throw-away and user-specific

Dim sInstanceIds As String
Dim oLetter As clsLetterInstance
Dim sErrMsg As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim vLabelsToSkip As Variant
Dim iSkip As Integer
    strProcName = ClassName & ".cmdPrintLabels_Click"
    
      ' We already have our selected letters:
    Set cdctSelectedLetters = GetSelectedItems(True, , , sErrMsg)
    If cdctSelectedLetters Is Nothing Then
        LogMessage strProcName, "WARNING", "No items selected!"
        GoTo Block_Exit
    End If

    Set oDb = CurrentDb()
    'oDb.Execute "DELETE FROM LETTER_Label_Instances WHERE User = '" & Identity.UserName & "'"
    oDb.Execute "DELETE FROM LETTER_Label_Instances "
    
    Set oRs = oDb.OpenRecordSet("SELECT * FROM LETTER_Label_Instances WHERE User = '" & Identity.UserName & "'")
    
    '' Now, if we want to add some blanks to accomodate for partially printed label sheets
    '' we do that here..
    If MsgBox("Is this a brand new label sheet?", vbYesNo, "New Label sheet?") = vbNo Then
        vLabelsToSkip = InputBox("How many labels should we skip on the FIRST sheet?" & vbCrLf & "(Count from the top, left to right...)", "Skip how many?")
        If IsNumeric(vLabelsToSkip) = False Then
            ' uh.. stop
            LogMessage strProcName, "ERROR", "We expected a number but got something else!?!", CStr(vLabelsToSkip), True
            GoTo Block_Exit
        End If
        ' just check that we aren't skipping more than are on a sheet
        ' NOTE: this is based on Avery label template: 5163, 2" X 4" (2 columns, 10 total)
    
        If vLabelsToSkip > 9 Then
            LogMessage strProcName, "ERROR", "There are only 10 Labels on each AVERY Lable # 5163!", , True
            GoTo Block_Exit
        End If
        
        For iSkip = 1 To CInt(vLabelsToSkip)
            oRs.AddNew
            oRs("InstanceId") = ""
            oRs("DateAdded") = Now()
            oRs("User") = Identity.UserName
            oRs.Update
        Next
    End If

    For Each oLetter In cdctSelectedLetters.Letters
        oRs.AddNew
        oRs("InstanceId") = oLetter.InstanceId
        oRs("DateAdded") = Now()
        oRs("User") = Identity.UserName
        oRs.Update
    Next

    DoCmd.OpenReport "rpt_LETTER_Static_Details_NEW", acViewPreview, , , acWindowNormal
    
Block_Exit:
    oRs.Close
    Set oRs = Nothing
    Set oDb = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSelectByBatch_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oListBox As listBox
Dim lRow As Long
Dim vItem As Variant
Dim lSelBatchid As Long

    strProcName = ClassName & ".cmdSelectByBatch_Click"
    'Call SelectComboBoxItemFromText(Me.cmbBatches, CStr(Me.MostRecentBatchId))
    
    lSelBatchid = Me.cmbBatches
    If cdctQueueColumns Is Nothing Then Call GetQueueColumns
                
    clBatchColumn = cdctQueueColumns.Item("BATCHID")
    For lRow = 0 To Me.lstQueue.ListCount
        If Me.lstQueue.Column(clBatchColumn, lRow) = lSelBatchid Then
            Me.lstQueue.Selected(lRow) = True
        Else
            Me.lstQueue.Selected(lRow) = False
        End If
    Next
    
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub GetQueueColumns()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetQueueColumns"
  
    Set cdctQueueColumns = New Scripting.Dictionary
    Set oRs = Me.lstQueue.RecordSet
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)
    
Block_Exit:
    Set oRs = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    
    If Me.BatchToSelect() <> 0 Then
        Me.cmbBatches = Me.BatchToSelect
        Call cmdSelectByBatch_Click
    End If

End Sub

Private Sub frmfilter_QueryFormRefresh()

    RefreshMain

End Sub

Private Sub frmFilter_UpdateSql()
    msAdvancedFilter = frmFilter.SQL.WherePrimary
End Sub



Private Sub cmdEndDate_Click()
    On Error GoTo Exit_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtThroughDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtThroughDate = mReturnDate
    
    txtThroughDate_Exit False

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub



Private Sub cmdRefresh_Click()
    RefreshMain
End Sub


Private Sub cmdSelectEntireQueue_Click()
    Dim idx As Integer

    For idx = 1 To Me.lstQueue.ListCount
        Me.lstQueue.Selected(idx) = True
    Next idx
    ' could add in detail view but left off b/c assuming analyst are just selecting all to print in batch processes - TH 10-8-07
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub

Private Sub cmdStartDate_Click()
    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtFromDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtFromDate = mReturnDate
    
    txtFromDate_Exit False

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub


Public Sub RefreshMain()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim sCmboSql As String

    strProcName = ClassName & ".RefreshMain"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintLabels"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
'        .Parameters("@pManualOnly") = Nz(Me.fraManualOnly, 0)
        .Parameters("@pProcessFromDt") = Format(Me.txtFromDate, "m/d/yyyy")
        .Parameters("@pProcessThruDt") = Format(Me.txtThroughDate, "m/d/yyyy")
        .Parameters("@pProcessType") = Me.cboViewType
        .Parameters("@pLetterType") = Me.cboLetterType
'        .Parameters("@pAuditor") = Me.cboAuditor
        .Parameters("@pPageLimit") = CStr(Nz(Me.txtPageThreshold, 9))
        
        If Me.tglAdvancedFilter = True Then
            .Parameters("@pExtraFilter") = msAdvancedFilter
        Else
            .Parameters("@pExtraFilter") = ""
        End If
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Problem getting the data: " & .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        Else
            Set Me.lstQueue.RecordSet = oRs
            Me.lstQueue.ColumnCount = oRs.Fields.Count
        End If
    End With
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)

    Me.cboLetterType.RowSource = "SELECT DISTINCT LetterType FROM Letter_Type WHERE AccountId = " & CStr(gintAccountID) & " AND IsLetterActive = 1 AND For_DS_Only = 0 "
            
    sCmboSql = AccessSqlToSqlServer("SELECT DISTINCT BatchId, QueueRunId FROM v_Letter_Print_Queue ORDER BY QueueRunId DESC ")
    
    Call RefreshComboBoxADO(sCmboSql, Me.cmbBatches, CStr(Me.MostRecentBatchId), , "v_Code_Database")

Block_Exit:
    'Auto sizing
    Set ColResize1 = New clsAutoSizeColumns
    ColResize1.SetControl Me.lstQueue
    'don't resize if lstclaims is null
    If Me.lstQueue.ListCount > 1 Then
        ColResize1.AutoSize
    End If
    Set oAdo = Nothing
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub RefreshMain_LEGACY()
On Error GoTo Block_Err
Dim strProcName As String
Dim dtThrouDt As Date
Dim strSelectedAuditor As String
Dim strSelectedLetterType As String
Dim sAdoSql As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrder As String
Dim sCmboSql As String

    strProcName = ClassName & ".RefreshMain"

    dtThrouDt = DateAdd("d", 1, txtThroughDate.Value)
    
    sSelect = "SELECT SD.PageCount, Q.*, SD.BatchId "
    
    sFrom = " FROM LETTER_Work_Queue Q INNER JOIN ( SELECT D.InstanceId, Max(LetterBatchId) AS BatchId, D.PageCount FROM LETTER_Static_Details D GROUP BY D.InstanceID, D.PageCount ) SD " & _
            " ON Q.InstanceId = SD.InstanceID "
    
    sWhere = " WHERE ProcessedDt >= #" & Nz(txtFromDate.Value, "01/01/1900") & "# and ProcessedDt < #" & _
            Format(dtThrouDt, "mm-dd-yyyy") & "# "

    sWhere = sWhere & " AND Status IN ('P','G') "
    
    sWhere = sWhere & " AND SD.PageCount > " & CStr(Nz(Me.txtPageThreshold, 9))
    
    
    strSelectedAuditor = Nz(cboAuditor.Value, "")   ' 20121010 KD: fixed this..
    
    If strSelectedAuditor <> "View All" And strSelectedAuditor <> "" Then
'                sQueueRowSource = sQueueRowSource & " and Auditor = " & Chr(34) & strSelectedAuditor & Chr(34)
        sWhere = sWhere & " AND Auditor = '" & strSelectedAuditor & "' "
    End If
    
    strSelectedLetterType = Nz(cboLetterType.Value, "")   ' 20121010 KD: fixed this..
    
    If strSelectedLetterType <> "View All" And strSelectedLetterType <> "" Then
'                sQueueRowSource = sQueueRowSource & " and LetterType = " & Chr(34) & strSelectedLetterType & Chr(34)
        
        sWhere = sWhere & " AND LetterType = '" & strSelectedLetterType & "' "
    End If
    
    If Me.tglAdvancedFilter = True Then
'                    sQueueRowSource = sQueueRowSource & "AND (" & msAdvancedFilter & ")"
        sWhere = sWhere & " AND ( " & msAdvancedFilter & " ) "
    End If
    
    sOrder = " ORDER BY SD.PageCount ASC, LetterType, CnlyProvId, LetterReqDt"
    
'    sAdoSql = AccessSqlToSqlServer(sQueueRowSource1 & sQueueRowSource & " order by lettertype, cnlyProvId, LetterReqdt")
    sAdoSql = AccessSqlToSqlServer(sSelect & sFrom & sWhere & sOrder)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = sAdoSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
'            Stop
'                        Me.lstQueue.RowSource = sQueueRowSource1 & sQueueRowSource & " order by lettertype, cnlyProvID, letterreqdt"
            Me.lstQueue.RowSource = sSelect & sFrom & sWhere & sOrder
        Else
            Set Me.lstQueue.RecordSet = oRs
        End If
    End With
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)


    Me.cboAuditor.RowSource = "Select UserID, OrderValue from (SELECT TOP 1 'View All' AS UserID, 1 AS OrderValue FROM LETTER_Work_Queue) " & _
                                " UNION (Select Auditor as UserID, 2 AS OrderValue " & sFrom & sWhere & " ) order by OrderValue, UserID; "

    
    Me.cboLetterType.RowSource = "Select LetterType, OrderValue from (SELECT TOP 1 'View All' AS LetterType, 1 AS OrderValue FROM LETTER_Work_Queue  " & _
                                " UNION Select LetterType as UserID, 2 AS OrderValue " & sFrom & sWhere & " ) As A order by OrderValue, LetterType; "


    ' Ok, so for the batches we only want to load the combo box with the batches that are actually in the list so we need to structure our query a little differently
    sCmboSql = AccessSqlToSqlServer("SELECT DISTINCT SD.LetterBatchId, SD.UserId FROM LETTER_WORK_Queue Q INNER JOIN (SELECT D.InstanceId, MAx(LetterBatchId) AS LetterBatchId, " & _
        " D.Auditor as UserId, D.PageCount FROM LETTER_Static_Details D GROUP BY D.InstanceId, D.Auditor, D.PageCount ) SD ON Q.InstanceID = SD.InstanceID " & _
        sWhere & " AND SD.InstanceId IS NOT NULL " & _
            " ORDER BY LetterBatchID DESC")
    Call RefreshComboBoxADO(sCmboSql, Me.cmbBatches, CStr(Me.MostRecentBatchId))


Block_Exit:
    'Auto sizing
    Set ColResize1 = New clsAutoSizeColumns
    ColResize1.SetControl Me.lstQueue
    'don't resize if lstclaims is null
    If Me.lstQueue.ListCount > 1 Then
        ColResize1.AutoSize
    End If
    Set oAdo = Nothing
    
    Exit Sub

Block_Err:
'    MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
    Else
        lblAppTitle.visible = True
    End If
    Me.lblNumSelected.Caption = " "
    
    Me.tglAdvancedFilter = False
    Me.tglAdvancedFilter.Caption = "Add Filter"
    
    ' KD Comeback: Should really get the most recent batch id
    ' but for now we'll just zero it out which will hide the label
    MostRecentBatchId = 0
    
    '' Start off on Process Type: 'View Processed Letters'
'    Me.cboViewType = "View Processed Letters"
    Me.cboViewType = "View Un-processed Letters"
    Call cboViewType_AfterUpdate
    
    RefreshMain
    
End Sub

Private Sub lblRecentBatchId_DblClick(Cancel As Integer)

    DoCmd.OpenReport "rpt_Letter_Static_Details", acViewPreview, , "[LetterBatchId] = " & CStr(Me.lblRecentBatchId)

    
End Sub

Private Sub lstQueue_AfterUpdate()
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub




Private Sub lstQueue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub

Private Sub tglAdvancedFilter_Click()

    If Me.tglAdvancedFilter.Value = True Then
    
        Set frmFilter = New Form_frm_GENERAL_Filter
        
        With frmFilter
            .CalledBy = Me.Name
            .FieldsTable = "LETTER_Work_Queue"
            .Setup
            .visible = True
        End With

        Me.tglAdvancedFilter.Caption = "Filter On"

    Else
        Me.tglAdvancedFilter.Caption = "Add Filter"
        msAdvancedFilter = ""
        RefreshMain
    End If


End Sub

Private Sub txtFromDate_Exit(Cancel As Integer)
    If Not (IsDate(txtFromDate.Value) Or IsNull(txtFromDate.Value)) Then
        MsgBox "Please enter a valid from date"
        txtFromDate.SetFocus
    Else
        If Me.txtFromDate > Me.txtThroughDate Then
            Me.txtThroughDate = Me.txtFromDate
        End If
        RefreshMain
    End If
End Sub

Private Sub txtPageThreshold_AfterUpdate()
    Me.lstQueue.SetFocus
    Me.txtPageThreshold.Locked = True
    Me.txtPageThreshold.Enabled = False

    Call RefreshMain
End Sub

Private Sub txtThroughDate_Exit(Cancel As Integer)
    If Not IsDate(txtThroughDate.Value) Then
        MsgBox "Please enter a valid through date"
        txtThroughDate.SetFocus
    Else
        If Me.txtThroughDate < Me.txtFromDate Then
            Me.txtFromDate = Me.txtThroughDate
        End If
        RefreshMain
    End If
End Sub

Private Function JoinPDFs(TargetPDF As String, FileToAppend As String) As Boolean 'This one works!
    Dim Project_Folder As String
    Dim fs As Scripting.FileSystemObject 'FileSystemObject - late bind this
    Dim AcroApp As Object
    Dim PDF_TargetFile As Object
    Dim PDF_DataSheet As Object
    Dim RowNr As Integer
    Dim iPathLen As Integer
    Dim j As Integer
    Dim i As Integer
    Dim strChkPath As String
    
    Project_Folder = left(TargetPDF, InStrRev(TargetPDF, "\") - 1)
    Set fs = New Scripting.FileSystemObject
    
    '--------------
    'test if folder exists if it does skip this logic...
    If Not fs.FolderExists(Project_Folder) Then
            'TGH taken from Damon - Added condition to handle UNC paths + check if folders exist
            iPathLen = Len(Project_Folder)
            If InStr(1, Project_Folder, "\\") > 0 Then  'in the UNC case
                'Start past the "\\cca-audit\"
                j = 12
                Do
                    i = InStr(j + 1, Project_Folder, "\")
                    If i > 0 Then
                        strChkPath = left(Project_Folder, i)
                        If Not fs.FolderExists(strChkPath) Then
                            fs.CreateFolder strChkPath
                        End If
                        j = i
                    Else
                        j = iPathLen
                    End If
                Loop Until j = iPathLen
             Else
                Do
                    i = InStr(j + 1, Project_Folder, "\")
                    If i > 0 Then
                        strChkPath = left(Project_Folder, i)
                        If Not fs.FolderExists(strChkPath) Then
                            fs.CreateFolder strChkPath
                        End If
                        j = i
                    Else
                        j = iPathLen
                    End If
                Loop Until j = iPathLen
            End If
    End If 'End of loop for when the destination folder does not exist
    '-------------
    'Now we know the folders all exist...
    'Need to create a preview PDF to append all the new PDF's into.
     
    Set AcroApp = CreateObject("AcroExch.App") 'works
    AcroApp.Hide
    Set PDF_TargetFile = CreateObject("AcroExch.PDDoc") 'This is the header file
    Set PDF_DataSheet = CreateObject("AcroExch.PDDoc") 'This will be each datasheet in turn
     
     
    If Not fs.FileExists(TargetPDF) Then
        PDF_TargetFile.Create
        PDF_TargetFile.Save 1, TargetPDF
    End If

     'Open the already created header file
  
    PDF_TargetFile.Open TargetPDF
    PDF_DataSheet.Open FileToAppend 'Project_Folder & "Test2.pdf"
     
    'Open the source document that will be added to the destination
  If PDF_TargetFile.InsertPages(PDF_TargetFile.GetNumPages - 1, PDF_DataSheet, 0, PDF_DataSheet.GetNumPages, 0) Then
    JoinPDFs = True
  Else
    JoinPDFs = False
    MsgBox ("Failure! Joining PDF's")
  End If
  
  PDF_TargetFile.Save 1, TargetPDF
  
  PDF_DataSheet.Close
  PDF_TargetFile.Close
  AcroApp.Exit 'works
    Set fs = Nothing
End Function


Private Function ViewLetters(TotalRecs As Integer, cn As Variant, fmrStatus As Form_ScrStatus) As Boolean
'This funciton has been updated by 'TGH 9-26-08
' I added in the option based off of the boolean PreviewViewAllowed to view non printed letters.

    On Error GoTo Error_Encountered:
    ViewLetters = True
    
    'make sure we have some items selected.
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected"
        ViewLetters = False
        Exit Function
    End If
          
'Variable declarations ------------------------------------------------------------------------
'    Dim Person As New ClsIdentity
Dim strErrMsg As String: strErrMsg = ""
Dim Usedword As Boolean
Dim db As Database
Dim rsLetterConfig As DAO.RecordSet
Dim strInstanceID As String
Dim strStatus As String
Dim strOutputFileName As String
Dim varItem As Variant
Dim bFirstLetter As Boolean
Dim iCnt As Integer
Dim strOutputPath As String
Dim strAuditor As String
Dim colLetters As Collection    'thieu 1/16/08
Dim bnewwordfile As Boolean     'thieu 1/16/08
Dim bnewpdffile As Boolean
Dim ShowPreview As Boolean
Dim curLtr As String
Dim strPreviewFileName As String
Dim sMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
'    ' ADO variables
Dim cmdGetLetter As ADODB.Command
    Set cmdGetLetter = New ADODB.Command
Dim ProgVal As Integer
Dim strSQLcmd As String
'pdf objects
Dim TargetPDF As String: TargetPDF = ""
Dim PassFail As Boolean

Dim objWordApp, _
    objMasterDoc, _
    objTemplateDoc ', objWordSelection
    
'    Dim objWordApp As Word.Application, _
'        objMasterDoc As Word.Document, _
'        objTemplateDoc As Word.Document ', objWordSelection
    
    
Dim strTemplateLoc As String
Dim CountCurLetter As Integer
Dim i As Long
Dim MyRecordset As DAO.RecordSet

    Set objWordApp = CreateObject("word.application")
'    Set objWordApp = New Word.Application
    
        ' Create an instance of objWord, and make it invisible.
    objWordApp.visible = False

    'End of declarations ------------------------------------------------------------------------

    'assign the preview from the config table
    Set db = CurrentDb
    
    'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    strOutputPath = rsLetterConfig("LetterOutputLocation").Value

    'Setup for letter viewing.  setting auditor and clearing out the preview folder
    ShowPreview = False
    Usedword = False
    strAuditor = Replace(Identity.UserName, ".", "")              ' thieu 1/16/08
    strOutputPath = strOutputPath & "\PREVIEW\" & strAuditor    'thieu 1/16/08
    DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
    CreateFolder (strOutputPath & "\")                                ' thieu 1/16/08
    Set colLetters = New Collection                             ' thieu 1/16/08
    bnewwordfile = False                                        ' thieu 1/16/08
    'TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.column(2, 1)) & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
    bnewpdffile = False
    
    'this is the first time through the view letter.
    bFirstLetter = True
    On Error GoTo Error_Encountered
        
    'set to the formstatus current progress
    iCnt = 0
    ProgVal = fmrStatus.ProgVal
    cmdGetLetter.ActiveConnection = cn
    cmdGetLetter.commandType = adCmdStoredProc
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("LetterName", adChar, adParamOutput, 255, "")
    cmdGetLetter.CommandText = "usp_LETTER_Get_Letter_Name"

    curLtr = ""
    CountCurLetter = 0
    

    Set MyRecordset = Me.lstQueue.RecordSet
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    For Each varItem In lstQueue.ItemsSelected
            'grab the instance id and status from the listqueue
            strInstanceID = Trim(Me.lstQueue.Column(MyRecordset.Fields("instanceID").OrdinalPosition, varItem))
            strStatus = Trim(Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem))
            
        If UCase(strStatus) = "P" Then
            'mark that the preview window should be shone b/c there is at least one instance marked as 'print'
            ShowPreview = True
            CountCurLetter = CountCurLetter + 1
            'lookup this instanceID and grag the appropriate LetterName
            cmdGetLetter.Parameters("InstanceID").Value = strInstanceID
            cmdGetLetter.Execute ' executing "usp_LETTER_Get_Letter_Name"
            'see the filename back to see how we want to handle the file pdf/word.
            strOutputFileName = Trim(cmdGetLetter.Parameters("LetterName").Value)
                
            'check current letter vs the last P record we looked at.
            'First time through CurLtr = "" so we don't enter this statement
            If StrComp(curLtr, "", vbTextCompare) = 1 And _
                curLtr <> Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem)) Then
                    If InStr(1, strOutputFileName, ".doc") > 0 Then
                        bnewwordfile = True 'we should start a new word file for a different word letter.
                    Else
                        TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.Column(2, varItem)) & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
                        bFirstLetter = True
                        bnewwordfile = False        'thieu
                    End If
            End If
            
            curLtr = Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem))
            
            'At this point we have marked if the letter is a new or if we have changed letters
            ' TEST if this is a PDF or Word Doc.
            If InStr(1, strOutputFileName, ".PDF", vbTextCompare) = 0 Then 'we are NOT dealing with a pdf document (so word doc unless we add more)...
                'set the flag to show word documents at the end and close open word app
                Usedword = True

                'if we have a new letter close the old one and save it.
                If bnewwordfile Then                        ' thieu 1/16/08
                    objMasterDoc.spellingchecked = True 'this is needed to shut down popup for too many spelling errors
                    Sleep 1000
                    objMasterDoc.SaveAs strPreviewFileName
                    objMasterDoc.Close
                    Set objMasterDoc = Nothing
                    strPreviewFileName = strOutputPath & "\" & curLtr & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    bFirstLetter = True         'mark this so we exit this clause until we are ready for a new word file
                    bnewwordfile = False        'thieu
                End If

                If bFirstLetter = True Then
                    strOutputFileName = Trim(cmdGetLetter.Parameters("LetterName").Value)
                    'need to get template's name
                    strTemplateLoc = DLookup("TemplateLoc", "LETTER_Type", "Lettertype = '" & curLtr & "'")
                    'reset the counter for how many for this letter
                    CountCurLetter = 0
                    
                    'Now check if we are dealing with a word doc or a PDF image.
                    If InStr(1, strOutputFileName, ".doc", vbTextCompare) > 0 Then
                        'set the word documents and begin setup...
                        'Set objMasterDoc = objWordApp.Documents.Add()   'tgh 3/18/08
                        Set objMasterDoc = objWordApp.Documents.Open(strOutputFileName)
                            'objMasterDoc
                        
                        'Set objTemplateDoc = objWordApp.Documents.Open(strTemplateLoc)
                        '    objMasterDoc.PageSetup.LeftMargin = objTemplateDoc.PageSetup.LeftMargin
                        '    objMasterDoc.PageSetup.RightMargin = objTemplateDoc.PageSetup.RightMargin
                        '    objMasterDoc.PageSetup.TopMargin = objTemplateDoc.PageSetup.TopMargin
                        '    objMasterDoc.PageSetup.BottomMargin = objTemplateDoc.PageSetup.BottomMargin
                         '   objMasterDoc.PageSetup.HeaderDistance = objTemplateDoc.PageSetup.HeaderDistance
                        '    objMasterDoc.PageSetup.FooterDistance = objTemplateDoc.PageSetup.FooterDistance
                        '    objMasterDoc.spellingchecked = True
                        '    objMasterDoc.showspellingerrors = False
                        '    objMasterDoc.showgrammaticalerrors = False
                        'objTemplateDoc.Close
                        objWordApp.ActiveDocument.spellingchecked = True
                        'objWordApp.Selection.InsertFile (strOutputFileName)
                        'viewing letter name is based on time, can use instance_id for first letter if we want...
                        objWordApp.selection.EndKey Unit:=wdStory
                        strPreviewFileName = strOutputPath & "\" & curLtr & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    End If
                
                    bFirstLetter = False
                Else 'else we are just adding to a file already created.
                    'Set objTemplateDoc = objWordApp.Documents.Open(strOutputFileName)
                    'objTemplateDoc.Close
'                    objWordApp.Selection.InsertBreak (7) 'wdPageBreak

                    objWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage)
                    objWordApp.selection.InsertFile (strOutputFileName)
                    objWordApp.selection.EndKey Unit:=wdStory
                    'need to re-set these after every insert! it gets defaluted back once a new file is inserted.
                    objMasterDoc.spellingchecked = True
                    objMasterDoc.showspellingerrors = False
                    objMasterDoc.showgrammaticalerrors = False
                    
                    'If iCnt > 0 And (iCnt Mod 10) = 0 Then
                    If iCnt > fmrStatus.ProgVal And (iCnt Mod 10) = 0 Then                          'i think this re-does the tabs
                        objMasterDoc.Repaginate                                                     'thieu 1/16/08
                    End If
                
                   
                    If CountCurLetter > 100 Then 'if the Letter type is comprised of over 100 letters split it
                    'this below does not work when we hit 100 pages it keeps going. not refreshing correctly.
                    'If objMasterDoc.BuiltInDocumentProperties("Number of Pages") > 100 Then 'wdPropertyPages       'thieu 1/16/08 property # 14
                        bnewwordfile = True                                                             'thieu 1/16/08
                    End If                                                                          'thieu 1/16/08
                End If ' End of check if were are in the bFirstLetter
            'else we are dealing with a PDF
            ElseIf InStr(1, strOutputFileName, ".PDF", vbTextCompare) > 0 Then
            'we have a PDF output so we call Join PDF's.
                'if targetPDF is empty that means this is the first time we are encountering a pdf so set the target pdf here.
                If TargetPDF = "" Or CountCurLetter > 100 Then 'if the Letter type is comprised of over 100 letters split it
                 TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem)) _
                                           & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
                End If
                    
                    PassFail = JoinPDFs(TargetPDF, strOutputFileName)
                    If Err Or PassFail = False Then
                        MsgBox ("error joining pdfs")
                        Exit Function
                    End If
            Else
                MsgBox ("File needs to be in word or pdf format")
            End If
        'keep track of how many times we have looped.
        iCnt = iCnt + 1
    End If
        
        ' display progress
        sMsg = "View Record " & iCnt & " / " & TotalRecs
        fmrStatus.ProgVal = ProgVal + iCnt
        fmrStatus.StatusMessage sMsg

        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If

        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.  if  so rollback and promt with error message.
        If fmrStatus.EvalStatus(2) = True Then
                sMsg = "Viewing Generated Letters Canceled!"
                fmrStatus.StatusMessage sMsg
                DoEvents
                strErrMsg = sMsg
                GoTo Error_Encountered
        End If
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Next varItem

 If ShowPreview Then
    If Usedword Then
        objMasterDoc.spellingchecked = True
        
        With objMasterDoc
            For i = 1 To .Sections.Count
                .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            Next i
        End With
        Sleep 1000
        objMasterDoc.SaveAs strPreviewFileName
        objMasterDoc.Close
        Set objMasterDoc = Nothing
    End If
      
      'right here we need to execute the print command with view only as a parameter.
      
        Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
    Else
        MsgBox ("Items selected must be printed already to view.")
    End If
    ViewLetters = True
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    ViewLetters = False

Cleanup:

    If Not objMasterDoc Is Nothing Then '07/01/2013
        objMasterDoc.Close wdDoNotSaveChanges
    End If

    
    Set objMasterDoc = Nothing
    Set objTemplateDoc = Nothing
    'Do not close the Connection passed in, taking care of that in the calling sub.
    Set cmdGetLetter = Nothing
'    Set Person = Nothing
    objWordApp.Quit wdDoNotSaveChanges
     Set objWordApp = Nothing
    ' make word visible for user to view
    'objWordApp.Visible = True
    'objWordApp.Activate

End Function


Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mReturnDate = ReturnDate
End Sub


Private Function PreviewViewLetters(oCn As ADODB.Connection, TotalRecs As Long, fmrStatus As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'    Dim PreViewFileArray() As String
Dim cmd As ADODB.Command
Dim cmdGetLetter As ADODB.Command
Dim strSQLcmd As String

' Letter configuration variables
'Dim rsLetterConfig As dao.Recordset
Dim strODCFile As String
Dim strBasedPath As String
Dim colLetterTemplate As Collection
Dim objLetterInfo As clsLetterTemplate
    
    ' Word objects setup as variants b/c of late binding (due to the various versions we have scattered around the environment)
'    Dim objWordApp As Word.Application, _
'        objMasterDoc As Word.Document, _
'        objWordDoc As Word.Document, _
'        objWordMergedDoc As Word.Document, _
'        objWordField As Word.Field, _
'        objWordSection As Word.Section

Dim objWordApp As Object, _
    objMasterDoc As Object, _
    objWordDoc As Object, _
    objWordMergedDoc As Object, _
    objWordField As Object, _
    objWordSection As Object
    
    'Letter generation variables
Dim rsProvList As ADODB.RecordSet
Dim rsLetterTemplate As ADODB.RecordSet
Dim strInstanceID As String
Dim strProvNum As String
Dim strAuditor As String
Dim strLetterType As String
Dim dtLetterReqDt As Date
Dim strStatus As String
Dim strLocalTemplate As String
Dim strLocalPath As String
Dim oLetter As clsLetterInstance

Dim bMergeError As Boolean
Dim strOutputPath As String
Dim strOutputFileName As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iRtnCd As Integer

Dim varItem As Variant
Dim iAnswer As Integer
Dim bFirstLetter As Boolean
Dim iCnt As Integer
Dim i As Integer
Dim sMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim bObjectExists  As Boolean
Dim rsLetterConfig As ADODB.RecordSet
Dim oRs As ADODB.RecordSet
Dim sCombinedDoc As String
Dim sSuperName As String
    
    
    strProcName = ClassName & ".PreviewViewLetters"
    
    bFirstLetter = True
    
    strErrMsg = ""
    
    Set cmd = New ADODB.Command
    Set rsProvList = New ADODB.RecordSet
    Set rsLetterTemplate = New ADODB.RecordSet
    Set objWordApp = CreateObject("Word.Application")
    objWordApp.visible = False

    Set cmdGetLetter = New ADODB.Command
    Set objLetterInfo = New clsLetterTemplate
    
    
    'set local path
    'USER ENTRY NEEDED
    strLocalPath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    'End USER ENTRY NEEDED
    If Not FolderExist(strLocalPath) Then CreateFolders (strLocalPath)
    
    ' get list of templates
    Set colLetterTemplate = New Collection
    'TL add account ID logic

    ' NOTE: we need to fully qualify the table here because the connection we are using
    ' (with the transaction started) is from the _CODE database)
    strSQLcmd = "select LetterType, TemplateLoc from CMS_AUDITORS_CLAIMS.dbo.LETTER_Type where AccountID = " & gintAccountID
    cmd.ActiveConnection = oCn
    cmd.CommandText = strSQLcmd
    cmd.CommandTimeout = 0
    cmd.commandType = adCmdText
    Set rsLetterTemplate = cmd.Execute

    'rst is a recset of all templates and locations.
    'Then populates the collection colLetterTemplate with the info from these templates.
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            objLetterInfo.LetterType = Trim(![LetterType])
            objLetterInfo.TemplateLoc = Trim(![TemplateLoc])
            colLetterTemplate.Add objLetterInfo, Trim(![LetterType])
            strLocalTemplate = strLocalPath & "\" & GetFileName(objLetterInfo.TemplateLoc)
            'see if the template exists at the strLocalPath if it does we are deleting it to make room for the copy over.
            If DeleteFile(strLocalTemplate, False) = False Then
                LogMessage strProcName, "ERROR", "Could not delete the template - may be locked open?", strLocalTemplate
            End If
            Set objLetterInfo = New clsLetterTemplate
            .MoveNext
        End With
    Loop
    If rsLetterTemplate.State = adStateOpen Then rsLetterTemplate.Close
    Set rsLetterTemplate = Nothing
    
    ' Set the based path for saving merge doc
    Set rsLetterConfig = GetLetterConfigDetails()
    
    strBasedPath = rsLetterConfig("LetterOutputLocation").Value
    strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    'select our preview folder and delete it if it exists...
    If rsLetterConfig.State = adStateOpen Then rsLetterConfig.Close
    Set rsLetterConfig = Nothing
    
    
    strOutputPath = strBasedPath & "\PREVIEW"
    strAuditor = Replace(Identity.UserName, ".", "")
    strOutputPath = strOutputPath & "\" & strAuditor
    DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
    CreateFolders (strOutputPath & "\")
    
    bMergeError = False
      
    cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("LetterName", adChar, adParamInput, 255, "")
    cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adChar, adParamOutput, 255, "")
    
    Set objLetterInfo = New clsLetterTemplate

    
    ' setup progress screen that is passed to this function
    ' start processing letters
    iCnt = 0
    
'    Dim MyRecordset As dao.Recordset

    Set oRs = Me.lstQueue.RecordSet
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    
    For Each oLetter In cdctSelectedLetters.Letters
        If oLetter.LetterQueueStatus = "W" Or oLetter.LetterQueueStatus = "R" Then

            strInstanceID = oLetter.InstanceId
            strProvNum = oLetter.cnlyProvID
            strLetterType = oLetter.LetterType
            dtLetterReqDt = oLetter.LetterReqDt
            strAuditor = oLetter.Auditor
            strStatus = oLetter.LetterQueueStatus
            
            iCnt = iCnt + 1

            If strLetterType <> objLetterInfo.LetterType Then
                'We are dealing with a new report in the queue!
                Set objLetterInfo = colLetterTemplate(strLetterType)
                 
                bObjectExists = False
                'ACCESS REPORT LOGIC
                'Check to see if we are talking about a report or Template
                If InStr(1, objLetterInfo.TemplateLoc, ".doc", vbBinaryCompare) = 0 Then 'if we are dealing with an access rpt
                Stop
                    'make sure the report exits in access.
'                    For i = 0 To db.Containers("Reports").Documents.Count - 1
'                        If db.Containers("Reports").Documents(i).Name = objLetterInfo.TemplateLoc Then
'                            bObjectExists = True
'                        End If
'                    Next i
'
'                    If bObjectExists = False Then
'                        strErrMsg = "Missing letter template in access." & vbCrLf & "Template Report name = " & objLetterInfo.TemplateLoc & ""
'                        GoTo Block_Err
'                    End If
                Else ' we have a Word template we are working from (WORD MAIL MERGE)
                    'WORD DOC AREA
                    'check if the template physically exists
                    If FileExists(objLetterInfo.TemplateLoc) = False Then
                        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
                        GoTo Block_Err
                    End If
                
                    
                   ' make a local copy so it would not impact other users
                    strLocalTemplate = strLocalPath & "\" & GetFileName(objLetterInfo.TemplateLoc)
                    
                    If FileExists(strLocalTemplate) = False Then
                        If CopyFile(objLetterInfo.TemplateLoc, strLocalTemplate, False, strErrMsg) = False Then
                            LogMessage strProcName, "ERROR", "Error copying template file to temp folder." & strErrMsg, objLetterInfo.TemplateLoc & " to " & strLocalTemplate
                        End If
                    End If
                 
                    'May put this back in someday.  i believe this is being done automatically via the MailMerge.
                    ' open template or objWordDoc and set margins to the objmasterdoc.
                    'When the mail merge runs it keeps the template's Margins....
                    Set objWordDoc = objWordApp.Documents.Add(strLocalTemplate, , False) 'tried didn't effect change
                    
                    
                    'add a connolly-internal watermark to the preview letters
                    If Not (ADDWATERMARK(objWordApp, objWordDoc, strErrMsg)) Then
                        LogMessage strProcName, "ERROR", "There was an error while adding the watermark: " & strErrMsg, strErrMsg
                        GoTo Block_Err
                    End If

                    
                End If  'End of split for access report or word mail merge
            End If 'end of check to see if the template report or word doc exists
            
            ' NOW CHECK TO SEE IF THE TEMPLATE IS NOT A WORD DOCUMENT
            If InStr(1, objLetterInfo.TemplateLoc, ".doc", vbTextCompare) <> 0 Then 'If we are dealing with a Word Mail Merge
            
                    ' Create temp table to address issues with Word VBA/sp multiple selection criterion.
                    AdoExeTxt "usp_LETTER_Get_Info_load '" & strInstanceID & "',''", "v_CODE_Database"
                                
                    ' Set data source for mail merge.  Data will be from new Temp Table
                    objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, _
                        SqlStatement:="exec usp_LETTER_Get_Info '" & strInstanceID & "'"
                    
                    ' Perform mail merge.
                    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
                    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
                    objWordDoc.MailMerge.Execute Pause:=False
                    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
                        objWordApp.visible = True
                        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
                        bMergeError = True
                        objWordApp.ActiveDocument.Activate
                        GoTo Block_Exit
                    End If
                    ''------------------- here is where we convert to pdf instead of word ----------------''
                    ' Save the output doc
                    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
                    
                    ' 20130219 KD: Add the sec pages field in the footer
                    Call AddSecPagesCode(objWordMergedDoc, oLetter)

                    'Added to rename reprints...
                    strOutputFileName = strOutputPath & "\" & strLetterType & "-Preview-" & strInstanceID & ".doc"
                    objWordMergedDoc.spellingchecked = True
                    

                        ' 20130219 KD: Make sure that the section pages start at 1
                    objWordMergedDoc.Repaginate
                    
                    With objWordMergedDoc
                        For i = 1 To .Sections.Count
                            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                            objWordMergedDoc.Repaginate
                            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                            objWordMergedDoc.Repaginate
                        Next i
                    End With
                    
'                   This section was to Unlink the SecPages code but it's not needed - or, perhaps it IS needed
                    For Each objWordSection In objWordApp.ActiveDocument.Sections
                        For i = 1 To objWordSection.Footers.Count
                            For Each objWordField In objWordSection.Footers.Item(i).Range.Fields
                                Debug.Print objWordField.Code
                                objWordField.Update
                                objWordField.Unlink
                            Next
                            
                        Next
                    Next

                    Sleep 1000
                    objWordMergedDoc.SaveAs strOutputFileName
                    objWordMergedDoc.Close

                    Set objWordMergedDoc = Nothing

            End If
                     
            strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
            If strErrMsg <> "" Then
                GoTo Block_Err
            End If
            
    
            ' display progress  '/ 2
            sMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax & vbCrLf & _
                        "Provider = " & strProvNum & vbCrLf & _
                        "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
            fmrStatus.ProgVal = iCnt
            fmrStatus.StatusMessage sMsg
    
            If fmrStatus.ProgMax = lngProgressCount Then
                msgIcon = vbInformation
            Else
                msgIcon = vbExclamation
            End If
            
            'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
            If fmrStatus.EvalStatus(2) = True Then
                    sMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
                    fmrStatus.StatusMessage sMsg
                    DoEvents
                    strErrMsg = sMsg
                    GoTo Block_Err
            End If
            
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            
            ' Clear TEMP LOAD Table
            AdoExeTxt "usp_LETTER_Get_Info_tmp_clear", "v_CODE_Database"
        End If 'end if to ensure the items are marked as W to print
    
    Next
    

    PreviewViewLetters = True



    ' HEre, ONLY prompt if the user is Data Services:
    sSuperName = Identity.UserSupervisorId()
    
    If UCase(sSuperName) <> "DATA CENTER" Then
        i = MsgBox("Would you like to combine the Preview Letters into one file?", vbYesNo)
        If i = vbYes Then
            If (Not CombineDocs(strOutputPath, True)) Then 'could pass stroutputpath, true to delete all but he combined here.  clients choice.
                GoTo Block_Err
            End If
        End If
    End If
    
    
    
    Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
          
    
    'should track all documents created and delete the ones when we error'd out or cancelled generation.
     'now empty out the array
'    On Error Resume Next
'    For i = 0 To UBound(PreViewFileArray)
'        If PreViewFileArray(i) <> "" Then
'        Kill PreViewFileArray(i)
'        'only deletes folder if it is empty
'            If Len(dir(Left(PreViewFileArray(i), InStrRev(PreViewFileArray(i), "\")) & "*.*")) = 0 Then
'                DeleteFolder (Left(PreViewFileArray(i), InStrRev(PreViewFileArray(i), "\") - 1))
'            End If
'        End If
'
'    Next i
    

      
Block_Exit:

    If Not objMasterDoc Is Nothing Then '07/01/2013
        objMasterDoc.Close wdDoNotSaveChanges
    End If


    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    ' Release references.

    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    Set objMasterDoc = Nothing
    
    Set cmd = Nothing
    Set cmdGetLetter = Nothing
    Set rsLetterConfig = Nothing
    Set rsProvList = Nothing
'    Set Person = Nothing
    objWordApp.Quit wdDoNotSaveChanges
    Set objWordApp = Nothing
    Exit Function
Block_Err:
    If strErrMsg <> "" Then
        LogMessage TypeName(Me) & ".PreviewViewLetters", "ERROR", "An error occurred: " & strErrMsg
        
        MsgBox strErrMsg, vbCritical
    Else
        ReportError Err, strProcName
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical

    End If
    PreviewViewLetters = False
    GoTo Block_Exit
End Function


'
''' This function will look for a bookmark named 'SecPages' and will replace that with
''' the Sec Pages field
'Private Function AddSecPagesCode(objWordDoc As Object) As Boolean
''Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim objWordField As FormField
'
''Dim objWordSection As Object
'Dim objWordSection As Object
'Dim iFoot As Integer
'Dim oFooter As Object
'Dim oField As Object
'Dim oRange As Object
'Dim sBookmarkName As String
'Dim saryBkmarks(1) As String
'Dim iBkmark As Integer
'
'
'    strProcName = TypeName(Me) & ".AddSecPagesCode"
'
'    saryBkmarks(0) = "SecPages"
'    saryBkmarks(1) = "SecPages2"
'
'    For iBkmark = 0 To UBound(saryBkmarks)
'        sBookmarkName = saryBkmarks(iBkmark)
'
'        If IsBookMark(objWordDoc, sBookmarkName) = False Then
'            ' nothing else to do
'            GoTo NextBkmark
'        End If
'
'        For Each objWordSection In objWordDoc.Sections
'            For iFoot = 1 To objWordSection.Footers.Count
'                    ' shape range should take care of the text box..
'                    '' Note: may want to do the Headers here too
'                Set oFooter = objWordSection.Footers.Item(iFoot)
'                    ' Looks like 2007 + use a different means..
'                If oFooter.Shapes.Count > 0 Then
'                    Set oRange = objWordDoc.bookmarks(sBookmarkName).Range
'
'                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'                ElseIf oFooter.Range.ShapeRange.Count > 0 Then
'
'                    Set oRange = objWordDoc.bookmarks(sBookmarkName).Range
'                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'                End If
'            Next
'        Next
'NextBkmark:
'    Next
'
'    AddSecPagesCode = True  ' In this case true = no error.. :)
'
'Block_Exit:
'    Set oField = Nothing
'    Set oRange = Nothing
'    Set oFooter = Nothing
'    Set objWordSection = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function

'
'Private Function IsBookMark(objWordDoc As Object, sBookmarkName As String) As Boolean
''Private Function IsBookMark(objWordDoc As Word.Document, sBookMarkName As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oBkmk As Object
''Dim oBkmk As Word.Bookmark
'
'    strProcName = TypeName(Me) & ".IsBookMark"
'
'    For Each oBkmk In objWordDoc.bookmarks
'        If UCase(oBkmk.Name) = UCase(sBookmarkName) Then
'            IsBookMark = True
'            GoTo Block_Exit
'        End If
'    Next
'
'Block_Exit:
'    Set oBkmk = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function

Private Function ADDWATERMARK(objWordApp As Variant, objWordDoc As Variant, ByRef strErrMsg As String) As Boolean
'Private Function ADDWATERMARK(objWordApp As Word.Application, objWordDoc As Word.Document, ByRef strErrMsg As String) As Boolean
On Error GoTo Error_Encountered

     objWordDoc.Select
    'take our open word document. add a watermark since we are previewing.  This will deter the auditors from printing this.  (ALTHOUGH they could remove it manually)
    objWordApp.ActiveDocument.Sections(1).Range.Select
    objWordApp.ActiveWindow.ActivePane.View.seekview = 9 'wdSeekCurrentPageHeader
    'objWordApp.activewindow.activepane.View.seekview = 9 'wdseekcurrentPageHeader
    
    objWordApp.selection.HeaderFooter.Shapes.AddTextEffect(vbNull, _
        "CONNOLLY - INTERNAL", "Times New Roman", 1, False, False, 0, 0).Select
    objWordApp.selection.ShapeRange.Name = "PowerPlusWaterMarkObject1"
    objWordApp.selection.ShapeRange.TextEffect.NormalizedHeight = False
    objWordApp.selection.ShapeRange.Line.visible = False
    objWordApp.selection.ShapeRange.Fill.visible = True
    objWordApp.selection.ShapeRange.Fill.Solid
    objWordApp.selection.ShapeRange.Fill.ForeColor.RGB = RGB(153, 153, 153)
    objWordApp.selection.ShapeRange.Fill.Transparency = 0
    objWordApp.selection.ShapeRange.Rotation = 315
    objWordApp.selection.ShapeRange.LockAspectRatio = True
    objWordApp.selection.ShapeRange.top = -999995  'wdShapeCenter
    objWordApp.selection.ShapeRange.left = -999995 'wdShapeCenter
    objWordApp.selection.ShapeRange.Height = objWordApp.inchestopoints(2) '1.69)
    objWordApp.selection.ShapeRange.Width = objWordApp.inchestopoints(6.77) '500 '487.45 'InchesToPoints(6.77)
    objWordApp.selection.ShapeRange.WrapFormat.AllowOverlap = True
    objWordApp.selection.ShapeRange.WrapFormat.Side = 3 'wdWrapNone
    objWordApp.selection.ShapeRange.WrapFormat.Type = 3
    objWordApp.selection.ShapeRange.RelativeHorizontalPosition = 0 ' wdRelativeVerticalPositionMargin
    objWordApp.selection.ShapeRange.RelativeVerticalPosition = 0 '  wdRelativeVerticalPositionMargin
    objWordApp.ActiveWindow.ActivePane.View.seekview = 0 'wdSeekMainDocument
    
    ADDWATERMARK = True
    Exit Function
                    
Error_Encountered:
    strErrMsg = Nz(strErrMsg, "") & "Error adding watermark.  " & Err.Source & " " & Err.Number & " " & Err.Description
    ADDWATERMARK = False
    
End Function

'Function created by TGH on 10-3-08
'   This function will go through all Word docs (excluding temp docs) and combine them together into one file.  It takes will by default create a new file named CombinedDocs unless it is passed a different name.
'           This does not include a preview window for progress but it runs pretty quick...
'  Parameters: StrPath -> The folder containing all of the word docs we would like to combine.
'                      StrPreviewFileName -> the name of the combined word docs file.  by defaule it will be called combinedDocs.
'         *Note: This function does not Run Recursivily to sub folders.  It would be a simple fix if we were looking to do this.  just run a alternative loop to recurse to the subfolder and call combine docs...

Private Function CombineDocs(strPath As String, Optional KillNonCombined As Boolean, Optional strPreviewFileName As String) As Boolean
On Error GoTo Err:
Dim First As Boolean: First = True
   
   'Dim strPreviewFileName As String
   'strpath = "\\ccaintranet.com\DFS-FLD-01\Audits\AmeriHealth\LETTER_REPOSITORY\LETTERS\PREVIEW\tomhartey"
   
   'the default would be false but doing this in case it is ever passed as null
    KillNonCombined = Nz(KillNonCombined, False)
 
   
    If Nz(strPreviewFileName, "") = "" Then
        strPreviewFileName = strPath & "\CombinedDocs.DOC"
    End If
   
Dim fso As Scripting.FileSystemObject
Dim CurrentFolder As Scripting.Folder
Dim Files As Scripting.Files, file As Scripting.file
   Set fso = New Scripting.FileSystemObject
   
    
    'Late bind the Word Object library 11
'    Dim objWordApp As Word.Application, _
'        objMasterDoc As Word.Document, _
'        objWordDoc As Word.Document, _
'        objWordField As Word.Field, _
'        objWordSection As Word.Section

Dim objWordApp As Object, _
    objMasterDoc As Object, _
    objWordDoc As Object, _
    objWordField As Object, _
    objWordSection As Object

Dim i As Integer

    
    Set objWordApp = CreateObject("Word.Application")
    'Set objWordApp = New Word.Application
    
        ' Create an instance of objWord, and make it invisible.
    objWordApp.visible = False
    
    Set objMasterDoc = objWordApp.Documents.Add()   'tgh 3/18/08
                    objMasterDoc.spellingchecked = True 'this is needed to shut down popup for too many spelling errors
                    Sleep 1000
                    objMasterDoc.SaveAs strPreviewFileName
        
    'get the appropriate folder
    Set CurrentFolder = fso.GetFolder(strPath)
    Set Files = CurrentFolder.Files
    For Each file In Files  'for each file in the folders
        If InStr(1, file.Name, ".doc") > 0 And file.Path <> strPreviewFileName And InStr(1, file.Name, "~") = 0 Then
        
            If First Then
                First = False
                Set objWordDoc = objWordApp.Documents.Open(file.Path)
        
                'set the margins for our newly created word doc to be the same as the first word doc in the folder.
                objMasterDoc.PageSetup.LeftMargin = objWordDoc.PageSetup.LeftMargin
                objMasterDoc.PageSetup.RightMargin = objWordDoc.PageSetup.RightMargin
                objMasterDoc.PageSetup.TopMargin = objWordDoc.PageSetup.TopMargin
                objMasterDoc.PageSetup.BottomMargin = objWordDoc.PageSetup.BottomMargin
                objMasterDoc.PageSetup.HeaderDistance = objWordDoc.PageSetup.HeaderDistance
                objMasterDoc.PageSetup.FooterDistance = objWordDoc.PageSetup.FooterDistance
                objWordDoc.Select   'make sure we close the right one and keep adding to the other...
        
                objWordDoc.Close
            Else
        '        objWordApp.Selection.InsertBreak    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
                objWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
                ' I need to make the fields static here I think..
        
            End If
        
            objWordApp.ActiveDocument.spellingchecked = True
        
        
        
            objWordApp.selection.InsertFile (file.Path)
            If KillNonCombined Then 'added if they only want the combined doc left...
                file.Delete True
                    '                Kill file.path
            End If
Dim intF As Integer

            '' 20130411 KD I'm not sure if we need to do this here or not but I don't think it hurts! (yes, need to do it!)
            With objWordApp.ActiveDocument
                For i = 1 To .Sections.Count
                    .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                    .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                    
                    Set objWordSection = .Sections(i)
'                    Stop
                    For intF = 1 To objWordSection.Footers.Count
                        For Each objWordField In objWordSection.Footers.Item(intF).Range.Fields
                            Debug.Print objWordField.Code
                            objWordField.Update
                            objWordField.Unlink
                        Next
                        
                    Next
                Next i
            End With
        
            objWordApp.ActiveDocument.Repaginate                                                     'thieu 1/16/08
            Sleep 1000
            objWordApp.ActiveDocument.SaveAs strPreviewFileName
        
        End If
    Next

    ' 20130219 KD: This was here before but I wanted to comment on it: It's to make sure that the sections (each document) starts at page 1
    With objWordApp.ActiveDocument
        For i = 1 To .Sections.Count
            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        Next i
    End With

    objWordApp.ActiveDocument.Repaginate                                                     'thieu 1/16/08

    Sleep 1000
    objWordApp.ActiveDocument.SaveAs strPreviewFileName

    'Destruct our footprint.
     Set objWordDoc = Nothing
    On Error Resume Next
       objWordApp.Quit (0)
     Set objWordApp = Nothing
     Set fso = Nothing
   
   CombineDocs = True
   Exit Function

Err:
    LogMessage TypeName(Me) & ".CombineDocs", "ERROR", Err.Description
    MsgBox Err.Description

   'Destruct our footprint.
   Set objWordDoc = Nothing
   On Error Resume Next
       objWordApp.Quit (0)
   
   Set objWordApp = Nothing
   Set fso = Nothing

    CombineDocs = False
End Function


Private Function GetLetterConfigDetails() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
    
    strProcName = ClassName & ".GetLetterConfigDetails"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_GetConfig"
        .Parameters.Refresh
        .Parameters("@pAccountId") = IIf(gintAccountID = 0, 1, gintAccountID)
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not get the letter configurations for some reason. DB Connectivity?"
            GoTo Block_Exit
        End If
    End With
    
Block_Exit:
    Set GetLetterConfigDetails = oRs
        ' don't close the RS because that'll close the returned RS
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function GenerateLetters(fmrStatus As Form_ScrStatus, Optional bAtLeastOneErrored As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim dctLetterTemplate As Scripting.Dictionary
Dim oLetterToGenerate As clsLetterInstance
Dim objLetterTemplate As clsLetterTemplate
Dim iCnt As Integer
Dim bTemplateFound As Boolean
Dim strInstanceID As String
Dim strProvNum As String
Dim strLetterType As String
Dim strOutputFileName As String
Dim strAuditor As String
Dim strErrMsg As String
Dim strProgressMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim rsLetterConfig As ADODB.RecordSet
Dim strODCFile As String
Dim strOutputLocation As String
Dim iPageCount As Integer

    
    strProcName = ClassName & ".GenerateLetters"
    
    Set rsLetterConfig = GetLetterConfigDetails()
    
    If rsLetterConfig.recordCount = 0 Then
        strErrMsg = "ERROR: Letter configuration parameters is missing"
        GoTo Block_Err
    ElseIf rsLetterConfig.recordCount > 1 Then
        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
        GoTo Block_Err
    Else
        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    End If
    
    ' setup progress screen that is passed to this function
    
    '' 20130821 KD: Not really sure what this is all about..
                DoEvents
                DoEvents
                DoEvents
                DoEvents

    ' in order to get here we need to have come through the Generate button
    ' which means we should have our form scoped cdctSelectedLetters
    
    If cdctSelectedLetters.Count < 1 Then
        LogMessage strProcName, "ERROR", "Got to this sub without any selected records for some reason!?!?!"
        GoTo Block_Exit
    End If
    
    Set dctLetterTemplate = New Scripting.Dictionary
    For Each oLetterToGenerate In cdctSelectedLetters.Letters
    
        ' display progress
        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
                    "Provider = " & strProvNum & vbCrLf & _
                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
        fmrStatus.ProgVal = iCnt
        fmrStatus.StatusMessage strProgressMsg

        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If
        
    
        strInstanceID = oLetterToGenerate.InstanceId
        strProvNum = oLetterToGenerate.ProvNum
        strLetterType = oLetterToGenerate.LetterType
        
        bTemplateFound = False
        
        If dctLetterTemplate.Exists(strLetterType) = False Then
            If CopyTemplates(dctLetterTemplate, strLetterType) = False Then
                strErrMsg = "There was a problem copying the letter templates to the users temp directory. Cannot proceed!"
                GoTo Block_Err
            End If
        End If
    
        Set objLetterTemplate = dctLetterTemplate.Item(strLetterType)
        
        iCnt = iCnt + 1
    
    
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If fmrStatus.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
            fmrStatus.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            GoTo Block_Exit
        End If

           
        If PrintLetterInstance(oLetterToGenerate, objLetterTemplate.TemplateLoc, strOutputFileName, strOutputLocation, strProvNum, _
                                        strODCFile, strLetterType, iPageCount) = False Then
            LogMessage strProcName, "ERROR", "Printing the letter instance failed for InstanceId: " & CStr(oLetterToGenerate.InstanceId), strErrMsg
            bAtLeastOneErrored = True
            GoTo NextLetter
        End If
        
        If cdctSelectedLetters.UpdateLetter(oLetterToGenerate) = False Then
            Stop
        End If


NextLetter:
        
            '' 20130821 KD: Not really sure what this is all about..
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If fmrStatus.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
            fmrStatus.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            GoTo Block_Exit
        End If

        ' display progress
        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
                    "Provider = " & strProvNum & vbCrLf & _
                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
        fmrStatus.ProgVal = iCnt
        fmrStatus.StatusMessage strProgressMsg
    Next
    

    
    ' Notify user we are done.
    cboViewType.SetFocus
    cboViewType = cboViewType.ItemData(1) 'JS change 20130305
    Call cboViewType_AfterUpdate
    
    Me.txtFromDate = Format(Now, "mm/dd/yyyy")
    Me.txtThroughDate = Format(Now, "mm/dd/yyyy")
    
    cmdRefresh_Click 'can't run with open trans
    
    GenerateLetters = True
    
Block_Exit:
    If GenerateLetters = False Then bAtLeastOneErrored = True

    Set rsLetterConfig = Nothing
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    End If
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    End If
    
    GoTo Block_Exit
End Function




Private Function PrintLetterInstance(oLetterInst As clsLetterInstance, pstrTemplateName As String, _
            pstrOutputFileName As String, pstrOutputBasePath As String, pstrProvNum As String, _
            pstrODCFile As String, pstrLetterType As String, _
            Optional iPageCount As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
    
Dim oCmd As ADODB.Command
Dim strSQLcmd As String
Dim bMergeError As Boolean
Dim strOutputPath As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iRtnCd As Integer
Dim varItem As Variant
Dim iAnswer As Integer
Dim iCnt As Integer
Dim i As Integer

Dim objLetterInfo As clsLetterTemplate

'    ' Word objects setup as variants b/c of late binding
'Dim objWordApp, _
'    objWordDoc, _
'    objWordMergedDoc
    
Dim objWordApp As Word.Application, _
    objWordDoc As Word.Document, _
    objWordMergedDoc As Word.Document
      
      Debug.Print "Outfile path: " & pstrOutputBasePath
      
    strProcName = ClassName & ".PrintLetterInstance"
    
    Set oAdo = New clsADO
    oAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    Set objLetterInfo = New clsLetterTemplate
    strErrMsg = ""
    

    Set objWordApp = CreateObject("Word.Application")
'    Set objWordApp = New Word.Application
    objWordApp.visible = False
    
    ' check if template exists
    strChkFile = Dir(pstrTemplateName)
    If strChkFile = "" Then
        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
        GoTo Block_Err
    End If

    ' open template
    Set objWordDoc = objWordApp.Documents.Add(pstrTemplateName, , False)
       
    ' load letter info
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Get_Info_load"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@InstanceID") = oLetterInst.InstanceId
    oCmd.Execute
    
    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
        GoTo Block_Err
    End If
    
    
    ' Set data source for mail merge.  Data will be from new Temp Table
    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
                        SqlStatement:="exec usp_LETTER_Get_Info '" & oLetterInst.InstanceId & "'"
                    
    
    ' Perform mail merge.
    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        strErrMsg = "Error encountered with mail merge."
        GoTo Block_Err
    End If
    
    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
'    Call AddSecPagesCode(objWordApp.ActiveDocument)
    
    
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
    Call CreateFolders(strOutputPath)

    If Not FolderExists(strOutputPath) Then
        strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        GoTo Block_Err
    End If
    
    If oLetterInst.LetterQueueStatus = "R" Then
    'If pstrInstanceStatus = "R" Then
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
    Else
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & oLetterInst.InstanceId & ".doc"
    End If
    
    objWordMergedDoc.spellingchecked = True
    objWordMergedDoc.Repaginate
    
 
    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
    Call AddSecPagesCode(objWordApp.ActiveDocument, oLetterInst)
        
    If UnlinkWordFields(objWordApp, objWordMergedDoc, oLetterInst.LetterType) = False Then
        LogMessage strProcName, "ERROR", "There was an error unlinking the fields. Check that the fields are correct!", pstrOutputFileName, True
    End If
    
    objWordMergedDoc.SaveAs pstrOutputFileName
    SleepEvents 1
    
    With oLetterInst
        If .LetterBatchId = 0 Then
            .LetterBatchId = Me.MostRecentBatchId
        End If
        If objWordMergedDoc.BuiltInDocumentProperties(14) = 1 Then
            Stop
        End If
        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
        .LetterPath = pstrOutputFileName
        .SaveStaticDetails
    End With
    
    objWordMergedDoc.Close
    
    Set objWordMergedDoc = Nothing
    
    If Not FileExists(pstrOutputFileName) Then
        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
        
        GoTo Block_Err
    End If

    
    ' clear letter info
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Get_Info_tmp_clear"
    oCmd.Parameters.Refresh
    'oCmd.Parameters("@pInstanceID") = pstrInstanceID
    oCmd.Execute
                                
    ' start letter transaction
    oAdo.BeginTrans
    
    
    ' update LETTER status
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Update_Status"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@InstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@LetterName").Value = pstrOutputFileName
    oCmd.Parameters("@pNextStatus").Value = "G" ' for Generated, not yet printed..
    oCmd.Execute
            
    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
        oAdo.RollbackTrans
        GoTo Block_Err
    End If
                            
                            
    ' update claim status & move to next queue
    ' note, only does this where the letter status is W
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_AuditClaims_Update"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@pInstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@pInstanceStatus").Value = oLetterInst.LetterQueueStatus
    oCmd.Execute
            
    strErrMsg = Trim(Nz(oCmd.Parameters("@pErrMsg").Value, ""))
    If strErrMsg <> "" Then
        oAdo.RollbackTrans
        GoTo Block_Err
    End If
                                
                                
    ' commit letter transaction
    oAdo.CommitTrans
    
    PrintLetterInstance = True
    


Block_Exit:
    ' Release references.
    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo 0
    End If
    Set objWordApp = Nothing
    
    Set oCmd = Nothing
    Set oAdo = Nothing
    
    Exit Function
Block_Err:

    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
        LogMessage TypeName(Me) & "PrintLetterInstance-2010", "USAGE DETAIL", strErrMsg
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
        LogMessage TypeName(Me) & "PrintLetterInstance-2010", "USAGE DETAIL", Err.Description
    End If
    PrintLetterInstance = False
    
    Call DeleteFile(pstrOutputFileName, False)
    
    GoTo Block_Exit
End Function


' ORIG
'Private Function PrintLetterInstance(pstrInstanceID As String, pstrInstanceStatus As String, pstrTemplateName As String, _
'            pstrOutputFileName As String, pstrOutputBasePath As String, pstrProvNum As String, _
'            pstrODCFile As String, pstrLetterType As String, _
'            Optional iPageCount As Integer) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oADO As clsADO
'
'Dim oCmd As ADODB.Command
'Dim strSQLCmd As String
'Dim bMergeError As Boolean
'Dim strOutputPath As String
'Dim strChkFile As String
'Dim strErrMsg As String
'Dim iRtnCd As Integer
'Dim varItem As Variant
'Dim iAnswer As Integer
'Dim iCnt As Integer
'Dim i As Integer
'
'Dim objLetterInfo As clsLetterTemplate
'Dim oLetterInst As clsLetterInstance
''    ' Word objects setup as variants b/c of late binding
'Dim objWordApp, _
'    objWordDoc, _
'    objWordMergedDoc
'
''    Dim objWordApp As Word.Application, _
''        objWordDoc As Word.Document, _
''        objWordMergedDoc As Word.Document
'
'
'
'    strProcName = ClassName & ".PrintLetterInstance"
'
'    Set oADO = New clsADO
'    oADO.ConnectionString = GetConnectString("v_CODE_Database")
'
'
'    Set objLetterInfo = New clsLetterTemplate
'    Set oLetterInst = New clsLetterInstance
'
'
'    If oLetterInst.LoadFromID(pstrInstanceID) = False Then
'        Stop
'    End If
'    strErrMsg = ""
'
'
'    Set objWordApp = CreateObject("Word.Application")
''    Set objWordApp = New Word.Application
'    objWordApp.visible = False
'
'    ' check if template exists
'    strChkFile = Dir(pstrTemplateName)
'    If strChkFile = "" Then
'        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
'        GoTo Block_Err
'    End If
'
'    ' open template
'    Set objWordDoc = objWordApp.Documents.Add(pstrTemplateName, , False)
'
'    ' load letter info
'    Set oCmd = New ADODB.Command
'    oCmd.ActiveConnection = oADO.CurrentConnection
'    oCmd.commandType = adCmdStoredProc
'    oCmd.CommandText = "usp_LETTER_Get_Info_load"
'    oCmd.Parameters.Refresh
'    oCmd.Parameters("@InstanceID") = pstrInstanceID
'    oCmd.Execute
'
'    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
'    If strErrMsg <> "" Then
'        GoTo Block_Err
'    End If
'
'
'    ' Set data source for mail merge.  Data will be from new Temp Table
'    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
'                        SQLStatement:="exec usp_LETTER_Get_Info '" & pstrInstanceID & "'"
'
'
'    ' Perform mail merge.
'    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
'    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
'    objWordDoc.MailMerge.Execute Pause:=False
'    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
'        objWordApp.visible = True
'        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
'        bMergeError = True
'        objWordApp.ActiveDocument.Activate
'        strErrMsg = "Error encountered with mail merge."
'        GoTo Block_Err
'    End If
'
'    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
'    Call AddSecPagesCode(objWordApp.ActiveDocument)
'
'    ' Save the output doc
'    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
'    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
'    Call CreateFolder(strOutputPath)
'    If Not FolderExists(strOutputPath) Then
'        strErrMsg = "Provider folder for letter was not created for instance: " & pstrInstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        GoTo Block_Err
'    End If
'
'    If pstrInstanceStatus = "R" Then
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & pstrInstanceID & ".doc"
'    Else
'        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & pstrInstanceID & ".doc"
'    End If
'
'    objWordMergedDoc.spellingchecked = True
'    Sleep 1000
'    objWordMergedDoc.SaveAs pstrOutputFileName
'
'    With oLetterInst
'        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
'        .LetterPath = pstrOutputFileName
'        .SaveStaticDetails
'    End With
'
'    ' Here we need to save the details to our LETTER_Static_Details table:
'    oLetterInst.SaveStaticDetails
'
'    objWordMergedDoc.Close
'
'    Set objWordMergedDoc = Nothing
'
'    If Not FileExists(pstrOutputFileName) Then
'        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & pstrInstanceID & vbNewLine + vbNewLine & "Process will continue for the instances left."
'        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
'
'        GoTo Block_Err
'    End If
'
'
'    ' clear letter info
'    Set oCmd = New ADODB.Command
'    oCmd.ActiveConnection = oADO.CurrentConnection
'    oCmd.commandType = adCmdStoredProc
'    oCmd.CommandText = "usp_LETTER_Get_Info_tmp_clear"
'    oCmd.Parameters.Refresh
'    'oCmd.Parameters("@pInstanceID") = pstrInstanceID
'    oCmd.Execute
'
'
'    ' start letter transaction
'    oADO.BeginTrans
'
'
'    ' update LETTER status
'    Set oCmd = New ADODB.Command
'    oCmd.ActiveConnection = oADO.CurrentConnection
'    oCmd.commandType = adCmdStoredProc
'    oCmd.CommandText = "usp_LETTER_Update_Status"
'    oCmd.Parameters.Refresh
'    oCmd.Parameters("@InstanceID").Value = pstrInstanceID
'    oCmd.Parameters("@LetterName").Value = pstrOutputFileName
'    oCmd.Execute
'
'    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
'    If strErrMsg <> "" Then
'        oADO.RollbackTrans
'        GoTo Block_Err
'    End If
'
'
'    ' update claim status & move to next queue
'    Set oCmd = New ADODB.Command
'    oCmd.ActiveConnection = oADO.CurrentConnection
'    oCmd.commandType = adCmdStoredProc
'    oCmd.CommandText = "usp_LETTER_AuditClaims_Update"
'    oCmd.Parameters.Refresh
'    oCmd.Parameters("@pInstanceID").Value = pstrInstanceID
'    oCmd.Parameters("@pInstanceStatus").Value = pstrInstanceStatus
'    oCmd.Execute
'
'    strErrMsg = Trim(Nz(oCmd.Parameters("@pErrMsg").Value, ""))
'    If strErrMsg <> "" Then
'        oADO.RollbackTrans
'        GoTo Block_Err
'    End If
'
'
'    ' commit letter transaction
'    oADO.CommitTrans
'
'    PrintLetterInstance = True
'
'
'
'Block_Exit:
'    ' Release references.
'    If Not objWordDoc Is Nothing Then '07/01/2013
'        objWordDoc.Close wdDoNotSaveChanges
'    End If
'
'    If Not objWordMergedDoc Is Nothing Then '07/01/2013
'        objWordMergedDoc.Close wdDoNotSaveChanges
'    End If
'
'    Set objWordDoc = Nothing
'    Set objWordMergedDoc = Nothing
'
'    If Not objWordApp Is Nothing Then
'        On Error Resume Next
'        objWordApp.Quit wdDoNotSaveChanges
'        On Error GoTo 0
'    End If
'    Set objWordApp = Nothing
'
'    Set oCmd = Nothing
'    Set oADO = Nothing
'
'    Exit Function
'Block_Err:
'
'
'    If strErrMsg <> "" Then
'        MsgBox strErrMsg, vbCritical
'        LogMessage TypeName(Me) & "PrintLetterInstance-2010", "USAGE DETAIL", strErrMsg
'    Else
'        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
'        LogMessage TypeName(Me) & "PrintLetterInstance-2010", "USAGE DETAIL", Err.Description
'    End If
'    PrintLetterInstance = False
'
'    Call DeleteFile(pstrOutputFileName, False)
'
'    GoTo Block_Exit
'End Function


Private Function GetUserTempDirectory() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim strTempPath As String
Dim strErrMsg As String

    strProcName = ClassName & ".GetUserTempDirectory"


    strTempPath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTempDir\"
    If CreateFolders(strTempPath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTempPath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTempPath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTempPath
        GoTo Block_Err
    End If
    
    
Block_Exit:
    GetUserTempDirectory = strTempPath
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    Else
        LogMessage strProcName, "ERROR", strErrMsg
    End If
    GoTo Block_Exit
End Function

Private Function CopyTemplates(dctTemplatesDict As Scripting.Dictionary, strLetterType As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

Dim oAdo As clsADO
Dim rsLetterTemplate As ADODB.RecordSet
Dim objLetterInfo As clsLetterTemplate
Dim strTemplatePath As String
Dim strLocalTemplate As String
Dim strSQL As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iFolderChkLoop As Integer

    strProcName = ClassName & ".CopyTemplates"

            '' just in case:
    If dctTemplatesDict Is Nothing Then Set dctTemplatesDict = New Scripting.Dictionary

    strSQL = "SELECT LetterType, TemplateLoc FROM LETTER_Type WHERE AccountID = " & gintAccountID & " and LetterType = '" & strLetterType & "'"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_DATA_Database")
            ' get list of templates
        .SQLTextType = sqltext
        .sqlString = strSQL
        Set rsLetterTemplate = .ExecuteRS
    End With
    

    ' create template directory
    iFolderChkLoop = 0
    strTemplatePath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    If CreateFolders(strTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTemplatePath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTemplatePath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTemplatePath
        GoTo Block_Err
    End If
    

    ' copy templates to local directory. Skip if template already there
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            strLocalTemplate = strTemplatePath & "\" & GetFileName(!TemplateLoc)
            
            strChkFile = Dir(strLocalTemplate) & ""
            If strChkFile = "" Then
                strChkFile = Dir(!TemplateLoc) & ""
                If strChkFile <> "" Then
                    If CopyFile(rsLetterTemplate("TemplateLoc").Value, strLocalTemplate, False, strErrMsg) = False Then
                        
                    End If
                Else
                    strErrMsg = "Error: source template " & rsLetterTemplate("TemplateLoc").Value & " not found"
                    
                End If
            End If
                    
            Set objLetterInfo = New clsLetterTemplate
            objLetterInfo.LetterType = Trim(!LetterType)
            objLetterInfo.TemplateLoc = strLocalTemplate
            
            If dctTemplatesDict.Exists(rsLetterTemplate("LetterType").Value) = True Then
                dctTemplatesDict.Item(rsLetterTemplate("LetterType").Value) = objLetterInfo
            Else
                dctTemplatesDict.Add rsLetterTemplate("LetterType").Value, objLetterInfo
            End If

            .MoveNext
        End With
    Loop
    
    CopyTemplates = True
    
Block_Exit:
    If Not rsLetterTemplate Is Nothing Then
        If rsLetterTemplate.State = adStateOpen Then rsLetterTemplate.Close
        Set rsLetterTemplate = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg
    Else
        ReportError Err, strProcName
    End If
    
    GoTo Block_Exit
End Function

'  ORIG
'Private Function CopyTemplates(pcolLetterTemplate As Collection, pstrLetterType) As Boolean
'
'    Dim myADO As clsADO
'    Dim rsLetterTemplate As ADODB.Recordset
'    Dim objLetterInfo As clsLetterTemplate
''    Dim Person As New ClsIdentity
'
'    Dim strTemplatePath As String
'    Dim strLocalTemplate As String
'    Dim strSQLCmd As String
'    Dim strChkFile As String
'    Dim strErrMsg As String
'    Dim iFolderChkLoop As Integer
'
'    Set myADO = New clsADO
'    myADO.ConnectionString = GetConnectString("v_DATA_Database")
'
'
'    CopyTemplates = False
'
'    ' create template directory
'    iFolderChkLoop = 0
'    strTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\" & Identity.UserName & "\LETTERTEMPLATE"
'    Do Until FolderExist(strTemplatePath) Or iFolderChkLoop = 5
'        CreateFolder (strTemplatePath)
'        iFolderChkLoop = iFolderChkLoop + 1
'    Loop
'
'    If Not FolderExist(strTemplatePath) Then
'        strErrMsg = "ERROR: can not create folder " & strTemplatePath
'        GoTo Error_Encountered
'    End If
'
'
'    ' get list of templates
'    strSQLCmd = "select LetterType, TemplateLoc from LETTER_Type where AccountID = " & gintAccountID & " and LetterType = '" & pstrLetterType & "'"
'
'    myADO.SQLTextType = SQLTEXT
'    myADO.sqlString = strSQLCmd
'    Set rsLetterTemplate = myADO.OpenRecordSet
'
'    ' copy templates to local directory. Skip if template already there
'    Do While Not rsLetterTemplate.EOF
'        With rsLetterTemplate
'            strLocalTemplate = strTemplatePath & "\" & GetFileName(!TemplateLoc)
'
'            strChkFile = Dir(strLocalTemplate) & ""
'            If strChkFile = "" Then
'                strChkFile = Dir(!TemplateLoc) & ""
'                If strChkFile <> "" Then
'                    FileCopy !TemplateLoc, strLocalTemplate
'                Else
'                    strErrMsg = "Error: source template " & !TemplateLoc & " not found"
'                End If
'            End If
'
'            Set objLetterInfo = New clsLetterTemplate
'            objLetterInfo.LetterType = Trim(!LetterType)
'            objLetterInfo.TemplateLoc = strLocalTemplate
'            pcolLetterTemplate.Add objLetterInfo, Trim(![LetterType])
'            .MoveNext
'        End With
'    Loop
'
'    CopyTemplates = True
'    GoTo Clean_Up
'
'Error_Encountered:
'    If strErrMsg <> "" Then
'        MsgBox strErrMsg, vbCritical
'    Else
'        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
'    End If
'    CopyTemplates = False
'
'Clean_Up:
'    Set rsLetterTemplate = Nothing
'    Set myADO = Nothing
'
'End Function

Private Sub DeleteTemplates()
'    Dim Person As New ClsIdentity
    Dim strTemplatePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file

    ' delete template directory
    strTemplatePath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    
    'JS Change 20130305 no more delete the folder, now it will only delete the contents
    'DeleteFolder (strTemplatePath)
    
'    On Error Resume Next
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(strTemplatePath)
    For Each oFile In oFldr.Files
        Dim sFilePath As String
        sFilePath = oFile.Path
        Set oFile = Nothing
        If DeleteFile(sFilePath, False) = False Then
            LogMessage ClassName & ".DeleteTemplates", , "Tried to delete a file - not there, probably fine!", strTemplatePath
        End If
    Next
    
    
'    Kill strTemplatePath & "\*.*"
'    On Error GoTo 0
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

'    Set Person = Nothing
End Sub


Private Function GetSelectedItems(Optional bViewOnly As Boolean = False, Optional sDesiredStatus As String, _
            Optional sDisallowedStatus As String, Optional sErrMsg As String) As clsLetterInstanceDct
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtrInstance As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim varItem As Variant
Dim oRs As ADODB.RecordSet
Dim iSelectedCount As Integer
Dim sLetterInstance As String
Dim sQStatus As String
Dim sProvNum As String
Dim sLetterType As String
Dim dctPos As Scripting.Dictionary
Dim bOk As Boolean

    strProcName = ClassName & ".GetSelectedItems"
    
    Set oLetters = New clsLetterInstanceDct

    If Not (TypeOf Me.lstQueue.RecordSet Is ADODB.RecordSet) Then
        Set oLetters = GetSelectedItemsDAO(bViewOnly)
        GoTo Block_Exit
    End If

    Set oRs = Me.lstQueue.RecordSet
    
    
    Set dctPos = GetADOFieldOrdinalPosition(oRs)
    
    
    ' Look at the whole list for selected, and correct status for printing
    For Each varItem In Me.lstQueue.ItemsSelected
    
        sLetterInstance = UCase(Trim(Me.lstQueue.Column(dctPos.Item("INSTANCEID"), varItem)))
        
        sQStatus = UCase(Trim(Me.lstQueue.Column(dctPos.Item("STATUS"), varItem)))
'Stop    ' Kev: why did they user CnlyProvId and call it ProvNum when there's a ProvNum in the recordset too??
                ' better check that out
        sProvNum = UCase(Trim(Me.lstQueue.Column(dctPos.Item("CNLYPROVID"), varItem)))
        
        sLetterType = UCase(Trim(Me.lstQueue.Column(dctPos.Item("LETTERTYPE"), varItem)))
            ' Only count Re-Printable and W (whatever that stands for)
            ' unless we are viewing instead of generating the letters
        
        If sDesiredStatus <> "" And sQStatus = sDesiredStatus Then bOk = True
        If (sQStatus = "W" Or sQStatus = "R") Or bViewOnly = True Then bOk = True
        
        If sDisallowedStatus <> "" Then
            If InStr(1, sQStatus, sDisallowedStatus, vbTextCompare) > 0 Then
                bOk = False
                ' this is disallowed, stop, don't process anything
                sErrMsg = "There is at least 1 record selected with a status of: '" & sQStatus & "' which cannot be processed this way!"
                Set oLetters = New clsLetterInstanceDct
                GoTo Block_Exit
            End If
        End If
        
        
        If bOk = True Then
            iSelectedCount = iSelectedCount + 1
            Call oLetters.AddLetterInstance(sLetterInstance, sQStatus, sProvNum, sLetterType)
        End If
        bOk = False ' reset
    Next varItem
    
    Set GetSelectedItems = oLetters
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function GetSelectedItemsDAO(Optional bViewOnly As Boolean = False) As clsLetterInstanceDct
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtrInstance As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim varItem As Variant
Dim oRs As DAO.RecordSet
Dim iSelectedCount As Integer
Dim sLetterInstance As String
Dim sQStatus As String
Dim sProvNum As String
Dim sLetterType As String


    strProcName = ClassName & ".GetSelectedItemsDAO"
    
    Set oLetters = New clsLetterInstanceDct

    Set oRs = Me.lstQueue.RecordSet
    oRs.MoveLast
    oRs.MoveFirst
    Sleep 750
    ' Look at the whole list for selected, and correct status for printing
    For Each varItem In Me.lstQueue.ItemsSelected
        sLetterInstance = Trim(Me.lstQueue.Column(oRs.Fields("InstanceID").OrdinalPosition, varItem))
        sQStatus = UCase(Trim(Me.lstQueue.Column(oRs.Fields("Status").OrdinalPosition, varItem)))
'Stop    ' Kev: why did they user CnlyProvId and call it ProvNum when there's a ProvNum in the recordset too??
                ' better check that out
        sProvNum = Trim(Me.lstQueue.Column(oRs.Fields("cnlyProvID").OrdinalPosition, varItem))
        sLetterType = Trim(Me.lstQueue.Column(oRs.Fields("LetterType").OrdinalPosition, varItem))
        
            ' Only count Re-Printable and W (whatever that stands for)
            ' unless we are viewing instead of generating the letters
            
        If (sQStatus = "W" Or sQStatus = "R") Or bViewOnly = True Then
            iSelectedCount = iSelectedCount + 1

            Call oLetters.AddLetterInstance(sLetterInstance, sQStatus, sProvNum, sLetterType)
        End If
    Next varItem
    
    Set GetSelectedItemsDAO = oLetters
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
Private Function GenerateBatchId() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim lBatchId As Long

    strProcName = ClassName & ".GenerateBatchID"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Assign_Generation_BatchId"
        .Parameters.Refresh
        .Parameters("@pUserId") = Identity.UserName
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error generating a Letter Generation Batch ID", .Parameters("@pErrMsg").Value, True
            GoTo Block_Exit
        Else
            lBatchId = .Parameters("@pBatchId").Value
        End If
    End With
    
    Me.MostRecentBatchId = lBatchId
    
Block_Exit:
    Set oAdo = Nothing
    GenerateBatchId = lBatchId
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function UpdateBatchId(lBatchId As Long, bSuccess As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO


    strProcName = ClassName & ".UpdateBatchId"
    
    If lBatchId = 0 Then
        lBatchId = Me.MostRecentBatchId
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Update_Generation_BatchId"
        .Parameters.Refresh
        .Parameters("@pBatchId") = lBatchId
        .Parameters("@pUserId") = Identity.UserName
        .Parameters("@pSuccess") = IIf(bSuccess, 1, 0)
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error trying to update the batch status", .Parameters("@pErrMsg").Value, True
            GoTo Block_Exit
        End If
    End With
    
Block_Exit:
    Set oAdo = Nothing
    UpdateBatchId = lBatchId
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function UnlinkWordFields(oWordApp As Word.Application, oDoc As Word.Document, sLetterType As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim objWordField As Word.Field
Dim objWordSection As Word.Section
Dim i As Integer

    strProcName = ClassName & ".UnlinkWordFields"
    
    ' 20130219 KD: Make sure that the section pages start at 1
      
    oDoc.Repaginate
      
    With oDoc
        .Fields.Unlink
        
          For i = 1 To .Sections.Count
              .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
              .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
              .Repaginate
              .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
              .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
              .Repaginate
          Next i
      End With
      
      oDoc.Activate
      
        '' Hardcoded (shame) for barcodes: need to make this data driven at some point..
    If sLetterType <> "VADRA_QR" Then
            ' by the way, this breaks the ADR footer's Page X of Y (even though it doesn't break the
        For Each objWordSection In oWordApp.ActiveDocument.Sections
            For i = 1 To objWordSection.Footers.Count
                For Each objWordField In objWordSection.Footers.Item(i).Range.Fields
                    Debug.Print objWordField.Code
                    objWordField.Update
                    objWordField.Unlink
                Next
                
            Next
        Next
        
    Else
        '' this should be unlinking the bar codes
        For Each objWordSection In oWordApp.ActiveDocument.Sections
            For Each objWordField In objWordSection.Range.Fields
    '            objWordField.Update
                objWordField.Unlink
            Next
        Next
    End If
        
      UnlinkWordFields = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
