Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frmReAssignSelect As Form_frm_QUEUE_ReAssign_Select
Attribute frmReAssignSelect.VB_VarHelpID = -1
Private WithEvents frmMassClaimRelease As Form_frm_QUEUE_Mass_Claim_Release
Attribute frmMassClaimRelease.VB_VarHelpID = -1

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private ColReSize As clsAutoSizeColumns

Private mbMassClaimReleaseUnload As Boolean
Private mOperMode As Integer
Private mGroupName As String
Private mrsDisplayConfig As ADODB.RecordSet
Private mrsAgeCalConfig As ADODB.RecordSet
Private mstrUserName As String
Private mstrDetailView As String
Private mbRefreshSelection As Boolean
Private mCurrState As CurrState
Private miStackID As Integer
Private mColStates As Collection
Private miStatusPos As Integer
Private miAssignedToPos As Integer

Private mstrUserProfile As String
Private miAppPermission As Integer

Private Type CurrState
    Hdr_ViewOrder As String
    Hdr_RowSelected As Integer
    Hdr_SQL As String
    Dtl_ViewOrder As String
    Dtl_RowSelected As Integer
    Dtl_SQL As String
    Dtl_Combo As String
End Type

Const CstrFrmAppID As String = "QueueMgt"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Let OperMode(ByVal vData As Integer)
    mOperMode = vData
End Property


Public Property Get OperMode() As Integer
    OperMode = mOperMode
End Property


Private Sub cboDetail_Change()
    Dim strViewOrder As String
    Dim i As Integer
    
    On Error GoTo Err_handler
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    cboDetail.SetFocus
    
    mrsDisplayConfig.MoveFirst
    mrsDisplayConfig.Find "ComboBoxOrder = " & cboDetail.Column(1)
    If mrsDisplayConfig.EOF <> True Then
        mrsDisplayConfig("CtrlColValue") = ""
    End If
    
    strViewOrder = mCurrState.Hdr_ViewOrder & "|" & cboDetail.Column(1)
    lstDetail.RowSource = SetQueueSQL(strViewOrder)
    
    myCode_ADO.sqlString = lstDetail.RowSource
    Set lstDetail.RecordSet = myCode_ADO.OpenRecordSet()
    lstDetail.ColumnCount = lstDetail.RecordSet.Fields.Count
    
    
    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstDetail
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstDetail.ListCount - 1 > 0 Then
        ColReSize.AutoSize
    End If
    Set ColReSize = Nothing
        
    'cmdReply.Enabled = False
    'cmdForward.Enabled = False
    
    If cboDetail.Column(0) = mstrDetailView Then
        cmdMassClaimUpdate.visible = True
        cmdViewDetail.visible = True
        cmdReAssign.visible = True
        cmdSelectAll.visible = True
        CmdClear.visible = True
        cmdRefreshDtl.visible = True
        'cmdForward.visible = True
        'cmdReply.visible = True
        cmdPrintChartReview.visible = True
        cmdPrintClaimDetail.visible = True
        miStatusPos = GetColumnPosition(lstDetail, "QueueStatus")
        miAssignedToPos = GetColumnPosition(lstDetail, "AssignedTo")
        'If miStatusPos < 0 Then cmdReply.visible = False
        If miAssignedToPos < 0 Then
            'cmdForward.visible = False
            cmdReAssign.visible = False
        End If
    Else
        cmdViewDetail.visible = False
        cmdReAssign.visible = False
        cmdSelectAll.visible = False
        CmdClear.visible = False
        cmdRefreshDtl.visible = False
        'cmdForward.visible = False
        'cmdReply.visible = False
        cmdPrintChartReview.visible = False
        cmdPrintClaimDetail.visible = False
        cmdMassClaimUpdate.visible = False
    End If
    
    mCurrState.Dtl_Combo = cboDetail
    mCurrState.Dtl_ViewOrder = cboDetail.Column(1)
    mCurrState.Dtl_SQL = lstDetail.RowSource
    
    
    ' permission checking & override
    If (miAppPermission And gcAllowReAssign) = False Then
        cmdReAssign.visible = False
    End If
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
    MsgBox Err.Number & " --- " & Err.Description
    Resume Exit_Sub
End Sub


Private Sub cboHdr_Change()

    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    ' reset data filters
    If mrsDisplayConfig.BOF = True And mrsDisplayConfig.EOF = True Then Exit Sub
    
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        mrsDisplayConfig("CtrlColValue") = ""
        mrsDisplayConfig.MoveNext
    Wend

    mCurrState.Hdr_ViewOrder = cboHdr.Column(1)
    mCurrState.Hdr_RowSelected = 0
    mCurrState.Hdr_SQL = SetQueueSQL(mCurrState.Hdr_ViewOrder)
    
    lstHdr.RowSource = mCurrState.Hdr_SQL
    myCode_ADO.sqlString = mCurrState.Hdr_SQL
    Set lstHdr.RecordSet = myCode_ADO.OpenRecordSet
    lstHdr.ColumnCount = lstHdr.RecordSet.Fields.Count
    
    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstHdr
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstHdr.ListCount - 1 > 0 Then
        ColReSize.AutoSize
    End If
    Set ColReSize = Nothing
    
    
    LoadDtlComboBox
    cboDetail = cboDetail.DefaultValue
    
    If lstHdr.ListCount > 1 Then
        lstHdr.Selected(1) = True
        Call lstHdr_Click
    End If
    
    Call cboDetail_Change
    
    Set myCode_ADO = Nothing
End Sub


Private Sub cmdClear_Click()
    Dim i As Integer

    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do

    For i = 1 To lstDetail.ListCount
        lstDetail.Selected(i) = False
    Next i

End Sub


Private Sub cmdExportToExcel_Click()
    Dim bExport As Boolean
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    If lstDetail.ListCount = 1 Then Exit Sub     'only row headers, nothing to do
    
    bExport = ExportDetails(Me.lstDetail.RecordSet, "QueueDetails.xls")
    
    'If bExport = False Then
    '     MsgBox "An error was encountered while attempting to export Detail data to Excel.", vbCritical
    'End If

End Sub


'Private Sub cmdForward_Click()
'    Dim varItem
'
'    If lstDetail.Recordset Is Nothing Then Exit Sub     'nothing to do
'
'    If lstDetail.ItemsSelected.Count < 1 Then
'        MsgBox "Please select items to be forwarded", vbInformation
'        Exit Sub
'    End If
'
'    For Each varItem In lstDetail.ItemsSelected
'        If UCase(lstDetail.Column(miAssignedToPos, varItem)) <> UCase(mstrUserName) Then
'            MsgBox "You can not forward items not assigned to you"
'            Exit Sub
'        End If
'
'        If UCase(lstDetail.Column(miStatusPos, varItem)) = "FORWARD" Then
'            MsgBox "You can not forward items FORWARDED to you.  Please reply instead"
'            Exit Sub
'        End If
'
'
'    Next
'
'    lstHdr.Selected(lstHdr.ListIndex + 1) = False
'    If lstDetail.ItemsSelected.Count = lstDetail.Recordset.RecordCount Then
'        mbRefreshSelection = False
'    End If
'
'    Set frmReAssignSelect = New Form_frm_QUEUE_ReAssign_Select
'    frmReAssignSelect.Action = "FORWARD"
'    ColObjectInstances.Add frmReAssignSelect, frmReAssignSelect.hWnd & ""
'    ShowFormAndWait frmReAssignSelect
'    Set frmReAssignSelect = Nothing
'End Sub


Private Sub cmdMassClaimUpdate_Click()
    If lstDetail.ItemsSelected.Count < 1 Then
        MsgBox "Please select the claim(s) you want to update first!", vbInformation
        Exit Sub
    End If

    mbMassClaimReleaseUnload = False
    
    Set frmMassClaimRelease = New Form_frm_QUEUE_Mass_Claim_Release
    Set frmMassClaimRelease.claimList = Me.lstDetail
    
    If mbMassClaimReleaseUnload = False Then
        ColObjectInstances.Add frmMassClaimRelease, frmMassClaimRelease.hwnd & ""
        ShowFormAndWait frmMassClaimRelease
    End If
    Set frmMassClaimRelease = Nothing
    
End Sub

Private Sub cmdPrintChartReview_Click()
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do

    If lstDetail.ItemsSelected.Count < 1 Then
        MsgBox "Please select the claim(s) you want to print first!", vbInformation
        Exit Sub
    End If

    Dim varItem
    Dim strCnlyClaimNum
    For Each varItem In lstDetail.ItemsSelected
        strCnlyClaimNum = lstDetail.Column(2, varItem)
        DoCmd.OpenReport "rpt_AUDITCLM_ChartReview", acViewNormal, , "cnlyClaimNum = '" & strCnlyClaimNum & "'"
        'DoCmd.OpenReport "rpt_AUDITCLM_ChartReview", acViewPreview, , "cnlyClaimNum = '" & strCnlyClaimNum & "'"
    Next
End Sub

Private Sub cmdPrintClaimDetail_Click()
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do

    If lstDetail.ItemsSelected.Count < 1 Then
        MsgBox "Please select the claim(s) you want to print first!", vbInformation
        Exit Sub
    End If

    Dim varItem
    Dim strCnlyClaimNum
    For Each varItem In lstDetail.ItemsSelected
        strCnlyClaimNum = lstDetail.Column(2, varItem)
        'DoCmd.OpenReport "rptClaimDetail", acViewPreview, "cnlyClaimNum = '" & strCnlyClaimNum & "'"
        LaunchClaimDetailReport (strCnlyClaimNum)
        DoCmd.OpenReport "rptClaimDetail", acViewNormal, "cnlyClaimNum = '" & strCnlyClaimNum & "'"
        DoCmd.Close acReport, "rptClaimDetail", acSaveNo
    Next
End Sub

Private Sub cmdReAssign_Click()
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    
    If lstDetail.ItemsSelected.Count < 1 Then
        MsgBox "Please select items to be re-assigned", vbInformation
        Exit Sub
    End If
    
    lstHdr.Selected(lstHdr.ListIndex + 1) = False
    If lstDetail.ItemsSelected.Count = lstDetail.RecordSet.recordCount Then
        mbRefreshSelection = False
    End If
    
    Set frmReAssignSelect = New Form_frm_QUEUE_ReAssign_Select
    frmReAssignSelect.Action = "REASSIGN"
    ColObjectInstances.Add frmReAssignSelect, frmReAssignSelect.hwnd & ""
    ShowFormAndWait frmReAssignSelect
    Set frmReAssignSelect = Nothing
End Sub


Private Sub cmdRefresh_Click()
    cboHdr_Change
End Sub

Private Sub cmdRefreshDtl_Click()
    'Refresh list
    Call cmdClear_Click
    Call RefreshData
End Sub

'Private Sub cmdReply_Click()
'    Dim varItem As Variant
'
'    If lstDetail.Recordset Is Nothing Then Exit Sub     'nothing to do
'
'    If lstDetail.ItemsSelected.Count < 1 Then
'        MsgBox "Please select items to reply", vbInformation
'        Exit Sub
'    End If
'
'    For Each varItem In lstDetail.ItemsSelected
'        If UCase(mstrUserName) <> UCase(lstDetail.Column(miAssignedToPos, varItem)) Then
'            MsgBox "You can only reply on items assigned to you"
'            Exit Sub
'        End If
'        If UCase(lstDetail.Column(miStatusPos, varItem) <> "FORWARD") Then
'            MsgBox "You can only reply to items forwarded to you."
'            Exit Sub
'        End If
'    Next
'
'    lstHdr.Selected(lstHdr.ListIndex + 1) = False
'    If lstDetail.ItemsSelected.Count = lstDetail.Recordset.RecordCount Then
'        mbRefreshSelection = False
'    End If
'
'    Set frmReAssignSelect = New Form_frm_QUEUE_ReAssign_Select
'    frmReAssignSelect.Action = "REPLY"
'    ColObjectInstances.Add frmReAssignSelect, frmReAssignSelect.hWnd & ""
'    ShowFormAndWait frmReAssignSelect
'    Set frmReAssignSelect = Nothing
'End Sub


Private Sub cmdRollBack_Click()
    If miStackID = 0 Then
        MsgBox "You're already at the top view"
    Else
        Call RollBackCurrentState
        Call RefreshData
    End If
End Sub


Private Sub cmdSelectAll_Click()
    Dim i As Integer
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to view

    For i = 1 To lstDetail.ListCount
        lstDetail.Selected(i) = True
    Next i
End Sub


Private Sub cmdViewDetail_Click()
    Dim iClaimIDPos As Integer
    Dim strCnlyClaimNum As String
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to view
    
    ' find claim ID position
    iClaimIDPos = GetColumnPosition(lstDetail, "CnlyClaimNum")
    
    If lstDetail.ListIndex >= 0 Then
        strCnlyClaimNum = lstDetail.Column(iClaimIDPos)
        DoCmd.OpenForm "frm_QUEUE_Detail_View", acNormal, , "CnlyClaimNum = '" & strCnlyClaimNum & "'"
    Else
        MsgBox "Please select a claim first", vbInformation
    End If
End Sub


Private Sub Form_Close()
    lstHdr.RowSource = ""
    lstDetail.RowSource = ""
End Sub


Private Sub Form_Load()
    Me.Caption = "Queue Management"
    
    Call Account_Check(Me)
    
    Dim rsPermission As ADODB.RecordSet
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    mstrUserName = Identity.UserName
    
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
   
    miStackID = 0
    mbRefreshSelection = True
    
    ' initialize current state
    mCurrState.Hdr_ViewOrder = ""
    mCurrState.Hdr_SQL = ""
    mCurrState.Hdr_RowSelected = 0
    mCurrState.Dtl_ViewOrder = ""
    mCurrState.Dtl_SQL = ""
    mCurrState.Dtl_RowSelected = 0
    mCurrState.Dtl_Combo = ""
    
    Set mColStates = New Collection
    
    ' check and assign operation mode (user vs manager)
    If IsSubForm(Me) Then
        mOperMode = Nz(Me.Parent.OperMode, OperationMode.Manager)
        mGroupName = "QUEUE_QA"
    Else
        mOperMode = OperationMode.Manager
        mGroupName = "QUEUE_QA"
    End If
    
    
    
    ' read display and age calculation configuration
    ' TL add account ID logic
    MyAdo.sqlString = "select * from QUEUE_Display_Config " & _
                      " where GroupName = '" & mGroupName & "'" & _
                      " and AccountID = " & gintAccountID & _
                      " and ComboBoxOrder <> 0" & _
                      " order by ComboBoxOrder"
    Set mrsDisplayConfig = MyAdo.OpenRecordSet()
    
    
    MyAdo.sqlString = "select * from QUEUE_Display_Config " & _
                      " where GroupName = '" & mGroupName & "'" & _
                      " and AccountID = " & gintAccountID & _
                      " and ComboBoxOrder = 0"
    Set mrsAgeCalConfig = MyAdo.OpenRecordSet()
    
    
    ' remove columns associated with manager display
    If mOperMode = OperationMode.User Then
        mrsDisplayConfig.Find "ColInd = 1"
        While mrsDisplayConfig.EOF <> True
            mrsDisplayConfig("ComboBoxName") = ""
            mrsDisplayConfig.Find "ColInd = 1", 1, adSearchForward
        Wend
        mrsDisplayConfig.UpdateBatch
    End If
    
    ' set detail view column
    If mrsDisplayConfig.BOF = True And mrsDisplayConfig.EOF = True Then
        MsgBox "Queue display configuration is not setup for this account.  Please check with your administrator", vbCritical
    Else
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find "ColInd = 2"
        If mrsDisplayConfig.EOF <> True Then
            mstrDetailView = mrsDisplayConfig("ComboBoxName")
        End If
    
        ' initialize header combo box
        cboHdr.RowSourceType = "Value List"
        cboHdr.BoundColumn = 1
        cboHdr.ColumnCount = 2
        cboHdr.ColumnHidden = False
        cboHdr.ColumnWidths = "1.5;0"
    
        ' initialize header list box
        lstHdr.ColumnHeads = True
        lstHdr.ColumnCount = 99
    
        ' initialize detail combo box
        cboDetail.RowSourceType = "Value List"
        cboDetail.BoundColumn = 1
        cboDetail.ColumnCount = 2
        cboDetail.ColumnHidden = False
        cboDetail.ColumnWidths = "1.5;0"
    
        ' initialize detail list box
        lstDetail.ColumnHeads = True
        lstDetail.ColumnCount = 99
    
        Call LoadHdrComboBox
        cboHdr = cboHdr.DefaultValue
        Call cboHdr_Change
    End If
    
    Set MyAdo = Nothing
    
End Sub


Private Sub frmMassClaimRelease_ClaimProcessed()
    Call RefreshData
End Sub


Private Sub frmMassClaimRelease_FormUnload()
    mbMassClaimReleaseUnload = True
End Sub

Private Sub lstDetail_Click()
    Dim strStatus As String
    Dim strAssignedTo As String
    
    'cmdReply.Enabled = False
    'cmdForward.Enabled = False
    cmdReAssign.Enabled = True
    
    If cboDetail.Column(0) = mstrDetailView Then
        strStatus = ""
        If miStatusPos >= 0 Then strStatus = UCase(lstDetail.Column(miStatusPos))
    
        strAssignedTo = ""
        If miAssignedToPos >= 0 Then strAssignedTo = UCase(lstDetail.Column(miAssignedToPos))
        
        If strStatus = "FORWARD" Then
            If UCase(mstrUserName) = strAssignedTo Then
                'cmdReply.Enabled = True
                'cmdForward.Enabled = False
                cmdReAssign.Enabled = False
            Else
                'cmdReply.Enabled = False
                'cmdForward.Enabled = False
                cmdReAssign.Enabled = False
            End If
        ElseIf strStatus = "OPEN" Then 'Or strStatus = "REPLY"
            If UCase(mstrUserName) = strAssignedTo Then
                cmdReAssign.Enabled = True
                'cmdForward.Enabled = True
                'cmdReply.Enabled = False
            Else
                cmdReAssign.Enabled = True
                'cmdForward.Enabled = False
                'cmdReply.Enabled = False
            End If
        End If
    End If
    
End Sub


Private Sub lstDetail_DblClick(Cancel As Integer)
    Dim strSQL As String
    Dim strICN As String
    Dim strDRG As String
    Dim strCnlyClaimNum As String
    Dim intMRReceived_Age As Integer
    Dim strAssignedTo As String
    Dim strHdrList As String
    Dim iResults As Integer
    
    If cboDetail.Column(0) = mstrDetailView Then
        If Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "QueueType")) = "PAC" Then
            NewProvider Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "cnlyProvID")), ""
        Else
            NewMain Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "cnlyClaimNum")), ""
            strHdrList = lstDetail.Column(8)
            strICN = lstDetail.Column(4)
            strDRG = lstDetail.Column(1)
            intMRReceived_Age = lstDetail.Column(0)
            strAssignedTo = lstDetail.Column(5)
            strCnlyClaimNum = lstDetail.Column(3)
            Set myCode_ADO = New clsADO
            myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
            strSQL = "insert into CMS_AUDITORS_Claims..QUEUE_Reviewed values('" & strICN & "', '" & strHdrList & "', " & intMRReceived_Age & ", '" & strDRG & "', '" & strAssignedTo & "', getdate(), '" & mstrUserName & "', '" & strCnlyClaimNum & "')"
            myCode_ADO.sqlString = strSQL
            myCode_ADO.SQLTextType = sqltext
            'Insert row into reviewed table so it doesn't show up again
            iResults = myCode_ADO.Execute()
            Set MyAdo = Nothing
            Set myCode_ADO = Nothing
            'Refresh list
            'Call RefreshData
        End If
    Else
        ' save header state
        mCurrState.Dtl_RowSelected = lstDetail.ListIndex + 1
        mCurrState.Dtl_SQL = lstDetail.RowSource
        mCurrState.Dtl_ViewOrder = cboDetail.Column(1)
        mCurrState.Dtl_Combo = cboDetail
        Call SaveCurrentState
    
        'set new header state
        mCurrState.Hdr_ViewOrder = mCurrState.Hdr_ViewOrder + "|" & cboDetail.Column(1, cboDetail.ListIndex)
        mCurrState.Hdr_RowSelected = lstDetail.ListIndex + 1
        mCurrState.Hdr_SQL = lstDetail.RowSource
        
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.sqlString = mCurrState.Hdr_SQL
        Set lstHdr.RecordSet = myCode_ADO.OpenRecordSet
        lstHdr.ColumnCount = lstHdr.RecordSet.Fields.Count
        lstHdr.Selected(mCurrState.Hdr_RowSelected) = True
        
        LoadDtlComboBox
        cboDetail = cboDetail.DefaultValue
        
        Call lstHdr_Click
    End If
    
    Set myCode_ADO = Nothing
End Sub


Private Sub lstHdr_Click()
    Dim ViewOrderList
    Dim iViewOrder As Integer
    Dim i As Integer
    Dim strListValue As String
    
    mCurrState.Hdr_RowSelected = lstHdr.ListIndex + 1
    
    ViewOrderList = Split(mCurrState.Hdr_ViewOrder, "|")
    
    For i = 0 To UBound(ViewOrderList)
        iViewOrder = val(ViewOrderList(i))
    
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find "ComboBoxOrder = " & iViewOrder
        
        If mrsDisplayConfig.EOF <> True Then
            strListValue = lstHdr.Column(mrsDisplayConfig("LstColOrder") - 1, lstHdr.ListIndex + 1) & ""
            If strListValue = "" Then strListValue = "NULL"
            mrsDisplayConfig("CtrlColValue") = strListValue
        End If
    Next i
    
    Call cboDetail_Change
End Sub


Private Function LoadHdrComboBox()
    Dim i, j As Integer
    Dim ViewOrder
    
    
    ' reset display config record
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        mrsDisplayConfig("ComboBoxName") = mrsDisplayConfig("ComboBoxName").OriginalValue
        mrsDisplayConfig.MoveNext
    Wend
    
    ViewOrder = Split(mCurrState.Hdr_ViewOrder, "|")
    For i = 0 To UBound(ViewOrder)
        j = val(ViewOrder(i))
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find ("ComboBoxOrder = " & j)
        If mrsDisplayConfig.EOF <> True Then
            mrsDisplayConfig("ComboBoxName") = ""
        End If
    Next i
    
    cboHdr.RowSource = ""
    cboHdr.DefaultValue = ""
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        If mrsDisplayConfig("ComboBoxName") <> "" And mrsDisplayConfig("ColInd") <> 2 Then
            If mrsDisplayConfig("ColInd") >= 0 Then
                cboHdr.RowSource = cboHdr.RowSource & mrsDisplayConfig("ComboBoxName") & ";" & Trim(CStr(mrsDisplayConfig("ComboBoxOrder"))) & ";"
                If cboHdr.DefaultValue = "" Then
                    cboHdr.DefaultValue = mrsDisplayConfig("ComboBoxName")
                End If
            End If
        End If
        mrsDisplayConfig.MoveNext
    Wend
End Function


Private Function LoadDtlComboBox()
    Dim i, j As Integer

    Dim ViewOrder
    
    
    ' reset display config record
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        mrsDisplayConfig("ComboBoxName") = mrsDisplayConfig("ComboBoxName").OriginalValue
        mrsDisplayConfig.MoveNext
    Wend
    
    ViewOrder = Split(mCurrState.Hdr_ViewOrder, "|")
    For i = 0 To UBound(ViewOrder)
        j = val(ViewOrder(i))
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find ("ComboBoxOrder = " & j)
        If mrsDisplayConfig.EOF = False Then
            mrsDisplayConfig("ComboBoxName") = ""
        End If
    Next i
    
    
    cboDetail.RowSource = ""
    cboDetail.DefaultValue = ""
    
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        If mrsDisplayConfig("ComboBoxName") <> "" Then
            If mrsDisplayConfig("ColInd") >= 0 Then
                cboDetail.RowSource = cboDetail.RowSource & mrsDisplayConfig("ComboBoxName") & ";" & Trim(CStr(mrsDisplayConfig("ComboBoxOrder"))) & ";"
                If cboDetail.DefaultValue = "" Then
                    cboDetail.DefaultValue = mrsDisplayConfig("ComboBoxName")
                End If
            End If
        End If
        mrsDisplayConfig.MoveNext
    Wend
End Function


Private Function RollBackCurrentState() As Boolean
    Dim temp
    Dim strState As String
    
    On Error GoTo Err_handler
    
    If miStackID > 0 Then
        miStackID = miStackID - 1
        strState = mColStates(Trim(str(miStackID)))
        temp = Split(strState, ";")
        mCurrState.Hdr_RowSelected = val(temp(0))
        mCurrState.Hdr_SQL = temp(1)
        mCurrState.Hdr_ViewOrder = temp(2)
        mCurrState.Dtl_RowSelected = val(temp(3))
        mCurrState.Dtl_SQL = temp(4)
        mCurrState.Dtl_ViewOrder = temp(5)
        mCurrState.Dtl_Combo = temp(6)
        mColStates.Remove (Trim(str(miStackID)))
    End If
   
    RollBackCurrentState = True
    Exit Function

Err_handler:
    RollBackCurrentState = False

End Function


Private Function SaveCurrentState() As Boolean
    Dim CurrState
    
    On Error GoTo Err_handler
    
    CurrState = mCurrState.Hdr_RowSelected & ";" & mCurrState.Hdr_SQL & ";" & mCurrState.Hdr_ViewOrder & ";" & _
                mCurrState.Dtl_RowSelected & ";" & mCurrState.Dtl_SQL & ";" & mCurrState.Dtl_ViewOrder & ";" & _
                mCurrState.Dtl_Combo
    mColStates.Add CurrState, Trim(str(miStackID))
    miStackID = miStackID + 1
    
    SaveCurrentState = True
    Exit Function

Err_handler:
    SaveCurrentState = False
    MsgBox Err.Description
End Function


Private Sub RefreshData()
    Dim strDetailVal As String
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    strDetailVal = mCurrState.Dtl_Combo
    lstHdr.RowSource = mCurrState.Hdr_SQL
    myCode_ADO.sqlString = lstHdr.RowSource
    Set lstHdr.RecordSet = myCode_ADO.OpenRecordSet
    lstHdr.ColumnCount = lstHdr.RecordSet.Fields.Count
    
    If mbRefreshSelection = True Then
        If lstHdr.ListCount > 1 Then
            If mCurrState.Hdr_RowSelected > lstHdr.ListCount Then mCurrState.Hdr_RowSelected = 1
            lstHdr.Selected(mCurrState.Hdr_RowSelected) = True
            Call lstHdr_Click
        End If
    Else
        lstHdr.Selected(lstHdr.ListIndex + 1) = False
    End If
    
    LoadDtlComboBox
    cboDetail = strDetailVal
    Call cboDetail_Change
   
    mbRefreshSelection = True
    
    Set myCode_ADO = Nothing

End Sub


Private Function SetQueueSQL(QueueLevel As String) As String
    Dim ViewOrder
    Dim SQL As SQLConstruct
    Dim strSQL As String
    Dim iViewOrder As Integer
    Dim bDetailView As Boolean
    Dim i, j As Integer
    Dim strValue As String
    Dim iLstColOrder As Integer
    
    SQL.Select = ""
    SQL.Where = ""
    SQL.GroupBy = ""
    SQL.OrderBy = ""
    SQL.From = "from v_QUEUE_QA"
    
    bDetailView = False
    iLstColOrder = 0
        
    ViewOrder = Split(QueueLevel, "|")
    For i = 0 To UBound(ViewOrder)
        iViewOrder = val(ViewOrder(i))
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find "ComboBoxOrder = " & iViewOrder
        If mrsDisplayConfig.EOF <> True Then
            If mrsDisplayConfig("ColInd") = 2 Then
                ' detail view
                bDetailView = True
                SQL.Select = "select "
                SQL.OrderBy = "order by "
                If mrsAgeCalConfig.BOF = True And mrsAgeCalConfig.EOF = True Then
                Else
                    mrsAgeCalConfig.MoveFirst
                    While mrsAgeCalConfig.EOF <> True
                        SQL.Select = SQL.Select & "(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1) as " & mrsAgeCalConfig("DispColName") & ", "
                        SQL.OrderBy = SQL.OrderBy & "(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1), "
                        mrsAgeCalConfig.MoveNext
                    Wend
                End If
                
                SQL.Select = SQL.Select & " * "
                If mrsAgeCalConfig.BOF = True And mrsAgeCalConfig.EOF = True Then
                    SQL.OrderBy = ""
                Else
                    SQL.OrderBy = left(Trim(SQL.OrderBy), Len(Trim(SQL.OrderBy)) - 1)
                End If
                SQL.GroupBy = ""
            Else
                ' top view
                If SQL.Select = "" Then
                    SQL.Select = "select " & mrsDisplayConfig("CtrlColumn") & " as " & mrsDisplayConfig("DispColName")
                    SQL.GroupBy = "group by " & mrsDisplayConfig("CtrlColumn")
                    SQL.OrderBy = "order by " & mrsDisplayConfig("CtrlColumn")
                    
                Else
                    SQL.Select = SQL.Select & ", " & mrsDisplayConfig("CtrlColumn") & " as " & mrsDisplayConfig("DispColName")
                    SQL.GroupBy = SQL.GroupBy & ", " & mrsDisplayConfig("CtrlColumn")
                    SQL.OrderBy = SQL.OrderBy & ", " & mrsDisplayConfig("CtrlColumn")
                End If
                iLstColOrder = iLstColOrder + 1
                mrsDisplayConfig("LstColOrder") = iLstColOrder
            End If
            
            
            ' add where clause
            If mrsDisplayConfig("CtrlColValue") & "" <> "" Then
                strValue = mrsDisplayConfig("CtrlColValue")
                If SQL.Where = "" Then
                    If mrsDisplayConfig("CtrlColType") = "C" Then
                        If strValue = "NULL" Then strValue = ""
                        SQL.Where = "where isnull(" & mrsDisplayConfig("CtrlColumn") & ",'') = '" & strValue & "'"
                    Else
                        SQL.Where = "where " & mrsDisplayConfig("CtrlColumn") & " = " & mrsDisplayConfig("CtrlColValue")
                    End If
                Else
                    If mrsDisplayConfig("CtrlColType") = "C" Then
                        If strValue = "NULL" Then strValue = ""
                        SQL.Where = SQL.Where & " and isnull(" & mrsDisplayConfig("CtrlColumn") & ",'') = '" & strValue & "'"
                    Else
                        SQL.Where = SQL.Where & " and " & mrsDisplayConfig("CtrlColumn") & " = " & strValue
                    End If
                End If
            End If
                
            ' TL add account ID filter
            If SQL.Where = "" Then
                SQL.Where = "where AccountID = " & gintAccountID
            Else
                SQL.Where = SQL.Where & " and AccountID = " & gintAccountID
            End If
                
            ' check for column desc
            If mrsDisplayConfig("ColInd") <> 2 Then
                mrsDisplayConfig.MoveNext
                If mrsDisplayConfig.EOF <> True Then
                    If mrsDisplayConfig("ColInd") < 0 Then
                        SQL.Select = SQL.Select & ", " & mrsDisplayConfig("CtrlColumn") & " as " & mrsDisplayConfig("DispColName")
                        SQL.GroupBy = SQL.GroupBy & ", " & mrsDisplayConfig("CtrlColumn")
                        SQL.OrderBy = SQL.OrderBy & ", " & mrsDisplayConfig("CtrlColumn")
                        
                        iLstColOrder = iLstColOrder + 1
                        mrsDisplayConfig("LstColOrder") = iLstColOrder
                    End If
                End If
            End If
        End If
    Next i
    
    
    ' add total count
    If Not bDetailView Then
        SQL.Select = SQL.Select & ", Count(1) as TotalCount "
        If mrsAgeCalConfig.BOF = True And mrsAgeCalConfig.EOF = True Then
        Else
            mrsAgeCalConfig.MoveFirst
            While mrsAgeCalConfig.EOF <> True
                SQL.Select = SQL.Select & ", max(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1) as " & mrsAgeCalConfig("DispColName")
                mrsAgeCalConfig.MoveNext
            Wend
        End If
    End If
    
    ' filter by user if operation mode = user
    If mOperMode = OperationMode.User Then
        mrsDisplayConfig.MoveFirst
        mrsDisplayConfig.Find "AuditorInd = 1"
        If mrsDisplayConfig.EOF <> True Then
            If SQL.Where = "" Then
                SQL.Where = "where " & mrsDisplayConfig("CtrlColumn") & " = '" & mstrUserName & "'"
            Else
                SQL.Where = SQL.Where & " and " & mrsDisplayConfig("CtrlColumn") & " = '" & mstrUserName & "'"
            End If
        End If
    End If
            
      
    strSQL = SQL.Select & Space(1) & SQL.From & Space(1) & SQL.Where & Space(1) & SQL.GroupBy & Space(1) & SQL.OrderBy
    SetQueueSQL = strSQL
End Function


Private Sub frmReAssignSelect_ReAssignQueue(AssignedTo As String, Comment As String, Action As String)
    Dim rsQueueHdr As ADODB.RecordSet
    Dim rsNote As ADODB.RecordSet
    Dim varItem As Variant
    Dim i, iMaxCol, iClaimIDPos, iStatusPos, iLastUpdatePos As Integer
    Dim iNoteID As Long
    Dim bResult As Boolean
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim strCnlyClaimNum As String
    Dim strStatus As String
    Dim strAssignedFrom As String
    Dim strAssignedTo As String
    Dim dLastUpdate As Date
    Dim dChkDate As Date
    
        
    On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        
    bResult = False
    iNoteID = 0
    strErrMsg = ""
    strErrSource = "frmReAssignSelect_ReAssignQueue"
    
    If Comment <> "" Then
        MyAdo.sqlString = "select * from NOTE_Detail where 1=2"
        Set rsNote = MyAdo.OpenRecordSet
        With rsNote
            .AddNew
            !AppID = Me.frmAppID
            !NoteType = "GENERAL"
            !NoteText = Comment
            !NoteUserID = Identity.UserName()
        End With
    End If
    
    ' find claim ID position
    iClaimIDPos = GetColumnPosition(lstDetail, "CnlyClaimNum")
    iStatusPos = GetColumnPosition(lstDetail, "QueueStatus")
    iLastUpdatePos = GetColumnPosition(lstDetail, "LastUpdate")
    
    ' reassign claims
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.BeginTrans
    
    For Each varItem In lstDetail.ItemsSelected
        strCnlyClaimNum = lstDetail.Column(iClaimIDPos, varItem)
        strStatus = UCase(lstDetail.Column(iStatusPos, varItem))
        dLastUpdate = lstDetail.Column(iLastUpdatePos, varItem)
        MyAdo.sqlString = "select * from QUEUE_Hdr where CnlyClaimNum = '" & strCnlyClaimNum & "'"
        Set rsQueueHdr = MyAdo.OpenRecordSet
    
        If rsQueueHdr.BOF = True And rsQueueHdr.EOF = True Then
            MsgBox "Item " & strCnlyClaimNum & " is no longer in queue!", vbInformation
        Else
            ' check if record has been update by someone else
            dChkDate = rsQueueHdr("LastUpdate")
            
            If Format(dLastUpdate, "ddmmyyhhmmss") <> Format(dChkDate, "ddmmyyhhmmss") Then
                strErrMsg = "Item " & strCnlyClaimNum & " has changed since opened.  Please refresh and try again!"
                GoTo Rollback
            End If
            
            ' add note
            If Comment <> "" Then
                iNoteID = GetAppKey("NOTE")
                rsNote.MoveFirst
                rsNote("NoteID") = iNoteID
    
                bResult = myCode_ADO.Update(rsNote, "usp_NOTE_Detail_Insert")
                If bResult = False Then
                    strErrMsg = "Error updating queue"
                    GoTo Err_handler
                End If
            End If
        
            Select Case UCase(Action)
                Case "FORWARD"
                    If strStatus <> "OPEN" And strStatus <> "REPLY" Then
                        strErrMsg = "Can not forward item: " & strCnlyClaimNum & "." & vbCrLf & "Status must be Open/Reply to forward!!"
                        GoTo Rollback
                    Else
                        rsQueueHdr("QueueStatus") = "Forward"
                        strAssignedFrom = rsQueueHdr("AssignedTo")
                        strAssignedTo = AssignedTo
                    End If
                Case "REPLY"
                    If strStatus <> "FORWARD" Then
                        strErrMsg = "Can not reply to item: " & strCnlyClaimNum & "." & vbCrLf & "Status must be Forward to reply!!"
                        GoTo Rollback
                    Else
                        rsQueueHdr("QueueStatus") = "Reply"
                        strAssignedFrom = rsQueueHdr("AssignedTo")
                        strAssignedTo = rsQueueHdr("AssignedFrom")
                    End If
                Case Else
                    If strStatus = "FORWARD" Then
                        strErrMsg = "Can not reassigned claim " & strCnlyClaimNum & " which was forwarded to you." & vbCrLf & "Status must be Open/Reply to reassign!!"
                        GoTo Rollback
                    End If
                    strAssignedFrom = rsQueueHdr("AssignedTo")
                    strAssignedTo = AssignedTo
            End Select
        
            ' update Queue header
            rsQueueHdr("LastUpdate") = Now()
            rsQueueHdr("UpdateUser") = Identity.UserName()
            rsQueueHdr("AssignedDate") = Date
            rsQueueHdr("AssignedFrom") = strAssignedFrom
            rsQueueHdr("AssignedTo") = strAssignedTo
            If iNoteID > 0 Then
                rsQueueHdr("NoteID") = iNoteID
            Else
                rsQueueHdr("NoteID") = Null
            End If
            bResult = myCode_ADO.Update(rsQueueHdr, "usp_QUEUE_Hdr_Update")
            If bResult = False Then
                strErrMsg = "Error updating queue"
                GoTo Err_handler
            End If
        End If
    Next
    
    myCode_ADO.CommitTrans
    
    Call RefreshData
    
Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Rollback:
    MsgBox strErrMsg, vbInformation
    myCode_ADO.RollbackTrans
    GoTo Exit_Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox Err.Number & " -- " & strErrSource & vbCrLf & vbCrLf & strErrMsg
    myCode_ADO.RollbackTrans
    Resume Exit_Sub
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub

Private Function ExportDetails(rst As ADODB.RecordSet, strFilePath As String) As Boolean

    Dim dlg As clsDialogs
    Dim cie As clsImportExport

    Set cie = New clsImportExport
    Set dlg = New clsDialogs

    
    If rst Is Nothing Then Exit Function     ' nothing to save
    
    With dlg
    
        strFilePath = .SavePath(Identity.CurrentFolder, xlsf, strFilePath)
        strFilePath = .CleanFileName(strFilePath, CleanPath)
     
        If strFilePath <> "" Then
        
            If .FileExists(strFilePath) = True Then
            
                If MsgBox("Overwrite existing file?", vbYesNo) = vbYes Then
                    .DeleteFile strFilePath
                Else
                    GoTo exitHere
                End If
            
            End If
            
            If rst.recordCount > 65535 Then
                MsgBox "Warning: Your recordset contains more than 65535 rows, the maximum number of rows allowed in Excel.  " & _
                Trim(str(rst.recordCount - 65535)) & " rows will not be displayed.", vbCritical
            End If
                        
        Else
            GoTo exitHere
        End If
     
        With cie
            .ExportExcelRecordset rst, strFilePath, True
        End With
         
    End With
    
    ExportDetails = True

exitHere:
    Set cie = Nothing
    Set dlg = Nothing
    Exit Function
    
HandleError:
    ExportDetails = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function
