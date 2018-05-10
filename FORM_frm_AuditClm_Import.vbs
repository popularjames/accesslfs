Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private ColReSize As clsAutoSizeColumns

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
Private msAdvancedFilter As String

Private mstrUserProfile As String
Private miAppPermission As Integer

Private Const strViewSource As String = "v_AuditClm_Import_Claims"

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents frmFilter As Form_frm_GENERAL_Filter
Attribute frmFilter.VB_VarHelpID = -1
Private Type CurrState
    Hdr_ViewOrder As String
    Hdr_RowSelected As Integer
    Hdr_SQL As String
    Dtl_ViewOrder As String
    Dtl_RowSelected As Integer
    Dtl_SQL As String
    Dtl_Combo As String
End Type

Const CstrFrmAppID As String = "AuditClm"

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
    
    
    'cboDetail.SetFocus
    
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
        
    
    If cboDetail.Column(0) = mstrDetailView Then
        cmdSelectAll.visible = True
        CmdClear.visible = True
    Else
        cmdSelectAll.visible = False
        CmdClear.visible = False
    End If
    
    mCurrState.Dtl_Combo = cboDetail
    mCurrState.Dtl_ViewOrder = cboDetail.Column(1)
    mCurrState.Dtl_SQL = lstDetail.RowSource
    
    
    
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


Private Sub chkFilter_Click()
    If Me.chkFilter.Value <> 0 And msAdvancedFilter = "" Then
        
            Set frmFilter = New Form_frm_GENERAL_Filter
            
            With frmFilter
                .CalledBy = Me.Name
                .FieldsTable = strViewSource
                .Setup
                .visible = True
            End With
    End If
    
    If Me.chkFilter.Value = 0 And msAdvancedFilter <> "" Then
            msAdvancedFilter = ""
            Call cboHdr_Change
    End If
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    
    For i = 1 To lstDetail.ListCount
        lstDetail.Selected(i) = False
    Next i
    
End Sub


Private Sub cmdClearMarked_Click()

  Dim strSQL As String
  Set MyAdo = New clsADO
  Dim rst As New ADODB.RecordSet
  Dim lngRecordsAffected As Long
  
  Set MyAdo = New clsADO
  MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
  MyAdo.sqlString = "DELETE aa FROM " & DLookup("DataBaseName", "v_Data_Database") & ".[dbo].[AuditClm_Import_Claims] aa"
  
  MyAdo.SQLTextType = sqltext
  lngRecordsAffected = MyAdo.Execute
  RefreshMarkedCount
End Sub

Private Sub cmdExportToExcel_Click()
    Dim bExport As Boolean
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    If lstDetail.ListCount = 1 Then Exit Sub     'only row headers, nothing to do
    
    bExport = ExportDetails(Me.lstDetail.RecordSet, "CLaim_Import.xls")
    
    'If bExport = False Then
    '     MsgBox "An error was encountered while attempting to export Detail data to Excel.", vbCritical
    'End If

End Sub
Private Sub cmdImport_Click()

  Dim rst As New ADODB.RecordSet
  Dim intCount As Long
  Dim intUserCount As Long
  Dim strSQLFromWhere As String
  Dim strSQL As String
  Dim lngRecordsAffected As Long
  Dim strSQLFrom As String
  Dim strDebug As String
  
  On Error GoTo ItHappened
  
  Me.txtDebug = ""
  
  strDebug = vbCrLf & "***SQL***" & vbCrLf & mCurrState.Dtl_SQL & vbCrLf
  Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
  
  'Get FROM WHERE based on user selections
  strSQLFromWhere = Mid(mCurrState.Dtl_SQL, InStr(1, mCurrState.Dtl_SQL, "FROM"), InStr(1, mCurrState.Dtl_SQL, "GROUP BY") - InStr(1, mCurrState.Dtl_SQL, "FROM"))
     
     
  strDebug = vbCrLf & "***SQL FROM WHERE***" & vbCrLf & strSQLFromWhere & vbCrLf
  Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
     
  'How many does the user want to import?
  intUserCount = Nz(Me.txtClaimCount, 0)

  'Connect to the database
  Set MyAdo = New clsADO
  MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
  
  'SELECT THE CLAIMS BASED ON THE TOP GRIDS SELECTION
  'GET THE COUNT
       
   
   MyAdo.sqlString = "SELECT SUm(1) as CT " & strSQLFromWhere
   
   strDebug = vbCrLf & "***GET COUNTS***" & vbCrLf & MyAdo.sqlString & vbCrLf
   Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
    
    Set rst = MyAdo.OpenRecordSet()
    If Not rst.EOF Then
        'Get the number of claims based on the user selections
        intCount = Nz(rst!CT, "")
        'if the count of the records is greater than what the user entered, we are only going to import the number they specified
        If (intCount > intUserCount And intUserCount > 0) Then
            strDebug = vbCrLf & CStr(intUserCount) & " Claims will be imported" & vbCrLf
            Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
            intCount = intUserCount
        Else
        'Otherwise, we just bring in everything
            intUserCount = intCount
            strDebug = vbCrLf & CStr(intCount) & " Claims will be imported" & vbCrLf
            Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
        End If
    Else
        'Recordset was empty, in this case, we'll do nothing
            strDebug = vbCrLf & "No Claims will be imported" & vbCrLf
            Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
    End If
    

    If MsgBox("You are about to import " & CStr(intCount) & " claims into the system. Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        
        
        'Build the SQL to insert the data into the table
        strSQL = " INSERT INTO " & DLookup("DataBaseName", "v_Data_Database") & ".[dbo].[AuditClm_Import_Claims] ([cnlyCLaimNum],[LineNum] ,[ConceptID] ,[Processed]) "
        strSQL = strSQL & " SELECT top " & str(intCount) & " cnlyCLaimNum, LineNum, ConceptID, 0 as Processed " & strSQLFromWhere
        strSQL = strSQL & " AND cnlyClaimNum NOT IN ( select cnlyClaimNum from " & DLookup("DataBaseName", "v_Data_Database") & ".[dbo].[AuditClm_Import_Claims] ) "
          
        If Nz(Me.cboFieldList, "") <> "" Then
            strSQL = strSQL & " ORDER BY " & Me.cboFieldList.Value & " " & IIf(Me.frmCriteria = 1, "ASC", "DESC") & "  "
        Else
            strSQL = strSQL & " "
        End If
    
        
        strDebug = vbCrLf & "***ADDING RECORDS***" & vbCrLf & strSQL & vbCrLf
        Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
        
        
        
        MyAdo.BeginTrans
        MyAdo.sqlString = strSQL
        MyAdo.SQLTextType = sqltext
        
        
        lngRecordsAffected = MyAdo.Execute
        If lngRecordsAffected <= 0 Then
            MsgBox CStr(lngRecordsAffected) & " Claims Imported", vbOKOnly
            strDebug = vbCrLf & "No Claims Imported" & vbCrLf
            Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
            MyAdo.RollbackTrans
        Else
            MsgBox CStr(lngRecordsAffected) & " Claims Imported", vbOKOnly
            strDebug = vbCrLf & CStr(lngRecordsAffected) & " Claims Imported" & vbCrLf
            Me.txtDebug = Nz(Me.txtDebug, "") & strDebug
            MyAdo.CommitTrans
        End If
    
    End If
   
RefreshMarkedCount


Exit Sub
ItHappened:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Set MyAdo = Nothing
End Sub
Private Sub RefreshMarkedCount()

  Dim strSQL As String
  Set MyAdo = New clsADO
  Dim rst As New ADODB.RecordSet
  
  Set MyAdo = New clsADO
  MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
  MyAdo.sqlString = "SELECT SUm(1) as CT FROM " & DLookup("DataBaseName", "v_Data_Database") & ".[dbo].[AuditClm_Import_Claims]"
  
  Set rst = MyAdo.OpenRecordSet()

  If Not rst.EOF Then
        Me.txtRecordsMarked = Nz(rst!CT, "")
  Else
        Me.txtRecordsMarked = 0
  End If
  
End Sub


Private Sub cmdRefresh_Click()
    cboHdr_Change
End Sub


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




Private Sub Form_Close()
    lstHdr.RowSource = ""
    lstDetail.RowSource = ""
End Sub


Private Sub Form_Load()
    Me.Caption = "Claim Import"
    
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
        Select Case UCase(Nz(Me.Parent.Parameter, "USER"))
            Case "USER", "MANAGER"
                mGroupName = "CLAIM_IMPORT"
            Case "USER_EXCEPTION", "MANAGER_EXCEPTION"
                mGroupName = "CLAIM_IMPORT"
            Case Else
                mGroupName = "CLAIM_IMPORT"

        End Select
            
    Else
        mOperMode = OperationMode.Manager
        mGroupName = "CLAIM_IMPORT"

    End If
    
    
    
    ' read display and age calculation configuration
    ' TL add account ID logic
    MyAdo.sqlString = "select * from GENERAL_DISPLAY_CONFIG " & _
                      " where GroupName = '" & mGroupName & "'" & _
                      " and AccountID = " & gintAccountID & _
                      " and ComboBoxOrder <> 0" & _
                      " order by ComboBoxOrder"
    Set mrsDisplayConfig = MyAdo.OpenRecordSet()
    
    
    MyAdo.sqlString = "select * from GENERAL_DISPLAY_CONFIG " & _
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
    Me.txtClaimCount = 0
    Me.cboFieldList.RowSource = strViewSource
    RefreshMarkedCount
End Sub


Private Sub frmfilter_QueryFormRefresh()
    Call cboHdr_Change
End Sub

Private Sub frmFilter_UpdateSql()
    msAdvancedFilter = Replace(frmFilter.SQL.WherePrimary, "#", "'")
End Sub
Private Sub lstDetail_DblClick(Cancel As Integer)

    
    If cboDetail.Column(0) = mstrDetailView Then
        'If Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "QueueType")) = "PAC" Then
        '    NewProvider Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "cnlyProvID")), ""
        'Else
        '    NewMain Me.lstDetail.Column(GetColumnPosition(Me.lstDetail, "cnlyClaimNum")), ""
        'End If
        MsgBox "Not Supported", vbInformation
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

    Me.txtDebug = vbCrLf & "***SQL***" & vbCrLf & mCurrState.Dtl_SQL & vbCrLf
   
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
    
    'Dynbamic
    SQL.From = "from " & strViewSource
    
    
    
    
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
            'DR add advanced filter
            
            If Me.chkFilter.Value <> 0 And Trim(msAdvancedFilter) <> "" Then
                
                If SQL.Where = "" Then
                    SQL.Where = "where AccountID = " & gintAccountID & " AND ( " & msAdvancedFilter & " ) "
                Else
                    SQL.Where = SQL.Where & " and AccountID = " & gintAccountID & " AND ( " & msAdvancedFilter & " ) "
                End If
            Else
                If SQL.Where = "" Then
                    SQL.Where = "where AccountID = " & gintAccountID
                Else
                    SQL.Where = SQL.Where & " and AccountID = " & gintAccountID
                End If
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


Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg, vbCritical, ErrSource
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg, vbCritical, ErrSource
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
