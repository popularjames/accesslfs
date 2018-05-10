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

' new stuff
Private mrsDisplayConfig As ADODB.RecordSet
Private mrsAgeCalConfig As ADODB.RecordSet
Private mstrUserName As String
Private mstrDetailView As String
Private mbRefreshSelection As Boolean
Private mCurrState As CurrState
Private miStackID As Integer
Private mColStates As Collection

Private Type CurrState
    Hdr_ViewOrder As String
    Hdr_RowSelected As Integer
    Hdr_SQL As String
    Dtl_ViewOrder As String
    Dtl_RowSelected As Integer
    Dtl_SQL As String
    Dtl_Combo As String
    
End Type

Public Property Let OperMode(ByVal vData As Integer)
    mOperMode = vData
End Property

Public Property Get OperMode() As Integer
    OperMode = mOperMode
End Property


Private Sub cboDetail_Change()
    
    Dim strViewOrder As String
    Dim i As Integer
    
    mrsDisplayConfig.MoveFirst
    mrsDisplayConfig.Find "ComboBoxOrder = " & cboDetail.Column(1)
    If mrsDisplayConfig.EOF <> True Then
        mrsDisplayConfig("CtrlColValue") = ""
    End If
    
    strViewOrder = mCurrState.Hdr_ViewOrder & "|" & cboDetail.Column(1)
    lstDetail.RowSource = SetViewSQL(strViewOrder)
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = lstDetail.RowSource
    Set lstDetail.RecordSet = myCode_ADO.OpenRecordSet()
    lstDetail.ColumnCount = lstDetail.RecordSet.Fields.Count
    
   
    mCurrState.Dtl_Combo = cboDetail
    mCurrState.Dtl_ViewOrder = cboDetail.Column(1)
    mCurrState.Dtl_SQL = lstDetail.RowSource
    
    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstDetail
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstDetail.ListCount - 1 > 0 Then
        ColReSize.AutoSize
    End If
    
    If Me.cboDetail.Value = "Detail" Then
        Me.cmdMarkLine.Enabled = True
    Else
        Me.cmdMarkLine.Enabled = False
    End If
    
    Set myCode_ADO = Nothing
    
End Sub

Private Sub cboHdr_Change()
    ' reset data filters
    If mrsDisplayConfig.BOF = True And mrsDisplayConfig.EOF = True Then Exit Sub
    
    mrsDisplayConfig.MoveFirst
    While mrsDisplayConfig.EOF <> True
        mrsDisplayConfig("CtrlColValue") = ""
        mrsDisplayConfig.MoveNext
    Wend

    mCurrState.Hdr_ViewOrder = cboHdr.Column(1)
    mCurrState.Hdr_RowSelected = 0
    mCurrState.Hdr_SQL = SetViewSQL(mCurrState.Hdr_ViewOrder)
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = mCurrState.Hdr_SQL
    lstHdr.RowSource = mCurrState.Hdr_SQL
    Set lstHdr.RecordSet = myCode_ADO.OpenRecordSet
    lstHdr.ColumnCount = lstHdr.RecordSet.Fields.Count
    
    
    LoadDtlComboBox
    cboDetail = cboDetail.DefaultValue
    
    If lstHdr.ListCount > 1 Then
        lstHdr.Selected(1) = True
        Call lstHdr_Click
    End If
    
    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstHdr
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstHdr.ListCount - 1 > 0 Then
        ColReSize.AutoSize
    End If
    
    Call cboDetail_Change
    
    Set myCode_ADO = Nothing
End Sub


Private Sub cmdExportToExcel_Click()
    Dim bExport As Boolean
    
    bExport = ExportDetails(Me.lstDetail.RecordSet, "PoolDetails.xls")
    
'    If bExport = False Then
'        MsgBox "An error was encountered while attempting to export Detail data to Excel.", vbCritical
'    End If

End Sub

Private Sub cmdMarkLine_Click()

  Dim sprocCheckClaim As clsAdoSproc
  Dim sprocCheckPool As clsAdoSproc
  Dim sprocMarkClaim As clsAdoSproc
  Dim Lst As listBox
  Dim lstStatus As listBox
  
  Dim lClaimNum As Long
  Dim lConceptCd As Long
  Dim lLinePosition As Long
  Dim lLineNum As Long
  Dim lInPool As Long
  Dim sClaimNum As String
  Dim varItem As Variant
  Dim sClmLevel As String
  
    Set sprocCheckClaim = New clsAdoSproc
    Set sprocCheckPool = New clsAdoSproc
    Set sprocMarkClaim = New clsAdoSproc
    
    Set Lst = Me.lstDetail
    Set lstStatus = Me.Parent.lstStatus
    
    If mGroupName = "" Then
        MsgBox "setup error"
        GoTo exitHere
    ElseIf mGroupName = "Pool_Dtl" Then
        sClmLevel = "L"
    Else
        sClmLevel = "H"
    End If

    If Lst.ItemsSelected.Count = 0 Then
        MsgBox "Select a Claim to add", vbOKOnly, "No Claim Selected"
        GoTo exitHere
    End If
    
    DoCmd.Hourglass True
    
    
    lClaimNum = GetColumnPosition(Lst, "CnlyClaimNum")
    
'Rob Swander 07/07/2010: Changed ConceptID to ConceptCd
'    lConceptCd = GetColumnPosition(Lst, "ConceptID")
    lConceptCd = GetColumnPosition(Lst, "ConceptCd")
    
'* Setup Sprocs

     sprocCheckPool.RefTable = "v_CODE_Database"
     sprocCheckPool.CommandText = "usp_POOL_CheckClaimMark"
     sprocCheckPool.Setup
    
     sprocCheckClaim.RefTable = "v_CODE_Database"
     sprocCheckClaim.CommandText = "usp_AUDITCLM_CheckClaimExists"
     sprocCheckClaim.Setup
        
    sprocMarkClaim.RefTable = "v_CODE_Database"
    
    If sClmLevel = "L" Then
        sprocMarkClaim.CommandText = "usp_POOL_SelectDtlForSubmission"
        lLinePosition = GetColumnPosition(Lst, "LineNum")
    Else
        sprocMarkClaim.CommandText = "usp_POOL_SelectHdrForSubmission"
    End If
            
    sprocMarkClaim.Setup
        
'* Mark Selected Claim lines

       For Each varItem In Lst.ItemsSelected
            
            sClaimNum = Lst.Column(lClaimNum, varItem)
            
            If sClmLevel = "L" Then
            
                lLineNum = Lst.Column(lLinePosition, varItem)
                
            End If
              
      '* Validate Row
        
            sprocCheckClaim.AddParam "@pCnlyClaimNum", sClaimNum
            sprocCheckClaim.Exec
                
            If sprocCheckClaim.GetParam("@pClaimExists") = True Then
                Me.Parent.lstStatus.AddItem sClaimNum & " is already a claim."
                GoTo exitHere
            End If
            
            sprocCheckPool.AddParam "@pCnlyClaimNum", sClaimNum
            
            If sClmLevel = "L" Then
                sprocCheckPool.AddParam ("@pLineNum"), lLineNum
            End If
            
            sprocCheckPool.Exec
            
            lInPool = sprocCheckPool.GetParam("@pClaimExists")
            
            Select Case lInPool
            
                Case Is = 1
                
                    If sClmLevel = "L" Then   '* Is in pool at header level
                        lstStatus.AddItem sClaimNum & " in pool at header level."
                    Else
                        lstStatus.AddItem sClaimNum & " already in pool."
                    End If
                    
                Case Is = 2
                    
                    If sClmLevel = "L" Then  '* Line In Pool
                        lstStatus.AddItem sClaimNum & " " & CStr(lLineNum) & " already in pool."
                    Else '* sClmLevel = "H"
                        lstStatus.AddItem sClaimNum & " in pool at header level."
                    End If
                    
                Case Else
                
                    sprocMarkClaim.AddParam "@pCnlyClaimNum", Lst.Column(lClaimNum, varItem)
                    If sClmLevel = "L" Then
                        sprocMarkClaim.AddParam "@pLineNum", lLineNum
                    End If
                    
                    'Rob Swander 07/07/2010: changed @ConceptCd to @pConceptCd
'                    sprocMarkClaim.AddParam "@ConceptCd", Lst.Column(lConceptCd, varItem)
                    sprocMarkClaim.AddParam "@pConceptCd", Lst.Column(lConceptCd, varItem)

                    
                    sprocMarkClaim.Exec
                    
                    If sprocMarkClaim.ReturnValue = 1 Then
                        MsgBox "Error flagging claim", vbOKOnly, "Error flagging claim"
                        GoTo exitHere
                    End If
                    
                     Me.Parent.lstStatus.AddItem sClaimNum & IIf(sClmLevel = "L", " " & CStr(lLineNum), "") & " added"
              
            End Select
        
        Next varItem

exitHere:
    On Error Resume Next
    
    Set Lst = Nothing
    Set sprocCheckClaim = Nothing
    Set sprocMarkClaim = Nothing
    Set sprocCheckPool = Nothing

    DoCmd.Hourglass False
    
    Exit Sub
HandleError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GoTo exitHere
End Sub

Private Sub cmdMoveToClm_Click()

  Dim sprocAddClaims As clsAdoSproc

    Set sprocAddClaims = New clsAdoSproc

    If mGroupName = "" Then
        MsgBox "setup error"
        GoTo exitHere
    End If

    If Nz(Me.cboAuditNum, "") = "" Then
        MsgBox "No Audit Number Selected."
        GoTo exitHere
    End If

' thieu add account id

    sprocAddClaims.RefTable = "v_CODE_Database"
    sprocAddClaims.CommandText = "usp_POOL_MoveToAuditClaims"
    sprocAddClaims.Setup
    
    sprocAddClaims.AddParam ("@pAccountID"), gintAccountID
    sprocAddClaims.AddParam ("@pAuditNum"), Me.cboAuditNum
    sprocAddClaims.AddParam ("@pErrMsg"), ""
    
    sprocAddClaims.Exec
        
    If sprocAddClaims.ReturnValue = 0 Then
         Me.Parent.lstStatus.AddItem "Claims Added"
    Else
        MsgBox "Error Moving Claims. " & vbCr & sprocAddClaims.GetParam("@pErrMsg")
    End If
        
    
exitHere:
    On Error Resume Next
    
    Set sprocAddClaims = Nothing

    Exit Sub
HandleError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GoTo exitHere
End Sub

Private Sub cmdRollBack_Click()
    
    If miStackID = 0 Then
        MsgBox "You're already at the top view"
    Else
        Call RollBackCurrentState
        Call RefreshData
    End If

End Sub



Private Sub Form_Close()
    lstHdr.RowSource = ""
    lstDetail.RowSource = ""
End Sub

Private Sub Form_Load()

    Me.Caption = "Pool Management"
    
    Call Account_Check(Me)
    
    mstrUserName = Identity.UserName
    miStackID = 0
    mbRefreshSelection = True
    
    RefreshComboBox "select * from ADMIN_Audit_Number WHERE AccountID = " & gintAccountID & "", Me.cboAuditNum
    
    ' initialize current state
    mCurrState.Hdr_ViewOrder = ""
    mCurrState.Hdr_SQL = ""
    mCurrState.Hdr_RowSelected = 0
    mCurrState.Dtl_ViewOrder = ""
    mCurrState.Dtl_SQL = ""
    mCurrState.Dtl_RowSelected = 0
    mCurrState.Dtl_Combo = ""
    
    Set mColStates = New Collection
    
    mOperMode = OperationMode.Manager
    mGroupName = "POOL"
        
    If IsSubForm(Me) Then
        Select Case UCase(Nz(Me.Parent.Parameter, "POOL_HDR"))
            Case "POOL_HDR"
                mGroupName = "POOL"
            Case "POOL_DTL"
                mGroupName = "POOL_Dtl"
            Case "POOL_HDR_SUBMIT"
                mGroupName = "POOL_Hdr_SUBMIT"
            Case "POOL_DTL_SUBMIT"
                mGroupName = "POOL_Dtl_Submit"
            Case Else
                mGroupName = "POOL"
        End Select
    Else
            mGroupName = "POOL"
    End If
    
    Set MyAdo = New clsADO
    
    ' read display and age calculation configuration
    ' TL add account ID logic
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from GENERAL_Display_Config where GroupName = '" & mGroupName & "'" & _
                      " and ComboBoxOrder <> 0" & _
                      " and AccountID = " & gintAccountID & _
                      " order by ComboBoxOrder"
    Set mrsDisplayConfig = MyAdo.OpenRecordSet()
    
    
       
    ' TL add account ID logic
    MyAdo.sqlString = "select * from GENERAL_Display_Config where GroupName = '" & mGroupName & "'" & _
                      " and ComboBoxOrder = 0" & _
                      " and AccountID = " & gintAccountID
    Set mrsAgeCalConfig = MyAdo.OpenRecordSet()
    
     ' set detail view column
    If mrsDisplayConfig.BOF = True And mrsDisplayConfig.EOF = True Then
        MsgBox "POOL display configuration is not setup for this account.  Please check with your administrator", vbCritical
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
    
    Me.cmdMarkLine.Enabled = False
    
    If mGroupName = "POOL_HDR_SUBMIT" Then
        Me.cmdMoveToClm.visible = True
        Me.cmdMarkLine.visible = False
    ElseIf mGroupName = "POOL_Dtl_SUBMIT" Then
        Me.cmdMoveToClm.visible = True
        Me.cmdMarkLine.visible = False
    Else
        Me.cmdMoveToClm.visible = False
    End If
    
    Set MyAdo = Nothing
    
End Sub

Private Sub lstDetail_DblClick(Cancel As Integer)

    If cboDetail.Column(0) = mstrDetailView Then
     '* JAC  MsgBox "Detail"
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
        
        Set myCode_ADO = Nothing
    End If
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
    Dim i As Integer
    Dim j As Integer
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
    Dim i As Integer
    Dim j As Integer

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


Private Function AppendCurrentState() As String
    'AppendCurrentState = CurrState & States
End Function

Private Function RollBackCurrentState() As Boolean

    Dim temp As Variant
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

Private Function SetViewSQL(QueueLevel As String) As String
    Dim ViewOrder
    Dim SQL As SQLConstruct
    Dim strSQL As String
    Dim iViewOrder As Integer
    Dim bDetailView As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim strValue As String
    Dim iLstColOrder As Integer
    
    SQL.Select = ""
    SQL.Where = ""
    SQL.GroupBy = ""
    SQL.OrderBy = ""
    
    '*JAC  this is listview recordsource
    
    Select Case mGroupName
    
        Case Is = "pool"
            SQL.From = "from v_POOL_Concept_Hdr"
        Case Is = "Pool_Dtl"
            SQL.From = "from v_POOL_Concept_DTL"
        Case Is = "Pool_Hdr_Submit"
            SQL.From = "from v_POOL_Hdr_Submission"
        Case Is = "POOL_DTL_Submit"
            SQL.From = "from v_POOL_DTL_Submission"
        
        Case Else
            MsgBox "No Setup info"
    End Select
    
        
    bDetailView = False
    
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
                        If UCase(mrsAgeCalConfig("CalcAge")) = "Y" Then
                            SQL.Select = SQL.Select & "(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1) as " & mrsAgeCalConfig("DispColName") & ", "
                            SQL.OrderBy = SQL.OrderBy & "(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1), "
                        End If
                        mrsAgeCalConfig.MoveNext
                    Wend
                End If
                
                SQL.Select = SQL.Select & " * "
                If mrsAgeCalConfig.BOF = True And mrsAgeCalConfig.EOF = True Then
                    SQL.OrderBy = ""
                Else
                    If SQL.OrderBy = "order by " Then
                        SQL.OrderBy = ""
                    Else
                        SQL.OrderBy = left(Trim(SQL.OrderBy), Len(Trim(SQL.OrderBy)) - 1)
                    End If
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
                If UCase(mrsAgeCalConfig("CalcAge")) = "Y" Then
                    SQL.Select = SQL.Select & ", max(datediff(d," & mrsAgeCalConfig("CtrlColumn") & ",getdate())+1) as " & mrsAgeCalConfig("DispColName")
                End If
                
                If UCase(mrsAgeCalConfig("CalcTotal")) = "Y" Then
                    SQL.Select = SQL.Select & ", sum(" & mrsAgeCalConfig("CtrlColumn") & ") as " & mrsAgeCalConfig("DispColName")
                End If
                
                mrsAgeCalConfig.MoveNext
            Wend
        End If
    End If
    
       
      
    strSQL = SQL.Select & Space(1) & SQL.From & Space(1) & SQL.Where & Space(1) & SQL.GroupBy & Space(1) & SQL.OrderBy
    SetViewSQL = strSQL
End Function

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
