Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'' Last Modified: 05/05/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''  The purpose of this is to
''
''
''
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/05/2013  - Created
''
'' AUTHOR
''  =====================================
'' Kevin Dearing
''
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################


Private coItem As clsBOLD_LetterRuleItemDetail

Public Event ItemChanged()

Private cbDirty As Boolean
Private cbInitialized As Boolean


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ItemName() As String
    ItemName = coItem.ItemName
End Property
Public Property Let ItemName(sItemName As String)
    coItem.ItemName = sItemName
End Property


Public Property Get LocalId() As Long
    LocalId = coItem.Id
End Property
Public Property Let LocalId(lLocalId As Long)
    coItem.Id = lLocalId
End Property


Public Property Get RuleId() As Long
    RuleId = coItem.RuleId
End Property
Public Property Let RuleId(lRuleId As Long)
    coItem.RuleId = lRuleId
End Property


Public Property Get RuleItemId() As Long
    RuleItemId = coItem.RuleItemId
End Property
Public Property Let RuleItemId(lRuleItemId As Long)
    coItem.RuleItemId = lRuleItemId
End Property


Public Property Get BooleanVal() As Long
    BooleanVal = coItem.BooleanVal
End Property
Public Property Let BooleanVal(lBooleanVal As Long)
    coItem.BooleanVal = lBooleanVal
End Property


Public Property Get Operator() As Long
    Operator = coItem.Operator
End Property
Public Property Let Operator(lOperator As Long)
    coItem.Operator = lOperator
End Property


Public Property Get ItemValue() As String
    ItemValue = coItem.ItemValue
End Property
Public Property Let ItemValue(sItemValue As String)
    coItem.ItemValue = sItemValue
End Property


Public Property Get IsDirty() As Boolean
    IsDirty = cbDirty
End Property
Public Property Let IsDirty(bIsDirty As Boolean)
    cbDirty = bIsDirty
End Property


Public Property Get WasInitialized() As Boolean
    WasInitialized = cbInitialized
End Property
Public Property Let WasInitialized(bInitialized As Boolean)
    cbInitialized = bInitialized
End Property



Public Function InitDataObj(oItem As clsBOLD_LetterRuleItemDetail) As Boolean
    Set coItem = oItem
    InitDataObj = True
    Call RefreshForm
End Function



Public Function InitData(lLocalId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".InitData"
    Set coItem = New clsBOLD_LetterRuleItemDetail
    
    InitData = coItem.LoadFromId(lLocalId)
    Call RefreshForm
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function RefreshForm() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".RefreshForm"
    If coItem Is Nothing Then GoTo Block_Exit
    
    With coItem
        Me.txtItemId = .Id
        Me.txtItemName = .ItemName
        Me.frmAndOr = .BooleanVal
        Me.cmbOperator = .Operator
        Me.txtValue = .ItemValue
    End With

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub cmbOperator_Change()
    DataChanged
End Sub

Private Sub CmdAdd_Click()
    If Me.IsDirty = True Then
        coItem.SaveNow
    End If
    Set coItem = Nothing
    Me.visible = False
''    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub CmdCancel_Click()
    Set coItem = Nothing
    Me.visible = False
End Sub

Private Sub Command6_Click()
Stop
    Call Me.InitData(19)

End Sub



Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String

    strProcName = ClassName & ".Form_Load"
    
    Set coItem = New clsBOLD_LetterRuleItemDetail
    
    sSql = "SELECT OperatorId, OperatorName FROM BOLD_Letter_Automation_XREF_Operators WHERE Active <> 0  "
    
    '' Now need to get the combo box:
    Me.cmbOperator.ColumnCount = 2
    Me.cmbOperator.ColumnWidths = "0;20"
    Me.cmbOperator.ColumnCount = 2
    
    Call RefreshComboBoxADO(sSql, Me.cmbOperator, , , "v_Data_Database")
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub frmAndOr_AfterUpdate()
    DataChanged
End Sub


Private Sub optAnd_Click()
    DataChanged
End Sub

Private Sub optOr_Click()
    DataChanged
End Sub

Private Sub txtValue_BeforeUpdate(Cancel As Integer)
    DataChanged
End Sub


Public Function DataChanged() As Boolean

    With coItem
        .BooleanVal = Me.frmAndOr
        .Operator = Me.cmbOperator
        .ItemValue = Me.txtValue
    End With
    Me.IsDirty = True
    
End Function

Private Sub txtValue_Dirty(Cancel As Integer)
    DataChanged
End Sub
