Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private IntListCostField As String
Private IntNetCostField As String
Private IntVenField As String
Private IntVenType As Byte
Private IntItemField As String
Private IntItemType As Byte
Private IntGraphSource As String
Private IntGroupField As String
Private IntExtraCriteria As String
Private IntQtyField As String
Private IntVenNum As String
Private IntItemNum As String
Private IntDateCriteria As String
Private IntDateFrom As String
Private IntDateTo As String


Property Let ExtraCriteria(Criteria As String)
    IntExtraCriteria = Criteria
    Me!lblFilter.Caption = Criteria
End Property
Property Let DateCriteria(Criteria As String)
    IntDateCriteria = Criteria
End Property
Property Let DateCriteriaTo(Criteria As String)
    IntDateTo = Criteria
    ' HC 5/2010 - removed 2010
    'If "" & Me.EndDte.Object.Value <> IntDateTo Then Me.EndDte.Object.Value = IntDateTo
    'HC 5/2010 - updated 2010
    If "" & Me.EndDte.Value <> IntDateTo Then Me.EndDte.Value = IntDateTo
End Property
Property Let DateCriteriaFrom(Criteria As String)
    IntDateFrom = Criteria
    ' HC 5/2010- removed 2010
    'If "" & Me.StartDte.Object.Value <> IntDateFrom Then Me.StartDte.Object.Value = IntDateFrom
    ' HC 5/2010 - updated 2010
    If "" & Me.StartDte.Value <> IntDateFrom Then Me.StartDte.Value = IntDateFrom
End Property

Property Get ExtraCriteria() As String
    ExtraCriteria = IntExtraCriteria
End Property
Property Get DateCriteria() As String
    DateCriteria = IntDateCriteria
End Property
Property Get DateCriteriaTo() As String
    DateCriteriaTo = IntDateTo
End Property
Property Get DateCriteriaFrom() As String
    DateCriteriaFrom = IntDateFrom
End Property
Property Get VenNum() As String
    VenNum = IntVenNum
End Property
Property Get ItemNum() As String
    VenNum = IntItemNum
End Property

Property Let ItemNum(ItemNum As String)
    IntItemNum = ItemNum
    Me!LblItemNum.Caption = ItemNum
End Property

Property Let VenNum(VenNum As String)
    IntVenNum = VenNum
    Me!LblVenNum.Caption = IntVenNum
End Property
Property Let GraphSource(Source As String)
    IntGraphSource = Source
End Property

Property Get GraphSource() As String
    GraphSource = IntGraphSource
End Property
Property Let GroupField(FieldName As String)
    IntGroupField = FieldName
End Property

Property Get GroupField() As String
    GroupField = IntGroupField
End Property



Public Sub InitGraph(ByList As Boolean)
Dim SQL As String

'Private IntListCostField As String
'Private IntNetCostField As String
'Private IntVenField As String
'Private IntVenType As Byte
'Private IntItemField As String
'Private IntItemType As Byte
'Private IntGraphSource As String
'Private IntGroupField As String
'Private IntExtraCriteria As String
'Private IntQtyField As String
'Private IntVenNum As String
'Private IntItemNum As String

' Build SQL for Graph

SQL = "SELECT " & IntGroupField & ", Sum(" & IntQtyField & ") AS SumOfCalcRecQty, "
SQL = SQL & "Avg(" & IIf(ByList, IntListCostField, IntNetCostField) & ") AS AvgOfPaidCost "
SQL = SQL & "From " & IntGraphSource & " "
SQL = SQL & "Where " & IntVenField & " =" & GetIdentifier(IntVenType) & IntVenNum & GetIdentifier(IntVenType) & " "
SQL = SQL & "And " & IntItemField & " =" & GetIdentifier(IntItemType) & IntItemNum & GetIdentifier(IntItemType) & " "
If Len(IntExtraCriteria) > 0 Then SQL = SQL & "And " & IntExtraCriteria & " "
SQL = SQL & "GROUP BY " & IntGroupField
'Debug.Print SQL
Me!DtlGraph.RowSource = SQL
End Sub

Property Get VenField() As String
    VenField = IntVenField
End Property

Property Get ItemType() As Byte
    ItemType = IntItemType
End Property
Property Let VenType(DataType As Byte)
    IntVenType = DataType
End Property
Property Let ItemType(DataType As Byte)
    IntItemType = DataType
End Property

Property Get VenType() As Byte
    VenType = IntVenType
End Property
Property Get ItemField() As String
    ItemField = IntItemField
End Property

Property Get QtyField() As String
    QtyField = IntQtyField
End Property

Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click


    DoCmd.Close acForm, Me.Name, acSaveNo

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
End Sub

Private Sub cmdPrint_Click()

On Error GoTo Err_Command2_Click


    DoCmd.PrintOut

Exit_Command2_Click:
    Exit Sub

Err_Command2_Click:
    MsgBox Err.Description
    Resume Exit_Command2_Click

End Sub

Property Get NetCostField() As String
    NetCostField = IntNetCostField
End Property

Property Let ListCostField(FieldName As String)
    IntListCostField = FieldName
End Property
Property Let ItemField(FieldName As String)
    IntItemField = FieldName
End Property

Property Let VenField(FieldName As String)
    IntVenField = FieldName
End Property
Property Let QtyField(FieldName As String)
    IntQtyField = FieldName
End Property
Property Let NetCostField(FieldName As String)
    IntNetCostField = FieldName
End Property

Private Sub cmdRefresh_Click()
On Error GoTo RefreshError
Dim SQL As String

With Me
    If .DateCriteria = "" Then .DateCriteria = .GroupField
    If .DateCriteria = "" Then Exit Sub 'If no date criteria provided - Exit
    ' HC 5/2010 - removed 2010
    '.DateCriteriaFrom = Me.StartDte.Object.Value
    '.DateCriteriaTo = Me.EndDte.Object.Value
    ' HC 5/2010 - updated 2010
    .DateCriteriaFrom = Me.StartDte.Value
    .DateCriteriaTo = Me.EndDte.Value
    SQL = .DateCriteria & " Between #" & Me.DateCriteriaFrom & "# And #" & Me.DateCriteriaTo & "#"
    Me.ExtraCriteria = SQL
    Me.InitGraph OptCost
End With


RefreshExit:
    On Error Resume Next
    Exit Sub

RefreshError:
    MsgBox Err.Description
    Resume RefreshExit

End Sub






Private Sub OptCost_AfterUpdate()
InitGraph OptCost
End Sub
