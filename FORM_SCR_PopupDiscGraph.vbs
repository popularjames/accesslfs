Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private IntInvAmtField As String
Private IntDiscAmtField As String
Private IntVenField As String
Private IntVenType As Byte
Private IntInvDateField As String
Private IntChkDateField As String
Private IntGraphSource As String
Private IntExtraCriteria As String
Private IntVenNum As String
Private IntDateCriteria As String
Private IntDateFrom As String
Private IntDateTo As String

Private Splitting As Boolean

Property Let ExtraCriteria(Criteria As String)
    IntExtraCriteria = Criteria
    Me!lblFilter.Caption = Criteria
End Property
Property Let DateCriteria(Criteria As String)
    IntDateCriteria = Criteria
End Property
Property Let DateCriteriaTo(Criteria As String)
    IntDateTo = Criteria
    ' HC 5/2010 removed for 2010
    'If "" & Me.EndDte.Object.Value <> IntDateTo Then Me.EndDte.Object.Value = IntDateTo
    ' HC 5/2010 updated for 2010
    If "" & Me.EndDte.Value <> IntDateTo Then Me.EndDte.Value = IntDateTo
End Property
Property Let DateCriteriaFrom(Criteria As String)
    IntDateFrom = Criteria
    ' HC 5/2010 removed for 2010
    'If "" & Me.StartDte.Object.Value <> IntDateFrom Then Me.StartDte.Object.Value = IntDateFrom
    ' HC 5/2010 updated for 2010
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
Private Sub Resize(Optional ChangeSize)
Static LastSplitterTop As Single

If IsMissing(ChangeSize) Then GoTo ExitIt


Me.DtlGraph.Height = Me.Splitter.top - Me.DtlGraph.top
With Me.DtlGraph2
    If LastSplitterTop - Me.Splitter.top < 0 Then
        .Height = Me.Section(acDetail).Height - (Me.Splitter.top + Me.Splitter.Height) - 80
        .top = (Me.Splitter.top + Me.Splitter.Height)
    Else
        .top = Me.Splitter.top + Me.Splitter.Height
        .Height = Me.Section(acDetail).Height - .top
    End If
End With

ExitIt:
LastSplitterTop = Me.Splitter.top
End Sub

Property Get VenNum() As String
    VenNum = IntVenNum
End Property
Property Get InvDateField() As String
    InvDateField = IntInvDateField
End Property

Property Let InvDateField(FieldName As String)
    IntInvDateField = FieldName
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




Public Sub InitGraph(ByList As Boolean)
Dim SQL As String

'Private IntInvAmtField As String
'Private IntDiscAmtField As String
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
'SELECT ScrApData.InvDte AS ByDate, Sum(ScrApData.DiscAmt) AS TotalDiscAmt, Sum(ScrApData.GrossAmt) AS TotalInvAmt FROM ScrApData WHERE (((ScrApData.VenNum)='M00162')) GROUP BY ScrApData.InvDte HAVING (((Sum(ScrApData.DiscAmt))>0));
SQL = "SELECT Format(" & IIf(ByList, IntInvDateField, IntChkDateField) & ",'mm/dd/yy') AS ByDate " & ", "
SQL = SQL & "Sum(" & IntInvAmtField & ") AS [Volume], "
SQL = SQL & "Sum(1) as CT "
SQL = SQL & "From " & IntGraphSource & " "
SQL = SQL & "Where " & IntInvAmtField & " > 0 and " & IntVenField & " =" & GetIdentifier(IntVenType) & IntVenNum & GetIdentifier(IntVenType) & " "
If Len(IntExtraCriteria) > 0 Then SQL = SQL & "And " & IntExtraCriteria & " "
SQL = SQL & "GROUP BY " & IIf(ByList, IntInvDateField, IntChkDateField)
'Debug.Print SQL
Me!DtlGraph.RowSource = SQL

SQL = "SELECT Format(" & IIf(ByList, IntInvDateField, IntChkDateField) & ",'mm/dd/yy') AS ByDate " & ", "
SQL = SQL & "Avg(Clng(Abs(" & IntDiscAmtField & ")/" & IntInvAmtField & "*10000)/100) AS [Avg Disc Pct], "
SQL = SQL & "Avg(DateDiff('d'," & IntInvDateField & "," & IntChkDateField & ")) AS [Avg Days To Pay] "
SQL = SQL & "From " & IntGraphSource & " "
SQL = SQL & "Where " & IntInvAmtField & " > 0 and " & IntVenField & " =" & GetIdentifier(IntVenType) & IntVenNum & GetIdentifier(IntVenType) & " "
If Len(IntExtraCriteria) > 0 Then SQL = SQL & "And " & IntExtraCriteria & " "
SQL = SQL & "GROUP BY " & IIf(ByList, IntInvDateField, IntChkDateField)
'Debug.Print SQL
Me!DtlGraph2.RowSource = SQL
End Sub

Property Get VenField() As String
    VenField = IntVenField
End Property

Property Let VenType(DataType As Byte)
    IntVenType = DataType
End Property

Property Get VenType() As Byte
    VenType = IntVenType
End Property
Property Get ChkDateField() As String
    ChkDateField = IntChkDateField
End Property



Private Sub cmdPrint_Click()

On Error GoTo Err_Command2_Click


    DoCmd.PrintOut

Exit_Command2_Click:
    Exit Sub

Err_Command2_Click:
    MsgBox Err.Description
    Resume Exit_Command2_Click

End Sub

Property Get DiscAmtField() As String
    DiscAmtField = IntDiscAmtField
End Property

Property Let InvAmtField(FieldName As String)
    IntInvAmtField = FieldName
End Property
Property Let ChkDateField(FieldName As String)
    IntChkDateField = FieldName
End Property

Property Let VenField(FieldName As String)
    IntVenField = FieldName
End Property
Property Let DiscAmtField(FieldName As String)
    IntDiscAmtField = FieldName
End Property

Private Sub cmdRefresh_Click()
On Error GoTo RefreshError
Dim SQL As String

With Me
    If .DateCriteria = "" Then .DateCriteria = IIf(.OptCost, IntInvDateField, IntChkDateField)
    If .DateCriteria = "" Then Exit Sub        'If no date criteria provided - Exit
    ' HC 5/2010 - removed for 2010
'    .DateCriteriaFrom = Me.StartDte.Object.Value
'    .DateCriteriaTo = Me.EndDte.Object.Value
    ' HC 5/2010 - 2010
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

Private Sub Command0_Click()
On Error GoTo Err_Command0_Click


    DoCmd.Close

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click
    
End Sub





Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub DtlGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub DtlGraph2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub Form_Deactivate()
screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Resize
    Resize 1
End Sub

Private Sub Form_LostFocus()
screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
screen.MousePointer = 0
End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub


Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub OptCost_AfterUpdate()
InitGraph OptCost
End Sub



Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Splitting = True
    Me.Splitter.BackColor = 255
End If
End Sub


Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ExitIt
If Button = 0 Then screen.MousePointer = 7
If Button = 1 Then

    If (Me.Splitter.top + Y > Me.Section(acDetail).Height - 200) Or (Me.Splitter.top + Y < 200) Then Exit Sub
    Me.Splitter.top = Me.Splitter.top + Y
    Resize Y
    Me.Repaint
End If

ExitIt:
End Sub


Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    If (Me.Splitter.top + Y > Me.Section(acDetail).Height - 200) Or (Me.Splitter.top + Y < 200) Then Exit Sub
    Me.Splitter.top = Me.Splitter.top + Y
    Resize Y
    Me.Splitter.BackColor = 12632256
    Me.Repaint
End If
screen.MousePointer = 0
End Sub
