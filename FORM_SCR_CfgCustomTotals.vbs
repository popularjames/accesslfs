Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 
'-- MODULE VARIABLES --'
Private MvScreenID As Long
Private mvTotalID As Long
Private mvFormId As Long
Private mvResults As Long
Private mvSource As String ' Name of the primary RecordSource
Private Enum SQLMethod
    Add = 1
    Delete = 2
End Enum
Private WithEvents MvImport As CT_ClsImport
Attribute MvImport.VB_VarHelpID = -1
Public Function CreateNewTotalName() As String
  CreateNewTotalName = Identity.UserName & " - " & Format(Now, "yyyy/mm/dd hh:nn:ss")
    Me.LblScreenName2 = "" & IIf([Global] = True, "GLB", [UserName]) & ": " & [TotalName]
End Function

Public Property Let CurrentScreenID(Value As Long)
    MvScreenID = Value
End Property
Public Property Get CurrentScreenID() As Long
    CurrentScreenID = mvTotalID
End Property
Public Property Let CurrentTotalID(Value As Long)
    mvTotalID = Value
End Property
Public Property Get CurrentTotalID() As Long
    CurrentTotalID = mvTotalID
End Property
Public Property Let CurrentFormID(Value As Long)
    mvFormId = Value
        mvSource = Scr(mvFormId).Config.PrimaryRecordSource
End Property
Public Property Get CurrentFormID() As Long
    CurrentFormID = mvFormId
End Property
Public Property Get Results() As Long
    Results = mvResults
End Property
Public Sub show(Optional NewTotal As Boolean = False)
Me.visible = True
DoEvents
    If NewTotal = True Then
        Call DoCmd.GoToRecord(acDataForm, Me.Name, acNewRec)
        Me.ScreenID = MvScreenID
        Me.UserName = Identity.UserName
        Me.TotalName = CreateNewTotalName()
        Me.Dirty = False
        PopulateLists
    Else
        Me.filter = "TotalID = " & mvTotalID
        Me.FilterOn = True
    End If
    Me.LblScreenName2 = "" & IIf([Global] = True, "GLB", [UserName]) & ": " & [TotalName]

End Sub
Private Sub cmdClose_Click()
    DBEngine.Idle dbRefreshCache + dbForceOSFlush
    DoEvents
    Me.Dirty = False
    mvResults = vbOK
    Me.visible = False
   
End Sub

Private Sub cmdCopy_Click()
On Error GoTo ErrorHappened
Dim Pairs(1) As ReplacePairs

Pairs(0).From = "$TotalName"
Pairs(0).To = CreateNewTotalName()


Pairs(1).From = "$TotalID"
Pairs(1).To = Me.TotalID

Set MvImport = New CT_ClsImport

With MvImport
    If .RunUtilityEx(6, CurrentDb.Name, Pairs()) = True Then
        mvTotalID = DLookup("TotalID", "SCR_ScreensTotals", "RefID=" & Pairs(1).To)
        Me.show
        Me.TotalName.BackColor = 255
        Me.TotalName.SetFocus
        Me.TimerInterval = 2000
    End If
End With

ExitNow:
    On Error Resume Next
    Set MvImport = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Loading Total Config Form"
    Resume ExitNow
    Resume
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim SQL As String

    If MsgBox("Are you sure you want to delete the following Custom Totals:" & vbCrLf & vbCrLf & Me.TotalName, vbQuestion + vbYesNo + vbDefaultButton2, "Delete Totals") = vbYes Then
        SQL = "Delete SCR_ScreensTotals.TotalID From SCR_ScreensTotals Where TotalID = " & Me.TotalID
        Set db = CurrentDb
        CurrentDb.Execute SQL, dbFailOnError
    End If

    mvResults = -2
    Me.visible = False

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Loading Total Config Form"
    Resume ExitNow
    Resume
End Sub

Private Sub CmdExpAdd_Click()
On Error GoTo ErrorHappened

Dim X As Integer, Ordinal As Long


For X = 0 To Me.LstExpFlds.ItemsSelected.Count - 1
    Ordinal = Nz(DMax("Ordinal", "SCR_ScreensTotalsCalculations", "TotalID = " & Me.TotalID), 0) + 1000
    SQLExpressions SQLMethod.Add, LstExpFlds.ItemData(LstExpFlds.ItemsSelected.Item(X)), Ordinal
Next X


LstExpFldsSel.Requery

'Deselect

For X = Me.LstGrpFlds.ItemsSelected.Count - 1 To 0 Step -1
    LstGrpFlds.Selected(LstGrpFlds.ItemsSelected.Item(X)) = False
Next X


ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Adding Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdExpFldsDown_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim id1 As Long, id2 As Long
Dim Ord1 As Long, Ord2 As Long
Dim index As Long, SQL As String

index = Me.LstExpFldsSel.ListIndex

Select Case index
Case -1, (Me.LstExpFldsSel.ListCount - 1) 'Nothing or At the bottom
    GoTo ExitNow
Case Else
    id1 = LstExpFldsSel.Column(0, index)
    Ord1 = LstExpFldsSel.Column(4, index)
    id2 = LstExpFldsSel.Column(0, index + 1)
    Ord2 = LstExpFldsSel.Column(4, index + 1)
End Select

SQL = "UPDATE SCR_ScreensTotalsCalculations "
SQL = SQL & "SET Ordinal = IIF(TotalCalcID = " & id1 & "," & Ord2 & "," & Ord1 & ") "
SQL = SQL & "WHERE TotalCalcID in (" & id1 & "," & id2 & ")"

Set db = CurrentDb
db.Execute SQL, dbFailOnError

ExitNow:
On Error Resume Next
    Set db = Nothing
    LstExpFldsSel.Requery
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Ordering Fields"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdExpFldsEdit_Click()
FieldEdit Calculations
End Sub
Private Sub FieldEdit(eMode As mode)
Dim FrmZoom As Form_SCR_CfgCustomTotalsExpression
Dim Txt As String
Dim ctrl As Control

Select Case eMode
Case mode.Calculations
    If Me.LstExpFldsSel.ListIndex = -1 Then GoTo ExitNow
    Set ctrl = LstExpFldsSel
Case mode.Grouping
    If Me.LstGrpFldsSel.ListIndex = -1 Then GoTo ExitNow
    Me.LstGrpFldsSel.Requery
    Set ctrl = LstGrpFldsSel
End Select
    
Txt = "" & LstExpFldsSel.Column(2, LstExpFldsSel.ListIndex)


Set FrmZoom = New Form_SCR_CfgCustomTotalsExpression
FrmZoom.Move ctrl.left

With FrmZoom
    .RunMode = eMode
    
    Select Case eMode
    Case mode.Calculations
        .filter = "TotalCalcID = " & LstExpFldsSel
    Case mode.Grouping
        .filter = "TotalFldID = " & LstGrpFldsSel
        Me.LstGrpFldsSel.Requery
    End Select

    
    .FilterOn = True
    .Modal = True
    .visible = True
    
    Do Until .Results <= 0
        DoEvents
    Loop
End With

DBEngine.Idle dbRefreshCache + dbForceOSFlush

DoEvents

Select Case eMode
Case mode.Calculations
    LstExpFldsSel.Requery
Case mode.Grouping
    Me.LstGrpFldsSel.Requery
End Select

ExitNow:
On Error Resume Next
    Set FrmZoom = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Editing Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub
Private Sub CmdExpFldsUp_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim id1 As Long, id2 As Long
Dim Ord1 As Long, Ord2 As Long
Dim index As Long, SQL As String

index = Me.LstExpFldsSel.ListIndex

Select Case index
Case 0, -1 'Top, Nothing
    GoTo ExitNow
Case Else
    id1 = LstExpFldsSel.Column(0, index)
    Ord1 = LstExpFldsSel.Column(4, index)
    id2 = LstExpFldsSel.Column(0, index - 1)
    Ord2 = LstExpFldsSel.Column(4, index - 1)
End Select

SQL = "UPDATE SCR_ScreensTotalsCalculations "
SQL = SQL & "SET Ordinal = IIF(TotalCalcID = " & id1 & "," & Ord2 & "," & Ord1 & ") "
SQL = SQL & "WHERE TotalCalcID in (" & id1 & "," & id2 & ")"

Set db = CurrentDb
db.Execute SQL, dbFailOnError


ExitNow:
On Error Resume Next
    
    Set db = Nothing
    LstExpFldsSel.Requery
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Ordering Fields"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdExpRemove_Click()
On Error GoTo ErrorHappened

Dim X As Integer


For X = 0 To Me.LstExpFldsSel.ItemsSelected.Count - 1
    SQLExpressions SQLMethod.Delete, LstExpFldsSel.ItemData(LstExpFldsSel.ItemsSelected.Item(X))
Next X

LstExpFldsSel.Requery

If LstExpFldsSel.ListCount > 0 Then
    LstExpFldsSel = LstExpFldsSel.ItemData(0)
End If

ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Removing Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdGrpAdd_Click()
On Error GoTo ErrorHappened

Dim X As Integer, Ordinal As Long



For X = 0 To Me.LstGrpFlds.ItemsSelected.Count - 1
    Ordinal = Nz(DMax("Ordinal", "SCR_ScreensTotalsFields", "TotalID = " & Me.TotalID), 0) + 1000
    SQLGrouping SQLMethod.Add, LstGrpFlds.ItemData(LstGrpFlds.ItemsSelected.Item(X)), Ordinal
Next X

LstGrpFldsSel.Requery

'Deselect

For X = Me.LstGrpFlds.ItemsSelected.Count - 1 To 0 Step -1
    LstGrpFlds.Selected(LstGrpFlds.ItemsSelected.Item(X)) = False
Next X


ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Adding Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub SQLGrouping(Method As SQLMethod, Name As String, Optional Ordinal As Long = 1000)
On Error GoTo ErrorHappened

Dim db As DAO.Database
Dim SQL As String

Set db = CurrentDb
Select Case Method
Case SQLMethod.Add
    SQL = "Insert into SCR_ScreensTotalsFields(TotalID, FldType,FldName,Ordinal) "
    SQL = SQL & "Values (" & TotalID & ",1," & Chr(34) & Name & Chr(34) & "," & Ordinal & ")"
Case SQLMethod.Delete
    SQL = "Delete SCR_ScreensTotalsFields.TotalFldID "
    SQL = SQL & "FROM SCR_ScreensTotalsFields "
    SQL = SQL & "Where TotalID = " & TotalID & " and "
    SQL = SQL & "    TotalFldID = " & Name 'More lazy coding
End Select

db.Execute SQL, dbFailOnError

ExitNow:
On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    Select Case Err.Number
    Case 3022 'Duplicate Index violation - Ignoring this because I am lazy about checking if the val exists before adding
        LogMessage "Duplicate Key - Item Skipped"
        
        Resume Next
    Case Else
        MsgBox Err.Description, vbCritical, "Custom Totals: Error Adding Field"
        LogMessage Err.Description
        Resume ExitNow
        Resume
    End Select

End Sub

Private Sub SQLExpressions(Method As SQLMethod, Name As String, Optional Ordinal As Long = 1000)
On Error GoTo ErrorHappened

Dim db As DAO.Database
Dim SQL As String

Set db = CurrentDb
Select Case Method
Case SQLMethod.Add
    SQL = "Insert into SCR_ScreensTotalsCalculations(TotalID, AggregateID,FldName,Alias,Ordinal) "
    SQL = SQL & "Values (" & TotalID & "," & Me.CmboAggr & ","
    SQL = SQL & "" & Chr(34) & Name & Chr(34) & ","
    SQL = SQL & "" & Chr(34) & Name & Me.CmboAggr.Column(1, Me.CmboAggr.ListIndex) & Chr(34) & "," 'ALIAS
    SQL = SQL & Ordinal & ")"
Case SQLMethod.Delete
    SQL = "Delete SCR_ScreensTotalsCalculations.TotalCalcID "
    SQL = SQL & "FROM SCR_ScreensTotalsCalculations "
    SQL = SQL & "Where TotalID = " & TotalID & " and "
    SQL = SQL & "    TotalCalcID = " & Name 'More lazy coding
End Select

db.Execute SQL, dbFailOnError

ExitNow:
On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    Select Case Err.Number
    Case 3022 'Duplicate Index violation - Ignoring this because I am lazy about checking if the val exists before adding
        LogMessage "Duplicate Key - Item Skipped"
        Resume Next
    Case Else
        MsgBox Err.Description, vbCritical, "Custom Totals: Error Adding Field"
        LogMessage Err.Description
        Resume ExitNow
        Resume
    End Select

End Sub
Private Sub CmdGrpFldsDown_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim id1 As Long, id2 As Long
Dim Ord1 As Long, Ord2 As Long
Dim index As Long, SQL As String

index = Me.LstGrpFldsSel.ListIndex

Select Case index
Case -1, (Me.LstGrpFldsSel.ListCount - 1) 'Nothing At the bottom or no selection
    GoTo ExitNow
Case Else
    id1 = LstGrpFldsSel.Column(0, index)
    Ord1 = LstGrpFldsSel.Column(3, index)
    id2 = LstGrpFldsSel.Column(0, index + 1)
    Ord2 = LstGrpFldsSel.Column(3, index + 1)
    
'    If Ord2 = Ord1 Then
'        Ord2 = Ord1 + 1000 'Not sure why this is needed, but at times the values are getting out of sync
'    End If
End Select

SQL = "UPDATE SCR_ScreensTotalsFields "
SQL = SQL & "SET Ordinal = IIF(TotalFldID = " & id1 & "," & Ord2 & "," & Ord1 & ") "
SQL = SQL & "WHERE TotalFldID in (" & id1 & "," & id2 & ")"

Set db = CurrentDb
db.Execute SQL, dbFailOnError
DBEngine.Idle dbRefreshCache + dbForceOSFlush

ExitNow:
On Error Resume Next
    Set db = Nothing
    LstGrpFldsSel.Requery
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Ordering Fields"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdGrpFldsEdit_Click()
FieldEdit Grouping
End Sub

Private Sub CmdGrpFldsUp_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim id1 As Long, id2 As Long
Dim Ord1 As Long, Ord2 As Long
Dim index As Long, SQL As String

index = Me.LstGrpFldsSel.ListIndex

Select Case index
Case 0, -1 'Top, Nothing
    GoTo ExitNow
Case Else
    id1 = LstGrpFldsSel.Column(0, index)
    Ord1 = LstGrpFldsSel.Column(3, index)
    id2 = LstGrpFldsSel.Column(0, index - 1)
    Ord2 = LstGrpFldsSel.Column(3, index - 1)
End Select

SQL = "UPDATE SCR_ScreensTotalsFields "
SQL = SQL & "SET Ordinal = IIF(TotalFldID = " & id1 & "," & Ord2 & "," & Ord1 & ") "
SQL = SQL & "WHERE TotalFldID in (" & id1 & "," & id2 & ")"

Set db = CurrentDb
db.Execute SQL, dbFailOnError


ExitNow:
On Error Resume Next
    Set db = Nothing
    LstGrpFldsSel.Requery
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Ordering Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub CmdGrpRemove_Click()
On Error GoTo ErrorHappened

Dim X As Integer


For X = 0 To Me.LstGrpFldsSel.ItemsSelected.Count - 1
    SQLGrouping SQLMethod.Delete, LstGrpFldsSel.ItemData(LstGrpFldsSel.ItemsSelected.Item(X))
Next X

LstGrpFldsSel.Requery

If LstGrpFldsSel.ListCount > 0 Then
    LstGrpFldsSel = LstGrpFldsSel.ItemData(0)
End If

ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Custom Totals: Error Removing Field"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub

Private Sub Form_Current()
    PopulateLists

End Sub
Private Sub Form_Timer()
    Me.Messages.ForeColor = 0
    Me.TimerInterval = 0
    Me.TotalName.BackColor = 16777215
    
End Sub


Private Sub Global_Click()
Me.LblScreenName2 = "" & IIf([Global] = True, "GLB", [UserName]) & ": " & [TotalName]
End Sub

Private Sub LstExpFlds_DblClick(Cancel As Integer)
CmdExpAdd_Click
End Sub

Private Sub LstExpFldsSel_DblClick(Cancel As Integer)
CmdExpFldsEdit_Click
End Sub

Private Sub LstGrpFlds_DblClick(Cancel As Integer)
CmdGrpAdd_Click
End Sub

Private Sub LstGrpFldsSel_DblClick(Cancel As Integer)
    FieldEdit Grouping
End Sub
Private Sub PopulateLists()
On Error GoTo ErrorHappened

If Me.visible = False Then GoTo ExitNow

If "" & mvSource = "" Then
    mvSource = DLookup("PrimaryRecordSource", "SCR_Screens", "ScreenID=" & Me.ScreenID)
End If
    Me.LstGrpFlds.RowSource = mvSource
    Me.LstExpFlds.RowSource = mvSource
    Me.LstGrpFldsSel.RowSource = "SELECT TotalFldID, FldName, IIF("" & Alias = "",FldName,Alias) as AliasCalc, Ordinal FROM SCR_ScreensTotalsFields WHERE TotalID=" & TotalID & " ORDER BY Ordinal;"
    Me.LstExpFldsSel.RowSource = "SELECT TotalCalcID, AggregateName AS Aggr, FldName, Alias, Ordinal FROM SCR_ScreensTotalsCalculationsAggr  as A INNER JOIN SCR_ScreensTotalsCalculations as C ON A.AggregateID = C.AggregateID Where TotalID = " & TotalID & " ORDER BY Ordinal, Alias;"

ExitNow:
On Error Resume Next

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "PopulateLists"
    LogMessage Err.Description
    Resume ExitNow
    Resume
End Sub
Private Sub LogMessage(Message As String)
    On Error Resume Next
    Dim MsgIn() As String
    Dim MsgOut() As String
    Dim X As Integer
    
    MsgIn() = Split("" & Me.Messages, vbCrLf)
    
    
    ReDim MsgOut(UBound(MsgIn) + 1)
    
    MsgOut(0) = Format(Now(), "hh:mm:ss") & " " & Message
    
    For X = 0 To UBound(MsgIn)
        MsgOut(X + 1) = MsgIn(X)
        If X > 9 Then Exit For
    Next X
    
    Me.Messages = Join(MsgOut, vbCrLf)
    Me.Messages.ForeColor = 255
    Me.TimerInterval = 750

End Sub

Private Sub MvImport_StatusMessage(Src As String, Msg As String, lvl As Integer)
 LogMessage Src & "-" & Msg
End Sub

Private Sub TotalName_Exit(Cancel As Integer)
Me.LblScreenName2 = "" & IIf([Global] = True, "GLB", [UserName]) & ": " & [TotalName]
End Sub
