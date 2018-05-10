Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long
Private mvFormId As Byte
Private MvScr As Form_SCR_MainScreens

Public Property Get ActiveCalcCount() As Long
    ActiveCalcCount = LstApply.ListCount
End Property

Public Property Get SQLList() As String
    Dim X As Long, StIDs As String
    For X = 0 To LstApply.ListCount - 1
        If X <> LstApply.ListCount - 1 Then
            StIDs = StIDs & LstApply.Column(0, X) & ","
        Else
            StIDs = StIDs & LstApply.Column(0, X)
        End If
    Next X
    SQLList = StIDs
End Property

Private Sub CmdAdd_Click()
If Lst.ListIndex <> -1 Then
    ApplyAdd Lst.Column(0, Lst.ListIndex), Lst.Column(1, Lst.ListIndex)
End If
End Sub

Private Sub cmdApply_Click()
If Me.Dirty = True Then Me.Dirty = False
    ApplyCalcs
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a Calculation to delete.", vbInformation, "Delete Calculation"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex)
lngId = Me.Lst.Column(0, Me.Lst.ListIndex)
If MsgBox("Delete Calculation '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete * From SCR_ScreensCalculations Where CalcID = " & CStr(lngId))
    Me.Lst.Requery
    DBEngine.Idle dbRefreshCache + dbForceOSFlush
    Me.Lst.Requery
    Me.Requery
    
End If



ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub





Private Sub CmdApplyDelete_Click()
    ApplyClear
End Sub

Private Sub cmdRemove_Click()
If LstApply.ListIndex <> -1 Then
    ApplyRemove LstApply.Column(0, LstApply.ListIndex)
End If
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim stName As String
Dim SQL As String
Dim NewID As Long



stName = InputBox("Enter Name For New Calculation", "New Calculation")
If "" & stName <> "" Then
    SQL = "Insert Into SCR_ScreensCalculations(ScreenID, CalcName, CalcFormula,Computer,UserName) "
    SQL = SQL & "Values(" & MvScreenID & ", " & Chr(34) & stName & Chr(34) & ",'NA', "
    SQL = SQL & Chr(34) & Identity.Computer & Chr(34) & ", " & Chr(34) & Identity.UserName & Chr(34) & ")"
    Set db = CurrentDb
    db.Execute SQL, dbFailOnError
    NewID = DLookup("CalcID", "SCR_ScreensCalculations", "ScreenID =" & MvScreenID & " and CalcName = " & Chr(34) & stName & Chr(34))
    Me.Lst.Requery
    If Me.RecordSet.BOF And Me.RecordSet.EOF Then
        Me.Undo
    End If
    Me.Requery
    Lst.Value = NewID
    Lst_AfterUpdate
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub cmdRefresh_Click()
Lst.Requery
End Sub


Private Sub CmdRename_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long
Dim StNewName As String

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a calculation to rename.", vbInformation, "Rename Calculation"
    GoTo ExitNow
End If

stName = "" & Me.Lst.Column(1, Me.Lst.ListIndex)
lngId = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename Calculation", stName)
If "" & StNewName = "" Or "" & StNewName = stName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensCalculations Set CalcName = " & Chr(34) & StNewName & Chr(34) & " Where CalcID = " & CStr(lngId))
Me.Lst.Requery
Me.Requery
Me.Lst.Value = lngId
Lst_AfterUpdate

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub CmdZoom1_Click()
ZoomText Me.Expression1, "Calculation"
End Sub

Private Sub Form_Current()
Dim BlEnabled As Boolean

If "" & Me.RecordSource = "" Then
    Exit Sub
End If

If Nz(Me!CalcID, 0) = 0 Then
    BlEnabled = False
ElseIf (Me.RecordSet.EOF Or Me.RecordSet.BOF) Then
    BlEnabled = ((Me.RecordSet.EOF Or Me.RecordSet.BOF) * -1) - 1
Else
    BlEnabled = True
End If

Me.Expression1.Enabled = BlEnabled
Me.CmdZoom1.Enabled = BlEnabled
Me.DataType.Enabled = BlEnabled
Me.FieldWidth.Enabled = BlEnabled
Me.Align.Enabled = BlEnabled
Me.Format.Enabled = BlEnabled
Me.Decimals.Enabled = BlEnabled

End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set MvScr = Nothing
End Sub

Private Sub Lst_AfterUpdate()
On Error GoTo ErrorHappened
Dim rst As DAO.RecordSet

If Lst.ListIndex = -1 Then
    GoTo ExitNow
End If

Set rst = Me.RecordsetClone

With rst
    .MoveFirst
    .FindFirst "CalcID = " & Lst.Column(0, Lst.ListIndex)
    If .NoMatch = False Then
        Me.Bookmark = .Bookmark
    End If
End With
ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error Selecting Record", vbCritical, "Calculations"
    Resume ExitNow
    Resume
End Sub


Public Sub InitData(ScreenID As Long, FormID As Byte)
On Error GoTo LoadError


If Me.Parent Is Nothing Then
    MsgBox "This form is not meant to run independatly of a screen!", vbCritical
    DoCmd.Close acForm, Me.Name, acSaveNo
    GoTo ExitNow
Else
    MvScreenID = ScreenID
    mvFormId = FormID
    Lst.RowSource = "SELECT CalcID, CalcName FROM SCR_ScreensCalculations Where ScreenID = " & MvScreenID & " ORDER BY CalcName "
    Set MvScr = Scr(mvFormId)
    Me.RecordSource = "SELECT * FROM SCR_ScreensCalculations Where ScreenID = " & MvScreenID
    Me.Requery
End If

ExitNow:
    Exit Sub
LoadError:
    MsgBox Err.Description, vbCritical, "Error Loading Conditional Formats"
    'DoCmd.Close acForm, Me.Name, acSaveNo
    Resume ExitNow
    Resume
End Sub
Public Sub ApplyAdd(ID As Long, Name As String)
Dim X As Long
For X = 0 To LstApply.ListCount - 1
    If LstApply.Column(0, X) = ID Then
        Exit Sub
    End If
Next X
LstApply.AddItem (ID & ";" & Chr(34) & Name & Chr(34))
End Sub
Public Sub ApplyRemove(ID As Long)
Dim X As Long
LstApply.Value = Null
For X = 0 To LstApply.ListCount - 1
    If LstApply.Column(0, X) = ID Then
        LstApply.RemoveItem (X)
        Exit For
    End If
Next X
LstApply.Requery
DoEvents
End Sub
Public Sub ApplyClear()
Dim X As Long

If LstApply.ListCount > 0 Then
    For X = LstApply.ListCount - 1 To 0 Step -1
        LstApply.RemoveItem (X)
    Next X
End If
End Sub
Private Sub Lst_DblClick(Cancel As Integer)
If Lst.ListIndex <> -1 Then
    ApplyAdd Lst.Column(0, Lst.ListIndex), Lst.Column(1, Lst.ListIndex)
End If
End Sub

Private Sub LstApply_DblClick(Cancel As Integer)
If LstApply.ListIndex <> -1 Then
    LstApply.RemoveItem LstApply.ListIndex
    LstApply.Requery
    DoEvents
End If
End Sub
Public Sub ApplyCalcs()
Dim X As Long, StIDs As String
Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
Dim fld As CnlyFldDef
DoCmd.Hourglass True

'SAVE CHANGES TO CURRENT RECORD IF NEEDED
If Me.Dirty = True Then
    Me.Dirty = False
End If

'FIRST REMOVE ALL OF THE EXISTING FORMATS
If MvScr.GridForm Is Nothing Then
    MsgBox "HACK"
    GoTo ExitNow
Else
MvScr.GridForm.CalcFieldsClear
End If
'NOW FIGURE OUT WHAT NEEDS TO BE APPLIED
With LstApply
    If .ListCount > 0 Then
        For X = 0 To LstApply.ListCount - 1
            If X <> .ListCount - 1 Then
                StIDs = StIDs & LstApply.Column(0, X) & ","
            Else
                StIDs = StIDs & LstApply.Column(0, X)
            End If
        Next X
    Else
        GoTo ExitNow
    End If
End With

'GET THE VALUES FROM THE DATABASE
SQL = "Select * From SCR_ScreensCalculations Where CalcID in (" & StIDs & ")"
DBEngine.Idle dbRefreshCache + dbForceOSFlush
DoEvents
Set db = CurrentDb
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
With rst
    Do Until .EOF
        'Get The Field Def
        fld.Name = .Fields("CalcName")
        fld.ControlSrc = .Fields("CalcFormula")
        With fld
            .Type = rst.Fields("DataType")
            .Decimal = Nz(rst.Fields("Decimals"), 255)
            .Alias = .Name
            .left = 0
            If IsNull(rst.Fields("FieldWidth")) Then
                .Width = GetFieldWidth(.Type, IIf(.Type = 10, 10, 0), "SCR_ScreensCalculations", .Name, Me.Parent.ScreenID)
            Else
                .Width = rst.Fields("FieldWidth")
            End If
            .Height = GetFieldHeight(Identity.DataSheetStyle.fontsize)
            If IsNull(rst.Fields("Align")) Then
                .Align = GetFieldAlign(.Type, "SCR_ScreensCalculations", .Name, Me.Parent.ScreenID)
            Else
                .Align = rst.Fields("Align")
            End If
            If "" & rst.Fields("Format") <> "" Then
                .Format = rst.Fields("Format")
            Else
                .Format = GetFieldFormat(.Type, .Decimal, "SCR_ScreensCalculations", .Name, Me.Parent.ScreenID)
            End If
        End With
        MvScr.GridForm.CalcFieldsAdd fld
        MvScr.GridForm.Recalc
NextRecord:
        .MoveNext
    Loop
    .Close
End With

ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Set rst = Nothing
    Set db = Nothing

    Exit Sub
End Sub

Private Sub ZoomText(ctl As TextBox, Optional ByVal Title As String = "")
Dim FrmZoom As Form_CT_Text
Dim Txt As String

Txt = "" & ctl

Set FrmZoom = New Form_CT_Text
FrmZoom.Move ctl.left
With FrmZoom
    .Text = Txt
    .visible = True
    .Title = Title
    Do Until .Results <= 0
        DoEvents
    Loop
    If .Results = True Then
        ctl = "" & .Text
    End If
End With

End Sub
Function GiveMeDaText(StInput As String, delim As String) As String
Dim ST() As String

ST = Split(StInput, delim)
    
GiveMeDaText = ST(1)

End Function
