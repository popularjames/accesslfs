Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 11/14/2012 - Changed RowSource property on CmboRecordSource and added row to CmboRecordSourceType

Private MvScreenID As Long
Private MvSortMode As Boolean


Private Sub CmboRecordSource_AfterUpdate()
On Error GoTo ErrorHappened
    'Update the object type combo
    Select Case CmboRecordSource.Column(1, CmboRecordSource.ListIndex)
        Case "Table"
            CmboRecordSourceType = 0
        Case "Query"
            CmboRecordSourceType = 1
        Case "Form"
            CmboRecordSourceType = 2
    End Select
    
    'Change The Tabbed SubformFields List Sources
    ChangeRecordSource
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub CmdAddBatch_Click()
On Error GoTo ErrorHappened
Dim frm As New Form_CT_PopupSelect
Dim SQL As String
With frm
    With .Lst
        SQL = CmboRecordSource.RowSource
        .RowSource = SQL
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "3.5" & Chr(34) & ";.5" & Chr(34)
        .Requery
    End With
    .Title = "Batch Add Tabs"
    .ListTitle = "Select Record Source(s)"
    .StartupWidth = -1  'AUTO
    .visible = True
    
    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        Dim X As Long, MyCol As Collection
        SQL = ""
        Set MyCol = .Selections
        DoCmd.Hourglass True
        If Not MyCol Is Nothing Then
            For X = 1 To MyCol.Count
                SQL = "Insert Into SCR_ScreensTabs(ScreenID,Feature, RecordSource,Type) Values("
                SQL = SQL & Me.ScreenIDCurrent & ","
                SQL = SQL & Chr(34) & MyCol.Item(X)(0) & Chr(34) & ","
                SQL = SQL & Chr(34) & MyCol.Item(X)(0) & Chr(34) & ","
                SQL = SQL & IIf(MyCol.Item(X)(1) = "Table", 0, 1) & ")"
                CurrentDb.Execute SQL, dbFailOnError
NextITEM:
            Next X
        End If
    End If
End With
ExitNow:
    On Error Resume Next
    Lst.Requery
    Me.Requery
    DoCmd.Hourglass False
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    Select Case Err.Number
    Case 3022 ' Duplicate Index
        If MsgBox("Failed Adding Tab (" & MyCol.Item(X)(0) & ") because it already exists." & vbCrLf & vbCrLf & "Would you like to continue adding tabs?", vbQuestion + vbDefaultButton1 + vbYesNo, "Add Failed --> " & CodeContextObject.Name) = vbYes Then
            Resume NextITEM
        Else
            Resume ExitNow
        End If
    Case Else
        MsgBox Err.Description, vbCritical, "Error Adding Tabs --> " & CodeContextObject.Name
        Resume ExitNow
    End Select
    Resume
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim StTabName As String, lngTabID As Long

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a tab to delete.", vbInformation, "Delete Tab"
    GoTo ExitNow
End If

StTabName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngTabID = Me.Lst.Value
If MsgBox("Delete Tab '" & StTabName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete From SCR_ScreensTabs Where TabID = " & CStr(lngTabID))
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

Private Sub CmdDown_Click()
SortChange -1
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim StTabName As String


StTabName = InputBox("Enter New Tab Name", "New Tab")
If "" & StTabName <> "" Then
    Set db = CurrentDb
    db.Execute "Insert Into SCR_ScreensTabs(ScreenID, Feature, RecordSource) Values(" & MvScreenID & ", " & Chr(34) & StTabName & Chr(34) & ",'NA')", dbFailOnError
    Me.Lst.Requery
    Me.Requery
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub cmdRefresh_Click()
    Me.Lst.Requery
End Sub

Private Sub CmdRename_Click()
On Error GoTo ErrorHappened
Dim StTabName As String, lngTabID As Long
Dim StNewName As String

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a tab to rename.", vbInformation, "Rename Tab"
    GoTo ExitNow
End If

StTabName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngTabID = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & StTabName & "'", "Rename Tab", StTabName)
If "" & StNewName = "" Or "" & StNewName = StTabName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensTabs Set Feature = " & Chr(34) & StNewName & Chr(34) & " Where TabID = " & CStr(lngTabID))
Me.Lst.Requery
Me.Requery
Me.Lst.Value = lngTabID
Lst_AfterUpdate

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub CmdSort_AfterUpdate()
    Me.SortMode = CmdSort.Value
End Sub



Private Sub CmdUp_Click()
SortChange 1
End Sub
Public Sub SortChange(Move As Integer)
On Error GoTo ErrorHappened
Dim LgID As Long, LgIDx As Long
Dim db As DAO.Database
Dim rst As DAO.RecordSet
Dim X As Long

If Me.Lst.ListIndex <> -1 Then
    LgID = Lst.Value
    LgIDx = Lst.ListIndex
   
    If LgIDx = 0 And Move = 1 Then 'IF FIRST ITEM AND MOVE UP THEN NOTHING TO DO
        GoTo ExitNow
    End If
    If LgIDx = Lst.ListCount - 1 And Move = -1 Then 'IF LAST ITEM AND MOVE DOWN THEN NOTHING TO DO
        GoTo ExitNow
    End If
    
    ReDim ArySort(Lst.ListCount - 1)
    For X = 0 To Lst.ListCount - 1
        If X = LgIDx Then ' THE ITEM TO MOVE
            ArySort(X) = X - Move
        ElseIf Move = 1 And X = LgIDx - 1 Then 'THE ITEM BEFORE
            ArySort(X) = X + Move
        ElseIf Move = -1 And X = LgIDx + 1 Then 'THE ITEM AFTER
            ArySort(X) = X + Move
        Else
            ArySort(X) = X
        End If
    Next X
    Set db = CurrentDb
    Set rst = Lst.RecordSet
    X = 0
    With rst
        .MoveFirst
        Do Until .EOF
            If rst("Sort") <> ArySort(X) Then
                .Edit
                rst("Sort") = ArySort(X)
                .Update
            End If
            .MoveNext
            X = X + 1
        Loop
        
    End With
    Lst.Requery
    Me.Requery
    Lst.Value = LgID
    Lst_AfterUpdate
End If

ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    DoCmd.Hourglass False
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub
Private Sub Form_Current()
On Error Resume Next

'GET THE CURRENT SCREEN ID PROPERTIES
If Nz(Me!ScreenID, 0) = 0 Then
    If Me.Parent.Name = "SCRCFG_Screens" Then
        Me.ScreenIDCurrent = Nz(Me.Parent!ScreenID, 0)
    Else
        Me.ScreenIDCurrent = 0
    End If
Else
   Me.ScreenIDCurrent = Nz(Me!ScreenID, 0)
End If
'Change The Tabbed SubformFields List Sources
ChangeRecordSource
End Sub

Public Property Let ScreenIDCurrent(data As Long)
    If MvScreenID <> data Then
        MvScreenID = data
        Lst.RowSource = "SELECT TabID, Feature as TabName, [Sort] FROM SCR_ScreensTabs WHERE ScreenID =" & MvScreenID & " ORDER BY Sort, Feature"
    End If
End Property

Public Property Get ScreenIDCurrent() As Long
 ScreenIDCurrent = MvScreenID
End Property

Public Property Get SortMode() As Boolean
    SortMode = MvSortMode
End Property

Public Property Let SortMode(data As Boolean)
DoCmd.Hourglass True
    MvSortMode = data
    
    If MvSortMode = True Then
      Me.Lst.BackColor = 10545662 'RGB(245, 245, 245)
    Else
      Me.Lst.BackColor = RGB(255, 255, 255)
    End If
    
    Dim MyCtrl As Control
    For Each MyCtrl In Me.Controls
        'Debug.Print MyCtrl.Name
        Select Case UCase(MyCtrl.Name)
        Case "CMDUP", "CMDDOWN"
            MyCtrl.Enabled = MvSortMode
        Case "CMDSORT", "LST"
            'DO NOTHING - MUST REMAIL ENABLED
        Case "LBLFEATURE1", "LBLFEATURE2"
            'DO NOTHING - MUST REMAIL DISABLED
        Case Else
            If MyCtrl.ControlType <> acLabel Then
                MyCtrl.Enabled = CBool(Abs(MvSortMode) - 1)
            End If
        End Select
    Next MyCtrl
    Set MyCtrl = Nothing
DoCmd.Hourglass False
'
End Property
Private Sub Lst_AfterUpdate()
If Me.SortMode = False Then
    Dim rst As DAO.RecordSet
    Set rst = Me.RecordsetClone
    With rst
        .MoveFirst
        .FindFirst "TabID =" & Me.Lst.Value
        If .NoMatch = False Then
            Me.Bookmark = .Bookmark
        End If
    End With
End If
End Sub

Sub ChangeRecordSource()
On Error GoTo ChangeRecordSourceError
Dim tmpStr As String

'Get The Table Name From The Current Screens Primary Definition
tmpStr = "" & DLookup("PrimaryRecordSource", "SCR_Screens", "ScreenID = " & MvScreenID)

With Me.SfTabsFields
    .Form!MasterField.RowSource = tmpStr
    .Form!ChildField.RowSource = "" & Me.CmboRecordSource
End With

ChangeRecordSourceExit:
    On Error Resume Next
    Exit Sub
    
ChangeRecordSourceError:
    Select Case Err.Number
    Case 13 'Type MisMatch
        Resume ChangeRecordSourceExit
    Case Else
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error getting field list!", vbInformation
        Resume ChangeRecordSourceExit
        Resume
    End Select
End Sub
