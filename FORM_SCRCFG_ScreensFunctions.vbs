Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long
Private MvSortMode As Boolean



Private Sub CmdAddSystem_Click()
On Error GoTo ErrorHappened
Dim frm As New Form_CT_PopupSelect
Dim SQL As String

With frm
    With .Lst
        SQL = "SELECT FunctionID, Function "
        SQL = SQL & "FROM SCR_ScreensMastersFunctions "
        SQL = SQL & "Where Not Function In (Select Function From SCR_ScreensFunctions Where ScreenID = " & Me.ScreenIDCurrent & ") "
        SQL = SQL & "ORDER BY Function"
        .RowSource = SQL
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0;10" & Chr(34)
        .Requery
    End With
    .Title = "Add Functions"
    .ListTitle = "Select Function(s)"
    .visible = True
    
    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        Dim X As Long, MyCol As Collection
        Dim LgID As Long
        DoCmd.Hourglass True
        Set MyCol = .Selections
        If Not MyCol Is Nothing Then
            For X = 1 To MyCol.Count
                SQL = ""
                SQL = SQL & MyCol.Item(X)(0)
                SQL = "Insert Into SCR_ScreensFunctions(ScreenID, ListName, Function, System) "
                SQL = SQL & "Values(" & Me.ScreenIDCurrent & ", " & Chr(34) & MyCol.Item(X)(1) & Chr(34)
                SQL = SQL & ", " & Chr(34) & MyCol.Item(X)(1) & Chr(34) & ", True)"
                CurrentDb.Execute SQL, dbFailOnError
                LgID = DMax("FunctionID", "SCR_ScreensFunctions", "ScreenID=" & Me.ScreenIDCurrent & " and Function = " & Chr(34) & MyCol.Item(X)(1) & Chr(34))
                SQL = "Insert Into SCR_ScreensFunctionsFields(FunctionID,FieldDef,FieldName,FieldType) "
                SQL = SQL & "Select " & LgID & " as FunctionID,FieldDef,FieldName,FieldType From SCR_ScreensMastersFunctionsFields "
                SQL = SQL & "Where FunctionID = " & MyCol.Item(X)(0)
                CurrentDb.Execute SQL, dbFailOnError
NextITEM:
            Next X
        End If
    End If
End With

ExitNow:
    On Error Resume Next
    Me.Requery
    Me.Lst.Requery
    DoCmd.Hourglass False
    Set MyCol = Nothing
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    Select Case Err.Number
    Case 3022 ' Duplicate Index
        If MsgBox("Failed Adding Function (" & MyCol.Item(X)(1) & ") because it already exists." & vbCrLf & vbCrLf & "Would you like to continue adding functions?", vbQuestion + vbDefaultButton1 + vbYesNo, "Add Failed --> " & CodeContextObject.Name) = vbYes Then
            Resume NextITEM
        Else
            Resume ExitNow
        End If
    Case Else
        MsgBox Err.Description, vbCritical, "Error Adding Functions --> " & CodeContextObject.Name
        Resume ExitNow
    End Select
    Resume


End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a function to delete.", vbInformation, "Delete Function"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value
If MsgBox("Delete Function '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete * From SCR_ScreensFunctions Where FunctionID = " & CStr(lngId))
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

Private Sub cmdNew_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim stName As String


stName = InputBox("Enter New Function Name", "New Function")
If "" & stName <> "" Then
    Set db = CurrentDb
    db.Execute "Insert Into SCR_ScreensFunctions(ScreenID, ListName, Function) Values(" & MvScreenID & ", " & Chr(34) & stName & Chr(34) & "," & Chr(34) & stName & Chr(34) & ")", dbFailOnError
    Me.Lst.Requery
    Me.Requery
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, CodeContextObject.Name & ": " & Err.Source
    Resume ExitNow
End Sub

Private Sub cmdRefresh_Click()
    Me.Lst.Requery
End Sub

Private Sub CmdRename_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long
Dim StNewName As String

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a function to rename.", vbInformation, "Rename Function"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename Function", stName)
If "" & StNewName = "" Or "" & StNewName = stName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensFunctions Set ListName = " & Chr(34) & StNewName & Chr(34) & " Where FunctionID = " & CStr(lngId))
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
If Nz(Me!System, False) = True Then
    Me.CmboFunction.Locked = True
    CmboFunction.BackColor = RGB(230, 230, 230)
Else
    Me.CmboFunction.Locked = False
    CmboFunction.BackColor = RGB(255, 255, 255)
End If
'Change The Tabbed SubformFields List Sources
ChangeRecordSource
End Sub

Public Property Let ScreenIDCurrent(data As Long)
    If MvScreenID <> data Then
        MvScreenID = data
        Lst.RowSource = "SELECT FunctionID, ListName FROM SCR_ScreensFunctions WHERE ScreenID =" & MvScreenID & " ORDER BY ListName"
    End If
End Property

Public Property Get ScreenIDCurrent() As Long
 ScreenIDCurrent = MvScreenID
End Property

Public Property Get SortMode() As Boolean
    SortMode = MvSortMode
End Property

Private Sub Lst_AfterUpdate()
Dim rst As DAO.RecordSet
Set rst = Me.RecordsetClone
With rst
    .MoveFirst
    .FindFirst "FunctionID =" & Me.Lst.Value
    If .NoMatch = False Then
        Me.Bookmark = .Bookmark
    End If
End With
End Sub

Sub ChangeRecordSource()
On Error GoTo ChangeRecordSourceError
Dim tmpStr As String

'Get The Table Name From The Current Screens Primary Definition
tmpStr = "" & DLookup("PrimaryRecordSource", "SCR_Screens", "ScreenID = " & MvScreenID)

With Me.SfFunctionFields
    .Form!FieldName.RowSource = tmpStr
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
