Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long
Private MvSortMode As Boolean


Private Sub CmboRecordSource_AfterUpdate()
Dim MyCtrl As ComboBox
Set MyCtrl = Me.Controls.Item("CmboRecordSource")
    Me.Controls.Item(MyCtrl.Name & "Type") = IIf(MyCtrl.Column(1, MyCtrl.ListIndex) = "Table", 0, 1)
Set MyCtrl = Nothing

'Change The Tabbed SubformFields List Sources
ChangeRecordSource

End Sub



Private Sub CmdAddReports_Click()
On Error GoTo ErrorHappened
Dim frm As New Form_CT_PopupSelect
Dim SQL As String
With frm
    With .Lst
        SQL = "SELECT O2.Name "
        SQL = SQL & "FROM MSysObjects AS O1 "
        SQL = SQL & "INNER JOIN MSysObjects AS O2 ON "
        SQL = SQL & "O1.Id = O2.ParentId "
        SQL = SQL & "WHERE O1.Name = " & Chr(34) & "Reports" & Chr(34) & " AND "
        SQL = SQL & "NOT O2.Name in (Select ListName From SCR_ScreensReports Where ScreenID = " & Me.ScreenIDCurrent & ") "
        SQL = SQL & "ORDER BY O2.Name;"
        .RowSource = SQL
        .BoundColumn = 1
        .ColumnCount = 1
        .ColumnWidths = "3" & Chr(34)
        .Requery
    End With
    .Title = "Batch Add Reports"
    .ListTitle = "Select Report(s)"
    .StartupWidth = (3 * 1440)
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
                SQL = "Insert Into SCR_ScreensReports(ScreenID,ListName,ReportName) Values("
                SQL = SQL & Chr(34) & Me.ScreenIDCurrent & Chr(34) & ","
                SQL = SQL & Chr(34) & MyCol.Item(X)(0) & Chr(34) & ","
                SQL = SQL & Chr(34) & MyCol.Item(X)(0) & Chr(34) & ")"
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
        If MsgBox("Failed Adding Report (" & MyCol.Item(X)(0) & ") because it already exists." & vbCrLf & vbCrLf & "Would you like to continue adding reports?", vbQuestion + vbDefaultButton1 + vbYesNo, "Add Failed --> " & CodeContextObject.Name) = vbYes Then
            Resume NextITEM
        Else
            Resume ExitNow
        End If
    Case Else
        MsgBox Err.Description, vbCritical, "Error Adding Reports --> " & CodeContextObject.Name
        Resume ExitNow
    End Select
    Resume
    

End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a report to delete.", vbInformation, "Delete Report"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value
If MsgBox("Delete Report '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete * From SCR_ScreensReports Where ReportID = " & CStr(lngId))
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


stName = InputBox("Enter New Report Name", "New Report")
If "" & stName <> "" Then
    Set db = CurrentDb
    db.Execute "Insert Into SCR_ScreensReports(ScreenID, ListName, ReportName) Values(" & MvScreenID & ", " & Chr(34) & stName & Chr(34) & ",'NA')", dbFailOnError
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
Dim stName As String, lngId As Long
Dim StNewName As String

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a report to rename.", vbInformation, "Rename Report"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename Report", stName)
If "" & StNewName = "" Or "" & StNewName = stName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensReports Set ListName = " & Chr(34) & StNewName & Chr(34) & " Where ReportID = " & CStr(lngId))
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
'Change The Tabbed SubformFields List Sources
ChangeRecordSource
End Sub

Public Property Let ScreenIDCurrent(data As Long)
    If MvScreenID <> data Then
        MvScreenID = data
        Lst.RowSource = "SELECT ReportID, ListName FROM SCR_ScreensReports WHERE ScreenID =" & MvScreenID & " ORDER BY ListName"
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
    .FindFirst "ReportID =" & Me.Lst.Value
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

With Me.SfReportFields
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
