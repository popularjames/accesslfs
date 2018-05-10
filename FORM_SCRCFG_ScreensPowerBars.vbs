Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long

'SA 11/14/2012 - Updated suggested Powerbar size on the label


Private Sub cmdDelete_Click()
On Error GoTo ErrorHappened
Dim stName As String, lngId As Long

If Me.Lst.ListIndex = -1 Then
    MsgBox "You must select a PowerBar to delete.", vbInformation, "Delete PowerBar"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value
If MsgBox("Delete PowerBar '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    CurrentDb.Execute ("Delete * From SCR_ScreensPowerBars Where PwrBarID = " & CStr(lngId))
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
    db.Execute "Insert Into SCR_ScreensPowerBars(ScreenID, ListName, Function) Values(" & MvScreenID & ", " & Chr(34) & stName & Chr(34) & "," & Chr(34) & stName & Chr(34) & ")", dbFailOnError
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
    MsgBox "You must select a PowerBar to rename.", vbInformation, "Rename PowerBar"
    GoTo ExitNow
End If

stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
lngId = Me.Lst.Value

StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename PowerBar", stName)
If "" & StNewName = "" Or "" & StNewName = stName Then
    GoTo ExitNow
End If

CurrentDb.Execute ("Update SCR_ScreensPowerBars Set ListName = " & Chr(34) & StNewName & Chr(34) & " Where PwrBarID = " & CStr(lngId))
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
End Sub

Public Property Let ScreenIDCurrent(data As Long)
    If MvScreenID <> data Then
        MvScreenID = data
        Lst.RowSource = "SELECT PwrBarID, ListName FROM SCR_ScreensPowerBars WHERE ScreenID =" & MvScreenID & " ORDER BY ListName"
    End If
End Property

Public Property Get ScreenIDCurrent() As Long
 ScreenIDCurrent = MvScreenID
End Property


Private Sub Lst_AfterUpdate()
Dim rst As DAO.RecordSet
Set rst = Me.RecordsetClone
With rst
    .MoveFirst
    .FindFirst "PwrBarID =" & Me.Lst.Value
    If .NoMatch = False Then
        Me.Bookmark = .Bookmark
    End If
End With
End Sub
