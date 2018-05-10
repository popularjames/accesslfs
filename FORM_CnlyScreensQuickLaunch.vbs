Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private MvScreenID As Long


Private Sub CmdBuildToolbar_Click()

    Call CCACommandBarMake

End Sub


Private Sub cmdDelete_Click()
    On Error GoTo ErrorHappened

    Dim stName As String
    Dim lngId As Long

    If Me.Lst.ListIndex = -1 Then
        MsgBox "You must select a PowerBar to delete.", vbInformation, "Delete PowerBar"
        GoTo ExitNow
    End If

    stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)

    lngId = Me.Lst.Value

    If MsgBox("Delete Menu Form '" & stName & "'?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
        CurrentDb.Execute ("Delete * From CnlyScreensQuickLaunch Where FormID = " & CStr(lngId))
        Me.Lst.Requery
        Me.Requery
    End If

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim stName As String


    stName = Nz(InputBox("Enter New Menu Form List Name", "New Menu Form"), "")

    If stName <> "" Then

        Set db = CurrentDb
        db.Execute "Insert Into CnlyScreensQuickLaunch( ListName, FormName) Values(" & Chr(34) & stName & Chr(34) & "," & Chr(34) & stName & Chr(34) & ")", dbFailOnError
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
    Dim stName As String
    Dim lngId As Long
    Dim StNewName As String

    If Me.Lst.ListIndex = -1 Then
        MsgBox "You must select a Menu Form to rename.", vbInformation, "Rename Menu Form"
        GoTo ExitNow
    End If

    stName = Me.Lst.Column(1, Me.Lst.ListIndex + 1)
    lngId = Me.Lst.Value

    StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename Menu Form", stName)

    If "" & StNewName = "" Or "" & StNewName = stName Then
        GoTo ExitNow
    End If

    CurrentDb.Execute ("Update CnlyScreensQuickLaunch Set ListName = " & Chr(34) & StNewName & Chr(34) & " Where FormID = " & CStr(lngId))
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

Private Sub Lst_AfterUpdate()

    Dim rst As DAO.RecordSet

    Set rst = Me.RecordsetClone

    With rst
        .MoveFirst
        .FindFirst "FormID =" & Me.Lst.Value
        If .NoMatch = False Then
            Me.Bookmark = .Bookmark
        End If
    End With

exitHere:
    Set rst = Nothing
End Sub
