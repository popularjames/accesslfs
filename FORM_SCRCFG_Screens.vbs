Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 03/22/2012 - CR2609 Moved and resized various screen objects
'SA 10/10/2012 - Added tab for user table management dialog

Private MvSortMode As Boolean
Private genUtils As New CT_ClsGeneralUtilities

Private Sub SetList2Visible(BL As Boolean)
    With Me
        .CkMulti2.visible = BL
        .LblTxtListCaption2.visible = BL
        .TxtListCaption2.visible = BL
        .LblCmboListSource2.visible = BL
        .CmboListSource2.visible = BL
        .CmboListSource2Type.visible = BL
        .SfCnlyScreensLists2.visible = BL
    End With
End Sub
Private Sub SetList3Visible(BL As Boolean)
    With Me
        .ckMulti3.visible = BL
        .LblTxtListCaption3.visible = BL
        .TxtListCaption3.visible = BL
        .LblCmboListSource3.visible = BL
        .CmboListSource3.visible = BL
        .CmboListSource3Type.visible = BL
        .SfCnlyScreensLists3.visible = BL
    End With
End Sub
Private Sub SetDateVisible(BL As Boolean)
    Me.SfDates.visible = BL
End Sub

Private Sub ChkList2Use_AfterUpdate()
    SetList2Visible Me.ChkList2Use
End Sub

Private Sub ChkList2Use_Click()
    SetList2Visible Me.ChkList2Use
End Sub

Private Sub ChkList3Use_AfterUpdate()
    SetList3Visible Me.ChkList3Use
End Sub

Private Sub ChkList3Use_Click()
    SetList3Visible Me.ChkList3Use
End Sub

Private Sub ckDateUse_AfterUpdate()
    SetDateVisible Me.ckDateUse
End Sub

Private Sub ckDateUse_Click()
    SetDateVisible Me.ckDateUse
End Sub

Private Sub CmboListSource1_AfterUpdate()
With Me.SfCnlyScreensLists1.Form!FieldName
    .RowSource = CmboListSource1
    .Requery
End With
' HC 9/22 - removed custom criteria list box record source from the screen.  If it is set to none, it will be
' set to the value of the primary.  If a different record source is required, the DA will need to set it by
' hand in the table
If UCase(Me!CustomCriteriaListBoxRecordSource.Value) = "NONE" Then
    On Error Resume Next
    CurrentDb.Execute ("UPDATE SCR_Screens SET CustomCriterialListBoxRecordSource = '" & CmboListSource1.Value & "' Where ScreenID = " & CStr(Me!ScreenID))
End If
End Sub

Private Sub CmboListSource2_AfterUpdate()
With Me.SfCnlyScreensLists2.Form!FieldName
    .RowSource = CmboListSource2
    .Requery
End With
End Sub

Private Sub CmboListSource3_AfterUpdate()
With Me.SfCnlyScreensLists3.Form!FieldName
    .RowSource = CmboListSource3
    .Requery
End With

End Sub

Private Sub CmboPrimary_AfterUpdate()
    RecordsourceUpdated "CmboPrimary"
With Me.SfDates.Form!FieldName
    .RowSource = "" & CmboPrimary
    .Requery
End With
End Sub

Private Sub CmboTotals_AfterUpdate()
    RecordsourceUpdated "CmboTotals"
End Sub

Private Sub cmdAddNote_Click()
On Error GoTo ErrorHappened
Dim SQL As String
If Me!ScreenID = 0 Or "" & Me.TxtNote = "" Then
    MsgBox "No Current Screen or Blank Note", vbCritical
    Exit Sub
End If

SQL = "Insert Into SCR_ScreensNotes (ScreenID,NoteText,Computer, UserName) "
SQL = SQL & "Values ("
SQL = SQL & Me!ScreenID & ", "
SQL = SQL & Chr(34) & Me.TxtNote & Chr(34) & ", "
SQL = SQL & Chr(34) & Identity.Computer & Chr(34) & ", "
SQL = SQL & Chr(34) & Identity.UserName & Chr(34) & ") "


CurrentDb.Execute SQL, dbFailOnError

SfNotes.Form.Requery

ExitNow:
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Adding Note"
    Resume ExitNow
End Sub


Private Sub CmdDatesApply_Click()
On Error GoTo ErrorHappened
'JL Added validation for dates and updated .RowSource select statement to select all screens
If Not IsDate(Me.StartDte) Then
    MsgBox "The start date has an invalid value", vbInformation, "Screens Config"
    GoTo ExitNow
End If

If Not IsDate(Me.EndDte) Then
    MsgBox "The End Date has an invalid value", vbInformation, "Screens Config"
    GoTo ExitNow
End If

Dim frm As New Form_CT_PopupSelect
'get selected row
Dim ListItem As String
Dim currRow As Variant
For Each currRow In Me.LstScreens.ItemsSelected
    If Me.LstScreens.Selected(currRow) Then
        ListItem = Me.LstScreens.ItemData(currRow)
    End If
Next currRow



With frm
    With .Lst
        .RowSource = "Select ScreenID, ScreenName From SCR_Screens Order By ScreenName "
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0;3"
        .Requery
    End With

    Dim xx As Integer
    For xx = 0 To .Lst.ListCount - 1
        If .Lst.Column(0, xx) = ListItem Then
            .Lst.Selected(xx) = True
            Exit For
        End If
    Next xx

    .Title = "Select Screens"
    .ListTitle = "Apply Dates to:"
    .StartupWidth = -1   'AUTO SIZE THE FORM TO LIST WIDTH
    .visible = True
    
    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        Dim X As Long, MyCol As Collection
        Dim SQL As String
        DoCmd.Hourglass True
        SQL = ""
        Set MyCol = .Selections
        If Not MyCol Is Nothing Then
            For X = 1 To MyCol.Count
                If "" & SQL <> "" Then
                    SQL = SQL & ", "
                End If
                SQL = SQL & MyCol.Item(X)(0)
            Next X
        End If
        SQL = "Update SCR_Screens Set StartDte = #" & Me.StartDte & "#, EndDte = #" & Me.EndDte & "#  Where ScreenID In (" & SQL & ")"
        CurrentDb.Execute SQL, dbFailOnError
    End If
End With

ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, CodeContextObject.Name & " --> Batch Add Reports"
    Resume ExitNow
End Sub

Public Property Get SortMode() As Boolean
    SortMode = MvSortMode
End Property

Public Property Let SortMode(data As Boolean)
DoCmd.Hourglass True
    MvSortMode = data
    
    If MvSortMode = True Then
      Me.LstScreens.BackColor = 10545662 'RGB(245, 245, 245)
    Else
      Me.LstScreens.BackColor = RGB(255, 255, 255)
    End If
    
    Dim MyCtrl As Control
    For Each MyCtrl In Me.Controls
        ''Debug.Print MyCtrl.Name
        Select Case UCase(MyCtrl.Name)
        Case "CMDUP", "CMDDOWN"
            MyCtrl.Enabled = MvSortMode
        Case "CMDSORT", "LSTSCREENS"
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
    
    If MvSortMode = False Then
        LstScreens_AfterUpdate
    End If
    
DoCmd.Hourglass False
'
End Property

Private Sub CmdSort_AfterUpdate()
    Me.SortMode = CmdSort.Value
End Sub

Private Sub CmdUp_Click()
SortChange 1
End Sub

Private Sub CmdDown_Click()
SortChange -1
End Sub

Private Sub SortChange(Move As Integer)
On Error GoTo ErrorHappened
Dim LgID As Long, LgIDx As Long
Dim db As DAO.Database
Dim rst As DAO.RecordSet
Dim X As Long

If Me.LstScreens.ListIndex <> -1 Then
    LgID = LstScreens.Value
    LgIDx = LstScreens.ListIndex
   
    If LgIDx = 0 And Move = 1 Then 'IF FIRST ITEM AND MOVE UP THEN NOTHING TO DO
        GoTo ExitNow
    End If
    If LgIDx = LstScreens.ListCount - 1 And Move = -1 Then 'IF LAST ITEM AND MOVE DOWN THEN NOTHING TO DO
        GoTo ExitNow
    End If
    
    ReDim ArySort(LstScreens.ListCount - 1)
    For X = 0 To LstScreens.ListCount - 1
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
    Set rst = LstScreens.RecordSet
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
    LstScreens.Requery
    Me.Requery
    LstScreens.Value = LgID
    LstScreens_AfterUpdate
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

Private Sub cmdImport_Click()
    'Open migration utility if installed
    If IsProductInstalled("Migration Utility") Then
        DoCmd.OpenForm "MUT_Migrate"
    Else
        MsgBox "Please install the Migration Utility using the App Manager to import objects and data from legacy versions of Decipher.", vbInformation, "Migration Utility"
    End If
End Sub

Private Sub CmdScreenDelete_Click()
On Error GoTo ErrorHappened
Dim StScreenName As String, LngScreenID As Long


If Me.LstScreens.ListIndex = -1 Then
    MsgBox "You must select a screen to delete.", vbInformation, "Delete Screen"
    GoTo ExitNow
End If

StScreenName = Me.LstScreens.Column(1, Me.LstScreens.ListIndex + 1)
LngScreenID = Me.LstScreens.Value
If MsgBox("Delete " & StScreenName & "?", vbQuestion + vbDefaultButton2 + vbOKCancel, "Confirm Delete") = vbOK Then
    If Me!ScreenID = LngScreenID Then
        Me.RecordSource = "Select Top 1 *  From SCR_Screens Where ScreenID <> " & LngScreenID
    End If
    
    'SA 11/12/12 - Switched to new method to make sure user data is deleted
    SCR_DeleteScreenByID LngScreenID

    Me.LstScreens.Requery
End If



ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
    Resume
End Sub

Private Sub CmdScreenNew_Click()
On Error GoTo ErrorHappened
Dim db As DAO.Database
Dim StScreenName As String, LgID As Long


StScreenName = InputBox("Enter New Screen Name", "New Screen")
If "" & StScreenName <> "" Then
    Set db = CurrentDb
    db.Execute "Insert Into SCR_Screens(ScreenName, Sort) Values('" & StScreenName & "', " & Nz(DMax("Sort", "SCR_Screens"), 0) + 1 & ")"
    Me.LstScreens.Requery
    LgID = Nz(DMax("ScreenID", "SCR_Screens", "ScreenName = " & Chr(34) & StScreenName & Chr(34)), 0)
    If LgID <> 0 Then
        Me.LstScreens.Value = LgID
        LstScreens_AfterUpdate
    End If
    
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub CmdScreenRefresh_Click()
Me.LstScreens.Requery
End Sub

Private Sub CmdScreenRename_Click()
On Error GoTo ErrorHappened
Dim StScreenName As String, LngScreenID As Long
Dim StNewName As String

If Me.LstScreens.ListIndex = -1 Then
    MsgBox "You must select a screen to rename.", vbInformation, "Rename Screen"
    GoTo ExitNow
End If

StScreenName = Me.LstScreens.Column(1, Me.LstScreens.ListIndex + 1)
LngScreenID = Me.LstScreens.Value

StNewName = InputBox("Please enter a new name for '" & StScreenName & "'", "Rename Screen", StScreenName)
If "" & StNewName = "" Or "" & StNewName = StScreenName Then
    GoTo ExitNow
End If


CurrentDb.Execute ("Update SCR_Screens Set ScreenName = '" & StNewName & "' Where ScreenID = " & CStr(LngScreenID))
Me.LstScreens.Requery

Me.Requery

ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub CmdSync_Click()
    'SA 8/9/2012 - Open Decipher Screens Sync if installed
    If IsProductInstalled("Decipher Screens Sync") Then
        DoCmd.OpenForm "SCRSYNC_Sync"
    Else
        MsgBox "Please install Decipher Screens Sync using the App Manager to import and sync screens from Decipher 3.0 and later.", vbInformation, "Sync Utility"
    End If
End Sub

Private Sub Form_Current()
On Error Resume Next
    SetList2Visible Me.ChkList2Use 'HIDE OR UNHIDE THE SECONDARY LIST INFO
    SetList3Visible Me.ChkList3Use 'HIDE OR UNHIDE THE Tertiary list info
    SetDateVisible Me.ckDateUse     ' HIDE OR UNHIDE THE DATES INFo
    'SET THE LIST BOX TO THE CORRECT TABLE FOR LIST FIELDS
    CmboListSource1_AfterUpdate
    CmboListSource2_AfterUpdate
    If Me.ckDateUse Then
        With Me.SfDates.Form!FieldName
            .RowSource = "" & CmboPrimary
            .Requery
        End With
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)
'Load form settings
'SA 8/9/2012 - Moved record source strings to code and select first screen in list
On Error GoTo ErrorHappened
    genUtils.ToggleAccessMenus (False)
    
    Me.RecordSource = "SELECT ScreenID, ScreenName, [Included], FormName, PrimaryRecordSource, PrimaryRecordSourceType, " & _
        "DateUse, StartDte, EndDte, PrimaryListBoxRecordSource, PrimaryListBoxRecordSourceType, PrimaryListBoxCaption, " & _
        "PrimaryListBoxMulti, CustomCriteriaListBoxRecordSource, SecondaryListBoxUse, SecondaryListBoxDependency, " & _
        "SecondaryListBoxMulti, SecondaryListBoxRecordSource, SecondaryListBoxRecordSourceType, SecondaryListBoxCaption, " & _
        "[Sort], RefID, TertiaryListBoxUse, TertiaryListBoxDependency, TertiaryListBoxMulti, TertiaryListBoxRecordSource, " & _
        "TertiaryListBoxRecordSourceType, TertiaryListBoxCaption, TertiaryListBoxPrimaryDependency " & _
        "FROM SCR_Screens ORDER BY [Sort],[ScreenName]"
    LstScreens.RowSource = "SELECT ScreenID, ScreenName, Included, Sort FROM SCR_Screens ORDER BY [Sort],[ScreenName]"
    
    If LstScreens.ListCount > 1 Then
        LstScreens.Value = LstScreens.Column(0, 1)
    End If
    
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_Screens:Form_Open"
    Resume ExitNow
End Sub

Private Sub LstScreens_AfterUpdate()
   
    Me.RecordSource = "Select * From SCR_Screens Where ScreenID = " & Nz(Me.LstScreens.Value, 0)
        
    If Me.SortMode = False Then
    On Error Resume Next
        Dim rst As DAO.RecordSet
        Set rst = Me.RecordsetClone
        With rst
            .MoveFirst
            .FindFirst "ScreenID =" & Nz(Me.LstScreens.Value, 0)
            Me.StartDte = !StartDte
            Me.EndDte = !EndDte
            If .NoMatch = False Then
                Me.Bookmark = .Bookmark
            End If
        End With
        
    End If
    
End Sub
Private Sub RecordsourceUpdated(CtrlName As String)
    Dim MyCtrl As ComboBox
    Set MyCtrl = Me.Controls.Item(CtrlName)
        Me.Controls.Item(CtrlName & "Type") = IIf(MyCtrl.Column(1, MyCtrl.ListIndex) = "Table", 0, 1)
    Set MyCtrl = Nothing
End Sub
