Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130205 KD FiXED  Come on guys, really?

Private genUtils As CT_ClsGeneralUtilities  ' 20130205 KD FIxed this too..
Dim StProductionDb As String

Private Sub ClearTable_Click()
On Error GoTo ErrorHappened

    Dim SqlStmt As String

'    SqlStmt = "DELETE * FROM CT_CurrentlyLoggedIn " & _
              "WHERE NOT (empCurrentlyLoggedIn = fOSUserName() AND " & _
              "UDate = (SELECT Max(UDate) FROM CT_CurrentlyLoggedIn WHERE empCurrentlyLoggedIn = fOSUserName()))"

    SqlStmt = "DELETE * FROM CT_CurrentlyLoggedIn " & _
              "WHERE NOT (empCurrentlyLoggedIn = '" & Identity.UserName & "' AND " & _
              "UDate = (SELECT Max(UDate) FROM CT_CurrentlyLoggedIn WHERE empCurrentlyLoggedIn = '" & Identity.UserName & "' AND " & _
              "Computer = '" & Identity.Computer & "'))"

    DoCmd.SetWarnings False
    DoCmd.RunSQL SqlStmt
    DoCmd.SetWarnings True

    lstActiveUsers.Requery
    
    If StProductionDb <> "" Then
        SqlStmt = "DELETE * FROM [" & StProductionDb & "].CT_CurrentlyLoggedIn " & _
                  "WHERE NOT (empCurrentlyLoggedIn = '" & Identity.UserName & "' AND " & _
                  "UDate = (SELECT Max(UDate) FROM [" & StProductionDb & "].CT_CurrentlyLoggedIn WHERE empCurrentlyLoggedIn = '" & Identity.UserName & "' AND " & _
                  "Computer = '" & Identity.Computer & "'))"
    
        DoCmd.SetWarnings False
        DoCmd.RunSQL SqlStmt
        DoCmd.SetWarnings True
    
        lstActiveUsersProd.Requery
    End If

ExitNow:
    Exit Sub

ErrorHappened:
    MsgBox Error$, vbCritical, "ERROR"
    Resume ExitNow

End Sub

Private Sub CloseForm_Click()
On Error GoTo Err_CloseForm_Click


    DoCmd.Close

Exit_CloseForm_Click:
    Exit Sub

Err_CloseForm_Click:
    MsgBox Err.Description
    Resume Exit_CloseForm_Click

End Sub

Private Sub Form_Load()
Me.lblCurrentDb.Caption = "Current DB:     " & DBEngine(0)(0).Name
lstActiveUsers.RowSource = "SELECT empCurrentlyLoggedIn, UDate, Computer FROM CT_CurrentlyLoggedIn;"

StProductionDb = Nz(DLookup("Value", "CT_Options", "OptionName = 'ProductionDb'"), "")
If StProductionDb <> "" Then
    Me.lblProductionDb.Caption = "Production DB:  " & StProductionDb
    lstActiveUsersProd.RowSource = "SELECT empCurrentlyLoggedIn, UDate, Computer FROM [" & StProductionDb & "].CT_CurrentlyLoggedIn;"
Else
    Me.lblProductionDb.Caption = "Production DB:  (Unknown)"
    lstActiveUsersProd.RowSource = ""
End If

End Sub

Private Sub RefreshList_Click()

    lstActiveUsers.Requery
    lstActiveUsersProd.Requery

End Sub

Private Sub Form_Resize()
On Error GoTo ErrorHappened
'SA 1/18/2012 - CR2392 Resize screen objects
    If genUtils Is Nothing Then Set genUtils = New CT_ClsGeneralUtilities
    
    genUtils.SuspendLayout
    
    'Set min width
    If Me.InsideWidth < 6400 Then
        Me.InsideWidth = 6400
    End If
    'Set min height
    If Me.InsideHeight < 4185 Then
        Me.InsideHeight = 4185
    End If
    
    'Width
    lblCurrentDb.Width = Me.InsideWidth - (lblCurrentDb.left * 2)
    lstActiveUsers.Width = lblCurrentDb.Width
    lblProductionDb.Width = lblCurrentDb.Width
    lstActiveUsersProd.Width = lblCurrentDb.Width
    
    'Height
    RefreshList.top = Me.InsideHeight - RefreshList.Height - 105
    ClearTable.top = RefreshList.top
    CloseForm.top = RefreshList.top
    lstActiveUsers.Height = (Me.InsideHeight - lblCurrentDb.Height - _
        lblProductionDb.Height - RefreshList.Height - 395) / 2
    lblProductionDb.top = lstActiveUsers.top + lstActiveUsers.Height + 90
    lstActiveUsersProd.top = lblProductionDb.top + lblProductionDb.Height + 15
    lstActiveUsersProd.Height = lstActiveUsers.Height
    
ExitNow:
    genUtils.ResumeLayout
    Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub
