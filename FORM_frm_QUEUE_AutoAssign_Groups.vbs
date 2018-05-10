Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdEdit_Click()
    Me.Productivity.Enabled = True
    Me.FilFactor.Enabled = True
    Me.MinAgeGroup.Enabled = True
    Me.MaxAgeGroup.Enabled = True
End Sub

Private Sub Form_AfterInsert()
    Me.cmdEdit.SetFocus
    Me.Productivity.Enabled = False
    Me.FilFactor.Enabled = False
    Me.MinAgeGroup.Enabled = False
    Me.MaxAgeGroup.Enabled = False
  
End Sub

Private Sub Form_AfterUpdate()
 If IsSubForm(Me) Then
    Me.Parent.RefreshData
    Me.cmdEdit.SetFocus
    Me.Productivity.Enabled = False
    Me.FilFactor.Enabled = False
    Me.MinAgeGroup.Enabled = False
    Me.MaxAgeGroup.Enabled = False
  
 End If
End Sub

Private Sub cmdSAave_Click()
On Error GoTo Err_cmdSAave_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSAave_Click:
    Exit Sub

Err_cmdSAave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSAave_Click
    
End Sub
