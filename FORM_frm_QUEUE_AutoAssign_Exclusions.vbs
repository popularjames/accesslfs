Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdEdit_Click()
Me.ExclusionType.Enabled = True
Me.ExclusionValue.Enabled = True
Me.Combo4.Enabled = True
End Sub

Private Sub cmdNew_Click()
On Error GoTo Err_cmdNew_Click


    DoCmd.GoToRecord , , acNewRec
    Me.ExclusionType.Enabled = True
    Me.ExclusionValue.Enabled = True
    Me.Combo4.Enabled = True

Exit_cmdNew_Click:
    Exit Sub

Err_cmdNew_Click:
    MsgBox Err.Description
    Resume Exit_cmdNew_Click
    
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click


    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_cmdDelete_Click:
    Exit Sub

Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
    
End Sub

Private Sub Form_AfterInsert()
 If IsSubForm(Me) Then
  Me.Parent.RefreshData
    Me.cmdEdit.SetFocus
   
    Me.ExclusionType.Enabled = False
    Me.ExclusionValue.Enabled = False
    Me.Combo4.Enabled = False
 End If
End Sub

Private Sub Form_AfterUpdate()
 If IsSubForm(Me) Then
    Me.Parent.RefreshData
    Me.cmdEdit.SetFocus
    Me.ExclusionType.Enabled = False
    Me.ExclusionValue.Enabled = False
    Me.Combo4.Enabled = False
 End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_cmdSave_Click


    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSave_Click:
    Exit Sub

Err_cmdSave_Click:
    MsgBox Err.Description
    Resume Exit_cmdSave_Click
    
End Sub
