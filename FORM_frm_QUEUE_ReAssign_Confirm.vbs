Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Event ConfirmReAssign(Action As String)


Private Sub cbAll_Click()
    RaiseEvent ConfirmReAssign("all")
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub cbNo_Click()
    RaiseEvent ConfirmReAssign("no")
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cbYes_Click()
    RaiseEvent ConfirmReAssign("yes")
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub Form_Close()
    RemoveObjectInstance Me
End Sub

Private Sub Form_GotFocus()
    rationaleText.Enabled = False
End Sub
