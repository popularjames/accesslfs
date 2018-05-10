Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event UpdateRow()

Private Sub CmdCancel_Click()

DoCmd.Close acForm, Me.Name

End Sub

Private Sub cmdUpdate_Click()

RaiseEvent UpdateRow

DoCmd.Close acForm, Me.Name


End Sub
