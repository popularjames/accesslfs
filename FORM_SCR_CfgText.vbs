Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdOk_Click()
DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub Form_Load()
On Error Resume Next
Me.TxtText = "" & Me.OpenArgs
End Sub
