Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
    DoCmd.Close
End Sub

Private Sub Form_Load()
    Dim temp
    If Me.OpenArgs & "" <> "" Then
        temp = Split(Me.OpenArgs(), ";")
        Me.Caption = "Error Message for claim " & temp(0)
        txtErrorMessage = temp(1)
        txtErrorMessage.SetFocus
        txtErrorMessage.SelLength = 0
    End If
End Sub
