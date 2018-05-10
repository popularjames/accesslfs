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
        Me.Caption = temp(0)
        txtMessage = temp(1)
    End If
End Sub
