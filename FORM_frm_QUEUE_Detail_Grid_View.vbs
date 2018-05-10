Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    If IsSubForm(Me) Then
        Me.RecordSource = ""
    Else
        Me.RecordSource = Me.Tag
    End If
End Sub
