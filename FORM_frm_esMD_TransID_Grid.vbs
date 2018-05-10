Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_DblClick(Cancel As Integer)

 
    If Me.CnlyClaimNum & "" <> "" Then
        Me.Parent.DisplayClaimScreen Me.CnlyClaimNum
    End If

End Sub
