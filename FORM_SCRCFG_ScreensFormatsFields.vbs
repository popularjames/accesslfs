Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If DataErr = 2107 Then 'Validation Rule ERror (BUG)
    Response = 0
End If

End Sub
