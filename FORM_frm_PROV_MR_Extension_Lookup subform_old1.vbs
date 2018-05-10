Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Click()

    'Dim frmCurrentForm As Form
    'Set frmCurrentForm = Screen.ActiveForm
    'Dim frmName As String
    'MsgBox "form name = " & frmCurrentForm.Name
    
    If Len(Me.RequestNumber) > 0 Then
        Forms!frm_Prov_MR_Extension.lstSelectedClaims.AddItem Me.RequestNumber & ";" & Me.CnlyClaimNum
    End If
    
End Sub
