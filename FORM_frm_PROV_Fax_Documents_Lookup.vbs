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
    
    If Len(Me.InstanceId) > 0 Then
    
        'MG syntax below works for main form->subform1
        'Forms!frm_PROV_Fax_Documents_Grid_View.lstSelectedClaims.AddItem Me.InstanceID & ";" & Me.LetterType & ";" & Me.LetterReqDt & ";" & Me.MaxRefLink
        
        'MG syntax below works for main form->subform1->subform2
        Me.Parent!lstSelectedClaims.AddItem Me.InstanceId & ";" & Me.LetterType & ";" & Me.LetterReqDt & ";" & Me.MaxRefLink

    End If
    
End Sub


Private Sub MaxRefLink_Click()

    Dim strFileName As String
    strFileName = Me.MaxRefLink
    SetFileReadOnly (strFileName)
    If UCase(Right(strFileName, 3)) = "TIF" Then
        Shell "explorer.exe " & strFileName, vbNormalFocus
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If
    
End Sub
