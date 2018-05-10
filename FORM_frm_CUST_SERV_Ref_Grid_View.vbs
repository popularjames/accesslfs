Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdViewImage_Click()

    Dim strFileName As String
    strFileName = Me.RecordSet("RefLink")
    SetFileReadOnly (strFileName)
    If UCase(Right(strFileName, 3)) = "TIF" Then
        If UCase(left(GetPCName(), 9)) = "TS-FLD-03" Then
            Shell "explorer.exe " & strFileName, vbNormalFocus
        Else
            Shell "C:\Program Files (x86)\Common Files\microsoft shared\MODI\11.0\mspview.exe " & strFileName, vbNormalFocus
        End If
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If
End Sub

Private Sub FaxIND_AfterUpdate()

If Me.RefType = "IMAGE" Then
    MsgBox "Faxing of Images are not allowed. Your changes will not be saved.", vbCritical, "Customer Service"
    Me.FaxIND = 0
End If

End Sub

Private Sub FaxIND_BeforeUpdate(Cancel As Integer)

'If Me.RefSubType = "MR" Then
'    MsgBox "Faxing of Medical Records are not allowed. Your changes will not be saved", vbCritical, "Customer Service"
'    Cancel = True
'End If
End Sub
