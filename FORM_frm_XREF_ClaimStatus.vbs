Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Const CstrFrmAppID As String = "ClmStatusCode"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub ClmStatusDesc_BeforeUpdate(Cancel As Integer)
    If Me.ClmStatusDesc & "" = "" Then
        MsgBox "Status description can not be blank.", vbCritical
        Cancel = True
    End If
End Sub

Private Sub ClmStatusGroup_BeforeUpdate(Cancel As Integer)
    If Me.ClmStatusGroup & "" = "" Then
        MsgBox "Claim status group can not be blank.", vbCritical
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    iAppPermission = UserAccess_Check(Me)
End Sub

Private Sub ValidationInd_BeforeUpdate(Cancel As Integer)
    If Me.ValidationInd & "" = "" Then
        MsgBox "Validation indicator can not be blank.", vbCritical
        Cancel = True
    End If
End Sub

Private Sub WebDesc_BeforeUpdate(Cancel As Integer)
    If Me.WebDesc & "" = "" Then
        MsgBox "Web portal description can not be blank.", vbCritical
        Cancel = True
    End If
End Sub
