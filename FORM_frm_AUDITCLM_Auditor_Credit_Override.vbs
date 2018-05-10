Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event OverrideReason(OverrideReason As String, Cancel As Boolean)


Private Sub CmdCancel_Click()
    RaiseEvent OverrideReason("", True)
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strErrMsg As String
    
    
    If Me.OverrideReason & "" = "" Then
        MsgBox "Error: Comment can not be blank.", vbInformation
        Exit Sub
    End If
    
    RaiseEvent OverrideReason(Me.OverrideReason, False)
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub
