Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event AttachmentSelected(strAttachmentType As String)

Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    If AttachmentType = "" Then
        MsgBox "Please select an attachment type first"
        Exit Sub
    End If
    
    RaiseEvent AttachmentSelected(AttachmentType)
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub
