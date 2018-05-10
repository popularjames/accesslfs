Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    If Not Me.RecordSet Is Nothing Then
        If Me.FileIsMR = 1 Then
            Me.Parent.btnMarkNotImage.Caption = "Mark not Image"
        Else
            Me.Parent.btnMarkNotImage.Caption = "Mark as Image"
        End If
    End If
End Sub
