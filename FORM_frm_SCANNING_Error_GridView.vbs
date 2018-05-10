Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ValidationDt_AfterUpdate()
    If ValidationDt & "" <> "" Then
        If PDFCnt <> TIFCnt And ValidationDt <> "1/1/1900" Then
            MsgBox "Error: PDTCnt and TIFCnt are not the same."
            ValidationDt = "1/1/1900"
        Else
            If IsDate(ValidationDt) = False Then
                MsgBox "Error: Please put in a valid date"
                ValidationDt.SetFocus
            End If
        End If
    End If
End Sub
