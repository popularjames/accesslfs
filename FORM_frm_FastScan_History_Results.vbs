Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub SplitCoverSheetNum_DblClick(Cancel As Integer)
    If Me.SplitCoverSheetNum.Value <> "--" Then
        Me.Parent.OpenCoverSheetNum = Me.SplitCoverSheetNum.Value
        Me.Parent.RefreshScreen
    End If
End Sub
