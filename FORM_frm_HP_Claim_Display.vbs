Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    RemoveWindow Me
End Sub

Public Sub AddItem(DisplayMsg As String)
    Me.lstStatDisplay.AddItem DisplayMsg
End Sub
