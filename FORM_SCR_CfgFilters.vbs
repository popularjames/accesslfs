Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_Close()
On Error GoTo UpdateFilterListError

'If no screen I was passed to the for them just quit
If Len(Me.OpenArgs) = 0 Then Exit Sub

Scr(CByte(Me.OpenArgs)).CmboFilters.Requery


UpdateFilterListExit:
    On Error Resume Next
    Exit Sub

UpdateFilterListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Updating Filter List Box!", vbInformation + vbOKOnly, "FILTERS"
    Resume UpdateFilterListExit
End Sub


Private Sub ScreenName_AfterUpdate()
On Error Resume Next
Me.Caption = "Edit Filters: " & Me.ScreenName_Label.Caption
End Sub
