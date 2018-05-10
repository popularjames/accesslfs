Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event DateSelected(SelectedDate As Date)


Private mDatePassed As Date

Public Property Let DatePassed(data As Date)
    mDatePassed = data
End Property


Private Sub Update_Click()
    On Error Resume Next
    RaiseEvent DateSelected(Me.Calendar0.Value)
    DoCmd.Close
End Sub


Public Sub RefreshData()
    Me.Calendar0.Value = mDatePassed
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub
