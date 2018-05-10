Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub FieldName_LostFocus()
   On Error Resume Next
   If FieldName.Text = "" And Me!ScreenID <> "" Then
        Dim SQL As String
        SQL = "Delete FROM SCR_ScreensDateFilters Where FieldName Is Null And ScreenID=" & Me!ScreenID
        CurrentDb.Execute SQL
        Me.Requery
   End If
End Sub
