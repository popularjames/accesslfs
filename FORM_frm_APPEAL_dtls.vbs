Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Current()
'This sub doesn't execute for 0 record case
On Error GoTo err_hndlr
    'Debug.Print Me.PackageID & "|" & Me.PackageName
    If Me.Count = 0 Then
        Me.Parent.PackageDetails.Form.filter = " PackageID=-1"
    Else
        Me.Parent.PackageDetails.Form.filter = " PackageID=" & Me.PackageID & " AND PackageName = '" & Me.PackageName & "'"
    End If
    Me.Parent.PackageDetails.Form.FilterOn = True
Exit Sub

err_hndlr:
End Sub
