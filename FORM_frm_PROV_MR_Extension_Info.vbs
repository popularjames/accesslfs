Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
    'MG refresh data sheet
    Dim sqlString As String
    sqlString = " CnlyClaimNum = " & Chr(34) & Me.Parent.Form.CnlyClaimNum & Chr(34)
    Me.Form.filter = sqlString
    Me.Form.FilterOn = True
    Me.Form.Requery
    Me.Form.Refresh
End Sub
