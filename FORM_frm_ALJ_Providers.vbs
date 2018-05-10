Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'2014-06-18 VS: Add ALJ Providers Screen
' ALJ


Private Sub cmdViewHistory_Click()

    Dim strSQL As String
    strSQL = "select * from v_ALJ_Providers_History where Id= " & Me.txtId & " order by Updatedate DESC"
    Me.RecordSource = strSQL
    
End Sub

Private Sub cmdSave_Click()

    Call Update_Contact(Nz(Me.txtId, 0), Provider)
    
End Sub
