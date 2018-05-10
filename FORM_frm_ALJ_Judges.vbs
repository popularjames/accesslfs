Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'2014-07-21 VS: Add ALJ Judges Screen
' ALJ

Private Sub cmdViewHistory_Click()

    Dim strSQL As String
    strSQL = "select * from v_ALJ_Hearing_Judges_History where Id= " & Me.txtId & " order by Updatedate DESC"
    Me.RecordSource = strSQL
    
End Sub

Private Sub cmdSave_Click()

    Call Update_Contact(Nz(Me.txtId, 0), Judge)
    
End Sub
