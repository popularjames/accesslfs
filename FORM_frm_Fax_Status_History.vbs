Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
        'Me.RecordSource = ""
        'Call Get_Data
End Sub


Public Sub Form_GotFocus()
      Dim SQL As String
       
       'MsgBox (Me.txtDocId)
       SQL = "select * from v_Fax_Status_History"
       Me.RecordSource = SQL
       Me.Refresh
         
End Sub
