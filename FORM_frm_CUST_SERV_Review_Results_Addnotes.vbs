Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdSave_Click()



Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim strSQL As String
Dim strProcCd As String
Dim ErrMsg As String
Dim strUser As String

strUser = Identity.UserName

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_CUST_AddNotes"
                cmd.Parameters.Refresh
                cmd.Parameters("@ActionID") = intActionID
                cmd.Parameters("@EventID") = intEventID
                cmd.Parameters("@Notes") = Me.TxtNote
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, msgboxtitle
End If
 
Set MyCodeAdo = Nothing
Set cmd = Nothing

'Forms("frm_CUST_Main").Requery
Set MyAdo = Nothing

DoCmd.Close
End Sub

Private Sub Form_Load()

Me.AddNote.Caption = "Add Note for Event ID " & intEventID & " and Action ID " & intActionID & ""

End Sub
