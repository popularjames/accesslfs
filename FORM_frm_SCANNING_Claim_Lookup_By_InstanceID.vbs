Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim strSQL As String
Dim strWhere As String
Dim strCalledFrom As String


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSearch_Click()
    strWhere = ""
    If txtInstanceID & "" <> "" Then
        strWhere = "Where InstanceID = '" & txtInstanceID & "'"
    End If
    
    If txtCnlyProvID & "" <> "" Then
        If strWhere = "" Then
            strWhere = "Where CnlyProvID = '" & txtCnlyProvID & "'"
        Else
            strWhere = strWhere & " and CnlyProvID = '" & txtCnlyProvID & "'"
        End If
    End If
    
    strSQL = "select * from v_SCANNING_Claim_Lookup_By_InstanceID " & strWhere
    
    Me.SubForm.Form.RecordSource = strSQL
End Sub

Private Sub Form_Close()
    If strCalledFrom <> "" Then
        DoCmd.OpenForm strCalledFrom
    End If
End Sub

Private Sub Form_Load()

    If Me.OpenArgs() & "" <> "" Then
        strCalledFrom = Me.OpenArgs()
    End If
    
End Sub
