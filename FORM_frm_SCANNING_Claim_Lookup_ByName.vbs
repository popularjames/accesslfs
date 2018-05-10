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
    If txtLastName & "" <> "" Then
        strWhere = "Where LastName like '" & txtLastName & "*'"
    End If
    
    If txtFirstName & "" <> "" Then
        If strWhere = "" Then
            strWhere = "Where FirstName like '" & txtFirstName & "*'"
        Else
            strWhere = strWhere & " and FirstName like '" & txtFirstName & "*'"
        End If
    End If
    
    strSQL = "select * from v_SCANNING_Claim_Lookup_ByName " & strWhere
    
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
