Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdFindMissingLetters_Click()
    Dim myCode_ADO As New clsADO
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    Dim fso As New FileSystemObject
    
    Dim strFileName As String
    Dim strInstanceID As String
    Dim strNewInstanceID As String
    Dim dLetterReqDt As Date
    
    Dim strSQL As String
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = sqltext
    strSQL = "select distinct lc.InstanceID, lc.LetterReqDt, lc.LetterName " & _
             " from v_LETTER_Claim lc " & _
             " left join CMS_AUDITORS_CLAIMS.dbo.LETTER_Reprint_Logs rl " & _
             "      on rl.InstanceID = lc.InstanceID " & _
             " where lc.LetterReqDt = Convert(varchar(20), getdate(), 101) " & _
             " and lc.status = 'p' " & _
             " and rl.InstanceID is null"
    
    myCode_ADO.sqlString = strSQL
    Set rs = myCode_ADO.OpenRecordSet
    
    If rs.recordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            strFileName = rs("LetterName")
            If Not fso.FileExists(strFileName) Then
                strInstanceID = rs("InstanceID")
                dLetterReqDt = rs("LetterReqDt")
                
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = myCode_ADO.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_LETTER_Reprint"
                cmd.Parameters.Refresh
                cmd.Parameters("@InstanceID") = strInstanceID
                cmd.Parameters("@Auditor") = Identity.UserName()
                cmd.Parameters("@LtrDate") = dLetterReqDt
                cmd.Execute
                lblMissingLetter.Caption = strInstanceID & " -- New instance = " & cmd.Parameters("@NewInstanceID")
            End If
            rs.MoveNext
        Wend
    End If
    
    Set fso = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Set cmd = Nothing
    MsgBox "done"
End Sub
