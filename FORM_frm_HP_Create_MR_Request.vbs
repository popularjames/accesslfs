Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdGenerate_Click()

    Dim MyAdo As New clsADO
    Dim MyCodeAdo As New clsADO
    
    Dim rsMRRequest As ADODB.RecordSet
    Dim rsLTTRHr As ADODB.RecordSet
    Dim rsSessID As ADODB.RecordSet
    
    Dim cmd As ADODB.Command
    
    Dim sqlString As String
    Dim SqlProvCheck As String
    Dim SqlProvClear As String
    Dim SqlLetterHr As String
    Dim SqlGetSeesionID As String
            
    Dim strCnlyProvID As String
    Dim strSessID As String
    Dim strInstanceID As String
    
    
    strCnlyProvID = Me!frm_HP_Providers_Sub_form.Form!cnlyProvID & ""
    strInstanceID = Nz(Me.txtInstanceID, "")
    
    sqlString = "Insert into dbo.HP_Create_MR_Key select '" & strCnlyProvID & "', USER, GETDATE()"
    SqlProvCheck = "Select * from dbo.HP_Create_MR_Key where ProvNum = '" & strCnlyProvID & "'"
    SqlProvClear = "Delete from dbo.HP_Create_MR_Key where ProvNum = '" & strCnlyProvID & "'"
    SqlLetterHr = "select 1 from dbo.LETTER_Header where CnlyProvID = '" & strCnlyProvID & "' and LetterType like 'VADR%'"
    SqlGetSeesionID = "Select CMS_AUDITORS_CODE.dbo.udf_GetInstanceID() as SessionID"
               
               
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
           
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = SqlLetterHr
    Set rsLTTRHr = MyAdo.OpenRecordSet
                                      
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = SqlProvCheck
    Set rsMRRequest = MyAdo.OpenRecordSet
                                       
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = SqlGetSeesionID
    Set rsSessID = MyAdo.OpenRecordSet
                                       
    strSessID = rsSessID.Fields.Item("SessionID").Value
                                      
' check for CnlyProvID to be filled
    If IsNull(Me!frm_HP_Providers_Sub_form.Form!cnlyProvID) Or Me!frm_HP_Providers_Sub_form.Form!cnlyProvID = "" Then
        MsgBox "Missing Connolly Provider ID.", vbCritical
        GoTo Exit_Sub
    End If

'Check for active status
    If Me!frm_HP_Providers_Sub_form.Form!CurStatus <> "ACTIVE" Or IsNull(Me!frm_HP_Providers_Sub_form.Form!CurStatus) Then
        MsgBox "Cannot Generate A MR for a provider which status is not Active.", vbCritical
        GoTo Exit_Sub
    End If
                                
'Check for Letter header
    If rsLTTRHr.recordCount < 1 Then
       MsgBox "No Letter for this Provider", vbCritical
       GoTo Exit_Sub
    End If
                                       
'Check if someone else is running the selected provider
    If rsMRRequest.recordCount > 0 Then
       MsgBox "Request is being ran by " & rsMRRequest("UserID") & " on " & Format(rsMRRequest("ModDate"), "mm-dd-yyyy hh:mm:ss"), vbInformation
       GoTo Exit_Sub
    Else
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = MyAdo.CurrentConnection
        cmd.commandType = adCmdText
        cmd.CommandText = sqlString
        cmd.Execute
    End If
      
'Execute the HP_Create_MR_Request Mod
    Call HP_Create_MR_Request(strCnlyProvID, strSessID, strInstanceID)
      
      
'Clear the HP_Create_MR_Key table
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = MyAdo.CurrentConnection
        cmd.commandType = adCmdText
        cmd.CommandText = SqlProvClear
        cmd.Execute
      
        MsgBox "Request sent successfully", vbInformation
        
Exit_Sub:
    Set MyAdo = Nothing
    Set rsMRRequest = Nothing
    Set rsLTTRHr = Nothing
End Sub


Sub mod_LookUp_Provider()

If Me.txtProvNum <> "" Then
    Me.frm_HP_Providers_Sub_form.Enabled = True
    Me.cmdGenerate.Enabled = True
End If

Forms!frm_hp_create_MR_request!.frm_HP_Providers_Sub_form.Form.RecordSource = "select * from hp_providers where provnum = '" & Me.txtProvNum.Value & "'"

End Sub


Private Sub cmdMASSRequest_Click()

If MsgBox("You are about to process multiple request files. Are you sure?", vbYesNo + vbCritical, "Send Mass Request files") = vbYes Then
    If MsgBox("Last chance to back out. Do you want to continue?", vbYesNo + vbCritical, "Send Mass Request files") = vbYes Then
        Call HP_Create_MR_Request_MASS_MOVE
    End If
End If
End Sub

Private Sub cmdRefresh_Click()

Me.txtProvNum.Requery

End Sub

Private Sub Form_Load()

Forms!frm_hp_create_MR_request!.frm_HP_Providers_Sub_form.Form.RecordSource = "select * from hp_providers where provnum = ''"
Me.frm_HP_Providers_Sub_form.Enabled = False
Me.cmdGenerate.Enabled = False
Me.txtProvNum = ""

End Sub

Private Sub txtProvNum_AfterUpdate()

mod_LookUp_Provider

If IsNull(Me.txtProvNum) Then
Form_Load
End If

End Sub
Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
    
End Sub
