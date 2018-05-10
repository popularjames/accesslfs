Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cbUserID_AfterUpdate()
'Clear out any previous selections from Profile & Company combo boxes
    NextSelections_Clear
    NextSelections_Refresh
End Sub

Private Sub cmdUpdate_Click()
'************************************************
' Revised Monday 4/15/2013 by Kathleen C Flanagan
' Set the SupervisorID value for the user
'************************************************
    Dim strUserID           As String
    Dim strCurProfileID     As String
    Dim strCurCompanyID     As String
    Dim strNextProfileID    As String
    Dim strNextCompanyID    As String
    Dim strNextSupervisorID As String '4/15/2013 KCF
    
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim strConfirmMsg As String
    
    On Error GoTo Err_handler
    
    strUserID = Me.frm_Admin_R3_Profile_User.Form.UserID
    strCurProfileID = Me.frm_Admin_R3_Profile_User.Form.ProfileID
    strCurCompanyID = Me.frm_Admin_R3_Profile_User.Form.CompanyID
    strNextProfileID = Me.cbProfileID
    strNextCompanyID = Me.cbCompany
    strNextSupervisorID = Me.cbSupervisorID '4/15/2013 KCF
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_ADMIN_User_R3_Permissions_Update"
    cmd.Parameters.Refresh
    cmd.Parameters("@pUserID") = strUserID
    cmd.Parameters("@pCurProfileID") = strCurProfileID
    cmd.Parameters("@pCurCompanyID") = strCurCompanyID
    cmd.Parameters("@pNextProfileID") = strNextProfileID
    cmd.Parameters("@pNextCompanyID") = strNextCompanyID
    cmd.Parameters("@pNextSupervisorID") = strNextSupervisorID '4/15/2013 KCF
    cmd.Execute
    
    strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        NextSelections_Clear
        NextSelections_Refresh
        
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_ADMIN_User_R3_Permissions_Update"
        Err.Raise 50001, "usp_QA_Review_Worktable_Update", strErrMsg
        
    End If
    
    strConfirmMsg = cmd.Parameters("@pConfirmMsg")
     If cmd.Parameters("@RETURN_VALUE") = 0 Or strConfirmMsg <> "" Then
        MsgBox (strConfirmMsg)
    End If
    
    NextSelections_Clear
    NextSelections_Refresh
    Me.frm_Admin_R3_Profile_User_Hist.Requery
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub

Private Sub NextSelections_Clear()
'Clear out any values from the Profile and Company ID combo boxes
    Me.cbProfileID = ""
    Me.cbCompany = ""
    Me.cbSupervisorID = ""
    
End Sub

Private Sub NextSelections_Refresh()
'Update the values in the combobox

If Me.cbUserID & "" <> "" Then
    Me.cbProfileID.Requery
    Me.cbCompany.Requery
    Me.cbSupervisorID.Requery
End If

End Sub
