Option Compare Database

'Needed for the RECON Screen Curlan Johnson 4/12/12
Global gbl_CnlyClmNum As String
Global gbl_CnlyClmNumLock As String
Global gbl_GUID As String
'Global gbl_sysUser As String
Global gbl_DocID As String
'Global gbl_UpdateUser As String
Global gbl_UserRights As String
Global Const gbl_MsgBoxTitleLTTR = "RECON Letters"
Global Const gbl_MsgBoxTitleMRLTR = "Incomplete Medical Records Request Letters"
Global Const gbl_FromFieldForMR = "Connolly Customer Service"
Global Const gbl_User = "lmtdAccess"
Global Const gbl_INC_Client_Id = 3
Global gbl_frmLoad As Integer
Global gbl_TriggerFormTotal
Global gbl_TriggerFormCurrent
Global Const gbl_Fax_Status_Send = "Sent"
Global Const gbl_Fax_Status_Waiting = "Waiting"


'MG 9/16/2013 Added new global variable for new fax interface on provider level. Basically, the old method get data from each record set and loop through
'This new method bypass SQL and use MS Access to pass data from form to report, which will be quicker
Global gbl_fax_To As String
Global gbl_fax_FaxNumber As String
Global gbl_fax_DocID As String
Global gbl_fax_Regarding As String
Global gbl_fax_Comment As String
Global gbl_fax_From As String
'------------End--------------------------


Function UserRights()

UserRights = Nz(DLookup("[Rights]", "QUEUE_RECON_Review_Admin", "[UpdateUser] ='" & Identity.UserName & "'"), "PwrUser")

End Function

Function UserRights_Inc()
'Linked table QUEUE_MR_Admin lists users who are allowed to send fax requests for MR
UserRights_Inc = Nz(DLookup("[Rights]", "QUEUE_MR_Admin", "[UpdateUser] ='" & GetUserName & "'"), "PwrUser")

End Function

Function gbl_sysUser()

If UserRights = "user" Or UserRights = "Admin" Then
    gbl_sysUser = "*"
Else
    gbl_sysUser = Identity.UserName
End If

End Function

Function gbl_sysUser_Inc()
'For Incomplete Medical Records Fax Queue only people who are allowed to send faxes have access to records on this screen.
If UserRights_Inc = "user" Or UserRights_Inc = "Admin" Then
    gbl_sysUser_Inc = "%"
Else
    gbl_sysUser_Inc = ""
End If

End Function