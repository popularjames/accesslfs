Option Compare Database
Option Explicit


Private Const msgboxtitle As String = ""


Global lngEventID As Long
Global lngActionID As Long
Global currActionID As Variant
Global isSave As String
Global GblParentEvent As String


Public Function setNotificaton(lngActionID As Variant, lngEventID As Long)

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim locEventID As Long
Dim locActionID As Long

locActionID = lngActionID
locEventID = lngEventID

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_CUST_Event_Notification_LOAD"
                cmd.Parameters.Refresh
                cmd.Parameters("@EventID") = locEventID
                cmd.Parameters("@ActionID") = locActionID
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")
                
If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, msgboxtitle
End If

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Function



Public Function SendNotificatonEmail(lngEventID As Long)

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim locEventID As Long
Dim locActionID As Long

locEventID = lngEventID
locActionID = lngActionID

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_CUST_Send_Notification"
                cmd.Parameters.Refresh
                cmd.Parameters("@EventID") = locEventID
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")
                
If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, msgboxtitle
End If

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Function


Public Sub EventSearch(lEventID As Variant)

Dim strClaimNum As String
Dim strMsgText As String
Dim strMsgText2 As String
Dim strRelClaimId As String

On Error GoTo ErrHandler

strMsgText = "The Event ID you've entered is incorrect or does not exists."
strMsgText2 = "Please save your work and close the current event window and try again."


If IsNumeric(lEventID) = False Then
MsgBox strMsgText, vbCritical + vbOKOnly, "Incorrect EventID"
        Exit Sub
End If


'If CurrentProject.AllForms("frm_CUST_Main").IsLoaded = True Then
'    MsgBox strMsgText2, vbInformation + vbOKOnly, "Close Event Window"
'    Exit Sub
'End If

lngEventID = lEventID
'lngEventID = Me.txtSearchFor

'MG 10/1/2013 be careful when using DLookup as it doesn't like underscore, dash and other symbols. It's safer to use alphanumeric characters
strRelClaimId = Nz(DMax("[RelatedClaimID]", "v_CUST_EVENT_Related_Claims", "[EventID] =" & lngEventID & ""), "0")

'If strRelClaimId = "0" Then
'        MsgBox "No claim found for eventID " & lngEventID, vbOKOnly, "Alert"
'        'Exit Sub
'End If

strClaimNum = Nz(DLookup("[CnlyClaimNum]", "v_CUST_EVENT_Related_Claims ", "[RelatedClaimID] =" & strRelClaimId & ""), "0")

LaunchNewCustClaimEvent strClaimNum

Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & "(" & Err.Description & ")", vbOKOnly + vbExclamation, "oops"
'DoCmd.Close acForm, "frm_CUST_Quick_Launch"
End Sub