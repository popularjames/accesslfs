Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mrsEventClaimActions As ADODB.RecordSet
Private mlEventID As Long
Private mstrCnlyClaimNum As String
Private mstrUserName As String
Private mstrAssignedUser As String
Private mstrDescription As String
Private mstrduedate As String

Public Property Let ActionCnlyClaimNum(ByVal vData As String)
    mstrCnlyClaimNum = vData
    mstrAssignedUser = "none"
    Me.Description = ""
    Me.Description = ""
    Me.DueDate = ""
End Property
Public Property Let ActionUserName(ByVal vData As String)
    mstrUserName = vData
End Property
Public Property Let ActionEventID(ByVal vData As Long)
    mlEventID = vData
End Property
Public Property Let ActionrsEventClaimActions(vData As ADODB.RecordSet)
    Set mrsEventClaimActions = vData
End Property

Private Sub Cancel_Click()
    DoCmd.Close acForm, "frm_CUST_Event_Claim_Action_Add", acSaveNo
End Sub
Private Sub Description_Change()
    mstrDescription = Me.Description
End Sub

Private Sub DueDate_BeforeUpdate(Cancel As Integer)

If IsDate(Me.DueDate) = False Then
    MsgBox "Please enter a valid date.", vbCritical + vbOKOnly, "Invalid Date"
End If

End Sub

'Private Sub DueDate_Change()
'    mstrduedate = Me.DueDate
'End Sub

Private Sub addNoteDetail(currActionID As Variant, CurrEventID As Long)

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
                cmd.Parameters("@ActionID") = currActionID
                cmd.Parameters("@EventID") = CurrEventID
                cmd.Parameters("@Notes") = Me.Notes
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, ""   ' msgboxtitle
End If
 
Set MyCodeAdo = Nothing
Set cmd = Nothing

'Forms("frm_CUST_Main").Requery
Set MyCodeAdo = Nothing

DoCmd.Close

End Sub



Private Sub Save_Click()

'Curlan Johnson 10/2/12

Dim CurrEventID As Long
    
Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim strSQL As String
Dim strProcCd As String
Dim ErrMsg As String
Dim strUser As String
        
   On Error GoTo ErrHandler
        
    CurrEventID = IIf(mlEventID = 0, lngEventID, mlEventID)
    
    If mstrAssignedUser = "none" Then
        MsgBox "You must select a user to assign the task.", vbOKOnly
        Exit Sub
    End If

    If IsDate(Me.DueDate) = 0 Then
        MsgBox "You must enter a due date for the task.", vbOKOnly
        Exit Sub
    End If
        
    If Len(Me.Description) = 0 Then
        MsgBox "You must enter a description for the task.", vbOKOnly
        Exit Sub
    End If

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_CUST_AddNotes_Dtl"
                cmd.Parameters.Refresh
                cmd.Parameters("@EventID") = CurrEventID
                cmd.Parameters("@ClaimNum") = mstrCnlyClaimNum
                cmd.Parameters("@Description") = Me.Description
                cmd.Parameters("@AssignedTo") = SelectedUser.Column(0, SelectedUser)
                cmd.Parameters("@DueDate") = Me.DueDate
                cmd.Parameters("@Notes") = Me.Notes
                cmd.Execute
                currActionID = cmd.Parameters("@NextActionID")
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, ""   ' msgboxtitle
End If
 
    addNoteDetail currActionID, CurrEventID
    'Add to email notification
    setNotificaton currActionID, CurrEventID
 
DoCmd.Close acForm, "frm_CUST_Event_Claim_Action_Add", acSaveNo
 
Cleanup:
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical, "Error encountered"
    Resume Cleanup

End Sub

Private Sub SelectedUser_Change()
    mstrAssignedUser = Nz(SelectedUser.Column(0, SelectedUser), "InvalidUser")
End Sub
