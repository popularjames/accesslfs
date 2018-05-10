Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim mstrCnlyClaimNum As String
Dim mstrImageCreateDt As String



Property Let CnlyClaimNum(data As String)
     mstrCnlyClaimNum = data
End Property

Property Let ImageCreateDt(data As String)
     mstrImageCreateDt = data
End Property

Public Sub RefreshData(Optional SelectMode As String = "NONE")
    
    Dim strProcName As String

    Dim strError As String
    On Error GoTo ErrHandler
    
    Dim oAdo As clsADO
   
    strProcName = "frm_AuditClm_RelatedImage_Assign"
   
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_AuditClm_RelatedImage_Assign_Load"
        .Parameters.Refresh
        .Parameters("@pOrigCnlyClaimNum") = mstrCnlyClaimNum
        .Parameters("@pUserID") = Identity.UserName
        .Parameters("@pImageCreateDt") = mstrImageCreateDt
        .Parameters("@pSelectMode") = SelectMode
        .Execute
        
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Or .Parameters("@RETURN_VALUE") <> 0 Then
            Stop
            LogMessage strProcName, "ERROR", "There was a problem loading the related claims!", .Parameters("@pErrMsg"), True
            GoTo ErrHandler
        End If
    End With
    
    Me.SubForm.Form.RecordSource = "select * from AuditClm_RelatedImage_Assign_Worktable where UserID = '" & Identity.UserName & "' and OrigCnlyClaimNum= '" & mstrCnlyClaimNum & "' order by PatLastName, PatFirstName, PatDOB"
    
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub

Private Sub CmdCancel_Click()
    MsgBox "You have cancelled the process. The image was not propagated to other claims.", vbInformation, "User Cancelled"
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdPropagate_Click()

    Dim strError As String
    Dim strProcName As String
    Dim UserAnswer As Integer
    
    UserAnswer = MsgBox("Are you sure you want to propagate the image to the selected claims?", vbYesNo + vbQuestion, "Confirmation")
    
    If UserAnswer = vbNo Then
        Exit Sub
    End If
    
    Dim oAdo As clsADO
   
    strProcName = "frm_AuditClm_RelatedImage_Process"
   
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_AuditClm_RelatedImage_Process"
        .Parameters.Refresh
        .Parameters("@pOrigCnlyClaimNum") = mstrCnlyClaimNum
        .Parameters("@pUserID") = Identity.UserName
        .Parameters("@pImageCreateDt") = mstrImageCreateDt
        .Execute
        
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Or .Parameters("@RETURN_VALUE") <> 0 Then
            Stop
            LogMessage strProcName, "ERROR", "There was a problem propagating the claims!", .Parameters("@pErrMsg"), True
            GoTo ErrHandler
        End If
    End With
    
exitHere:

    DoCmd.Close acForm, Me.Name

    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
    
    
End Sub

Private Sub cmdSelectSwitch_Click()
    Me.SubForm.Form.RecordSource = "select * from AuditClm_RelatedImage_Assign_Worktable where 1=2"
    RefreshData "SWITCH"
End Sub
