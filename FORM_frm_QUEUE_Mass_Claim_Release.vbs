Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event ClaimProcessed()
Public Event FormUnload()

Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Private mbProcessed As Boolean
Private mlstClaimList As listBox
Private mstrErrMsg As String


Property Set claimList(data As listBox)
    Dim varItem
    Dim strSQL As String
    Dim iClaimIDPos As Integer
    Dim strCnlyClaimNum As String
    Dim strClaimStatus As String
    Dim strQueueType As String
    
    Set mlstClaimList = data
    
    strCnlyClaimNum = ""
    iClaimIDPos = GetColumnPosition(mlstClaimList, "CnlyClaimNum")
    
    For Each varItem In mlstClaimList.ItemsSelected
        strCnlyClaimNum = mlstClaimList.Column(iClaimIDPos, varItem)
        Exit For
    Next
    
    If strCnlyClaimNum <> "" Then
        strClaimStatus = DLookup("ClmStatus", "AUDITCLM_Hdr", "CnlyClaimNum = '" & strCnlyClaimNum & "'")
        strQueueType = DLookup("QueueType", "QUEUE_Hdr", "CnlyClaimNum = '" & strCnlyClaimNum & "'")
        
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.SQLTextType = sqltext
        
        strSQL = "select x.ClmStatus, x.ClmStatusDesc " & _
                " from AUDITCLM_Process_Logics p " & _
                " join XREF_ClaimStatus x on x.ClmStatus = p.NextStatus " & _
                " where p.ProcessModule = 'AuditClm' " & _
                " and p.ProcessType = 'Manual' " & _
                " and p.CurrStatus = '" & strClaimStatus & "' " & _
                " and p.CurrQueue = '" & strQueueType & "' " & _
                " and p.NextStatus <> '" & strClaimStatus & "' "
        
        MyAdo.sqlString = strSQL
        
        Set Me.NewClaimStatus.RecordSet = MyAdo.OpenRecordSet
        
        If Me.NewClaimStatus.RecordSet.recordCount = 0 Then
            MsgBox "There is no processing logic supporting the operation for this queue"
            Call cmdExit_Click
        End If
        Set MyAdo = Nothing
    End If
End Property


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub CmDRun_Click()
    Dim LocCmd As New ADODB.Command
    
    Dim varItem
    Dim strCnlyClaimNum As String
    Dim iClaimIDPos As Integer
    Dim strProcessMsg As String
    
    Dim iTotalCnt As Integer
    Dim iSuccessCnt As Integer
    Dim iErrorCnt As Integer
    Dim iResult As Integer
    
    
    iTotalCnt = 0
    iSuccessCnt = 0
    iErrorCnt = 0
    Me.TotalCnt = iTotalCnt
    Me.SuccessCnt = iSuccessCnt
    Me.ErrorCnt = iErrorCnt
    
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.sqlString = "usp_QUEUE_Mass_Claim_Release"
    
    LocCmd.ActiveConnection = myCode_ADO.CurrentConnection
    LocCmd.commandType = adCmdStoredProc
    LocCmd.CommandText = "usp_QUEUE_Mass_Claim_Release"
    LocCmd.Parameters.Refresh
    
    iClaimIDPos = GetColumnPosition(mlstClaimList, "CnlyClaimNum")
    
    Me.Result.visible = True
    Me.TotalCnt.visible = True
    Me.SuccessCnt.visible = True
    Me.ErrorCnt.visible = True
    
    For Each varItem In mlstClaimList.ItemsSelected
        iTotalCnt = iTotalCnt + 1
        Me.TotalCnt = iTotalCnt
        
        strCnlyClaimNum = mlstClaimList.Column(iClaimIDPos, varItem)
        
        LocCmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
        LocCmd.Parameters("@pNewClaimStatus") = Me.NewClaimStatus
        LocCmd.Parameters("@pClaimNote") = Me.ClaimNote
        LocCmd.Parameters("@pClaimRationale") = Me.Rationale
        LocCmd.Parameters("@pErrMsg") = ""
        
        iResult = myCode_ADO.Execute(LocCmd.Parameters)
        iResult = LocCmd.Parameters("@RETURN_VALUE")
        strProcessMsg = LocCmd.Parameters("@pErrMsg")
        
        If iResult = 0 Then
            iSuccessCnt = iSuccessCnt + 1
            Me.SuccessCnt = iSuccessCnt
        Else
            iErrorCnt = iErrorCnt + 1
            Me.ErrorCnt = iErrorCnt
        End If
        
        If Me.Result = "" Then
            Me.Result = strProcessMsg
        Else
            Me.Result = Me.Result & vbCrLf & strProcessMsg
        End If
        
    Next
    
    mbProcessed = True
    MsgBox "Processing completed"
    
    Set myCode_ADO = Nothing
    Set LocCmd = Nothing
End Sub

Private Sub Form_Close()
    If mbProcessed Then RaiseEvent ClaimProcessed
    
    RaiseEvent FormUnload
    
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    Me.Result.visible = False
    Me.TotalCnt.visible = False
    Me.SuccessCnt.visible = False
    Me.ErrorCnt.visible = False
    Me.Result = ""
End Sub


Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    mstrErrMsg = "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource & vbCrLf & "; ADO Error"
    MsgBox mstrErrMsg
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    mstrErrMsg = "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource & vbCrLf & "; ADO Error"
    MsgBox mstrErrMsg
End Sub

Private Sub NewClaimStatus_AfterUpdate()
    lblClaimStatusDesc.Caption = Me.NewClaimStatus.Column(1)
End Sub
