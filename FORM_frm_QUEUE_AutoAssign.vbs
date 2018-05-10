Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Const CstrFrmAppID As String = "QueueTeamAutoAssign"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Sub RefreshData()

    Dim strSQL As String
    Dim strSelectedGrouping As String
    
    
    strSelectedGrouping = Nz(Me.Combo41, "")
    
    strSQL = " SELECT * from QUEUE_AutoAssign_Exclusions "
    Me.lstExclusions.RowSource = strSQL
    Me.lstExclusions.Requery
    
    strSQL = " SELECT * from QUEUE_AutoAssign_Groups "
    Me.lstQueueAssignGroups.RowSource = strSQL
    Me.lstQueueAssignGroups.Requery
    
    If (strSelectedGrouping = "" Or strSelectedGrouping = "Auditor") Then
        strSQL = " select Auditor, SUM(1) as Ct FROM v_QUEUE_AutoAssign_UnAssignedClaims GROUP BY Auditor "
    Else
        strSQL = " select Auditor," & strSelectedGrouping & ", SUM(1) as Ct FROM v_QUEUE_AutoAssign_UnAssignedClaims GROUP BY Auditor ," & strSelectedGrouping & " order by Auditor, " & strSelectedGrouping & ""
    End If
        
    Me.lstPreviewSummary.RowSource = strSQL
    Me.lstPreviewSummary.Requery
    
    
    
    strSQL = "SELECT * from v_QUEUE_AutoAssign_GroupCapacity"
    Me.frm_QUEUE_AutoAssign_GroupCapacity.Form.RecordSource = strSQL
    Me.frm_QUEUE_AutoAssign_GroupCapacity.Form.Requery
    
    
    '2012-06-05 - Damon Plased in to prevent locking
    'The object is set on the refresh
    Me.frm_QUEUE_AutoAssign_UnAssignedClaims.SourceObject = "frm_QUEUE_AutoAssign_UnAssignedClaims"
    strSQL = " select * FROM v_QUEUE_AutoAssign_UnAssignedClaims "
    Me.frm_QUEUE_AutoAssign_UnAssignedClaims.Form.RecordSource = strSQL
    Me.frm_QUEUE_AutoAssign_UnAssignedClaims.Form.Requery
    

End Sub

Private Sub cmdExecute_Click()
    Dim fmrStatus As Form_ScrStatus
    Dim lngProgressCount As Long
    Dim sMsg As String
    Dim intStatus As Integer
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    
    Dim iActionConfirm As Integer
    
    Dim strErrMsg As String
     
    iActionConfirm = MsgBox("You are about to commit claim assignments.  Do you want to proceed?", vbYesNo)
    If iActionConfirm <> vbYes Then Exit Sub

    Set myCode_ADO = New clsADO
    
    intStatus = 1
    Set fmrStatus = New Form_ScrStatus
    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        .ProgMax = 3
        .TimerInterval = 50
        .show
        .visible = True
    End With
        
            
    sMsg = "Prepping Procedure " & intStatus & " / " & fmrStatus.ProgMax
    fmrStatus.ProgVal = intStatus
    fmrStatus.StatusMessage sMsg
    DoEvents
           
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_QUEUE_AutoAssign_Claims"
    
           
    intStatus = intStatus + 1
    sMsg = "Calling Procedure " & intStatus & " / " & fmrStatus.ProgMax
    fmrStatus.ProgVal = intStatus
    fmrStatus.StatusMessage sMsg
    DoEvents
    
    cmd.CommandTimeout = 0
    cmd.Execute
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    If strErrMsg <> "" Then
       MsgBox "ERROR: " & strErrMsg, vbCritical
    End If
       
    
    intStatus = intStatus + 1
    sMsg = "Refreshing Form " & intStatus & " / " & fmrStatus.ProgMax
    fmrStatus.ProgVal = intStatus
    fmrStatus.StatusMessage sMsg
    DoEvents
    
    DoCmd.Close acForm, fmrStatus.Name
    Set fmrStatus = Nothing
    
    Me.RefreshData
    
    Set cmd = Nothing
    Set myCode_ADO = Nothing
    If strErrMsg = "" Then
       MsgBox "Process completed"
    End If
End Sub

Private Sub cmdPreview_Click()
 
        Dim fmrStatus As Form_ScrStatus
        Dim lngProgressCount As Long
        Dim sMsg As String
        Dim intStatus As Integer
        Dim myCode_ADO As clsADO
        Dim cmd As ADODB.Command
        
        Set myCode_ADO = New clsADO
        
        '2012-06-05 - Damon Plased in to prevent locking
        'The object is set on the refresh
        Me.frm_QUEUE_AutoAssign_UnAssignedClaims.SourceObject = ""
        

        intStatus = 1
        Set fmrStatus = New Form_ScrStatus
        With fmrStatus
            .ShowCancel = True
            .ShowMessage = False
            .ShowMessage = True
            .ProgVal = 0
            .ProgMax = 3
            .TimerInterval = 50
            .show
            .visible = True
        End With
            
                
        sMsg = "Prepping Procedure " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        DoEvents


        intStatus = intStatus + 1
        sMsg = "Calling Procedure " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        DoEvents

        With myCode_ADO
            .ConnectionString = GetConnectString("v_CODE_Database")
            .sqlString = "usp_QUEUE_AutoAssign_Precal_Assignments"
            .SQLTextType = StoredProc
            .Parameters.Refresh
            .CurrentConnection.CommandTimeout = 10000
'DPR REMOVED THIS CALL AND USING TRHE CONNECTION TO CALL MANUALLY TO DEAL WITH A TIMEOUT ISSUE
'            .Execute
'            iResult = .Parameters("@RETURN_VALUE")
        End With
        
        
'MANUALLY CALLING TO MAKE A LONG TIMEOUT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = myCode_ADO.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandTimeout = 1000
        cmd.CommandText = "dbo.usp_QUEUE_AutoAssign_Precal_Assignments"
        
'Stop
'cmd.CommandTimeout = 300
        cmd.Execute
'
'        iResult = cmd.Parameters("@RETURN_VALUE")
        If iResult <> 0 Then
            MsgBox "Error executing [usp_QUEUE_AutoAssign_Precal_Assignments]." & myCode_ADO.Parameters("@pErrMsg")
        End If
        
        
        intStatus = intStatus + 1
        sMsg = "Refreshing Form " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        DoEvents
       
        DoCmd.Close acForm, fmrStatus.Name
        
        Set fmrStatus = Nothing
        
        Me.RefreshData
        
        Me.TabCtl25.Pages.Item(1).SetFocus
        
        Me.cmdExecute.Enabled = True
        
        Set myCode_ADO = Nothing
        Set cmd = Nothing
        
End Sub



Private Sub Combo41_Click()
    RefreshData
End Sub

Private Sub Command19_Click()
        Dim fmrStatus As Form_ScrStatus
        Dim lngProgressCount As Long
        Dim sMsg As String
        Dim intStatus As Integer
        Dim myCode_ADO As clsADO
        
        Set myCode_ADO = New clsADO
        
        intStatus = 1
        Set fmrStatus = New Form_ScrStatus
        With fmrStatus
            .ShowCancel = True
            .ShowMessage = False
            .ShowMessage = True
            .ProgVal = 0
            .ProgMax = 3
            .TimerInterval = 50
            .show
            .visible = True
        End With
                
        sMsg = "Prepping Procedure " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        intStatus = intStatus + 1
        DoEvents
        
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.SQLTextType = sqltext
        myCode_ADO.sqlString = "Usp_QUEUE_AutoAssign_Cal_Productivity"
        
        sMsg = "Prepping Procedure " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        intStatus = intStatus + 1
        DoEvents
        
        iResult = myCode_ADO.Execute
        
        sMsg = "Called Procedure " & intStatus & " / " & fmrStatus.ProgMax
        fmrStatus.ProgVal = intStatus
        fmrStatus.StatusMessage sMsg
        intStatus = intStatus + 1
        
        Set fmrStatus = Nothing
        Set fmrStatus = Nothing
        Me.RefreshData

End Sub
Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Action Code"
    
    iAppPermission = UserAccess_Check(Me)
    
    Me.Combo41 = "Auditor"
    RefreshData
End Sub
Private Sub lstExclusions_Click()
    Dim strSQL As String
    
    Me.frm_QUEUE_AutoAssign_Exclusions.Form.RecordSource = " select * from QUEUE_AutoAssign_exclusions where exclusiontype = '" & Me.lstExclusions & "' AND ExclusionValue = '" & Me.lstExclusions.Column(1) & "'"


End Sub

Private Sub lstQueueAssignGroups_Click()
    Dim strSQL As String
    Me.frm_QUEUE_AutoAssign_Groups.Form.RecordSource = " select * from QUEUE_AutoAssign_Groups where GroupName = '" & Me.lstQueueAssignGroups & "'"

End Sub
