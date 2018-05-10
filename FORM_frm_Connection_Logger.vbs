Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const cs_CONN_STRING_PTRN As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=[SERVER];Initial Catalog=[DATABASE];Application Name=ConnectionTest;"

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Private Sub cmdStartLogging_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdStartLogging_Click"
    
    
    
    If Me.cmdStartLogging.Caption = "Start Logging" Then
        Me.cmdStartLogging.Caption = "Stop logging"
        Me.Detail.BackColor = RGB(0, 255, 0)
        Me.TimerInterval = Me.txtSecondsForLoop * 1000
    Else
        Me.cmdStartLogging.Caption = "Start Logging"
        Me.Detail.BackColor = 16777215
        Me.TimerInterval = 0
        
        Me.lblProdStatus.BackColor = Me.Detail.BackColor
    
        Me.lblDevStatus.BackColor = Me.Detail.BackColor
    End If
    
    
    
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Command5_Click()
    Debug.Print Me.Detail.BackColor
Stop
End Sub

Private Sub Command10_Click()
    Debug.Print Me.InsideHeight
    Debug.Print Me.InsideWidth
    
Stop
End Sub

Private Sub Form_Resize()
    
    Me.InsideHeight = 4605
    Me.InsideWidth = 4800

End Sub

Private Sub Form_Timer()
Dim lCurrentInterval As Long

    ' stop our timer
    lCurrentInterval = Me.TimerInterval
    Me.TimerInterval = 0
    
    Call TestConnection("DS-FLD-009")
    Call TestConnection("DC-BIGSKY")
    
    ' turn it back on:
    Me.TimerInterval = lCurrentInterval
    
End Sub


Public Function TestConnection(strServer As String, Optional strDb As String, Optional strSQL As String = "SELECT TOP 1 * FROM AUDITCLM_Hdr") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim sConnString As String
Dim bSuccess As Boolean
Dim sErrMsg As String
Dim iCurSetting As Integer

    strProcName = ClassName & ".TestConnection"
    
    If strDb = "" Then
        strDb = "CMS_AUDITORS_CLAIMS"
    End If
    
    
    sConnString = Replace(cs_CONN_STRING_PTRN, "[SERVER]", strServer)
    sConnString = Replace(sConnString, "[DATABASE]", strDb)
    
    bSuccess = True ' optimistic

    iCurSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = sConnString
        .CursorLocation = adUseNone
        .ConnectionTimeout = Nz(Me.txtTimeoutSeconds, 5)
        .CommandTimeout = Nz(Me.txtTimeoutSeconds, 5)
        On Error Resume Next
        
        .Open
        
        If Err.Number <> 0 Then
            sErrMsg = Err.Description
            bSuccess = False
            Err.Clear
        End If
        
        On Error GoTo Block_Err
    End With
    
    Call LogResults(strServer, bSuccess, sErrMsg, "CONNECT", CLng(Nz(Me.txtTimeoutSeconds, 5)))
    
    If bSuccess = False Then GoTo Block_Exit
    
    
    Set oRs = New ADODB.RecordSet
    With oRs
        .CursorLocation = adUseNone
        .LockType = adLockReadOnly
        Set .ActiveConnection = oCn

        On Error Resume Next
        
        .Open (strSQL)
        
        If Err.Number <> 0 Then
            sErrMsg = Err.Description
            bSuccess = False
            Err.Clear
        End If
        
        On Error GoTo Block_Err

    End With
    
    Call LogResults(strServer, bSuccess, sErrMsg, "SELECT", CLng(Nz(Me.txtTimeoutSeconds, 5)), strSQL)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Application.SetOption "Error Trapping", iCurSetting
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function LogResults(sToServer As String, bSuccess As Boolean, sErrMsg As String, sOperationPerformed As String, lTimeout As Long, Optional sCommand As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim lColor As Long

    strProcName = ClassName & ".LogResults"
    If bSuccess = False Then
        lColor = RGB(255, 0, 0)
    Else
        lColor = RGB(0, 255, 0)
    End If
    
    Select Case UCase(sToServer)
    Case "DS-FLD-009"
        Me.lblProdStatus.BackColor = lColor
    Case Else
        Me.lblDevStatus.BackColor = lColor
    End Select
    
    
    Set oDb = CurrentDb()
    
    Set oRs = oDb.OpenRecordSet("SELECT * FROM tbl_ConnectionLogging WHERE 1 = 2")
    
    oRs.AddNew
    oRs("LogTime").Value = Now()
    oRs("FromServer") = GetPCName
    oRs("ToServer") = sToServer
    oRs("Success").Value = bSuccess
    oRs("ErrorMsg").Value = sErrMsg
    oRs("OperationPerformed").Value = sOperationPerformed
    oRs("MaxTimeoutAllowed").Value = lTimeout
    If sCommand <> "" Then
        oRs("CommandIssued").Value = sCommand
    End If
    oRs.Update
    
    Me.txtLogList = Now() & " " & sToServer & " " & sOperationPerformed & " " & IIf(bSuccess, "Success", "Fail") & vbCrLf & Me.txtLogList
    ' trim it up
    Me.txtLogList = left(Me.txtLogList, 2000)
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
