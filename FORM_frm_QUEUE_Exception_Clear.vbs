Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private mstrFormFilter As String
Private mrsException As ADODB.RecordSet
Private mstrErrMsg As String
Private mstrErrSource As String

Property Let FormFilter(data As String)
    mstrFormFilter = data
End Property

Private Sub cmdClearException_Click()
'12/17/2013 MG change variable name from rsExceptionValidation to rs because this rs can reference multiple fields, not just ValidationFlag

    Dim rs As ADODB.RecordSet
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select ValidationFlag,IsReadOnly from QUEUE_Xref_Type where QueueType = '" & Me.ExceptionType & "'"
    Set rs = MyAdo.OpenRecordSet
    
    mstrErrSource = "cmdClearException_Click"
    
    If MsgBox("Please make sure to save your record before proceeding" & vbCrLf & vbCrLf & "Do you want to proceed?", vbYesNo + vbInformation) = vbYes Then
    
        'MG do the following if it's marked as READ ONLY
        'JS 20140304 I added the readonly flag check to [usp_QUEUE_Exception_Delete] so it does not get done here in Access
        
'        If rs("IsReadOnly") = "1" Then
'            MsgBox "Exception " & Me.ExceptionType & " is read-only. If you wish to clear it out, please contact Mike Guan.", vbCritical
'        Else
            Select Case UCase(Me.ExceptionType)
                Case "EX001"                        ' Additional medical record information received
                    Clear_Exception_InfoOnly
                Case "EX002"                        ' Provider Address Change Form Received
                    Clear_Exception_InfoOnly
                Case "EX003"                        ' Provider does not exists
                    Clear_Exception_EX003
                Case "EX004"                        ' Provider Medical Record Address does not exist
                    Clear_Exception_EX004
                Case "EX005"                        ' Provider Finance Address does not exist
                    Clear_Exception_EX005
                Case "EX006"
                    Clear_Exception_InfoOnly
                Case "EX007"
                    Clear_Exception_InfoOnly
                Case "EX008"
                    Clear_Exception_InfoOnly
                Case "EX051"
                    'Clear_Exception_EX051
                    'DPR Changed to prevent creation of EX060
                    Clear_Exception_InfoOnly
                Case Else
                    If rs("ValidationFlag") & "" <> "Y" Then
                        Clear_Exception_InfoOnly
                    Else
                        MsgBox "ERROR: There is no logic defined to handle exception type: '" & Me.ExceptionType & "'.  Please notify IT.", vbCritical
                    End If
                    Exit Sub
            End Select
            
            
'        End If
    End If
    
End Sub
Private Sub Clear_Exception_EX051()

    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As Variant
    Dim ErrMsg As String
    Dim bResult As Boolean
    On Error GoTo Err_handler
    Set MyCodeAdo = New clsADO
        
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    'Check if annotated record is there?
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_QUEUE_Exception_apply"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    cmd.Parameters("@pExceptionType") = "EX060"
    cmd.Parameters("@pExceptionStatus") = "OPEN"
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pLastUpdate") = Now()
    cmd.Parameters("@pUpdateUser") = Identity.UserName
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    
    mstrErrSource = "Clear_Exception_" & mrsException("QueueType")
   
    If spReturnVal <> 0 Then
        ErrMsg = cmd.Parameters("@pErrMsg")
        Err.Raise 65000, mstrErrSource, ErrMsg
    End If
        
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.BeginTrans
    bResult = myCode_ADO.Update(mrsException, "usp_QUEUE_Exception_Clear")
    If bResult = False Then
        GoTo Err_handler
    End If
    
    myCode_ADO.CommitTrans
    
    MsgBox "Exception cleared!", vbInformation
    DoCmd.Close acForm, Me.Name

Exit_Sub:
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
    If mstrErrMsg = "" Then
        mstrErrMsg = Err.Description
        myCode_ADO.RollbackTrans
        Err.Raise Err.Number, mstrErrSource, mstrErrMsg
    Else
        MsgBox mstrErrMsg, vbCritical
        myCode_ADO.RollbackTrans
    End If
    mstrErrMsg = ""
    GoTo Exit_Sub
End Sub
Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub

Public Sub RefreshData()
    If mstrFormFilter = "" Then
        MsgBox "Error: Where condition is not set.  Please notify IT."
        DoCmd.Close acForm, Me.Name
    End If
    
    'TL add account ID logic
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "select * from v_QUEUE_Exception_Info where AccountID = " & gintAccountID & " and " & mstrFormFilter
    Set mrsException = myCode_ADO.OpenRecordSet
    If Not (mrsException.BOF And mrsException.EOF) Then
        If mrsException.recordCount <> 1 Then
            MsgBox "Error: Selection process returned more than 1 record.  Please notify IT."
            DoCmd.Close acForm, Me.Name
        Else
            Set Me.RecordSet = mrsException
        End If
    Else
        MsgBox "Error:  There is no record to edit"
        DoCmd.Close acForm, Me.Name
    End If
    
    Set myCode_ADO = Nothing
End Sub

' clear info only exception
Private Sub Clear_Exception_InfoOnly()
    Dim bResult As Boolean
    
    On Error GoTo Err_handler
    
    mstrErrSource = "Clear_Exception_" & mrsException("QueueType")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.BeginTrans
    bResult = myCode_ADO.Update(mrsException, "usp_QUEUE_Exception_Clear")
    If bResult = False Then
        GoTo Err_handler
    End If
    
    myCode_ADO.CommitTrans
    MsgBox "Exception cleared!", vbInformation
    
    DoCmd.Close acForm, Me.Name

Exit_Sub:
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
    If mstrErrMsg = "" Then
        mstrErrMsg = Err.Description
        myCode_ADO.RollbackTrans
        Err.Raise Err.Number, mstrErrSource, mstrErrMsg
    Else
        MsgBox mstrErrMsg, vbCritical
        myCode_ADO.RollbackTrans
    End If
    mstrErrMsg = ""
    GoTo Exit_Sub
End Sub


' Provider does not exists
Private Sub Clear_Exception_EX003()
    Dim bResult As Boolean
    Dim rs As ADODB.RecordSet
    
    On Error GoTo Err_handler
    
    mstrErrSource = "Clear_Exception_EX003"
    
    ' check to see if provider exists
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from PROV_Hdr where CnlyProvID = '" & Me.cnlyProvID & "'"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.BOF And rs.EOF Then
        'provider does not exists
        MsgBox "Provider " & Me.cnlyProvID & " does not exists. " & vbCrLf & _
               "Please create this provider first before attempt to clear this exception", vbCritical
        Exit Sub
    Else
        'provider exists
        'clear all exceptions associated with this provider
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.sqlString = "select * from v_QUEUE_Exception_Info where CnlyProvID = '" & Me.cnlyProvID & "' and ExceptionType = '" & Me.ExceptionType & "'"
        Set rs = myCode_ADO.OpenRecordSet
        
        myCode_ADO.BeginTrans
        
        bResult = myCode_ADO.Update(rs, "usp_QUEUE_Exception_Clear")
        If bResult = False Then
            GoTo Err_handler
        End If
    
        myCode_ADO.CommitTrans
        MsgBox "Exception cleared!", vbInformation
        
        Set rs = Nothing
        DoCmd.Close acForm, Me.Name
    End If
    
Exit_Sub:
    Set rs = Nothing
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
    If mstrErrMsg = "" Then
        mstrErrMsg = Err.Description
        myCode_ADO.RollbackTrans
        Err.Raise Err.Number, mstrErrSource, mstrErrMsg
    Else
        MsgBox mstrErrMsg, vbCritical
        myCode_ADO.RollbackTrans
    End If
    mstrErrMsg = ""
End Sub


'Provider Medical Record Address does not exist
Private Sub Clear_Exception_EX004()
    Dim bResult As Boolean
    Dim rs As ADODB.RecordSet
    
    On Error GoTo Err_handler
    
    mstrErrSource = "Clear_Exception_EX004"
    
    'select to see if provider address exists
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from PROV_Address where AddrType = '01' " & _
                      " and getdate() between EffDt and isnull(TermDt,'12/31/9999')" & _
                      " and CnlyProvID = '" & Me.cnlyProvID & "'"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.BOF And rs.EOF Then
        'provider address does not exists
        MsgBox "There is NO effective MEDICAL ADDRESS for provider " & Me.cnlyProvID & vbCrLf & vbCrLf & _
                " Please create the address first before attempt to clear this exception", vbCritical
        Exit Sub
    Else
        'provider address does exists
        
        
        'clear all exceptions associated with this provider
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.sqlString = "select * from v_QUEUE_Exception_Info where CnlyProvID = '" & Me.cnlyProvID & "' and ExceptionType = '" & Me.ExceptionType & "'"
        Set rs = myCode_ADO.OpenRecordSet
        
        myCode_ADO.BeginTrans
        bResult = myCode_ADO.Update(rs, "usp_QUEUE_Exception_Clear")
        If bResult = False Then
            GoTo Err_handler
        End If
    
        myCode_ADO.CommitTrans
        MsgBox "Exception cleared!", vbInformation
        
        Set rs = Nothing
        DoCmd.Close acForm, Me.Name
    End If

Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Exit Sub

Err_handler:
    If mstrErrMsg = "" Then
        mstrErrMsg = Err.Description
        myCode_ADO.RollbackTrans
        Err.Raise Err.Number, mstrErrSource, mstrErrMsg
    Else
        MsgBox mstrErrMsg, vbCritical
        myCode_ADO.RollbackTrans
    End If
    mstrErrMsg = ""
    GoTo Exit_Sub

End Sub


'Provider Finance Address does not exist
Private Sub Clear_Exception_EX005()
    Dim bResult As Boolean
    Dim rs As ADODB.RecordSet
    
    On Error GoTo Err_handler
    
    mstrErrSource = "Clear_Exception_EX005"
    
    'select to see if provider address exists
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from PROV_Address where AddrType = '02' " & _
                      " and getdate() between EffDt and isnull(TermDt,'12/31/9999')" & _
                      " and CnlyProvID = '" & Me.cnlyProvID & "'"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.BOF And rs.EOF Then
        'provider address does not exists
        MsgBox "There is NO effective FINANCE ADDRESS for provider " & Me.cnlyProvID & vbCrLf & vbCrLf & _
                " Please create the address first before attempt to clear this exception", vbCritical
        Exit Sub
    Else
        'provider address does exists
        
        'clear all exceptions associated with this provider
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.sqlString = "select * from v_QUEUE_Exception_Info where CnlyProvID = '" & Me.cnlyProvID & "' and ExceptionType = '" & Me.ExceptionType & "'"
        Set rs = myCode_ADO.OpenRecordSet
        
        myCode_ADO.BeginTrans
        bResult = myCode_ADO.Update(rs, "usp_QUEUE_Exception_Clear")
        If bResult = False Then
            GoTo Err_handler
        End If
    
        myCode_ADO.CommitTrans
        MsgBox "Exception cleared!", vbInformation
        
        Set rs = Nothing
        DoCmd.Close acForm, Me.Name
    End If

Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Exit Sub

Err_handler:
    If mstrErrMsg = "" Then
        mstrErrMsg = Err.Description
        myCode_ADO.RollbackTrans
        Err.Raise Err.Number, mstrErrSource, mstrErrMsg
    Else
        MsgBox mstrErrMsg, vbCritical
        myCode_ADO.RollbackTrans
    End If
    mstrErrMsg = ""
    GoTo Exit_Sub

End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    mstrErrMsg = "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource & vbCrLf & vbCrLf & mstrErrSource & ": ADO Error"
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    mstrErrMsg = "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource & vbCrLf & vbCrLf & mstrErrSource & ": ADO Error"
End Sub
