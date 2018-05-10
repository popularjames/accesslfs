Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class provide a disconnected recordset back to the calling routine
Option Compare Database
Option Explicit

Public Event ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)


Private mConnectionString As String
Private mSQLType As clsAdoSQLType
Private mSQLString As String
Private mCN As ADODB.Connection
Private mCmd As ADODB.Command
Private mRS As ADODB.RecordSet
Private mParam As ADODB.Parameter
Private mInTransaction As Boolean

'* JC renamed from SqlType

Public Enum clsAdoSQLType
    StoredProc = adCmdStoredProc
   ' Table = adCmdTable
    sqltext = adCmdText
End Enum

Private Enum SQLDataType
    adoDateTime = 135
    adoFloat = 5
    adoReal = 4
    adomoney = 6
    adoBigInt = 20
    adoInt = 3
    adoSmallInt = 2
    adoTinyInt = 17
    adoNText = 203
    adoText = 201
    adoImage = 205
    adoTimeStamp = 128
    adoUniqueidentifier = 72
    adoNVarchar = 202
    adoNChar = 130
    adoVarchar = 200
    adoChar = 129
    adoVarBinary = 204
    adoBinary = 128
    adoNumeric = 131
End Enum

'' KD 20120416

Private bGotData As Boolean
Private bRefreshedPs As Boolean


Public Property Get GotData() As Boolean
    GotData = bGotData
End Property

Public Property Get cmd() As ADODB.Command
    Call InsureCmdIsSetup
    Set cmd = mCmd
End Property

Public Property Get Parameters() As ADODB.Parameters
    If mCmd Is Nothing Then
        Call InsureCmdIsSetup
    End If
    Set Parameters = mCmd.Parameters
End Property


Private Sub InsureCmdIsSetup()
    If mCmd Is Nothing Then
        Set mCmd = New ADODB.Command
        bRefreshedPs = False
    End If
    With mCmd
        .commandType = Me.SQLTextType
        .CommandText = Me.sqlString
        If Not Me.CurrentConnection Is Nothing Then
            If .ActiveConnection Is Nothing Then
                If Me.CurrentConnection.State = adStateOpen Then
                    .ActiveConnection = Me.CurrentConnection
                Else
                    Me.CurrentConnection.Open
                    
                    .ActiveConnection = Me.CurrentConnection
                End If
            End If
        End If
        If bRefreshedPs = False Then
            .Parameters.Refresh
            bRefreshedPs = True
        End If
        
    End With
End Sub


'' End / KD 20120416


Public Sub BatchUpdate(ByVal rs As ADODB.RecordSet)
    Dim strErrSource As String
    strErrSource = "clsADO.BatchUpdate" ' KD 20120416 - this must've been copied / pasted from OpenRecordset
    
    On Error GoTo Err_handler
    
    If Check_Connection Then
        rs.ActiveConnection = mCN
        rs.UpdateBatch
    End If

Exit_Function:
    Exit Sub

Err_handler:
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
End Sub

Public Function Update(ByVal rs As ADODB.RecordSet, StoredProc As String) As Boolean
    
    Dim fld As Field
    Dim i As Integer
    Dim iParamDatatype As Integer
    Dim iRtnCd As Integer
    Dim strErrSource As String
    Dim strErrMsg As String
    
    strErrSource = "clsADO.Update"
    strErrMsg = ""
    
    On Error GoTo Err_handler
    
    If Check_Connection Then
        ' get stored procedure parameters
        Set mCmd = New ADODB.Command    '' KD 20120416
        mCmd.ActiveConnection = mCN
        mCmd.commandType = adCmdStoredProc
        mCmd.CommandText = StoredProc
        mCmd.Parameters.Refresh
        ' update data
        rs.MoveFirst
        While Not rs.EOF
            ' set stored proc parameter values
            For i = 1 To mCmd.Parameters.Count - 2
                If isField(rs, Mid$(mCmd.Parameters(i).Name, 3)) Then
                    mCmd.Parameters(i).Value = Trim(rs(Mid$(mCmd.Parameters(i).Name, 3)))
                End If
'Debug.Print mCmd.Parameters(i).Name & "='" & rs(Mid$(mCmd.Parameters(i).Name, 3)) & "'"
            Next i
            mCmd.Execute
            
            iRtnCd = mCmd.Parameters("@RETURN_VALUE")
            If iRtnCd <> 0 Then
                strErrMsg = mCmd.Parameters("@pErrMsg")
                GoTo Err_handler
            End If
            mCmd.Parameters.Refresh
        
            rs.MoveNext
        Wend
    End If

Exit_Function:
    Update = True
    Exit Function

Err_handler:
    Update = False
    If strErrMsg <> "" Then
        RaiseEvent ADOError(strErrMsg, vbObjectError + 513, strErrSource)
    Else
        RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    End If
End Function


Public Function CopyRecordset(ByVal rs As ADODB.RecordSet, Optional rsBookMark As Variant) As ADODB.RecordSet
Dim strErrSource As String
Dim rsTemp As ADODB.RecordSet
Dim fld As ADODB.Field
Dim i As Integer
        
    strErrSource = "clsADO.CopyRecordSet"
    
    On Error GoTo Err_handler
    
    Set CopyRecordset = rs.Clone
    GoTo Exit_Function
    
    Set rsTemp = CreateObject("ADODB.Recordset")
    rsTemp.CursorType = adOpenKeyset
    rsTemp.CursorLocation = rs.CursorLocation
    rsTemp.Source = rs.Source
    rsTemp.DataMember = rs.DataMember
    rsTemp.index = rs.index
    
    For Each fld In rs.Fields
        rsTemp.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Next
    
    rsTemp.Open
    rs.MoveFirst
    While Not rs.EOF
        rsTemp.AddNew
        For i = 0 To rs.Fields.Count - 1
           rsTemp(i).Value = rs(i).Value
        Next i
        rs.MoveNext
    Wend
    Set CopyRecordset = rsTemp

Exit_Function:
    Exit Function

Err_handler:
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
End Function

Public Function OpenRecordSet(Optional SQLCmd As String = "", Optional DisConnectedRS As Boolean = True) As ADODB.RecordSet
    Dim strErrSource As String
    strErrSource = "clsADO.OpenRecordSet"
    
    On Error GoTo Err_handler
    bGotData = False
    
    
    '8/17/2012 - DPR RUNNING ASYNC TO TEST THROUGHPUT
    
    
    If SQLCmd <> "" Then mSQLString = SQLCmd
    
    If Check_Connection Then
        Set mRS = New ADODB.RecordSet   '' 20120416 KD early bound
        mRS.CursorLocation = adUseClient
        
        'DPR ASYNC ADD
        'mRS.Open mSQLString, mCN, adOpenStatic, adLockBatchOptimistic, adAsyncExecute
        'DPR REMOVED
        mRS.Open mSQLString, mCN, adOpenStatic, adLockBatchOptimistic
        'DPR REMOVED
        'mRS.Open mSQLString, mCN, adOpenStatic, adLockBatchOptimistic, adLockReadOnly
        
        
        'DPR ASYNC ADD Wait for recordset to finish fetching
'        Do While mRS.State <> adStateOpen
'            Sleep 20
'        Loop
'
        
        If DisConnectedRS Then
            mRS.ActiveConnection = Nothing
        End If
        Set OpenRecordSet = mRS
            '' 20120416 KD Added Got Data property
        If Not mRS Is Nothing Then
            If mRS.recordCount > 0 Then
                bGotData = True
            End If
        End If
        
        'DPR -2012 - 9 - 16
        ' KD: CAN'T Close the connection because any recordsets that use this are killed.. 9/11/2012
'        On Error Resume Next
'        mCN.Close
        'mRS.Close
        Set mRS = Nothing
    End If

Exit_Function:
    Exit Function

Err_handler:
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
End Function


'' -- 20120416 KD, extended to use the mCmd params
Public Function Execute(Optional ByRef Params As ADODB.Parameters, Optional strErrMsg As String) As Long
Dim lRecordAffected As Long
Dim strErrSource As String
Dim Param As ADODB.Parameter
Dim i As Integer
Dim bUsePassedParams As Boolean

    strErrSource = "clsADO.Execute"
    
    On Error GoTo Err_handler
    bGotData = False    '' KD 20120416 - added property to class
    
    If mSQLType = 0 Then
        strErrMsg = "Error: SQL type is not defined!"
        GoTo Err_handler
    End If
    
    If mSQLString = "" Then
        strErrMsg = "Error: SQL text is not defined!"
        GoTo Err_handler
    End If
    
    
    
    '' 20120416: KD added mCmd property.
    '' so, if user doesn't pass Params, then check to se
    '' if the mCmd object has any set..
    If Not Params Is Nothing Then
        'Set Params = mCmd.Parameters
        bUsePassedParams = True
        Set mCmd = New ADODB.Command
        mCmd.ActiveConnection = mCN
        mCmd.CommandText = mSQLString
        mCmd.commandType = mSQLType
    Else
        Set mCmd = Me.cmd
    End If
   
    '' 20120416 KD - changed to use mCmd, OR the params passed in
    If Check_Connection Then
        If bUsePassedParams = True Then
            If Not (Params Is Nothing) Then
                For Each Param In Params
                    mCmd.Parameters.Append mCmd.CreateParameter(Param.Name, Param.Type, Param.Direction, Param.Size, Param.Value)
                    If mCmd.Parameters(mCmd.Parameters.Count - 1).Type = adNumeric Then
                        mCmd.Parameters(mCmd.Parameters.Count - 1).Precision = Param.Precision
                        mCmd.Parameters(mCmd.Parameters.Count - 1).NumericScale = Param.NumericScale
                    End If
    'Debug.Print Param.Name & " -- " & mCmd.Parameters(Param.Name).Value
                Next
            End If
        End If
        mCmd.Execute lRecordAffected
    
        If bUsePassedParams = True Then
            If Not (Params Is Nothing) Then
                For i = 0 To Params.Count - 1
                    Params(i).Value = mCmd.Parameters(i).Value
                Next
            End If
        End If
    End If
    '' 20120416 KD / end changes..
    
    
                ''      KD: Legacy code:
                '    If Check_Connection Then
                '        Set mCmd = CreateObject("ADODB.Command")
                '        mCmd.ActiveConnection = mCN
                '        mCmd.CommandText = mSQLString
                '        mCmd.CommandType = mSQLType
                '        If Not (Params Is Nothing) Then
                '            For Each Param In Params
                '                mCmd.Parameters.Append mCmd.CreateParameter(Param.Name, Param.Type, Param.Direction, Param.Size, Param.Value)
                '                If mCmd.Parameters(mCmd.Parameters.Count - 1).Type = adNumeric Then
                '                    mCmd.Parameters(mCmd.Parameters.Count - 1).Precision = Param.Precision
                '                    mCmd.Parameters(mCmd.Parameters.Count - 1).NumericScale = Param.NumericScale
                '                End If
                ''Debug.Print Param.Name & " -- " & mCmd.Parameters(Param.Name).Value
                '            Next
                '        End If
                '        mCmd.Execute lRecordAffected
                '
                '        If Not (Params Is Nothing) Then
                '            For i = 0 To Params.Count - 1
                '                Params(i).Value = mCmd.Parameters(i).Value
                '            Next
                '        End If
                '    End If
    
    
    bGotData = IIf(lRecordAffected > 0, True, False)    '' 20120416 KD Added got data prop
    
Exit_Function:

    Execute = lRecordAffected

''    Set mCmd = Nothing '' 20120416 KD we want this command now.
    Exit Function

Err_handler:
    lRecordAffected = -1
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function


Public Function ExecuteRS(Optional Params As ADODB.Parameters, Optional strErrMsg As String) As ADODB.RecordSet
Dim lRecordAffected As Long
Dim strErrSource As String
Dim Param As ADODB.Parameter
Dim i As Integer
Dim bUsePassedParams As Boolean

    strErrSource = "clsADO.ExecuteRS"   '' KD 20120416 copy / pasted...
    bGotData = False    '' KD 20120416 gotdata property
    
    On Error GoTo Err_handler
    
    If mSQLType = 0 Then
        strErrMsg = "Error: SQL type is not defined!"
        GoTo Err_handler
    End If
    
    If mSQLString = "" Then
        strErrMsg = "Error: SQL text is not defined!"
        GoTo Err_handler
    End If
    
    '' 20120327: KD added mCmd property.
    '' so, if user doesn't pass Params, then check to see
    '' if the mCmd object has any set..
    If Not Params Is Nothing Then
        'Set Params = mCmd.Parameters
        bUsePassedParams = True
        Set mCmd = New ADODB.Command
        mCmd.ActiveConnection = mCN
        mCmd.CommandText = mSQLString
        mCmd.commandType = mSQLType
    Else
        Set mCmd = Me.cmd
    End If
    
    
    If Check_Connection Then

        If bUsePassedParams = True Then
            If Not (Params Is Nothing) Then
                For Each Param In Params
                    mCmd.Parameters.Append mCmd.CreateParameter(Param.Name, Param.Type, Param.Direction, Param.Size, Param.Value)
                Next
            End If
        End If
        
        Set mRS = mCmd.Execute
        If bUsePassedParams = True Then
            If Not (Params Is Nothing) Then
                For i = 0 To Params.Count - 1
                    Params(i).Value = mCmd.Parameters(i).Value
                Next
            End If
        End If
    End If
            '' 20120416 end KD see Execute method
            '''    '' LEGACY CODE:
            '''    If Check_Connection Then
            '''        Set mCmd = CreateObject("ADODB.Command")
            '''        mCmd.ActiveConnection = mCN
            '''        mCmd.CommandText = mSQLString
            '''        mCmd.CommandType = mSQLType
            '''        If Not (Params Is Nothing) Then
            '''            For Each Param In Params
            '''                mCmd.Parameters.Append mCmd.CreateParameter(Param.Name, Param.Type, Param.Direction, Param.Size, Param.Value)
            '''            Next
            '''        End If
            '''        Set mRS = mCmd.Execute
            '''
            '''        If Not (Params Is Nothing) Then
            '''            For i = 0 To Params.Count - 1
            '''                Params(i).Value = mCmd.Parameters(i).Value
            '''            Next
            '''        End If
            '''    End If
    
    Set ExecuteRS = mRS

        ''' 20120416 KD: for GotData property:
    If Not mRS Is Nothing Then
        If mRS.State = adStateOpen Then
            If mRS.recordCount > 0 Then
                bGotData = True
            End If
        End If
    End If
    
Exit_Function:
'    Set mCmd = Nothing     ''' 20120416 KD
    Set mRS = Nothing
    Exit Function

Err_handler:
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    GoTo Exit_Function
End Function


Public Function Connect() As Boolean
    Dim strErrSource As String
    Dim strErrMsg As String
    
    strErrSource = "clsADO.Connect"
    
    On Error GoTo Err_handler
    Set mCmd = Nothing  '' 20120416 KD made mCmd more useful
    
    If mInTransaction Then
        strErrMsg = "Pending transaction exists.  Please rollback or commit transaction first"
        GoTo Err_handler
    End If
    
    If mConnectionString = "" Then
        strErrMsg = "Connection string is blank."
        GoTo Err_handler
    End If
    
    If mCN Is Nothing Then
        Set mCN = New ADODB.Connection
    End If
    
    If mCN.State = adStateOpen Then
        mCN.Close
    End If
    
    mCN.ConnectionString = mConnectionString
    mCN.CursorLocation = adUseClient
    mCN.CommandTimeout = 180    ' 2 minutes by default..
    mCN.Open
    
Exit_Function:
    Connect = True
    Exit Function

Err_handler:
    Connect = False
    If strErrMsg <> "" Then
        RaiseEvent ADOError(strErrMsg, vbObjectError + 513, strErrSource)
    Else
        RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    End If
End Function


Public Function DisConnect() As Boolean
    Dim strErrSource As String
    strErrSource = "clsADO.DisConnect"
    
    On Error GoTo Err_handler
    
    If mCN Is Nothing Then
        GoTo Exit_Function
    End If
    
    If mCN.State = adStateOpen Then
        mCN.Close
    End If
    

Exit_Function:
    DisConnect = True
    Exit Function

Err_handler:
    DisConnect = False
    RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
End Function

Public Function BeginTrans() As Boolean
    Dim strErrSource As String
    Dim strErrMsg As String

    strErrSource = "clsADO.BeginTran"
    
    On Error GoTo Err_handler
    
    If mCN Is Nothing Then
        strErrMsg = "Cannot start transaction.  Connection object is not defined."
        GoTo Err_handler
    End If
    
    If mCN.State = adStateClosed Then
        strErrMsg = "Cannot start transaction.  Connection object is closed."
        GoTo Err_handler
    End If
    
    If mInTransaction Then
        strErrMsg = "Cannot start transaction.  Pending transaction exists."
        GoTo Err_handler
    End If
    
    mCN.BeginTrans
    mInTransaction = True
    
Exit_Function:
    BeginTrans = True
    Exit Function

Err_handler:
    BeginTrans = False
    If strErrMsg <> "" Then
        RaiseEvent ADOError(strErrMsg, vbObjectError + 513, strErrSource)
    Else
        RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    End If
End Function

Public Function RollbackTrans() As Boolean
    Dim strErrSource As String
    Dim strErrMsg As String

    strErrSource = "clsADO.RollbackTrans"
    
    On Error GoTo Err_handler
    
    If mCN Is Nothing Then
        strErrMsg = "Cannot rollback transaction.  Connection object is not defined."
        GoTo Err_handler
    End If
    
    If mCN.State = adStateClosed Then
        strErrMsg = "Cannot rollback transaction.  Connection object is closed."
        GoTo Err_handler
    End If
    
    If mInTransaction Then
        mCN.RollbackTrans
        mInTransaction = False
    End If
    
    
Exit_Function:
    RollbackTrans = True
    Exit Function

Err_handler:
    RollbackTrans = False
    If strErrMsg <> "" Then
        RaiseEvent ADOError(strErrMsg, vbObjectError + 513, strErrSource)
    Else
        RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    End If
End Function

Public Function CommitTrans() As Boolean
    Dim strErrSource As String
    Dim strErrMsg As String

    strErrSource = "clsADO.CommitTrans"
    
    On Error GoTo Err_handler
    
    If mCN Is Nothing Then
        strErrMsg = "Cannot commit transaction.  Connection object is not defined."
        GoTo Err_handler
    End If
    
    If mCN.State = adStateClosed Then
        strErrMsg = "Cannot commit transaction.  Connection object is closed."
        GoTo Err_handler
    End If
    
    If mInTransaction Then
        mCN.CommitTrans
        mInTransaction = False
    End If
    
    
Exit_Function:
    CommitTrans = True
    Exit Function

Err_handler:
    CommitTrans = False
    If strErrMsg <> "" Then
        RaiseEvent ADOError(strErrMsg, vbObjectError + 513, strErrSource)
    Else
        RaiseEvent ADOError(Err.Description, Err.Number, strErrSource)
    End If
End Function

Private Function Check_Connection() As Boolean
    Check_Connection = True
    
    If mCN Is Nothing Then
        Check_Connection = Me.Connect
    End If
    
    If mCN.State = adStateClosed Then
        Check_Connection = Me.Connect
    End If
End Function




Public Property Let sqlString(ByVal vData As String)
    mSQLString = vData
    bGotData = False    ' 20120416  KD:
    Set mCmd = Nothing  '20120416  KD: made mCmd more useful
End Property

Public Property Get sqlString() As String
    sqlString = mSQLString
End Property

Public Property Let ConnectionString(ByVal vData As String)
    ' only reset the connection string if it a new connection string
    Set mCmd = Nothing  '20120416  KD: made mCmd more useful
    
    If mConnectionString <> vData Then
        mConnectionString = vData
        Call Me.Connect
    End If
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mConnectionString
End Property

Public Property Get SQLTextType() As clsAdoSQLType
    SQLTextType = mSQLType
End Property

Public Property Let SQLTextType(ByVal vData As clsAdoSQLType)
    mSQLType = vData
    Set mCmd = Nothing  '20120416  KD: made mCmd more useful
End Property

Public Property Get CurrentConnection() As ADODB.Connection
    Set CurrentConnection = mCN
End Property

Private Sub Class_Initialize()
    mInTransaction = False
End Sub

Private Sub Class_Terminate()
    '' KD Note: we can't close the connection here because it closes  references to it
'    If mRS.State = adStateOpen Then mRS.Close
'    If mCN.State = adStateOpen Then mCN.Close
    
    Set mCN = Nothing
    Set mCmd = Nothing
    Set mRS = Nothing
End Sub