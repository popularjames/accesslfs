Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Change Log
'2010-06-25:Gautam
'   Changed code to implement address change requests from the online portal.
'2010-08-05:Gautam
'   Changed code to encapsulate portal updates and address updates into a single transaction

Option Compare Database

Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Public Event ProvError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private mstrCnlyProvID As String
Private miNoteID As Long

Private mrsPROVHdr As ADODB.RecordSet
Private mrsPROVAddr As ADODB.RecordSet
Public mrsPROVAddrPortal As ADODB.RecordSet

Private mrsPROVAddrDeleted As ADODB.RecordSet
Private mrsPROVNotes As ADODB.RecordSet
Private mbLockedForEdit As Boolean
Private mbProvExists As Boolean
Public mbProvAddrChanged As Boolean
Private mstrLockedUser As String
Private mbLockedDt As Date
Private mstrCurrentUser As String



Public Function LoadProv(cnlyProvID As String, Optional bAllowChange As Boolean = True) As Boolean
    
    On Error GoTo Err_handler
        
    Dim strErrSource As String
    Dim strErrMsg As String
    
    strErrSource = "clsPROV_LoadProv"
    strErrMsg = ""
    
    ' read provider header record
    mstrCnlyProvID = cnlyProvID
    Me.LoadProvider
    
    mbLockedForEdit = False
    mbProvExists = False
    
    If Not (mrsPROVHdr.BOF And mrsPROVHdr.EOF) Then
        ' provider exists, check if record is locked by other.
        ' If not locked then lock it and set lock indicator
        mstrLockedUser = mrsPROVHdr("LockUserID") & ""
        If (mstrLockedUser = "" Or mstrLockedUser = mstrCurrentUser) And bAllowChange Then
            'lock record
            If LockProv Then
                mbLockedForEdit = True
                mstrLockedUser = mstrCurrentUser
            End If
        End If
        miNoteID = Nz(mrsPROVHdr("NoteID"), -1)
        mbProvExists = True
    Else
        ' provider does not exist, set note id to return blank record
        miNoteID = -1
    End If
    
    ' read provider address record
    Me.LoadAddress
    
    If mrsPROVAddrPortal.recordCount > 0 Then
        mbProvAddrChanged = vbTrue
    Else
        mbProvAddrChanged = vbFalse
    End If
    ' read provider notes
    Me.LoadNotes
    
    LoadProv = True
    Exit Function

Err_handler:
    LoadProv = False
    If strErrMsg = "" Then strErrMsg = Err.Description
    RaiseEvent ProvError(strErrMsg, 0, strErrSource)
End Function

Public Sub LoadProvider()
    MyAdo.sqlString = "select * from PROV_Hdr WHERE CnlyProvID = '" & mstrCnlyProvID & "' and AccountID = " & gintAccountID
    Set mrsPROVHdr = MyAdo.OpenRecordSet()
End Sub

Public Sub LoadAddress()

    'original code
    'myADO.SqlString = "select * from PROV_Address WHERE CnlyProvID = '" & mstrCnlyProvID & "'"
    'Set mrsPROVAddr = myADO.OpenRecordSet()
    'myADO.SqlString = "select * from PROV_Address WHERE 1=2"
    'Set mrsPROVAddrDeleted = myADO.OpenRecordSet
    'myADO.SqlString = "select * from PROV_Address_PortalPending WHERE CnlyProvID = '" & mstrCnlyProvID & "' AND ReqStatusCode = '01'"
    'Set mrsPROVAddrPortal = myADO.OpenRecordSet()
        
    'MG only show active providers based on view, but cause issue with update and need more coding changes
    'mycode_Ado.SqlString = "select * from v_PROV_Address WHERE CnlyProvID = '" & mstrCnlyProvID & "'"
    'Set mrsPROVAddr = mycode_Ado.OpenRecordSet()
    'mycode_Ado.SqlString = "select * from v_PROV_Address WHERE 1=2"
    'Set mrsPROVAddrDeleted = mycode_Ado.OpenRecordSet
    'mycode_Ado.SqlString = "select * from v_PROV_Address_PortalPending WHERE CnlyProvID = '" & mstrCnlyProvID & "'"
    'Set mrsPROVAddrPortal = mycode_Ado.OpenRecordSet()
    
    'MG Add filter in where clause pulling table
    MyAdo.sqlString = "select * from PROV_Address pa WHERE CnlyProvID = '" & mstrCnlyProvID & "' AND getDate() BETWEEN EffDt AND TermDt AND exists (select 1 from prov_xref_address_code where addrtype = pa.addrtype and Active = 'Y')"
    Set mrsPROVAddr = MyAdo.OpenRecordSet()
    MyAdo.sqlString = "select * from PROV_Address WHERE 1=2"
    Set mrsPROVAddrDeleted = MyAdo.OpenRecordSet
    MyAdo.sqlString = "select * from PROV_Address_PortalPending pap WHERE CnlyProvID = '" & mstrCnlyProvID & "' AND ReqStatusCode = '01' AND exists (select 1 from prov_xref_address_code where addrtype = pap.addrtype and Active = 'Y')"
    Set mrsPROVAddrPortal = MyAdo.OpenRecordSet()
    
End Sub

Public Sub LoadNotes()
    MyAdo.sqlString = "select * from NOTE_Detail WHERE NoteID = " & miNoteID & ""
    Set mrsPROVNotes = MyAdo.OpenRecordSet()
End Sub

Public Function SaveProv() As Boolean
    
    Dim bResult As Boolean
    Dim strErrSource As String
    Dim cmd As ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String

    On Error GoTo Err_handler
    
    strErrSource = "clsPROV_SaveClaim"

    
    bResult = False
    If Not (mrsPROVHdr.BOF And mrsPROVHdr.EOF) Then
        myCode_ADO.BeginTrans
        mrsPROVHdr.MoveFirst
        mrsPROVHdr("LastUpdateDt") = Now
        mrsPROVHdr("LastUPdateUser") = mstrCurrentUser
       
        mrsPROVHdr("LockUserID") = mstrCurrentUser
        mrsPROVHdr("LockDt") = Now
        If Not (mrsPROVNotes.BOF And mrsPROVNotes.EOF) Then
            mrsPROVNotes.MoveFirst
            mrsPROVHdr("NoteID") = mrsPROVNotes("NoteID")
            miNoteID = mrsPROVHdr("NoteID")
        End If
        
        bResult = myCode_ADO.Update(mrsPROVHdr, "usp_PROV_Hdr_Apply")
        If bResult = False Then
            strErrMsg = "Error: can not save provider record"
            GoTo Err_handler
        End If
        
        ' remove deleted addresses if exist
        If Not (mrsPROVAddrDeleted.BOF And mrsPROVAddrDeleted.EOF) Then
            Set cmd = CreateObject("ADODB.Command")
            cmd.ActiveConnection = myCode_ADO.CurrentConnection
            cmd.commandType = adCmdStoredProc
            cmd.CommandText = "usp_PROV_Address_Delete"
            cmd.Parameters.Refresh
            
            mrsPROVAddrDeleted.MoveFirst
            With mrsPROVAddrDeleted
                While Not .EOF
                    cmd.Parameters("@pCnlyProvID") = !cnlyProvID
                    cmd.Parameters("@pAddrID") = !AddrId
                    cmd.Execute
                    If cmd.Parameters("@RETURN_VALUE") <> 0 Then
                        strErrMsg = "Error: can not save provider address"
                        GoTo Err_handler
                    End If
                    .MoveNext
                Wend
            End With
        End If

        ' save address if exist
        If Not (mrsPROVAddr.BOF And mrsPROVAddr.EOF) Then
            bResult = myCode_ADO.Update(mrsPROVAddr, "usp_PROV_Address_Apply")
            If bResult = False Then
                strErrMsg = "Error: can not save provider address"
                GoTo Err_handler
            End If
            
            'Update Portal line status to 02 and lastupdt to now
            If Not (mrsPROVAddrPortal.BOF And mrsPROVAddrPortal.EOF) Then
                bResult = myCode_ADO.Update(mrsPROVAddrPortal, "usp_PROV_PortalAddress_Apply")
                If bResult = False Then
                    strErrMsg = "Error: can not Update Portal Status"
                    GoTo Err_handler
                End If
            End If
        End If
    
        ' save notes if exists
        If Not (mrsPROVNotes.BOF And mrsPROVNotes.EOF) Then
            bResult = myCode_ADO.Update(mrsPROVNotes, "usp_NOTE_Detail_Apply")
            If bResult = False Then
                strErrMsg = "Error: can not save provider note"
                GoTo Err_handler
            End If
        End If
    
        myCode_ADO.CommitTrans
        
        Me.LoadProvider
        Me.LoadAddress
        Me.LoadNotes
        mbProvExists = True
        mbLockedForEdit = True
    Else
        mbProvExists = False
        MsgBox "Can not save blank provider record"
    End If
    
    SaveProv = bResult
    
Exit_Function:
    Exit Function

Err_handler:
    'Rollback anything we did up until this point
    strErrMsg = strErrMsg & vbCrLf & Err.Description
    RaiseEvent ProvError(strErrMsg, Err.Number, strErrSource)
    myCode_ADO.RollbackTrans
    SaveProv = False
    GoTo Exit_Function
End Function


Public Property Let cnlyProvID(ByVal vData As String)
    mstrCnlyProvID = vData
End Property

Public Property Get cnlyProvID() As String
    cnlyProvID = mstrCnlyProvID
End Property

Public Property Get PROVAddrPortal() As ADODB.RecordSet
    Set PROVAddrPortal = mrsPROVAddrPortal
End Property

Public Property Get PROVHdr() As ADODB.RecordSet
    Set PROVHdr = mrsPROVHdr
End Property

Public Property Get PROVAddr() As ADODB.RecordSet
    Set PROVAddr = mrsPROVAddr
End Property

Public Property Set PROVAddr(data As ADODB.RecordSet)
    Set mrsPROVAddr = data
End Property

Public Property Get PROVAddrDeleted() As ADODB.RecordSet
    Set PROVAddrDeleted = mrsPROVAddrDeleted
End Property

Public Property Get PROVNotes() As ADODB.RecordSet
    Set PROVNotes = mrsPROVNotes
End Property

Public Property Get LockedForEdit() As Boolean
     LockedForEdit = mbLockedForEdit
End Property

Public Property Get LockedUser() As String
     LockedUser = mstrLockedUser
End Property

Public Property Get LockedDate() As Date
     LockedDate = mbLockedDt
End Property

Public Property Get ProviderExists() As Boolean
    ProviderExists = mbProvExists
End Property

Private Sub Class_Initialize()
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    mstrCurrentUser = Identity.UserName
End Sub
Private Sub Class_Terminate()
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set mrsPROVHdr = Nothing
    Set mrsPROVAddr = Nothing
    Set mrsPROVNotes = Nothing
End Sub

Public Function LockProv() As Boolean
           
    On Error GoTo Err_handler
           
    
    Dim iResult As Long
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsPROV_LockProv"
    
    'set the objects claim number to the passed in claim
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandText = "usp_PROV_Hdr_Lock"
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyProvID").Value = mstrCnlyProvID
    cmd.Parameters("@pLockUserID").Value = mstrCurrentUser
    cmd.Parameters("@pLockDt").Value = Now()
    cmd.Parameters("@pErrMsg").Value = ""
    
    myCode_ADO.sqlString = "usp_PROV_Hdr_Lock"
    myCode_ADO.SQLTextType = StoredProc

    iResult = myCode_ADO.Execute(cmd.Parameters)
    strErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
    If iResult = -1 Then
        MsgBox strErrMsg
        LockProv = False
        GoTo Exit_Function
    End If
    
    LockProv = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
    
Err_handler:
    LockProv = False
    Err.Raise Err.Description, Err.Number, strErrSource

    Resume Exit_Function
    
End Function

Public Function UnLockProv() As Boolean
           
    On Error GoTo Err_handler
           
    
    Dim iResult As Long
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsPROV_UnLockClaim"
    
    If mstrCnlyProvID = "" Then
        MsgBox "Provider ID is not defined"
        UnLockProv = False
        Exit Function
    End If
    
    'set the objects claim number to the passed in claim
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandText = "usp_PROV_Hdr_UnLock"
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyProvID") = mstrCnlyProvID
    cmd.Parameters("@pLockUserID") = Identity.UserName()
    
    myCode_ADO.sqlString = "usp_PROV_Hdr_UnLock"
    myCode_ADO.SQLTextType = StoredProc
    
    iResult = myCode_ADO.Execute(cmd.Parameters)
    strErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
    If iResult = -1 Then
        MsgBox strErrMsg
        UnLockProv = False
        GoTo Exit_Function
    End If
    UnLockProv = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
    
Err_handler:
    UnLockProv = False
    Err.Raise Err.Description, Err.Number, strErrSource
    
    Resume Exit_Function
End Function


Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "clsPROV : ADO Error"
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "clsPROV : ADO Error"
End Sub