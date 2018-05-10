Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Public Event CustServiceError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private mfrmCustMain As Form_frm_CUST_Main

Private mrsEvent As ADODB.RecordSet
Private mrsContacts As ADODB.RecordSet
Private mrsNotes As ADODB.RecordSet
Private mrsTopics As ADODB.RecordSet
Private mrsRelatedClaims As ADODB.RecordSet
Private mrsProviderContacts As ADODB.RecordSet
Private mrsEventProviderContact As ADODB.RecordSet
Private mrsEventClaimActions As ADODB.RecordSet
Private mrsAuditClmHdr As ADODB.RecordSet
Private mrsAuditClmDtl As ADODB.RecordSet
Private mrsCaseTrackHdr As ADODB.RecordSet
Private mrsCaseTrackDtl As ADODB.RecordSet


Private mAppID As String
Private mlEventID As Long
Private mstrCnlyClaimNum As String
Private mstrCnlyProvID As String

Private mbEventExists As Boolean

Private mNoteID As Long
Private mbLockedForEdit As Boolean
Private mstrLockedUser As String
Private mdLockedDt As Date
Private mstrCurrentUser As String
Public Property Get rsEvent() As ADODB.RecordSet
    Set rsEvent = mrsEvent
End Property
Public Property Get rsTopics() As ADODB.RecordSet
    'Get current list of topics
    MyAdo.sqlString = "select * CUST_Event_Topic"
    Set mrsTopics = MyAdo.OpenRecordSet("", False)
    
    Set rsTopics = mrsTopics
End Property
Public Property Get rsCustomTopics() As ADODB.RecordSet
    'Get current list of topics
    MyAdo.sqlString = "select * CUST_Event_Custom_Topic"
    Set mrsTopics = MyAdo.OpenRecordSet("", False)
    
    Set rsCustomTopics = mrsTopics
End Property
Public Property Get rsRelatedClaims() As ADODB.RecordSet
    Set rsRelatedClaims = mrsRelatedClaims
End Property
Public Property Get EventID() As Long
    EventID = mlEventID
End Property
Public Function custMain(vData As Form_frm_CUST_Main)
    Set mfrmCustMain = vData
End Function
Public Property Let CnlyClaimNum(ByVal vData As String)
    mstrCnlyClaimNum = vData
End Property
Public Property Let cnlyProvID(ByVal vData As String)
    mstrCnlyProvID = vData
End Property
Public Property Get CnlyClaimNum() As String
    CnlyClaimNum = mstrCnlyClaimNum
End Property
Public Property Get rsEventClaimActions() As ADODB.RecordSet
    MyAdo.sqlString = "SELECT * from cust_event_claim_action where EventID = " & IIf(mlEventID = 0, lngEventID, mlEventID) & "" '" and CnlyClaimNum = '" & mstrCnlyClaimNum & "'"
    Set mrsEventClaimActions = MyAdo.OpenRecordSet("", False)
    Set rsEventClaimActions = mrsEventClaimActions
End Property
Public Property Get rsProviderContacts() As ADODB.RecordSet
    myCode_ADO.sqlString = "SELECT * from v_prov_address where CnlyProvID = '" & mstrCnlyProvID & "'"
    Set mrsProviderContacts = myCode_ADO.OpenRecordSet("", False)
    Set rsProviderContacts = mrsProviderContacts
End Property
Public Property Get rsEventProviderContact() As ADODB.RecordSet
    MyAdo.sqlString = "SELECT * from CUST_Org_Contacts where EventID = " & mlEventID
    Set mrsEventProviderContact = MyAdo.OpenRecordSet("", False)
    Set rsEventProviderContact = mrsEventProviderContact
End Property
Public Property Get rsNotes() As ADODB.RecordSet
    
    'get the noteid associated with the claim
    mNoteID = Nz(mrsAuditClmHdr("NoteID"), -1)
        
    'open the note and disconnect
    MyAdo.sqlString = " SELECT * from NOTE_Detail where NoteID = " & mNoteID
    
    'myADO.SQLstring = "SELECT * from NOTE_Detail a join auditclm_hdr b on b.cnlyclaimnum = '" & mstrCnlyClaimNum & "' where a.NoteID = b.NoteID"
    Set mrsNotes = MyAdo.OpenRecordSet("", False)
    Set rsNotes = mrsNotes
End Property
Public Function DeleteEvent(iEventID As Long) As Long
    
    'call the stored procedure which deletes all related records for the event and the event itself
    myCode_ADO.sqlString = "exec usp_CUST_Delete_Event @EventID = " & iEventID
    myCode_ADO.Execute
    
End Function

Public Function UpdateRelatedClaimTopic(strCnlyClaimNum As String, iTopicID As Integer) As Integer
    
    'open the note and disconnect
    MyAdo.sqlString = "update CUST_Event_Related_Claim set TopicID = " & iTopicID & "where EventID = " & mlEventID & " and CnlyClaimNum = '" & strCnlyClaimNum & "'"
    MyAdo.Execute
    
End Function
Public Property Get rsAuditClmHdr() As ADODB.RecordSet
    MyAdo.sqlString = "SELECT * from auditclm_hdr where cnlyclaimnum = '" & mstrCnlyClaimNum & "'"
    Set mrsAuditClmHdr = MyAdo.OpenRecordSet("", False)
    Set rsAuditClmHdr = mrsAuditClmHdr
End Property
Public Property Get rsAuditClmDtl() As ADODB.RecordSet
    MyAdo.sqlString = "SELECT * from auditclm_dtl where cnlyclaimnum = '" & mstrCnlyClaimNum & "'"
    Set mrsAuditClmDtl = MyAdo.OpenRecordSet("", False)
    Set rsAuditClmDtl = mrsAuditClmDtl
End Property
Public Property Get rsCaseTrackHdr() As ADODB.RecordSet
    '2014-10-09 tk add case tracking
    MyAdo.sqlString = "SELECT * from CtsCaseHdr where CaseID = " & "'" & mlEventID & "'"
    Set mrsCaseTrackHdr = MyAdo.OpenRecordSet("", False)
    Set rsCaseTrackHdr = mrsCaseTrackHdr
End Property
Public Property Get rsCaseTrackDtl() As ADODB.RecordSet
    '2014-10-09 tk add case tracking history
    MyAdo.sqlString = "SELECT * from CtsCaseHdr_HIST where CaseID = " & "'" & mlEventID & "'"
    Set mrsCaseTrackDtl = MyAdo.OpenRecordSet("", False)
    Set rsCaseTrackDtl = mrsCaseTrackDtl
End Property

Public Property Get LockedForEdit() As Boolean
     LockedForEdit = mbLockedForEdit
End Property

Public Property Get LockedUser() As String
     LockedUser = mstrLockedUser
End Property
Public Property Get LockedDate() As Date
     LockedDate = mdLockedDt
End Property
'Remove after cleanup?
Public Property Get EventExists() As Boolean
    EventExists = mbEventExists
End Property
Public Function RefreshRelatedClaimRecordSet() As ADODB.RecordSet
   
    MyAdo.sqlString = "select EventID, TopicID, a.CnlyClaimNum, c.ConceptID, c.ConceptDesc from cust_event_related_claim a join auditclm_hdr b on b.cnlyclaimnum = a.cnlyclaimnum join concept_hdr c on c.conceptid = b.adj_conceptid where EventID = " & mlEventID
    Set mrsRelatedClaims = MyAdo.OpenRecordSet("", True)

End Function
Public Function CanUserCreateEvent(UserName As String) As Boolean
Dim rs As ADODB.RecordSet

    MyAdo.sqlString = "select userid from cust_security_user where UserID = '" & UserName & "'"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.AbsolutePosition = 1 Then
        CanUserCreateEvent = True
    Else
        CanUserCreateEvent = False
    End If
    
    Set rs = Nothing
    
    Exit Function

End Function
Public Function AddClaimToEvent(strCnlyClaimNum As String) As Integer

    On Error GoTo ErrHandler

    Dim strErrSource As String
    Dim i As Integer
    
    strErrSource = "clsCUSTSERVICE_AddClaimToEvent"

    'Create the new Event record
    MyAdo.SQLTextType = sqltext
    
    mstrCnlyClaimNum = strCnlyClaimNum
    
    If Not (mrsEvent.BOF And mrsEvent.EOF) Then
        mlEventID = mrsEvent("EventID")
        
        'Associate the claim with the new event
        MyAdo.sqlString = " Insert into CUST_Event_Related_Claim (EventID, CnlyClaimNum,TopicID, LastUpdateUser, LastUpdateDt) "
        MyAdo.sqlString = MyAdo.sqlString & "Values(" & mlEventID & ",'" & strCnlyClaimNum & "',0,'" & mstrCurrentUser & "',getdate())"
    
        MyAdo.Execute
        
        myCode_ADO.SQLTextType = sqltext
        myCode_ADO.sqlString = "select * from v_CUST_EVENT_Related_Claims where EventID = " & mlEventID
        Set mrsRelatedClaims = myCode_ADO.OpenRecordSet("", True)
   
        If mrsEvent.BOF And mrsEvent.EOF Then
            RaiseEvent CustServiceError("Could not insert Related Claim for existing Event", 99, strErrSource)
            GoTo ErrHandler
        End If
        
        'Refresh the list
        'mfrmCustMain.RefreshCurrentClaim (strCnlyClaimNum)
        
    End If

Exit_Function:
    Exit Function

ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function
Private Function AddOtherProviderAddress()
Dim strSQL As String
Dim rs As ADODB.RecordSet
Dim rs2 As ADODB.RecordSet
Dim strNow As String
Dim NewAddr01, NewAddr02, NewAddr03, NewZipCode, NewState, NewCity As String

    On Error GoTo ErrHandler

    Dim strErrSource As String
    strErrSource = "clsCUSTSERVICE_AddOtherProviderAddress"

    strNow = Format(Now(), "mm/dd/yyyy")
    
    'If there are no Other ('OT') address lines for this provider, add one
    strSQL = "select * from prov_address where cnlyprovid = '" & mstrCnlyProvID & "' and effdt <= '" & strNow & "' and termdt >= '" & strNow & "' and addrtype = 'OT'"
    MyAdo.sqlString = strSQL
    Set rs = MyAdo.OpenRecordSet
    If rs.BOF And rs.EOF Then

        ' the new provider address (OT) will copy the addr01, addr02, addr03, City, State and Zip from the first other valid address from the same provider
        ' this way a state code exists for this provider address.
        'James Segura 06/11/2012
        
        strSQL = "select CnlyProvID, Addr01, Addr02, Addr03, City, State, Zip from prov_address where cnlyprovid = '" & mstrCnlyProvID & "' and effdt <= '" & strNow & "' and termdt >= '" & strNow & "' ORDER BY AddrType"
        MyAdo.sqlString = strSQL
        Set rs2 = MyAdo.OpenRecordSet
        If Not (rs2.BOF And rs2.EOF) Then
            NewAddr01 = Nz(rs2("Addr01"), "")
            NewAddr02 = Nz(rs2("Addr02"), "")
            NewAddr03 = Nz(rs2("Addr03"), "")
            NewCity = Nz(rs2("City"), "")
            NewState = Nz(rs2("State"), "")
            NewZipCode = Nz(rs2("Zip"), "")
            If NewAddr01 = "" Then NewAddr01 = " "
            If NewAddr02 = "" Then NewAddr02 = " "
            If NewAddr03 = "" Then NewAddr03 = " "
            If NewCity = "" Then NewCity = " "
            If NewState = "" Then NewState = " "
            If NewZipCode = "" Then NewZipCode = " "
        Else
            NewAddr01 = " "
            NewAddr01 = " "
            NewAddr01 = " "
            NewCity = " "
            NewState = " "
            NewZipCode = " "
        End If
       
       ' 2013-08-13 TK BEGIN:
'        strSQL = "exec usp_prov_address_insert @pcnlyprovid = '" & mstrCnlyProvID & "', @paddrtype = 'OT', @peffdt = '" & strNow & "', @pTermDt = '12/31/9999', @pFirstname = ' ', @pMiddleInit = ' ', @pLastname = ' ', @pTitle = ' ', @pAddr01 = '" & NewAddr01 & "', @pAddr02 = '" & NewAddr02 & "', @pAddr03 = '" & NewAddr03 & "', @pCity = '" & NewCity & "', @pState = '" & NewState & "', @pZip = '" & NewZipCode & "', @pZipExt = ' ', @pPhone = ' ', @pPhoneExt = ' ', @pFax = ' ', @pEmail = ' ', @pComments = 'Default address to provide Customer Service with Other option.',  @plastupdateuser = 'Alex.Cannon', @plastupdatedt = '" & strNow & "', @pSFileName = ' ', @pDataSource = ' ', @pErrMsg = ' '"
'        mycode_ADO.sqlString = strSQL
'        mycode_ADO.Execute
        strSQL = "exec CMS_Auditors_Code.dbo.usp_prov_address_insert @pcnlyprovid = '" & mstrCnlyProvID & "', @paddrtype = 'OT', @peffdt = '" & strNow & "', @pTermDt = '12/31/9999', @pFirstname = ' ', @pMiddleInit = ' ', @pLastname = ' ', @pTitle = ' ', @pAddr01 = '" & NewAddr01 & "', @pAddr02 = '" & NewAddr02 & "', @pAddr03 = '" & NewAddr03 & "', @pCity = '" & NewCity & "', @pState = '" & NewState & "', @pZip = '" & NewZipCode & "', @pZipExt = ' ', @pPhone = ' ', @pPhoneExt = ' ', @pFax = ' ', @pEmail = ' ', @pComments = 'Default address to provide Customer Service with Other option.',  @plastupdateuser = 'Alex.Cannon', @plastupdatedt = '" & strNow & "', @pSFileName = ' ', @pDataSource = ' ', @pErrMsg = ' '"
        MyAdo.sqlString = strSQL
        MyAdo.Execute
        ' 2013-08-13 TK END:
    End If


Exit_Function:
    Exit Function

ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function

End Function
Public Function CreateEventFromClaim(strCnlyClaimNum As String, Optional intEventID As Long) As Integer

    On Error GoTo ErrHandler

    Dim strErrSource As String
    Dim mlEventIDlk As Long
    
    strErrSource = "clsCUSTSERVICE_CreateEventFromClaim"
    mlEventIDlk = IIf(intEventID = 0, -1, intEventID)
    

    'Create the new Event record
    MyAdo.SQLTextType = sqltext
    
    mstrCnlyClaimNum = strCnlyClaimNum
    
    MyAdo.sqlString = " Insert into CUST_Event (CreatedByUserID, EventDate, MediumName, EventDuration, LastUpdateUser, LastUpdateDt) "
    MyAdo.sqlString = MyAdo.sqlString & "Values('" & mstrCurrentUser & "',getdate(),'Telephone',0,'" & mstrCurrentUser & "',getdate())"
                    
    'Create the eventrecord
    MyAdo.Execute
    
    
    
    If mlEventIDlk = -1 Then
    'Associate the claim with the new event
            MyAdo.sqlString = "select * from CUST_Event where LastUpdateDt = (select max(LastUpdateDt) from CUST_Event)"
            Set mrsEvent = MyAdo.OpenRecordSet("", False)
            
'            Set mrsEvent = MYADO.ExecuteRS
            
            mlEventID = mrsEvent("EventID")
                               
            
            MyAdo.sqlString = " Insert into CUST_Event_Related_Claim (EventID, CnlyClaimNum,TopicID, LastUpdateUser, LastUpdateDt) "
            MyAdo.sqlString = MyAdo.sqlString & "Values(" & mlEventID & ",'" & strCnlyClaimNum & "',0,'" & mstrCurrentUser & "',getdate())"
            MyAdo.Execute
                    
    
    Else
            'If exsits get record for event
            MyAdo.sqlString = "select * from cms_auditors_claims.dbo.CUST_Event where EventID ='" & mlEventIDlk & "'"
            Set mrsEvent = MyAdo.OpenRecordSet("", False)
            mlEventID = mrsEvent("EventID")
            
    End If
    
    'miEventID = mrsEvent("EventID")
    
    If mrsEvent.BOF And mrsEvent.EOF Then
        RaiseEvent CustServiceError("Could not insert Related Claim for new Event", 99, strErrSource)
        GoTo ErrHandler
    End If
    
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.sqlString = "select * from v_CUST_EVENT_Related_Claims where EventID = " & mlEventID
    Set mrsRelatedClaims = myCode_ADO.OpenRecordSet("", True)
   
    mstrCnlyProvID = mrsRelatedClaims("cnlyprovid")
    
    'If there's no detault address for the provider of type Other ('OT'), add one
    AddOtherProviderAddress

Exit_Function:
    Exit Function

ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function
Public Function CreateEventFromProvider(strCnlyProvID As String) As Integer

    On Error GoTo ErrHandler

    Dim strErrSource As String

    
    strErrSource = "clsCUSTSERVICE_CreateEventFromProvider"

    'Create the new Event record
    MyAdo.SQLTextType = sqltext
    
    mstrCnlyClaimNum = "0"
    mstrCnlyProvID = strCnlyProvID
    
    'If there's no detault address for the provider of type Other ('OT'), add one
    AddOtherProviderAddress
    
    MyAdo.sqlString = " Insert into CUST_Event (CreatedByUserID, EventDate, MediumName, EventDuration, LastUpdateUser, LastUpdateDt) "
    MyAdo.sqlString = MyAdo.sqlString & "Values('" & mstrCurrentUser & "',getdate(),'Telephone',0,'" & mstrCurrentUser & "',getdate())"
                    
    'Create the empty related claim recordset
    MyAdo.Execute
    
    MyAdo.sqlString = "select * from cms_auditors_claims.dbo.CUST_Event where LastUpdateDt = (select max(LastUpdateDt) from cms_auditors_claims.dbo.CUST_Event)"
    Set mrsEvent = MyAdo.OpenRecordSet("", False)
    
    If mrsEvent.BOF And mrsEvent.EOF Then
        RaiseEvent CustServiceError("Could not obtain Related Claim recordset for new Event from Provider", 99, strErrSource)
        GoTo ErrHandler
    End If
    
    mlEventID = mrsEvent("EventID")
        
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.sqlString = "select * from v_CUST_EVENT_Related_Claims where EventID = " & mlEventID
    Set mrsRelatedClaims = myCode_ADO.OpenRecordSet("", True)
   
Exit_Function:
    Exit Function

ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function


' Alex C - modify this to load an existing event record
Public Function LoadClaim(strCnlyClaimNum As String, Optional bAllowChange As Boolean = True) As Boolean
'Damon 7/15/08
'Loads a claim based on what is passed to the method
    On Error GoTo ErrHandler
        
    Dim strErrSource As String
    strErrSource = "clsAuditClm_LoadClaim"
    
   
    mbLockedForEdit = False
    mbEventExists = False
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = strCnlyClaimNum
    
    MyAdo.sqlString = " SELECT * from AUDITCLM_Hdr WHERE cnlyClaimNum = '" & strCnlyClaimNum & "' and AccountID = " & gintAccountID
    'open the audit claims header and disconnect
    Set mrsAuditClmHdr = MyAdo.OpenRecordSet("", False)
    
    If Not (mrsAuditClmHdr.BOF And mrsAuditClmHdr.EOF) Then
        'Assigning the current Claim Status at the time the claim was loaded
        'This is a property of the class that will allow us to build processing logic around status code changes
        'mstrLoadClaimStatus = mrsAuditClmHdr("ClmStatus")
        
        
        ' TKL 3/2/2011: auditor credit override
        'mstrClaimAuditor = mrsAuditClmHdr("Adj_Auditor") & ""
        'mstrOverrideAuditor = ""
        
        'Check if the claim has an existing lock on it
        'if so, lock it down
        mstrLockedUser = mrsAuditClmHdr("LockUserID") & ""
        If Not IsNull(mrsAuditClmHdr("LockDt")) Then
            mdLockedDt = mrsAuditClmHdr("LockDt")
        End If
        
        'If (mstrLockedUser = "" Or mstrLockedUser = mstrCurrentUser) And bAllowChange Then
        '    If LockEvent Then
        '        mbLockedForEdit = True
        '        mstrLockedUser = mstrCurrentUser
        '    End If
        'End If
        
        
        'open the audit claims detail and disconnect
        MyAdo.sqlString = " SELECT * from AUDITCLM_Dtl WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'"
        Set mrsAuditClmDtl = MyAdo.OpenRecordSet()
        
        Select Case gintAccountID
            Case 1
            Case 2
            Case 3
            Case 4
                MyAdo.sqlString = " SELECT * from AMERIGROUP_AUDITCLM_Hdr_AdditionalInfo WHERE cnlyClaimNum = '" & strCnlyClaimNum & "'"
        End Select
        
        'get the noteid associated with the claim
        mNoteID = Nz(mrsAuditClmHdr("NoteID"), -1)
        
        'open the note and disconnect
        MyAdo.sqlString = " SELECT * from NOTE_Detail where NoteID = " & mNoteID
        Set mrsNotes = MyAdo.OpenRecordSet("", False)
        LoadClaim = True
        mbEventExists = True
    Else
        mbEventExists = False
        LoadClaim = False
    End If
    
Exit_Function:
    Exit Function
    
ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function
'Alex C - adapt this to save the event
Public Function SaveEvent() As Boolean
    
    Dim bResult As Boolean
    Dim strErrSource As String
    Dim iResult As Integer
    Dim strErrMsg As String

    On Error GoTo ErrHandler
    
    strErrSource = "clsAuditClm_SaveEvent"
    
   ' myCode_Ado.BeginTrans
    bResult = True
    
   ' If bResult Then
   '     ' save the notes
   '     bResult = SaveData_Notes
   '     If bResult = False Then
   '         Err.Raise 65000, strErrSource, "Error saving claim notes.  Record not saved"
   '     End If
   ' End If
    
   ' 'Save the main event record
   'mrsEvent.UpdateBatch
    
   ' 'Save related claims
   ' mrsRelatedClaims.UpdateBatch
    
   ' myCode_Ado.CommitTrans
    SaveEvent = bResult
    
Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    myCode_ADO.RollbackTrans
    SaveEvent = bResult
    GoTo Exit_Sub
End Function
'Alex C - adapted this to save claim notes from the customer service screen
Public Function SaveData_Notes() As Boolean
    Dim bResult As Boolean
    
    On Error GoTo ErrHandler
    
    If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
        'If the noteID is -1 then we need to create a new ID
        If mNoteID = -1 Then
            'This is a public function that gets a unique ID based on the app being passed to the method
            mNoteID = GetAppKey("NOTE")
            'Set the recordset of the header to contain the new note ID
            Me.UpdateField "NoteID", mNoteID
            'Apply this new noteID to all of the records in the note recordset
            If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
                mrsNotes.MoveFirst
                While Not mrsNotes.EOF
                    mrsNotes.Update
                    mrsNotes("NoteID") = mNoteID
                    mrsNotes.MoveNext
                Wend
            End If
        End If
        'Pass the recordset back to SQL synching the results
        bResult = myCode_ADO.Update(mrsNotes, "usp_NOTE_Detail_Apply")
        If bResult = False Then
            Err.Raise 65000, "", "Error updating claim note"
        End If
   Else
        bResult = True
    End If
    
    SaveData_Notes = bResult

Exit_Sub:
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    SaveData_Notes = False
    GoTo Exit_Sub
End Function
Private Sub Class_Initialize()
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    'Alex C - check if this works
    mAppID = "AUDITCLM"
    'mAppID = "CUSTSERVICE"
    mstrCurrentUser = Identity.UserName
End Sub
Private Sub Class_Terminate()
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    
    Set mrsEvent = Nothing
    Set mrsContacts = Nothing
    Set mrsProviderContacts = Nothing
    Set mrsEventProviderContact = Nothing
    Set mrsNotes = Nothing
    Set mrsTopics = Nothing
    Set mrsEventClaimActions = Nothing
    Set mrsRelatedClaims = Nothing
    Set mrsAuditClmHdr = Nothing
    Set mrsAuditClmDtl = Nothing
    Set mrsCaseTrackHdr = Nothing
    Set mrsCaseTrackDtl = Nothing

End Sub
'Alex C - adapt to Event
Public Function LockEvent() As Boolean
           
    On Error GoTo ErrHandler

    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsCUST_LockEvent"
    
    'set the objects claim number to the passed in claim
    Me.CnlyClaimNum = Me.CnlyClaimNum
    
    ' lock claim for edit
'    myCODE_Ado.SQLstring = "usp_AUDITCLM_Hdr_Lock"
'    myCODE_Ado.SQLTextType = StoredProc

'    Set LocCmd = New ADODB.Command
'    LocCmd.ActiveConnection = myCODE_Ado.CurrentConnection
'    LocCmd.CommandText = myCODE_Ado.SQLstring
'    LocCmd.CommandType = adCmdStoredProc
'    LocCmd.Parameters.Refresh
            
'    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
'    LocCmd.Parameters("@pLockUser") = Identity.Username()
'    LocCmd.Parameters("@pLockTime") = Now()
   

'    iResult = myCODE_Ado.Execute(LocCmd.Parameters)
'    iResult = LocCmd.Parameters("@RETURN_VALUE")
'   strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
'    If iResult <> 0 Then
'        MsgBox strErrMsg
'        LockClaim = False
'        GoTo ErrHandler
'    End If
    
    LockEvent = True
Exit Function

ErrHandler:
    LockEvent = False
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    
End Function
'Alex C - adapt to Event
Public Function UnLockEvent(Optional MsgSuppress As Boolean = False) As Boolean
           
    On Error GoTo ErrHandler
           
    Dim iResult As Long
    Dim LocCmd As New ADODB.Command
    Dim strErrMsg As String
        
    Dim strErrSource As String
    strErrSource = "clsCUST_UnLockEvent"
    
    'set the objects claim number to the passed in claim
'    Me.CnlyClaimNum = Me.CnlyClaimNum
    
    'lock claim for edit
'    myCODE_Ado.Connect
'    myCODE_Ado.SQLstring = "usp_AUDITCLM_Hdr_UnLock"
'    myCODE_Ado.SQLTextType = StoredProc

'    Set LocCmd = New ADODB.Command
'    LocCmd.ActiveConnection = myCODE_Ado.CurrentConnection
'    LocCmd.CommandText = myCODE_Ado.SQLstring
'    LocCmd.CommandType = adCmdStoredProc
'    LocCmd.Parameters.Refresh
            
'    LocCmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
'    LocCmd.Parameters("@pLockUserID") = Identity.Username()
    
    
'    iResult = myCODE_Ado.Execute(LocCmd.Parameters)
'    iResult = LocCmd.Parameters("@RETURN_VALUE")
'    strErrMsg = Nz(LocCmd.Parameters("@pErrMsg").Value, "")
    
'    Select Case iResult
'        Case Is <> 0
'            If MsgSuppress = False Then
'                ' this is to suppress message when use exit claim screen
'                MsgBox "Claim is locked by " & strErrMsg
'            End If
'            UnLockClaim = False
'        Case Else
'            UnLockClaim = True
'    End Select
'Exit Function

ErrHandler:
    UnLockEvent = False
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    
End Function
Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub
Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub
' Alex C - remove after cleanup?
Public Function UpdateField(FieldName As String, FieldValue As Variant)
    Dim strErrSource As String
    
    On Error GoTo ErrHandler
    strErrSource = "clsAuditClm_UpdateField"
    mrsAuditClmHdr(FieldName).Value = FieldValue
    mrsAuditClmHdr.UpdateBatch adAffectAllChapters 'must do this to avoid recordset not syncing when displaying on forms
Exit Function

ErrHandler:
    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
End Function