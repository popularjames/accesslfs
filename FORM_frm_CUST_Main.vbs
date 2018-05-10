Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_CUST_Main
' Description:
'   Main Customer Service Event maintenance form.
'
'
' =============================================

Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1
Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents myCustService As clsCUSTSERVICE
Attribute myCustService.VB_VarHelpID = -1

Private mrsEvent As ADODB.RecordSet
Private mrsAuditClmHdr As ADODB.RecordSet
Private mrsAuditClmDtl As ADODB.RecordSet
Private mrsCaseTrackHdr As ADODB.RecordSet
Private mrsCaseTrackDtl As ADODB.RecordSet

Private mrsRelatedClaims As ADODB.RecordSet
Private mrsEventClaimActions As ADODB.RecordSet
Private mrsProviderContacts As ADODB.RecordSet
Private mrsEventProviderContact As ADODB.RecordSet
Private mrsNotes As ADODB.RecordSet

'Private mrsRelatedConcepts As ADODB.Recordset
'Private mrsContacts As ADODB.Recordset

Private mlEventID As String
Private mstrCnlyClaimNum As String
Private mstrCnlyProvID As String

Private mstrUserProfile As String
Private mstrUserName As String
Private miAppPermission As Integer
Private mbPermissionDenied As Boolean
Private mbEventDeleted As Boolean

Private mbAllowChange As Boolean
Private mbIsLoaded As Boolean
Private mbIsRefreshing As Boolean
Private mbHasRefreshed As Boolean
Private mbRelatedClaimsSetup As Boolean
Private mbIsProviderEvent As Boolean

Private mbRecordLocked As Boolean
Private mbRecordChanged
Private mdtOpenTime As Date

'mg 10/3/2013
Dim EmailID As String
Dim EmailSender As String
Dim emailMessage As String

Const CstrFrmAppID As String = "CustomerService"
Public Property Let IsLoaded(ByVal vData As Boolean)
    mbIsLoaded = vData
End Property
Public Property Get IsLoaded() As Boolean
    IsLoaded = mbIsLoaded
End Property
Public Property Get AppPermission() As Integer
    AppPermission = miAppPermission
End Property
Public Property Get RelatedClaimsSetup() As Boolean
    RelatedClaimsSetup = mbRelatedClaimsSetup
End Property
Public Property Let EventID(ByVal vData As Long)
    mlEventID = vData
End Property
Public Property Get EventID() As Long
    EventID = mlEventID
End Property
Public Property Get rsEvent() As ADODB.RecordSet
    Set rsEvent = mrsEvent
End Property
Public Property Get rsEventClaimActions() As ADODB.RecordSet
    Set rsEventClaimActions = mrsEventClaimActions
End Property
Public Property Let cnlyProvID(ByVal vData As String)
    mstrCnlyProvID = vData
End Property
Public Property Get IsRefreshing() As Boolean
    IsRefreshing = mbIsRefreshing
End Property
Public Property Get cnlyProvID() As String
    cnlyProvID = mstrCnlyProvID
End Property
Public Property Let CnlyClaimNum(ByVal vData As String)
    mstrCnlyClaimNum = vData
    If myCustService Is Nothing Then
    Else
        myCustService.CnlyClaimNum = mstrCnlyClaimNum
    End If
End Property
Public Property Get CnlyClaimNum() As String
    CnlyClaimNum = mstrCnlyClaimNum
End Property
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
    Me.txtAppID = CstrFrmAppID
End Property
Public Property Get rsRelatedClaims() As ADODB.RecordSet
    Set rsRelatedClaims = mrsRelatedClaims
End Property
Public Property Get CustService() As clsCUSTSERVICE
    Set CustService = myCustService
End Property

Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property

Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property
Property Let RecordLocked(data As Boolean)
    mbRecordLocked = data
    If mbRecordLocked Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Property
Public Function RefreshEvent()

    On Error GoTo ErrHandler

    RefreshMain
    
Exit_Function:
    Exit Function

ErrHandler:
'    RaiseEvent CustServiceError(Err.Description, Err.Number, strErrSource)
    Resume Exit_Function
End Function

Private Function ShouldSendEmail(Optional bReset As Boolean)
Static bAlreadySent As Boolean

    If bReset = True Then
        bAlreadySent = False
        Exit Function
    End If

    If bAlreadySent = False Then
        MsgBox "Going to send!"
        bAlreadySent = True
    End If

End Function


Public Function UpdateRelatedClaimTopic(strCnlyClaimNum As String, iTopicID As Integer)
Dim returnCode As Integer

    returnCode = myCustService.UpdateRelatedClaimTopic(strCnlyClaimNum, iTopicID)

End Function
Private Sub LookUpClaim(strParentEvent As String)
Dim frmCustQuicklookup As Form_frm_CUST_QuickLookup

GblParentEvent = strParentEvent

    On Error GoTo Err_btnAddClaim_Click
    
    Set frmCustQuicklookup = New Form_frm_CUST_QuickLookup
    
     Select Case GblParentEvent
            Case "AddClaim"
                frmCustQuicklookup.SearchType = "AUDITCLM"
            Case "Event"
                frmCustQuicklookup.SearchType = "CUSTEVENT"
      End Select
                
    frmCustQuicklookup.CustService = CustService
     
    frmCustQuicklookup.RefreshData
    ShowFormAndWait frmCustQuicklookup
     
    Set frmCustQuicklookup = Nothing

    'Reset the recordset, since it was reopened
    Set CUST_EVENT_Related_Claims.Form.RecordSet = myCustService.rsRelatedClaims

    If myCustService.rsRelatedClaims.recordCount > 0 Then
        mstrCnlyClaimNum = myCustService.CnlyClaimNum
        RefreshCurrentClaim (myCustService.CnlyClaimNum)
    Else
        mstrCnlyClaimNum = ""
    End If

     'RefreshMain
     
Exit_btnAddClaim_Click:
    Exit Sub

Err_btnAddClaim_Click:
    MsgBox Err.Description
    Resume Exit_btnAddClaim_Click

End Sub

Private Function RefreshClaimNotes()
Dim strError As String
Dim strTmpNotes As String
Dim rsNotes As ADODB.RecordSet
    
    On Error GoTo ErrHandler
    
    'Me.Caption = "View Notes"
    Set rsNotes = myCustService.rsNotes
    
    NotesText.SetFocus
    NotesText = ""
    
    If Not (rsNotes.BOF = True And rsNotes.EOF = True) Then
        rsNotes.MoveFirst
    End If
    
    strTmpNotes = ""
    While Not rsNotes.EOF
         strTmpNotes = strTmpNotes & "Added by " & Trim(UCase(rsNotes("NoteUserID"))) & " @ " & rsNotes("NoteDate") & " Note Type " & Trim(UCase(rsNotes("NoteType"))) & vbCrLf
         strTmpNotes = strTmpNotes & String(100, "-") & vbCrLf
         strTmpNotes = strTmpNotes & Trim(rsNotes("NoteText")) & vbCrLf & vbCrLf
         rsNotes.MoveNext
    Wend
    
    NotesText = strTmpNotes
    NotesText.SelLength = 0

    
exitHere:
    Exit Function
    
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "RefreshClaimNotes : "
    Resume exitHere

End Function


Private Sub btnAddClaim_Click()
    LookUpClaim ("addClaim")
End Sub

Private Sub btnDeleteClaim_Click()
    
    If Me.CUST_EVENT_Related_Claims.Form.RecordSet.recordCount = 1 Then
    
        MsgBox "You cannot delete the last Claim from a Claim Event!", vbCritical, "Error deleting related Claim"
        
    Else
    
        Dim MyAdo As clsADO
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "DELETE FROM CUST_Event_Related_Claim WHERE EventID = " & Me.CUST_EVENT_Related_Claims.Form.EventID & " AND CnlyClaimNum = '" & Me.CUST_EVENT_Related_Claims.Form.CnlyClaimNum & "'"
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute

        myCustService.RefreshRelatedClaimRecordSet
        
        'Get list of claims associated with this event and display them
        Set Me.CUST_EVENT_Related_Claims.Form.RecordSet = rsRelatedClaims
      
        
        CUST_EVENT_Related_Claims.Form.Requery
        
        
    End If
End Sub


Private Sub cmd_AddActionNotes_Click()

If Me.frmCustEventClaimActions.Form.RecordSet.recordCount = 0 Then
     MsgBox "Notes can't be added to an event without an action", vbOKOnly + vbCritical, "No Action ID Found"
    Exit Sub
End If

lngActionID = Me.frmCustEventClaimActions.Form.Controls("ActionID").Value
DoCmd.OpenForm "frm_CUST_Serv_Review_Results_Addnotes", acNormal




End Sub

Private Sub cmdAction_Click()
    'Launches another action based on what is selected
    'Damon 05/08
    On Error GoTo ErrHandler
    Dim strError As String
    Dim aParameters
    Dim rst As DAO.RecordSet
    Dim strSQL As String
    Dim strFunction As String
    Dim strFunctionResult As String
    Dim strParameterName As String
    Dim lngi As Long
    
    If Nz(Me.cboAction, "") = "" Then
        Exit Sub
    End If
    
    
    'Get the Access function for the selected action in the AUDITCLM_Action table
    strSQL = " select * from GENERAL_Action where FormName = '" & Me.Name & "' and AutoID = " & Me.cboAction & " "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL)
    If Not rst.EOF Then
        strFunction = rst!Function
        strParameterName = ""
        
        'Multiple parameters can be passed to a function
        'In the table, they are a comma delimited list.  Here, we split them into an array
        'and build an arguement
        'ACTION function variables must be passed as strings.  Do your conversion once they are in
        aParameters = Split(rst!ParameterName, ",")
        For lngi = 0 To UBound(aParameters)
            If lngi = 0 Then
                strParameterName = strParameterName & "'" & Me.Controls(aParameters(lngi)) & "'"
            Else
                strParameterName = strParameterName & ",'" & Me.Controls(aParameters(lngi)) & "'"
            End If
        Next lngi
               
        'The EVAL() function calls the Action.  The "StrFunction" variable has to point to a public module somewhere in Decipher
        'These exist in mod_AUDITCLM_Action
        strFunctionResult = Eval(strFunction & "(" & strParameterName & ")")
        RefreshMain
                
    End If
Exit Sub
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : cmdAction_Click"
End Sub
'Alex C - keep this as an example of using the calendar form
Private Sub cmdAdj_From_Click()
    On Error GoTo Err_btnChkDt_Click
    
'    Set frmCalendar = New Form_frm_GENERAL_Calendar
'    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
'    frmCalendar.DatePassed = Nz(Me.Adj_From, Date)
'    frmCalendar.RefreshData
'    ShowFormAndWait frmCalendar
    
'    Me.Adj_From = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click


End Sub
Private Sub cmdAddAction_Click()
Dim frm As Form_frm_CUST_Event_Claim_Action_Add

    On Error GoTo Err_handler
    
    If mstrCnlyClaimNum = "" Or mstrCnlyClaimNum = "" Then
        MsgBox "There is no current claim to enter an action.", vbOKOnly
        Exit Sub
    End If
    
    Set frm = New Form_frm_CUST_Event_Claim_Action_Add
   
    frm.ActionCnlyClaimNum = mstrCnlyClaimNum
    frm.ActionEventID = mlEventID
    frm.ActionUserName = mstrUserName
    frm.ActionrsEventClaimActions = mrsEventClaimActions
    
    ShowFormAndWait frm

    RefreshTabControl

    Set frm = Nothing

     RefreshMain
     
Exit_AddAction:
    Exit Sub

Err_handler:
    MsgBox Err.Description
    Resume Exit_AddAction

End Sub
'Alex C - adapt for use with Event
Private Sub cmdddNote_Click()
    On Error GoTo Err_cmdddNote_Click
    
     Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
    
     frmGeneralNotes.frmAppID = "AuditClm"
     Set mrsNotes = myCustService.rsNotes
     Set frmGeneralNotes.NoteRecordSource = mrsNotes
     
    Select Case Me.frmCustEvent.Form.Controls("mediumname")
        Case "Voicemail"
            frmGeneralNotes.DefaultNoteType = "PHONE"
        Case "Telephone"
            frmGeneralNotes.DefaultNoteType = "PHONE"
        Case "Fax"
            frmGeneralNotes.DefaultNoteType = "FAX"
        Case "Email"
            frmGeneralNotes.DefaultNoteType = "E-MAIL"
        Case Else
            frmGeneralNotes.DefaultNoteType = "TELEPHONE"
    End Select
    
     frmGeneralNotes.RefreshData
     
     ShowFormAndWait frmGeneralNotes
     
     Set frmGeneralNotes = Nothing
     
     myCustService.SaveData_Notes
     
     'Set mrsNotes = myCustService.rsNotes
     RefreshTabControl

Exit_cmdddNote_Click:
    Exit Sub

Err_cmdddNote_Click:
    MsgBox Err.Description
    Resume Exit_cmdddNote_Click
End Sub

Private Sub cmdExit_Click()

    On Error GoTo HandleError
      
   
    DoCmd.Close acForm, "frm_CUST_Main" 'Me.Name
    
exitHere:
    Exit Sub
HandleError:
'* Error 2501 will be caused by canceling the form close.
'Alex C 3/8/2012 - so will 3018, 3021, 3709

    If Err.Number = 2501 Or Err.Number = 3018 Or Err.Number = 3021 Or Err.Number = 3709 Then
        Resume Next
    Else
        MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
        GoTo exitHere
    End If
End Sub

Private Sub cmdLaunchTab_Click()
    'Launches a new window displaying the data selected in the tab list box
    'Damon 05/08
    On Error GoTo ErrHandler
    Dim strError As String
    
    Dim rst As DAO.RecordSet
    Dim lngTabID As Long
    Dim strSQL As String
    Dim strTabName As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    
    'Get the ID of the currently selected tab
    'lngTabID = Me.cboTabs
    'strSQL = GetNavigateTabSQL(lngTabID, Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
    'NewMainTab strSQL, mstrCnlyClaimNum, strTabName

Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : cmdLaunchTab_Click"
End Sub

Private Sub cmdLkupEvent_Click()


DoCmd.OpenForm "frm_CUST_Quick_Launch", acNormal

'LookUpClaim ("event")


End Sub

Private Sub cmdOpen_Click()
    
'Open the search form to get another event
    
    
End Sub




Private Sub Command389_Click()

End Sub

Private Sub Command405_Click()
    DoCmd.OpenForm "frm_CTS_Hdr_Create"
       
    
End Sub

Private Sub DeleteEvent_Click()

    If MsgBox("Are you sure you want to delete this event record and close?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        myCustService.DeleteEvent (mlEventID)
        mbEventDeleted = True
        
        'Detach the record sets from the forms so they cannot update the records
        If Not mrsEvent Is Nothing Then
            mrsEvent.CancelUpdate
            mrsEvent.Close
        End If
        If Not mrsRelatedClaims Is Nothing Then
            If Not mrsRelatedClaims.BOF And Not mrsRelatedClaims.EOF Then
                mrsRelatedClaims.CancelUpdate
                mrsRelatedClaims.Close
            End If
        End If
        If Not mrsEventClaimActions Is Nothing Then
            If Not mrsEventClaimActions.BOF And Not mrsEventClaimActions.EOF Then
                mrsEventClaimActions.CancelUpdate
                mrsEventClaimActions.Close
            End If
        End If
        
        'Close the form
        DoCmd.Close acForm, Me.Name
    End If

End Sub

Private Sub Form_Close()
Dim minutesDuration As Long

'    If myCustService.EventExists Then
'        If myCustService.LockedForEdit = True Then
'            myCustService.UnLockEvent
'        Else
'            myCustService.UnLockEvent (True)
'        End If
'    End If
    
    'If this form is closing because user does not have permission, don't use these objects - they don't exist yet
    'If the user has deleted the event, don't attempt to save
    If (mbEventDeleted = False And miAppPermission <> 0 And mbPermissionDenied = False) Then
        'Either save the seconds of duratio, or multiply user's input of minutes as seconds
        If mrsEvent("EventDuration") = 0 Then
            minutesDuration = Now() - mdtOpenTime
            minutesDuration = DateDiff("s", mdtOpenTime, Now())
            mrsEvent("EventDuration") = minutesDuration
        Else
            minutesDuration = mrsEvent("EventDuration") * 60
       End If
   
        mrsEvent.UpdateBatch

       'Save any unsaved changes
       'SaveData
    End If
    
    RemoveObjectInstance Me
    
    lngActionID = 0
    lngEventID = 0
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub

Private Sub Form_Load()
Dim strSQL As String

    
    Me.Caption = "Claim Processing"
    Me.RecordSource = ""
    
    Me.txtAppID = CstrFrmAppID
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    
    'Alex C 2/12/2012 - disable for initial release - there is separate security for CUST
    If miAppPermission = 0 Then Exit Sub
    
    Set myCustService = New clsCUSTSERVICE
    
    'Setting the APPID hardcoded for now to test other functionality
    mstrUserName = Identity.UserName
    
    mstrUserProfile = GetUserProfile()

    mbAllowChange = (miAppPermission And gcAllowChange)
           
    mbAllowChange = True
    If mbAllowChange = False Then
        'Me.cmdddNote.Enabled = False
        'Me.cmdSave.Enabled = False
    End If
   
    Me.Detail.visible = False
    cmdSave.Enabled = True
        
    Me.frmCustEventClaimActions.SourceObject = "frm_CUST_Event_Claim_Action"
    
    'Save the start time for calculating duration on close
    
    mdtOpenTime = Now()
    
    mbIsLoaded = True
    
    ''''****************TEST
    Set mrsEventClaimActions = myCustService.rsEventClaimActions
    'Set frmCustEventClaimActions.Form.Recordset = mrsEventClaimActions
    
End Sub
'The claim number has changed - new one added, or row has changed.
'Reset the recordsets for the claim, and update the tab control display
Public Function RefreshCurrentClaim(strCnlyClaimNum As String)
    
    CnlyClaimNum = strCnlyClaimNum
    myCustService.CnlyClaimNum = strCnlyClaimNum
    
    Set mrsAuditClmHdr = myCustService.rsAuditClmHdr
    Set mrsAuditClmDtl = myCustService.rsAuditClmDtl
    

    CurrentClaimLabel = "Selected Claim is " & mrsAuditClmHdr("ICN")

    'Get the provider id
    myCustService.cnlyProvID = mrsAuditClmHdr("CnlyProvID")
    mstrCnlyProvID = mrsAuditClmHdr("CnlyProvID")
    
    RefreshTabControl

End Function

'Try to create the event - check for user access, then set up initial record sets
Public Function CreateEventFromClaim() As Boolean
Dim strSQL As String
Dim strError As String

    On Error GoTo ErrHandler
    
    If myCustService.CanUserCreateEvent(mstrUserName) = False Then
        MsgBox "You are not authorized to use the Customer Service function.", vbOKOnly
        mbPermissionDenied = True
        CreateEventFromClaim = False
        Exit Function
    End If

    'Have the class create the data for the new event
    myCustService.CreateEventFromClaim mstrCnlyClaimNum, lngEventID

    'Get the Event recordset, module variables that should be set up just once per event
    mlEventID = myCustService.EventID
    
    Me.Caption = gstrAcctDesc & ": Event : " & IIf(mlEventID = 0, lngEventID, mlEventID)
    Me.Detail.visible = True
    
    
    'Get the recordsets for the initial related claim
    Set mrsAuditClmHdr = myCustService.rsAuditClmHdr
    Set mrsAuditClmDtl = myCustService.rsAuditClmDtl
    
    'Give the event recordset to the class and display
    frmCustEvent.visible = True
    Set mrsEvent = myCustService.rsEvent
    frmCustEvent.SourceObject = "frm_Cust_Event"
    Set frmCustEvent.Form.RecordSet = mrsEvent
    frmCustEvent.Form.Controls("addrtype") = ""
    frmCustEvent.Form.Controls("topicid") = 0
    
    'mg 10/01/2013 changed bc of we still want CS users to see it and add claims to it
    If mstrCnlyClaimNum = "0" Then
        'MsgBox "no claim test"
        frmCustEvent.Form.Controls("cnlyprovid") = "0"
    Else
        frmCustEvent.Form.Controls("cnlyprovid") = mrsAuditClmHdr("cnlyprovid")
    End If
    
    
    
    frmCustEvent.Form.Controls("cmb_EventTopic").visible = False
    frmCustEvent.Form.Refresh
    
    'txtSelectedProviderContact = cmbSelectedProviderContact.Column(1, cmbSelectedProviderContact.ListIndex) & IIf(Trim(cmbSelectedProviderContact.Column(2, cmbSelectedProviderContact.ListIndex)) = "", "", " - " & cmbSelectedProviderContact.Column(2, cmbSelectedProviderContact.ListIndex) & " - " & cmbSelectedProviderContact.Column(3, cmbSelectedProviderContact.ListIndex))
    
    'Get list of claims associated with this event and display them
    Set mrsRelatedClaims = myCustService.rsRelatedClaims
    CUST_EVENT_Related_Claims.SourceObject = "frm_cust_event_related_claims"
    Set Me.CUST_EVENT_Related_Claims.Form.RecordSet = rsRelatedClaims
    CUST_EVENT_Related_Claims.Form.UniqueTable = "CUST_Event_Related_Claim"
        
    mbRelatedClaimsSetup = True
        
    'mg 10/01/2013 changed bc of we still want CS users to see it and add claims to it
    If mstrCnlyClaimNum = "0" Then
        CurrentClaimLabel = "Selected Claim is " & "0"
        mstrCnlyProvID = "0"
    Else
        CurrentClaimLabel = "Selected Claim is " & mrsAuditClmHdr("ICN")
        'Get the provider id
        myCustService.cnlyProvID = mrsAuditClmHdr("CnlyProvID")
        mstrCnlyProvID = mrsAuditClmHdr("CnlyProvID")
    End If
    

       
    'Set the list of available contacts for the provider contact combobox - use the same provider regardless of additional claims added
    'strSQL = "SELECT addrtype, contactname from v_cust_prov_contacts where provaddrid = max(provaddrid) and CnlyProvID = '" & mstrCnlyProvID & "'"
    'StrSQL = "SELECT addrtype, ContactName from v_cust_prov_contacts group by AddrType, ContactName, AddrId, CnlyProvID having AddrId = max(AddrId) and CnlyProvID = '" & mstrCnlyProvID & "'"
    strSQL = " SELECT    a.AddrType, b.Description AS Type, a.Firstname + ' ' + a.LastName AS Name, a.Title, a.AddrId " & _
             " FROM      PROV_Address AS a INNER JOIN PROV_Xref_Address_code AS b ON b.AddrType = a.AddrType " & _
             " GROUP BY  a.AddrType, b.Description, a.Firstname + ' ' + a.LastName, a.Title, a.AddrId, a.CnlyProvID " & _
             " HAVING CnlyProvID = '" & mstrCnlyProvID & "' "
    
    frmCustEvent.Form.Controls("cmbSelectedProviderContact").RowSource = strSQL

    'Refresh the Event page
    RefreshMain
    
    
    CreateEventFromClaim = True
    
    
    'MG 10/03/2013 Added new tab to show email received from Provider
    'MsgBox "test from mike g"
    'Dim strSQL As String
    Dim recordCount As Integer
    Dim index As Integer
    
    Dim db As Database
    Dim rs As DAO.RecordSet
    
    
    strSQL = "SELECT EmailID,EmailSender,EmailDesc,EventDate FROM CUST_Event WHERE eventID=" & Me.EventID
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(strSQL)
        
    'rs.MoveLast
    'recordCount = rs.recordCount 'get record count
    'rs.MoveFirst
    
    'For index = 1 To recordCount
    'For index = 0 To recordCount - 1
        'MsgBox rs.Fields(1)
        'Me.lstSelectedClaims.AddItem rs.Fields(0) & ";" & rs.Fields(1)
        'rs.MoveNext
    'Next index
    
    EmailID = CStr(Nz(rs.Fields(0), ""))
    EmailSender = CStr(Nz(rs.Fields(1), ""))
    emailMessage = CStr(Nz(rs.Fields(2), ""))
    
    rs.Close
    db.Close
    
    'txtEmailID.SetFocus
    txtEmailID.Value = EmailID
    
    'txtEmailSender.SetFocus
    txtEmailSender.Value = EmailSender
    
    'txtEmailSender.SetFocus
    txtEmailMessage.Value = emailMessage
    
Exit Function

ErrHandler:
    mbIsRefreshing = False
        strError = Err.Description
        If Err.Number = 2467 Then
            MsgBox "Error: The Event you are trying to open does not exsits." & vbCr & "Please review and try again.", vbOKOnly + vbExclamation, "oops"
        Else
            MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbExclamation, "oops"
        End If
            
End Function

'Try to create the event - check for user access, then set up initial record sets
Public Function CreateEventFromProvider() As Boolean
Dim strSQL As String
Dim strError As String

    On Error GoTo ErrHandler
    
    If myCustService.CanUserCreateEvent(mstrUserName) = False Then
        MsgBox "You are not authorized to use the Customer Service function.", vbOKOnly
        mbPermissionDenied = True
        CreateEventFromProvider = False
        Exit Function
    End If

    'Have the class create the data for the new event
    myCustService.CreateEventFromProvider (mstrCnlyProvID)
    mbIsProviderEvent = True

    'Get the Event recordset, module variables that should be set up just once per event
    mlEventID = myCustService.EventID
    
    Me.Caption = gstrAcctDesc & ": Event : " & mlEventID
    Me.Detail.visible = True
    
    '2013-08-13 TK doesn't exist
    'Me("EventTopic").visible = True
    
    'Give the event recordset to the class and display
    frmCustEvent.visible = True
    Set mrsEvent = myCustService.rsEvent
    frmCustEvent.SourceObject = "frm_Cust_Event"
    Set frmCustEvent.Form.RecordSet = mrsEvent
    frmCustEvent.Form.Controls("addrtype") = ""
    frmCustEvent.Form.Controls("topicid") = 0
    frmCustEvent.Form.Controls("cnlyprovid") = mstrCnlyProvID
    frmCustEvent.Form.Controls("cmb_EventTopic").visible = True
    
    'Get list of claims associated with this event and display them
    Set mrsRelatedClaims = myCustService.rsRelatedClaims
    Set Me.CUST_EVENT_Related_Claims.Form.RecordSet = rsRelatedClaims
    CUST_EVENT_Related_Claims.Form.UniqueTable = "CUST_Event_Related_Claim"
        
    mbRelatedClaimsSetup = True
        
    CurrentClaimLabel = "Provider is " & mstrCnlyProvID
    
    'Set the list of available contacts for the provider contact combobox - use the same provider regardless of additional claims added
    strSQL = "SELECT addrtype, ContactName from CMS_AUDITORS_CODE.dbo.v_cust_prov_contacts group by AddrType, ContactName, AddrId, CnlyProvID having AddrId = max(AddrId) and CnlyProvID = '" & mstrCnlyProvID & "'"
    
    'frmCustEvent.Form.Controls("SelectedProviderContact").RowSource = strSQL
    frmCustEvent.Form.Controls("cmbSelectedProviderContact").RowSource = strSQL

    'Refresh the Event page
    RefreshMain

    CreateEventFromProvider = True
Exit Function

ErrHandler:
    mbIsRefreshing = False
        strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : CreateEventFromClaim"
End Function
Private Sub RefreshMain()
    'Refresh the main form
    'This form has control names that match the column names in the AuditCLM_Hdr table
    'Only textboxes and comboboxes with a TAG property of "R" are filled
    'To add a field to this form
    '   Add the column to the data table
    '   Create a textbox (or combobox) on the form wit the same name as the column name in the table
    '   Set the control's tag property to "R"
    
Dim strError As String
Dim rsRelatedClaims As ADODB.RecordSet
Dim strSQL As String
              
    On Error GoTo ErrHandler
              
    'Prevent any sub forms from attempting to refresh in On Change or other events
    'until the main refresh is finished
    mbIsRefreshing = True
        
    'Get list of claims associated with this event and display them
'    Set mrsRelatedClaims = myCustService.rsRelatedClaims
'    Set Me.CUST_EVENT_Related_Claims.Form.Recordset = rsRelatedClaims
'    CUST_EVENT_Related_Claims.Form.UniqueTable = "CUST_Event_Related_Claim"
   
    If mstrCnlyClaimNum <> "" Then
        Set mrsAuditClmHdr = myCustService.rsAuditClmHdr
        Set mrsAuditClmDtl = myCustService.rsAuditClmDtl
    End If
    

    
    'Refresh the current tab in tab control details
    RefreshTabControl
    
    mbIsRefreshing = False
    mbHasRefreshed = True
Exit Sub

ErrHandler:
    mbIsRefreshing = False
        strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : RefreshMain"
End Sub
Private Sub CheckEventOwner()
    
    'Alex C - check for who is assigned the Event
'    Dim strLOB As String
        
'    If myAuditClaim.rsAuditClmHdr.EOF <> True And myAuditClaim.rsAuditClmHdr.BOF <> True Then
'        strLOB = Nz(myAuditClaim.rsAuditClmHdr("LOB"), "")
'        If strLOB <> "" Then
'            MsgBox "ALERT: the claim belongs to '" & strLOB & "'", vbCritical
'        End If
'    End If

End Sub

Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim strerr As String

    'Don't try these tests unless the form has refreshed an event record
    'Also don't try these tests if the user has deleted the event and form is closing
    If mbEventDeleted = True Or mbHasRefreshed = False Then
        Exit Sub
    End If

    strerr = ""
    

    'MG 10/1/2013 Disable the below as it's very annoying and confusing for users
    'If Me.CUST_EVENT_Related_Claims.Form.areTopicsSet() = False Then
    '    strerr = "You must select a topic for each related claim before exiting."
    'End If
    
    'If Nz(Me.frmCustEvent.Form.Controls("ProvAddrID"), "") = "" Then
    '    If strerr = "" Then
    '        strerr = "You must select a provider contact before exiting."
    '    Else
    '        strerr = strerr & vbCrLf & "You must select a provider contact before exiting."
    '    End If
    'End If
    
    'If mbIsProviderEvent = True And Me.frmCustEvent.Form.Controls("topicid") = 0 Then
    '    If strerr = "" Then
    '        strerr = "You must select a topic for this Provider event."
    '    Else
    '        strerr = strerr & vbCrLf & "You must select a topic for this Provider event."
    '    End If
    'End If
    
    'If there are any missing required fields, notify the user and cancel the form close
    'If strerr <> "" Then
    '    MsgBox strerr, vbInformation + vbOKOnly, "Exit"
    '    Cancel = True
    'End If

        
exitHere:
    Exit Sub
End Sub
Private Sub cmdSave_Click()
Dim CurrEventID As Long
    
    
CurrEventID = IIf(Me.EventID = 0, lngEventID, Me.EventID)

'If Me.frmCustEventClaimActions.Form.Recordset.RecordCount <> 0 Then
'    If Me.frmCustEventClaimActions.Form.Controls("TriggerEmail") = 1 Then
'        Me.frmCustEventClaimActions.Form.Controls("TriggerEmail") = 0
'    End If
'End If

SaveData
    
SendNotificatonEmail CurrEventID
  
End Sub

Private Sub cmdSearch_Click()
'On Error GoTo Err_cmdSearch_Click
'    NewMainSearch "AUDITCLM", "v_AUDITCLM_Hdr", "Claims"
    
Exit_cmdSearch_Click:
    Exit Sub
Err_cmdSearch_Click:
    MsgBox Err.Description
    Resume Exit_cmdSearch_Click
End Sub

Private Sub frmGeneralNotes_NoteAdded()
    mbRecordChanged = True
End Sub

Public Sub LoadData()
'    Dim bLoaded As Boolean
'    bLoaded = myAuditClaim.LoadClaim(Me.CnlyClaimNum, mbAllowChange)
    
'    cmdSave.Enabled = False
'    cmdddNote.Enabled = False
    
'    If myAuditClaim.ClaimExists Then
'        If mbAllowChange Then
'            If myAuditClaim.LockedForEdit = False Then
'                MsgBox "Record is being locked by " & myAuditClaim.LockedUser & " at " & myAuditClaim.LockedDate
'            ElseIf mbAllowChange Then
'                cmdSave.Enabled = True
'                cmdddNote.Enabled = True
'            End If
'        End If
'    Else
'        MsgBox "Claim '" & CnlyClaimNum & "' does not exist for this account"
'    End If
'
    RefreshMain
'
'    CheckProviderStatus
'    CheckClaimOwner
    
End Sub
Private Sub SaveData()

'    Dim rs As DAO.Recordset
'    Dim strClaimStatusGroups As String
    Dim strError As String
    Dim bSaved As Boolean
'    Dim bLoaded As Boolean
'
'    On Error GoTo Err_SaveData

'    If mbRecordChanged = False And Me.Dirty = False Then
'        MsgBox "There are no changes to save."
'        Exit Sub
'    End If

    
    strError = ""
    
     bSaved = myCustService.SaveEvent
     If bSaved Then
        MsgBox "Record saved.", vbOKOnly + vbInformation, "Customer Service"
        RefreshMain
        mbRecordChanged = False
        Else
            MsgBox "Record not saved." & vbCrLf & strError, vbCritical
    End If

Exit_SaveData:
   Exit Sub

Err_SaveData:
   strError = Err.Description
   MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
   Resume Exit_SaveData

End Sub

Private Sub myCUSTSERVICE_EventError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_Event_Main : ADO Error"
End Sub

Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    'mReturnDate = ReturnDate
End Sub

Private Sub NotesText_GotFocus()
    'Don't refresh notes for this claim if there isn't one in focuc
    If mstrCnlyClaimNum = "" Or mstrCnlyClaimNum = "0" Then
        Exit Sub
    End If
    RefreshClaimNotes
End Sub

Private Sub TabCtl128_Change()
    RefreshTabControl
End Sub

Public Function RefreshTabControl()
    
'    'If no related claims, don't refresh the claim detail or notes
'    If (mstrCnlyClaimNum = "" Or mstrCnlyClaimNum = "0") And (TabCtl128.Value = 0 Or TabCtl128.Value = 1 Or TabCtl128.Value = 2) Then
'        Exit Function
'    End If
    
    Select Case TabCtl128.Value
        'Claim case tracking
        Case 0
'            frmclaimDetails.SourceObject = "frm_AUDITCLM_Main_Dtl"
'            Set mrsAuditClmDtl = myCustService.rsAuditClmDtl
'            Set frmclaimDetails.Form.RecordSet = mrsAuditClmDtl
'            frmclaimDetails.Locked = True

            '2014-10-16 TK adding case tracking info
            frmCaseTrackingHdr.SourceObject = "frm_CTS_Hdr"
            Set mrsCaseTrackHdr = myCustService.rsCaseTrackHdr
            Set frmCaseTrackingHdr.Form.RecordSet = mrsCaseTrackHdr
            frmCaseTrackingHdr.Locked = False
            
'            '2014-10-16 TK adding case tracking info history
'            frmCaseTrackingDtl.SourceObject = "frm_CTS_Hist"
'            Set mrsCaseTrackDtl = myCustService.rsCaseTrackDtl
'            Set frmCaseTrackDtl.Form.RecordSet = mrsCaseTrackDtl
'            frmCaseTrackingDtl.Locked = False
            
            
        ' Claim Notes
        Case 1
        RefreshClaimNotes
        
        ' Claim Actions
        Case 2
            Set mrsEventClaimActions = myCustService.rsEventClaimActions
            Set frmCustEventClaimActions.Form.RecordSet = mrsEventClaimActions
        
        'Provider Contacts
        Case 3
        'Set CUST_Org_Contact.Form.RecordSource = myCustService.rsProviderContacts
            Set mrsProviderContacts = myCustService.rsProviderContacts
            Set frmProvContacts.Form.Form.RecordSet = mrsProviderContacts
        
         'Claim Details; moved from case 0 to case 4
        Case 5
            frmclaimDetails.SourceObject = "frm_AUDITCLM_Main_Dtl"
            Set mrsAuditClmDtl = myCustService.rsAuditClmDtl
            Set frmclaimDetails.Form.RecordSet = mrsAuditClmDtl
            frmclaimDetails.Locked = True
        
        
    End Select
End Function
