Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private ColReSize3 As clsAutoSizeColumns
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents co_ADO As clsADO
Attribute co_ADO.VB_VarHelpID = -1

Private mstrFieldSource As String
Private mstrWhere As String
Private lngQueryType As Long '* These are values from msysobjects 1/4/6 = table, 5 = query
Private rstLetterWorkTable As ADODB.RecordSet

Private WithEvents frmFilter As Form_frm_GENERAL_Filter
Attribute frmFilter.VB_VarHelpID = -1
Private msAdvancedFilter As String

Private mstrGroupsArray() As String
Private mstrAuditorsArray() As String

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private mdReturnDate As Date
Private cofrmParent As Form_frm_LETTER_Main


Private cbClickedAdd As Boolean

Private clContractId As Long

'' 9/17/2014 KD: Need to fix this add to queue stuff - put the insert query into a stored proc

Const CstrFrmAppID As String = "LetterQueueAdd"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get SelectedContractId() As Long
    SelectedContractId = clContractId
End Property
Public Property Let SelectedContractId(lContractId As Long)
    clContractId = clContractId
End Property


Private Sub dtLetterDate_AfterUpdate()
    If Me.dtLetterDate <> "" And IsNull(Me.dtLetterDate) = False Then
        If IsDate(Me.dtLetterDate) = False Then
            MsgBox "Letter Request Date must be a valid date."
            Me.dtLetterDate = Date
        ElseIf CDate(Me.dtLetterDate) < Date Then
            MsgBox "Letter Request Date cannot be a past date."
            Me.dtLetterDate.Value = Date
        End If
    End If
End Sub

Private Sub frmfilter_QueryFormRefresh()
    RefreshMain
End Sub

Private Sub frmFilter_UpdateSql()
    msAdvancedFilter = frmFilter.SQL.WherePrimary
    RefreshMain
End Sub

Private Sub lstSenders_AfterUpdate()
 Dim rstGroupAuditors As DAO.RecordSet

 Dim strSQL As String
 Dim varItemSelected As Variant
 Dim varGroupAuditor As Variant
 Dim varAuditor As Variant
 
 Dim bolFirstItemSelected As Boolean
 Dim bolFirstGroup As Boolean
 Dim bolFirstAuditor As Boolean
 Dim bolUserIDFound As Boolean
    
    Erase mstrGroupsArray
    Erase mstrAuditorsArray

    ReDim mstrGroupsArray(0)
    ReDim mstrAuditorsArray(0)
    
    Me.txtSenders = ""
    bolFirstGroup = True
    bolFirstAuditor = True
   
    'Loop through all selected items
    For Each varItemSelected In Me.lstSenders.ItemsSelected
    
        'FYI:
        'Me.lstSenders.Column(0, varItemSelected) = ID
        'Me.lstSenders.Column(1, varItemSelected) = Name
        'Me.lstSenders.Column(2, varItemSelected) = Type

        'If Type = Group, then get all UserIDs associated with group
        If Me.lstSenders.Column(2, varItemSelected) = "Group" Then
            strSQL = "SELECT GA.UserName " & _
            "FROM ADMIN_User_Group G INNER JOIN ADMIN_User_GroupAssign GA ON G.GroupID = GA.GroupID " & _
            "WHERE G.GroupID = '" & Me.lstSenders.Column(0, varItemSelected) & "' ORDER BY GA.UserID;"
            
            Set rstGroupAuditors = CurrentDb.OpenRecordSet(strSQL)
            rstGroupAuditors.MoveFirst
            
            Me.txtSenders.Value = Me.lstSenders.Column(1, varItemSelected) & "{"
            
            If bolFirstGroup = False Then
                ReDim Preserve mstrGroupsArray(UBound(mstrGroupsArray) + 1)
            Else
                bolFirstGroup = False
            End If
            mstrGroupsArray(UBound(mstrGroupsArray)) = Mid(Me.lstSenders.Column(1, varItemSelected), 8)
            
            'Loop through all UserIDs in Group to display them
            Do Until rstGroupAuditors.EOF

                If bolFirstAuditor = False Then
                    Me.txtSenders.Value = Me.txtSenders.Value & ", "
                    ReDim Preserve mstrAuditorsArray(UBound(mstrAuditorsArray) + 1)
                Else
                    bolFirstAuditor = False
                End If
                Me.txtSenders.Value = Me.txtSenders.Value & rstGroupAuditors!UserID
                mstrAuditorsArray(UBound(mstrAuditorsArray)) = rstGroupAuditors!UserID

                rstGroupAuditors.MoveNext
            Loop
            
            Me.txtSenders.Value = Me.txtSenders.Value + "}"
            rstGroupAuditors.Close
        'Else the selected row is not a group, so just display the UserID
        Else
            bolUserIDFound = False
            
            For Each varAuditor In mstrAuditorsArray
                If Me.lstSenders.Column(1, varItemSelected) = varAuditor Then
                    bolUserIDFound = True
                End If
            Next varAuditor
            'If the UserID has already been placed in txtSenders, then do not display again.
            'Otherwise, place UserID in txtSenders and strAuditorsArray
            If bolUserIDFound = False Then

                If bolFirstAuditor = False Then
                    Me.txtSenders.Value = Me.txtSenders.Value & ", "
                    ReDim Preserve mstrAuditorsArray(UBound(mstrAuditorsArray) + 1)
                Else
                    bolFirstAuditor = False
                End If
                Me.txtSenders.Value = Me.txtSenders.Value & DLookup("UserName", "Admin_User", "UserID = '" & Me.lstSenders.Column(1, varItemSelected) & "'")
                mstrAuditorsArray(UBound(mstrAuditorsArray)) = DLookup("UserName", "Admin_User", "UserID = '" & Me.lstSenders.Column(1, varItemSelected) & "'")
            
            End If
        End If
    Next varItemSelected
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub

Private Sub co_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub

Private Sub btnDate_Click()

    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.dtLetterDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    If mdReturnDate <> "12:00:00 AM" Then
        If mdReturnDate < Date Then
            MsgBox "Letter Request Date cannot be a past date."
            Me.dtLetterDate = Date
            Me.dtLetterDate.SetFocus
        Else
            Me.dtLetterDate = mdReturnDate
        End If
    End If

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub

Public Sub RefreshMain()
 On Error GoTo Error_Encountered

'This dropdown will refresh the letter types available from the Letter_type Table...
'once a type is selected we populate the listbox below based from the table in the letter_type table.
Dim ClaimSource As String
Dim clmcount As Integer: clmcount = 0
Dim ltrcount As Integer: ltrcount = 0
Dim lngRecCount As Long
Dim strErrMsg As String: strErrMsg = ""
Dim sFrom As String
Dim MyAdo As clsADO
Dim sqlString As String
Dim Result As Integer
Dim Sstr As String
Dim dtStart As Date
Dim strProcName As String
Dim sWhere As String
Dim sContractIdClause As String
Dim cmd As ADODB.Command

    strProcName = ClassName & ".RefreshMain"
    
    DoCmd.Hourglass True
         
    If cofrmParent Is Nothing Then Set cofrmParent = Me.Parent.Form
    If cofrmParent.ContractId <> 0 Then
        sContractIdClause = " AND A.ContractId = " & CStr(cofrmParent.ContractId)
    Else
        Stop ' need a contract id here
    End If
         
    'Checking if there is a souce for this letter (should be a SQL server view or Query)
    'Get the view or query source of the letter
dtStart = Now()
    sFrom = Nz(DLookup("LetterSource", "LETTER_Type", "LetterType = '" & Me.cboLetterType & "'"), "")
    If sFrom = "" Then
        GoTo Error_Encountered
    End If


    'Create a primary key on the Letter table to allow DAO inserts.
    'This "should" work anyway, if there is a PK on the table already, but there is no harm in it being here
    CreatePK "LETTER_Selection_Temp", "SelectionID"
    
        
    'Delete the current user's data from the temp letter table
    sqlString = "delete from LETTER_Selection_Temp  where LETTER_Selection_Temp.username = '" & Identity.UserName & "'"
    
    'Create a new instance of the ADO class to handle the delete call we just defined in the query above
    Set MyAdo = New clsADO
       
    'Connect to the database and execute the delete
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = sqlString
    Result = MyAdo.Execute
          
    'Done with the ADO class, clean it out
    Set MyAdo = Nothing

    'Set the SQL string equal to nothing to reinitialize it
    Sstr = ""
   
    ' KD COMEBACK - If we are returning from inserting to the queue then we do not need to do this which may take a LONG time..
'    Stop
'    If cbClickedAdd = True Then
'Stop
'        GoTo Block_Exit
'    End If
   
    'TGH NOTES: 12/2/08
    ' this is pretty crazy the below code is how access executes an insert.
    'DPR NOTES 03/3/09 WOW I did not know that there was this bizzaro world SQL syntax for an insert!
    '***This SQL takes the data from the letter source and puts it into the Letter Selection temp table***
    'This is a local Access insert, and since we usually source letters from SQL, this should be moved to an ADO call.
            '    Sstr = "SELECT " & sFrom & ".*, '" & Identity.UserName & "' as UserName FROM " & sFrom & IIf(msAdvancedFilter <> "", " WHERE " & msAdvancedFilter, "")
            '    Sstr = "insert into LETTER_Selection_Temp(UserName) " & Sstr
            '    CurrentDb.Execute (Sstr)

    'Creazy no more, I ADOized the above. Making sure the timeout is rather long

    sWhere = " WHERE (LT.UseLegacySystem != 0) "
    If msAdvancedFilter <> "" Then
        msAdvancedFilter = " AND (" & msAdvancedFilter & ")"
    End If


    sFrom = "FROM " & IIf(left(sFrom, 2) = "v_", "cms_auditors_code.dbo." & sFrom, "cms_auditors_claims.dbo." & sFrom) & " a "
    sFrom = sFrom & " INNER JOIN CMS_AUDITORS_CLAIMS.dbo.LETTER_Type LT ON LT.LetterType = '" & Me.cboLetterType & "' AND LT.AccountId = " & CStr(IIf(gintAccountID = 0, 1, gintAccountID))
    sFrom = sFrom & " AND a.ContractId = LT.ContractId "
    Sstr = "SELECT a.*, '" & Identity.UserName & "' as UserName " & sFrom & sWhere & msAdvancedFilter
    
    Sstr = Sstr & sContractIdClause
    
    Sstr = "insert into cms_auditors_claims.dbo.LETTER_Selection_Temp " & Sstr
       

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = GetConnectString("v_CODE_Database")
    cmd.CommandText = Sstr
    cmd.CommandTimeout = 0
    cmd.commandType = adCmdText
    cmd.Execute

    Set MyAdo = Nothing
    Set cmd = Nothing


        'TGH added 01/08/09 to ensure they are using their only table.
        'Code to initialize the list
    Set Me.lstClaims.RecordSet = Nothing
    Me.lstClaims.RowSource = vbNullString
        'Set our listbox columns to be the same as letter_selection_temp
    Me.lstClaims.ColumnCount = CurrentDb.TableDefs("Letter_selection_temp").Fields.Count
        'Set the data in the list equal to what is in the Selection table with any additional filters applied
    Me.lstClaims.RowSource = "SELECT * FROM LETTER_Selection_Temp A " & IIf(msAdvancedFilter <> "", " WHERE " & msAdvancedFilter & " and Username = '" & Identity.UserName & "'", " WHERE Username = '" & Identity.UserName & "'" & sContractIdClause) & " order by cnlyprovid, cnlyclaimnum "


        'added last minute to exclude error messag 04/24/08 purpose if recordset is empty then there was a problem with
        ' the report vs just an empty dataset
    If Me.lstClaims.RecordSet Is Nothing Or Nz(Me.lstClaims.ItemData(1), "") = "" Then
        strErrMsg = "No Records Returned for this query. "
        GoTo Error_Encountered
    End If
    
    DoCmd.Hourglass False

        'update the count for the listbox.
    recountselected clmcount, ltrcount, Me.lstClaims, True
    Me.txtLtrCount = ltrcount
    Me.txtClaimCount = clmcount

        ' resizing of the listbox.  max columns is 58 columns to be reformatted.
    Set ColReSize3 = New clsAutoSizeColumns
    ColReSize3.SetControl Me.lstClaims
        'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstClaims.ListCount - 1 > 0 Then
        ColReSize3.AutoSize
    End If
    
Block_Exit:
    Exit Sub
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox ("View did not come back correctly: " & Err.Description)
    End If
    DoCmd.Hourglass False
End Sub

Private Sub cboLetterType_AfterUpdate()
    RefreshMain
End Sub

Private Sub AddLettersToQueue()
On Error GoTo Block_Err
Dim intI As Long
Dim strInstanceID As String
Dim strSessionID As String
Dim strLetterType As String
Dim strLetterReqDt As String
Dim i As Integer
Dim j As Integer
Dim intStatus As Integer
Dim sMsg As String
Dim strErrMsg As String
Dim bMultipleAuditors As Boolean
Dim iMaxCol As Integer
Dim varItem As Variant
Dim bResult As Boolean
Dim strAuditor As String
Dim iReturn As Integer
Dim fmrStatus As Form_ScrStatus
Dim colPrms As ADODB.Parameters
Dim prm As ADODB.Parameter
Dim LocCmd As New ADODB.Command
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim dicProviderList As Scripting.Dictionary
Dim strProcName As String
Dim dtStart As Date
Dim LetterTypeIsADR As Boolean

    strProcName = ClassName & ".AddLettersToQueue"
    
    'Connect to the databases with ADO
    Set MyAdo = New clsADO
    Set co_ADO = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    co_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    'Allow ADR letter types to have a letter date in the future
    
    LetterTypeIsADR = DLookup("ADR", "LETTER_Type", "LetterType = '" & Me.cboLetterType & "'")
    
    'Damon 7/08
    'Some checking early on in the process
    'Just making sure all of the selections are made before moving on
    If lstClaims.ListCount = 0 Then
        MsgBox "There is no record to add to queue"
        GoTo Block_Exit
    ElseIf lstClaims.ItemsSelected.Count = 0 Then
        MsgBox "Letters must be selected."
        SelectAllLButton.SetFocus
        GoTo Block_Exit
    ElseIf IsNull(Me.dtLetterDate) Or Me.dtLetterDate = "" Then
        MsgBox "Please enter a letter request date"
        Me.dtLetterDate.SetFocus
        GoTo Block_Exit
    'CMS SPECIFIC RULE
    ElseIf CDate(Nz(Me.dtLetterDate, "")) < Date Then
        MsgBox "Letter Request Date cannot be a past date."
        Me.dtLetterDate.SetFocus
        GoTo Block_Exit
    ElseIf CDate(Nz(Me.dtLetterDate, "")) > Date And Not LetterTypeIsADR Then
        MsgBox "Letter Request Date cannot be a future date."
        Me.dtLetterDate.SetFocus
        GoTo Block_Exit
    ElseIf CDate(Nz(Me.dtLetterDate, "")) > DateAdd("d", 7, Date) And LetterTypeIsADR Then
        MsgBox "Letter Request Date cannot be over a week into the future."
        Me.dtLetterDate.SetFocus
        GoTo Block_Exit
    ElseIf Me.lstSenders.ListIndex = -1 Then
        MsgBox "Please choose a sender"
        Me.lstSenders.SetFocus
        GoTo Block_Exit
    End If

    'Get a dictionary to keep track of providers
    Set dicProviderList = New Scripting.Dictionary

    'Damon 7/08
    'Setup progress screen
    intStatus = 1
    Set fmrStatus = New Form_ScrStatus
    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        .ProgMax = lstClaims.ItemsSelected.Count '- 1
        .TimerInterval = 50
        .show
    End With
    
    'Damon 7/08
    'Start a Transaction

    co_ADO.BeginTrans
'LogMessage strProcName, "PERFORMANCE TEST", "1. " & ProcessTookHowLong(dtStart)
dtStart = Now()

    'Get a new unique identifier that will be associated with this batch of letters
    strSessionID = GetInstanceID
'LogMessage strProcName, "PERFORMANCE TEST", "2. " & ProcessTookHowLong(dtStart)
   
    'Loop through what is selected in the list
    For Each varItem In Me.lstClaims.ItemsSelected
'dtStart = Now()
        'Move through the recordset of work claims and associate an auditor with a provider
        'While Not rstLetterWorkTable.EOF
        If j > UBound(mstrAuditorsArray) Then
           j = 0
        End If
        
        'DPR creates an association of provider and auditor in a dictionary to allow multiple auditors to split the letters evenly
        'TGH took out code to loop through the recordest because if we are looping already just run this each time first so we know if we have a new provider it will be in the
        'array before we run the below code.  So removing an extra loop thorough the recordset.
        If Not dicProviderList.Exists(Trim(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("CnlyProvID").OrdinalPosition, varItem))) Then
            dicProviderList.Add Trim(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("CnlyProvID").OrdinalPosition, varItem)), mstrAuditorsArray(j)
            j = j + 1
        End If
'LogMessage strProcName, "PERFORMANCE TEST", "3. " & ProcessTookHowLong(dtStart)
'dtStart = Now()

        'Call usp_LETTER_Work_Table_Insert via ADO.Execute
        'Also passing in the auditor based off of the auditors dictionary so we can handle multiple auditors assigned.
        co_ADO.sqlString = "usp_LETTER_Work_Table_Insert"
        co_ADO.SQLTextType = StoredProc
        'SelectedID based on the current row in the loop
        Set prm = LocCmd.CreateParameter("@pSelectionID", adVarChar, adParamInput, 20, Trim(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("selectionid").OrdinalPosition, varItem)))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@pInstanceID", adVarChar, adParamInput, 20, strSessionID)
        LocCmd.Parameters.Append prm
        'Auditor based on the dictionary paring
        Set prm = LocCmd.CreateParameter("@pAuditor", adVarChar, adParamInput, 50, dicProviderList.Item(Trim(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("CnlyProvID").OrdinalPosition, varItem))))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@pLetterType", adVarChar, adParamInput, 50, Me.cboLetterType.Value)
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@pLetterReqDt", adDBDate, adParamInput, , CStr(Format(CDate(Nz(Me.dtLetterDate, Date)), "MM/DD/YYYY")))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@ErrMsg", adVarChar, adParamOutput, 255)
        LocCmd.Parameters.Append prm
        iReturn = co_ADO.Execute(LocCmd.Parameters)
        strErrMsg = Nz(LocCmd.Parameters("@ErrMsg").Value, "")

'LogMessage strProcName, "PERFORMANCE TEST", "4. " & ProcessTookHowLong(dtStart)
'dtStart = Now()

        If strErrMsg <> "" Then
Stop
            GoTo Block_Err
        End If
        Set LocCmd = Nothing
        
       'Update Progress Bar.
       intStatus = intStatus + 1
       sMsg = "Adding Record " & intStatus & " / " & fmrStatus.ProgMax
       fmrStatus.ProgVal = intStatus
       fmrStatus.StatusMessage sMsg
       
       'Just checking if cancel is selected
       If fmrStatus.EvalStatus(2) = True Then
                sMsg = "Cancel has been selected. No records added!"
                fmrStatus.StatusMessage sMsg
                DoEvents
                strErrMsg = sMsg
                GoTo Block_Err
       End If

       'Damon 7/08
       'I have no idea.....
       'this is to communicate with the application running the progress screen,  doing this over and over to allow for active refresh of the progress screen.
       'DR - Yes but why do you have to call DOEVENTS 3000 times? KD: ROFL!!!
       DoEvents
       DoEvents
       DoEvents

    
    Next varItem
    
'' Need to commit this now..
'Stop
'co_ADO.CommitTrans
'
'co_ADO.BeginTrans
'Stop
    'Calling usp_LETTER_Work_Queue_Mass_Insert.  This procedure takes everything that was placed in the Work table and adds it to the letter system to be printed
    'This needs to be called for each "Auditor" that was associated with a letter, hence we are looping through the dictionary paring
    j = 0
    While j <= UBound(mstrAuditorsArray)
dtStart = Now()

        co_ADO.sqlString = "usp_LETTER_Work_Queue_Mass_Insert"
        co_ADO.SQLTextType = StoredProc
        Set prm = LocCmd.CreateParameter("@SessionID", adVarChar, adParamInput, 20, strSessionID)
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("Auditor", adVarChar, adParamInput, 75, mstrAuditorsArray(j))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@InstanceID", adVarChar, adParamOutput, 20)
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@ErrMsg", adVarChar, adParamOutput, 255)
        LocCmd.Parameters.Append prm
        iReturn = co_ADO.Execute(LocCmd.Parameters)
        strErrMsg = Nz(LocCmd.Parameters("@ErrMsg").Value, "")

'LogMessage strProcName, "PERFORMANCE TEST", "5. " & ProcessTookHowLong(dtStart)

        'Damon 7/08
        'I need to define what I am doing here, what I want is that if this fails, kill the whole thing
        'Select Case iReturn
        If strErrMsg <> "" Then
        Stop
            LogMessage strProcName, "ERROR", "Error in " & co_ADO.sqlString & " was: " & strErrMsg
            GoTo Block_Err
        End If
        Set LocCmd = Nothing
        j = j + 1
    Wend
dtStart = Now()
    'Commit the transaction!
    co_ADO.CommitTrans
    
   
'LogMessage strProcName, "PERFORMANCE TEST", "6. " & ProcessTookHowLong(dtStart)

    
    
    'Refresh the letters!
    RefreshMain
    
Block_Exit:
    'Demolish the status form
    Set fmrStatus = Nothing
    Set MyAdo = Nothing
    Set co_ADO = Nothing
    Exit Sub
    
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg, , True
    Else
        ReportError Err, strProcName
    End If
    co_ADO.RollbackTrans
    GoTo Block_Exit
End Sub

Private Sub cmdAddLettersToQueue_Click()
    cbClickedAdd = True
    AddLettersToQueue
    cbClickedAdd = False
End Sub

Private Sub cmdPrintFromQueue_Click()
    DoCmd.Hourglass False
End Sub

Private Sub btnEnDate_Click()
    On Error GoTo Err_btnEnDate_Click

    Dim stDocName As String
    Dim stForm As String
    Dim stField As String
    Dim stLinkCriteria As String

    stDocName = "frmCalendar"
    stForm = Me.Name
    stField = "dtEndDate"
    DoCmd.OpenForm stDocName, , , , , , stForm & "|" & stField

Exit_btnEnDate_Click:
    Exit Sub

Err_btnEnDate_Click:
    MsgBox Err.Description
    Resume Exit_btnEnDate_Click
End Sub

Private Sub btnStDate_Click()
    On Error GoTo Err_btnStDate_Click

    Dim stDocName As String
    Dim stForm As String
    Dim stField As String
    Dim stLinkCriteria As String

    stDocName = "frmCalendar"
    stForm = Me.Name
    stField = "dtStartDate"
    DoCmd.OpenForm stDocName, , , , , , stForm & "|" & stField

Exit_btnStDate_Click:
    Exit Sub

Err_btnStDate_Click:
    MsgBox Err.Description
    Resume Exit_btnStDate_Click
End Sub

Public Sub Form_Load()
Dim iAppPermission As Integer
Dim sContractIdClause As String
Dim strSQL As String
Dim sUserProfile As String

    Call Account_Check(Me)
    
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    

    
    ' check letter views
    Check_Letter_Views
    
    If cofrmParent Is Nothing Then Set cofrmParent = Me.Parent.Form
    If cofrmParent.ContractId <> 0 Then
        sContractIdClause = " AND ContractId = " & CStr(cofrmParent.ContractId)
    Else
        sContractIdClause = ""
    End If
    
    ' set up letter type row source
    Me.cboLetterType.RowSource = " select LetterType, LetterDesc, ClaimCnt " & _
                                 " from qry_AddToQueue_LetterType_List " & _
                                 " where AccountID = " & gintAccountID
    Me.cboLetterType.Requery


    sUserProfile = GetUserProfile()



    'strSQL = "select lt.LetterType, lt.LetterDesc, COUNT(1) " & _
    '            "from CMS_AUDITORS_CLAIMS.dbo.LETTER_Type lt with (nolock) " & _
    '            "join CMS_AUDITORS_CLAIMS.dbo.auditclm_process_logics pl with (nolock) on lt.LetterType = pl.DataType and pl.AccountID = " & gintAccountID & _
    '            "join CMS_AUDITORS_CLAIMS.dbo.QUEUE_Hdr qh with (nolock) on qh.QueueType = pl.CurrQueue and qh.AccountID = " & gintAccountID & _
    '            " where lt.For_DS_only <= " & str(IIf(sUserProfile = "CM_Admin", 1, 0)) & " " & _
    '            "group by lt.LetterType, lt.LetterDesc " & _
    '            "order by lt.LetterType "

    'MG 10/29/2013 Display letters based on queueType and exceptions base

    ' KD: 12/17/2014 - Added UseLegacySystem to where clause for new letter system
    strSQL = "select LetterType,LetterDesc,ClaimCount=MIN(ClaimCount)" & _
                " From" & _
                " (" & _
                " select lt.LetterType, lt.LetterDesc, ClaimCount=COUNT(1)" & _
                " from CMS_AUDITORS_CLAIMS.dbo.LETTER_Type lt with (nolock)" & _
                    " inner join CMS_AUDITORS_CLAIMS.dbo.auditclm_process_logics pl with (nolock) on lt.LetterType = pl.DataType and pl.AccountID = lt.accountid " & _
                    " inner join CMS_AUDITORS_CLAIMS.dbo.QUEUE_Hdr qh with (nolock) on qh.QueueType = pl.CurrQueue and qh.AccountID = pl.AccountID " & _
                    " WHERE lt.UseLegacySystem != 0 AND lt.isLetterActive = 1 and lt.accountid = " & gintAccountID & " and lt.For_DS_only <= " & str(IIf(sUserProfile = "CM_Admin", 1, 0)) & _
                    sContractIdClause & " group by lt.LetterType, lt.LetterDesc" & _
                " Union" & _
                " select lt.LetterType, lt.LetterDesc, ClaimCount=COUNT(1)" & _
                " from CMS_AUDITORS_CLAIMS.dbo.LETTER_Type lt with (nolock)" & _
                    " inner join CMS_AUDITORS_CLAIMS.dbo.auditclm_process_logics pl with (nolock) on lt.LetterType = pl.DataType and pl.AccountID = lt.accountid " & _
                    " inner join CMS_AUDITORS_CLAIMS.dbo.QUEUE_Exception qe with (nolock) on qe.ExceptionType=pl.comment" & _
                    " WHERE  lt.UseLegacySystem != 0 AND lt.isLetterActive = 1 and lt.accountid = " & gintAccountID & " and lt.For_DS_only <= " & str(IIf(sUserProfile = "CM_Admin", 1, 0)) & _
                    " and ProcessModule='LETTER' and comment like 'EX%'" & _
                    sContractIdClause & " group by lt.LetterType, lt.LetterDesc" & _
                " ) z" & _
                " group by LetterType,LetterDesc" & _
                " order by LetterType"



    Dim MyAdo As New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = strSQL
    Set cboLetterType.RecordSet = MyAdo.OpenRecordSet

    ' set up sender row source
    Me.lstSenders.RowSource = " select ID, Name, Type " & _
                              " from qry_AddToQueue_Senders_List " & _
                              " where AccountID = " & gintAccountID         '       & sContractIdClause
    Me.lstSenders.Requery

    Dim listloop As Integer
    For listloop = 1 To Me.lstSenders.ListCount
        If Me.lstSenders.ItemData(listloop) = Identity.UserName() Then
            Me.lstSenders.Selected(listloop) = True
            On Error Resume Next
            Me.lstSenders.ListIndex = listloop - 1 'JS 07/01/2013   ' uh.. whatever..
            
            Exit For
        End If
    Next

    Call lstSenders_AfterUpdate

    Me.lstClaims.RowSource = ""
    Me.txtClaimCount = ""
    Me.txtLtrCount = ""
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
    Else
        lblAppTitle.visible = True
    End If
End Sub

Private Sub lstClaims_Click()
    Dim clmcount As Integer: clmcount = 0
    Dim ltrcount As Integer: ltrcount = 0
    
    'update the count for the listbox.
    recountselected clmcount, ltrcount, Me.lstClaims, False
    Me.txtLtrCount = ltrcount
    Me.txtClaimCount = clmcount
End Sub

Private Sub lstClaims_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim clmcount As Integer: clmcount = 0
    Dim ltrcount As Integer: ltrcount = 0
    
    'update the count for the listbox.
    recountselected clmcount, ltrcount, Me.lstClaims, False
    Me.txtLtrCount = ltrcount
    Me.txtClaimCount = clmcount
End Sub

Private Sub recountselected(ByRef Claims As Integer, ByRef Providers As Integer, lstBox As listBox, CountAll As Boolean)
On Error GoTo err_msg:

     'This counts the letters
     Dim varItem
     Dim claimsct As Integer: claimsct = 0
     Dim provsct As Integer: provsct = 0
     Dim i As Integer: i = 0
     Dim dictProvs As Object
     
     Set dictProvs = CreateObject("Scripting.Dictionary")
     Dim rs As DAO.RecordSet
     
     If lstBox.RecordSet Is Nothing Then
        Claims = 0
        Providers = 0
         GoTo Clean_Up:
     End If
     
     If CountAll Then
        'This will pull us a count distinct from the table sql table linked to access. count of providernums.
             Set rs = CurrentDb.OpenRecordSet("select count(*) as Count from (SELECT ProvNum From Letter_Selection_temp A where username = '" & Identity.UserName & "' GROUP BY ProvNum) T;")
             If Not (rs.BOF And rs.EOF) Then
                provsct = rs("Count")
            End If
            
            Set rs = Nothing
            Set rs = CurrentDb.OpenRecordSet("Select count(*) as Count from Letter_Selection_temp A where username = '" & Identity.UserName & "' ")
             If Not (rs.BOF And rs.EOF) Then
                claimsct = rs("Count")
            End If
            Set rs = Nothing
     Else
        For Each varItem In lstBox.ItemsSelected
            If Not dictProvs.Exists(lstBox.Column(lstBox.RecordSet.Fields("Provnum").OrdinalPosition, varItem)) Then
                   provsct = provsct + 1
                   dictProvs.Add lstBox.Column(lstBox.RecordSet.Fields("Provnum").OrdinalPosition, varItem), ""
               End If
               claimsct = claimsct + 1
        Next varItem
     End If
    Claims = claimsct
    Providers = provsct
Clean_Up:
    Set rs = Nothing
    Set dictProvs = Nothing
    Exit Sub
err_msg:
    MsgBox ("error in recountselected sub - > " & Err.Description)
    GoTo Clean_Up:
End Sub
    
Private Sub SelectAllLButton_Click()
    Dim clmcount As Integer: clmcount = 0
    Dim ltrcount As Integer: ltrcount = 0
    Dim idx As Integer

    For idx = 1 To Me.lstClaims.ListCount
        Me.lstClaims.Selected(idx) = True
    Next idx
    'could add in detail view but left off b/c assuming analyst are just selecting all to print in batch processes - TH 10-8-07

    'update the count for the listbox.
    recountselected clmcount, ltrcount, Me.lstClaims, True
    Me.txtLtrCount = ltrcount
    Me.txtClaimCount = clmcount
End Sub

Private Sub Form_Close()
    Set MyAdo = Nothing
    Set co_ADO = Nothing
End Sub

Private Sub tglAdvancedFilter_Click()

'* FIlter on
    If Me.tglAdvancedFilter.Value = True Then
    
        Set frmFilter = New Form_frm_GENERAL_Filter
        
        With frmFilter
            .CalledBy = Me.Name
            .FieldsTable = Nz(DLookup("LetterSource", "LETTER_Type", "LetterType = '" & Me.cboLetterType & "'"), "")
            .Setup
            .visible = True
        End With

        Me.tglAdvancedFilter.Caption = "Filter On"

    Else
        Me.tglAdvancedFilter.Caption = "Add Filter"
        msAdvancedFilter = ""
        RefreshMain
    End If

End Sub

Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mdReturnDate = ReturnDate
End Sub


Private Sub CreatePK(ByVal TableName As String, ByVal Fields As String)
Dim iSetting As Integer
    iSetting = Application.GetOption("Error Trapping")
    Call Application.SetOption("Error Trapping", 2)

On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
    Call Application.SetOption("Error Trapping", iSetting)
End Sub

Private Sub Check_Letter_Views()
Dim rsLetterViews As ADODB.RecordSet

Dim strLocation As String
Dim strServer As String
Dim strDatabase As String
Dim strLetterView As String

Dim strChk As String
Dim strSQL As String
    
    On Error GoTo Err_handler
    
    If cofrmParent Is Nothing Then Set cofrmParent = Me.Parent.Form
    

    'get server and database name from connectstring in workfile linked table
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = "select * from LETTER_View_Relink"

    Set rsLetterViews = MyAdo.OpenRecordSet
    If rsLetterViews.recordCount > 0 Then
        rsLetterViews.MoveFirst
        While Not rsLetterViews.EOF
            strLocation = rsLetterViews("Location")
            strDatabase = rsLetterViews("DatabaseName")
            strServer = rsLetterViews("ServerName")
            strLetterView = rsLetterViews("LetterView")
            ' Do we need the Schema here?
            LinkTable "SQL", strServer, strDatabase, strLetterView
            strChk = "" & DLookup("Table", "Link_Table_Config", "Location = '" & strLocation & "' and [Server] = '" & strServer & "' and [Database] = '" & strDatabase & "' and [Table] = '" & strLetterView & "'")
            If strChk = "" Then
                strSQL = "INSERT INTO Link_Table_Config ( Location, Server, [Database], [Table] )" & _
                        " SELECT '" & strLocation & "' as Location, '" & strServer & "' as ServerName, '" & strDatabase & "' as DatabaseName, " & _
                        "'" & strLetterView & "' as TableName "
                
                CurrentDb.Execute (strSQL)
            End If
            rsLetterViews.MoveNext
        Wend
    End If


Exit_Sub:
    Set MyAdo = Nothing
    Set rsLetterViews = Nothing
    Exit Sub
    
Err_handler:
    MsgBox "ERROR: " & Err.Description, vbCritical
    Resume Exit_Sub
End Sub
