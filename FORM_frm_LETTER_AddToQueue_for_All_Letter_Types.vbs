Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private ColReSize3 As clsAutoSizeColumns
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

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

Const CstrFrmAppID As String = "LetterQueueAdd"
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
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

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
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
    
    DoCmd.Hourglass True
         
    'Checking if there is a souce for this letter (should be a SQL server view or Query)
    If Nz(DLookup("LetterSource", "LETTER_Type", "LetterType = '" & Me.cboLetterType & "'"), "") = "" Then
        GoTo Error_Encountered
    End If
    
    'Create a primary key on the Letter table to allow DAO inserts.
    'This "should" work anyway, if there is a PK on the table already, but there is no harm in it being here
    CreatePK "LETTER_Selection_Temp", "SelectionID"
    
        
    'Delete the current user's data from the temp letter table
    sqlString = "delete from LETTER_Selection_Temp where LETTER_Selection_Temp.username = '" & Identity.UserName & "'"
    
    'Create a new instance of the ADO class to handle the delete call we just defined in the query above
    Set MyAdo = New clsADO
       
    'Connect to the database and execute the delete
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = sqlString
    Result = MyAdo.Execute
          
    'Done with the ADO class, clean it out
    Set MyAdo = Nothing
    
    'Get the view or query source of the letter
    sFrom = Nz(DLookup("LetterSource", "LETTER_Type", "LetterType = '" & Me.cboLetterType & "'"), "")
    'Set the SQL string equal to nothing to reinitialize it
    Sstr = ""
   
    
    'TGH NOTES: 12/2/08
    ' this is pretty crazy the below code is how access executes an insert.
    'DPR NOTES 03/3/09 WOW I did not know that there was this bizzaro world SQL syntax for an insert!
    '***This SQL takes the data from the letter source and puts it into the Letter Selection temp table***
    'This is a local Access insert, and since we usually source letters from SQL, this should be moved to an ADO call.
    Sstr = "SELECT " & sFrom & ".*, '" & Identity.UserName & "' as UserName FROM " & sFrom & IIf(msAdvancedFilter <> "", " WHERE " & msAdvancedFilter, "")
    Sstr = "insert into LETTER_Selection_Temp(UserName) " & Sstr & " order by ProvNum"
    CurrentDb.Execute (Sstr)
    
    'TGH added 01/08/09 to ensure they are using their only table.
    'Code to initialize the list
    Set Me.lstClaims.RecordSet = Nothing
    Me.lstClaims.RowSource = vbNullString
    'Set our listbox columns to be the same as letter_selection_temp
    Me.lstClaims.ColumnCount = CurrentDb.TableDefs("Letter_selection_temp").Fields.Count
    'Set the data in the list equal to what is in the Selection table with any additional filters applied
    Me.lstClaims.RowSource = "SELECT * FROM LETTER_Selection_Temp " & IIf(msAdvancedFilter <> "", " WHERE " & msAdvancedFilter & " and Username = '" & Identity.UserName & "'", " WHERE Username = '" & Identity.UserName & "'")
 
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
    
   'resizing of the listbox.  max columns is 58 columns to be reformatted.
    Set ColReSize3 = New clsAutoSizeColumns
    ColReSize3.SetControl Me.lstClaims
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstClaims.ListCount - 1 > 0 Then
        ColReSize3.AutoSize
    End If
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
    On Error GoTo Exit_With_Error
    
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
    Dim dicProviderList As Object 'As New Dictionary
    
    'Connect to the databases with ADO
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    'Get a dictionary to keep track of providers
    Set dicProviderList = CreateObject("Scripting.Dictionary")
    
    'Damon 7/08
    'Some checking early on in the process
    'Just making sure all of the selections are made before moving on
    If lstClaims.ListCount = 0 Then
        MsgBox "There is no record to add to queue"
        Exit Sub
    ElseIf lstClaims.ItemsSelected.Count = 0 Then
        MsgBox "Letters must be selected."
        SelectAllLButton.SetFocus
        Exit Sub
    ElseIf IsNull(Me.dtLetterDate) Or Me.dtLetterDate = "" Then
        MsgBox "Please enter a letter request date"
        Me.dtLetterDate.SetFocus
        Exit Sub
    'CMS SPECIFIC RULE
    ElseIf CDate(Nz(Me.dtLetterDate, "")) < Date Then
        MsgBox "Letter Request Date cannot be a past date."
        Me.dtLetterDate.SetFocus
        Exit Sub
    ElseIf Me.lstSenders.ListIndex = -1 Then
        MsgBox "Please choose a sender"
        Me.lstSenders.SetFocus
        Exit Sub
    End If

    'Damon 7/08
    'Setup progress screen
    intStatus = 1
    Set fmrStatus = New Form_ScrStatus
    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        .ProgMax = lstClaims.ItemsSelected.Count - 1
        .TimerInterval = 50
        .show
    End With
    
    'Damon 7/08
    'Start a Transaction
    
    myCode_ADO.BeginTrans
    'Get a new unique identifier that will be associated with this batch of letters
    strSessionID = GetInstanceID
    
   
    'Loop through what is selected in the list
    For Each varItem In Me.lstClaims.ItemsSelected
        'Move through the recordset of work claims and associate an auditor with a provider
        'While Not rstLetterWorkTable.EOF
        If j > UBound(mstrAuditorsArray) Then
           j = 0
        End If
        
        'DPR creates an association of provider and auditor in a dictionary to allow multiple auditors to split the letters evenly
        'TGH took out code to loop through the recordest because if we are looping already just run this each time first so we know if we have a new provider it will be in the
        'array before we run the below code.  So removing an extra loop thorough the recordset.
        If Not dicProviderList.Exists(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("CnlyProvID").OrdinalPosition, varItem)) Then
            dicProviderList.Add Trim(Me.lstClaims.Column(Me.lstClaims.RecordSet.Fields("CnlyProvID").OrdinalPosition, varItem)), mstrAuditorsArray(j)
            j = j + 1
        End If
              
        'Call usp_LETTER_Work_Table_Insert via ADO.Execute
        'Also passing in the auditor based off of the auditors dictionary so we can handle multiple auditors assigned.
        myCode_ADO.sqlString = "usp_LETTER_Work_Table_Insert"
        myCode_ADO.SQLTextType = StoredProc
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
        Set prm = LocCmd.CreateParameter("@pLetterReqDt", adDBDate, adParamInput, , CStr(Format(Now(), "MM/DD/YYYY")))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@ErrMsg", adVarChar, adParamOutput, 255)
        LocCmd.Parameters.Append prm
        iReturn = myCode_ADO.Execute(LocCmd.Parameters)
        strErrMsg = Nz(LocCmd.Parameters("@ErrMsg").Value, "")
        Set LocCmd = Nothing
        If strErrMsg <> "" Then
            GoTo Exit_With_Error
        End If
                 
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
                GoTo Exit_With_Error
       End If

       'Damon 7/08
       'I have no idea.....
       'this is to communicate with the application running the progress screen,  doing this over and over to allow for active refresh of the progress screen.
       'DR - Yes but why do you have to call DOEVENTS 3000 times?
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
    
    Next varItem
    
    'Calling usp_LETTER_Work_Queue_Mass_Insert.  This procedure takes everything that was placed in the Work table and adds it to the letter system to be printed
    'This needs to be called for each "Auditor" that was associated with a letter, hence we are looping through the dictionary paring
    j = 0
    While j <= UBound(mstrAuditorsArray)
        myCode_ADO.sqlString = "usp_LETTER_Work_Queue_Mass_Insert"
        myCode_ADO.SQLTextType = StoredProc
        Set prm = LocCmd.CreateParameter("@SessionID", adVarChar, adParamInput, 20, strSessionID)
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("Auditor", adVarChar, adParamInput, 75, mstrAuditorsArray(j))
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@InstanceID", adVarChar, adParamOutput, 20)
        LocCmd.Parameters.Append prm
        Set prm = LocCmd.CreateParameter("@ErrMsg", adVarChar, adParamOutput, 255)
        LocCmd.Parameters.Append prm
        iReturn = myCode_ADO.Execute(LocCmd.Parameters)
        strErrMsg = Nz(LocCmd.Parameters("@ErrMsg").Value, "")
        Set LocCmd = Nothing
    
    'Damon 7/08
    'I need to define what I am doing here, what I want is that if this fails, kill the whole thing
    'Select Case iReturn
    If strErrMsg <> "" Then
        GoTo Exit_With_Error
    End If
      j = j + 1
    Wend
    
    'Commit the transaction!
    myCode_ADO.CommitTrans
    
    'Demolish the status form
    Set fmrStatus = Nothing
    
    'Refresh the letters!
    RefreshMain
    
Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Exit Sub
    
Exit_With_Error:
    If strErrMsg <> "" Then
      MsgBox strErrMsg, vbCritical
    Else
      MsgBox Err.Description, vbCritical
    End If
    myCode_ADO.RollbackTrans
    Resume Exit_Sub
End Sub
Private Sub cmdAddLettersToQueue_Click()
    AddLettersToQueue
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

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    ' relink letter view if applicable
    Check_Letter_Views
    
    ' set up letter type row source
    Me.cboLetterType.RowSource = " select LetterType, LetterDesc " & _
                                 " from Letter_Type " & _
                                 " where AccountID = " & gintAccountID
    Me.cboLetterType.Requery
    
    
    ' set up sender row source
    Me.lstSenders.RowSource = " select ID, Name, Type " & _
                              " from qry_AddToQueue_Senders_List " & _
                              " where AccountID = " & gintAccountID
    Me.lstSenders.Requery
                                
                                
    
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
             Set rs = CurrentDb.OpenRecordSet("select count(*) as Count from (SELECT ProvNum From Letter_Selection_temp where username = '" & Identity.UserName & "' GROUP BY ProvNum) T;")
             If Not (rs.BOF And rs.EOF) Then
                provsct = rs("Count")
            End If
            
            Set rs = Nothing
            Set rs = CurrentDb.OpenRecordSet("Select count(*) as Count from Letter_Selection_temp where username = '" & Identity.UserName & "' ")
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
    Set myCode_ADO = Nothing
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
On Error Resume Next
    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & TableName & " On " & TableName & "(" & Fields & ")"
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
