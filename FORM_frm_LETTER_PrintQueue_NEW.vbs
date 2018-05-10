Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




''' Last Modified: 09/18/2014
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This Print Letter Form is to take all the letters selected and generate letters.
'''     Status: 'W' means it's been queued up and ready to be generated
'''             'G' means it's been generated
'''             'P' means it's been printed
'''             'R' means it's ready to be re-generated (then re-printed)
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 9/19/2014    KD: Changed this to point to the new Letter_Print_Queue tables
'''     so this can function as the manual override
'''
'''  - 08/05/2013 - KD: Added sorting to print the letters with the most pages last
'''     in order to make sure that the operator doesn't have to baby sit the envelope
'''     stuffer
'''     Also, added a mail merge for letter lables for the documents with > 10 (or was it 8) pages
'''     which have to be sent in manila envelopes so operator doesn't have to manually type them up
'''
'''
'''  - 04/16/2013 - added logging throughout and more tweaking to barcode stuff
'''         also minor cleanup of some of the code (ado stuff was early bound, badly
'''         indented, etc..)
'''  - 03/28/2013 - added some logging and very minor tweaking to the barcode stuff..
'''  - 02/19/2013 - Added AddSecPagesCode and some minor other adjustments for adding
'''     bar codes to letter's footer for the envelop stuffer
'''  - ?????? - Created...
'''
''' AUTHOR
'''  =====================================
''' ????? Thieu ??????
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################

'
Private mstrAuditor As String

Private Const cs_USER_TEMPLATE_PATH_ROOT As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\"

Private cdctSelectedLetters As clsLetterInstanceDct

'for resizing
Private ColResize1 As clsAutoSizeColumns
Private ColReSize2 As clsAutoSizeColumns
Private lngQueryType As Long '* These are values from msysobjects 1/4/6 = table, 5 = query


Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents frmFilter As Form_frm_GENERAL_Filter
Attribute frmFilter.VB_VarHelpID = -1

Private mReturnDate As Date
Private msAdvancedFilter As String

Private clBatchId As Long
Private clBatchColumn As Long
Private cdctQueueColumns As Scripting.Dictionary
Private cdctBatches As Scripting.Dictionary

Private ccolErrors As Collection
Private gbVerboseLogging As Boolean

Const CstrFrmAppID As String = "LetterQueuePrint"


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Public Property Get FilteredByUngenerated() As Boolean
    If UCase(Me.cboViewType) = "VIEW UN-PROCESSED LETTERS" Then
        FilteredByUngenerated = True
    Else    'If UCase(Me.cboViewType) = "VIEW PROCESSED LETTERS" Then
        FilteredByUngenerated = False
    End If
End Property

Public Property Get FilteredByGenerated() As Boolean
    If UCase(Me.cboViewType) = "VIEW PROCESSED LETTERS" Then
        FilteredByGenerated = True
    Else
        FilteredByGenerated = False
    End If
End Property

Public Property Get FilteredByErrors() As Boolean
     If UCase(Me.cboViewType) = "VIEW ERRORS" Then
         FilteredByErrors = True
     Else
         FilteredByErrors = False
     End If
End Property


Public Property Let NumSelected(lNumberSelected As Long)

'    Me.lblNumSelected.Caption = Format(Me.lstQueue.ItemsSelected.Count, "###,###") & " selected items"

    Me.lblNumSelected.Caption = Format(Nz(lNumberSelected, 0), "###,##0") & " items selected"
End Property


Public Property Get MostRecentBatchId() As Long
    MostRecentBatchId = clBatchId
End Property
Public Property Let MostRecentBatchId(lBatchId As Long)
    clBatchId = lBatchId
    Me.lblRecentBatchId.Caption = "Most recent batchid: " & CStr(clBatchId)
    If clBatchId = 0 Then
        Me.lblRecentBatchId.visible = False
    Else
        Me.lblRecentBatchId.visible = True
    End If
End Property


'Private Sub cboAuditor_AfterUpdate()
'    If Me.cboAuditor <> "View All" Then
'        Me.cboAuditor.ForeColor = 16711680
'    Else
'        Me.cboAuditor.ForeColor = 0
'    End If
'    RefreshMain
'End Sub

Private Sub cboLetterType_AfterUpdate()
    If Me.cboLetterType <> "View All" Then
        Me.cboLetterType.ForeColor = 16711680
    Else
        Me.cboLetterType.ForeColor = 0
    End If
    RefreshMain
End Sub

Private Sub cboViewType_AfterUpdate()

    Me.cboViewType.ForeColor = 0

    ' 20130815 KD: Ok, so, if they choose Unprocessed then we are going to change the
    ' 'View Letter(s)' button to read: PRE-view Letter(s)'
    ' otherwise we'll say Print Letters (or something like that..)
    Select Case UCase(Me.cboViewType)
    Case "VIEW UN-PROCESSED LETTERS"
        Me.cmdViewLetters.Caption = "PRE-View Letter(s)"
        
        RefreshMain
        Me.cmbStatus = "Q"
        Me.cmbBatches = ""
        Call cmdSelectByStatus_Click
    
    Case "VIEW PROCESSED LETTERS"
        Me.cmdViewLetters.Caption = "Print Letter(s)"
        ' should also select the 'G' status

        RefreshMain
        
        If Me.MostRecentBatchId <> 0 Then
            Me.cmbStatus = ""
            Me.cmbBatches = Me.MostRecentBatchId
            Call cmdSelectByBatch_Click
        Else
            Me.cmbStatus = "G"
            Me.cmbBatches = ""
            
            Call cmdSelectByStatus_Click
        End If

    Case "View Errors"
        Stop
    End Select

    
End Sub





Private Sub ckManualOnly_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Nz(Me.fraManualOnly, 0) = 0 Then
        Me.fraManualOnly = 1
    Else
        Me.fraManualOnly = 0
    End If
End Sub

Private Sub cmbBatches_Change()
    Call cmdSelectByBatch_Click
End Sub


Private Sub cmbStatus_Change()
    Call cmdSelectByStatus_Click
End Sub

Private Sub cmdSelectByBatch_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oListBox As listBox
Dim lRow As Long
Dim vItem As Variant
Dim lSelBatchid As Long

    strProcName = ClassName & ".cmdSelectByBatch_Click"
    
    lSelBatchid = Me.cmbBatches
    If cdctQueueColumns Is Nothing Then Call GetQueueColumns
                
    clBatchColumn = cdctQueueColumns.Item("BATCHID")
    For lRow = 0 To Me.lstQueue.ListCount
        If Me.lstQueue.Column(clBatchColumn, lRow) = lSelBatchid Then
            Me.lstQueue.Selected(lRow) = True
        Else
            Me.lstQueue.Selected(lRow) = False
        End If
    Next
    
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub




Private Sub GetQueueColumns()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetQueueColumns"
  
    Set cdctQueueColumns = New Scripting.Dictionary
    Set oRs = Me.lstQueue.RecordSet
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)
    
Block_Exit:
    Set oRs = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSelectByStatus_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oListBox As listBox
Dim lRow As Long
Dim vItem As Variant
Dim sSelectCd As String

    strProcName = ClassName & ".cmdSelectByStatus_Click"
        
    sSelectCd = Me.cmbStatus
    If cdctQueueColumns Is Nothing Then Call GetQueueColumns
                
    clBatchColumn = cdctQueueColumns.Item("STATUS")
    For lRow = 0 To Me.lstQueue.ListCount
        If Me.lstQueue.Column(clBatchColumn, lRow) = sSelectCd Then
            Me.lstQueue.Selected(lRow) = True
        Else
            Me.lstQueue.Selected(lRow) = False
        End If
    Next
    
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub Form_Resize()
    Me.lstQueue.Width = Me.InsideWidth - Me.lstQueue.left - 300
'    Me.lstQueue.Height = Me.Detail.Height - (Me.lstQueue.top + 300)
    
    Me.lstQueue.Height = Me.InsideHeight - (Me.lstQueue.top + 3000)
    
    Me.lblAppTitle.Width = Me.InsideWidth
End Sub

Private Sub frmfilter_QueryFormRefresh()
    RefreshMain
End Sub

Private Sub frmFilter_UpdateSql()
    msAdvancedFilter = frmFilter.SQL.WherePrimary
End Sub


Private Sub cmdDeleteQueue_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim strPowerUsers As String
Dim bPowerUser As Boolean
Dim strFileName As String

' ADO variables & late bind them
Dim cn As ADODB.Connection
Dim cmd As ADODB.Command
Dim cmdGetLetter As ADODB.Command
    
Set cn = New ADODB.Connection
Set cmd = New ADODB.Command
Set cmdGetLetter = New ADODB.Command

Dim strSQLcmd As String
Dim strInstanceID As String
Dim strStatus As String
Dim strErrMsg As String
Dim iRtnCd As Integer
Dim i As Integer

Dim varItem
    
    strProcName = ClassName & ".cmdDeleteQueue_Click"
    
    'User Entry Needed***************************************************************
    'These users have total power and can delete printed letters
    'individual users can delete these when they have a status of 'W'
    'strPowerUsers = UCase("Alex.Dremann|Damon.Ramaglia|Thieu.Le|Joe.Casella|Tom.Hartey|Robert.Swander")
    'If InStr(1, strPowerUsers, UCase(Person.UserName)) > 0 Then bPowerUser = True
    
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected"
        GoTo Block_Exit
    End If
    
    bPowerUser = True
    
    'If bPowerUser = True Then
    '    MsgBox "You are a POWER user!!", vbInformation
    i = MsgBox("Power user: Are you sure you want to delete these records?", vbYesNo)
    If i <> vbYes Then Exit Sub
    
    
    'MsgBox "You are a POWER user!!", vbInformation
    'i = MsgBox("Power user: Are you sure you want to delete these records?", vbYesNo)
    'If i <> vbYes Then Exit Sub
    
    cn.ConnectionString = CodeConnString
    cn.CommandTimeout = 0
    cn.Open
    cn.CursorLocation = adUseClient
    
    'Begin our transactions
    cn.BeginTrans
    
    'CmdGetLetter setup before we enter the loop
    ' this first one gets the path of the letter (in field, 'LetterName' in the old system)
    
    cmdGetLetter.ActiveConnection = cn
    cmdGetLetter.commandType = adCmdStoredProc
    cmdGetLetter.CommandText = "usp_LETTER_Automation_Get_Letter_Path"
    cmdGetLetter.Parameters.Refresh
    
    
    'setup command for usp_LETTER_Work_Queue_Forced_Delete
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = cn
    cmd.commandType = adCmdStoredProc

    If bPowerUser Then
        cmd.CommandText = "usp_LETTER_Automation_Forced_DELETE"
    Else
        cmd.CommandText = "usp_LETTER_Automation_DELETE"
    End If
    cmd.Parameters.Refresh
    
    
    
    Dim FileArrays() As String
    ReDim Preserve FileArrays(lstQueue.ItemsSelected.Count - 1)
    'initilize the loop
    i = 0
    'So for each item selected go through and store the file path then run the delete stored proc.
    'If we error out before the end of the sub the rollback will correct the sql and the kill statement will
    'not be executed.
    
    Dim MyRecordset As ADODB.RecordSet
    Set MyRecordset = Me.lstQueue.RecordSet
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(MyRecordset)
    For Each varItem In lstQueue.ItemsSelected

'        strInstanceID = Trim(Me.lstQueue.Column(MyRecordset.Fields("InstanceID").OrdinalPosition, varItem))
'        strStatus = Trim(Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem))
        

        strInstanceID = Trim(Me.lstQueue.Column(cdctQueueColumns("INSTANCEID"), varItem))
        strStatus = Trim(Me.lstQueue.Column(cdctQueueColumns("STATUS"), varItem))
        'if the user is a power user or the case is W, E or R aka Not Printed
        'weird logic
'        If InStr(1, "W|E|R|G", UCase(strStatus)) > 0 Or bPowerUser Then
        If InStr(1, "Q|QR|W|E|R|G", UCase(strStatus)) > 0 Or bPowerUser Then
            If UCase(strStatus) = "P" Then  'if printed already and we are a power user
                'get the LetterPath and add it to array of letters to be printed
                cmdGetLetter.Parameters("@pInstanceID").Value = strInstanceID
                cmdGetLetter.Execute
                'only if the letter has been printed do we need to populate the array to delete files.
                FileArrays(i) = Trim(Nz(cmdGetLetter.Parameters("@pLetterPath").Value, ""))
                i = i + 1
            End If
            
            cmd.Parameters("@pInstanceID").Value = strInstanceID
            cmd.Execute
            
            iRtnCd = cmd.Parameters("Return").Value
            strErrMsg = Trim(Nz(cmd.Parameters("@pErrMsg").Value, ""))
            
            If iRtnCd <> 0 Or strErrMsg <> "" Then
                GoTo Block_Err
            End If
        End If
    Next varItem
    cn.CommitTrans 'we went through the entire process.
    
 
Block_Exit:

    Set cmd = Nothing
    Set cmdGetLetter = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Set Me.lstQueue.RecordSet = Nothing
    RefreshMain
    Exit Sub
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg, , True
    Else
        ReportError Err, strProcName
    End If
    If Not cn Is Nothing Then cn.RollbackTrans
    GoTo Block_Exit
End Sub

Private Sub cmdEndDate_Click()
    On Error GoTo Exit_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtThroughDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtThroughDate = mReturnDate
    
    txtThroughDate_Exit False

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub



Private Sub cmdPrintSelectedItems_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim fmrStatus As Form_ScrStatus
Dim iTotalRecs As Integer
Dim oRs As DAO.RecordSet
Dim lBatchId As Long
Dim bAtLeastOneErrored As Boolean
Dim sErrMsg As String
Dim sOrigPrinter As String

    strProcName = ClassName & ".cmdPrintSelectedItems_Click"
    
    ' ensure we have some selections
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "Nothing is selected!", vbCritical, "Please select something to generate!"
        GoTo Block_Exit
    End If
    
    Set fmrStatus = New Form_ScrStatus
    
    
    ' 20130809 KD: So a change to the process...
    ' The Generate button (this button) is ONLY going to generate the letters
    ' not also view them (they weren't using that anyway)
    ' Additionally, we'll generate a batch id for each batch of letters that are generated
    ' that will help them to select them for viewing (printing) easier
    ' I'll add a drop down and a notification to let them know what THEIR latest batch id was
    ' Status 'G' (Generated, not yet printed) should not be re-generated..
    
    Set cdctSelectedLetters = GetSelectedItems(, , "G", sErrMsg)
    If cdctSelectedLetters Is Nothing Then
        bAtLeastOneErrored = True
        GoTo Block_Exit
    End If
    iTotalRecs = cdctSelectedLetters.Count
    
    If iTotalRecs > 0 Then
'        lBatchId = GenerateBatchId()
        lBatchId = AssignBatchId(cdctSelectedLetters)
    Else
        LogMessage strProcName, "WARNING", "Nothing to process..." & vbCrLf & "('G' status cannot be regenerated!)", sErrMsg, True
        bAtLeastOneErrored = True
        GoTo Block_Exit
    End If
    
    If lBatchId = 0 Then
        LogMessage strProcName, "ERROR", "No batchid generated!"
        bAtLeastOneErrored = True
        GoTo Block_Exit
    End If
    
    DoCmd.Hourglass True
    
    '' set the default printer to Adobe Acrobat
    '' and grab the name of what the NORMAL default printer is:


    Call SetDefaultPrinterToAcrobatAPI(sOrigPrinter)
    
    
    cdctSelectedLetters.BatchID = lBatchId
    
    '' KD: I'm going to leave this here to make sure that it works since I changed things.. The Stop command will show me what's up
    If Me.lstQueue.ItemsSelected.Count = 1 And iTotalRecs = 0 Then
            ' Not exactly true here.. It has to be W or R status as well
            Stop        ' Why???? Fix this. need to make sure that we get all of them in GetSelectedItems()
        iTotalRecs = 1 'JS Change 20130305 I had to add this because for some reason it will not detect only one item selected
    End If
    
    
    ' delete all templates in temp directory
    Call DeleteTemplates

    'BEGIN THE PROGRESS FORM, MAKE THE MAX PROGRESS DOUBLE THE ITEMS SELECTED B/C WE HAVE TO TOUCH EACH REC TWICE WHEN GENERATING.
    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgMax = iTotalRecs * 2
        .TimerInterval = 50
        .show
    End With

   
    ' Print letter
    If GenerateLetters(fmrStatus, bAtLeastOneErrored, lBatchId) = False Then
        LogMessage strProcName, "ERROR", "Print letters failed. Please check logs for details"
       
        Call UpdateBatchId(lBatchId, False)
        bAtLeastOneErrored = True
        GoTo Block_Exit
    End If
    
    '' KD: OK now show the user their batch id..
    '' actually, we should mark it as completed without error
    Call UpdateBatchId(lBatchId, True)
    


    '' kd comeback: select the batch id in the drop down as soon as you add it (numbscull)!
    ' but first we need to change the ProcessType to Processed:
    Me.cboViewType = "View Processed Letters"
    

    Call RefreshMain
    
    Me.cmbBatches = lBatchId
    Call cmdSelectByBatch_Click
    
    
    fmrStatus.visible = False
    
    
Block_Exit:
    Call SetDefaultPrinterToAcrobatAPI("", sOrigPrinter)
    
    
    DoCmd.Hourglass False
    Call NotifyUserOfErrors
    
    If bAtLeastOneErrored = True Then
        If sErrMsg <> "" Then
            LogMessage strProcName, "ERROR", sErrMsg, , True
        Else
            MsgBox "Finished with at least 1 error!" & vbCrLf & "To print the letters, select the ones you wish to print and click the View Letter(s) button."
        End If
        
        '2014:05:01:JS
        'Making it refresh because i have seen cases where the generation failed for some instances only and the listbox is not updating them, therefore the operator tries to generate all of them again and the ones that were sucessfully generated now will throw an error
        RefreshMain
        
    Else
        MsgBox "Finished!" & vbCrLf & "To print the letters, select the ones you wish to print and click the View Letter(s) button. This batch has automatically been selected for you."
    End If
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Call UpdateBatchId(lBatchId, False)
    GoTo Block_Exit
End Sub



Private Function PrintSelectedItems(cn As Variant, TotalRecs As Integer, fmrStatus As Form_ScrStatus) As Boolean
'On Error GoTo Error_encountered
'set the function to False in the beginning.  This ensures we get to the end to set it as correct.
PrintSelectedItems = False

Dim PrintFileArray() As String: ReDim Preserve PrintFileArray(TotalRecs - 1)
Dim UpdateLettersArray() As String: ReDim Preserve UpdateLettersArray(TotalRecs - 1)
'    Dim Person As New ClsIdentity
' ADO variables - Late binded here.
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

Dim cmdGetLetter As ADODB.Command
Set cmdGetLetter = New ADODB.Command
        
Dim strSQLcmd As String

' Letter configuration variables
Dim db As Database
Dim rsLetterConfig As DAO.RecordSet
Dim strODCFile As String
Dim strBasedPath As String
Dim colLetterTemplate As Collection
Dim objLetterInfo As New clsLetterTemplate
    
    ' Word objects setup as variants b/c of late binding
'    Dim objWordApp As Word.Application, _
'        objWordDoc As Word.Document, _
'        objMasterDoc As Word.Document, _
'        objWordMergedDoc As Word.Document
        
Dim objWordApp, _
    objWordDoc, _
    objMasterDoc, _
    objWordMergedDoc
Dim strProcName As String
    Set objWordApp = CreateObject("word.application")
'    Set objWordApp = New Word.Application
    
    objWordApp.visible = False
    
    'Letter generation variables
    Dim rsProvList As ADODB.RecordSet
    Set rsProvList = New ADODB.RecordSet
    
    Dim rsLetterTemplate As ADODB.RecordSet
    Set rsLetterTemplate = New ADODB.RecordSet

    Dim strInstanceID As String
    Dim strProvNum As String
    Dim strAuditor As String
    Dim strLetterType As String
    Dim dtLetterReqDt As Date
    Dim strStatus As String
    Dim strLocalTemplate As String
    Dim strLocalPath As String
    
    Dim bMergeError As Boolean
    Dim strOutputPath As String
    Dim strOutputFileName As String
    Dim strChkFile As String
    Dim strErrMsg As String
    Dim iRtnCd As Integer
    
    Dim varItem As Variant
    Dim iAnswer As Integer
    Dim bFirstLetter As Boolean
    Dim iCnt As Integer
    Dim i As Integer
    
    Dim MyAdo As clsADO
    
    strProcName = ClassName & ".PrintSelectedItem"
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = DataConnString
    
    cmd.ActiveConnection = cn 'open connection passed in
    bFirstLetter = True
    
    strErrMsg = ""

    'set local path
    'USER ENTRY NEEDED
    strLocalPath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    'End USER ENTRY NEEDED
    If Not FolderExist(strLocalPath) Then CreateFolder (strLocalPath)
        
    ' get list of templates
    Set colLetterTemplate = New Collection
    'TL add account ID logic
    strSQLcmd = "select LetterType, TemplateLoc from LETTER_Type where AccountID = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = strSQLcmd
    'cmd.ActiveConnection = cn 'open connection passed in
    'cmd.CommandText = strSQLCmd
    'cmd.CommandTimeout = 0
    'cmd.CommandType = adCmdText
    'Set rsLetterTemplate = cmd.Execute
    Set rsLetterTemplate = MyAdo.OpenRecordSet

    'rst is a recset of all templates and locations.
    'Then populates the collection colLetterTemplate with the info from these templates.
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            objLetterInfo.LetterType = Trim(UCase(![LetterType]))
            objLetterInfo.TemplateLoc = Trim(![TemplateLoc])
            colLetterTemplate.Add objLetterInfo, Trim(UCase(![LetterType]))
            strLocalTemplate = strLocalPath & "\" & GetFileName(objLetterInfo.TemplateLoc)
            On Error Resume Next
                'see if the template exists at the strLocalPath if it does we are deleting it to make room for the copy over.
                Kill strLocalTemplate
            On Error GoTo Error_Encountered
            Set objLetterInfo = New clsLetterTemplate
            .MoveNext
        End With
    Loop
    rsLetterTemplate.Close
    Set rsLetterTemplate = Nothing
    
    ' Set the based path for saving merge doc
    Set db = CurrentDb
    'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    
    strBasedPath = rsLetterConfig("LetterOutputLocation").Value
    strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    
    bMergeError = False
      
    cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("LetterName", adChar, adParamInput, 255, "")
    cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adChar, adParamOutput, 255, "")
    
    Set objLetterInfo = New clsLetterTemplate
    
    ' setup progress screen that is passed to this function
    Dim sMsg As String
    Dim lngProgressCount As Long
    Dim msgIcon As Integer
    Dim ObjectExists As Boolean
    
    ' start processing letters
    iCnt = 0
    
    Dim MyRecordset As DAO.RecordSet
    Set MyRecordset = Me.lstQueue.RecordSet
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    For Each varItem In lstQueue.ItemsSelected
        If Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem) = "Q" Or _
           Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem) = "r" Then
            strInstanceID = Trim(Me.lstQueue.Column(MyRecordset.Fields("InstanceID").OrdinalPosition, varItem))
            strProvNum = Trim(Me.lstQueue.Column(MyRecordset.Fields("cnlyProvID").OrdinalPosition, varItem))
            strLetterType = UCase(Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem)))
            dtLetterReqDt = Me.lstQueue.Column(MyRecordset.Fields("LetterReqDt").OrdinalPosition, varItem)
            strAuditor = Trim(Me.lstQueue.Column(MyRecordset.Fields("Auditor").OrdinalPosition, varItem))
            strStatus = Trim(Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem))
            iCnt = iCnt + 1
    
                If strLetterType <> objLetterInfo.LetterType Then
                    'We are dealing with a new report in the queue!
                    Set objLetterInfo = colLetterTemplate(strLetterType)
                     
                    ObjectExists = False
                    'Check to see if we are talking about a report or Template
                    If InStr(1, objLetterInfo.TemplateLoc, ".doc", vbBinaryCompare) = 0 Then 'if we are dealing with an access rpt
                        'make sure the report exits in access.
                        For i = 0 To db.Containers("Reports").Documents.Count - 1
                            If db.Containers("Reports").Documents(i).Name = objLetterInfo.TemplateLoc Then
                            ObjectExists = True
                            End If
                        Next i
                    
                        If ObjectExists = False Then
                            strErrMsg = "Missing letter template in access." & vbCrLf & "Template Report name = " & objLetterInfo.TemplateLoc & ""
                            GoTo Error_Encountered
                        End If
                    Else ' we have a Word template we are working from (WORD MAIL MERGE)
                        
                        'check if the template physically exists
                        strChkFile = Dir(objLetterInfo.TemplateLoc)
                        If strChkFile = "" Then
                            strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
                            GoTo Error_Encountered
                        End If
                    
                        
                       ' make a local copy so it would not impact other users
                        strLocalTemplate = strLocalPath & "\" & GetFileName(objLetterInfo.TemplateLoc)
                        strChkFile = Dir(strLocalTemplate)
                        If strChkFile = "" Then
                            FileCopy objLetterInfo.TemplateLoc, strLocalTemplate
                        End If
                       
                        'May put this back in someday.  i believe this is being done automatically via the MailMerge.
                        ' open template or objWordDoc and set margins to the objmasterdoc.
                        'When the mail merge runs it keeps the template's Margins....
                        Set objWordDoc = objWordApp.Documents.Add(strLocalTemplate, , False) 'tried didn't effect change
                        'Set objWordDoc = objWordApp.Documents.Open(strLocalTemplate)
                        '                    objMasterDoc.PageSetup.LeftMargin = objWordDoc.PageSetup.LeftMargin
                        '                    objMasterDoc.PageSetup.RightMargin = objWordDoc.PageSetup.RightMargin
                        '                    objMasterDoc.PageSetup.TopMargin = objWordDoc.PageSetup.TopMargin
                        '                    objMasterDoc.PageSetup.BottomMargin = objWordDoc.PageSetup.BottomMargin
                        'objWordDoc.Select
                        'objWordDoc.spellingchecked = True
                        'open template
                        
                    End If  'End of split for access report or word mail merge
                End If 'end of check to see if the template report or word doc exists
                
                ' NOW CHECK TO SEE IF THE TEMPLATE IS NOT A WORD DOCUMENT
                If InStr(1, objLetterInfo.TemplateLoc, ".doc", vbTextCompare) <> 0 Then 'If we are dealing with a Word Mail Merge
                
                        ' Create temp table to address issues with Word VBA/sp multiple selection criterion.
                        
                        AdoExeTxt "usp_LETTER_Get_Info_load '" & strInstanceID & "',''", "v_CODE_Database"
                                    
                        ' Set data source for mail merge.  Data will be from new Temp Table
                        objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, _
                            SqlStatement:="exec usp_LETTER_Get_Info '" & strInstanceID & "'"
                        
                        ' Perform mail merge.
                        objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
                        objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
                        objWordDoc.MailMerge.Execute Pause:=False
                        If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
                            objWordApp.visible = True
                            'MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
                            Call ErrorCallStack_Add(clBatchId, "Error encountered with mail merge.", strProcName, "Instance ID: " & strInstanceID, , , strInstanceID, objLetterInfo.LetterType)
                            bMergeError = True
                            objWordApp.ActiveDocument.Activate
                            GoTo Cleanup
                        End If
                        ''------------------- here is where we convert to pdf instead of word ----------------''
                        ' Save the output doc
                        Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
                        strOutputPath = strBasedPath & "\" & strProvNum & "\"
                        Call CreateFolder(strOutputPath)
    
                        'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                        'Added to rename reprints...
                        If strStatus = "R" Then
                            strOutputFileName = strOutputPath & "" & strLetterType & "-Reprint-" & strInstanceID & ".doc"
                        Else
                            strOutputFileName = strOutputPath & "" & strLetterType & "-" & strInstanceID & ".doc"
                        End If
                        objWordMergedDoc.spellingchecked = True
                        Sleep 1000
                        objWordMergedDoc.SaveAs strOutputFileName
                        objWordMergedDoc.Close
                        
                        DoEvents
                        DoEvents
                        
                         '2014:04:25:JS:sometimes the letter file is not created! +-
                         If Not FileExists(strOutputFileName) Then
                             GoTo Error_Encountered
                         End If
                        
                        
                        'save the file location we just generated in case user cancels or error
                        PrintFileArray(iCnt - 1) = strOutputFileName
    
                        Set objWordMergedDoc = Nothing
    
                        'save the word doc's in original form.  could convert these to pdf if we are so inclined...
                        'ConvertWordToPDF Replace(strOutputFileName, ".doc", ".PDF"), strOutputFileName
                Else
'Stop
'                    'we are working with an access report now.
'                    'load report info into temp table...
'                    'Set the output path, and create the folder if it does not exist
'                    'then we call ConvertRPTtoPDF to save the access report as a pdf image.
'                    AdoExeTxt "usp_LETTER_Get_Info_load '" & strInstanceID & "',''", "v_AMERIHEALTH_Auditors_Code"
'                    strOutputPath = strBasedPath & "\" & strProvNum
'                    CreateFolder (strOutputPath)
'                    'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".PDF"
'                    If strStatus = "r" Then
'                         strOutputFileName = strOutputPath & "\" & strLetterType & "-Reprint-" & strInstanceID & ".PDF"""
'                    Else
'                         strOutputFileName = strOutputPath & "\" & strLetterType & "-" & strInstanceID & ".PDF"
'                    End If
'
'                    If Not ConvertRPTToPDF(strOutputFileName, objLetterInfo.TemplateLoc) Then
'                        strErrMsg = "pdf conversion took too long to run, most likely a printing spooling issue, please try again"
'                        GoTo Error_Encountered
'                    End If
'
'                    'save the file location we just generated in case user cancels or error
'                    PrintFileArray(iCnt - 1) = strOutputFileName
'
'                    'DoCmd.OutputTo acOutputReport, objLetterInfo.TemplateLoc, acFormatRTF, strOutputFileName, False
                End If
                         
            ' update status
Stop
            cmd.Parameters("InstanceID").Value = strInstanceID
            cmd.Parameters("LetterName").Value = strOutputFileName
            strSQLcmd = "usp_LETTER_Update_Status"
            cmd.CommandText = strSQLcmd
            cmd.commandType = adCmdStoredProc
            cmd.Execute

            strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
            If strErrMsg <> "" Then
                GoTo Error_Encountered
            End If
            
             'save the file location we just generated in case user cancels or error
            UpdateLettersArray(iCnt - 1) = strInstanceID
           
    
            ' display progress
            sMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
                        "Provider = " & strProvNum & vbCrLf & _
                        "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
            fmrStatus.ProgVal = iCnt
            fmrStatus.StatusMessage sMsg
    
            If fmrStatus.ProgMax = lngProgressCount Then
                msgIcon = vbInformation
            Else
                msgIcon = vbExclamation
            End If
            
            'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
            If fmrStatus.EvalStatus(2) = True Then
                    sMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
                    fmrStatus.StatusMessage sMsg
                    DoEvents
                    strErrMsg = sMsg
                    GoTo Error_Encountered
            End If
            
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            
            ' Clear TEMP LOAD Table
            AdoExeTxt "usp_LETTER_Get_Info_tmp_clear", "v_CODE_Database"
        End If 'end if to ensure the items are marked as W to print
    
    Next varItem
    
             'so we have printed all the letters and have had no errors.
            'this Function Call is client specific for after the letters are generated.

        If Not UpdateAfterLetterGenerated(cn, UpdateLettersArray(), strErrMsg) Then
            GoTo Error_Encountered
        End If
        
        'temp cause error
        'GoTo Error_encountered:
    
    cboViewType.SetFocus
    'cboViewType.ListIndex = 1
    cboViewType = cboViewType.ItemData(1) 'JS change 20130305
    Me.txtFromDate = Format(Now, "mm/dd/yyyy")
    
    cn.CommitTrans
    cmdRefresh_Click 'can't run with open trans
    PrintSelectedItems = True
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        Call ErrorCallStack_Add(clBatchId, strErrMsg, strProcName, "Instance ID: " & strInstanceID, , , strInstanceID, objLetterInfo.LetterType)
    
        'MsgBox strErrMsg, vbCritical
    Else
        'MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
        Call ErrorCallStack_Add(clBatchId, Err.Description, strProcName, "Instance ID: " & strInstanceID, , , strInstanceID, objLetterInfo.LetterType)

    End If
    PrintSelectedItems = False
    
    'should track all documents created and delete the ones when we error'd out or cancelled generation.
     'now empty out the array
    On Error Resume Next
    For i = 0 To UBound(PrintFileArray)
        If PrintFileArray(i) <> "" Then
        Kill PrintFileArray(i)
        'only deletes folder if it is empty
            If Len(Dir(left(PrintFileArray(i), InStrRev(PrintFileArray(i), "\")) & "*.*")) = 0 Then
                DeleteFolder (left(PrintFileArray(i), InStrRev(PrintFileArray(i), "\") - 1))
            End If
        End If
        
    Next i
    
Cleanup:
    
    ' Release references.
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    Set cmd = Nothing
    Set cmdGetLetter = Nothing
    Set rsLetterConfig = Nothing
    Set rsProvList = Nothing

    objWordApp.Quit (0) 'wdDoNotSaveChanges
    Set objWordApp = Nothing
    Set MyAdo = Nothing

End Function

Function UpdateAfterLetterGenerated(cn As Variant, UpdateLettersArray() As String, ByRef strErrMsg As String) As Boolean
    'CLIENT SPECIFIC FUNCTION TGH 10/28/08 per Jeremy's info
    On Error GoTo Error_Encountered
    UpdateAfterLetterGenerated = True
        
    Dim RetCd As Integer
    Dim ErrMsg As String
    Dim cmd As ADODB.Command
        
    Dim strSQLcmd As String
    Dim intI As Integer
            
    'Zero based array
    For intI = 0 To UBound(UpdateLettersArray)
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = cn
        cmd.commandType = adCmdStoredProc
        RetCd = 1   'set return to false
        cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
        cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
        cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adVarChar, adParamOutput, 255, "")
                
        cmd.Parameters("InstanceID").Value = UpdateLettersArray(intI)
        'USER ENTRY NEEDED '******
        'usp_LETTER_UpdateAfterGenerated needs to be populated and called below...
Stop
        strSQLcmd = "usp_LETTER_AuditClaims_Update"
        cmd.CommandText = strSQLcmd
        cmd.commandType = adCmdStoredProc
        cmd.Execute
                
        strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
        RetCd = cmd.Parameters("Return").Value
            
        If RetCd <> 0 Then
            GoTo Error_Encountered
        End If
        Set cmd = Nothing
    Next intI

    UpdateAfterLetterGenerated = True
Exit_Sub:
    
    Exit Function

Error_Encountered:
    UpdateAfterLetterGenerated = False
    cn.RollbackTrans
    strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
    Resume Exit_Sub
End Function


Private Sub cmdRefresh_Click()
    RefreshMain
End Sub

Private Sub cmdReprint_Click()
On Error GoTo Block_Err
Dim strProcName As String

Dim strLetterDate As String
Dim strSQLcmd As String
Dim strInstanceID As String
Dim strNewInstanceID As String
Dim strStatus As String
Dim strErrMsg As String
Dim iRtnCd As Integer
Dim iAnswer As Integer
Dim i As Integer
Dim varItem
Dim Reprinted As Boolean: Reprinted = False
Dim curLtr As String
Dim LastLtr As String
Dim ProvNum As String


' setup progress screen
Dim sMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim fmrStatus As Form_ScrStatus
Dim iCnt As Integer: iCnt = 1 'count through loop
Dim oRs As ADODB.RecordSet
Dim cn As ADODB.Connection
Dim cmd As ADODB.Command
Dim oLetter As clsLetterInstance
Dim dtLetterReqDt As Date
Dim sNewInstanceFilter As String

    strProcName = ClassName & ".cmdRePrint"
    
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There are no items selected", vbOKOnly, "No Items Selected"
        GoTo Block_Exit
    End If
    
    Set cdctSelectedLetters = GetSelectedItems()
        
    ' open connection
    Set cn = New ADODB.Connection
    cn.ConnectionString = CodeConnString
    cn.Open
    cn.CursorLocation = adUseClient
    cn.BeginTrans
        

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.commandType = adCmdStoredProc

    cmd.CommandText = "usp_LETTER_Automation_Reprint_Letter"

    
    Set fmrStatus = New Form_ScrStatus

    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        .ProgMax = lstQueue.ItemsSelected.Count
        .TimerInterval = 50
        .show
    End With

        'example of removing hard coding from column call...
        'Get the autoID of the selected row this example is for a double click on an item
        '    lngAutoID = Me.lstQueue.column(Me.lstQueue.Recordset.Fields("AutoID").OrdinalPosition, lstQueue.ListIndex + 1)

    Set oRs = Me.lstQueue.RecordSet
    Set cdctSelectedLetters = GetSelectedItems(True)
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)
    
    For Each oLetter In cdctSelectedLetters.Letters
    
        strInstanceID = oLetter.InstanceId
        ProvNum = oLetter.cnlyProvID
        curLtr = oLetter.LetterType
        strStatus = oLetter.LetterQueueStatus
        
        
            'For P status letters only
        If UCase(strStatus) = "P" Then  ' Or UCase(strStatus) = "G" Or UCase(strStatus) = "Q" Then
            'if this is the first time set current letter and last letter = and prompt for date.  USER Request 4-17-08 TGH
            If Reprinted = False Then
                LastLtr = curLtr
                cmd.Parameters("@pLtrDate").Value = InputBox("What Date would you like on the reprint for letter " & curLtr, "Reprint Date", Format(Now(), "mm/dd/yyyy"))
            Else
                'else this isn't the first time in this loop.  check the last letter and see if it differs if so re-ask for the date for the reprint
                If curLtr <> LastLtr Then
                    LastLtr = curLtr
                    cmd.Parameters("@pLtrDate").Value = InputBox("What Date would you like on the reprint for letter " & curLtr, "Reprint Date", Format(Now(), "mm/dd/yyyy"))
                End If
                
            End If
            
    cmd.Parameters.Refresh
    cmd.Parameters("@pInstanceId") = ""
    cmd.Parameters("@pAuditor") = Identity.UserName
    cmd.Parameters("@pLtrDate") = ""
            
            'flag to let us know we have some that will be reprinted (Marked as "P")
            Reprinted = True
    
            cmd.Parameters("@pInstanceID").Value = strInstanceID
            cmd.Execute
            
            iRtnCd = cmd.Parameters("Return").Value
            strErrMsg = Trim(cmd.Parameters("@pErrMsg").Value)
            
            If iRtnCd <> 0 Or strErrMsg <> "" Then
                GoTo Block_Err
            End If
'            sNewInstanceFilter = sNewInstanceFilter & "'" & Trim(cmd.Parameters("NewInstanceID").Value) & "',"
        End If
        
        
        
        ' display progress
        If UCase(strStatus) <> "P" Then
            sMsg = "Skipping over non 'P' Records"
        Else
            sMsg = "Queuing Record " & iCnt & " / " & fmrStatus.ProgMax & vbCrLf & _
                    "Provider = " & ProvNum & vbCrLf & _
                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & curLtr
        End If
        fmrStatus.ProgVal = iCnt
        fmrStatus.StatusMessage sMsg
        iCnt = iCnt + 1
        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If
 
        DoEvents
        DoEvents
        DoEvents
        

    Next 'varItem
Stop
    If Right(sNewInstanceFilter, 1) = "," Then
        sNewInstanceFilter = left(sNewInstanceFilter, Len(sNewInstanceFilter) - 1)
    End If



    'when done all loops and no errors commit the transaction
    cn.CommitTrans
    
    If Not Reprinted Then
        iAnswer = MsgBox("Please note that only records with status ""P""  can be reprinted", vbInformation)
    Else
        '' 20140401: KD:
        '' This was confusing me:
        '' * I didn't know they'd get a new instance id
        '' * the instance id's that I wanted to reprint were all still there with a P status
        ''    so it looked like it wasn't working
        If sNewInstanceFilter <> "" Then
            lstQueue.RowSource = "SELECT wq.* FROM LETTER_Work_Queue wq INNER JOIN LETTER_Reprint_Queue rq " & _
                                 " ON wq.InstanceID = rq.InstanceID " & _
                                 " WHERE wq.Status = ""R"" " & _
                                 " AND wq.InstanceId IN (" & sNewInstanceFilter & ")"
        Else
            lstQueue.RowSource = "SELECT wq.* FROM LETTER_Work_Queue wq INNER JOIN LETTER_Reprint_Queue rq " & _
                                 " ON wq.InstanceID = rq.InstanceID " & _
                                 " WHERE wq.Status = ""R"" "
        
        End If
            
                             'and rq.Auditor = """ & Person.UserName & """"
        lstQueue.Requery
        
        If Me.lstQueue.ListCount > 0 Then
            For i = 1 To Me.lstQueue.ListCount
                Me.lstQueue.Selected(i) = True
            Next i
                         'User Entry Needed '*************
                         'Call CmdPrintSelectedItems_click if you want the reprint to automatically generate a letter. I recommend this
                         'so the user doesn't reprint a letter and it gets added to queue then they delete it or forget to reprint.
                         ' also note reprint letters are saved with -Reprint- in the filename.
                '    Call cmdPrintSelectedItems_Click
                    ' PrintSelectedItems
        End If
        
    End If
    
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
    
    
Block_Exit:
    Set cmd = Nothing
    cn.Close
    Set cn = Nothing
    Exit Sub
    
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg
    Else
        ReportError Err, strProcName
    End If
    cn.RollbackTrans
    GoTo Block_Exit
End Sub

Private Sub cmdSelectEntireQueue_Click()
Dim idx As Integer

    For idx = 1 To Me.lstQueue.ListCount
        Me.lstQueue.Selected(idx) = True
    Next idx
    ' could add in detail view but left off b/c assuming analyst are just selecting all to print in batch processes - TH 10-8-07
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub

Private Sub cmdStartDate_Click()
    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtFromDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtFromDate = mReturnDate
    
    txtFromDate_Exit False

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub

Public Sub RefreshMain()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim sCmboSql As String

    strProcName = ClassName & ".RefreshMain"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_PrintQueue_ManualOverrides"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pManualOnly") = Nz(Me.fraManualOnly, 0)
        .Parameters("@pProcessFromDt") = Format(Me.txtFromDate, "m/d/yyyy")
        .Parameters("@pProcessThruDt") = Format(Me.txtThroughDate, "m/d/yyyy")
        .Parameters("@pProcessType") = Me.cboViewType
        .Parameters("@pLetterType") = Me.cboLetterType
'        .Parameters("@pAuditor") = Me.cboAuditor
        
        If Me.tglAdvancedFilter = True Then
            .Parameters("@pExtraFilter") = msAdvancedFilter
        Else
            .Parameters("@pExtraFilter") = ""
        End If
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Problem getting the data: " & .Parameters("@pErrMsg").Value
            GoTo Block_Exit
        Else
            Set Me.lstQueue.RecordSet = oRs
            Me.lstQueue.ColumnCount = oRs.Fields.Count
        End If
    End With
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)

    Me.cboLetterType.RowSource = "SELECT DISTINCT LetterType FROM Letter_Type WHERE AccountId = " & CStr(gintAccountID) & " AND IsLetterActive = 1 AND For_DS_Only = 0 "
            
    sCmboSql = AccessSqlToSqlServer("SELECT DISTINCT BatchId, QueueRunId FROM v_Letter_Print_Queue ORDER BY QueueRunId DESC ")
    
    Call RefreshComboBoxADO(sCmboSql, Me.cmbBatches, CStr(Me.MostRecentBatchId), , "v_Code_Database")

Block_Exit:
    'Auto sizing
    Set ColResize1 = New clsAutoSizeColumns
    ColResize1.SetControl Me.lstQueue
    'don't resize if lstclaims is null
    If Me.lstQueue.ListCount > 1 Then
        ColResize1.AutoSize
    End If
    Set oAdo = Nothing
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub RefreshMain_LEGACY()
On Error GoTo Block_Err
Dim strProcName As String
Dim dtThrouDt As Date
Dim strSelectedAuditor As String
Dim strSelectedLetterType As String
Dim sAdoSql As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSelect As String
Dim sFrom As String
Dim sWhere As String
Dim sOrder As String
Dim sCmboSql As String

    strProcName = ClassName & ".RefreshMain"

    dtThrouDt = DateAdd("d", 1, txtThroughDate.Value)
    
    sSelect = "SELECT Q.*, SD.BatchId " ', D.PageCount "
    
    sFrom = " FROM LETTER_Work_Queue Q LEFT JOIN LETTER_Static_Details D ON Q.InstanceId = D.InstanceID "
    
    sFrom = " FROM LETTER_Work_Queue Q LEFT JOIN ( SELECT D.InstanceId, Max(LetterBatchId) AS BatchId FROM LETTER_Static_Details D GROUP BY D.InstanceID ) SD " & _
            " ON Q.InstanceId = SD.InstanceID "
    
    ''' Add in the left join for 2d barcodes
'    sFrom = sFrom & " LEFT JOIN LETTER_Barcode_Service_Details BS WITH (NOLOCK) ON Q.InstanceId = BS.InstanceId  "
    
    sWhere = " WHERE ProcessedDt >= #" & Nz(txtFromDate.Value, "01/01/1900") & "# and ProcessedDt < #" & _
            Format(dtThrouDt, "mm-dd-yyyy") & "# "

    Select Case cboViewType
    
        Case "View Un-Processed Letters"
            sWhere = " WHERE Status In ('W','R') "
            
            Me.cmbStatus.RowSource = "W;Ready to be Generated;R;Re-Generate;G;Generated but not printed"
            
        Case "View Processed Letters"
           sWhere = sWhere & " AND Status IN ('P','G') "
            Me.cmbStatus.RowSource = "R;Re-Generate;G;Generated but not printed"
                            
        Case "View Errors"
            sWhere = sWhere & " AND Status = 'E' "
            
        Case Else
            '* Keep Original String
    End Select
    
    ''strSelectedAuditor = Nz(cboAuditor.Value, "")   ' 20121010 KD: fixed this..
    
    If strSelectedAuditor <> "View All" And strSelectedAuditor <> "" Then
        sWhere = sWhere & " AND Auditor = '" & strSelectedAuditor & "' "
    End If
    
    strSelectedLetterType = Nz(cboLetterType.Value, "")   ' 20121010 KD: fixed this..
    
    If strSelectedLetterType <> "View All" And strSelectedLetterType <> "" Then
        sWhere = sWhere & " AND LetterType = '" & strSelectedLetterType & "' "
    End If
    
    If Me.tglAdvancedFilter = True Then
        sWhere = sWhere & " AND ( Q." & msAdvancedFilter & " ) "
    End If
    
    ' Now add in the left join for the barcodes
    
    sOrder = " ORDER BY LetterType, CnlyProvId, LetterReqDt"
    
    sAdoSql = AccessSqlToSqlServer(sSelect & sFrom & sWhere & sOrder)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sAdoSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Me.lstQueue.RowSource = sSelect & sFrom & sWhere & sOrder
        Else
            Set Me.lstQueue.RecordSet = oRs
        End If
    End With
    
    ' Get the field positions for later
    Set cdctQueueColumns = GetADOFieldOrdinalPosition(oRs)


'    Me.cboAuditor.RowSource = "Select UserID, OrderValue from (SELECT TOP 1 'View All' AS UserID, 1 AS OrderValue FROM LETTER_Work_Queue) " & _
                                " UNION (Select Auditor as UserID, 2 AS OrderValue " & sFrom & sWhere & " ) order by OrderValue, UserID; "

    
    Me.cboLetterType.RowSource = "Select LetterType, OrderValue from (SELECT TOP 1 'View All' AS LetterType, 1 AS OrderValue FROM LETTER_Work_Queue  " & _
                                " UNION Select LetterType as UserID, 2 AS OrderValue " & sFrom & sWhere & " ) As A order by OrderValue, LetterType; "


    ' Ok, so for the batches we only want to load the combo box with the batches that are actually in the list so we need to structure our query a little differently
    sCmboSql = AccessSqlToSqlServer("SELECT DISTINCT SD.LetterBatchId, SD.UserId FROM LETTER_Work_Queue Q LEFT JOIN (SELECT D.InstanceId, MAx(LetterBatchId) AS LetterBatchId, D.Auditor as UserId FROM LETTER_Static_Details D GROUP BY D.InstanceId, D.Auditor ) SD ON Q.InstanceID = SD.InstanceID " & _
        sWhere & " AND SD.InstanceId IS NOT NULL " & _
            " ORDER BY LetterBatchID DESC")
            
    Call RefreshComboBoxADO(sCmboSql, Me.cmbBatches, CStr(Me.MostRecentBatchId))

Block_Exit:
    'Auto sizing
    Set ColResize1 = New clsAutoSizeColumns
    ColResize1.SetControl Me.lstQueue
    'don't resize if lstclaims is null
    If Me.lstQueue.ListCount > 1 Then
        ColResize1.AutoSize
    End If
    Set oAdo = Nothing
    
    Exit Sub

Block_Err:
'    MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
    Else
        lblAppTitle.visible = True
    End If
    Me.lblNumSelected.Caption = " "
    
    Me.tglAdvancedFilter = False
    Me.tglAdvancedFilter.Caption = "Add Filter"
    
    
    Me.fraManualOnly = 1
    
    ' KD Comeback: Should really get the most recent batch id
    ' but for now we'll just zero it out which will hide the label
    MostRecentBatchId = 0
    
    '' Start off on Process Type: 'View Processed Letters'
'    Me.cboViewType = "View Processed Letters"
    Me.cboViewType = "View Un-processed Letters"
    Call cboViewType_AfterUpdate
    
    RefreshMain
    
End Sub

Private Sub lblRecentBatchId_DblClick(Cancel As Integer)
    DoCmd.OpenReport "rpt_LETTER_Static_Details", acViewPreview, , "[LetterBatchId] = " & CStr(Nz(Me.lblRecentBatchId, 0))
End Sub

Private Sub lstQueue_AfterUpdate()
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub

Private Sub lstQueue_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        cmdDeleteQueue_Click
    Else
        'lstQueue_Click
    End If
    
End Sub

Private Sub lstQueue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.NumSelected = Me.lstQueue.ItemsSelected.Count
End Sub

Private Sub tglAdvancedFilter_Click()

    If Me.tglAdvancedFilter.Value = True Then
    
        Set frmFilter = New Form_frm_GENERAL_Filter
        
        With frmFilter
            .CalledBy = Me.Name
            .FieldsTable = "LETTER_Work_Queue"
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

Private Sub txtFromDate_Exit(Cancel As Integer)
    If Not (IsDate(txtFromDate.Value) Or IsNull(txtFromDate.Value)) Then
        MsgBox "Please enter a valid from date"
        txtFromDate.SetFocus
    Else
'        Me.txtThroughDate = Me.txtFromDate
        RefreshMain
    End If
End Sub

Private Sub txtThroughDate_Exit(Cancel As Integer)
    If Not IsDate(txtThroughDate.Value) Then
        MsgBox "Please enter a valid through date"
        txtThroughDate.SetFocus
    Else
        If Me.txtThroughDate < Me.txtFromDate Then
            Me.txtFromDate = Me.txtThroughDate
        End If
        RefreshMain
    End If
End Sub

Private Function JoinPDFs(TargetPDF As String, FileToAppend As String) As Boolean 'This one works!
    Dim Project_Folder As String
    Dim fs As Scripting.FileSystemObject 'FileSystemObject - late bind this
    Dim AcroApp As Object
    Dim PDF_TargetFile As Object
    Dim PDF_DataSheet As Object
    Dim RowNr As Integer
    Dim iPathLen As Integer
    Dim j As Integer
    Dim i As Integer
    Dim strChkPath As String
    
    Project_Folder = left(TargetPDF, InStrRev(TargetPDF, "\") - 1)
    Set fs = New Scripting.FileSystemObject
    
    '--------------
    'test if folder exists if it does skip this logic...
    If Not fs.FolderExists(Project_Folder) Then
            'TGH taken from Damon - Added condition to handle UNC paths + check if folders exist
            iPathLen = Len(Project_Folder)
            If InStr(1, Project_Folder, "\\") > 0 Then  'in the UNC case
                'Start past the "\\cca-audit\"
                j = 12
                Do
                    i = InStr(j + 1, Project_Folder, "\")
                    If i > 0 Then
                        strChkPath = left(Project_Folder, i)
                        If Not fs.FolderExists(strChkPath) Then
                            fs.CreateFolder strChkPath
                        End If
                        j = i
                    Else
                        j = iPathLen
                    End If
                Loop Until j = iPathLen
             Else
                Do
                    i = InStr(j + 1, Project_Folder, "\")
                    If i > 0 Then
                        strChkPath = left(Project_Folder, i)
                        If Not fs.FolderExists(strChkPath) Then
                            fs.CreateFolder strChkPath
                        End If
                        j = i
                    Else
                        j = iPathLen
                    End If
                Loop Until j = iPathLen
            End If
    End If 'End of loop for when the destination folder does not exist
    '-------------
    'Now we know the folders all exist...
    'Need to create a preview PDF to append all the new PDF's into.
     
    Set AcroApp = CreateObject("AcroExch.App") 'works
    AcroApp.Hide
    Set PDF_TargetFile = CreateObject("AcroExch.PDDoc") 'This is the header file
    Set PDF_DataSheet = CreateObject("AcroExch.PDDoc") 'This will be each datasheet in turn
     
     
    If Not fs.FileExists(TargetPDF) Then
        PDF_TargetFile.Create
        PDF_TargetFile.Save 1, TargetPDF
    End If

     'Open the already created header file
  
    PDF_TargetFile.Open TargetPDF
    PDF_DataSheet.Open FileToAppend 'Project_Folder & "Test2.pdf"
     
    'Open the source document that will be added to the destination
  If PDF_TargetFile.InsertPages(PDF_TargetFile.GetNumPages - 1, PDF_DataSheet, 0, PDF_DataSheet.GetNumPages, 0) Then
    JoinPDFs = True
  Else
    JoinPDFs = False
    MsgBox ("Failure! Joining PDF's")
  End If
  
  PDF_TargetFile.Save 1, TargetPDF
  
  PDF_DataSheet.Close
  PDF_TargetFile.Close
  AcroApp.Exit 'works
    Set fs = Nothing
End Function

Private Sub CmdViewLetters_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim iTotalSelected As Integer   ' if they have more than 36,xxx selected then there SHOULD be an error!
Dim sErrMsg As String
Dim oStatusFrm As Form_ScrStatus
Dim oMainFrm As Form_frm_LETTER_Main

        ' 20130815 KD: This has changed.
        ' First, we need to determine if we are looking at generated letters or un-generated letters
        ' if generated letters there's no need to redo the MS Word mail merge
        ' if not generated, then we need to do the MS Word mail merge and add a watermark
        
    strProcName = ClassName & ".cmdViewLetters_Click"


        ' View Letter Button is written to open a new explorer window contaning the combined letters selected.
        ' You can only view generated letters.  NOTE if you leave the folder open you may have to refresh to see the file
    Set oStatusFrm = New Form_ScrStatus


    'make sure we have some items selected.
    '' now, (here) we just need to call the right procedure
    If Me.FilteredByGenerated = True Or Me.FilteredByErrors = True Then
        ' This is 'PRINT' (not Pre-View)
        ' which is really a 'Combine' and "send" to the mail room
Stop
        Set cdctSelectedLetters = GetSelectedItems(True, "GE", , sErrMsg)
        If cdctSelectedLetters Is Nothing Then
            iTotalSelected = 0
        Else
            iTotalSelected = cdctSelectedLetters.Count
        End If
        If iTotalSelected = 0 Then
            LogMessage strProcName, "WARNING", "No items selected have a 'G' status!", sErrMsg, True
            GoTo Block_Exit
        End If
        
        DoCmd.Hourglass True
        
        If PrintGeneratedLetters(oStatusFrm) = False Then
            LogMessage strProcName, "ERROR", "Print Ungenerated Letters returned false for some reason, please check the logs for details!"
            GoTo Block_Exit
        End If
        
        ' Now, lets just switch to the Label form and try to select the batch we just did:

        If IsSubForm(Me) Then
            Set oMainFrm = Me.Parent.Form
            If Nz(Me.cmbBatches, 0) <> 0 Then
                oMainFrm.BatchIdToSelect = CInt("0" & Nz(Me.cmbBatches, 0))
            End If
            oMainFrm.LoadNextTab = 3
            Set oMainFrm = Nothing
        End If
        
        
    ElseIf Me.FilteredByUngenerated = True Then
        ' This is the Preview section
        Set cdctSelectedLetters = GetSelectedItems(True, , "G", sErrMsg)
        If cdctSelectedLetters Is Nothing Then
            iTotalSelected = 0
        Else
            iTotalSelected = cdctSelectedLetters.Count
        End If
        If iTotalSelected = 0 Then
            LogMessage strProcName, "WARNING", "No items are selected!", sErrMsg, True
            GoTo Block_Exit
        End If

        DoCmd.Hourglass True

        If PreviewUngeneratedLetters(oStatusFrm) = False Then
            LogMessage strProcName, "ERROR", "Preview Ungenerated Letters returned false for some reason, please check the logs for details!"
            GoTo Block_Exit
        End If
    End If
    
    
Block_Exit:
    DoCmd.Hourglass False
    Call NotifyUserOfErrors
    Set oStatusFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Function PrintGeneratedLetters(oStatusForm As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sWhere As String
Dim oLetter As clsLetterInstance
Dim sInstanceIds As String
Dim sTempFldr As String
Dim SFileName As String
Dim sOrigFilePath As String
Dim sCurLetterType As String
Dim dctWordDocs As Scripting.Dictionary
Dim iErrCnt As Integer
Dim strErrMsg As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim oWord As Word.Application
Dim oMergedDoc As Word.Document
Dim iPageCount As Integer
Dim lAddressRow As Long
Dim sLastInstanceId As String
Dim sCurInstanceId As String
Dim bFirst As Boolean
Dim oCmd As ADODB.Command
Dim strProgressMsg As String
Dim lCnt As Long
Dim lngProgressCount As Long
Dim sNewStatus As String
Dim lTtlPages As Long
Dim lCombinedDocCount As Long
Dim lPagesForCombinedDocs As Long
Dim bForward As Boolean
Dim dtStart As Date
Dim oCn As ADODB.Connection

    bForward = False

    strProcName = ClassName & ".PrintGeneratedLetters"
Stop
    '' 20130815 KD: All we have to do here is:
    '' - get a RS with the ones to select including the document paths (Sorted by pagecount ASC)
    
    lPagesForCombinedDocs = GetConfigPagesForCombined()
    
        ' We already have our selected letters:
    For Each oLetter In cdctSelectedLetters.Letters
        sInstanceIds = sInstanceIds & "'" & oLetter.InstanceId & "',"
    Next
    sInstanceIds = left(sInstanceIds, Len(sInstanceIds) - 1) ' remove final comma
    sWhere = "WHERE X.InstanceID IN (" & sInstanceIds & ") "

    
    sSql = "SELECT DISTINCT X.InstanceID, X.LetterType, LetterName As DocumentPath, PageCount, ProvName, ContactTitle " & _
            ", ContactName, ContactType, Addr01, Addr02, Addr03, City, State, Zip " & _
            " FROM Letter_Xref X LEFT JOIN LETTER_Static_Details SD ON X.InstanceId = SD.InstanceID " & sWhere & _
            " ORDER BY X.LetterType, PageCount  "  ' the order by is very important so we put the large ones at the end of the stack so
                                        ' the envelope machine goes through the batch until it hits the first letter with 8 or more
                                        ' pages... we are going by letter type first, because we are going to group all letter types
                                        ' together. Typically users already do this, but there's no reason they HAVE to so we're going
                                        ' to make sure we take care of it in code..
            
    
    If bForward = True Then
        sSql = sSql & " ASC"
    Else
        sSql = sSql & " DESC"
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "WARNING", "No data received for selected InstanceIDs!", sInstanceIds, True
            GoTo Block_Exit
        End If
    End With
    
    '' - copy the documents to a temp user folder ' cs_USER_TEMPLATE_PATH_ROOT
    sTempFldr = GetUserTempDirectory()
    If sTempFldr = "" Then
        LogMessage strProcName, "ERROR", "There was an error creating a work folder!", , True
        GoTo Block_Exit
    End If
    
    Set dctWordDocs = New Scripting.Dictionary
    
    ' since there's a bit of a file system lag,
    ' I decided to copy all of the files to the directory first
    lngProgressCount = oRs.recordCount * 2 ' (times 2 because copying is step 1, then combining is step2)
    
    'BEGIN THE PROGRESS FORM, MAKE THE MAX PROGRESS DOUBLE THE ITEMS SELECTED B/C WE HAVE TO TOUCH EACH REC TWICE WHEN GENERATING.
    With oStatusForm
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgMax = lngProgressCount
        .ProgAllMax = lngProgressCount
        .TimerInterval = 50
        .show
    End With
    
    While Not oRs.EOF
        sOrigFilePath = oRs("DocumentPath").Value
        SFileName = GetFileName(sOrigFilePath)
        lCnt = lCnt + 1

        ' display progress
        strProgressMsg = "Copying generated letters to temp folder " & lCnt & " / " & CStr(lngProgressCount / 2)
        oStatusForm.ProgVal = lCnt
        oStatusForm.StatusMessage strProgressMsg

            ' If it's already there, then remove it..
        If FileExists(sTempFldr & SFileName) = True Then
            ' what do we do now? I guess it's just a copy so we should probably delete, then copy it again
            ' assuming that the source is still there
            If FileExists(sOrigFilePath) = False Then
                LogMessage strProcName, "ERROR", "The source file to load for InstanceID: " & oRs("InstanceId").Value & " is missing", sOrigFilePath
                Call ErrorCallStack_Add(0, "Source file to load for Instance id: " & oRs("InstanceId").Value & " is missing", strProcName, sOrigFilePath, , , oRs("InstanceID").Value, oRs("LetterType").Value)
                iErrCnt = iErrCnt + 1
                GoTo NextRow
            Else
                If DeleteFile(sTempFldr & SFileName, False) = False Then
                    LogMessage strProcName, "WARNING", "The file being copied already exists and appears to be locked open", sTempFldr & SFileName
                    Call ErrorCallStack_Add(0, "The file being copied already exists and appears to be locked open.", strProcName, sTempFldr & SFileName, , , oRs("InstanceID").Value, oRs("LetterType").Value)
                    GoTo NextRow
                End If
            End If
        Else
'        Stop
        End If
        
        If CopyFile(sOrigFilePath, sTempFldr & SFileName, False, strErrMsg) = False Then
            LogMessage strProcName, "ERROR", "There was a problem copying a file to the temp folder: " & strErrMsg, sOrigFilePath & " to " & sTempFldr & SFileName
            Call ErrorCallStack_Add(0, "There was a problem copying a file to the temp folder: " & strErrMsg, strProcName, sOrigFilePath & " to " & sTempFldr & SFileName, , , oRs("InstanceID").Value, oRs("LetterType").Value)
            iErrCnt = iErrCnt + 1
        End If
        
NextRow:
        oRs.MoveNext
    Wend

    ' reset because now we are going to go through and make 1 MS Word document
    ' and clean up the copies at the end
    oRs.MoveFirst
    
    Set oWord = New Word.Application
    oWord.visible = False

        ' for letters with page count of 10 or more, they go in a different envelop
        ' so we will need labels.  for now, I'm being lazy and I'm just going to dump this to a Spreadsheet
    lAddressRow = 2 ' excel starts with 1, but we're going to have a header row

    bFirst = True
    lCombinedDocCount = 1

    dtStart = Now
    LogMessage strProcName, "DEBUG", "Started printing", Format(dtStart, "m/d/yyyy hh:nn:ss AMPM")
    
        '' We are going to put this in a transaction too so if the user cancels the status isn't set to P for ones that they didn't
        '' actually get the chance to print.
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = CodeConnString
    oCn.Open
    oCn.BeginTrans
    
    '' - open the first one in MS Word,
    While Not oRs.EOF
        sCurLetterType = oRs("LetterType").Value
        sOrigFilePath = oRs("DocumentPath").Value
        SFileName = GetFileName(sOrigFilePath)
        iPageCount = Nz(oRs("PageCount").Value, 0)
        sCurInstanceId = Nz(oRs("InstanceID").Value, "")
        sNewStatus = "P"    ' optimistic thinking...
        
        lCnt = lCnt + 1

        ' display progress
        strProgressMsg = "Merging letters into 1 document for printing. " & CStr(lCnt - (lngProgressCount / 2)) & " / " & CStr(lngProgressCount / 2)
        oStatusForm.ProgVal = lCnt
        oStatusForm.StatusMessage strProgressMsg
 
        DoEvents    ' This is for the status form...
        DoEvents
        DoEvents
        
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If oStatusForm.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected!"
            oStatusForm.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            oCn.RollbackTrans
            GoTo Block_Exit
        End If
        

        If lTtlPages >= lPagesForCombinedDocs Then
            ' we want to close and save this puppy.
            If Not oMergedDoc Is Nothing Then
                oMergedDoc.Save
                oMergedDoc.Close
                Set oMergedDoc = Nothing
            End If
            ' we want to start a new document..
            lTtlPages = 0
            bFirst = True
            lCombinedDocCount = lCombinedDocCount + 1
'        Else
''Stop
        End If
        
        
        If bFirst = True Then
            ' new word document
            ' well, do we need to save the last one?
            If Not oMergedDoc Is Nothing Then
                    Stop    ' what the?
            End If
            Set oMergedDoc = oWord.Documents.Open(sTempFldr & SFileName, False, True, False, , , , , , , , False)
            ' we are going to rename this one so we know it's a merged document
            If FileExists(sTempFldr & Format(lCombinedDocCount, "0##") & "_" & sCurLetterType & "_MergedDoc.doc") = True Then
                Call DeleteFile(sTempFldr & Format(lCombinedDocCount, "0##") & "_" & sCurLetterType & "_MergedDoc.doc", False)
            End If
                ' Probably don't need to do this but it's only on the first one so not hurting much
            If UnlinkWordFields(oWord, oMergedDoc, oRs("LetterType").Value) = False Then
                LogMessage strProcName, "ERROR", "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!"
                Call ErrorCallStack_Add(0, "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!", strProcName, , , , oRs("InstanceID").Value, oRs("LetterType").Value)
            End If
            
            '2014:04:29:JS Addded this here because InsertWordDocAtStartOfCurrentDoc only does it for the inserted documents now.
            With oMergedDoc
                .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                .Repaginate
            End With
            
            oMergedDoc.SaveAs2 (sTempFldr & Format(lCombinedDocCount, "0##") & "_" & sCurLetterType & "_MergedDoc.doc")
            bFirst = False
            
        Else
            ' not a new letter type so we just append this to our current mergedDoc
            ' the next row should always be a different instance but lets check to make sure..
            If sLastInstanceId <> sCurInstanceId Then
                    '' - use Word to add the next document to the end of that first one.
                If bForward = True Then
                    If InsertWordDocAtEndOfCurrentDoc(oMergedDoc, sTempFldr & SFileName, lTtlPages) = False Then
    
                        Call ErrorCallStack_Add(0, "There was a problem adding a letter to the beginning of the marged document", strProcName, SFileName, , , oRs("InstanceID").Value, oRs("LetterType").Value)
                        iErrCnt = iErrCnt + 1
                        sNewStatus = "E"
                    Else
                        sNewStatus = "P"
                        oMergedDoc.Save
                    End If

                Else
                    If InsertWordDocAtStartOfCurrentDoc(oMergedDoc, sTempFldr & SFileName, lTtlPages) = False Then
    
                        Call ErrorCallStack_Add(0, "There was a problem adding a letter to the beginning of the marged document", strProcName, SFileName, , , oRs("InstanceID").Value, oRs("LetterType").Value)
                        iErrCnt = iErrCnt + 1
                        sNewStatus = "E"
                    Else
                        sNewStatus = "P"
                        If lCnt Mod 100 = 0 Then
                            oMergedDoc.Save
                        End If
                    End If

                End If
'''                LogMessage strProcName, "EFFICIENCY TESTING", "Finished InsertWordDoc," & sCurLetterType, CStr(lCnt / 2) & ", Total Pages now: " & CStr(lTtlPages)
            End If
        End If
        
        '' 20130821 KD: Ok, now we have to change the status to P (should only be G here (g = generated, not printed)

        
            ' update LETTER status
        Set oCmd = New ADODB.Command

        Set oCmd.ActiveConnection = oCn
        oCmd.commandType = adCmdStoredProc
        oCmd.CommandText = "usp_LETTER_Work_Queue_Update_Status"
        oCmd.Parameters.Refresh
        oCmd.Parameters("@InstanceID").Value = sCurInstanceId
        oCmd.Parameters("@StatusCD").Value = sNewStatus '    "P" ' for Generated And printed.. (we think
            ' - no way to tell if the user specifically printed - unless
            '   we programatically print for them..
        oCmd.Execute
        Set oCmd = Nothing
        
        sLastInstanceId = sCurInstanceId
        oRs.MoveNext

        DoEvents    ' This is for the status form...
        DoEvents
        DoEvents
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If oStatusForm.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected!" ' at " & i & " / " & fmrStatus.ProgMax
            oStatusForm.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            oCn.RollbackTrans
            GoTo Block_Exit
        End If
        
    Wend
    
    
    
    If FileExists(sTempFldr & SFileName) = False Then
        LogMessage strProcName, "ERROR", "File to append does not seem to exist where specified!", sTempFldr & SFileName
        GoTo Block_Err
    End If
    

        
    ' If we get this far then the user didn't cancel
    oCn.CommitTrans
    
    '' - clean up the copies, leaving only the combined document present
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sTempFldr)
    For Each oFile In oFldr.Files
        If oFile.Name <> oMergedDoc.Name Then
            ' we can delete it..
            oFile.Delete True
        End If
    Next
    
    LogMessage strProcName, "DEBUG", "Finished printing", ProcessTookHowLong(dtStart, Now())
 

    ' We can close the document because they'll just right click on the
    ' document in explorer and choose the print option.
    
    If Not oMergedDoc Is Nothing Then
        oMergedDoc.Activate
        oMergedDoc.Repaginate
        oMergedDoc.Save
        oMergedDoc.Close True
        Set oMergedDoc = Nothing
    End If
    
    Shell "explorer.exe " & Chr$(34) & sTempFldr & Chr$(34), vbNormalFocus
    
    PrintGeneratedLetters = True
    RefreshMain
    
Block_Exit:
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then
            
            oCn.Close
        End If
        Set oCn = Nothing
    End If
    If Not oMergedDoc Is Nothing Then
        oMergedDoc.Save
        oMergedDoc.Close True
        Set oMergedDoc = Nothing
    End If
    If Not oWord Is Nothing Then
        oWord.Quit
        Set oWord = Nothing
    End If
    
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Set oMergedDoc = Nothing
    Set dctWordDocs = Nothing
    Exit Function
Block_Err:
    
    If Err.Number <> 0 Then
        'ReportError Err, strProcName
        Call ErrorCallStack_Add(0, Err.Description, strProcName)
        LogMessage strProcName, "ERROR", Err.Description
    Else
        Call ErrorCallStack_Add(0, strErrMsg, strProcName)
    
        LogMessage strProcName, "ERROR", strErrMsg
    End If
    If Not oCn Is Nothing Then
        oCn.RollbackTrans
    End If
    GoTo Block_Exit
End Function


Private Function GetConfigPagesForCombined() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim lRet As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetConfigPagesForCombined"
    
    lRet = 500 ' default to something safe if something goes wrong we're still good
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM LETTER_Config WHERE AccountId = 1"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Encountered an error while trying to get the number of pages for the combined documents", , True

            GoTo Block_Exit
        End If
    End With
    
    lRet = oRs("NumberOfPagesForCombinedDocs").Value
    
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    If Not oAdo Is Nothing Then
        Set oAdo = Nothing
    End If
    GetConfigPagesForCombined = lRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Function InsertWordDocAtEndOfCurrentDoc(oWordDoc As Word.Document, sFileToInsert As String, Optional lTotalPagesAfterMerge As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oWordApp As Word.Application
Dim lTtlPages As Long
'Dim lCurPage As Long

'Const iMethodToTry As Integer = 1

    strProcName = ClassName & ".InsertWordDocAtEndOfCurrentDoc"
    If FileExists(sFileToInsert) = False Then
        LogMessage strProcName, "ERROR", "File to append does not seem to exist where specified!", sFileToInsert
        GoTo Block_Exit
    End If
    
    ' to insert a file we need to get a selection where we want the document to start..
    ' so, loop through the pages
'''    LogMessage strProcName, "EFFICIENCY TESTING", "Starting to insert word doc", "Method: " & CStr(iMethodToTry)
    
    
    Set oWordApp = oWordDoc.Application
''''    Select Case iMethodToTry
''''    Case 1
''''        oWordDoc.Select
''''        lTtlPages = oWordApp.selection.Information(4)   'wdNumberOfPagesInDocument)
''''
''''        oWordApp.selection.Goto 1, 2, lTtlPages
''''        oWordApp.selection.WholeStory
''''        oWordApp.selection.EndKey Unit:=wdStory
''''    Case 2
        oWordDoc.Activate
        lTtlPages = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
        
        oWordApp.selection.GoTo 1, 2, lTtlPages
'        oWordApp.selection.WholeStory
        oWordApp.selection.EndKey Unit:=wdStory

''''    End Select
    '' below may be quicker than selecting..
    '' ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)

   
    
    oWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
    
    oWordApp.selection.InsertFile (sFileToInsert)
    
    lTotalPagesAfterMerge = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
''''    LogMessage strProcName, "EFFICIENCY TESTING", "Finished inserting word doc", "Method: " & CStr(iMethodToTry) & "," & CStr(lTtlPages)


    InsertWordDocAtEndOfCurrentDoc = True
    
Block_Exit:
    Set oWordApp = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function InsertWordDocAtStartOfCurrentDoc(oWordDoc As Word.Document, sFileToInsert As String, Optional lTotalPagesAfterMerge As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oWordApp As Word.Application
Dim lTtlPages As Long
'Dim lCurPage As Long
' KD: ok, so to deal with a bug, design flaw, whatever
' we need to first open the file to be inserted, and make sure it starts and finishes with a Continuous section break.
Dim oWordInsertDoc As Word.Document
'Const iMethodToTry As Integer = 1

    strProcName = ClassName & ".InsertWordDocAtStartOfCurrentDoc"
    If FileExists(sFileToInsert) = False Then
        LogMessage strProcName, "ERROR", "File to append does not seem to exist where specified!", sFileToInsert
        GoTo Block_Exit
    End If
    
    ' to insert a file we need to get a selection where we want the document to start..
    ' so, loop through the pages
'''    LogMessage strProcName, "EFFICIENCY TESTING", "Starting to insert word doc", "Method: " & CStr(iMethodToTry)
    
    
    Set oWordApp = oWordDoc.Application
    'Set oWordInsertDoc = oWordApp.OpenAttachmentsInFullScreen(sFileToInsert)
'''    Set oWordInsertDoc = oWordApp.Documents.Open(sFileToInsert, False, False, False, , , , , , , , False)
'''
'''    With oWordApp
'''        .selection.HomeKey Unit:=6   '   wdStory
'''        .selection.InsertBreak Type:=3      ' wdSectionBreakContinuous
'''        .selection.Delete Unit:=1, Count:=1 ' 1 = wdCharacter
'''        .selection.EndKey Unit:=6   'wdStory
'''        .selection.InsertBreak Type:=3      'wdSectionBreakContinuous
'''    End With
'''    oWordInsertDoc.Save
'''    oWordInsertDoc.Close
'''    Set oWordInsertDoc = Nothing
    
''''    Select Case iMethodToTry
''''    Case 1
''''        oWordDoc.Select
''''        lTtlPages = oWordApp.selection.Information(4)   'wdNumberOfPagesInDocument)
''''
''''        oWordApp.selection.Goto 1, 2, lTtlPages
''''        oWordApp.selection.WholeStory
''''        oWordApp.selection.EndKey Unit:=wdStory
''''    Case 2
        oWordDoc.Activate
'        lTtlPages = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
        
'' Lets make sure that all of the footers (and headers) start the number at 1:

Dim objWordField As Word.Field
Dim objWordSection As Word.Section
Dim oShape As Word.Shape
Dim i As Integer

'    With oWordDoc
        'For i = 1 To .Sections.Count

'            .Sections(i).headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'            .Sections(i).headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'            .Repaginate
'            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
'            .Sections(i).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
'            .Repaginate

            '' how about shapes?

'            For Each oShape In .Sections(i).Footers(wdHeaderFooterPrimary).Shapes
'                oShape.TextFrame.TextRange.Fields.Unlink
'            Next
'
'            For Each oShape In .Sections(i).headers(wdHeaderFooterPrimary).Shapes
'                oShape.TextFrame.TextRange.Fields.Unlink
'            Next

        'Next i
        '.Fields.Unlink
'   End With
        
        oWordApp.selection.GoTo 1, 1, 1
'        oWordApp.selection.WholeStory
        oWordApp.selection.HomeKey Unit:=wdStory

''''    End Select
    '' below may be quicker than selecting..
    '' ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)

   
    
    
    oWordApp.selection.InsertFile (sFileToInsert)
    oWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
    
    
    '2014:04:29:JS: Why do this for each section every time? change to do it for the last one inserted, which would be always section 1
    With oWordDoc
        .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        '.Repaginate
        .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
    End With
    
'    lTotalPagesAfterMerge = oWordDoc.BuiltInDocumentProperties(14)  ' wdPropertyPages
''''    LogMessage strProcName, "EFFICIENCY TESTING", "Finished inserting word doc", "Method: " & CStr(iMethodToTry) & "," & CStr(lTtlPages)


    InsertWordDocAtStartOfCurrentDoc = True
    
Block_Exit:
    Set oWordApp = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function PreviewUngeneratedLetters(oStatusForm As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim lTranReturn As Long
Dim rsLetterConfig As ADODB.RecordSet
Dim bPreviewViewAllowed As Boolean
Dim fmrStatus As Form_ScrStatus
Dim lTotalPrintedRecs As Long
Dim lTotalToPrint As Long
'Dim oCn As ADODB.Connection
Dim sErrMsg As String
Dim bAtLeastOneErrored As Boolean


    strProcName = ClassName & ".PreviewUngeneratedLetters"
    '' 20130815 KD: Basically we have to re-create the generate process here
    ''  except since we aren't going to SAVE the document, we are going to add a watermark
    '' therefore, I'll probably have to redo the generate process cause i don't like
    '' duplicate code flowing around
    

    'TL add account ID logic
    Set rsLetterConfig = GetLetterConfigDetails()
    If UCase(rsLetterConfig("AllowPreview").Value) = "TRUE" Or UCase(rsLetterConfig("AllowPreview").Value) = "YES" Then
        bPreviewViewAllowed = True
    Else
        bPreviewViewAllowed = False
    End If
    If rsLetterConfig.State = adStateOpen Then rsLetterConfig.Close
    Set rsLetterConfig = Nothing


      ' We already have our selected letters:
            '''    For Each oLetter In cdctSelectedLetters.Letters
            '''        sInstanceIds = sInstanceIds & "'" & oLetter.InstanceID & "',"
            '''    Next
            '''    sInstanceIds = left(sInstanceIds, Len(sInstanceIds) - 1) ' remove final comma
            '''    sWhere = "WHERE X.InstanceID IN (" & sInstanceIds & ") "


    ' in order to get here we need to have come through the Generate button
    ' which means we should have our form scoped cdctSelectedLetters
    ' I don't like that idea though.. Let's get it here
Stop
    If cdctSelectedLetters Is Nothing Then
        Set cdctSelectedLetters = GetSelectedItems(, , "G", sErrMsg)
        If cdctSelectedLetters Is Nothing Then
            bAtLeastOneErrored = True
            lTotalToPrint = 0
        End If
    End If
    lTotalToPrint = cdctSelectedLetters.Letters.Count

    'get the count of Printed leters.  Because the auditor can select all letters in the grid they could pick a mix of unprocessed and processed.
    'we handle the 'P'  and 'R' Lletters differently - we pull the actual letter generated
    ' all others we create a temp preview file
    ' 20130821: KD: For 'R' (Regenerate) we are indeed generating a new letter - unless someone wants to see the actual letter
    '   that was already generated (but then they could just go to the claim and look at it there)

    'run QC Error check to see if no Printed recs selected and preview is not allowed:
    If Not bPreviewViewAllowed And lTotalToPrint = 0 Then
        MsgBox ("Selected items need to be printed to be viewed")
        GoTo Block_Exit
    ElseIf bPreviewViewAllowed And (lTotalToPrint = 0) Then
        MsgBox ("Please Select Letters to View")
        GoTo Block_Exit
    End If


    ' setup progress screen
    Set fmrStatus = New Form_ScrStatus

    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        If bPreviewViewAllowed Then 'TGH added 9-26-08
            .ProgMax = lTotalToPrint
        Else
            .ProgMax = lTotalToPrint ' Using  b/c we only want to view the printed letters.
        End If
        .TimerInterval = 50
        .show
    End With

    PreviewUngeneratedLetters = True
    
Block_Exit:
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function ViewLetters(TotalRecs As Integer, cn As Variant, fmrStatus As Form_ScrStatus) As Boolean
'This funciton has been updated by 'TGH 9-26-08
' I added in the option based off of the boolean PreviewViewAllowed to view non printed letters.

    On Error GoTo Error_Encountered:
    ViewLetters = True
    
    'make sure we have some items selected.
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected"
        ViewLetters = False
        Exit Function
    End If
          
'Variable declarations ------------------------------------------------------------------------
'    Dim Person As New ClsIdentity
Dim strErrMsg As String: strErrMsg = ""
Dim Usedword As Boolean
Dim db As Database
Dim rsLetterConfig As DAO.RecordSet
Dim strInstanceID As String
Dim strStatus As String
Dim strOutputFileName As String
Dim varItem As Variant
Dim bFirstLetter As Boolean
Dim iCnt As Integer
Dim strOutputPath As String
Dim strAuditor As String
Dim colLetters As Collection    'thieu 1/16/08
Dim bnewwordfile As Boolean     'thieu 1/16/08
Dim bnewpdffile As Boolean
Dim ShowPreview As Boolean
Dim curLtr As String
Dim strPreviewFileName As String
Dim sMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
'    ' ADO variables
Dim cmdGetLetter As ADODB.Command
    Set cmdGetLetter = New ADODB.Command
Dim ProgVal As Integer
Dim strSQLcmd As String
'pdf objects
Dim TargetPDF As String: TargetPDF = ""
Dim PassFail As Boolean

Dim objWordApp, _
    objMasterDoc, _
    objTemplateDoc ', objWordSelection
    
'    Dim objWordApp As Word.Application, _
'        objMasterDoc As Word.Document, _
'        objTemplateDoc As Word.Document ', objWordSelection
    
    
Dim strTemplateLoc As String
Dim CountCurLetter As Integer
Dim i As Long
Dim MyRecordset As DAO.RecordSet
Stop
Stop
Stop

    Set objWordApp = CreateObject("word.application")
'    Set objWordApp = New Word.Application
    
        ' Create an instance of objWord, and make it invisible.
    objWordApp.visible = False

    'End of declarations ------------------------------------------------------------------------

    'assign the preview from the config table
    Set db = CurrentDb
    
    'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    strOutputPath = rsLetterConfig("LetterOutputLocation").Value

    'Setup for letter viewing.  setting auditor and clearing out the preview folder
    ShowPreview = False
    Usedword = False
    strAuditor = Replace(Identity.UserName, ".", "")              ' thieu 1/16/08
    strOutputPath = strOutputPath & "\PREVIEW\" & strAuditor    'thieu 1/16/08
    DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
    CreateFolder (strOutputPath & "\")                                ' thieu 1/16/08
    Set colLetters = New Collection                             ' thieu 1/16/08
    bnewwordfile = False                                        ' thieu 1/16/08
    'TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.column(2, 1)) & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
    bnewpdffile = False
    
    'this is the first time through the view letter.
    bFirstLetter = True
    On Error GoTo Error_Encountered
        
    'set to the formstatus current progress
    iCnt = 0
    ProgVal = fmrStatus.ProgVal
    cmdGetLetter.ActiveConnection = cn
    cmdGetLetter.commandType = adCmdStoredProc
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("LetterName", adChar, adParamOutput, 255, "")
    cmdGetLetter.CommandText = "usp_LETTER_Get_Letter_Name"

    curLtr = ""
    CountCurLetter = 0
    

    Set MyRecordset = Me.lstQueue.RecordSet
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    For Each varItem In lstQueue.ItemsSelected
            'grab the instance id and status from the listqueue
            strInstanceID = Trim(Me.lstQueue.Column(MyRecordset.Fields("instanceID").OrdinalPosition, varItem))
            strStatus = Trim(Me.lstQueue.Column(MyRecordset.Fields("Status").OrdinalPosition, varItem))
            
        If UCase(strStatus) = "P" Then
            'mark that the preview window should be shone b/c there is at least one instance marked as 'print'
            ShowPreview = True
            CountCurLetter = CountCurLetter + 1
            'lookup this instanceID and grag the appropriate LetterName
            cmdGetLetter.Parameters("InstanceID").Value = strInstanceID
            cmdGetLetter.Execute ' executing "usp_LETTER_Get_Letter_Name"
            'see the filename back to see how we want to handle the file pdf/word.
            strOutputFileName = Trim(cmdGetLetter.Parameters("LetterName").Value)
                
            'check current letter vs the last P record we looked at.
            'First time through CurLtr = "" so we don't enter this statement
            If StrComp(curLtr, "", vbTextCompare) = 1 And _
                curLtr <> Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem)) Then
                    If InStr(1, strOutputFileName, ".doc") > 0 Then
                        bnewwordfile = True 'we should start a new word file for a different word letter.
                    Else
                        TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.Column(2, varItem)) & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
                        bFirstLetter = True
                        bnewwordfile = False        'thieu
                    End If
            End If
            
            curLtr = Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem))
            
            'At this point we have marked if the letter is a new or if we have changed letters
            ' TEST if this is a PDF or Word Doc.
            If InStr(1, strOutputFileName, ".PDF", vbTextCompare) = 0 Then 'we are NOT dealing with a pdf document (so word doc unless we add more)...
                'set the flag to show word documents at the end and close open word app
                Usedword = True

                'if we have a new letter close the old one and save it.
                If bnewwordfile Then                        ' thieu 1/16/08
                    objMasterDoc.spellingchecked = True 'this is needed to shut down popup for too many spelling errors
                    Sleep 1000
                    objMasterDoc.SaveAs strPreviewFileName
                    objMasterDoc.Close
                    Set objMasterDoc = Nothing
                    strPreviewFileName = strOutputPath & "\" & curLtr & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    bFirstLetter = True         'mark this so we exit this clause until we are ready for a new word file
                    bnewwordfile = False        'thieu
                End If

                If bFirstLetter = True Then
                    strOutputFileName = Trim(cmdGetLetter.Parameters("LetterName").Value)
                    'need to get template's name
                    strTemplateLoc = DLookup("TemplateLoc", "LETTER_Type", "Lettertype = '" & curLtr & "'")
                    'reset the counter for how many for this letter
                    CountCurLetter = 0
                    
                    'Now check if we are dealing with a word doc or a PDF image.
                    If InStr(1, strOutputFileName, ".doc", vbTextCompare) > 0 Then
                        'set the word documents and begin setup...
                        'Set objMasterDoc = objWordApp.Documents.Add()   'tgh 3/18/08
                        Set objMasterDoc = objWordApp.Documents.Open(strOutputFileName)
                            'objMasterDoc
                        
                        'Set objTemplateDoc = objWordApp.Documents.Open(strTemplateLoc)
                        '    objMasterDoc.PageSetup.LeftMargin = objTemplateDoc.PageSetup.LeftMargin
                        '    objMasterDoc.PageSetup.RightMargin = objTemplateDoc.PageSetup.RightMargin
                        '    objMasterDoc.PageSetup.TopMargin = objTemplateDoc.PageSetup.TopMargin
                        '    objMasterDoc.PageSetup.BottomMargin = objTemplateDoc.PageSetup.BottomMargin
                         '   objMasterDoc.PageSetup.HeaderDistance = objTemplateDoc.PageSetup.HeaderDistance
                        '    objMasterDoc.PageSetup.FooterDistance = objTemplateDoc.PageSetup.FooterDistance
                        '    objMasterDoc.spellingchecked = True
                        '    objMasterDoc.showspellingerrors = False
                        '    objMasterDoc.showgrammaticalerrors = False
                        'objTemplateDoc.Close
                        objWordApp.ActiveDocument.spellingchecked = True
                        'objWordApp.Selection.InsertFile (strOutputFileName)
                        'viewing letter name is based on time, can use instance_id for first letter if we want...
                        objWordApp.selection.EndKey Unit:=wdStory
                        strPreviewFileName = strOutputPath & "\" & curLtr & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    End If
                
                    bFirstLetter = False
                Else 'else we are just adding to a file already created.
                    'Set objTemplateDoc = objWordApp.Documents.Open(strOutputFileName)
                    'objTemplateDoc.Close
'                    objWordApp.Selection.InsertBreak (7) 'wdPageBreak

                    objWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage)
                    objWordApp.selection.InsertFile (strOutputFileName)
                    objWordApp.selection.EndKey Unit:=wdStory
                    'need to re-set these after every insert! it gets defaluted back once a new file is inserted.
                    objMasterDoc.spellingchecked = True
                    objMasterDoc.showspellingerrors = False
                    objMasterDoc.showgrammaticalerrors = False
                    
                    'If iCnt > 0 And (iCnt Mod 10) = 0 Then
                    If iCnt > fmrStatus.ProgVal And (iCnt Mod 40) = 0 Then                          'i think this re-does the tabs
                        objMasterDoc.Repaginate                                                     'thieu 1/16/08
                    End If
                
                   
                    If CountCurLetter > 100 Then 'if the Letter type is comprised of over 100 letters split it
                    'this below does not work when we hit 100 pages it keeps going. not refreshing correctly.
                    'If objMasterDoc.BuiltInDocumentProperties("Number of Pages") > 100 Then 'wdPropertyPages       'thieu 1/16/08 property # 14
                        bnewwordfile = True                                                             'thieu 1/16/08
                    End If                                                                          'thieu 1/16/08
                End If ' End of check if were are in the bFirstLetter
            'else we are dealing with a PDF
            ElseIf InStr(1, strOutputFileName, ".PDF", vbTextCompare) > 0 Then
            'we have a PDF output so we call Join PDF's.
                'if targetPDF is empty that means this is the first time we are encountering a pdf so set the target pdf here.
                If TargetPDF = "" Or CountCurLetter > 100 Then 'if the Letter type is comprised of over 100 letters split it
                 TargetPDF = strOutputPath & "\" & Trim(Me.lstQueue.Column(MyRecordset.Fields("LetterType").OrdinalPosition, varItem)) _
                                           & "-" & Format(Now, "yyyymmddhhmmss") & ".pdf"
                End If
                    
                    PassFail = JoinPDFs(TargetPDF, strOutputFileName)
                    If Err Or PassFail = False Then
                        MsgBox ("error joining pdfs")
                        Exit Function
                    End If
            Else
                MsgBox ("File needs to be in word or pdf format")
            End If
        'keep track of how many times we have looped.
        iCnt = iCnt + 1
    End If
        
        ' display progress
        sMsg = "View Record " & iCnt & " / " & TotalRecs
        fmrStatus.ProgVal = ProgVal + iCnt
        fmrStatus.StatusMessage sMsg

        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If

        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.  if  so rollback and promt with error message.
        If fmrStatus.EvalStatus(2) = True Then
                sMsg = "Viewing Generated Letters Canceled!"
                fmrStatus.StatusMessage sMsg
                DoEvents
                strErrMsg = sMsg
                GoTo Error_Encountered
        End If
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Next varItem

 If ShowPreview Then
    If Usedword Then
        objMasterDoc.spellingchecked = True
        
        With objMasterDoc
            For i = 1 To .Sections.Count
                .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            Next i
        End With
        Sleep 1000
        objMasterDoc.SaveAs strPreviewFileName
        objMasterDoc.Close
        Set objMasterDoc = Nothing
    End If
      
      'right here we need to execute the print command with view only as a parameter.
      
        Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
    Else
        MsgBox ("Items selected must be printed already to view.")
    End If
    ViewLetters = True
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    ViewLetters = False

Cleanup:

    If Not objMasterDoc Is Nothing Then '07/01/2013
        objMasterDoc.Close wdDoNotSaveChanges
    End If

    
    Set objMasterDoc = Nothing
    Set objTemplateDoc = Nothing
    'Do not close the Connection passed in, taking care of that in the calling sub.
    Set cmdGetLetter = Nothing
'    Set Person = Nothing
    objWordApp.Quit wdDoNotSaveChanges
     Set objWordApp = Nothing
    ' make word visible for user to view
    'objWordApp.Visible = True
    'objWordApp.Activate

End Function


Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mReturnDate = ReturnDate
End Sub


Private Function PreviewViewLetters(TotalRecs As Long, fmrStatus As Form_ScrStatus) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim cmdGetLetter As ADODB.Command
Dim strSQLcmd As String

' Letter configuration variables
Dim rsLetterConfig As ADODB.RecordSet
Dim strODCFile As String
Dim strBasedPath As String
    
    ' Word objects setup as variants b/c of late binding (due to the various versions we have scattered around the environment)
Dim objWordApp As Word.Application, _
    objMasterDoc As Word.Document, _
    objWordDoc As Word.Document, _
    objWordMergedDoc As Word.Document, _
    objWordField As Word.Field, _
    objWordSection As Word.Section

    'Letter generation variables
Dim strInstanceID As String
Dim strProvNum As String
Dim strAuditor As String
Dim strLetterType As String
Dim dtLetterReqDt As Date
Dim strStatus As String
Dim strLocalTemplate As String
Dim strLocalPath As String
Dim oLetterInst As clsLetterInstance

Dim bMergeError As Boolean
Dim strOutputPath As String
Dim strOutputFileName As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iRtnCd As Integer

Dim iCnt As Integer
Dim i As Integer
Dim sMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim sMergeSproc As String
Dim sCombinedDoc As String
Dim sSuperName As String
Dim dctLetterTemplate As Scripting.Dictionary
Dim dtSt As Date
Dim objLetterTemplate As clsLetterTemplate
    
    
    strProcName = ClassName & ".PreviewViewLetters"
    strErrMsg = ""

    Set objWordApp = New Word.Application
    objWordApp.visible = False

    Set cmdGetLetter = New ADODB.Command
'    Set objLetterInfo = New clsLetterTemplate
    
    
    'set local path
    strLocalPath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    If Not FolderExist(strLocalPath) Then CreateFolders (strLocalPath)
    
    ' Set the based path for saving merge doc
    Set rsLetterConfig = GetLetterConfigDetails()
    
    strBasedPath = rsLetterConfig("LetterOutputLocation").Value
    strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    If rsLetterConfig.State = adStateOpen Then rsLetterConfig.Close
    Set rsLetterConfig = Nothing
    
    
    strOutputPath = strBasedPath & "\PREVIEW"
    strAuditor = Replace(Identity.UserName, ".", "")
    strOutputPath = strOutputPath & "\" & strAuditor
    DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
    CreateFolders (strOutputPath & "\")
    
    bMergeError = False
      

    ' start processing letters
    iCnt = 0
    
    DoEvents
    DoEvents
    
    Set dctLetterTemplate = New Scripting.Dictionary
    For Each oLetterInst In cdctSelectedLetters.Letters
        If oLetterInst.LetterQueueStatus <> "P" Then ' Or oLetterInst.LetterQueueStatus = "R" Then

            strInstanceID = oLetterInst.InstanceId
            strProvNum = oLetterInst.cnlyProvID
            strLetterType = oLetterInst.LetterType
            dtLetterReqDt = oLetterInst.LetterReqDt
            strAuditor = oLetterInst.Auditor
            strStatus = oLetterInst.LetterQueueStatus
            
            iCnt = iCnt + 1

            ' Just copy the templates over again..
            If CopyTemplates(dctLetterTemplate, strLetterType) = False Then
    Stop
                strErrMsg = "There was a problem copying the letter templates to the users temp directory. Cannot proceed!"
                Call ErrorCallStack_Add(clBatchId, "There was a problem copying the letter templates to the users temp directory. Cannot proceed!", strProcName, strLetterType)
                GoTo Block_Err
            End If
            
            If IsEmpty(dctLetterTemplate.Item(strLetterType)) Then
                Stop
            End If
            
            Set objLetterTemplate = dctLetterTemplate.Item(strLetterType)
            strLocalTemplate = objLetterTemplate.TemplateLoc
            
            If objWordApp Is Nothing Then
                Set objWordApp = New Word.Application
            End If
            'When the mail merge runs it keeps the template's Margins....
            Set objWordDoc = objWordApp.Documents.Add(strLocalTemplate, , False) 'tried didn't effect change
            
 
            sMergeSproc = "usp_LETTER_Automation_MailMergeSource_ManualOverrides"
                
'            'add a connolly-internal watermark to the preview letters
            If Not (ADDWATERMARK(objWordApp, objWordDoc, strErrMsg)) Then
                LogMessage strProcName, "ERROR", "There was an error while adding the watermark: " & strErrMsg, strErrMsg
                GoTo Block_Err
            End If

            dtSt = Now  ' For tracking how much time this one takes
            
            ' Set data source for mail merge.  Data will be from new Temp Table
            objWordDoc.MailMerge.OpenDataSource Name:=strODCFile, _
                                SqlStatement:="exec " & sMergeSproc & " '" & oLetterInst.InstanceId & "'"
                            

            ' Perform mail merge.
            objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
            objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
            objWordDoc.MailMerge.Execute Pause:=False
            If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
                objWordApp.visible = True
                '''MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
                bMergeError = True
                objWordApp.ActiveDocument.Activate
                strErrMsg = "Error encountered with mail merge."
                GoTo Block_Err
            End If

            objWordApp.visible = True
            
            Call AddSecPagesCode(objWordApp.ActiveDocument, oLetterInst)
            
            ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
            If oLetterInst.InstanceQRCodePath <> "" Then
                Call AddInstanceIdQRCode(objWordApp.ActiveDocument, oLetterInst.InstanceId, oLetterInst.InstanceQRCodePath)
            End If
    
            ' Save the output doc
            Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
                    
            'Added to rename reprints...
            strOutputFileName = strOutputPath & "\" & strLetterType & "-Preview-" & strInstanceID & ".doc"
            
            Call CreateFolders(strOutputPath)

            If Not FolderExists(strOutputPath) Then
                strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
                GoTo Block_Err
            End If
    
            If oLetterInst.LetterQueueStatus = "R" Then
                strOutputFileName = QualifyFldrPath(strOutputPath) & "" & strLetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
            Else
                strOutputFileName = QualifyFldrPath(strOutputPath) & "" & strLetterType & "-" & oLetterInst.InstanceId & ".doc"
            End If
    
            objWordMergedDoc.spellingchecked = True
            objWordMergedDoc.Repaginate
            DoEvents
            DoEvents

            
            If UnlinkWordFields(objWordApp, objWordMergedDoc, oLetterInst.LetterType) = False Then
                LogMessage strProcName, "LETTER ERROR", "Failed to unlink word fields for some reason!", oLetterInst.InstanceId
            End If

            
            On Error Resume Next
                '' Note to Data Services user:
                '' sometimes this fails - there seems to be some sort of
                '' delay in the filesystem - perhaps it's anti-virus doing it's scan
                '' either way, if you get a run-time error on the next line of code, just press f5
                '' to continue execution - it should work the 2nd time.
                '' the code I've put in will hopefully take care of it if you don't have break on all errors turned
                '' on..
            CreateFolders (strOutputFileName)
            
            objWordMergedDoc.SaveAs strOutputFileName
            If Err.Number <> 0 Then
                SleepEvents 1
                objWordMergedDoc.SaveAs strOutputFileName
                Err.Clear
            End If
            On Error GoTo Block_Err

        
            If gbVerboseLogging = True Then Debug.Print ProcessTookHowLong(dtSt)
            
            If Not objWordDoc Is Nothing Then '07/01/2013
                objWordDoc.Close wdDoNotSaveChanges
            End If
            
            If Not objWordMergedDoc Is Nothing Then '07/01/2013
                objWordMergedDoc.Close wdDoNotSaveChanges
            End If
            
            Set objWordDoc = Nothing
            Set objWordMergedDoc = Nothing
            
            ' display progress  '/ 2
            sMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax & vbCrLf & _
                        "Provider = " & strProvNum & vbCrLf & _
                        "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
            fmrStatus.ProgVal = iCnt
            fmrStatus.StatusMessage sMsg
    
            If fmrStatus.ProgMax = lngProgressCount Then
                msgIcon = vbInformation
            Else
                msgIcon = vbExclamation
            End If
            
            'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
            If fmrStatus.EvalStatus(2) = True Then
                    sMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
                    fmrStatus.StatusMessage sMsg
                    DoEvents
                    strErrMsg = sMsg
                    GoTo Block_Err
            End If
            
            DoEvents
            DoEvents
        End If 'end if to ensure the items are marked as W to print
    Next
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo Block_Err
    End If
    Set objWordApp = Nothing

    PreviewViewLetters = True


    ' HEre, ONLY prompt if the user is Data Services:
    sSuperName = Identity.UserSupervisorId()
    
    If UCase(sSuperName) <> "DATA CENTER" Then
        i = MsgBox("Would you like to combine the Preview Letters into one file?", vbYesNo)
        If i = vbYes Then
            If (Not CombineDocs(strOutputPath, True)) Then 'could pass stroutputpath, true to delete all but he combined here.  clients choice.
                GoTo Block_Err
            End If
        End If
    End If
    
    
    
    Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
          


      
Block_Exit:

    If Not objMasterDoc Is Nothing Then '07/01/2013
        objMasterDoc.Close wdDoNotSaveChanges
    End If


    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    ' Release references.
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    Set objMasterDoc = Nothing
    
    Set cmdGetLetter = Nothing
    Set rsLetterConfig = Nothing
    If Not objWordApp Is Nothing Then
        objWordApp.Quit wdDoNotSaveChanges
        Set objWordApp = Nothing
    End If
    Exit Function
Block_Err:
    If strErrMsg <> "" Then
        LogMessage TypeName(Me) & ".PreviewViewLetters", "ERROR", "An error occurred: " & strErrMsg
        
        MsgBox strErrMsg, vbCritical
    Else
        ReportError Err, strProcName
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical

    End If
    PreviewViewLetters = False
    GoTo Block_Exit
End Function


''' This function will look for a bookmark named 'SecPages' and will replace that with
''' the Sec Pages field
'Private Function AddSecPagesCode(objWordDoc As Object) As Boolean
''Private Function AddSecPagesCode(objWordDoc As Word.Document) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim objWordField As FormField
'
''Dim objWordSection As Object
'Dim objWordSection As Object
'Dim iFoot As Integer
'Dim oFooter As Object
'Dim oField As Object
'Dim oRange As Object
'Dim sBookmarkName As String
'Dim saryBkmarks(1) As String
'Dim iBkmark As Integer
'
'
'    strProcName = TypeName(Me) & ".AddSecPagesCode"
'
'    saryBkmarks(0) = "SecPages"
'    saryBkmarks(1) = "SecPages2"
'
'    For iBkmark = 0 To UBound(saryBkmarks)
'        sBookmarkName = saryBkmarks(iBkmark)
'
'        If IsBookMark(objWordDoc, sBookmarkName) = False Then
'            ' nothing else to do
'            GoTo NextBkmark
'        End If
'
'        For Each objWordSection In objWordDoc.Sections
'            For iFoot = 1 To objWordSection.Footers.Count
'                    ' shape range should take care of the text box..
'                    '' Note: may want to do the Headers here too
'                Set oFooter = objWordSection.Footers.Item(iFoot)
'                    ' Looks like 2007 + use a different means..
'                If oFooter.Shapes.Count > 0 Then
'                    Set oRange = objWordDoc.bookmarks(sBookmarkName).Range
'
'                    Set oField = oFooter.Shapes(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'                ElseIf oFooter.Range.ShapeRange.Count > 0 Then
'
'                    Set oRange = objWordDoc.bookmarks(sBookmarkName).Range
'                    Set oField = oFooter.Range.ShapeRange.Item(1).TextFrame.TextRange.Fields.Add(oRange, wdFieldSectionPages, "\#""00""", True)
'
'                End If
'            Next
'        Next
'NextBkmark:
'    Next
'
'    AddSecPagesCode = True  ' In this case true = no error.. :)
'
'Block_Exit:
'    Set oField = Nothing
'    Set oRange = Nothing
'    Set oFooter = Nothing
'    Set objWordSection = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function


'
'Private Function IsBookMark(objWordDoc As Object, sBookmarkName As String) As Boolean
''Private Function IsBookMark(objWordDoc As Word.Document, sBookMarkName As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oBkmk As Object
''Dim oBkmk As Word.Bookmark
'
'    strProcName = TypeName(Me) & ".IsBookMark"
'
'    For Each oBkmk In objWordDoc.bookmarks
'        If UCase(oBkmk.Name) = UCase(sBookmarkName) Then
'            IsBookMark = True
'            GoTo Block_Exit
'        End If
'    Next
'
'Block_Exit:
'    Set oBkmk = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function

Private Function ADDWATERMARK(objWordApp As Variant, objWordDoc As Variant, ByRef strErrMsg As String) As Boolean
'Private Function ADDWATERMARK(objWordApp As Word.Application, objWordDoc As Word.Document, ByRef strErrMsg As String) As Boolean
On Error GoTo Error_Encountered

     objWordDoc.Select
    'take our open word document. add a watermark since we are previewing.  This will deter the auditors from printing this.  (ALTHOUGH they could remove it manually)
    objWordApp.ActiveDocument.Sections(1).Range.Select
    objWordApp.ActiveWindow.ActivePane.View.seekview = 9 'wdSeekCurrentPageHeader
    'objWordApp.activewindow.activepane.View.seekview = 9 'wdseekcurrentPageHeader
    
    objWordApp.selection.HeaderFooter.Shapes.AddTextEffect(vbNull, _
        "CONNOLLY - INTERNAL", "Times New Roman", 1, False, False, 0, 0).Select
    objWordApp.selection.ShapeRange.Name = "PowerPlusWaterMarkObject1"
    objWordApp.selection.ShapeRange.TextEffect.NormalizedHeight = False
    objWordApp.selection.ShapeRange.Line.visible = False
    objWordApp.selection.ShapeRange.Fill.visible = True
    objWordApp.selection.ShapeRange.Fill.Solid
    objWordApp.selection.ShapeRange.Fill.ForeColor.RGB = RGB(153, 153, 153)
    objWordApp.selection.ShapeRange.Fill.Transparency = 0
    objWordApp.selection.ShapeRange.Rotation = 315
    objWordApp.selection.ShapeRange.LockAspectRatio = True
    objWordApp.selection.ShapeRange.top = -999995  'wdShapeCenter
    objWordApp.selection.ShapeRange.left = -999995 'wdShapeCenter
    objWordApp.selection.ShapeRange.Height = objWordApp.inchestopoints(2) '1.69)
    objWordApp.selection.ShapeRange.Width = objWordApp.inchestopoints(6.77) '500 '487.45 'InchesToPoints(6.77)
    objWordApp.selection.ShapeRange.WrapFormat.AllowOverlap = True
    objWordApp.selection.ShapeRange.WrapFormat.Side = 3 'wdWrapNone
    objWordApp.selection.ShapeRange.WrapFormat.Type = 3
    objWordApp.selection.ShapeRange.RelativeHorizontalPosition = 0 ' wdRelativeVerticalPositionMargin
    objWordApp.selection.ShapeRange.RelativeVerticalPosition = 0 '  wdRelativeVerticalPositionMargin
    objWordApp.ActiveWindow.ActivePane.View.seekview = 0 'wdSeekMainDocument
    
    ADDWATERMARK = True
    Exit Function
                    
Error_Encountered:
    strErrMsg = Nz(strErrMsg, "") & "Error adding watermark.  " & Err.Source & " " & Err.Number & " " & Err.Description
    ADDWATERMARK = False
    
End Function

'Function created by TGH on 10-3-08
'   This function will go through all Word docs (excluding temp docs) and combine them together into one file.  It takes will by default create a new file named CombinedDocs unless it is passed a different name.
'           This does not include a preview window for progress but it runs pretty quick...
'  Parameters: StrPath -> The folder containing all of the word docs we would like to combine.
'                      StrPreviewFileName -> the name of the combined word docs file.  by defaule it will be called combinedDocs.
'         *Note: This function does not Run Recursivily to sub folders.  It would be a simple fix if we were looking to do this.  just run a alternative loop to recurse to the subfolder and call combine docs...

Private Function CombineDocs(strPath As String, Optional KillNonCombined As Boolean, Optional strPreviewFileName As String) As Boolean
On Error GoTo Err:
Dim First As Boolean: First = True
   
   'Dim strPreviewFileName As String
   'strpath = "\\ccaintranet.com\DFS-FLD-01\Audits\AmeriHealth\LETTER_REPOSITORY\LETTERS\PREVIEW\tomhartey"
   
   'the default would be false but doing this in case it is ever passed as null
    KillNonCombined = Nz(KillNonCombined, False)
 
   
    If Nz(strPreviewFileName, "") = "" Then
        strPreviewFileName = strPath & "\CombinedDocs.DOC"
    End If
   
Dim fso As Scripting.FileSystemObject
Dim CurrentFolder As Scripting.Folder
Dim Files As Scripting.Files, file As Scripting.file
   Set fso = New Scripting.FileSystemObject
   
    
    'Late bind the Word Object library 11
'    Dim objWordApp As Word.Application, _
'        objMasterDoc As Word.Document, _
'        objWordDoc As Word.Document, _
'        objWordField As Word.Field, _
'        objWordSection As Word.Section

Dim objWordApp As Object, _
    objMasterDoc As Object, _
    objWordDoc As Object, _
    objWordField As Object, _
    objWordSection As Object

Dim i As Integer

    
    Set objWordApp = CreateObject("Word.Application")
    'Set objWordApp = New Word.Application
    
        ' Create an instance of objWord, and make it invisible.
    objWordApp.visible = False
    
    Set objMasterDoc = objWordApp.Documents.Add()   'tgh 3/18/08
                    objMasterDoc.spellingchecked = True 'this is needed to shut down popup for too many spelling errors
                    Sleep 1000
                    objMasterDoc.SaveAs strPreviewFileName
        
    'get the appropriate folder
    Set CurrentFolder = fso.GetFolder(strPath)
    Set Files = CurrentFolder.Files
    For Each file In Files  'for each file in the folders
        If InStr(1, file.Name, ".doc") > 0 And file.Path <> strPreviewFileName And InStr(1, file.Name, "~") = 0 Then
        
            If First Then
                First = False
                Set objWordDoc = objWordApp.Documents.Open(file.Path)
        
                'set the margins for our newly created word doc to be the same as the first word doc in the folder.
                objMasterDoc.PageSetup.LeftMargin = objWordDoc.PageSetup.LeftMargin
                objMasterDoc.PageSetup.RightMargin = objWordDoc.PageSetup.RightMargin
                objMasterDoc.PageSetup.TopMargin = objWordDoc.PageSetup.TopMargin
                objMasterDoc.PageSetup.BottomMargin = objWordDoc.PageSetup.BottomMargin
                objMasterDoc.PageSetup.HeaderDistance = objWordDoc.PageSetup.HeaderDistance
                objMasterDoc.PageSetup.FooterDistance = objWordDoc.PageSetup.FooterDistance
                objWordDoc.Select   'make sure we close the right one and keep adding to the other...
        
                objWordDoc.Close
            Else
        '        objWordApp.Selection.InsertBreak    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
                objWordApp.selection.InsertBreak (2) '   wdSectionBreakNextPage    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
                ' I need to make the fields static here I think..
        
            End If
        
            objWordApp.ActiveDocument.spellingchecked = True
        
        
        
            objWordApp.selection.InsertFile (file.Path)
            If KillNonCombined Then 'added if they only want the combined doc left...
                file.Delete True
                    '                Kill file.path
            End If
Dim intF As Integer

            '' 20130411 KD I'm not sure if we need to do this here or not but I don't think it hurts! (yes, need to do it!)
            With objWordApp.ActiveDocument
                For i = 1 To .Sections.Count
                    .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                    .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                    
                    Set objWordSection = .Sections(i)
'                    Stop
                    For intF = 1 To objWordSection.Footers.Count
                        For Each objWordField In objWordSection.Footers.Item(intF).Range.Fields
                            Debug.Print objWordField.Code
                            objWordField.Update
'                            objWordField.Unlink
                        Next
                        
                    Next
                Next i
            End With
        
            objWordApp.ActiveDocument.Repaginate                                                     'thieu 1/16/08
            Sleep 1000
            objWordApp.ActiveDocument.SaveAs strPreviewFileName
        
        End If
    Next

    ' 20130219 KD: This was here before but I wanted to comment on it: It's to make sure that the sections (each document) starts at page 1
    With objWordApp.ActiveDocument
        For i = 1 To .Sections.Count
            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
            .Sections(i).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        Next i
    End With

    objWordApp.ActiveDocument.Repaginate                                                     'thieu 1/16/08

    Sleep 1000
    objWordApp.ActiveDocument.SaveAs strPreviewFileName

    'Destruct our footprint.
     Set objWordDoc = Nothing
    On Error Resume Next
       objWordApp.Quit (0)
     Set objWordApp = Nothing
     Set fso = Nothing
   
   CombineDocs = True
   Exit Function

Err:
    LogMessage TypeName(Me) & ".CombineDocs", "ERROR", Err.Description
    MsgBox Err.Description

   'Destruct our footprint.
   Set objWordDoc = Nothing
   On Error Resume Next
       objWordApp.Quit (0)
   
   Set objWordApp = Nothing
   Set fso = Nothing

    CombineDocs = False
End Function


Private Function GetLetterConfigDetails() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
    
    strProcName = ClassName & ".GetLetterConfigDetails"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_GetConfig"
        .Parameters.Refresh
        .Parameters("@pAccountId") = IIf(gintAccountID = 0, 1, gintAccountID)
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not get the letter configurations for some reason. DB Connectivity?"
            GoTo Block_Exit
        End If
    End With
    
Block_Exit:
    Set GetLetterConfigDetails = oRs
        ' don't close the RS because that'll close the returned RS
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'
'Private Function GenerateLetters(fmrStatus As Form_ScrStatus, Optional bAtLeastOneErrored As Boolean = False) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim dctLetterTemplate As Scripting.Dictionary
'Dim oLetterToGenerate As clsLetterInstance
'Dim objLetterTemplate As clsLetterTemplate
'Dim iCnt As Integer
'Dim bTemplateFound As Boolean
'Dim strInstanceID As String
'Dim strProvNum As String
'Dim strLetterType As String
'Dim strOutputFileName As String
'Dim strAuditor As String
'Dim strErrMsg As String
'Dim strProgressMsg As String
'Dim lngProgressCount As Long
'Dim msgIcon As Integer
'Dim rsLetterConfig As ADODB.RecordSet
'Dim strODCFile As String
'Dim strOutputLocation As String
'Dim iPageCount As Integer
'Dim dtStart As Date
'
'
'    strProcName = ClassName & ".GenerateLetters"
'
'    Set rsLetterConfig = GetLetterConfigDetails()
'
'    If rsLetterConfig.recordCount = 0 Then
'        strErrMsg = "ERROR: Letter configuration parameters is missing"
'        GoTo Block_Err
'    ElseIf rsLetterConfig.recordCount > 1 Then
'        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
'        GoTo Block_Err
'    Else
'        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
'        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
'        gbVerboseLogging = IIf(rsLetterConfig("VerboseLogging").Value = 0, False, True)
'    End If
'
'    ' setup progress screen that is passed to this function
'
'    '' 20130821 KD: Not really sure what this is all about..
'    DoEvents
'
'
'    ' in order to get here we need to have come through the Generate button
'    ' which means we should have our form scoped cdctSelectedLetters
'
'    If cdctSelectedLetters.Count < 1 Then
'        LogMessage strProcName, "ERROR", "Got to this sub without any selected records for some reason!?!?!"
'        GoTo Block_Exit
'    End If
'
'    dtStart = Now()
'    LogMessage strProcName, "LETTER GEN START", "Starting to generate " & CStr(cdctSelectedLetters.Count) & " letters"
'
'    Set dctLetterTemplate = New Scripting.Dictionary
'    For Each oLetterToGenerate In cdctSelectedLetters.Letters
'
'        ' display progress
'        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
'                    "Provider = " & strProvNum & vbCrLf & _
'                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
'        fmrStatus.ProgVal = iCnt
'        fmrStatus.StatusMessage strProgressMsg
'
'        If fmrStatus.ProgMax = lngProgressCount Then
'            msgIcon = vbInformation
'        Else
'            msgIcon = vbExclamation
'        End If
'
'
'        strInstanceID = oLetterToGenerate.InstanceID
'        strProvNum = oLetterToGenerate.ProvNum
'        strLetterType = oLetterToGenerate.LetterType
'
'        bTemplateFound = False
'
'        ' Just copy the templates over again..
'        If CopyTemplates(dctLetterTemplate, strLetterType) = False Then
'            strErrMsg = "There was a problem copying the letter templates to the users temp directory. Cannot proceed!"
'            Call ErrorCallStack_Add(clBatchId, "There was a problem copying the letter templates to the users temp directory. Cannot proceed!", strProcName, strLetterType)
'            GoTo Block_Err
'        End If
'
'        If IsEmpty(dctLetterTemplate.Item(strLetterType)) Then
'            Stop
'        End If
'        Set objLetterTemplate = dctLetterTemplate.Item(strLetterType)
'
'        iCnt = iCnt + 1
'
'
'        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
'        If fmrStatus.EvalStatus(2) = True Then
'            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
'            fmrStatus.StatusMessage strProgressMsg
'            DoEvents
'            strErrMsg = strProgressMsg
'            GoTo Block_Exit
'        End If
'
'
'        If PrintLetterInstance(oLetterToGenerate, objLetterTemplate.TemplateLoc, strOutputFileName, strOutputLocation, strProvNum, _
'                                        strODCFile, strLetterType, iPageCount) = False Then
'            LogMessage strProcName, "ERROR", "Printing the letter instance failed for InstanceId: " & CStr(oLetterToGenerate.InstanceID), strErrMsg
'                Call ErrorCallStack_Add(clBatchId, "There was a problem generating a letter. Cannot proceed!", strProcName, strLetterType)
'            bAtLeastOneErrored = True
'            GoTo NextLetter
'        End If
'
'        If cdctSelectedLetters.UpdateLetter(oLetterToGenerate) = False Then
'            Stop
'        End If
'
'
'NextLetter:
'
'            '' 20130821 KD: Not really sure what this is all about..
'        DoEvents
'        DoEvents
'        DoEvents
'
'        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
'        If fmrStatus.EvalStatus(2) = True Then
'            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
'            fmrStatus.StatusMessage strProgressMsg
'            DoEvents
'            strErrMsg = strProgressMsg
'            GoTo Block_Exit
'        End If
'
'        ' display progress
'        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
'                    "Provider = " & strProvNum & vbCrLf & _
'                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
'        fmrStatus.ProgVal = iCnt
'        fmrStatus.StatusMessage strProgressMsg
'    Next
'
'    LogMessage strProcName, "LETTERS", ProcessTookHowLong(dtStart)
'
'
'    ' Notify user we are done.
'    cboViewType.SetFocus
'    cboViewType = cboViewType.ItemData(1) 'JS change 20130305
'    Call cboViewType_AfterUpdate
'
'    Me.txtFromDate = Format(Now, "mm/dd/yyyy")
'    Me.txtThroughDate = Format(Now, "mm/dd/yyyy")
'
'    cmdRefresh_Click 'can't run with open trans
'
'    GenerateLetters = True
'
'Block_Exit:
'    If GenerateLetters = False Then bAtLeastOneErrored = True
'
'    Set rsLetterConfig = Nothing
'    Exit Function
'Block_Err:
'    If Err.Number <> 0 Then
'        ReportError Err, strProcName
'    End If
'    If strErrMsg <> "" Then
'        MsgBox strErrMsg, vbCritical
'    End If
'
'    GoTo Block_Exit
'End Function



Private Function GenerateLetters(fmrStatus As Form_ScrStatus, Optional bAtLeastOneErrored As Boolean = False, Optional lThisBatch As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim dctLetterTemplate As Scripting.Dictionary
Dim oLetterToGenerate As clsLetterInstance
Dim objLetterTemplate As clsLetterTemplate
Dim iCnt As Integer
Dim bTemplateFound As Boolean
Dim strInstanceID As String
Dim strProvNum As String
Dim strLetterType As String
Dim strOutputFileName As String
Dim strAuditor As String
Dim strErrMsg As String
Dim strProgressMsg As String
Dim lngProgressCount As Long
Dim msgIcon As Integer
Dim rsLetterConfig As ADODB.RecordSet
Dim strODCFile As String
Dim strOutputLocation As String
Dim iPageCount As Integer
Dim dtStart As Date
Dim oRs As ADODB.RecordSet
Dim oCn As ADODB.Connection
Dim oCmd As ADODB.Command
Dim sErrMsg As String


    
    strProcName = ClassName & ".GenerateLetters"
    
    Set rsLetterConfig = GetLetterConfigDetails()
    
    If rsLetterConfig.recordCount = 0 Then
        strErrMsg = "ERROR: Letter configuration parameters is missing"
        GoTo Block_Err
    ElseIf rsLetterConfig.recordCount > 1 Then
        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
        GoTo Block_Err
    Else
        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
        gbVerboseLogging = IIf(rsLetterConfig("VerboseLogging").Value = 0, False, True)
    End If
    
    ' setup progress screen that is passed to this function
    
    '' 20130821 KD: Not really sure what this is all about..
    DoEvents
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = CodeConnString
        .CursorLocation = adUseClient
        .Open
    End With

    Set oCmd = New ADODB.Command
    With oCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_LETTER_Automation_GetTemplateDetailsByType"
        .ActiveConnection = oCn
        .Parameters.Refresh
    End With



    ' in order to get here we need to have come through the Generate button
    ' which means we should have our form scoped cdctSelectedLetters
    ' I don't like that idea though.. Let's get it here
Stop
    If cdctSelectedLetters Is Nothing Then
        Set cdctSelectedLetters = GetSelectedItems(, , "G", sErrMsg)
        If cdctSelectedLetters Is Nothing Then
            bAtLeastOneErrored = True
            GoTo Block_Exit
        End If
    End If
    
    If cdctSelectedLetters.Count < 1 Then
        LogMessage strProcName, "ERROR", "Got to this sub without any selected records for some reason!?!?!"
        GoTo Block_Exit
    End If
    
    dtStart = Now()
    LogMessage strProcName, "LETTER GEN START", "Starting to generate " & CStr(cdctSelectedLetters.Count) & " letters"
    
    If lThisBatch = 0 Then
    Stop
        lThisBatch = AssignBatchId(cdctSelectedLetters)
    End If
    
    Set dctLetterTemplate = New Scripting.Dictionary
    For Each oLetterToGenerate In cdctSelectedLetters.Letters
    
        ' display progress
        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
                    "Provider = " & strProvNum & vbCrLf & _
                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
        fmrStatus.ProgVal = iCnt
        fmrStatus.StatusMessage strProgressMsg

        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If
        
    
        strInstanceID = oLetterToGenerate.InstanceId
        strProvNum = oLetterToGenerate.ProvNum
        strLetterType = oLetterToGenerate.LetterType
        
        bTemplateFound = False
        
        ' Just copy the templates over again..
        
        If CopyTemplates(dctLetterTemplate, strLetterType) = False Then
Stop
            strErrMsg = "There was a problem copying the letter templates to the users temp directory. Cannot proceed!"
            Call ErrorCallStack_Add(clBatchId, "There was a problem copying the letter templates to the users temp directory. Cannot proceed!", strProcName, strLetterType)
            GoTo Block_Err
        End If
    
        If IsEmpty(dctLetterTemplate.Item(strLetterType)) Then
            Stop
        End If
        Set objLetterTemplate = dctLetterTemplate.Item(strLetterType)
        
        iCnt = iCnt + 1
    
    
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If fmrStatus.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
            fmrStatus.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            GoTo Block_Exit
        End If

        
        ' Assign instance id and batchid for this run
        Call AssignInstanceId(oLetterToGenerate, lThisBatch)
        
'        If PerformIndividualMailMerges(oRs, dctLetterTemplate) = True Then
'Stop
'        End If
        If PrintLetterInstance(oLetterToGenerate, objLetterTemplate.TemplateLoc, strOutputFileName, strOutputLocation, strProvNum, _
                                        strODCFile, strLetterType, iPageCount) = False Then
            LogMessage strProcName, "ERROR", "Printing the letter instance failed for InstanceId: " & CStr(oLetterToGenerate.InstanceId), strErrMsg
                Call ErrorCallStack_Add(clBatchId, "There was a problem generating a letter. Cannot proceed!", strProcName, strLetterType)
            bAtLeastOneErrored = True
            GoTo NextLetter
        End If
        
        If cdctSelectedLetters.UpdateLetter(oLetterToGenerate) = False Then
            Stop
        End If


NextLetter:
        
            '' 20130821 KD: Not really sure what this is all about..
        DoEvents
        DoEvents
        DoEvents
        
        'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
        If fmrStatus.EvalStatus(2) = True Then
            strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
            fmrStatus.StatusMessage strProgressMsg
            DoEvents
            strErrMsg = strProgressMsg
            GoTo Block_Exit
        End If

        ' display progress
        strProgressMsg = "Generating Record " & iCnt & " / " & fmrStatus.ProgMax / 2 & vbCrLf & _
                    "Provider = " & strProvNum & vbCrLf & _
                    "InstanceID = " & strInstanceID & vbCrLf & "Letter type = " & strLetterType
        fmrStatus.ProgVal = iCnt
        fmrStatus.StatusMessage strProgressMsg
    Next
    
    LogMessage strProcName, "LETTERS", ProcessTookHowLong(dtStart)

    
    ' Notify user we are done.
    cboViewType.SetFocus
    cboViewType = cboViewType.ItemData(1) 'JS change 20130305
    Call cboViewType_AfterUpdate
    
    Me.txtFromDate = Format(Now, "mm/dd/yyyy")
    Me.txtThroughDate = Format(Now, "mm/dd/yyyy")
    
    cmdRefresh_Click 'can't run with open trans
    
    GenerateLetters = True
    
Block_Exit:
    If GenerateLetters = False Then bAtLeastOneErrored = True

    Set rsLetterConfig = Nothing
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    End If
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    End If
    
    GoTo Block_Exit
End Function

Private Function AssignInstanceId(oLetterToGenerate As clsLetterInstance, lBatchId As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & "AssignInstanceId"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AssignInstanceIds_ManualOverride"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pDynamicInstanceId") = oLetterToGenerate.InstanceId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        oLetterToGenerate.InstanceId = .Parameters("@pRealInstanceId").Value
    End With

    AssignInstanceId = oLetterToGenerate.InstanceId
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function AssignBatchId(dctSelLetters As clsLetterInstanceDct) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sIdList As String
Dim saryIdList() As String
Dim lBatchId As Long
Dim oLtr As clsLetterInstance
Dim vKey As Variant

    strProcName = ClassName & "AssignBatchId"
    
    For Each oLtr In cdctSelectedLetters.Letters
        sIdList = sIdList & oLtr.InstanceId + ","
    Next
    If InStr(1, sIdList, ",") > 0 Then
        sIdList = left(sIdList, Len(sIdList) - 1)
    End If
    saryIdList = Split(sIdList, ",")
    sIdList = MultipleValuesToXml("dynamicinstanceid", saryIdList)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AssignBatchIds_ManualOverride"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pIDList") = sIdList
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        lBatchId = Nz(.Parameters("@pBatchId").Value, 0)
    End With


    AssignBatchId = lBatchId
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Function PrintLetterInstance(oLetterInst As clsLetterInstance, pstrTemplateName As String, _
            pstrOutputFileName As String, pstrOutputBasePath As String, pstrProvNum As String, _
            pstrODCFile As String, pstrLetterType As String, _
            Optional iPageCount As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oCn As ADODB.Connection

Dim oCmd As ADODB.Command
Dim strSQLcmd As String
Dim bMergeError As Boolean
Dim strOutputPath As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iRtnCd As Integer
Dim varItem As Variant
Dim iAnswer As Integer
Dim iCnt As Integer
Dim i As Integer
Dim dThisOne As Date
Dim dtSt As Date
Dim dFileSize As Double
Dim dSleepTime As Double
Dim objLetterInfo As clsLetterTemplate
Dim objWordApp As Word.Application, _
    objWordDoc As Word.Document, _
    objWordMergedDoc As Word.Document
      
      Debug.Print "Outfile path: " & pstrOutputBasePath
      
    strProcName = ClassName & ".PrintLetterInstance"
    
    Set oAdo = New clsADO
    oAdo.ConnectionString = CodeConnString
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = CodeConnString
    oCn.CursorLocation = adUseClient
    oCn.Open
    
    '' check to make sure that the transaction is supported
    Dim oProp As ADODB.Property, iPropCnt As Integer
'    For iPropCnt = 0 To oCn.Properties.Count
'        Set oProp = oCn.Properties(iPropCnt)
'    Next
'    For Each oProp In oCn.Properties
'        Debug.Print oProp.Name
''        Stop
'    Next
'    Set oProp = Nothing
'   Stop
'   For Each oProp In oAdo.CurrentConnection.Properties
'        Debug.Assert oProp.Name <> "Transaction DDL"
'   Next
'
'   Stop
   
    Set objLetterInfo = New clsLetterTemplate
    strErrMsg = ""

    Set objWordApp = New Word.Application
    objWordApp.visible = False
    
    ' check if template exists
    strChkFile = Dir(pstrTemplateName)
    If strChkFile = "" Then
        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
        GoTo Block_Err
    End If

    ' open template
    Set objWordDoc = objWordApp.Documents.Add(pstrTemplateName, , False)
       
    ' load letter info
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    
    '' KD Added this for now to deal with the QR 2D barcodes..
    '' we'll change this and modify the real usp when we decide to
    ''  "go live" with it.
Dim sMergeSproc As String

'    If UCase(oLetterInst.LetterType) = "VADRA_QR" Then
'        oCmd.CommandText = "usp_LETTER_Get_Info_load"
        sMergeSproc = "usp_LETTER_Automation_MailMergeSource"
'    Else
'        oCmd.CommandText = "usp_LETTER_Get_Info_load"
'        sMergeSproc = "usp_LETTER_Get_Info"
'    End If
    
''    oCmd.Parameters.Refresh
''    oCmd.Parameters("@InstanceID") = oLetterInst.InstanceID
''    oCmd.Execute
    
'    strErrMsg = Trim(oCmd.Parameters("@ErrMsg").Value) & ""
'    If strErrMsg <> "" Then
'        GoTo Block_Err
'    End If
    
    dtSt = Now  ' For tracking how much time this one takes
    
    ' Set data source for mail merge.  Data will be from new Temp Table
    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
                        SqlStatement:="exec " & sMergeSproc & " '" & oLetterInst.InstanceId & "'"
                    


    ' Perform mail merge.
    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        '''MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        strErrMsg = "Error encountered with mail merge."
        GoTo Block_Err
    End If

'    objWordApp.visible = True
    
    Call AddSecPagesCode(objWordApp.ActiveDocument, oLetterInst)
    
    ' 20130219 KD: Add the Sec Pages field in the footer for barcodes
    If oLetterInst.InstanceQRCodePath <> "" Then
        Call AddInstanceIdQRCode(objWordApp.ActiveDocument, oLetterInst.InstanceId, oLetterInst.InstanceQRCodePath)
    End If
    
    
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
    Call CreateFolders(strOutputPath)

    If Not FolderExists(strOutputPath) Then
        strErrMsg = "Provider folder for letter was not created for instance: " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        GoTo Block_Err
    End If
    
    If oLetterInst.LetterQueueStatus = "R" Then
    'If pstrInstanceStatus = "R" Then
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & oLetterInst.InstanceId & ".doc"
    Else
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & oLetterInst.InstanceId & ".doc"
    End If
    
    objWordMergedDoc.spellingchecked = True
    objWordMergedDoc.Repaginate
    DoEvents
    DoEvents

    
    If UnlinkWordFields(objWordApp, objWordMergedDoc, oLetterInst.LetterType) = False Then
        LogMessage strProcName, "LETTER ERROR", "Failed to unlink word fields for some reason!", oLetterInst.InstanceId
    End If

    
    On Error Resume Next
    '' Note to Data Services user:
    '' sometimes this fails - there seems to be some sort of
    '' delay in the filesystem - perhaps it's anti-virus doing it's scan
    '' either way, if you get a run-time error on the next line of code, just press f5
    '' to continue execution - it should work the 2nd time.
    '' the code I've put in will hopefully take care of it if you don't have break on all errors turned
    '' on..
''    pstrOutputFileName = Replace(pstrOutputFileName, "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\LETTERS\", "\\ccaintranet.com\dfs-cms-ds\Data\CMS\AnalystFolders\KevinD\QR_Codes\pdf_test")
    CreateFolders (pstrOutputFileName)
    
    objWordMergedDoc.SaveAs pstrOutputFileName
    If Err.Number <> 0 Then
        SleepEvents 1
        objWordMergedDoc.SaveAs pstrOutputFileName
        Err.Clear
    End If
    On Error GoTo Block_Err

    
    With oLetterInst
        If .LetterBatchId = 0 Then
            .LetterBatchId = Me.MostRecentBatchId
        End If

            '' KD: Idea: the below takes a long time for Word to determine how many pages
            '' so we should probably move this to some other process
            '' that runs virtually all the time.  We would look at a QUEUE table
            '' and get the number of pages in those documents and then save them in the
            '' LETTER_Static_Detail table
            '' but this will have to happen #1: Quickly before the user gets the chance
            '' to print the letters
            '' and #2 quietly (without locking the documents and such)
            
            ''' Temporarilly commenting this out until I can come up with a better way..

        DoEvents
        .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
        
        If .PageCount = 1 Then
            dFileSize = FileLen(objWordMergedDoc.Path)
            dSleepTime = (0.25 + (dFileSize * 0.0000001)) * 1000
            
            If dSleepTime > 2000 Then dSleepTime = 1500
            DoEvents

            If gbVerboseLogging = True Then LogMessage strProcName, "LETTER", "Sleeping for " & CStr(dSleepTime) & " milliseconds"
            objWordMergedDoc.Repaginate
            DoEvents
            Sleep dSleepTime
'            Call SleepEvents(CLng(dSleepTime / 1000))

            DoEvents
            .PageCount = objWordMergedDoc.BuiltInDocumentProperties(14)    ' wdPropertyPages = 14
            If gbVerboseLogging = True Then LogMessage strProcName, "LETTER", "Page count: " & CStr(.PageCount)
        End If
'        .PageCount = 0  ' just doing this for now
        .LetterPath = pstrOutputFileName

        
        .SaveStaticDetails
    End With

    
   
    If gbVerboseLogging = True Then Debug.Print ProcessTookHowLong(dtSt)
    
    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo Block_Err
    End If
    Set objWordApp = Nothing

    DoEvents
    DoEvents

    If Not FileExists(pstrOutputFileName) Then
        strErrMsg = "PrintLetterInstance: Letter was not created for instance " & oLetterInst.InstanceId & vbNewLine + vbNewLine & "Process will continue for the instances left."
        LogMessage strProcName, "ERROR", "Generated letter does not exist where it should", pstrOutputFileName
        
        GoTo Block_Err
    End If

    
    ' KD: 5/8/2014: So until today, this, clear tmp table stuff
    ' was BEFORE the claims status got updated, but one of the criteria for the sproc is:
    ' where t2.Status not in ('R','W')
    ' Which means that it's not getting cleared..
    
    ' clear letter info
'    Set oCmd = New ADODB.Command
'    oCmd.ActiveConnection = oAdo.CurrentConnection
'    oCmd.commandType = adCmdStoredProc
'    oCmd.CommandText = "usp_LETTER_Get_Info_tmp_clear"
'    oCmd.Parameters.Refresh
'    'oCmd.Parameters("@pInstanceID") = pstrInstanceID
'    oCmd.Execute
    
                                
    ' start letter transaction
'    oAdo.BeginTrans
    oCn.BeginTrans
    
    ' update LETTER status
    Set oCmd = New ADODB.Command
    ' Update the Letter_Queue_Status
    oCmd.ActiveConnection = oCn
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_Automation_Update_Status"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@pInstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@pLetterPath").Value = pstrOutputFileName
    oCmd.Parameters("@pNextStatus").Value = "G" ' for Generated, not yet printed..
    oCmd.Execute
            
    strErrMsg = Trim(oCmd.Parameters("@pErrMsg").Value) & ""
    If strErrMsg <> "" Then
'        oAdo.RollbackTrans
        oCn.RollbackTrans
        GoTo Block_Err
    End If
                            

                            
                            
    ' update claim status & move to next queue
    ' note, only does this where the letter status is W
    Set oCmd = New ADODB.Command
    'oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.ActiveConnection = oCn
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_LETTER_AuditClaims_Update"
    oCmd.Parameters.Refresh
    oCmd.Parameters("@pInstanceID").Value = oLetterInst.InstanceId
    oCmd.Parameters("@pInstanceStatus").Value = oLetterInst.LetterQueueStatus
    oCmd.Execute
            
    strErrMsg = Trim(Nz(oCmd.Parameters("@pErrMsg").Value, ""))
    If strErrMsg <> "" Then
'        oAdo.RollbackTrans
        oCn.RollbackTrans
        GoTo Block_Err
    End If
                                
                                
    ' commit letter transaction
    oAdo.CommitTrans
    oCn.CommitTrans
    PrintLetterInstance = True
    

Block_Exit:

'    Call SetDefaultPrinterToAcrobat("", sOrigPrinter)

    ' Release references.
    If Not objWordDoc Is Nothing Then '07/01/2013
        objWordDoc.Close wdDoNotSaveChanges
    End If
    
    If Not objWordMergedDoc Is Nothing Then '07/01/2013
        objWordMergedDoc.Close wdDoNotSaveChanges
    End If
    
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
    
    If Not objWordApp Is Nothing Then
        On Error Resume Next
        objWordApp.Quit wdDoNotSaveChanges
        On Error GoTo 0
    End If
    Set objWordApp = Nothing
    
    Set oCmd = Nothing
    Set oAdo = Nothing
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
Block_Err:

    If strErrMsg <> "" Then
'        MsgBox strErrMsg, vbCritical
        LogMessage TypeName(Me) & "PrintLetterInstance-2010", "USAGE DETAIL", strErrMsg
        Call ErrorCallStack_Add(clBatchId, strErrMsg, strProcName)
    Else
'        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
        ReportError Err, strProcName
        Call ErrorCallStack_Add(clBatchId, Err.Description, strProcName)
    End If
    PrintLetterInstance = False
    
    'Call DeleteFile(pstrOutputFileName, False)
    
    GoTo Block_Exit
End Function


Private Function GetUserTempDirectory() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim strTempPath As String
Dim strErrMsg As String

    strProcName = ClassName & ".GetUserTempDirectory"


    strTempPath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTempDir\"
    If CreateFolders(strTempPath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTempPath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTempPath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTempPath
        GoTo Block_Err
    End If
    
    
Block_Exit:
    GetUserTempDirectory = strTempPath
    Exit Function
Block_Err:
    If Err.Number <> 0 Then
        ReportError Err, strProcName
    Else
        LogMessage strProcName, "ERROR", strErrMsg
    End If
    GoTo Block_Exit
End Function

Private Function CopyTemplates(dctTemplatesDict As Scripting.Dictionary, strLetterType As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

Dim oAdo As clsADO
Dim rsLetterTemplate As ADODB.RecordSet
Dim objLetterInfo As clsLetterTemplate
Dim strTemplatePath As String
Dim strLocalTemplate As String
Dim strSQL As String
Dim strChkFile As String
Dim strErrMsg As String
Dim iFolderChkLoop As Integer

    strProcName = ClassName & ".CopyTemplates"

            '' just in case:
    
    If dctTemplatesDict Is Nothing Then Set dctTemplatesDict = New Scripting.Dictionary

    strSQL = "SELECT LetterType, TemplateLoc FROM LETTER_Type WHERE (AccountID = " & CStr(gintAccountID) & " or " & CStr(gintAccountID) & " = 0) AND LetterType = '" & strLetterType & "'"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
            ' get list of templates
        .SQLTextType = sqltext
        .sqlString = strSQL
        Set rsLetterTemplate = .ExecuteRS
    End With
    

    ' create template directory
    iFolderChkLoop = 0
    strTemplatePath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    If CreateFolders(strTemplatePath) = False Then
        LogMessage strProcName, "ERROR", "Could not create user temp folder!", strTemplatePath, True
        GoTo Block_Exit
    End If
            
    If FolderExist(strTemplatePath) = False Then
        strErrMsg = "ERROR: can not create folder " & strTemplatePath
        GoTo Block_Err
    End If
    

    ' copy templates to local directory. Skip if template already there
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            strLocalTemplate = strTemplatePath & "\" & GetFileName(!TemplateLoc)
            
            strChkFile = Dir(strLocalTemplate) & ""
            If strChkFile = "" Then
                strChkFile = Dir(!TemplateLoc) & ""
                If strChkFile <> "" Then
                    If CopyFile(rsLetterTemplate("TemplateLoc").Value, strLocalTemplate, False, strErrMsg) = False Then
Stop
                    End If
                Else
                    strErrMsg = "Error: source template " & rsLetterTemplate("TemplateLoc").Value & " not found"
                    
                End If
            End If
                    
            Set objLetterInfo = New clsLetterTemplate
            objLetterInfo.LetterType = Trim(!LetterType)
            objLetterInfo.TemplateLoc = strLocalTemplate
            
            If dctTemplatesDict.Exists(rsLetterTemplate("LetterType").Value) = True Then
                Set dctTemplatesDict.Item(rsLetterTemplate("LetterType").Value) = objLetterInfo
            Else
                dctTemplatesDict.Add rsLetterTemplate("LetterType").Value, objLetterInfo
            End If

            .MoveNext
        End With
    Loop
    
    CopyTemplates = True
    
Block_Exit:
    If Not rsLetterTemplate Is Nothing Then
        If rsLetterTemplate.State = adStateOpen Then rsLetterTemplate.Close
        Set rsLetterTemplate = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    If strErrMsg <> "" Then
        LogMessage strProcName, "ERROR", strErrMsg
    Else
        ReportError Err, strProcName
    End If
    
    GoTo Block_Exit
End Function


Private Sub DeleteTemplates()
'    Dim Person As New ClsIdentity
    Dim strTemplatePath As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file

    ' delete template directory
    strTemplatePath = cs_USER_TEMPLATE_PATH_ROOT & Identity.UserName & "\LETTERTEMPLATE"
    
    'JS Change 20130305 no more delete the folder, now it will only delete the contents
    'DeleteFolder (strTemplatePath)
    
'    On Error Resume Next
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(strTemplatePath)
    For Each oFile In oFldr.Files
        Dim sFilePath As String
        sFilePath = oFile.Path
        Set oFile = Nothing
        If DeleteFile(sFilePath, False) = False Then
            LogMessage ClassName & ".DeleteTemplates", , "Tried to delete a file - not there, probably fine!", strTemplatePath
        End If
    Next
   

    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

End Sub


Private Function GetSelectedItems(Optional bViewOnly As Boolean = False, Optional sDesiredStatus As String, _
            Optional sDisallowedStatus As String, Optional sErrMsg As String) As clsLetterInstanceDct
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtrInstance As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim varItem As Variant
Dim oRs As ADODB.RecordSet
Dim iSelectedCount As Integer
Dim sLetterInstance As String
Dim sQStatus As String
Dim sProvNum As String
Dim sLetterType As String
Dim dctPos As Scripting.Dictionary
Dim bOk As Boolean

    strProcName = ClassName & ".GetSelectedItems"
    
    Set oLetters = New clsLetterInstanceDct

    If Not (TypeOf Me.lstQueue.RecordSet Is ADODB.RecordSet) Then
        Set oLetters = GetSelectedItemsDAO(bViewOnly)
        GoTo Block_Exit
    End If

    Set oRs = Me.lstQueue.RecordSet
    
    
    Set dctPos = GetADOFieldOrdinalPosition(oRs)
    
    
    ' Look at the whole list for selected, and correct status for printing
    For Each varItem In Me.lstQueue.ItemsSelected
    
        sLetterInstance = UCase(Trim(Me.lstQueue.Column(dctPos.Item("INSTANCEID"), varItem)))

        
        sQStatus = UCase(Trim(Me.lstQueue.Column(dctPos.Item("STATUS"), varItem)))
'Stop    ' Kev: why did they user CnlyProvId and call it ProvNum when there's a ProvNum in the recordset too??
                ' better check that out
        sProvNum = UCase(Trim(Me.lstQueue.Column(dctPos.Item("CNLYPROVID"), varItem)))
        
        sLetterType = UCase(Trim(Me.lstQueue.Column(dctPos.Item("LETTERTYPE"), varItem)))
            ' Only count Re-Printable and W (whatever that stands for)
            ' unless we are viewing instead of generating the letters
        
        If sLetterInstance = "" Then
            If isField(oRs, "DynamicInstanceId") = True Then
                sLetterInstance = UCase(Trim(Me.lstQueue.Column(dctPos.Item("DYNAMICINSTANCEID"), varItem)))
            Else
Stop
                sLetterInstance = sProvNum & "-" & sLetterType & "-" & Trim(Me.lstQueue.Column(dctPos.Item("LETTERREQDT"), varItem))
            End If
        End If
        
        'If sDesiredStatus <> "" And sQStatus = sDesiredStatus Then bOk = True
        '2014:04:25:JS: added the chance to specify a list of statuses:
        If sDesiredStatus <> "" And Nz(InStr(1, sDesiredStatus, sQStatus, vbTextCompare)) > 0 Then bOk = True
        
        If (sQStatus = "Q" Or sQStatus = "R") Or bViewOnly = True Then bOk = True
        
        If sDisallowedStatus <> "" Then
            If InStr(1, sQStatus, sDisallowedStatus, vbTextCompare) > 0 Then
                bOk = False
                ' this is disallowed, stop, don't process anything
                sErrMsg = "There is at least 1 record selected with a status of: '" & sQStatus & "' which cannot be processed this way!"
                Set oLetters = New clsLetterInstanceDct
                GoTo Block_Exit
            End If
        End If
        
        
        If bOk = True Then
            iSelectedCount = iSelectedCount + 1
            Call oLetters.AddLetterInstance(sLetterInstance, sQStatus, sProvNum, sLetterType)
        End If
        bOk = False ' reset
    Next varItem
    
    Set GetSelectedItems = oLetters
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function GetSelectedItemsDAO(Optional bViewOnly As Boolean = False) As clsLetterInstanceDct
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtrInstance As clsLetterInstance
Dim oLetters As clsLetterInstanceDct
Dim varItem As Variant
Dim oRs As DAO.RecordSet
Dim iSelectedCount As Integer
Dim sLetterInstance As String
Dim sQStatus As String
Dim sProvNum As String
Dim sLetterType As String


    strProcName = ClassName & ".GetSelectedItemsDAO"
    
    Set oLetters = New clsLetterInstanceDct

    Set oRs = Me.lstQueue.RecordSet
    oRs.MoveLast
    oRs.MoveFirst
    Sleep 750
    ' Look at the whole list for selected, and correct status for printing
    For Each varItem In Me.lstQueue.ItemsSelected
        sLetterInstance = Trim(Me.lstQueue.Column(oRs.Fields("InstanceID").OrdinalPosition, varItem))
        sQStatus = UCase(Trim(Me.lstQueue.Column(oRs.Fields("Status").OrdinalPosition, varItem)))
'Stop    ' Kev: why did they user CnlyProvId and call it ProvNum when there's a ProvNum in the recordset too??
                ' better check that out
        sProvNum = Trim(Me.lstQueue.Column(oRs.Fields("cnlyProvID").OrdinalPosition, varItem))
        sLetterType = Trim(Me.lstQueue.Column(oRs.Fields("LetterType").OrdinalPosition, varItem))
        
            ' Only count Re-Printable and W (whatever that stands for)
            ' unless we are viewing instead of generating the letters
            
        If (sQStatus = "Q" Or sQStatus = "R") Or bViewOnly = True Then
            iSelectedCount = iSelectedCount + 1

            Call oLetters.AddLetterInstance(sLetterInstance, sQStatus, sProvNum, sLetterType)
        End If
    Next varItem
    
    Set GetSelectedItemsDAO = oLetters
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function GenerateBatchId() As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim lBatchId As Long

    strProcName = ClassName & ".GenerateBatchID"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Assign_Generation_BatchId"
        .Parameters.Refresh
        .Parameters("@pUserId") = Identity.UserName
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error generating a Letter Generation Batch ID", .Parameters("@pErrMsg").Value, True
            GoTo Block_Exit
        Else
            lBatchId = .Parameters("@pBatchId").Value
        End If
    End With
    
    Me.MostRecentBatchId = lBatchId
    
Block_Exit:
    Set oAdo = Nothing
    GenerateBatchId = lBatchId
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function UpdateBatchId(lBatchId As Long, bSuccess As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO


    strProcName = ClassName & ".UpdateBatchId"
    
    If lBatchId = 0 Then
        lBatchId = Me.MostRecentBatchId
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Update_Generation_BatchId"
        .Parameters.Refresh
        .Parameters("@pBatchId") = lBatchId
        .Parameters("@pUserId") = Identity.UserName
        .Parameters("@pSuccess") = IIf(bSuccess, 1, 0)
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error trying to update the batch status", .Parameters("@pErrMsg").Value, True
            GoTo Block_Exit
        End If
    End With
    
Block_Exit:
    Set oAdo = Nothing
    UpdateBatchId = lBatchId
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Private Sub ErrorCallStack_Add(lBatchId As Long, sErrMsg As String, sErrProc As String, Optional sErrDetails As String, Optional lErrNum As Long, Optional bFatal As Boolean, _
    Optional sInstanceId As String, Optional sLetterType As String, Optional sCnlyClaimNums As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oPrintError As clsLetterError

    strProcName = ClassName & ".ErrorCallStack_Add"
    
    Set oPrintError = New clsLetterError
    
    With oPrintError
        .Auditor = Identity.UserName
        .BatchID = lBatchId
        .ErrorDetails = sErrDetails
        .ErrorMessage = sErrMsg
        .ErrorNum = lErrNum
        .ErrorProc = sErrProc
        .FatalError = bFatal
        .InstanceId = sInstanceId
        .LetterType = sLetterType
        .CnlyClaimNums = sCnlyClaimNums
    End With
    
    If ccolErrors Is Nothing Then Set ccolErrors = New Collection
    ccolErrors.Add oPrintError
    
Block_Exit:
    Exit Sub
Block_Err:
     ReportError Err, strProcName
     GoTo Block_Exit
End Sub

Private Sub NotifyUserOfErrors()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtrErr As clsLetterError

Dim sMsgToShow As String


    strProcName = ClassName & "NotifyUserOfErrors"
    
    If ccolErrors Is Nothing Then
        ' no errors, do nothing
        GoTo Block_Exit
    End If
    
    For Each oLtrErr In ccolErrors
        sMsgToShow = AppendToErrMsg(sMsgToShow, oLtrErr)
    Next
    
    If sMsgToShow = "" Then
        sMsgToShow = "One or more errors have occurred. Following are the details that you can save and send to Data Services if you feel something is not quite right." & vbCrLf & vbCrLf & _
            String(100, "#") & vbCrLf & vbCrLf
    End If
    
    ' Gonna stuff it in the clipboard then into notepad..
    Call ClipBoard_SetData(sMsgToShow)
    Shell "Notepad.exe", vbMaximizedFocus
    SendKeys "^v"
    ' we'll open notepad for the user..
    
Block_Exit:
    Set ccolErrors = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Function AppendToErrMsg(ByVal sMessageToAppendTo As String, oLetterErr As clsLetterError) As String
On Error GoTo Block_Exit
Dim strProcName As String
Const s_MSG_TEMPLATE As String = "An error occurred:" & vbCrLf & _
        "USER: [%AUDITOR%] " & vbCrLf & _
        "InstanceID: [%INSTANCEID%] " & vbCrLf & _
        "CnlyClaimNums: [%CNLYCLAIMNUMS%] " & vbCrLf & _
        "Letter Type: [%LETTERTYPE%] " & vbCrLf & _
        "Batch ID: [%BATCHID%] " & vbCrLf & vbCrLf & vbCrLf & _
        "Error [%ERRORNUM%] Details: [%ERRORMESSAGE%] " & vbCrLf & _
        "In: [%ERRORPROC%] " & vbCrLf & vbCrLf & _
        "Was fatal?  [%FATALERROR%] " & vbCrLf & vbCrLf & vbCrLf & _
        "Additional Information: " & vbCrLf & _
        "[%ERRORDETAILS%] " & vbCrLf

Dim sNewMsg As String

    strProcName = ClassName & ".AppendToErrMsg"
    
    sNewMsg = s_MSG_TEMPLATE
    With oLetterErr
        sNewMsg = Replace(sNewMsg, "[%AUDITOR%]", .Auditor, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%BATCHID%]", CStr(.BatchID), , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%ERRORDETAILS%]", .ErrorDetails, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%ERRORMESSAGE%]", .ErrorMessage, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%ERRORNUM%]", CStr(.ErrorNum), , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%ERRORPROC%]", .ErrorProc, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%FATALERROR%]", CStr(.FatalError), , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%INSTANCEID%]", .InstanceId, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%LETTERTYPE%]", .LetterType, , , vbTextCompare)
        sNewMsg = Replace(sNewMsg, "[%CNLYCLAIMNUMS%]", .CnlyClaimNums, , , vbTextCompare)
        
    End With
    
    If sMessageToAppendTo <> "" Then
        sMessageToAppendTo = sMessageToAppendTo & vbCrLf & vbCrLf
        sMessageToAppendTo = sMessageToAppendTo & String(100, "#")
        sMessageToAppendTo = sMessageToAppendTo & vbCrLf & vbCrLf
    End If
    
    
    AppendToErrMsg = sMessageToAppendTo & sNewMsg
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
