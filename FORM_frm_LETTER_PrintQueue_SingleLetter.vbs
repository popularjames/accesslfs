Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'This Print Letter Form is to take all the letters selected and generate letters.  The letters marked as 'w' or 'R' are setup to print and are archived
Private mstrAuditor As String

'for resizing
Private ColResize1 As clsAutoSizeColumns
Private ColReSize2 As clsAutoSizeColumns
Private lngQueryType As Long '* These are values from msysobjects 1/4/6 = table, 5 = query
Public blab As String

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents frmFilter As Form_frm_GENERAL_Filter
Attribute frmFilter.VB_VarHelpID = -1

Private mReturnDate As Date
Private msAdvancedFilter As String

Const CstrFrmAppID As String = "LetterQueuePrint"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub frmfilter_QueryFormRefresh()

    RefreshMain

End Sub

Private Sub frmFilter_UpdateSql()
    msAdvancedFilter = frmFilter.SQL.WherePrimary
End Sub


Private Sub cmdDeleteQueue_Click()
'    Dim Person As New ClsIdentity
    Dim strPowerUsers As String
    Dim bPowerUser As Boolean
    Dim strFileName As String
    
    ' ADO variables & late bind them
    Dim cn As Variant 'ADODB.Connection
    Dim cmd As Variant 'ADODB.Command
    Dim cmdGetLetter As Variant 'ADODB.Command
        
    Set cn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    Set cmdGetLetter = CreateObject("ADODB.Command")
    
    Dim strSQLcmd As String
    Dim strInstanceID As String
    Dim strStatus As String
    Dim strErrMsg As String
    Dim iRtnCd As Integer
    Dim i As Integer
    
    Dim varItem
    
    'User Entry Needed***************************************************************
    'These users have total power and can delete printed letters
    'individual users can delete these when they have a status of 'W'
    'strPowerUsers = UCase("Alex.Dremann|Damon.Ramaglia|Thieu.Le|Joe.Casella|Tom.Hartey|Robert.Swander")
    'If InStr(1, strPowerUsers, UCase(Person.UserName)) > 0 Then bPowerUser = True
    
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected"
        Exit Sub
    End If
    
    bPowerUser = True
    
    'If bPowerUser = True Then
    '    MsgBox "You are a POWER user!!", vbInformation
    i = MsgBox("Power user: Are you sure you want to delete these records?", vbYesNo)
    If i <> vbYes Then Exit Sub
    
    
    'MsgBox "You are a POWER user!!", vbInformation
    'i = MsgBox("Power user: Are you sure you want to delete these records?", vbYesNo)
    'If i <> vbYes Then Exit Sub
    
    On Error GoTo Error_Encountered
    
    cn.ConnectionString = GetConnectString("v_CODE_Database")
    cn.CommandTimeout = 0
    cn.Open
    cn.CursorLocation = adUseClient
    
    'Begin our transactions
    cn.BeginTrans
    
    'CmdGetLetter setup before we enter the loop
    cmdGetLetter.ActiveConnection = cn
    cmdGetLetter.commandType = adCmdStoredProc
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("LetterName", adChar, adParamOutput, 255, "")
    cmdGetLetter.CommandText = "usp_LETTER_Get_Letter_Name"
    
    'setup command for usp_LETTER_Work_Queue_Forced_Delete
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = cn
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adChar, adParamOutput, 255, "")
    If bPowerUser Then
        cmd.CommandText = "usp_LETTER_Work_Queue_Forced_Delete"
    Else
        cmd.CommandText = "usp_LETTER_Work_Queue_Delete"
    End If
    
    Dim FileArrays() As String
    ReDim Preserve FileArrays(lstQueue.ItemsSelected.Count - 1)
    'initilize the loop
    i = 0
    'So for each item selected go through and store the file path then run the delete stored proc.
    'If we error out before the end of the sub the rollback will correct the sql and the kill statement will
    'not be executed.
    For Each varItem In lstQueue.ItemsSelected
        strInstanceID = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
        strStatus = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem))
        'if the user is a power user or the case is W, E or R aka Not Printed
        'weird logic
        If InStr(1, "W|E|R", UCase(strStatus)) > 0 Or bPowerUser Then
            If UCase(strStatus) = "P" Then  'if printed already and we are a power user
                'get the lettername and add it to array of letters to be printed
                cmdGetLetter.Parameters("InstanceID").Value = strInstanceID
                cmdGetLetter.Execute
                'only if the letter has been printed do we need to populate the array to delete files.
                FileArrays(i) = Trim(Nz(cmdGetLetter.Parameters("LetterName").Value, ""))
                i = i + 1
            End If
            
            cmd.Parameters("InstanceID").Value = strInstanceID
            cmd.Execute
            
            iRtnCd = cmd.Parameters("Return").Value
            strErrMsg = Trim(Nz(cmd.Parameters("ErrMsg").Value, ""))
            
            If iRtnCd <> 0 Or strErrMsg <> "" Then
                GoSub Error_Encountered
            End If
        End If
    Next varItem
    cn.CommitTrans 'we went through the entire process.
    
    'now empty out the array
    'On Error Resume Next
    'goes through a few extra times but i am not worried about performance here.  TGH
    'For i = 0 To UBound(FileArrays)
    '    If FileArrays(i) <> "" Then
    '        Kill FileArrays(i)
    '
    '        If Len(dir(Left(FileArrays(i), InStrRev(FileArrays(i), "\")) & "*.*")) = 0 Then
    '            DeleteFolder (Left(FileArrays(i), InStrRev(FileArrays(i), "\") - 1))
    '        End If
    '        End If
    '
    'Next i
    On Error GoTo Error_Encountered
    GoTo Cleanup
Error_Encountered:
        If strErrMsg <> "" Then
            MsgBox strErrMsg
        Else
            MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
        End If
    cn.RollbackTrans
Cleanup:

    Set cmd = Nothing
    Set cmdGetLetter = Nothing
    cn.Close
    Set cn = Nothing
    RefreshMain
End Sub
Private Sub cmdEndDate_Click()
    On Error GoTo Exit_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtThroughDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtThroughDate = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
    End Sub
Private Sub cmdPrintSelectedItems_Click()

    Dim fmrStatus As Form_ScrStatus
    ''TGH Added 9-26-08
    'Dim PreviewViewAllowed As Boolean: PreviewViewAllowed = False
    'Dim PreviewMode As Boolean: PreviewMode = False
    Set fmrStatus = New Form_ScrStatus


    'declare variables
    Dim i As Integer
    Dim TotalRecs As Integer: TotalRecs = 0
    Dim InstSelArray() As String

    Dim varItem As Variant
    Dim Z As Integer
    Dim cn As Variant

    ' ensure we have some selections
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected"
        Exit Sub
    End If


    ' open connection for all printing work to be done so we can rollback if any errors
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = GetConnectString("v_CODE_Database")
    cn.Open
    cn.CursorLocation = adUseClient

    i = 0
    'get the size for the array.  looping through here so we don't have to run a re-dim in a loop (Poor Performance).
    For Each varItem In Me.lstQueue.ItemsSelected
        If UCase(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem)) = "W" Or _
            UCase(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem)) = "R" Then
            i = i + 1 'after loop i will be the correct # count (so one more then the arrary)
        End If
    Next varItem

    If i = 0 Then
        ReDim InstSelArray(0)
    Else
        ReDim InstSelArray(i - 1)
    End If

    TotalRecs = i
    'reset counter
    i = 0
    'populate an array with the instance id's we want to print. this is needed for when we view these
    For Each varItem In Me.lstQueue.ItemsSelected
        If Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "w" Or _
            Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "r" Then
            InstSelArray(i) = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
            i = i + 1 'after loop i will be the correct # count (so one more then the arrary)
        End If
    Next varItem

    'if i was not increased from zero we know we don't have any "W"s in the selected items so we exit as nothing to view.
    If TotalRecs = 0 Then
        MsgBox ("Generate letters is only for letters with a status of 'W' or 'R'")
        'cn.RollbackTrans
        GoTo Error_msg
    End If

    'BEGIN THE PROGRESS FORM, MAKE THE MAX PROGRESS DOUBLE THE ITEMS SELECTED B/C WE HAVE TO TOUCH EACH REC TWICE WHEN GENERATING.
    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgMax = TotalRecs * 2 'lstQueue.ItemsSelected.Count can't use selected amount b/c could select 'P' letters
        .TimerInterval = 50
        .show
    End With

   
    ' Print letter
    If Not PrintLetters(fmrStatus) Then
        GoTo Error_msg
    End If
    
    
    'now refresh the selected and highlight the onese we want to view (saved from above)
    For i = 1 To Me.lstQueue.ListCount
        For Z = 0 To UBound(InstSelArray)
            If Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, i) = InstSelArray(Z) Then
                Me.lstQueue.Selected(i) = True
            End If
        Next Z
    Next i
    
    'view these guys
    If Not ViewLetters(TotalRecs, cn, fmrStatus) Then
        GoTo Error_msg:
    End If

    Set cn = Nothing

    Exit Sub

Error_msg:
    'MsgBox "Error printing all work rolledback!"
End Sub

Private Function PrintSelectedItems(cn As Variant, TotalRecs As Integer, fmrStatus As Form_ScrStatus) As Boolean
    'On Error GoTo Error_encountered
    'set the function to False in the beginning.  This ensures we get to the end to set it as correct.
    PrintSelectedItems = False
    
    Dim PrintFileArray() As String: ReDim Preserve PrintFileArray(TotalRecs - 1)
    Dim UpdateLettersArray() As String: ReDim Preserve UpdateLettersArray(TotalRecs - 1)
'    Dim Person As New ClsIdentity
    ' ADO variables - Late binded here.
    Dim cmd As Variant 'ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    
    Dim cmdGetLetter As Variant 'ADODB.Command
    Set cmdGetLetter = CreateObject("ADODB.Command")
        
    Dim strSQLcmd As String
    
    ' Letter configuration variables
    Dim db As Database
    Dim rsLetterConfig As DAO.RecordSet
    Dim strODCFile As String
    Dim strBasedPath As String
    Dim colLetterTemplate As Collection
    Dim objLetterInfo As New clsLetterTemplate
    
    
    ' Word objects setup as variants b/c of late binding
    Dim objWordApp, _
        objWordDoc, _
        objMasterDoc, _
        objWordMergedDoc
        
    Set objWordApp = CreateObject("word.application")
    objWordApp.visible = False
    
    'Letter generation variables
    Dim rsProvList As Variant 'ADODB.Recordset
    Set rsProvList = CreateObject("ADODB.Recordset")
    
    Dim rsLetterTemplate As Variant 'ADODB.Recordset
    Set rsLetterTemplate = CreateObject("ADODB.Recordset")

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
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    cmd.ActiveConnection = cn 'open connection passed in
    bFirstLetter = True
    
    strErrMsg = ""

    'set local path
    'USER ENTRY NEEDED
    strLocalPath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\" & Identity.UserName & "\LETTERTEMPLATE"
    'End USER ENTRY NEEDED
    CreateFolder (strLocalPath)
    
    ' get list of templates
    Set colLetterTemplate = New Collection
    'TL add account ID logic
    'strSQLCmd = "select LetterType, TemplateLoc from HC_AUDITORS_Claims..LETTER_Type where AccountID = " & gintAccountID
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
            objLetterInfo.LetterType = Trim(![LetterType])
            objLetterInfo.TemplateLoc = Trim(![TemplateLoc])
            colLetterTemplate.Add objLetterInfo, Trim(![LetterType])
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
    For Each varItem In lstQueue.ItemsSelected
    If Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "w" Or _
       Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "r" Then
        strInstanceID = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
        strProvNum = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("cnlyProvID").OrdinalPosition, varItem))
        strLetterType = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterType").OrdinalPosition, varItem))
        dtLetterReqDt = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterReqDt").OrdinalPosition, varItem)
        strAuditor = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Auditor").OrdinalPosition, varItem))
        strStatus = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem))
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
                        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
                        bMergeError = True
                        objWordApp.ActiveDocument.Activate
                        GoTo Cleanup
                    End If
                    ''------------------- here is where we convert to pdf instead of word ----------------''
                    ' Save the output doc
                    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
                    strOutputPath = strBasedPath & "\" & strProvNum & "\"
                    CreateFolder (strOutputPath)
                    'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    'Added to rename reprints...
                    If strStatus = "R" Then
                        strOutputFileName = strOutputPath & "" & strLetterType & "-Reprint-" & strInstanceID & ".doc"
                    Else
                        strOutputFileName = strOutputPath & "" & strLetterType & "-" & strInstanceID & ".doc"
                    End If
                    objWordMergedDoc.spellingchecked = True
                    Sleep 2000
                    objWordMergedDoc.SaveAs strOutputFileName
                    objWordMergedDoc.Close
                    
                    'save the file location we just generated in case user cancels or error
                    PrintFileArray(iCnt - 1) = strOutputFileName
                    

                    Set objWordMergedDoc = Nothing

                    'save the word doc's in original form.  could convert these to pdf if we are so inclined...
                    'ConvertWordToPDF Replace(strOutputFileName, ".doc", ".PDF"), strOutputFileName
'            Else
'                'we are working with an access report now.
'                'load report info into temp table...
'                'Set the output path, and create the folder if it does not exist
'                'then we call ConvertRPTtoPDF to save the access report as a pdf image.
'                AdoExeTxt "usp_LETTER_Get_Info_load '" & strInstanceID & "',''", "v_AMERIHEALTH_Auditors_Code"
'                strOutputPath = strBasedPath & "\" & strProvNum
'                CreateFolder (strOutputPath)
'                'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".PDF"
'                If strStatus = "r" Then
'                     strOutputFileName = strOutputPath & "\" & strLetterType & "-Reprint-" & strInstanceID & ".PDF"""
'                Else
'                     strOutputFileName = strOutputPath & "\" & strLetterType & "-" & strInstanceID & ".PDF"
'                End If
'
'                If Not ConvertRPTToPDF(strOutputFileName, objLetterInfo.TemplateLoc) Then
'                    strErrMsg = "pdf conversion took too long to run, most likely a printing spooling issue, please try again"
'                    GoTo Error_encountered
'                End If
'
'                'save the file location we just generated in case user cancels or error
'                PrintFileArray(Icnt - 1) = strOutputFileName
'
'                'DoCmd.OutputTo acOutputReport, objLetterInfo.TemplateLoc, acFormatRTF, strOutputFileName, False
            End If
                     
        ' update status
        
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
            GoSub Error_Encountered
        End If
        
        
        
        
        
        'temp cause error
        'GoTo Error_encountered:
    
    ' Notify user we are done.
    cboViewType.SetFocus
    cboViewType.ListIndex = 1
    cn.CommitTrans
    cmdRefresh_Click 'can't run with open trans
    PrintSelectedItems = True
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
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
    Dim cmd As Variant 'ADODB.Command
        
    Dim strSQLcmd As String
    Dim intI As Integer
            
    'Zero based array
    For intI = 0 To UBound(UpdateLettersArray)
        Set cmd = CreateObject("ADODB.Command")
        cmd.ActiveConnection = cn
        cmd.commandType = adCmdStoredProc
        RetCd = 1   'set return to false
        cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
        cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
        cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adVarChar, adParamOutput, 255, "")
                
        cmd.Parameters("InstanceID").Value = UpdateLettersArray(intI)
        'USER ENTRY NEEDED '******
        'usp_LETTER_UpdateAfterGenerated needs to be populated and called below...
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

'Comment back in for use with Claims Plus
'Private Sub CreatePK(ByVal Tablename As String, ByVal Fields As String)
'On Error Resume Next
'    CurrentDb.Execute "CREATE UNIQUE INDEX PK_" & Tablename & " On " & Tablename & "(" & Fields & ")"
'End Sub
Private Sub cmdRefresh_Click()
    RefreshMain
End Sub

Private Sub cmdReprint_Click()
    Dim strLetterDate As String
    
    
    'While IsDate(strLetterDate) = False
    '    strLetterDate = InputBox("Enter a Reprint Date", "Reprint Date", Format(Now, "mm/dd/yyyy"))
    'Wend
    
    
    On Error GoTo Error_Encountered
    
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There is no item selected", vbOKOnly, "No Item Selected"
        Exit Sub
    End If
'    Dim Person As New ClsIdentity
    
    ' ADO variables
    Dim cn As Variant 'ADODB.Connection
    ' open connection
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = GetConnectString("v_CODE_Database")
    cn.Open
    cn.CursorLocation = adUseClient
    cn.BeginTrans
        
    Dim cmd As Variant 'ADODB.Command
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = cn
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmd.Parameters.Append cmd.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("Auditor", adChar, adParamInput, 75, Identity.UserName)
    cmd.Parameters.Append cmd.CreateParameter("LtrDate", adChar, adParamInput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("NewInstanceID", adChar, adParamOutput, 20, "")
    cmd.Parameters.Append cmd.CreateParameter("ErrMsg", adChar, adParamOutput, 255, "")
    cmd.CommandText = "usp_LETTER_Reprint"
        
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
    Set fmrStatus = New Form_ScrStatus
    Dim iCnt As Integer: iCnt = 1 'count through loop


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
    
    For Each varItem In lstQueue.ItemsSelected
    
        strInstanceID = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem)
        strStatus = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem)
        curLtr = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Lettertype").OrdinalPosition, varItem)
        ProvNum = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("cnlyProvID").OrdinalPosition, varItem)
        
        'For P status letters only
        If UCase(strStatus) = "P" Then
            'if this is the first time set current letter and last letter = and prompt for date.  USER Request 4-17-08 TGH
            If Reprinted = False Then
                LastLtr = curLtr
                cmd.Parameters("LtrDate").Value = InputBox("What Date would you like on the reprint for letter " & curLtr, "Reprint Date", Format(Now(), "mm/dd/yyyy"))
            Else
                'else this isn't the first time in this loop.  check the last letter and see if it differs if so re-ask for the date for the reprint
                If curLtr <> LastLtr Then
                    LastLtr = curLtr
                    cmd.Parameters("LtrDate").Value = InputBox("What Date would you like on the reprint for letter " & curLtr, "Reprint Date", Format(Now(), "mm/dd/yyyy"))
                End If
                
            End If
            
            'flag to let us know we have some that will be reprinted (Marked as "P")
            Reprinted = True
    
            cmd.Parameters("InstanceID").Value = strInstanceID
            cmd.Execute
            
            iRtnCd = cmd.Parameters("Return").Value
            strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
            
            If iRtnCd <> 0 Or strErrMsg <> "" Then
                GoSub Error_Encountered
            End If
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
        DoEvents
        DoEvents
        DoEvents
        

    Next varItem

    'when done all loops and no errors commit the transaction
    cn.CommitTrans
    
    If Not Reprinted Then
        iAnswer = MsgBox("Please note that only records with status ""P""  can be reprint", vbInformation)
    Else
        lstQueue.RowSource = "SELECT wq.* FROM LETTER_Work_Queue wq INNER JOIN LETTER_Reprint_Queue rq " & _
                             " ON wq.InstanceID = rq.InstanceID " & _
                             " WHERE wq.Status = ""R"" "
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
    
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    
    'if any error rollback
    On Error Resume Next
    cn.RollbackTrans
    
Cleanup:
    Set cmd = Nothing
    cn.Close
    Set cn = Nothing
    
End Sub

Private Sub cmdSelectEntireQueue_Click()
    Dim idx As Integer

    For idx = 1 To Me.lstQueue.ListCount
        Me.lstQueue.Selected(idx) = True
    Next idx
    ' could add in detail view but left off b/c assuming analyst are just selecting all to print in batch processes - TH 10-8-07
End Sub

Private Sub cmdStartDate_Click()
    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtFromDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.txtFromDate = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub


Public Sub RefreshMain()
    On Error GoTo Error_Encountered
    
    Dim dtThrouDt As Date
    Dim strSelectedAuditor As String
    'If txtThroughDate.Value <> Null Then
        dtThrouDt = DateAdd("d", 1, txtThroughDate.Value)
    'Else
        'Exit Sub
    'End If
    
'* JC Simplified.  The original section was requerying 3 times.

    Dim sQueueRowSource As String
    
    sQueueRowSource = "SELECT * FROM LETTER_Work_Queue WHERE RowCreateDt >= #" & Nz(txtFromDate.Value, "01/01/1900") & "# and RowCreateDt < #" & _
                                Format(dtThrouDt, "mm-dd-yyyy") & "# "
    Select Case cboViewType
    
        Case "View Un-Processed Letters"
        
            sQueueRowSource = sQueueRowSource & " AND Status in ('W','R')"
            
        Case "View Processed Letters"
        
           sQueueRowSource = sQueueRowSource & " AND Status = 'P'"
                            
        Case "View Errors"
        
            sQueueRowSource = sQueueRowSource & " AND Status = 'E'"
            
        Case Else
            '* Keep Original String
    End Select
    
    strSelectedAuditor = cboAuditor.Value
    
    If strSelectedAuditor <> "View All" And strSelectedAuditor <> "" Then
        sQueueRowSource = sQueueRowSource & " and Auditor = " & Chr(34) & strSelectedAuditor & Chr(34)
    End If
    
    If Me.tglAdvancedFilter = True Then
        sQueueRowSource = sQueueRowSource & "AND (" & msAdvancedFilter & ")"
    End If
    
    Me.lstQueue.RowSource = sQueueRowSource & " order by lettertype, cnlyProvID, letterreqdt"
        
'*     lstQueue.requery JC Don't need this


    'Auto sizing
    Set ColResize1 = New clsAutoSizeColumns
    ColResize1.SetControl Me.lstQueue
    'don't resize if lstclaims is null
    If Me.lstQueue.ListCount > 1 Then
        ColResize1.AutoSize
    End If

Exit Sub

Error_Encountered:
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    'LogMessage TypeName(Me) & "CmdViewLetters_Click-2010", "USAGE DETAIL", "SOmeone is using this form!!!"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    If IsSubForm(Me) Then
        lblAppTitle.visible = False
    Else
        lblAppTitle.visible = True
    End If
    
    Me.tglAdvancedFilter = False
    Me.tglAdvancedFilter.Caption = "Add Filter"
    
    RefreshMain
    
End Sub

Private Sub lstQueue_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        cmdDeleteQueue_Click
    Else
        'lstQueue_Click
    End If
    
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
    End If
End Sub
Private Sub txtThroughDate_Exit(Cancel As Integer)
    If Not IsDate(txtThroughDate.Value) Then
        MsgBox "Please enter a valid through date"
        txtThroughDate.SetFocus
    End If
End Sub

Private Function JoinPDFs(TargetPDF As String, FileToAppend As String) As Boolean 'This one works!
    Dim Project_Folder As String
    Dim fs As Variant 'FileSystemObject - late bind this
    Dim AcroApp As Object
    Dim PDF_TargetFile As Object
    Dim PDF_DataSheet As Object
    Dim RowNr As Integer
    Dim iPathLen As Integer
    Dim j As Integer
    Dim i As Integer
    Dim strChkPath As String
    
    Project_Folder = left(TargetPDF, InStrRev(TargetPDF, "\") - 1)
    Set fs = CreateObject("scripting.filesystemobject")
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
       
End Function
Private Sub CmdViewLetters_Click()
' open connection for all Viewing work to be done so we can rollback if any errors
' View Letter Button is written to open a new explorer window contaning the combined letters selected.  You can only view generated letters.  NOTE if you leave the folder open you may have to refresh to see the file

On Error GoTo Err:

    Dim cn As Variant
    Dim varItem As Variant
    Dim TotalPrintedRecs As Integer: TotalPrintedRecs = 0
    Dim TotalPreviewRecs As Integer: TotalPreviewRecs = 0 'TGH added 9-26-08
    Dim PreviewViewAllowed As Boolean ': PreviewViewAllowed = True 'TGH added 9-26-08
    
    Dim db As Database 'assign the AllowPreview from the config table
        Set db = CurrentDb
    
    Dim rsLetterConfig As DAO.RecordSet
    
    'TL add account ID logic
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    If rsLetterConfig("AllowPreview").Value = "TRUE" Or rsLetterConfig("AllowPreview").Value = "YES" Then
        PreviewViewAllowed = True
    Else
        PreviewViewAllowed = False
    End If
    
    Set rsLetterConfig = Nothing
    Set db = Nothing
    
  
    'make sure we have some items selected.
    If lstQueue.ItemsSelected.Count = 0 Then
        MsgBox "There are no items selected"
        Exit Sub
    End If
    
    'get the count of Printed leters.  Because the auditor can select all letters in the grid they could pick a mix of unprocessed and processed.
    'we handle the 'P'  and 'R' Lletters differently - we pull the actual letter generated
    ' all others we create a temp preview file
    
    For Each varItem In lstQueue.ItemsSelected
        If UCase(Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem))) = "P" Then
            TotalPrintedRecs = TotalPrintedRecs + 1
            Else 'TGH Added 9-26-08 Count non P records.
                TotalPreviewRecs = TotalPreviewRecs + 1
        End If
    Next varItem
    
    'run QC Error check to see if no Printed recs selected and preview is not allowed:
        If Not PreviewViewAllowed And TotalPrintedRecs = 0 Then
            MsgBox ("Selected items need to be printed to be viewed")
            GoTo Err:
        ElseIf PreviewViewAllowed And (TotalPrintedRecs + TotalPreviewRecs = 0) Then
            MsgBox ("Please Select Letters to View")
             GoTo Err:
        End If
                
    
    ' setup progress screen
    Dim fmrStatus As Form_ScrStatus
    Set fmrStatus = New Form_ScrStatus

    With fmrStatus
        .ShowCancel = True
        .ShowMessage = False
        .ShowMessage = True
        .ProgVal = 0
        If PreviewViewAllowed Then 'TGH added 9-26-08
            .ProgMax = TotalPrintedRecs + TotalPreviewRecs
        Else
            .ProgMax = TotalPrintedRecs ' Using  b/c we only want to view the printed letters.
        End If
        .TimerInterval = 50
        .show
    End With
    
    'open our connection
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = GetConnectString("v_CODE_Database")
    cn.Open
    cn.CursorLocation = adUseClient
    cn.BeginTrans
    
    
    'TGH UPdated this has been added in for previewing letters.
    If TotalPreviewRecs > 0 And PreviewViewAllowed Then
        If Not PreviewViewLetters(cn, TotalPreviewRecs, fmrStatus) Then
            cn.RollbackTrans
            'MsgBox ("error running view letters")
        End If
    End If
    
        
    If TotalPrintedRecs > 0 Then
        If ViewLetters(TotalPrintedRecs, cn, fmrStatus) Then
            cn.RollbackTrans
            '   MsgBox ("error running view letters")
'        Else
'            cn.CommitTrans
        End If
   End If
    
    cn.CommitTrans
    cn.Close
    Set cn = Nothing
    
    Exit Sub
Err:
     On Error Resume Next   'in case cn has not been openedcn.Rollback
     cn.Close
    Set cn = Nothing
    
End Sub


' NOTE: all letter should have been generated before entering this function
Private Function ViewLetters(TotalRecs As Integer, cn As Variant, fmrStatus As Form_ScrStatus) As Boolean
    
    Dim strErrMsg As String: strErrMsg = ""

    ViewLetters = False
    
    On Error GoTo Error_Encountered:
    
    
    'make sure we have some items selected.
    Dim iTotalSelected As Integer
    
    iTotalSelected = lstQueue.ItemsSelected.Count
    If iTotalSelected = 0 Then
        MsgBox "There is no item selected"
        ViewLetters = False
        Exit Function
    End If
    
    
    
    
    ' progress bar variables
    Dim ProgVal As Integer
    Dim strProgressMsg As String
    Dim lngProgressCount As Long
    Dim msgIcon As Integer
    
    
    Dim fso As New FileSystemObject
    

    ' get letter output path from letter config table
    Dim db As Database
    Dim rsLetterConfig As DAO.RecordSet
    Dim strOutputPath As String
'    Dim Person As New ClsIdentity
    
    Set db = CurrentDb
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    If rsLetterConfig.recordCount = 1 Then
        strOutputPath = rsLetterConfig("LetterOutputLocation").Value & "\PREVIEW\" & Replace(Identity.UserName, ".", "") & "\"
    Else
        strErrMsg = "ERROR: Letter configuration data is missing or not setup correctly."
        GoTo Error_Encountered
    End If
    
    
    ' Setup output directory
    Dim bRetCd As Boolean
    
    DeleteFolder (strOutputPath)       ' remove dir to make sure previous entries are deleted
    
    bRetCd = CreateFolder(strOutputPath)   ' create the output directory
    If bRetCd = False Then
        strErrMsg = "ERROR: Can not remove preview directory " & strOutputPath
        GoTo Error_Encountered
    End If
    
    
    
    ' get user options
    Dim iCombinedDoc As Integer
    Dim iPrintDoubleSided As Integer
    Dim iMaxPagesPerFile As Integer
    
    iCombinedDoc = MsgBox("Combined multiple documents into one doc?", vbYesNo)
    If iCombinedDoc = vbYes Then
        iMaxPagesPerFile = InputBox("Please enter number of pages per file (500 pages max): ", vbYesNo & vbInformation)
        If iMaxPagesPerFile = 0 Then iMaxPagesPerFile = 250      ' default to 250 pages per output file if not specified
        If iMaxPagesPerFile > 500 Then iMaxPagesPerFile = 500
    
        iPrintDoubleSided = MsgBox("Print double sided?", vbYesNo)
    End If
    
    
    

    ' setup command object to read letter name
    Dim cmdGetLetter As New ADODB.Command
    
    cmdGetLetter.ActiveConnection = cn      ' connection is passed in to improve performance
    cmdGetLetter.commandType = adCmdStoredProc
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("Return", adInteger, adParamReturnValue, 1)
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("InstanceID", adChar, adParamInput, 20, "")
    cmdGetLetter.Parameters.Append cmdGetLetter.CreateParameter("LetterName", adChar, adParamOutput, 255, "")
    cmdGetLetter.CommandText = "usp_LETTER_Get_Letter_Name"
    
    
    
    ' set up word for processing
'    Dim objWordApp As Word.Application
'    Dim objMasterDoc As Word.Document
'    Dim objWordDoc As Word.Document
    
    Dim objWordApp As Object
    Dim objMasterDoc As Object
    Dim objWordDoc As Object
    
    
    Set objWordApp = CreateObject("word.application")
    objWordApp.visible = False
    
    
    ' BEGIN PROCESSING
    Dim varItem
    Dim iCnt As Integer
    Dim iNumberOfPages As Integer
    Dim iTotalPages As Integer
    Dim strInstanceID As String
    Dim strStatus As String
    Dim strOutputFileName As String
    Dim strLetterFileName As String
    Dim strLetterType As String
    Dim strPreviousLetterType As String
    Dim strPreviewFileName As String
    Dim bFirstFile As Boolean
    Dim bNewFile As Boolean
    
    bFirstFile = True
    bNewFile = True
    For Each varItem In lstQueue.ItemsSelected
        strStatus = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem))
        
        If UCase(strStatus) = "P" Then
            strInstanceID = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
            strLetterType = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterType").OrdinalPosition, varItem))
            
            cmdGetLetter.Parameters("InstanceID").Value = strInstanceID
            cmdGetLetter.Execute
            strLetterFileName = Trim(cmdGetLetter.Parameters("LetterName").Value)
            
            If fso.FileExists(strLetterFileName) Then
                If iCombinedDoc = vbYes Then
                    If bNewFile Then                    ' new output file
                        If Not bFirstFile Then          ' save current output file if not the first file
                            Sleep 2000
                            objMasterDoc.SaveAs strPreviewFileName
                        End If
                    
                        ' set output file name
                        strPreviewFileName = strOutputPath & strLetterType & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                        Set objMasterDoc = objWordApp.Documents.Open(strLetterFileName)
                        objWordApp.selection.EndKey Unit:=wdStory
                        objWordApp.selection.InsertBreak Type:=wdPageBreak
                        objWordApp.selection.InsertFile (strLetterFileName)
                        Sleep 2000
                        objMasterDoc.SaveAs strPreviewFileName
    
                        ' insert extra page for printing double sided
                        If iPrintDoubleSided = vbYes Then
                            iNumberOfPages = objWordApp.selection.Information(wdNumberOfPagesInDocument)
                            
                            If (iNumberOfPages Mod 2) = 1 Then
                                objWordApp.selection.EndKey Unit:=wdStory
                                objWordApp.selection.InsertBreak Type:=wdPageBreak
                                objWordApp.selection.InsertFile ("m:\temp\blankpage.docx")
                            End If
                        End If
                    
                        bFirstFile = False
                        bNewFile = False
                    Else
                    End If
                Else
                    Call fso.CopyFile(strLetterFileName, strOutputPath)
                End If
            Else
                MsgBox "Error: file " & strLetterFileName & " does not exists.", vbExclamation
            End If
        End If
        iCnt = iCnt + 1
        
        ' display progress
        strProgressMsg = "View Record " & iCnt & " / " & iTotalSelected
        fmrStatus.ProgVal = ProgVal + iCnt
        fmrStatus.StatusMessage strProgressMsg

        If fmrStatus.ProgMax = lngProgressCount Then
            msgIcon = vbInformation
        Else
            msgIcon = vbExclamation
        End If

        'Check if the form's status has been selected as cancel.  if  so rollback and promt with error message.
        If fmrStatus.EvalStatus(2) = True Then
                strProgressMsg = "Viewing Generated Letters Canceled!"
                fmrStatus.StatusMessage strProgressMsg
                DoEvents
                strErrMsg = strProgressMsg
                GoTo Error_Encountered
        End If
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Next varItem
    
    

    
    'this is the first time through the view letter.
    On Error GoTo Error_Encountered
        
    'set to the formstatus current progress
    iCnt = 0
    ProgVal = fmrStatus.ProgVal

    Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
    
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    ViewLetters = False

Cleanup:
    'Do not close the Connection passed in, taking care of that in the calling sub.
    Set cmdGetLetter = Nothing

    Set objWordApp = Nothing

End Function


Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mReturnDate = ReturnDate
End Sub


Private Function PreviewViewLetters(cn As Variant, TotalRecs As Integer, fmrStatus As Form_ScrStatus) As Boolean
    'On Error GoTo Error_encountered
    'set the function to False in the beginning.  This ensures we get to the end to set it as correct.
    PreviewViewLetters = False
    
    Dim PreViewFileArray() As String: ReDim Preserve PreViewFileArray(TotalRecs - 1)
'    Dim Person As New ClsIdentity
    ' ADO variables - Late binded here.
    Dim cmd As Variant  'ADODB.Command
        Set cmd = CreateObject("ADODB.Command")
    Dim cmdGetLetter As Variant 'ADODB.Command
        Set cmdGetLetter = CreateObject("ADODB.Command")
        
    Dim strSQLcmd As String
    
    ' Letter configuration variables
    Dim db As Database
    Dim rsLetterConfig As DAO.RecordSet
    Dim strODCFile As String
    Dim strBasedPath As String
    Dim colLetterTemplate As Collection
    Dim objLetterInfo As New clsLetterTemplate
    
    ' Word objects setup as variants b/c of late binding
    Dim objWordApp, _
        objWordDoc, _
        objMasterDoc, _
        objWordMergedDoc
        
    Set objWordApp = CreateObject("word.application")
    objWordApp.visible = False

    
    'Letter generation variables
    Dim rsProvList As Variant 'ADODB.Recordset
        Set rsProvList = CreateObject("ADODB.Recordset")
    Dim rsLetterTemplate As Variant 'ADODB.Recordset
        Set rsLetterTemplate = CreateObject("ADODB.Recordset")

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
    Dim cnLocal As ADODB.Connection
    
    bFirstLetter = True
    
    strErrMsg = ""
    

    
    'set local path
    'USER ENTRY NEEDED
    strLocalPath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\" & Identity.UserName & "\LETTERTEMPLATE"
    'End USER ENTRY NEEDED
    CreateFolder (strLocalPath)
    
    ' get list of templates
    Set colLetterTemplate = New Collection
    'TL add account ID logic
    
    Set cnLocal = CreateObject("ADODB.Connection")
    cnLocal.ConnectionString = GetConnectString("v_DATA_Database")
    cnLocal.Open
    cnLocal.CursorLocation = adUseClient
    cnLocal.BeginTrans
    
    strSQLcmd = "select LetterType, TemplateLoc from LETTER_Type where AccountID = " & gintAccountID
    cmd.ActiveConnection = cnLocal 'open connection passed in
    cmd.CommandText = strSQLcmd
    cmd.CommandTimeout = 0
    cmd.commandType = adCmdText
    Set rsLetterTemplate = cmd.Execute

    'rst is a recset of all templates and locations.
    'Then populates the collection colLetterTemplate with the info from these templates.
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            objLetterInfo.LetterType = Trim(![LetterType])
            objLetterInfo.TemplateLoc = Trim(![TemplateLoc])
            colLetterTemplate.Add objLetterInfo, Trim(![LetterType])
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
    'select our preview folder and delete it if it exists...
    strOutputPath = strBasedPath & "\PREVIEW"
    strAuditor = Replace(Identity.UserName, ".", "")
    strOutputPath = strOutputPath & "\" & strAuditor
    DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
    CreateFolder (strOutputPath & "\")
    
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
    For Each varItem In lstQueue.ItemsSelected 'ensure we don't re-create printed ones.
    If Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "w" Or _
       Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem) = "r" Then
        strInstanceID = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
        strProvNum = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("cnlyProvID").OrdinalPosition, varItem))
        strLetterType = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterType").OrdinalPosition, varItem))
        dtLetterReqDt = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterReqDt").OrdinalPosition, varItem)
        strAuditor = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Auditor").OrdinalPosition, varItem))
        strStatus = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem))
        iCnt = iCnt + 1

            If strLetterType <> objLetterInfo.LetterType Then
                'We are dealing with a new report in the queue!
                Set objLetterInfo = colLetterTemplate(strLetterType)
                 
                ObjectExists = False
                'ACCESS REPORT LOGIC
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
                    'WORD DOC AREA
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
                    
                    
                    'add a connolly-internal watermark to the preview letters
                    If Not (ADDWATERMARK(objWordApp, objWordDoc, strErrMsg)) Then
                        GoTo Error_Encountered
                    End If
                    
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
                        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
                        bMergeError = True
                        objWordApp.ActiveDocument.Activate
                        GoTo Cleanup
                    End If
                    ''------------------- here is where we convert to pdf instead of word ----------------''
                    ' Save the output doc
                    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
                    'strOutputPath = strBasedPath & "\PREVIEW"
                    
                       'strAuditor = Replace(Person.UserName, ".", "")              ' thieu 1/16/08
                        'strOutputPath = strOutputPath & "\" & strAuditor    'thieu 1/16/08
                        'DeleteFolder (strOutputPath)                                'clear out the folder if it exists...
                        'CreateFolder (strOutputPath & "\")                                ' thieu 1/16/08

           
                   ' CreateFolder (strOutputPath)
                    'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".doc"
                    'Added to rename reprints...
                    strOutputFileName = strOutputPath & "\" & strLetterType & "-Preview-" & strInstanceID & ".doc"
                    objWordMergedDoc.spellingchecked = True
                    Sleep 2000
                    objWordMergedDoc.SaveAs strOutputFileName
                    objWordMergedDoc.Close
                    
                    'save the file location we just generated in case user cancels or error
                    PreViewFileArray(iCnt - 1) = strOutputFileName
                    

                    Set objWordMergedDoc = Nothing

                    'save the word doc's in original form.  could convert these to pdf if we are so inclined...
                    'ConvertWordToPDF Replace(strOutputFileName, ".doc", ".PDF"), strOutputFileName
'            Else
'                'we are working with an access report now.
'                'load report info into temp table...
'                'Set the output path, and create the folder if it does not exist
'                'then we call ConvertRPTtoPDF to save the access report as a pdf image.
'                AdoExeTxt "usp_LETTER_Get_Info_load '" & strInstanceID & "',''", "v_AMERIHEALTH_Auditors_Code"
'                strOutputPath = strBasedPath & "\" & strProvNum
'                CreateFolder (strOutputPath)
'                'strOutputFileName = strOutputPath & "\" & strLetterType & "-" & Format(dtLetterReqDt, "yyyymmdd") & "-" & Format(Now, "yyyymmddhhmmss") & ".PDF"
'                If strStatus = "r" Then
'                     strOutputFileName = strOutputPath & "\" & strLetterType & "-Reprint-" & strInstanceID & ".PDF"""
'                Else
'                     strOutputFileName = strOutputPath & "\" & strLetterType & "-" & strInstanceID & ".PDF"
'                End If
'
'                If Not ConvertRPTToPDF(strOutputFileName, objLetterInfo.TemplateLoc) Then
'                    strErrMsg = "pdf conversion took too long to run, most likely a printing spooling issue, please try again"
'                    GoTo Error_encountered
'                End If
'
'                'save the file location we just generated in case user cancels or error
'                PreViewFileArray(Icnt - 1) = strOutputFileName
'
'                'DoCmd.OutputTo acOutputReport, objLetterInfo.TemplateLoc, acFormatRTF, strOutputFileName, False
            End If
                     
        ' update status
        'cmd.Parameters("InstanceID").Value = strInstanceID
        'cmd.Parameters("LetterName").Value = strOutputFileName
        'strSQLCmd = "usp_LETTER_Update_Status"
        'cmd.CommandText = strSQLCmd
        'cmd.commandType = adCmdStoredProc
        'cmd.Execute
            
        strErrMsg = Trim(cmd.Parameters("ErrMsg").Value)
        If strErrMsg <> "" Then
            GoTo Error_Encountered
        End If
        
        'this Function Call is client specific for after the letters are generated.
        'If Not UpdateAfterLetterGenerated(cn, strInstanceID, strErrMsg) Then
        '   GoSub Error_encountered
        'End If
        

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
    
    cn.CommitTrans
    PreviewViewLetters = True

    i = MsgBox("Would you like to combine the Preview Letters into one file?", vbYesNo)
        If i = vbYes Then
            If (Not CombineDocs(strOutputPath)) Then 'could pass stroutputpath, true to delete all but he combined here.  clients choice.
                GoTo Error_Encountered
            End If
       End If
    
Shell "explorer.exe " & Chr$(34) & strOutputPath & Chr$(34), vbNormalFocus
          
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    PreviewViewLetters = False
    
    'should track all documents created and delete the ones when we error'd out or cancelled generation.
     'now empty out the array
'    On Error Resume Next
'    For i = 0 To UBound(PreViewFileArray)
'        If PreViewFileArray(i) <> "" Then
'        Kill PreViewFileArray(i)
'        'only deletes folder if it is empty
'            If Len(dir(Left(PreViewFileArray(i), InStrRev(PreViewFileArray(i), "\")) & "*.*")) = 0 Then
'                DeleteFolder (Left(PreViewFileArray(i), InStrRev(PreViewFileArray(i), "\") - 1))
'            End If
'        End If
'
'    Next i
    

      
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
End Function

Private Function ADDWATERMARK(objWordApp As Variant, objWordDoc As Variant, ByRef strErrMsg As String) As Boolean
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

   'Dim strpath As String
   Dim First As Boolean: First = True
   
   'Dim strPreviewFileName As String
   'strpath = "\\ccaintranet.com\DFS-FLD-01\Audits\AmeriHealth\LETTER_REPOSITORY\LETTERS\PREVIEW\tomhartey"
   
   'the default would be false but doing this in case it is ever passed as null
   KillNonCombined = Nz(KillNonCombined, False)
 
   
   If Nz(strPreviewFileName, "") = "" Then
    strPreviewFileName = strPath & "\CombinedDocs.DOC"
   End If
   
   Dim fso
   Dim CurrentFolder
   Dim Files, file
   Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Late bind the Word Object library 11
    Dim objWordApp, _
           objMasterDoc, _
           objWordDoc

    
    Set objWordApp = CreateObject("word.application")
        ' Create an instance of objWord, and make it invisible.
    objWordApp.visible = False
    
    Set objMasterDoc = objWordApp.Documents.Add()   'tgh 3/18/08
                    objMasterDoc.spellingchecked = True 'this is needed to shut down popup for too many spelling errors
                    Sleep 2000
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
        objWordApp.selection.InsertBreak    'insert a page break befor the file so next insert does not Bunch together. doing it before so you don't get an empty last page
    End If
    
    objWordApp.ActiveDocument.spellingchecked = True
    objWordApp.selection.InsertFile (file.Path)
    If KillNonCombined Then 'added if they only want the combined doc left...
        Kill file.Path
    End If
    Sleep 2000
    objWordApp.ActiveDocument.SaveAs strPreviewFileName
    
    End If
Next
    Sleep 2000
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
    MsgBox Err.Description

   'Destruct our footprint.
   Set objWordDoc = Nothing
   On Error Resume Next
       objWordApp.Quit (0)
   
   Set objWordApp = Nothing
   Set fso = Nothing

    CombineDocs = False
End Function

Private Function PrintLetters(fmrStatus As Form_ScrStatus) As Boolean
    Dim colLetterTemplate As New Collection
    Dim objLetterTemplate As clsLetterTemplate
    
    Dim iCnt As Integer
    Dim bRetCd As Boolean
    Dim bTemplateFound As Boolean
    
    Dim varItem
    Dim varLetterTemplate
    Dim strInstanceStatus As String
    Dim strInstanceID As String
    Dim strProvNum As String
    Dim strLetterType As String
    Dim strOutputFileName As String
    Dim strErrMsg As String
    
    PrintLetters = False
    
    ' Letter configuration variables
    Dim db As Database
    Dim rsLetterConfig As DAO.RecordSet
    Dim strODCFile As String
    Dim strOutputLocation As String
    
    Set db = CurrentDb
    Set rsLetterConfig = db.OpenRecordSet("select * from LETTER_Config where AccountID = " & gintAccountID)
    If rsLetterConfig.recordCount = 0 Then
        strErrMsg = "ERROR: Letter configuration parameters is missing"
        GoTo Error_Encountered
    ElseIf rsLetterConfig.recordCount > 1 Then
        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
        GoTo Error_Encountered
    Else
        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
    End If
    
    ' setup progress screen that is passed to this function
    Dim strProgressMsg As String
    Dim lngProgressCount As Long
    Dim msgIcon As Integer
    Dim ObjectExists As Boolean
    

    ' delete all templates in temp directory
    'DeleteTemplates


    ' start processing
    For Each varItem In lstQueue.ItemsSelected
        strInstanceStatus = Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("Status").OrdinalPosition, varItem)
        If strInstanceStatus = "W" Or strInstanceStatus = "R" Then
            strInstanceID = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("InstanceID").OrdinalPosition, varItem))
            strProvNum = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("cnlyProvID").OrdinalPosition, varItem))
            strLetterType = Trim(Me.lstQueue.Column(Me.lstQueue.RecordSet.Fields("LetterType").OrdinalPosition, varItem))
            
            bTemplateFound = False
            
            
            For Each varLetterTemplate In colLetterTemplate
                Set objLetterTemplate = varLetterTemplate
                If objLetterTemplate.LetterType = strLetterType Then
                    bTemplateFound = True
                    Exit For
                End If
                
                'Debug.Print objLetterTemplate.LetterType
            Next varLetterTemplate
            
            If bTemplateFound = False Then
                bRetCd = CopyTemplates(colLetterTemplate, strLetterType)
                If bRetCd = False Then
                    GoTo Error_Encountered
                End If
            End If
            
            Set objLetterTemplate = colLetterTemplate(strLetterType)
                
                
            
            bRetCd = PrintLetterInstance(strInstanceID, strInstanceStatus, objLetterTemplate.TemplateLoc, _
                                            strOutputFileName, strOutputLocation, strProvNum, _
                                            strODCFile, strLetterType)
            
            iCnt = iCnt + 1
        
        
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
        
            'ADDED TGH 4-16-08 This is to check if the form's status has been selected as cancel.
            If fmrStatus.EvalStatus(2) = True Then
                strProgressMsg = "Cancel has been selected. No records Generated!" ' at " & i & " / " & fmrStatus.ProgMax
                fmrStatus.StatusMessage strProgressMsg
                DoEvents
                strErrMsg = strProgressMsg
            End If
        End If
        
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
                
    Next varItem
    
    ' Notify user we are done.
    cboViewType.SetFocus
    cboViewType.ListIndex = 1
    cmdRefresh_Click 'can't run with open trans
    
    PrintLetters = True
    GoTo Clean_Up
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If


Clean_Up:
    Set colLetterTemplate = Nothing
    Set objLetterTemplate = Nothing
    Set db = Nothing
    Set rsLetterConfig = Nothing
    
End Function

Private Function PrintLetterInstance(pstrInstanceID As String, pstrInstanceStatus As String, pstrTemplateName As String, _
                                        pstrOutputFileName As String, pstrOutputBasePath As String, pstrProvNum As String, _
                                        pstrODCFile As String, pstrLetterType As String) As Boolean
    ' set ADO class
    Dim myCode_ADO As New clsADO
    
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    'set the function to false
    PrintLetterInstance = False
    
    
    ' ADO variables - Late binded here.
    Dim cmd As ADODB.Command
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
    
    Dim objLetterInfo As New clsLetterTemplate
    
    
    
    ' Word objects setup as variants b/c of late binding
    Dim objWordApp, _
        objWordDoc, _
        objWordMergedDoc
        
    Set objWordApp = CreateObject("Word.Application")
    objWordApp.visible = False
    
    
    ' check if template exists
    strChkFile = Dir(pstrTemplateName)
    If strChkFile = "" Then
        strErrMsg = "Missing letter template." & vbCrLf & "Template name = " & objLetterInfo.TemplateLoc & ""
        GoTo Error_Encountered
    End If

    
    ' open template
    Set objWordDoc = objWordApp.Documents.Add(pstrTemplateName, , False)

       
       
    ' load letter info
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_LETTER_Get_Info_load"
    cmd.Parameters.Refresh
    cmd.Parameters("@InstanceID") = pstrInstanceID
    cmd.Execute
    
    strErrMsg = Trim(cmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
        GoTo Error_Encountered
    End If
    
    
    ' Set data source for mail merge.  Data will be from new Temp Table
    objWordDoc.MailMerge.OpenDataSource Name:=pstrODCFile, _
                        SqlStatement:="exec usp_LETTER_Get_Info '" & pstrInstanceID & "'"
                    
    
    ' Perform mail merge.
    objWordDoc.MailMerge.MainDocumentType = 3 'wdDirectory
    objWordDoc.MailMerge.Destination = 0 'wdSendToNewDocument
    objWordDoc.MailMerge.Execute Pause:=False
    If left(objWordApp.ActiveDocument.Name, 16) = "Mail Merge Error" Then
        objWordApp.visible = True
        MsgBox "Error encountered with mail merge." & vbCrLf & vbCrLf & "Please notify IT staff.", vbCritical
        bMergeError = True
        objWordApp.ActiveDocument.Activate
        GoTo Cleanup
    End If
    
    
    ' Save the output doc
    Set objWordMergedDoc = objWordApp.Documents(objWordApp.ActiveDocument.Name)
    strOutputPath = pstrOutputBasePath & "\" & pstrProvNum & "\"
    CreateFolder (strOutputPath)
    
    If pstrInstanceStatus = "R" Then
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-Reprint-" & pstrInstanceID & ".DOC"
    Else
        pstrOutputFileName = strOutputPath & "" & pstrLetterType & "-" & pstrInstanceID & ".DOC"
    End If
    
    objWordMergedDoc.spellingchecked = True
    
    On Error Resume Next
    objWordMergedDoc.SaveAs pstrOutputFileName
    If Err.Number <> 0 Then
        SleepEvents 2
        objWordMergedDoc.SaveAs pstrOutputFileName
        Err.Clear
    End If
    On Error GoTo Error_Encountered:
    
    objWordMergedDoc.Close
    
    Set objWordMergedDoc = Nothing

    
    ' clear letter info
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_LETTER_Get_Info_tmp_clear"
    cmd.Parameters.Refresh
    'cmd.Parameters("@pInstanceID") = pstrInstanceID
    cmd.Execute
    
                                
    ' start letter transaction
    myCode_ADO.BeginTrans
    
    
    ' update LETTER status
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_LETTER_Update_Status"
    cmd.Parameters.Refresh
    cmd.Parameters("@InstanceID").Value = pstrInstanceID
    cmd.Parameters("@LetterName").Value = pstrOutputFileName
    cmd.Execute
            
    strErrMsg = Trim(cmd.Parameters("@ErrMsg").Value) & ""
    If strErrMsg <> "" Then
        myCode_ADO.RollbackTrans
        GoTo Error_Encountered
    End If
                            
                            
    ' update claim status & move to next queue
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_LETTER_AuditClaims_Update"
    cmd.Parameters.Refresh
    cmd.Parameters("@pInstanceID").Value = pstrInstanceID
    cmd.Parameters("@pInstanceStatus").Value = pstrInstanceStatus
    cmd.Execute
            
    strErrMsg = Trim(cmd.Parameters("@pErrMsg").Value)
    If strErrMsg <> "" Then
        myCode_ADO.RollbackTrans
        GoTo Error_Encountered
    End If
                                
                                
    ' commit letter transaction
    myCode_ADO.CommitTrans

    PrintLetterInstance = True
    
    GoTo Cleanup
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    PrintLetterInstance = False
    
    On Error Resume Next
    Kill pstrOutputFileName
    
Cleanup:
    
    ' Release references.
    Set objWordDoc = Nothing
    Set objWordMergedDoc = Nothing
  
    objWordApp.Quit (0) 'wdDoNotSaveChanges
    Set objWordApp = Nothing
    
    Set cmd = Nothing
    Set myCode_ADO = Nothing
    
End Function


Private Function CopyTemplates(pcolLetterTemplate As Collection, pstrLetterType) As Boolean

    Dim MyAdo As clsADO
    Dim rsLetterTemplate As ADODB.RecordSet
    Dim objLetterInfo As clsLetterTemplate
    
'    Dim Person As New ClsIdentity
    
    Dim strTemplatePath As String
    Dim strLocalTemplate As String
    Dim strSQLcmd As String
    Dim strChkFile As String
    Dim strErrMsg As String
    Dim iFolderChkLoop As Integer
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    
    CopyTemplates = False

    ' create template directory
    iFolderChkLoop = 0
    strTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\" & Identity.UserName & "\LETTERTEMPLATE"
    Do Until FolderExist(strTemplatePath) Or iFolderChkLoop = 5
        CreateFolder (strTemplatePath)
        iFolderChkLoop = iFolderChkLoop + 1
    Loop
            
    If Not FolderExist(strTemplatePath) Then
        strErrMsg = "ERROR: can not create folder " & strTemplatePath
        GoTo Error_Encountered
    End If
    
    
    ' get list of templates
    strSQLcmd = "select LetterType, TemplateLoc from LETTER_Type where AccountID = " & gintAccountID & " and LetterType = '" & pstrLetterType & "'"
    
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = strSQLcmd
    Set rsLetterTemplate = MyAdo.OpenRecordSet

    ' copy templates to local directory. Skip if template already there
    Do While Not rsLetterTemplate.EOF
        With rsLetterTemplate
            strLocalTemplate = strTemplatePath & "\" & GetFileName(!TemplateLoc)
            
            strChkFile = Dir(strLocalTemplate) & ""
            If strChkFile = "" Then
                strChkFile = Dir(!TemplateLoc) & ""
                If strChkFile <> "" Then
                    FileCopy !TemplateLoc, strLocalTemplate
                Else
                    strErrMsg = "Error: source template " & !TemplateLoc & " not found"
                End If
            End If
                    
            Set objLetterInfo = New clsLetterTemplate
            objLetterInfo.LetterType = Trim(!LetterType)
            objLetterInfo.TemplateLoc = strLocalTemplate
            pcolLetterTemplate.Add objLetterInfo, Trim(![LetterType])
            .MoveNext
        End With
    Loop
    
    CopyTemplates = True
    GoTo Clean_Up
    
Error_Encountered:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbCritical
    Else
        MsgBox Err.Source & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
    CopyTemplates = False
    
Clean_Up:
    Set rsLetterTemplate = Nothing
    Set MyAdo = Nothing

End Function

Private Sub DeleteTemplates()

    
    Dim strTemplatePath As String
    
    ' delete template directory
    strTemplatePath = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_REPOSITORY\_PROCESSING\" & Identity.UserName & "\LETTERTEMPLATE"
    DeleteFolder (strTemplatePath)

    
End Sub
