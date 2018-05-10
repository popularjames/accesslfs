Option Compare Database
Option Explicit

Public Sub HEALTHPORT_MANUALL_PROCESS()
    
    Call HP_Process_Files
    'MsgBox "DONE"
End Sub



Public Function HP_Process_Files()
    Dim db As Database
    Dim rsDAO As DAO.RecordSet
    
    Dim strSQLcmd As String
    Dim strMailSubject As String
    Dim strMailMsg As String
    Dim strMailTo As String
    
    Dim bInProgress As Boolean
    Dim bResult As Boolean
    Dim bLogFileCreated As Boolean
    
    Dim iACKFileProcessed As Integer
    Dim iRESFileProcessed As Integer
    Dim iERRFileProcessed As Integer
    
    Dim iACKFileErredOut As Integer
    Dim iRESFileErredOut As Integer
    
    Dim strACKLogFile As String
    Dim strERRLogFile As String
    Dim strFullRESLogFile As String
    
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    strMailTo = "thieu.le@connolly.com" ',alex.cannon@connolly.com,tuan.khong@connolly.com"
    
    '--------------------------------------------------------------
    ' move files: this is to circumvent SuperFlex issue.  It got
    ' issues of it's own but it still OK to use.
    '
    '  2012-02-09 disable HP_Move_File
    '--------------------------------------------------------------
    'bResult = HP_Move_File(strErrMsg)

    
    '--------------------------------------------------------------
    ' check last run status to see if it's ok to proceed
    '--------------------------------------------------------------
    Set db = CurrentDb()
    Set rsDAO = db.OpenRecordSet("HP_Auto_Process_Log")
    
    If rsDAO.recordCount = 0 Then
        ' first run: insert log record.
        strSQLcmd = "insert into HP_Auto_Process_Log " & _
                    " select now() as LastRunDate, 'BEGIN' as RunModule, 'R' as RunResult, True as InProgress"
        db.Execute (strSQLcmd)
        
        Set rsDAO = db.OpenRecordSet("HP_Auto_Process_Log")
        bInProgress = False
        
    Else
        bInProgress = rsDAO("InProgress")
    End If
    
    If bInProgress = True Then
        ' last run still in progress.   Send email and stop
        strMailSubject = "HP File processing already in progress"
        strMailMsg = "Last run still in progress.  Process not started.  Last run result = " & rsDAO("RunResult") & _
                     ".  Last run date = " & Format(rsDAO("LastRunDate"), "mm-dd-yyyy hh:mm:ss")
        Call Send_Mail(strMailSubject, strMailMsg, strMailTo)
        GoTo Exit_Function
    End If
    
    
    
    '----------------------------------------------
    ' process MR Request ACK files
    '----------------------------------------------
    strSQLcmd = "update HP_Auto_Process_Log " & _
                "set LastRunDate = now(), RunModule = 'MR_REQUEST_ACK', RunResult = 'R', InProgress = True "
    db.Execute (strSQLcmd)
        
    iACKFileProcessed = 0
    
    bResult = HP_Process_MR_ACK_Files(strACKLogFile, iACKFileProcessed, iACKFileErredOut, strErrMsg)
            
    If bResult = False Then
        If strACKLogFile <> "" Then
            strErrMsg = "MR request ACK file processing returned an error: " & strErrMsg & vbCrLf & vbCrLf & _
                        "Please check the log file <file://" & strACKLogFile & "> for more detail"
        Else
            strErrMsg = "MR request ACK file processing returned an error: " & strErrMsg
        End If
        Err.Raise vbObjectError + 513, "", strErrMsg
    End If
        
        
        
    '----------------------------------------------
    ' check for error files and send email.
    ' process will continue
    '----------------------------------------------
    strSQLcmd = "update HP_Auto_Process_Log " & _
                "set LastRunDate = now(), RunModule = 'MR_REQUEST_ERR', RunResult = 'R', InProgress = True "
    db.Execute (strSQLcmd)
        
    bResult = HP_Process_MR_ERR_Files(strERRLogFile, iERRFileProcessed, strErrMsg)
            
    If bResult = False Then
        If strACKLogFile <> "" Then
            strErrMsg = "MR request ERR file processing returned an error: " & strErrMsg & vbCrLf & vbCrLf & _
                        "Please check the log file <file://" & strERRLogFile & "> for more detail"
        Else
            strErrMsg = "MR request ERR file processing returned an error: " & strErrMsg
        End If
        Err.Raise vbObjectError + 513, "", strErrMsg
    End If
        
    

    '----------------------------------------------
    ' process MR response file
    '----------------------------------------------
    strSQLcmd = "update HP_Auto_Process_Log " & _
                "set LastRunDate = now(), RunModule = 'MR_RESPONSE', RunResult = 'R', InProgress = True "
    db.Execute (strSQLcmd)
        
    bResult = HP_Process_RESPONSE_Files(strFullRESLogFile, iRESFileProcessed, iRESFileErredOut, strErrMsg)
            
    If bResult = False Then
        If strFullRESLogFile <> "" Then
            strErrMsg = "MR RESPONSE file processing returned an error: " & strErrMsg & vbCrLf & vbCrLf & _
                        "Please check the log file <file://" & strFullRESLogFile & "> for more detail"
        Else
            strErrMsg = "MR RESPONSE file processing returned an error: " & strErrMsg
        End If
        Err.Raise vbObjectError + 513, "", strErrMsg
    End If
    
    
    
    '----------------------------------------------
    ' generate outstanding MR report
    '----------------------------------------------
    If Format(rsDAO("LastOMRReportDt"), "mm-dd-yyyy") <> Format(Date, "mm-dd-yyyy") Then
        HP_Create_Outstanding_MR_Requests
    
        strSQLcmd = "update HP_Auto_Process_Log " & _
                    "set LastOMRReportDt = now()"
        db.Execute (strSQLcmd)
    End If
    
    
    '----------------------------------------------
    ' update log file with complete status and
    ' send completion mail only if there are files
    ' processed.
    '----------------------------------------------
    strSQLcmd = "update HP_Auto_Process_Log " & _
                "set LastRunDate = now(), RunModule = 'COMPLETED', RunResult = 'C', InProgress = False "
    db.Execute (strSQLcmd)
        
    If (iERRFileProcessed + iACKFileProcessed + iRESFileProcessed) <> 0 Then
        strMailSubject = "HP - File processing completed."
        strMailMsg = ""
    
        If iERRFileProcessed > 0 Then
            strMailMsg = strMailMsg & "MR request ERR file processing result:" & vbCrLf & _
                        String(40, "-") & vbCrLf & "Total MR request ERR files detected: " & CStr(iERRFileProcessed)
            
            If strERRLogFile <> "" Then
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & "Please see log file <file://" & strERRLogFile & "> for more detail" & vbCrLf & vbCrLf & vbCrLf
            Else
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & vbCrLf
            End If
        End If
        
        If iACKFileProcessed > 0 Then
            strMailMsg = "MR request ACK file processing results:" & vbCrLf & _
                         String(40, "-") & vbCrLf & "Total MR request ACK files processed: " & CStr(iACKFileProcessed) & vbCrLf & _
                         "Files with error: " & CStr(iACKFileErredOut)
            
            If strACKLogFile <> "" Then
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & "Please see log file <file://" & strACKLogFile & "> for more detail" & vbCrLf & vbCrLf & vbCrLf
            Else
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & vbCrLf
            End If
        
        End If
    
        If iRESFileProcessed > 0 Then
            strMailMsg = CStr(iRESFileProcessed) & " MR RESPONSE files:" & vbCrLf & _
                         String(40, "-") & vbCrLf & "Total MR RESPONSE files processed: " & CStr(iRESFileProcessed) & vbCrLf & _
                         "Files with error: " & CStr(iRESFileErredOut)
            
            If strFullRESLogFile <> "" Then
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & "Please see log file <file://" & strFullRESLogFile & "> for more detail" & vbCrLf & vbCrLf & vbCrLf
            Else
                strMailMsg = strMailMsg & vbCrLf & vbCrLf & vbCrLf
            End If
        End If
    
        Call Send_Mail(strMailSubject, strMailMsg, strMailTo)
    End If
        

Exit_Function:
    Set db = Nothing
    Set rsDAO = Nothing
    Pause (1)
    Exit Function

Err_handler:
    strMailSubject = "HP file processing error"
    strMailMsg = Err.Description
    Call Send_Mail(strMailSubject, strMailMsg, strMailTo)
    
    strSQLcmd = "update HP_Auto_Process_Log " & _
                "set RunResult = 'E', InProgress = False "
    db.Execute (strSQLcmd)
    
    Resume Exit_Function
    
End Function

Public Sub HP_Create_MR_Request_MASS_MOVE()

    Dim MyAdo As New clsADO
    Dim MyCodeAdo As New clsADO
    Dim rsProviders As ADODB.RecordSet
    Dim rsMRRequest As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    
    Dim iMRFileNum
    Dim iRetCd As Integer
    Dim strMRFileName As String
    Dim strMRFtpFileName As String
    Dim strMRArchivedFileName As String
    Dim strCnlyProvID As String
    Dim strInstanceID As String
    Dim strMRFolder As String
    Dim strMRFtpFolder As String
    Dim strMRArchivedFolder As String
    
    Dim strErrMsg As String
    
    Dim bRtnCd As Boolean
    
    On Error GoTo Err_handler
    
    '  thieu hardcode this.
    strMRFolder = "Y:\Raw\CMS\Healthport\Outbound\Request\"
    strMRFtpFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Outbound\Request\"
    strMRArchivedFolder = "Y:\Raw\CMS\Healthport\Archived\Outbound\Request\"
    
       
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = "select * from HP_Adhoc_MR_Request_Providers"
    
    Set rsProviders = MyAdo.OpenRecordSet
    Close
    
    If rsProviders.recordCount > 0 Then
        rsProviders.MoveFirst
        While Not rsProviders.EOF
            strCnlyProvID = rsProviders("CnlyProvID")
            If strCnlyProvID & "" <> "" Then
                strInstanceID = Nz(rsProviders("InstanceID"), "")
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_HP_Create_MR_Request"
                cmd.Parameters.Refresh
                cmd.Parameters("@pCnlyProvID") = strCnlyProvID
                cmd.Parameters("@pInstanceID") = strInstanceID
                cmd.Execute
                
                iRetCd = cmd.Parameters("@RETURN_VALUE")
                strErrMsg = cmd.Parameters("@pErrMsg")
                If iRetCd <> 0 Or strErrMsg <> "" Then
                    MsgBox "ERRROR ENCOUNTER.  Provider  =[" & strCnlyProvID & "]"
                Else
                    strMRFileName = cmd.Parameters("@pFileName")
                    strMRFtpFileName = strMRFtpFolder & strMRFileName
                    strMRArchivedFileName = strMRArchivedFolder & strMRFileName
                    
                    MyCodeAdo.SQLTextType = sqltext
                    MyCodeAdo.sqlString = "select * from v_HP_MR_Request_Export where ExportFileName = '" & strMRFileName & "' order by RowOrder"
                    Set rsMRRequest = MyCodeAdo.OpenRecordSet
                                    
                    If rsMRRequest.recordCount > 0 Then
                        iMRFileNum = FreeFile
                        strMRFileName = strMRFolder & strMRFileName
                        
                        Open strMRFileName For Output As #iMRFileNum
                        
                        rsMRRequest.MoveFirst
                        While Not rsMRRequest.EOF
                            Print #iMRFileNum, rsMRRequest("ADR_Export_Txt")
                            rsMRRequest.MoveNext
                        Wend
                        Close #iMRFileNum
                        
                        bRtnCd = CopyFile(strMRFileName, strMRFtpFileName, True, strErrMsg)
                        If bRtnCd = False Then
                            Err.Raise vbObjectError + 513, "HP_Create_MR_Request - copy to FPT", strErrMsg
                        End If
                        
                        bRtnCd = MoveFile(strMRFileName, strMRArchivedFileName, True, strErrMsg)
                        If bRtnCd = False Then
                            Err.Raise vbObjectError + 513, "HP_Create_MR_Request - move to archive", strErrMsg
                        End If
                    End If
                    
                    
                End If
            End If
            
            rsProviders.MoveNext
        Wend
    End If
    MsgBox "Done", vbInformation + vbOKOnly, "Process Completed"
    Close
    
Exit_Sub:
    Set MyAdo = Nothing
    Set MyCodeAdo = Nothing
    Set rsProviders = Nothing
    Set rsMRRequest = Nothing
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then strErrMsg = Err.Description & " -- " & strErrMsg Else strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub



Public Sub HP_Create_MR_Request(strCnlyProvID As String, strSessionID As String, Optional strInstanceID As String)
    
    Dim MyAdo As New clsADO
    Dim MyCodeAdo As New clsADO
    Dim rsProviders As ADODB.RecordSet
    Dim rsMRRequest As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    
    Dim iMRFileNum
    Dim iRetCd As Integer
    Dim strMRFileName As String
    Dim strMRFtpFileName As String
    Dim strMRArchivedFileName As String
    '''''Dim strCnlyProvID As String
    '''''Dim strInstanceID As String
    Dim strMRFolder As String
    Dim strMRFtpFolder As String
    Dim strMRArchivedFolder As String
    
    Dim strErrMsg As String
    
    Dim bRtnCd As Boolean
    
    On Error GoTo Err_handler
    
    '  thieu hardcode this.
    
    strMRFolder = "Y:\Raw\CMS\Healthport\Outbound\Request\"
    strMRFtpFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Outbound\Request\"
    strMRArchivedFolder = "Y:\Raw\CMS\Healthport\Archived\Outbound\Request\"
           
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyCodeAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    '''''myADO.SQLTextType = sqltext
    '''''myADO.SqlString = "select * from HP_Adhoc_MR_Request_Providers"
    
    ''''Set rsProviders = myADO.OpenRecordSet
    ''''Close
    
    '''''If rsProviders.RecordCount > 0 Then
    '''''    rsProviders.MoveFirst
    '''''    While Not rsProviders.EOF
           ''''' strCnlyProvID = rsProviders("CnlyProvID")
           ''''' If strCnlyProvID & "" <> "" Then
               ''''' strInstanceID = Nz(rsProviders("InstanceID"), "")
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_HP_Create_MR_Request"
                cmd.Parameters.Refresh
                cmd.Parameters("@pCnlyProvID") = strCnlyProvID
                cmd.Parameters("@pInstanceID") = strInstanceID
                cmd.Parameters("@SessionID") = strSessionID
                cmd.Execute
                
                iRetCd = cmd.Parameters("@RETURN_VALUE")
                strErrMsg = cmd.Parameters("@pErrMsg")
                If iRetCd <> 0 Or strErrMsg <> "" Then
                    MsgBox "ERRROR ENCOUNTER.  Provider  =[" & strCnlyProvID & "]"
                Else
                    strMRFileName = cmd.Parameters("@pFileName")
                    strMRFtpFileName = strMRFtpFolder & strMRFileName
                    strMRArchivedFileName = strMRArchivedFolder & strMRFileName
                    
                    MyCodeAdo.SQLTextType = sqltext
                    ' need to add where clause to select only for this provider.
                    MyCodeAdo.sqlString = "select * from v_HP_MR_Request_Export where ExportFileName = '" & strMRFileName & "' order by RowOrder"
                    'myCodeADO.SqlString = "select * from v_HP_MR_Request_Export where SessionID = '" & strSessionID & "' order by RowOrder"
                    Set rsMRRequest = MyCodeAdo.OpenRecordSet
                                    
                    If rsMRRequest.recordCount > 0 Then
                        iMRFileNum = FreeFile
                        strMRFileName = strMRFolder & strMRFileName
                        
                        Open strMRFileName For Output As #iMRFileNum
                        
                        rsMRRequest.MoveFirst
                        While Not rsMRRequest.EOF
                            Print #iMRFileNum, rsMRRequest("ADR_Export_Txt")
                            rsMRRequest.MoveNext
                        Wend
                        Close #iMRFileNum
                        
                        bRtnCd = CopyFile(strMRFileName, strMRFtpFileName, True, strErrMsg)
                        If bRtnCd = False Then
                            Err.Raise vbObjectError + 513, "HP_Create_MR_Request - copy to FPT", strErrMsg
                        End If
                        
                        bRtnCd = MoveFile(strMRFileName, strMRArchivedFileName, True, strErrMsg)
                        If bRtnCd = False Then
                            Err.Raise vbObjectError + 513, "HP_Create_MR_Request - move to archive", strErrMsg
                        End If
                    End If
                    
                    
                End If
            '''''End If
            
            '''''rsProviders.MoveNext
       ''''' Wend
    '''''End If
    '''''MsgBox "done"
    Close
    
Exit_Sub:
    Set MyAdo = Nothing
    Set MyCodeAdo = Nothing
    Set rsProviders = Nothing
    Set rsMRRequest = Nothing
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then strErrMsg = Err.Description & " -- " & strErrMsg Else strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
End Sub


Public Sub HP_Create_Outstanding_MR_Requests()
    Dim MyCodeAdo As New clsADO
    Dim rsMRRequest As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    
    Dim iMRFileNum
    Dim iRetCd As Integer
    Dim strMRFileName As String
    Dim strMRFtpFileName As String
    Dim strCNLYFileName As String
    Dim strMRArchivedFileName As String
    Dim strCnlyProvID As String
    Dim strInstanceID As String
    Dim strMRFolder As String
    Dim strMRFtpFolder As String
    Dim strMRArchivedFolder As String
    
    Dim strErrMsg As String
    
    Dim bRtnCd As Boolean
    
    On Error GoTo Err_handler
    
    '  thieu hardcode this.
    strMRFolder = "Y:\Raw\CMS\Healthport\Outbound\Request\"
    strMRFtpFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Outbound\Outstanding_Requests\"
    strMRArchivedFolder = "Y:\Raw\CMS\Healthport\Archived\Outbound\Request\"
    
       
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    MyCodeAdo.SQLTextType = sqltext
    MyCodeAdo.sqlString = "select * from v_HP_Outstanding_MR_Requests where PaperRec = 'N' order by ProcessedDt, ExportFileName, RowOrder"
    Set rsMRRequest = MyCodeAdo.OpenRecordSet
                    
    If rsMRRequest.recordCount > 0 Then
        strMRFileName = "Connolly_" & Trim(str$(rsMRRequest.recordCount)) & "_Outstanding_MR_Requests_" & Format(Date, "yyyy_mm_dd") & ".csv"
        strMRFtpFileName = strMRFtpFolder & strMRFileName
        strCNLYFileName = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\HealthPort\Outstanding MR requests\" & strMRFileName
        strMRArchivedFileName = strMRArchivedFolder & strMRFileName
        
        iMRFileNum = FreeFile
        strMRFileName = strMRFolder & strMRFileName
        
        Open strMRFileName For Output As #iMRFileNum
        
        rsMRRequest.MoveFirst
        While Not rsMRRequest.EOF
            Print #iMRFileNum, rsMRRequest("ADR_Export_Txt")
            rsMRRequest.MoveNext
        Wend
        Close #iMRFileNum
        
        bRtnCd = CopyFile(strMRFileName, strMRFtpFileName, True, strErrMsg)
        If bRtnCd = False Then
            Err.Raise vbObjectError + 513, "HP_Create_Outstanding_MR_Requests - copy to FPT", strErrMsg
        End If
        
        
        bRtnCd = CopyFile(strMRFileName, strCNLYFileName, True, strErrMsg)
        If bRtnCd = False Then
            Err.Raise vbObjectError + 513, "HP_Create_Outstanding_MR_Requests - copy to field", strErrMsg
        End If
        
        bRtnCd = MoveFile(strMRFileName, strMRArchivedFileName, True, strErrMsg)
        If bRtnCd = False Then
            Err.Raise vbObjectError + 513, "HP_Create_Outstanding_MR_Requests - move to archive", strErrMsg
        End If
    End If
    
    Close
    
Exit_Sub:
    Set MyCodeAdo = Nothing
    Set rsMRRequest = Nothing
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then strErrMsg = Err.Description & " -- " & strErrMsg Else strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
End Sub




Public Function HP_Process_MR_ACK_Files(ACKLogFile As String, ACKFileProcessed As Integer, _
                                        ACKFileErredOut As Integer, ErrMsg As String) As Boolean
    Dim fso As FileSystemObject
    
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    Dim strLogFolder As String
    Dim strInboundFolder As String
    Dim strArchiveFolder As String
    
    Dim strInFile As String
    Dim strMRACKFile As String
    Dim strMRACKArchiveFile As String
    Dim strCosmosMapFile As String
    
    Dim strLineData As String
    
    Dim iMRACKFileRowCount As Integer
    Dim iMRACKFileNum
    
    Dim iLogFile
    
    Dim iRetCd As Integer
    
    Dim bResult As Boolean
    Dim bLogCreated As Boolean


    On Error GoTo Err_handler
    
    
    ' init variables
    ACKLogFile = ""
    ACKFileProcessed = 0
    ACKFileErredOut = 0
    ErrMsg = ""
    bLogCreated = False
   
    Set fso = New FileSystemObject
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    '------------------------------------------------------------------
    ' close any open files
    '------------------------------------------------------------------
    Close
    
    
    '--------------------------------------------------------------
    ' get file path
    '--------------------------------------------------------------
        ' check log path
        MyAdo.sqlString = "select * from HP_Config"
        MyAdo.SQLTextType = sqltext
    
        Set rs = MyAdo.OpenRecordSet
    
        strLogFolder = FixPath(rs("Log_Path"))
        
        If fso.FolderExists(strLogFolder) = False Then
            ' error encountered, return error message and exit
            HP_Process_MR_ACK_Files = False
            ErrMsg = "Log path <file://" & strLogFolder & "> is invalid or does not exists.  Please check."
            GoTo Exit_Function
        End If
            
        
        ' get inbound/archive folder paths
        strInboundFolder = FixPath(rs("Inbound_Work_Path")) & "Ack\"
        strArchiveFolder = FixPath(rs("Inbound_Archive_Path")) & "Ack\"

        
        If fso.FolderExists(strInboundFolder) = False Then
            HP_Process_MR_ACK_Files = False
            ErrMsg = "ERROR: Inbound path <file://" & strInboundFolder & "> is invalid or does not exists.  Please check."
            GoTo Exit_Function
        End If
            
        If fso.FolderExists(strArchiveFolder) = False Then
            HP_Process_MR_ACK_Files = False
            ErrMsg = "ERROR: Archive path <file://" & strArchiveFolder & "> is invalid or does not exists.  Please check."
            GoTo Exit_Function
        End If
    
    
    '--------------------------------------------------------------
    ' browse ACK inbound directory and process ACK files
    '--------------------------------------------------------------
    strInFile = Dir(strInboundFolder & "*.ACK")
    
    If strInFile <> "" Then
        ' open log file
        ACKLogFile = strLogFolder & "HP_MR_ACK_Processing_" & Format(Now(), "yyyy-mm-dd hhmmss") & ".log"
        iLogFile = FreeFile()
        bLogCreated = True
    
        Open ACKLogFile For Output As iLogFile
        Print #iLogFile, "Processing started @ "; Format(Now(), "hh:mm:ss")
        Print #iLogFile, ""
        
        Print #iLogFile, "Inbound path = " & strInboundFolder
        Print #iLogFile, "Archive path = " & strArchiveFolder
        Print #iLogFile, ""
    End If
        
    
    Do While strInFile <> ""
        ' count number of rows in ACK file for to validate file load later
        ACKFileProcessed = ACKFileProcessed + 1
        iMRACKFileRowCount = 0
        strMRACKFile = strInboundFolder & strInFile
        iMRACKFileNum = FreeFile()
        
        Open strMRACKFile For Input As #iMRACKFileNum
        While Not EOF(iMRACKFileNum)
            Line Input #iMRACKFileNum, strLineData
            iMRACKFileRowCount = iMRACKFileRowCount + 1
        Wend
        Close #iMRACKFileNum
        
        ' load file
        Print #iLogFile, "Processing <file://" & strMRACKFile & ">"
        strCosmosMapFile = "\\cca-audit\dfs-cms-ds\Data\CMS\Documentation and Training\_INTERNAL\PROJECT\HealthPort\Maps\HP_Import_MR_Request_ACK.tf.xml"
        bResult = RunCosmosMap(strMRACKFile, strCosmosMapFile, ErrMsg)
        
        If bResult = False Then
            ' error encountered, log error message and continue to next file
            ACKFileErredOut = ACKFileErredOut + 1
            Print #iLogFile, "Error loading ACK file <file://" & strMRACKFile & ">"
        Else
            'check record count after loading
            MyAdo.sqlString = "select * from HP_MR_Request_Worktable where FileName like '%" & strInFile & "'"
            MyAdo.SQLTextType = sqltext
    
            Set rs = MyAdo.OpenRecordSet
            If rs.recordCount <> iMRACKFileRowCount Then
                ' error encountered, log error message and continue to next file
                ACKFileErredOut = ACKFileErredOut + 1
                Print #iLogFile, "Error loading ACK file <file://" & strMRACKFile & ">.  Record count after loading does not match original record count."
            Else
                ' file loaded OK.  Run proc to update MR requests
                strMRACKFile = rs("FileName")
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = myCode_ADO.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_HP_Process_MR_Request_ACK"
                cmd.Parameters.Refresh
                cmd.Parameters("@pFileName") = strMRACKFile
                
                myCode_ADO.BeginTrans
                cmd.Execute
                
                iRetCd = cmd.Parameters("@RETURN_VALUE")
                ErrMsg = cmd.Parameters("@pErrMsg") & ""
                If iRetCd <> 0 Or ErrMsg <> "" Then
                    ' error encountered, log error message and continue to next file
                    myCode_ADO.RollbackTrans
                    ACKFileErredOut = ACKFileErredOut + 1
                    Print #iLogFile, "Error processing ACK file <file://" & strMRACKFile & ">.  " & ErrMsg
                Else
                    ' move file to archive
                    myCode_ADO.CommitTrans
                    strMRACKArchiveFile = strArchiveFolder & strInFile
                    Call fso.MoveFile(strMRACKFile, strMRACKArchiveFile)
                    If fso.FileExists(strMRACKArchiveFile) = False Then
                        ' error encountered, log error message and continue to next file
                        ACKFileErredOut = ACKFileErredOut + 1
                        Print #iLogFile, "Error moving file <file://" & strMRACKFile & "> file to archive."
                    End If
                End If
            End If
        End If
        
        strInFile = Dir()
        Print #iLogFile, ""
    Loop
        
    If bLogCreated Then Print #iLogFile, "Process ended @ "; Format(Now(), "hh:mm:ss")
    HP_Process_MR_ACK_Files = True
    
Exit_Function:
    Close
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Set cmd = Nothing
    Set fso = Nothing
    Exit Function
    
Err_handler:
    ErrMsg = Err.Description
    HP_Process_MR_ACK_Files = False
    
    Resume Exit_Function
End Function



Public Function HP_Process_MR_ERR_Files(ERRLogFile As String, ERRFileProcessed As Integer, ErrMsg As String) As Boolean
    Dim fso As FileSystemObject
    
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
   
    Dim strLogFolder As String
    Dim strInboundPath As String
    
    Dim strInFile As String
    
    Dim strMRERRFile As String
    Dim strMRERRArchiveFile As String
    Dim strCosmosMapFile As String
    
    Dim iLogFile
    

    On Error GoTo Err_handler
    
    
    ' init variables
    ERRLogFile = ""
    ERRFileProcessed = 0
    ErrMsg = ""
    
    
    Set fso = New FileSystemObject
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    
    '------------------------------------------------------------------
    ' close any open files
    '------------------------------------------------------------------
    Close
    
    
    '--------------------------------------------------------------
    ' get file path
    '--------------------------------------------------------------
        ' check log path
        MyAdo.sqlString = "select * from HP_Config"
        MyAdo.SQLTextType = sqltext
    
        Set rs = MyAdo.OpenRecordSet
    
        strLogFolder = FixPath(rs("Log_Path"))
' THIEU OVERRIDE
'strLogFolder = "m:\thieu\temp\hp\logs\"
        
        If fso.FolderExists(strLogFolder) = False Then
            ' error encountered, return error message and exit
            HP_Process_MR_ERR_Files = False
            ErrMsg = "Log path <file://" & strLogFolder & "> is invalid or does not exists.  Please check."
            GoTo Exit_Function
        End If
            
        
        ' get inbound/archive folder paths
        strInboundPath = FixPath(rs("Inbound_Work_Path")) & "ACK\"

'thieu override
'strInboundPath = "m:\thieu\temp\hp\inbound\ACK\"
        
        If fso.FolderExists(strInboundPath) = False Then
            HP_Process_MR_ERR_Files = False
            ErrMsg = "ERROR: Inbound path <file://" & strInboundPath & "> is invalid or does not exists.  Please check."
            GoTo Exit_Function
        End If
            
    
    
    '--------------------------------------------------------------
    ' browse ERR inbound directory and process ERR files
    '--------------------------------------------------------------
    strInFile = Dir(strInboundPath & "*.ERR")
    
    If strInFile <> "" Then
        ' open log file
        ERRLogFile = strLogFolder & "HP_MR_ERR_Processing_" & Format(Now(), "yyyy-mm-dd hhmmss") & ".log"
        iLogFile = FreeFile()
    
        Open ERRLogFile For Output As iLogFile
        Print #iLogFile, "Processing started @ "; Format(Now(), "hh:mm:ss")
        Print #iLogFile, ""
        
        Print #iLogFile, "Inbound path = " & strInboundPath
        Print #iLogFile, ""
    End If
        
    
    Do While strInFile <> ""
        ERRFileProcessed = ERRFileProcessed + 1
        Print #iLogFile, "Processing " & strInFile
        
        strInFile = Dir()
        Print #iLogFile, ""
    Loop
        
    
    HP_Process_MR_ERR_Files = True
    
Exit_Function:
    Close
    Set MyAdo = Nothing
    Set rs = Nothing
    Set fso = Nothing
    Exit Function
    
Err_handler:
    ErrMsg = Err.Description
    HP_Process_MR_ERR_Files = False
    
    Resume Exit_Function
End Function



Public Function HP_Process_RESPONSE_Files(strFullRESLogFile As String, iRESFileProcessed As Integer, _
                                          iRESFileErredOut As Integer, strErrMsg As String) As Boolean
    
    Dim fso As FileSystemObject
    
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    Dim strLogFolder As String
    Dim strInboundPath As String
    Dim strOutboundPath As String
    Dim strArchivePath As String
    Dim strDailyScanPath As String
    Dim strFPTPAth As String
    
    Dim strInboundFolder As String
    
    
    Dim strHPZipFile As String
    Dim strFullHPZipFile As String
    
    Dim iLogFile
    
    Dim bResult As Boolean
    

    '---------------------------------------------------------
    
    On Error GoTo Err_handler
    
    
    
    ' init variables
    strFullRESLogFile = ""
    iRESFileProcessed = 0
    iRESFileErredOut = 0
    strErrMsg = ""
    
    Set fso = New FileSystemObject
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    
    '------------------------------------------------------------------
    ' close any open files
    '------------------------------------------------------------------
    Close
    
    
    '--------------------------------------------------------------
    ' get file path
    '--------------------------------------------------------------
    ' check log path
    MyAdo.sqlString = "select * from HP_Config"
    MyAdo.SQLTextType = sqltext

    Set rs = MyAdo.OpenRecordSet

    strLogFolder = FixPath(rs("Log_Path"))
    
    If fso.FolderExists(strLogFolder) = False Then
        ' error encountered, return error message and exit
        HP_Process_RESPONSE_Files = False
        strErrMsg = "Log path <file://" & strLogFolder & "> is invalid or does not exists.  Please check."
        GoTo Exit_Function
    End If
    
    ' get inbound/archive folder paths
    strInboundPath = FixPath(rs("Inbound_Work_Path"))
    strOutboundPath = FixPath(rs("Outbound_Work_Path"))
    strArchivePath = FixPath(rs("Inbound_Archive_Path"))
    strFPTPAth = FixPath(rs("FTP_Outbound"))
    strDailyScanPath = "\\ccaintranet.com\dfs-cms-ds\raw\cms\Healthport\Inbound\DailyScans\"
    

    If fso.FolderExists(strInboundPath) = False Then
        HP_Process_RESPONSE_Files = False
        strErrMsg = "ERROR: Inbound path <file://" & strInboundPath & "> is invalid or does not exists.  Please check."
        GoTo Exit_Function
    End If
        
    If fso.FolderExists(strOutboundPath) = False Then
        HP_Process_RESPONSE_Files = False
        strErrMsg = "ERROR: Outbound path <file://" & strOutboundPath & "> is invalid or does not exists.  Please check."
        GoTo Exit_Function
    End If
    
    If fso.FolderExists(strArchivePath) = False Then
        HP_Process_RESPONSE_Files = False
        strErrMsg = "ERROR: Archive path <file://" & strArchivePath & "> is invalid or does not exists.  Please check."
        GoTo Exit_Function
    End If

    
    If fso.FolderExists(strDailyScanPath) = False Then
        HP_Process_RESPONSE_Files = False
        strErrMsg = "ERROR: DailyScan path <file://" & strDailyScanPath & "> is invalid or does not exists.  Please check."
        GoTo Exit_Function
    End If
        
        
    '-----------------------------------------------------------------------
    ' browse response inbound directory and process ZIP response files
    '-----------------------------------------------------------------------
    strInboundFolder = strInboundPath & "Response\"
    strHPZipFile = Dir(strInboundFolder & "*.zip") & ""
    
    If strHPZipFile <> "" Then
        ' open log file
        strFullRESLogFile = strLogFolder & "HP_MR_RESPONSE_Processing_" & Format(Now(), "yyyy-mm-dd hhmmss") & ".log"
        iLogFile = FreeFile()
        
        Open strFullRESLogFile For Output As iLogFile
        Print #iLogFile, "Processing started @ "; Format(Now(), "hh:mm:ss")
                
        ' log paths
        Print #iLogFile, "Inbound path  = " & strInboundPath & ""
        Print #iLogFile, "Outbound path = " & strOutboundPath & ""
        Print #iLogFile, "Archive path  = " & strArchivePath & ""
        Print #iLogFile, "DailyScan path  = " & strDailyScanPath & ""
    End If
    
    Do While strHPZipFile <> ""
        iRESFileProcessed = iRESFileProcessed + 1
        strFullHPZipFile = strInboundFolder & strHPZipFile
        
        Print #iLogFile, ""
        Print #iLogFile, "Processing " & strHPZipFile
        
        bResult = HP_Process_Single_RESPONSE_File(strFullHPZipFile, strInboundPath, strOutboundPath, strFPTPAth, strArchivePath, strDailyScanPath, iLogFile, strErrMsg)
        If bResult = False Then
            ' error encountered
            Print #iLogFile, strErrMsg
            GoTo Exit_Function
        End If
        
        strHPZipFile = Dir(strInboundFolder & "*.zip") & ""
    Loop
        
    
    If strFullRESLogFile <> "" Then
        Print #iLogFile, ""
        Print #iLogFile, "Processing ended @ "; Format(Now(), "hh:mm:ss")
    End If
    
    HP_Process_RESPONSE_Files = True
        
Exit_Function:
    Close
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Set fso = Nothing
    Exit Function
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    HP_Process_RESPONSE_Files = False

    Resume Exit_Function
End Function



Public Function HP_Process_Single_RESPONSE_File(ByVal FullHPZipFile As String, ByVal InboundPath As String, ByVal OutboundPath As String, _
                                                ByVal FPTOutboundPath, ByVal ArchivePath As String, ByVal DailyScanPath As String, iLogFile, _
                                                ErrMsg As String) As Boolean
    
    Dim fso As FileSystemObject
    Dim oFile As file
    Dim oFolder As Folder
    
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    

    Dim strFullHPFile As String
    Dim strFullHPErrFile As String
    Dim strFullCnlyFile As String
    Dim strHPFileType As String

    Dim strProvFolder As String
    Dim strOutboundFolder As String
    Dim strArchiveFolder As String
    Dim strWorkingFolder As String
    Dim strErrorFolder As String
    
    Dim strCnlyZipFile As String
    Dim strFullCnlyZipFile As String
    Dim strFullCnlyCSVFile As String
    Dim strCnlyImageFile As String
    
    Dim strHPZipFile As String
    Dim strBaseHPZipFile As String
    Dim strFullHPImageFile As String
    Dim strFullHPACKFile As String
    Dim strFullHPCSVFile As String
    
    
    Dim strFullArchiveZipFile As String
    Dim strFullErrFile As String
   
    Dim strCnlyProvID As String
    Dim strInstanceID As String
    Dim strLineData As String
    Dim strReceivedDt As String
    Dim strIndent As String
    
    Dim strSQL As String
    Dim strCosmosMapName As String
    
    Dim iCSVFileCnt As Integer
    Dim iOtherFileCnt As Integer
    Dim iImageCnt As Integer
    Dim iCSVRowCnt As Integer
    Dim iRowCnt As Integer
    Dim iMetadataFileCnt As Integer
    Dim iFolderFileCnt As Integer
    Dim iRtnCd As Integer
    Dim iProcessStatus As Integer
    Dim iFileNum
    Dim iTotalRecords As Integer
    Dim iGoodRecords As Integer
    
    Dim bResult As Boolean
    Dim bFullACK As Boolean
    
'---------------------------------
    
    On Error GoTo Err_handler
    
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")

    strIndent = ""
    
    '-----------------------------------------------------------------------------------------------------
    ' STEP #1:
    '       a - unzip file to the working folder, rename the original zip file using
    '           the naming convention: new zip file name = original zip file name + "_yyyymmddhhmmss"
    '       b - unzip file
    '       c - log file with process status = -1
    '
    '   Error handling:
    '       if any error is encountered in a,b,c,d delete the working folder and the log
    '-----------------------------------------------------------------------------------------------------
    
    
    '--------------------------------------------------------------------------
    ' a - back up the zip file first and copy zip file to the working folder
    '--------------------------------------------------------------------------
    Set fso = New FileSystemObject
    
    strHPZipFile = fso.GetFileName(FullHPZipFile)
    strBaseHPZipFile = fso.GetBaseName(FullHPZipFile)
    
    strFullErrFile = OutboundPath & "ResponseAckWithError\" & fso.GetBaseName(strHPZipFile) & ".Err"
    
    strArchiveFolder = ArchivePath & "Response\"
    strFullArchiveZipFile = strArchiveFolder & strHPZipFile
    
    ' archive original HP zip file
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Archiving zip file"
    strIndent = Space(10)
    
    Call fso.CopyFile(FullHPZipFile, strFullArchiveZipFile, True)
    If Not fso.FileExists(strFullArchiveZipFile) Then
        ErrMsg = "Error archiving file <file://" & FullHPZipFile & ">.  Process aborted"
        GoTo Exit_With_Error
    End If
    
    
    ' create work folder
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Creating work folder"
    strIndent = Space(10)
    
    strInstanceID = "_" & Format(Now(), "yyyymmddhhmmss")
    strWorkingFolder = InboundPath & "Response\" & strBaseHPZipFile & strInstanceID & "\"
    
    If CreateFolder(strWorkingFolder) = False Then
        ErrMsg = "Error in creating folder <file://" & strWorkingFolder & ">.  Please check"
        GoTo Exit_With_Error
    End If
    
    ' copy zip file to working folder
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Copying zip file to work folder"
    strIndent = Space(10)
    
    strFullCnlyZipFile = strWorkingFolder & strBaseHPZipFile & strInstanceID & "." & fso.GetExtensionName(FullHPZipFile)
    strCnlyZipFile = fso.GetFileName(strFullCnlyZipFile)
    
    Call fso.CopyFile(FullHPZipFile, strFullCnlyZipFile, True)
    If Not fso.FileExists(strFullCnlyZipFile) Then
        ErrMsg = "Error copying file <file://" & FullHPZipFile & "> to <file://" & strFullCnlyZipFile & ">.  Please check."
        GoTo Exit_With_Error
    End If
        
    
    
    '---------------------------------------------------------
    ' b - unzip file and wait
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Unzipping zip file"
    strIndent = Space(10)
    
    ErrMsg = UnZipFile(strFullCnlyZipFile, strWorkingFolder)
    If ErrMsg <> "" Then
        ErrMsg = "Error unzipping file <file://" & strFullCnlyZipFile & ">.  Error messge = [" & ErrMsg & "]"
        GoTo Exit_With_Error
    End If


    '---------------------------------------------------------
    ' c - log file with process status = -1
    '---------------------------------------------------------
    iProcessStatus = -1
    Set oFile = fso.GetFile(FullHPZipFile)
    strReceivedDt = Format(oFile.DateCreated, "mm-dd-yyyy")
    
    bResult = HP_Log_File(strCnlyZipFile, strHPZipFile, strCnlyZipFile, _
                          strReceivedDt, "ZIP_IMPORT", 1, -1, "MR response file received from Healthport", ErrMsg)

    If bResult = False Then
        GoTo Exit_With_Error
    End If




    '-----------------------------------------------------------------------------------------------------
    ' STEP #2: CHECK FILE FORMAT AND FILE CONTENT FOR FATAL ERROR
    '       a - up date log file with process status = -2
    '       b - check for fatal errors
    '               - check to make sure we have exactly ONE CSV file
    '               - check to make sure that the number of rows in the CSV file matches the number of row in
    '                  the file name
    '               - check to make sure that image file names in CSV file match physical image file name
    '       e - log each file in the working folder giving each a unique file name
    '
    '   No error proceed to step #3
    '   Fatal error encountered:
    '       a - create ERR file
    '       b - update log file with process status = 1
    '       c - move working folder to archive
    '-----------------------------------------------------------------------------------------------------
    
    '---------------------------------------------------------
    ' a - up date log file with process status = -2
    '---------------------------------------------------------
    iProcessStatus = -2
    bResult = HP_Update_Log(strCnlyZipFile, iProcessStatus, ErrMsg)
    If bResult = False Then
        GoTo Exit_With_Error
    End If
    
    
    
    '---------------------------------------------------------
    ' b - check for fatal errors
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Checking for fatal errors - Error file will be created if fatal error is encountered"
    strIndent = Space(10)
    
    Set oFolder = fso.GetFolder(strWorkingFolder)
    iFolderFileCnt = oFolder.Files.Count
    
    
    iCSVFileCnt = 0
    iImageCnt = 0
    iOtherFileCnt = 0
    
    ' log all files received.
    For Each oFile In oFolder.Files
        iRowCnt = 1
        
        ' Check for metadata CSV file
        If UCase(Right(oFile.Name, 4)) = ".CSV" Then
            iCSVFileCnt = iCSVFileCnt + 1
            strFullHPFile = oFile.Path
            strFullCnlyFile = strWorkingFolder & fso.GetBaseName(strFullHPFile) & strInstanceID & "." & fso.GetExtensionName(strFullHPFile)
            strFullHPCSVFile = strFullHPFile
            strFullCnlyCSVFile = strFullCnlyFile
            strHPFileType = "CSV_IMPORT"
            
            ' count number of rows in CSV file
            iCSVRowCnt = 0
            iFileNum = FreeFile()
            
            Open strFullHPFile For Input As #iFileNum
            While Not EOF(iFileNum)
                Line Input #iFileNum, strLineData
                iCSVRowCnt = iCSVRowCnt + 1
            Wend
            Close #iFileNum
            
            iRowCnt = iCSVRowCnt
        ElseIf InStr(1, ".PDF/.TIF", UCase(Right(oFile.Name, 4))) > 0 Then
            ' image file
            iImageCnt = iImageCnt + 1
            strFullHPFile = oFile.Path
            strFullCnlyFile = strWorkingFolder & fso.GetBaseName(strFullHPFile) & strInstanceID & "." & fso.GetExtensionName(strFullHPFile)
            strHPFileType = "MR_IMPORT"
        ElseIf InStr(1, ".CSV/.PDF/.TIF", UCase(Right(oFile.Name, 4))) = 0 Then
            ' other file
            strFullHPFile = oFile.Path
            strFullCnlyFile = strWorkingFolder & fso.GetBaseName(strFullHPFile) & strInstanceID & "." & fso.GetExtensionName(strFullHPFile)
            strHPFileType = "OTHER_IMPORT"
            If UCase(fso.GetFileName(strFullHPFile)) <> UCase(strCnlyZipFile) Then
                iOtherFileCnt = iOtherFileCnt + 1
            End If
        End If
'Debug.Print strFullCnlyFile
        
        ' log file (if not the original zip file)
        If fso.GetFileName(strFullHPFile) <> strCnlyZipFile Then
            Print #iLogFile, strIndent & "Logging file " & fso.GetFileName(strFullHPFile)
            bResult = HP_Log_File(fso.GetFileName(strFullCnlyFile), fso.GetFileName(strFullHPFile), strCnlyZipFile, _
                              strReceivedDt, strHPFileType, iRowCnt, 1, "MR response file received from Healthport", ErrMsg)
    
            If bResult = False Then
                GoTo Exit_With_Error
            End If
        End If
    Next
    
    ' case #1: no CSV file
    If iCSVFileCnt = 0 Then
        ErrMsg = "Missing CSV response file! Creating error file"
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If
        
    ' case #2: more than 1 CSV file
    If iCSVFileCnt > 1 Then
        ErrMsg = "More than one CSV response file detected! Creating error file."
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If
        
    ' case #3: CSV file does not follow naming standard
    If IsNumeric(Mid(fso.GetFileName(strFullHPCSVFile), Len("HealthPort_") + 1, 4)) Then
        ' extract file count from CSV filename
        iMetadataFileCnt = val(Mid(fso.GetFileName(strFullHPCSVFile), Len("HealthPort_") + 1, 4))
    Else
        ErrMsg = "CSV file does not follow naming standard. Creating error file."
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If
    
    ' case #4: CSV file count not matching actual file count
    If iMetadataFileCnt <> iCSVRowCnt Then
        ErrMsg = "Number of records as indicated in CSV file name does not match number of rows in CSV file.  Creating error file."
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If

    ' case #5: CSV file count not matching actual image file count
    If iCSVRowCnt <> iImageCnt Then
        ErrMsg = "Number of images does not match number of rows in CSV file. Creating error file."
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If

    ' case #6: Other files received
    If iOtherFileCnt > 0 Then
        ErrMsg = "Non-image file(s) detected in package.  Creating error file."
        Print #iLogFile, strIndent & ErrMsg
        
        bResult = HP_Create_ERR_File(strFullErrFile, "3", ErrMsg)
        If bResult = False Then
            GoTo Exit_With_Error
        End If
    End If



    '-----------------------------------------------------------------------------------------------------
    ' STEP #3: PROCESS CSV FILE
    '       a - up date log file with process status = -3
    '       b - load CSV file
    '       c - Do page count for each image.  If image is not readable, update ACK_Code in work table
    '           else update work table with image count
    '       d - run proc to process entries
    '       e - create ACK file
    '       f - copy working folder to DailyScans (image will be renamed when copying)
    '       g - move working folder to archive
    '       h - up date log file with process status = 1
    '
    '   Error handling:
    '       Will need to investigate and resolved
    '-----------------------------------------------------------------------------------------------------
    '---------------------------------------------------------
    ' a - up date log file with process status = -3
    '---------------------------------------------------------
    iProcessStatus = -3
    bResult = HP_Update_Log(strCnlyZipFile, iProcessStatus, ErrMsg)
    If bResult = False Then
        Print #iLogFile, strIndent & ErrMsg
        GoTo Exit_With_Error
    End If


    '---------------------------------------------------------
    ' b - load CSV file
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Copying CSV file to Cnly CSV file"
    strIndent = Space(10)
    
    Call fso.CopyFile(strFullHPCSVFile, strFullCnlyCSVFile, True)
    
    If Not fso.FileExists(strFullCnlyCSVFile) Then
        ErrMsg = "Error: can not copy  file " & strFullHPCSVFile & " to " & strFullCnlyCSVFile
        GoTo Exit_With_Error
    End If
    
    
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Loading CSV file"
    strIndent = Space(10)
    strCosmosMapName = "\\cca-audit\dfs-cms-ds\Data\CMS\Documentation and Training\_INTERNAL\PROJECT\HealthPort\Maps\HP_Import_MR_Response.tf.xml"
    bResult = RunCosmosMap(strFullCnlyCSVFile, strCosmosMapName, ErrMsg)
    If bResult = False Then
        GoTo Exit_With_Error
    End If
    
    
    '---------------------------------------------------------
    ' c - update page count
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Image page count check"
    strIndent = Space(10)
    
    strSQL = " SELECT * from CMS_AUDITORS_CLAIMS.dbo.HP_MR_Response_Worktable where ProcessFlag = 0 and Filename = '" & fso.GetFileName(strFullCnlyCSVFile) & "'"
    Set rs = MyAdo.OpenRecordSet(strSQL)

    If rs.BOF = True And rs.EOF = True Then
        ErrMsg = "ERROR: no row return for CSV file " & strFullCnlyCSVFile & " from work table"
        GoTo Exit_With_Error
    Else
        strFullHPImageFile = ""
        With rs
            .MoveFirst
            While Not .EOF
                strFullHPImageFile = strWorkingFolder & !ImageName
                Print #iLogFile, strIndent & "Checking image " & !ImageName
                
                If fso.FileExists(strFullHPImageFile) Then
                    Select Case UCase(Right(strFullHPImageFile, 4))
                        Case ".PDF"
                            !Cnly_PageCnt = Count_PDF_Pages(strFullHPImageFile)
                        Case ".TIF"
                            !Cnly_PageCnt = Count_TIF_Pages(strFullHPImageFile)
                        Case Else
                            !Cnly_PageCnt = -1      ' can not read image
                            Print #iLogFile, strIndent & "Image " & !ImageName & " failed page count validation"
                    End Select
                Else
                    !Cnly_PageCnt = 0              ' image does not exists
                    Print #iLogFile, strIndent & "Image " & !ImageName & " does not exists"
                End If
                
                .Update
                MyAdo.BatchUpdate rs
                
                .MoveNext
            Wend
        End With
    End If

   

    '---------------------------------------------------------
    ' d - process images
    '---------------------------------------------------------
    bFullACK = True
    
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Running proc usp_HP_Process_MR_Response"
    strIndent = Space(10)
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_HP_Process_MR_Response"
    cmd.Parameters.Refresh
    cmd.Parameters("@pMRResponseFile") = fso.GetFileName(strFullCnlyCSVFile)
    
    myCode_ADO.BeginTrans
    cmd.Execute
    iTotalRecords = cmd.Parameters("@pTotalRecords")
    iGoodRecords = cmd.Parameters("@pGoodRecords") & ""
    ErrMsg = cmd.Parameters("@pErrMsg") & ""
    
    If iTotalRecords <> iGoodRecords Then bFullACK = False
    
    iRtnCd = cmd.Parameters("@RETURN_VALUE")
    If iRtnCd <> 0 Or ErrMsg <> "" Then
        myCode_ADO.RollbackTrans
        GoTo Exit_With_Error
    Else
        myCode_ADO.CommitTrans
    End If
    
    
    '---------------------------------------------------------
    ' e - up date log file with process status = -4
    '     if error occurs after this stage it would have to be
    '     manually corrected
    '---------------------------------------------------------
    
    strIndent = Space(5)
    Print #iLogFile, strIndent; "NOTE: Any error encountered after this point would have to be manuall corrected"
    
    iProcessStatus = -4
    bResult = HP_Update_Log(strCnlyZipFile, iProcessStatus, ErrMsg)
    If bResult = False Then
        Print #iLogFile, strIndent & ErrMsg
        HP_Process_Single_RESPONSE_File = False
        GoTo Exit_Function
    End If
    
    
    '---------------------------------------------------------
    ' e - create ACK file
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Creating ACK file"
    strIndent = Space(10)
    
    strOutboundFolder = OutboundPath & "ResponseACK\"
    bResult = HP_Create_ACK_File(strOutboundFolder, fso.GetFileName(strFullCnlyCSVFile), strCnlyZipFile, strFullHPACKFile, ErrMsg)
    If bResult = False Then
        Print #iLogFile, strIndent & ErrMsg
        HP_Process_Single_RESPONSE_File = False
        GoTo Exit_Function
    End If
    
    
    
    '---------------------------------------------------------
    ' f - copy images to DailyScan
    '---------------------------------------------------------
    strIndent = Space(5)
    Print #iLogFile, strIndent & "Copying images to DailyScan folder"
    strIndent = Space(10)
    
    strSQL = "SELECT il.CnlyProvID, il.ImageName as CnlyImageFile, wk.* " & _
             "from dbo.HP_MR_Response wk " & _
             "join dbo.SCANNING_Image_Log_Tmp il on il.ScannedDt = wk.ScannedDt " & _
             "where wk.ProcessFlag = 1 " & _
             "and wk.AckCode in (1,4) " & _
             "and wk.Filename = '" & fso.GetFileName(strFullCnlyCSVFile) & "'"
    Set rs = MyAdo.OpenRecordSet(strSQL)

    If rs.BOF = True And rs.EOF = True Then
        ErrMsg = "Error: No row with ACKCode = 1 returned for CSV file " & strFullCnlyCSVFile
        Print #iLogFile, strIndent & ErrMsg
        bFullACK = False
    Else
        With rs
            .MoveFirst
            While Not .EOF
                strCnlyProvID = rs("CnlyProvID")
                strProvFolder = DailyScanPath & strCnlyProvID
                bResult = CreateFolder(strProvFolder)
                If bResult = False Then
                    ErrMsg = "Error: can not create provider folder " & strProvFolder
                    Print #iLogFile, strIndent & ErrMsg
                    HP_Process_Single_RESPONSE_File = False
                    GoTo Exit_Function
                End If
                
                strFullHPImageFile = strWorkingFolder & rs("ImageName")
                strCnlyImageFile = strProvFolder & "\" & rs("CnlyImageFile")
                fso.CopyFile strFullHPImageFile, strCnlyImageFile, True
                If fso.FileExists(strCnlyImageFile) = False Then
                    ErrMsg = "Error: can not copy image file " & strFullHPImageFile & " to DailyScan folder"
                    Print #iLogFile, strIndent & ErrMsg
                    HP_Process_Single_RESPONSE_File = False
                    GoTo Exit_Function
                End If
                .MoveNext
            Wend
        End With
    End If


'    If bFullACK = True Then
        '---------------------------------------------------------
        ' copy ACK file to ftp outbound
        '---------------------------------------------------------
        Call fso.CopyFile(strFullHPACKFile, FPTOutboundPath & "ResponseACK\" & fso.GetFileName(strFullHPACKFile))
        
        '---------------------------------------------------------
        ' g - archive folder
        '---------------------------------------------------------
        strIndent = Space(5)
        Print #iLogFile, strIndent & "Archiving working folder"
        strIndent = Space(10)
        
        If Right(strWorkingFolder, 1) = "\" Then strWorkingFolder = left(strWorkingFolder, Len(strWorkingFolder) - 1)
        
        Call fso.MoveFolder(strWorkingFolder, strArchiveFolder)
        
        
        '---------------------------------------------------------
        ' d - delete inbound zip file
        '---------------------------------------------------------
        strIndent = Space(5)
        Print #iLogFile, strIndent & "deleting original zip file"
        strIndent = Space(10)
        
        Call fso.DeleteFile(FullHPZipFile)
        If fso.FileExists(FullHPZipFile) Then
            ErrMsg = "Error deleting original zip file.  File = <file://" & FullHPZipFile & ">"
            Print #iLogFile, strIndent & ErrMsg
            HP_Process_Single_RESPONSE_File = False
            GoTo Exit_Function
        End If
'    Else
'        '---------------------------------------------------------
'        ' move zip file and folder to error folder
'        '---------------------------------------------------------
'        strErrorFolder = InboundPath & "ResponseWithIssue\"
'        Call fso.MoveFile(FullHPZipFile, strErrorFolder & fso.GetFileName(FullHPZipFile))
'        'Call fso.MoveFolder(strWorkingFolder, strErrorFolder & fso.GetFileName(strWorkingFolder))
'        Call fso.CopyFile(strFullHPACKFile, FPTOutboundPath & "ResponseAckWithError\" & fso.GetFileName(strFullHPACKFile))
'    End If
    
    '---------------------------------------------------------
    ' f - update log
    '---------------------------------------------------------
    iProcessStatus = 1
    bResult = HP_Update_Log(strCnlyZipFile, iProcessStatus, ErrMsg)
    If bResult = False Then
        HP_Process_Single_RESPONSE_File = False
        Print #iLogFile, strIndent & ErrMsg
    Else
        HP_Process_Single_RESPONSE_File = True
    End If

    strIndent = Space(5)
    Print #iLogFile, strIndent & "Processing file " & strHPZipFile & " completed"
    
    GoTo Exit_Function


Exit_With_Error:
    HP_Process_Single_RESPONSE_File = False
    Print #iLogFile, strIndent & ErrMsg
    
    MyAdo.sqlString = "delete from dbo.HP_MR_Response_Worktable where Filename = '" & fso.GetFileName(strFullCnlyCSVFile) & "'"
    MyAdo.Execute
    
    If fso.FolderExists(strWorkingFolder) Then
        If Right(strWorkingFolder, 1) = "\" Then strWorkingFolder = left(strWorkingFolder, Len(strWorkingFolder) - 1)
        fso.DeleteFolder (strWorkingFolder)
    End If
    
    bResult = HP_Purge_All_Logs(strCnlyZipFile, ErrMsg)
    If bResult = False Then
        Print #iLogFile, strIndent & ErrMsg
    End If
    

Exit_Function:
    Set fso = Nothing
    Set oFile = Nothing
    Set oFolder = Nothing
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set rs = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_handler:
    If ErrMsg = "" Then ErrMsg = Err.Description
    
    ' this block is a repeat of the Exit_With_Error block but it has to be repeated here for error trapping.
    HP_Process_Single_RESPONSE_File = False
    Print #iLogFile, strIndent & ErrMsg
    
    MyAdo.sqlString = "delete from dbo.HP_MR_Response_Worktable where Filename = '" & fso.GetFileName(strFullCnlyCSVFile) & "'"
    MyAdo.Execute
    
    If fso.FolderExists(strWorkingFolder) Then
        If Right(strWorkingFolder, 1) = "\" Then strWorkingFolder = left(strWorkingFolder, Len(strWorkingFolder) - 1)
        fso.DeleteFolder (strWorkingFolder)
    End If
    
    bResult = HP_Purge_All_Logs(strCnlyZipFile, ErrMsg)
    If bResult = False Then
        Print #iLogFile, strIndent & ErrMsg
    End If
    
    Resume Exit_With_Error
End Function



Public Function HP_Process_Single_RESPONSE_Row(folderName As String, ProvID As String, ImageName As String, Icn As String, InstanceId As String) As Boolean
    
    Dim fso As New FileSystemObject
    Dim oFile As file
    
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    
    Dim strBaseFolderName As String
    Dim strInstance As String
    Dim strFullCnlyImageName As String
    Dim strFullHPImageName As String

    Dim strErrMsg As String
    Dim iRtnCd As Integer
    
'---------------------------------
    
    On Error GoTo Err_handler
    
    ' set variables
    strBaseFolderName = fso.GetBaseName(folderName)
    strInstance = Right(strBaseFolderName, 15)
    strFullHPImageName = folderName + "\" + ImageName
    strFullCnlyImageName = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\MEDICALRECORD_TEMP\DailyScans\" & ProvID & "\" & ImageName & _
                            strInstance & "." & fso.GetExtensionName(strFullHPImageName)
    
    
    
    ' process database entry
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_HP_Process_MR_Response_Row"
    cmd.Parameters.Refresh
    cmd.Parameters("@pMRResponseFile") = strBaseFolderName
    cmd.Parameters("@pImageName") = ImageName
    cmd.Parameters("@pICN") = Icn
    cmd.Parameters("@pInstanceID") = InstanceId
    
    
    myCode_ADO.BeginTrans
    cmd.Execute
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    
    iRtnCd = cmd.Parameters("@RETURN_VALUE")
    If iRtnCd <> 0 Or strErrMsg <> "" Then
        myCode_ADO.RollbackTrans
        GoTo Exit_With_Error
    Else
        myCode_ADO.CommitTrans
    End If
    
    
    
    '---------------------------------------------------------
    ' copy image to DailyScan
    '---------------------------------------------------------
    
    Call fso.CopyFile(strFullHPImageName, strFullCnlyImageName)
    
    
    
    GoTo Exit_Function


Exit_With_Error:
    HP_Process_Single_RESPONSE_Row = False
    

Exit_Function:
    Set fso = Nothing
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Function
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    myCode_ADO.RollbackTrans
    Resume Exit_With_Error
End Function





Private Function RunCosmosMap(strSourceFile As String, strMapFile As String, strErrMsg As String) As Boolean
    Dim strCosmosMapName As String
    Dim strSourceFileName As String
    Dim iErrorCount As Integer
    Dim lRecordCount As Long
    Dim strTargetFolder As String
    
    
    On Error GoTo ErrorHandler

    strErrMsg = ""
    strSourceFileName = strSourceFile
    strCosmosMapName = strMapFile
    

    '-------------------------------------------------------------------
    ' 1. Creating an engine object and its members.
    '-------------------------------------------------------------------
    ' Creating the engine object.
    Dim djEngine
    Set djEngine = CreateObject("DJEC.Engine")
    djEngine.InitializationFile = "C:\Program Files\Pervasive\Cosmos\Common800\dj800.ini"

    ' Creating the conversion object.
    Dim djConversion
    Set djConversion = CreateObject("DJEC.Conversion")

    ' Creating the log object.
    Dim djLog
    Set djLog = CreateObject("DJEC.LogManager")
    Set djLog = djConversion.MessageLog
    djLog.FileName = strSourceFileName & "_ErrorMessageLog.log"


    '-------------------------------------------------------------------
    ' 2. Load map
    '-------------------------------------------------------------------
    djConversion.Load (strCosmosMapName)
    

    '-------------------------------------------------------------------
    '3. set data source
    '-------------------------------------------------------------------
    djConversion.Sources(0).connectioninfo.file = strSourceFileName


    '-------------------------------------------------------------------
    ' 4. Running map conversion
    '-------------------------------------------------------------------
    djConversion.Run


    RunCosmosMap = True

Exit_Function:
    Set djEngine = Nothing
    Set djConversion = Nothing
    Set djLog = Nothing
    Exit Function

ErrorHandler:
    strErrMsg = Err.Description
    RunCosmosMap = False
    Resume Exit_Function

End Function


Public Sub Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Pause
    
    Dim PauseTime As Variant, start As Variant
    
    PauseTime = NumberOfSeconds
    start = Timer
    Do While Timer < start + PauseTime
        DoEvents
    Loop
    
Exit_Pause:
    Exit Sub
    
Err_Pause:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Pause

End Sub


Public Sub Send_Mail(MailSubject As String, MailMsg As String, MailTo As String)
    Dim ShellCmd As String
    
    ShellCmd = "sqlcmd -E -S " & CurrentCMSServer() & " -d Cnly  -Q ""declare @result as tinyint; declare @mailoutput as xml; " & _
               " exec Mail.SqlNotifySend '" & MailSubject & "', '" & MailTo & "', NULL, NULL, '" & _
               MailMsg & "', @Result OUT, @MailOutput OUT"""

    Shell ShellCmd
End Sub


Private Function FixPath(InPath As String) As String
    InPath = Replace(InPath, "/", "\")
    If Right(InPath, 1) <> "\" Then InPath = InPath + "\"
    
    FixPath = InPath
End Function


Public Function HP_Log_File(strFileName As String, strOrigFileName As String, strZipFileName As String, strFileDate As String, _
                            strFileType As String, iTotalRecord As Integer, iProcessFlag As Integer, strProcessMsg As String, strErrMsg As String)
    
    Dim strSQLCode As String
    Dim MyAdo As clsADO
    Dim iRetCd As Integer
    
    On Error GoTo ErrorHandler

    strErrMsg = ""

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQLCode = "insert into CMS_AUDITORS_CLAIMS.dbo.HP_File_Log " & _
                        "values (" & _
                        "'" & strFileName & "'," & _
                        "'" & strOrigFileName & "'," & _
                        "'" & strZipFileName & "'," & _
                        "'" & strFileDate & "'," & _
                        "'" & strFileType & "'," & _
                        "'" & iTotalRecord & "'," & _
                        "'" & iProcessFlag & "'," & _
                        "'" & strProcessMsg & "'," & _
                        "'" & Now() & "')"
    
    MyAdo.sqlString = strSQLCode
    MyAdo.SQLTextType = sqltext
    iRetCd = MyAdo.Execute
    If iRetCd = -1 Then
        HP_Log_File = False
        strErrMsg = "Error executing SQL command [" & MyAdo.sqlString & "]"
        GoTo Exit_Function
    End If

    HP_Log_File = True

Exit_Function:
    Set MyAdo = Nothing
    Exit Function

ErrorHandler:
    strErrMsg = Err.Description
    HP_Log_File = False
    GoTo Exit_Function
    
End Function



Public Function HP_Update_Log(strFileName As String, iStatus As Integer, strErrMsg As String) As Boolean
    Dim strSQLCode As String
    Dim MyAdo As clsADO
    Dim iRetCd As Integer
    
    On Error GoTo Err_handler

    strErrMsg = ""
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQLCode = "update dbo.HP_File_Log set ProcessFlag = " & CStr(iStatus) & " where FileName = '" & strFileName & "'"
    MyAdo.sqlString = strSQLCode
    MyAdo.SQLTextType = sqltext
    iRetCd = MyAdo.Execute
    If iRetCd = -1 Then
        HP_Update_Log = False
        strErrMsg = "Error executing SQL command [" & MyAdo.sqlString & "]"
        GoTo Exit_Function
    End If

    HP_Update_Log = True
    
Exit_Function:
    Set MyAdo = Nothing
    Exit Function

Err_handler:
    strErrMsg = Err.Description
    HP_Update_Log = False
    Resume Exit_Function
    
End Function


Public Function HP_Delete_Log(FileName As String, ErrMsg As String) As Boolean
    Dim strSQLCode As String
    Dim MyAdo As clsADO
    Dim iRetCd As Integer
    
    On Error GoTo Err_handler

    ErrMsg = ""
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQLCode = "delete from dbo.HP_File_Log where FileName = '" & FileName & "'"
    MyAdo.sqlString = strSQLCode
    MyAdo.SQLTextType = sqltext
    iRetCd = MyAdo.Execute
    If iRetCd = -1 Then
        HP_Delete_Log = False
        ErrMsg = "Error executing SQL command [" & MyAdo.sqlString & "]"
        GoTo Exit_Function
    End If

    HP_Delete_Log = True
    
Exit_Function:
    Set MyAdo = Nothing
    Exit Function

Err_handler:
    ErrMsg = Err.Description
    HP_Delete_Log = False
    Resume Exit_Function
    
End Function


Public Function HP_Purge_All_Logs(ZipFile As String, ErrMsg As String) As Boolean
    Dim strSQLCode As String
    Dim MyAdo As clsADO
    Dim iRetCd As Integer
    
    On Error GoTo Err_handler

    ErrMsg = ""
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQLCode = "delete from dbo.HP_File_Log where ZipFileName = '" & ZipFile & "'"
    MyAdo.sqlString = strSQLCode
    MyAdo.SQLTextType = sqltext
    iRetCd = MyAdo.Execute
    If iRetCd = -1 Then
        HP_Purge_All_Logs = False
        ErrMsg = "Error executing SQL command [" & MyAdo.sqlString & "]"
        GoTo Exit_Function
    End If

    HP_Purge_All_Logs = True
    
Exit_Function:
    Set MyAdo = Nothing
    Exit Function

Err_handler:
    ErrMsg = Err.Description
    HP_Purge_All_Logs = False
    Resume Exit_Function
    
End Function


Public Function HP_Create_ACK_File(ByVal OutboundFolder As String, ByVal CnlyCSVFileName As String, ByVal CnlyZipFileName As String, _
                        FullHPACKFileName As String, ErrMsg As String) As Boolean
    
    Dim MyAdo As clsADO
    
    Dim fso As FileSystemObject
    Dim FSOFile As TextStream
    Dim strFilePath As String
    Dim strExportACKText As String
    Dim strExportFileName As String
    Dim strSQL As String
    
    Dim strCnlyACKFileName As String
    Dim strHPACKFileName As String
    
    Dim strOutFile As String
    
    Dim rs As ADODB.RecordSet
    
    Dim iFileNum
    Dim bResult As Boolean
    
    On Error GoTo Err_handler

    
    Set MyAdo = New clsADO
    Set fso = New FileSystemObject
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
    strSQL = "SELECT * from dbo.v_HP_MR_Response_ACK_Export WHERE Filename = '" & CnlyCSVFileName & "'"
    Set rs = MyAdo.OpenRecordSet(strSQL)
    
    FullHPACKFileName = ""
    
    If rs.BOF And True And rs.EOF Then
        ErrMsg = "Error: no data for ACK file"
        GoTo Exit_Function
    Else
        strCnlyACKFileName = Replace(rs.Fields("FileName"), "csv", "ack")
        strHPACKFileName = rs.Fields("ExportFileName")
    End If
    
   
    
    strOutFile = OutboundFolder & strHPACKFileName
    iFileNum = FreeFile()
    
    Open strOutFile For Output As #iFileNum
    rs.MoveFirst
    While Not rs.EOF
        Print #iFileNum, rs("ACK_Export_Txt")
        rs.MoveNext
    Wend
    Close #iFileNum

    If fso.FileExists(strOutFile) Then
        bResult = HP_Log_File(strCnlyACKFileName, strHPACKFileName, CnlyZipFileName, fso.GetFile(strOutFile).DateCreated, _
                              "CSV_IMPORT_ACK", 1, 1, "ACK response file sent to Healthport", ErrMsg)
        If bResult = False Then
            HP_Create_ACK_File = False
            GoTo Exit_Function
        End If
        
    End If

    FullHPACKFileName = strOutFile
    HP_Create_ACK_File = True

Exit_Function:
    Set fso = Nothing
    Set FSOFile = Nothing
    Set rs = Nothing
    Exit Function

Err_handler:
    HP_Create_ACK_File = False
    ErrMsg = Err.Description
    Resume Exit_Function
End Function



Public Function HP_Create_ERR_File(FileName As String, ErrCode As String, ErrMsg As String) As Boolean
    Dim fso As FileSystemObject
    Dim FSOFile As TextStream
    Dim bResult As Boolean
    
    On Error GoTo Err_handler
     
    Set fso = New FileSystemObject
    
    ' opens  file in write mode
    Set FSOFile = fso.OpenTextFile(FileName, 2, True)
    
    FSOFile.WriteLine ("""" & ErrCode & """" & "," & """" & ErrMsg & """")
    FSOFile.Close

    If fso.FileExists(FileName) Then
        bResult = HP_Log_File(FileName, FileName, "", fso.GetFile(FileName).DateCreated, "ERR_RESPONSE", 1, 1, _
                             "Error response file sent to Healthport", ErrMsg)
                             
        HP_Create_ERR_File = True
        
        ' manually copy ERR file FTP site
        Call fso.CopyFile(FileName, "\\ccaintranet.com\DFS-CMS-DS\ftp\Healthport\Outbound\" & fso.GetFileName(FileName))
        
    Else
        HP_Create_ERR_File = False
    End If



Exit_Function:
    Set fso = Nothing
    Set FSOFile = Nothing
    Exit Function

Err_handler:
    ErrMsg = Err.Description
    HP_Create_ERR_File = False
    Resume Exit_Function
End Function



Public Sub Copy_Image_To_DailyScan()
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim fso As New FileSystemObject
    
    Dim strWorkingFolder As String
    Dim strCnlyCSVFile As String
    Dim DailyScanPath As String
    
    Dim strCnlyProvID As String
    Dim strProvFolder As String
    Dim strHPImageFile As String
    Dim strCnlyImageFile As String
    Dim strSQL As String
        
    Dim bResult As Boolean
    
    strCnlyCSVFile = "Y:\Raw\CMS\Healthport\Inbound\Response\HealthPort_0005_102520110847_20111025115957\HealthPort_0005_102520110847_20111025115957.csv"
    DailyScanPath = "Y:\Raw\CMS\Healthport\Inbound\DailyScans\"
    strWorkingFolder = "Y:\Raw\CMS\Healthport\Inbound\Response\HealthPort_0005_102520110847_20111025115957\"
    
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.SQLTextType = sqltext
    
    strSQL = "SELECT il.CnlyProvID, il.ImageName as CnlyImageFile, wk.* " & _
             "from dbo.HP_MR_Response wk " & _
             "join dbo.SCANNING_Image_Log_Tmp il on il.ScannedDt = wk.ScannedDt " & _
             "where wk.ProcessFlag = 1 " & _
             "and wk.AckCode in (1,4) " & _
             "and wk.Filename = '" & fso.GetFileName(strCnlyCSVFile) & "'"
    Set rs = MyAdo.OpenRecordSet(strSQL)

    If rs.BOF = True And rs.EOF = True Then
        MsgBox "Error moving images to Dailyscan.  No row return for CSV file " & strCnlyCSVFile
    Else
        With rs
            .MoveFirst
            While Not .EOF
                strCnlyProvID = rs("CnlyProvID")
                strProvFolder = DailyScanPath & strCnlyProvID
                bResult = CreateFolder(strProvFolder)
                If bResult = False Then
                    Debug.Print "Error: can not create provider folder " & strProvFolder
                    MsgBox "Error: can not create provider folder " & strProvFolder
                    Exit Sub
                End If
                
                strHPImageFile = strWorkingFolder & rs("ImageName")
                strCnlyImageFile = strProvFolder & "\" & rs("CnlyImageFile")
                fso.CopyFile strHPImageFile, strCnlyImageFile, True
                If fso.FileExists(strCnlyImageFile) = False Then
                    Debug.Print "Error: can not copy image file " & strHPImageFile & " to DailyScan folder"
                    MsgBox "Error: can not copy image file " & strHPImageFile & " to DailyScan folder"
                End If
                .MoveNext
            Wend
        End With
    End If


End Sub

Public Function UnZipFile(ZipFileName As String, TargetFolder As String, Optional bCreateSubFolder As Boolean = False) As String

    '---------------------------------------------------------------------------------------
    ' Purpose   : unzips a file to a certain location
    '---------------------------------------------------------------------------------------
    Const PATHWINZIP        As String = "C:\Program Files (x86)\WinZip\"

    Dim fso As New FileSystemObject
    
    Dim ShellStr As String
    Dim strBaseFileName As String
    Dim strNewTargetPath As String
    
    
    On Error GoTo Err_handler

    TargetFolder = FixPath(TargetFolder)

    'This will check if this is the path where WinZip is installed.
    If Dir(PATHWINZIP & "wzunzip.exe") = "" Then
        UnZipFile = "ERROR: can not find wzunzip.exe."
        GoTo Exit_Function
    End If
    
    If bCreateSubFolder = True Then
        strBaseFileName = fso.GetBaseName(ZipFileName)
        strNewTargetPath = TargetFolder & strBaseFileName
    Else
        strNewTargetPath = TargetFolder
    End If
    
    
    'Unzip the zip file in the folder FolderName
    ShellStr = PATHWINZIP & "Wzunzip -e" _
             & " " & Chr(34) & ZipFileName & Chr(34) _
             & " " & Chr(34) & strNewTargetPath & Chr(34)

    
    If Not bShellAndWait(ShellStr) Then Err.Raise 9999

    ' success return empty string
    UnZipFile = ""
    
Exit_Function:
    Set fso = Nothing
    Exit Function

Err_handler:
    ' error return error description
    UnZipFile = Err.Description
    Resume Exit_Function
End Function


Public Function UnZipFile_OLD(ZipFileName As String, TargetFolder As String, Optional bCreateSubFolder As Boolean = False) As String

    '---------------------------------------------------------------------------------------
    ' Purpose   : unzips a file to a certain location
    '---------------------------------------------------------------------------------------
    Const PATHWINZIP        As String = "C:\Program Files (x86)\WinZip\"
    
    Dim fso As New FileSystemObject
    
    Dim ShellStr As String
    Dim strBaseFileName As String
    Dim strNewTargetPath As String
    
    
    On Error GoTo Err_handler

    TargetFolder = FixPath(TargetFolder)

    'This will check if this is the path where WinZip is installed.
    If Dir(PATHWINZIP & "winzip32.exe") = "" Then
        UnZipFile_OLD = "ERROR: Can not find Winzip32.exe."
        GoTo Exit_Function
    End If
    
    
    If bCreateSubFolder = True Then
        strBaseFileName = fso.GetBaseName(ZipFileName)
        strNewTargetPath = TargetFolder & strBaseFileName
    Else
        strNewTargetPath = TargetFolder
    End If
    
    
    'Unzip the zip file in the folder FolderName
    ShellStr = PATHWINZIP & "Winzip32 -min -e" _
             & " " & Chr(34) & ZipFileName & Chr(34) _
             & " " & Chr(34) & strNewTargetPath & Chr(34)

    
    ' Unzip file - VB DOES wait for Winzip
    If Not bShellAndWait(ShellStr, vbNormalFocus) Then Err.Raise 9999

    ' success return empty string
    UnZipFile_OLD = ""


Exit_Function:
    Set fso = Nothing
    Exit Function

Err_handler:
    ' error return error description
    UnZipFile_OLD = Err.Description
    Resume Exit_Function
End Function

Public Sub thieu()
    Dim ErrMsg As String
    Call HP_Move_File(ErrMsg)
    Debug.Print ErrMsg
End Sub
Public Function HP_Move_File(Optional ErrMsg As String = "") As Boolean
    Dim fso As New FileSystemObject
    Dim flder As Folder
    Dim f As file
    
    Dim bResult As Boolean
    Dim bErrFlag As Boolean
    Dim strProcessStep As String
    Dim strDestFile As String
    Dim strLogFile As String
    
    Dim strScanFolder As String
    Dim strDestFolder As String
    Dim strLogFolder As String
    
    Dim strMailSubject As String
    Dim strMailMsg As String
    Dim strMailTo As String
    Dim iLogFile
    
    On Error GoTo Err_handler
    
    strLogFile = ""
    bErrFlag = False
    Close
    
    
    ' move ACK/ERR files
    strProcessStep = "move ACK/ERR files"
    
    strScanFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Inbound\ACK"
    strDestFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Inbound\ACK"
    strLogFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Logs"
    
    Set flder = fso.GetFolder(strScanFolder)
    For Each f In flder.Files
        If UCase(Right(f.Name, 4)) = ".ACK" Or UCase(Right(f.Name, 4)) = ".ERR" Then
            If strLogFile = "" Then
                strLogFile = strLogFolder & "\" & "HP_File_Move " & Format(Now(), "mm-dd-yy hhmmss") & ".log"
                iLogFile = FreeFile()
                Open strLogFile For Output As #iLogFile
            End If
            'Debug.Print f.path
            strDestFile = strDestFolder & "\" & f.Name
            If DateDiff("n", f.DateLastModified, Now()) > 15 Then  ' file not updated for 15 minutes
                Print #iLogFile, "moving file " & f.Path & " to " & strDestFile
                bResult = MoveFile(f.Path, strDestFile, True, ErrMsg)
                If bResult = False Then
                    bErrFlag = True
                    Print #iLogFile, Space(5) & ErrMsg
                End If
            End If
        End If
    Next
    
    ' move BILLING files
    strProcessStep = "move BILLING files"
    
    strScanFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Inbound\Billing"
    strDestFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Inbound\Billing"
    strLogFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Logs"
    
    Set flder = fso.GetFolder(strScanFolder)
    For Each f In flder.Files
        If UCase(Right(f.Name, 4)) = ".ACK" Or UCase(Right(f.Name, 4)) = ".ERR" Then
            If strLogFile = "" Then
                strLogFile = strLogFolder & "\" & "HP_File_Move " & Format(Now(), "mm-dd-yy hhmmss") & ".log"
                iLogFile = FreeFile()
                Open strLogFile For Output As #iLogFile
            End If
            strDestFile = strDestFolder & "\" & f.Name
            If DateDiff("n", f.DateLastModified, Now()) > 15 Then  ' file not updated for 15 minutes
                Print #iLogFile, "moving file " & f.Path & " to " & strDestFile
                bResult = MoveFile(f.Path, strDestFile, True, ErrMsg)
                If bResult = False Then
                    bErrFlag = True
                    Print #iLogFile, Space(5) & ErrMsg
                End If
            End If
        End If
    Next
    
    
    
    ' move response files
    strProcessStep = "move response files"
    
    strScanFolder = "\\ccaintranet.com\dfs-cms-ds\ftp\Healthport\Inbound\Response"
    strDestFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Inbound\Response"
    strLogFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Logs"
    
    Set flder = fso.GetFolder(strScanFolder)
    For Each f In flder.Files
        If UCase(Right(f.Name, 4)) = ".ZIP" Then
            If strLogFile = "" Then
                strLogFile = strLogFolder & "\" & "HP_File_Move " & Format(Now(), "mm-dd-yy hhmmss") & ".log"
                iLogFile = FreeFile()
                Open strLogFile For Output As #iLogFile
            End If
            'Debug.Print f.path
            strDestFile = strDestFolder & "\" & f.Name
            If DateDiff("n", f.DateLastModified, Now()) > 30 Then  ' file not updated for 15 minutes
                Print #iLogFile, "moving file " & f.Path & " to " & strDestFile
                bResult = MoveFile(f.Path, strDestFile, True, ErrMsg)
                If bResult = False Then
                    bErrFlag = True
                    Print #iLogFile, Space(5) & ErrMsg
                End If
            End If
        End If
    Next

    If bErrFlag Then
        strMailSubject = "HP File Move (with Error))"
        strMailMsg = "Error(s) encountered with file move.  Please check the log: <file://" & strLogFile & ">"
        strMailTo = "thieu.le@connolly.com"
        Call Send_Mail(strMailSubject, strMailMsg, strMailTo)
    End If
    HP_Move_File = True
    
Exit_Function:
    Set fso = Nothing
    Set flder = Nothing
    Set f = Nothing
    Close
    Exit Function
    
Err_handler:
    HP_Move_File = False
    ErrMsg = "ERROR#: [" & CStr(Err.Number) & "].   ERROR DESCRIPTION: [" & Err.Description & "] at " & strProcessStep
    If strLogFile <> "" Then
        Print #iLogFile, Space(5) & ErrMsg
    End If
    Resume Exit_Function
End Function

Public Sub HP_Process_Single_RESPONSE_Row_RUN()
    Dim strFolder As String
    Dim strProvID As String
    Dim strImageName As String
    Dim strICN As String
    Dim strInstanceID As String
    
    strFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Healthport\Inbound\Response\HealthPort_0010_012620120857_20120126100004"
    strProvID = "100236"
    strICN = "21023601219002NTA    01"
    strInstanceID = "20120120084754393"
    strImageName = "104094598.pdf"
    
    Call HP_Process_Single_RESPONSE_Row(strFolder, strProvID, strImageName, strICN, strInstanceID)
End Sub

Public Sub test()
    Dim xx As Date
    xx = Now()
    
    
    MsgBox xx
    MsgBox Format(xx, "mm-dd-yyyy")
    If CDate(xx) < Date Then
        MsgBox CDate(xx)
    End If
    
End Sub