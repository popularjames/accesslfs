Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




''' Last Modified: 05/20/2015
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This is the main "automator"
'''     That will do all of the work!
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 05/20/2015 - KD: Changing the file naming to include the letter number in batch in the filename
'''  - 04/22/2015 - KD: added reprint functionality to the processor (just set it to status = 'R' and Held = 0 to have it be reprocessed
'''     without trying to update the status or requiring the claim to be in the correct status / queue..
'''  - 05/20/2014 - Created class
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################



Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetProfileStringA Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long


''##########################################################
''##########################################################
''##########################################################
'' Error enum and other enums
''##########################################################
''##########################################################
''##########################################################

Public Enum ProcessorErrors
    MissingSettingsData = 500
    ProcessorAlreadyRunning = 850
    NoItemsToProcess = 1000
    FileToConvertNotFound = 1001
    JobBeingProcessedAlready = 1003
    BatchInFolderNotFound = 2001
    NoCmdToCall = 3001
    DataIssue = 4001
    FileSystemIssue = 5001
End Enum


''##########################################################
''##########################################################
''##########################################################
'' Events
''##########################################################
''##########################################################
''##########################################################
Public Event ProcessorError(ErrMsg As String, ErrNum As ProcessorErrors, ErrSource As String)

Public Event ProcessorStatusChange(CurBatchId As Long, CurJobId As Long, BatchFile As String, JobFile As String, _
    BatchToType As String, JobToType As String, BatchStatus As String, JobStatus As String, _
    BatchRemaining As Long, JobRemaining As Long)

Public Event Finished()

Public Event Stopped(Reason As String)

Public Event Complete()

Private cblnVerboseLog As Boolean
Private cblnIsWorking As Boolean
Private cblnErrorOccurred As Boolean




'' Ok, so since we are going to run at least 1 process asynchronously
'' we will need a CN object and a Cmd object

Private WithEvents coCN As ADODB.Connection
Attribute coCN.VB_VarHelpID = -1
Private coCmd As ADODB.Command
' QueueFinishedLoading
Private cbQueueFinishedLoading As Boolean


'' Class state properties
''##########################################################
''##########################################################
''##########################################################

Private cbErrorOccurred As Boolean
Private cbFatalErrorOccurred As Boolean

Private clQueueRecordsLoadedThisTime  As Long



    ' THis one is if it's OFFICIALLY running
    ' no public properties
Private cblnRunning As Boolean

Private cblnPaused As Boolean
Private cblnOn As Boolean
Private cstrStopReason As String

Private csPausedAt As String
Private csLastFileProcessed As String
Private clLastIdBeforePaused As Long
Private cblnPauseIdIsBatch As Boolean
Private cobjPausedAtObject As Variant


Private ciMaxTimesRetried As Integer

Private Const csTableName As String = "CONVERT_Jobs"


Private clThisQueueRunId As Long
Private clCurBatch As Long
Private csCurBatchType As String


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property




''##########################################################
''##########################################################
''##########################################################
'' Class state properties
''##########################################################
''##########################################################
''##########################################################
Public Property Get Paused() As Boolean
    Paused = cblnPaused
End Property
Public Property Let Paused(blnPaused As Boolean)
    If cblnPaused = True And blnPaused = False Then
        ' Unpausing it!
        Call UnPause
    End If
    cblnPaused = blnPaused
    If blnPaused = True Then
        cstrStopReason = "Paused"
    Else
        cstrStopReason = ""
    End If
End Property


Public Property Get IsWorking() As Boolean
    IsWorking = cblnIsWorking
End Property
Public Property Let IsWorking(blnIsWorking As Boolean)
    cblnIsWorking = blnIsWorking
End Property




    ''##########################################################
Public Property Get Verbose() As Boolean
    Verbose = cblnVerboseLog
End Property
Public Property Let Verbose(bVerbose As Boolean)
    cblnVerboseLog = bVerbose
End Property

    ''##########################################################
Public Property Get IsOn() As Boolean
    IsOn = cblnOn
End Property
Public Property Let IsOn(blnIsOn As Boolean)
    cblnOn = blnIsOn
    If cblnOn = True Then
        cstrStopReason = ""
            ' if it's not already running then start it up
        If Not cblnRunning Then
            Call StartProcessing
        End If
    
    Else
        cstrStopReason = "Client Turned off"
    End If
End Property

Public Property Get CodeConnString() As String
    CodeConnString = GetSetting("SPROC_CONN_STRING")
End Property


Public Property Get DataConnString() As String
    DataConnString = GetSetting("DATA_CONN_STRING")
End Property

    ''##########################################################
Public Property Get QueueFinishedLoading() As Boolean
    QueueFinishedLoading = cbQueueFinishedLoading
End Property
Public Property Let QueueFinishedLoading(bQueueFinishedLoading As Boolean)
    cbQueueFinishedLoading = bQueueFinishedLoading
End Property



''##########################################################
''##########################################################
''##########################################################
'' General properties
''##########################################################
''##########################################################
''##########################################################
'
Public Property Get CurrentBatchId() As Long
    CurrentBatchId = clCurBatch
End Property
Public Property Let CurrentBatchId(lCurrentBatchId As Long)
    clCurBatch = lCurrentBatchId
End Property


'    ''##########################################################
'Public Property Get CurrentJobId() As Long
'Stop
''    CurrentJobId = coCurJob.JobId
'End Property



    ''##########################################################
Public Property Get ThisQueueRunId() As Long
    ThisQueueRunId = clThisQueueRunId
End Property
Public Property Let ThisQueueRunId(lThisQueueRunId As Long)
    clThisQueueRunId = lThisQueueRunId
End Property

    ''##########################################################
Public Property Get QueueRecordsLoadedThisTime() As Long
    QueueRecordsLoadedThisTime = clQueueRecordsLoadedThisTime
End Property
Public Property Let QueueRecordsLoadedThisTime(lQueueRecordsLoadedThisTime As Long)
    clQueueRecordsLoadedThisTime = lQueueRecordsLoadedThisTime
End Property


    ''##########################################################
Public Property Get CurrentBatchType() As String
    CurrentBatchType = csCurBatchType
End Property
Public Property Let CurrentBatchType(sCurBatchType As String)
    csCurBatchType = sCurBatchType
End Property

''##########################################################
''##########################################################
''##########################################################
'' Business logic type functions
''##########################################################
''##########################################################
''##########################################################

Public Function RunWholeProcess() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".RunWholeProcess"
    
    ' get the next queue run id:
    ThisQueueRunId = GetQueueRunId(False)
    
    LogMessage strProcName, , "Starting Run " & CStr(ThisQueueRunId)
    ' run Load Q Proc
    
    ' So, for whatever reason I can execute the sproc but it always fails and the error message is the dynamic query!?!?
'    Call RunLoadQueue
'
'    ' Run Generate Individuals
'    While InsureSqlIsFinished = False
'        SleepEvents 60
'    Wend
    
    Call GenerateIndividualLetters
    
    ' How about any reprints?
    Call GenerateReprintLetters
    
        
    ' Run Combine Individuals (and generate data)
    Call CombineLettersForMailRoom
    

    
    ' Run Copy To Mail Room
    
    ' Run Reports
    
    
    LogMessage strProcName, , "Finished Run " & CStr(ThisQueueRunId)
    
    
Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CombineLettersForMailRoom() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oSummaryRs As ADODB.RecordSet
Dim oDetailsRs As ADODB.RecordSet
Dim lThisBatch As Long
Dim sBatchType As String

Dim sOutFolder As String
Dim sTempFldr As String
Dim sLetterType As String
Dim sLetterDt As String

        ' This is the current Combined Doc Num's ID
    Dim lCombinedDocNum As Long
        ' how many files copied to the temp fldr (should = oRs.RecordCount)
    Dim lCopyCnt As Long
        ' How many letters in all of the combined letters?
    Dim lTotalLtrCnt As Long
        ' How many batches
    Dim lBatchCnt As Long
        ' How many pages for this particular iniidual letter?
    Dim lThisPageCnt As Long
        ' How many letters in this batch
    Dim lThisBatchLtrCount As Long
        ' How many Pages in this batch (combined doc)
    Dim lThisBatchPageCount As Long

Dim lPctDone As Long
Dim lRsTotal As Long
Dim lErrCnt As Long
Dim lDocCount As Long

Dim sOrigFilePath As String
Dim sCurInstanceId As String
Dim sLastIntanceId As String
Dim SFileName As String
Dim sNewStatus As String
Dim bFirst As Boolean
Dim oStatAdo As clsADO
Dim oLtrInst As clsLetterInstance
Dim oRSInst As ADODB.RecordSet
Dim sFinalPath As String
Dim lDbUpdates As Long
Dim sFileNameOnly As String
Dim sExt As String
Dim sBatchSubFldr As String
Dim saryBatchIds() As String
Dim iBatchCount As Integer
Dim sFileNum As String

    strProcName = ClassName & ".CombineLettersForMailRoom"
    
    Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, 0, , , False, True)
        
    
    sOutFolder = GetSetting("MAILROOM_LETTER_PATH")
    If sOutFolder = "" Then
        Stop
        LogMessage strProcName, "ERROR", "Could not get the mailroom letter path from the settings table"
        RaiseEvent ProcessorError("Could not get the mailroom letter path from the settings table", ProcessorErrors.MissingSettingsData, strProcName)
        GoTo Block_Exit
    End If
    
    sOutFolder = QualifyFldrPath(sOutFolder)
    
    sOutFolder = sOutFolder & QualifyFldrPath(Format(Now(), "yyyy-mm-dd"))
    
    ' Make sure it's empty:
    '' Nope: Not our call: Call DeleteFullFolder(sOutFolder)
    
    CreateFolders (sOutFolder)
    
    ' kd: makefolderhidden (sOutFolder)
    
        ' Basically just get the recordset of batches to print
        ' in the proper order, and don't forget that we are doing 2 "batches"
        ' for those that have > 12 pages (make configurable)
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetCombinationOverview"
        .Parameters.Refresh
        Set oSummaryRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "Problem executing " & .sqlString, .Parameters("@pErrMsg").Value
            RaiseEvent ProcessorError("Problem executing " & .sqlString & " : " & .Parameters("@pErrMsg").Value, DataIssue, strProcName)
            GoTo Block_Exit
        End If
        
    End With


    lRsTotal = oSummaryRs.recordCount
    
    bFirst = True
    
'    Set oWordApp = New Word.Application
    sTempFldr = GetUserTempDirectory()
    
    iBatchCount = -1
        
            'This returns the first recordset for the Combination routine
            '- this "outer" loop goes over the batches, the 'inner loop'
            '(from usp_LETTER_Automation_GetCombinationDetails )
            'Loops over each instanceId that needs to be in this combined document (or batchId)
    While Not oSummaryRs.EOF
        lBatchCnt = lBatchCnt + 1
        iBatchCount = iBatchCount + 1
        lThisBatch = oSummaryRs("BatchId").Value
        Me.CurrentBatchId = lThisBatch
        sBatchType = oSummaryRs("BatchType").Value
        
        ReDim Preserve saryBatchIds(iBatchCount)
        saryBatchIds(UBound(saryBatchIds)) = CStr(lThisBatch)
        
        Me.CurrentBatchType = sBatchType
        sLetterType = oSummaryRs("LetterType").Value
        sLetterDt = CStr(oSummaryRs("LetterReqDt").Value)
        
        
        ' Get the Instanceid details for this batch of letters
        With oAdo
            .sqlString = "usp_LETTER_Automation_GetCombinationDetails"
            .Parameters.Refresh
            .Parameters("@pBatchId") = lThisBatch
            .Parameters("@pBatchType") = sBatchType
            Set oDetailsRs = .ExecuteRS
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                Stop
                LogMessage strProcName, "ERROR", "Problem executing " & .sqlString, .Parameters("@pErrMsg").Value
                RaiseEvent ProcessorError("Problem executing " & .sqlString & " : " & .Parameters("@pErrMsg").Value, DataIssue, strProcName)
                GoTo NextBatch
            End If
                ' get this guy's ID
                ' KD: NOTE: this has changed a bit. We are no longer going to use this parameter we'll be getting it from the detail recordset..
            lCombinedDocNum = .Parameters("@pCombinedDocNum").Value
            
        End With
        
        lPctDone = (lBatchCnt / lRsTotal) * 100
        Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, lPctDone, lThisBatch)

            ' Get all of the individual letters that need to be combined..
'        If CopyIndividualLtrsToTempFldr(oDetailsRs, "LetterPath", sTempFldr, lCopyCnt) = False Then
'            LogMessage strProcName, "ERROR", "There was a problem copying the individual letters to the temp folder for combining"
'            Stop
'        End If
'        If lCopyCnt <> oDetailsRs.recordCount Then
'            LogMessage strProcName, "ERROR", "There was a problem copying the individual letters to the temp folder for combining", "No letters copied out of " & CStr(oDetailsRs.recordCount)
'            Stop
'        End If
        
        '' KD: Add copy files to out folder here ()
        sBatchSubFldr = Format(lThisBatch, "0000") & "\"
        CreateFolders (sOutFolder & sBatchSubFldr)
        
        ' KD: make the batchsubfldr hidden to not confuse with the legacy process as well as to keep these offical print file copies safe
        Call SetFolderHidden(sOutFolder & sBatchSubFldr)

        While Not oDetailsRs.EOF
            Set oLtrInst = New clsLetterInstance
            ' Load our object so we have all the details for this instance
            ' (including page count)
            With oLtrInst
                Set oRSInst = GetInstanceIdRS(oDetailsRs("InstanceID").Value)
                If .LoadFromRS(oRSInst) = False Then
                    LogMessage strProcName, "ERROR", "There was a problem with instance: '" & Nz(oDetailsRs("InstanceID").Value) & "'", "Could not load object from ID"
Stop
                    RaiseEvent ProcessorError("There was a problem with instance: '" & Nz(oDetailsRs("InstanceID").Value) & "' - Could not load object from Id", DataIssue, strProcName)
                    GoTo NextBatch
                End If
                .BatchID = lThisBatch
            End With
            
            lThisBatchLtrCount = lThisBatchLtrCount + 1
            sOrigFilePath = oDetailsRs("LetterPath").Value

            Call PathInfoFromPath(sOrigFilePath, sFileNameOnly, , sExt)
            sFileNameOnly = sFileNameOnly & "." & sExt
'Stop
            lCombinedDocNum = oDetailsRs("CombinedDocNum").Value
            sFileNum = Format(lCombinedDocNum, "00000") & "_" & Format(oDetailsRs("LetterNumInCombinedDoc").Value, "00000") & "_"
            
            sCurInstanceId = oDetailsRs("InstanceId").Value
            
            SFileName = GetFileName(sOrigFilePath)
            lThisPageCnt = Nz(oDetailsRs("PageCount").Value, 0)
            lThisBatchLtrCount = lThisBatchLtrCount + 1
            sNewStatus = "P"    ' optimistic thinking...
    
            If FileExists(sOrigFilePath) = False Then
                ' mark this as an error
                ' note that this will throw off my page / letter count for combined docs
                ' so I'll have to deal with that later
                LogMessage strProcName, "ERROR", "File already exists", sOrigFilePath
                Stop
                RaiseEvent ProcessorError("File already exists: " & sOrigFilePath, FileSystemIssue, strProcName)
                Stop
            End If
    
            ' Just copy it to the mail room location
            
            If CopyFile(sOrigFilePath, sOutFolder & sBatchSubFldr & sFileNum & sFileNameOnly, False) = False Then
                LogMessage strProcName, "ERROR", "CopyFile Failed from: " & sOrigFilePath, sOutFolder & sBatchSubFldr & sFileNum & sFileNameOnly
                Stop
                RaiseEvent ProcessorError("CopyFile Failed from: " & sOrigFilePath & " TO: " & sOutFolder & sBatchSubFldr & sFileNum & sFileNameOnly, FileSystemIssue, strProcName)
            End If
'Stop

            Call oLtrInst.UpdateStaticDetails(sOutFolder & sBatchSubFldr, lCombinedDocNum, sOutFolder & sBatchSubFldr & sFileNum & sFileNameOnly)
NextDetail:
            sLastIntanceId = sCurInstanceId
            oDetailsRs.MoveNext
            sFileNum = ""
        Wend

        bFirst = True
        lThisBatchLtrCount = 0
        
        ' For now this is how I'm going to update the status
        Set oStatAdo = New clsADO
        With oStatAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_LETTER_Automation_AdvanceQueue_FromCombined"
            .Parameters.Refresh
            .Parameters("@pBatchType") = sBatchType
            .Parameters("@pBatchId") = lThisBatch
            .Parameters("@pQueueRunId") = Me.ThisQueueRunId
            .Parameters("@pNextStatus") = sNewStatus
            
            .Execute
            
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                LogMessage strProcName, "ERROR", "Problem executing: " & .sqlString, .Parameters("@pErrMsg").Value
                Stop
                RaiseEvent ProcessorError("Problem executing: " & .sqlString & " : " & .Parameters("@pErrMsg").Value, FileSystemIssue, strProcName)

            End If

            lDbUpdates = lDbUpdates + Nz(.Parameters("@pRecordsAffected").Value, 0)

            '            .SQLTextType = SQLtext
            '            .sqlString = "UPDATE PQ SET Status = 'P' FROM Letter_Print_Queue PQ " & _
            '                " INNER Join (    SELECT PrintQueueId FROM    Letter_Print_Queue Q " & _
            '                "    INNER Join   LETTER_Automation_Static_Details SD ON Q.InstanceId = SD.InstanceId " & _
            '                "    Where Q.BatchID = " & CStr(lThisBatch) & " AND CASE WHEN SD.PageCount < 12 THEN 'Regular Batch' ELSE 'Manual Batch' END = '" & sBatchType & "' " & _
            '                "   Group BY  PrintQueueId,      BatchId,   CASE WHEN SD.PageCount < 12 THEN 'Regular Batch' ELSE 'Manual Batch' END " & _
            '                " ) As D ON PQ.PrintQueueId = D.PrintQueueID"
            '            .Execute
        End With
        
        
NextBatch:
        oSummaryRs.MoveNext
    Wend
    
    ' Update the statistics table: LETTER_Automation_Processor_Run_Hist
    ' and LETTER_Automation_Processor_Status
    
 

    '' Ok, now we need to get the data xml files ready for the mail room
    If iBatchCount > -1 Then
        Call ExportXmlForMailRoom(saryBatchIds)
    End If
    
    
Block_Exit:
    ' Clean up any word objects..
    
    Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, 100, , , True, False, lBatchCnt)
    If Not oSummaryRs Is Nothing Then
        If oSummaryRs.State = adStateOpen Then oSummaryRs.Close
        Set oSummaryRs = Nothing
    End If
    If Not oDetailsRs Is Nothing Then
        If oDetailsRs.State = adStateOpen Then oDetailsRs.Close
        Set oDetailsRs = Nothing
    End If

    Set oAdo = Nothing

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CombineLettersForMailRoom_LEGACY() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oSummaryRs As ADODB.RecordSet
Dim oDetailsRs As ADODB.RecordSet
Dim lThisBatch As Long
Dim sBatchType As String

Dim oWordApp As Word.Application
Dim oCombinedDoc As Word.Document
Dim sOutFolder As String
Dim sTempFldr As String
Dim sLetterType As String
Dim sLetterDt As String

        ' This is the current Combined Doc Num's ID
    Dim lCombinedDocNum As Long
        ' how many files copied to the temp fldr (should = oRs.RecordCount)
    Dim lCopyCnt As Long
        ' How many letters in all of the combined letters?
    Dim lTotalLtrCnt As Long
        ' How many batches
    Dim lBatchCnt As Long
        ' How many pages for this particular iniidual letter?
    Dim lThisPageCnt As Long
        ' How many letters in this batch
    Dim lThisBatchLtrCount As Long
        ' How many Pages in this batch (combined doc)
    Dim lThisBatchPageCount As Long

Dim lPctDone As Long
Dim lRsTotal As Long


Dim lErrCnt As Long
Dim lDocCount As Long

Dim sOrigFilePath As String
Dim sCurInstanceId As String
Dim sLastIntanceId As String
Dim SFileName As String
Dim sNewStatus As String
Dim bFirst As Boolean
Dim oStatAdo As clsADO
Dim oLtrInst As clsLetterInstance
Dim oRSInst As ADODB.RecordSet
Dim sFinalPath As String
Dim lDbUpdates As Long

    strProcName = ClassName & ".ombineLettersForMailRoom"
    
    Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, 0, , , False, True)
        
    
    sOutFolder = GetSetting("MAILROOM_LETTER_PATH")
    If sOutFolder = "" Then
        Stop
    End If
    
    sOutFolder = QualifyFldrPath(sOutFolder)
    
    sOutFolder = sOutFolder & QualifyFldrPath(Format(Now(), "yyyy-mm-dd"))
    
    ' Make sure it's empty:
    '' Nope: Not our call: Call DeleteFullFolder(sOutFolder)
    
    CreateFolders (sOutFolder)
    
        ' Basically just get the recordset of batches to print
        ' in the proper order, and don't forget that we are doing 2 "batches"
        ' for those that have > 12 pages (make configurable)
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetCombinationOverview"
        .Parameters.Refresh
        Set oSummaryRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        
    End With


    lRsTotal = oSummaryRs.recordCount
    
    bFirst = True
    
    Set oWordApp = New Word.Application
    sTempFldr = GetUserTempDirectory()

        ' This loop is the
    While Not oSummaryRs.EOF
        lBatchCnt = lBatchCnt + 1
        lThisBatch = oSummaryRs("BatchId").Value
        Me.CurrentBatchId = lThisBatch
        sBatchType = oSummaryRs("BatchType").Value
        Me.CurrentBatchType = sBatchType
        sLetterType = oSummaryRs("LetterType").Value
        sLetterDt = CStr(oSummaryRs("LetterReqDt").Value)
        
        
        ' Get the Instanceid details for this batch of letters
        With oAdo
            .sqlString = "usp_LETTER_Automation_GetCombinationDetails"
            .Parameters.Refresh
            .Parameters("@pBatchId") = lThisBatch
            .Parameters("@pBatchType") = sBatchType
            Set oDetailsRs = .ExecuteRS
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                Stop
            End If
                ' get this guy's ID
            lCombinedDocNum = .Parameters("@pCombinedDocNum").Value
            
        End With
        
        lPctDone = (lBatchCnt / lRsTotal) * 100
        Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, lPctDone, lThisBatch)

            ' Get all of the individual letters that need to be combined..
        If CopyIndividualLtrsToTempFldr(oDetailsRs, "LetterPath", sTempFldr, lCopyCnt) = False Then
            LogMessage strProcName, "ERROR", "There was a problem copying the individual letters to the temp folder for combining"
            Stop
        End If
        If lCopyCnt <> oDetailsRs.recordCount Then
            LogMessage strProcName, "ERROR", "There was a problem copying the individual letters to the temp folder for combining", "No letters copied out of " & CStr(oDetailsRs.recordCount)
            Stop
        End If
        
        '' KD: Add copy files to out folder here ()

        While Not oDetailsRs.EOF
            Set oLtrInst = New clsLetterInstance
            ' Load our object so we have all the details for this instance
            ' (including page count)
            With oLtrInst
                Set oRSInst = GetInstanceIdRS(oDetailsRs("InstanceID").Value)
                If .LoadFromRS(oRSInst) = False Then
                    LogMessage strProcName, "ERROR", "There was a problem with instance: '" & Nz(oDetailsRs("InstanceID").Value) & "'", "Could not load object from ID"
Stop
                End If
                .BatchID = lThisBatch
            End With
            
            lThisBatchLtrCount = lThisBatchLtrCount + 1
            sOrigFilePath = oDetailsRs("LetterPath").Value
            sCurInstanceId = oDetailsRs("InstanceId").Value
            
            SFileName = GetFileName(sOrigFilePath)
            lThisPageCnt = Nz(oDetailsRs("PageCount").Value, 0)
            lThisBatchLtrCount = lThisBatchLtrCount + 1
            sNewStatus = "P"    ' optimistic thinking...
    
            If FileExists(sOrigFilePath) = False Then
                ' mark this as an error
                ' note that this will throw off my page / letter count for combined docs
                ' so I'll have to deal with that later
                Stop
            End If
    
    
            If bFirst = True Then
                ' new word document
                If Not oCombinedDoc Is Nothing Then
                    Stop    ' what the?
                End If
                Set oCombinedDoc = oWordApp.Documents.Open(sTempFldr & SFileName, False, True, False, , , , , , , , False)
                ' we are going to rename this one so we know it's a merged document
                
'                sFinalPath = sOutFolder & Format(lBatchCnt, "0###") & "_" & sLetterType & "_" & Replace(sBatchType, " ", "_") & "_MergedDoc.doc"
'                sFinalPath = sOutFolder & Format(lThisBatch, "0###") & "_" & Replace(sBatchType, " ", "_") & sLetterType & "_MergedDoc.doc"

                lDocCount = 1

                sFinalPath = sOutFolder & Format(lThisBatch, "0###") & "_" & Replace(sBatchType, " ", "_") & sLetterType & "_MergedDoc_" & Format(lDocCount, "0###") & ".doc"
                While FileExists(sFinalPath) = True
                    lDocCount = lDocCount + 1
                    sFinalPath = sOutFolder & Format(lThisBatch, "0###") & "_" & Replace(sBatchType, " ", "_") & sLetterType & "_MergedDoc_" & Format(lDocCount, "0###") & ".doc"
                Wend
                
                'Call DeleteFile(sFinalPath, False)
                    ' Probably don't need to do this but it's only on the first one so not hurting much
                If UnlinkWordFields(oWordApp, oCombinedDoc) = False Then
                    LogMessage strProcName, "ERROR", "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!"
    '                Call ErrorCallStack_Add(0, "There was a problem unlinking Word Fields - bar codes may be incorrect?!?!", strProcName, , , , oRS("InstanceID").Value, oRS("LetterType").Value)
                End If
                
                '2014:04:29:JS Addded this here because InsertWordDocAtStartOfCurrentDoc only does it for the inserted documents now.
                With oCombinedDoc
                    .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                    .Sections(1).Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                    .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                    .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
                    .Repaginate
                End With

                oCombinedDoc.SaveAs2 (sFinalPath)
                bFirst = False
            Else    ' not the first letter in this batch
             ' not a new letter type so we just append this to our current mergedDoc
                ' the next row should always be a different instance but lets check to make sure..
                If sLastIntanceId <> sCurInstanceId Then
                        '' - use Word to add the next document to the end of that first one.
    '                If bForward = True Then
    '                    If InsertWordDocAtEndOfCurrentDoc(oCombinedDoc, sTempFldr & sFileName, lTtlPages) = False Then
    '
    ''                        Call ErrorCallStack_Add(0, "There was a problem adding a letter to the end of the merged document", strProcName, sFileName, , , oRS("InstanceID").Value, oRS("LetterType").Value)
    '                        lErrCnt = lErrCnt + 1
    '                        sNewStatus = "E"
    '                    Else
    '                        sNewStatus = "P"
    '                        oMergedDoc.Save
    '                    End If
    '
    '                Else
                        If InsertWordDocAtStartOfCurrentDoc(oCombinedDoc, sTempFldr & SFileName) = False Then
        
    '                        Call ErrorCallStack_Add(0, "There was a problem adding a letter to the beginning of the merged document", strProcName, sFileName, , , oRS("InstanceID").Value, oRS("LetterType").Value)
                            lErrCnt = lErrCnt + 1
                            sNewStatus = "E"
                        Else
                            sNewStatus = "P"
                            If lThisBatchLtrCount Mod 50 = 0 Then
                                oCombinedDoc.Save
                            End If
                        End If
    
    '                End If
    '''                LogMessage strProcName, "EFFICIENCY TESTING", "Finished InsertWordDoc," & sCurLetterType, CStr(lCnt / 2) & ", Total Pages now: " & CStr(lTtlPages)
                Else
                    Stop ' this should have been a different instanceid
                End If
            End If
            Call oLtrInst.UpdateStaticDetails(sFinalPath, lCombinedDocNum)
NextDetail:
            sLastIntanceId = sCurInstanceId
            oDetailsRs.MoveNext
        Wend

        ' When we get here we should be done with the current combined document and we need to save and close it
        oCombinedDoc.Save
        oCombinedDoc.Close True
        Set oCombinedDoc = Nothing
        bFirst = True
        lThisBatchLtrCount = 0
        
        ' For now this is how I'm going to update the status
        Set oStatAdo = New clsADO
        With oStatAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_LETTER_Automation_AdvanceQueue_FromCombined"
            .Parameters.Refresh
            .Parameters("@pBatchType") = sBatchType
            .Parameters("@pBatchId") = lThisBatch
            .Parameters("@pQueueRunId") = Me.ThisQueueRunId
            .Parameters("@pNextStatus") = sNewStatus
            
            .Execute
            
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                Stop
            End If

            lDbUpdates = lDbUpdates + Nz(.Parameters("@pRecordsAffected").Value, 0)

            '            .SQLTextType = SQLtext
            '            .sqlString = "UPDATE PQ SET Status = 'P' FROM Letter_Print_Queue PQ " & _
            '                " INNER Join (    SELECT PrintQueueId FROM    Letter_Print_Queue Q " & _
            '                "    INNER Join   LETTER_Automation_Static_Details SD ON Q.InstanceId = SD.InstanceId " & _
            '                "    Where Q.BatchID = " & CStr(lThisBatch) & " AND CASE WHEN SD.PageCount < 12 THEN 'Regular Batch' ELSE 'Manual Batch' END = '" & sBatchType & "' " & _
            '                "   Group BY  PrintQueueId,      BatchId,   CASE WHEN SD.PageCount < 12 THEN 'Regular Batch' ELSE 'Manual Batch' END " & _
            '                " ) As D ON PQ.PrintQueueId = D.PrintQueueID"
            '            .Execute
        End With
        
        
NextBatch:
        oSummaryRs.MoveNext
    Wend
    
    ' Update the statistics table: LETTER_Automation_Processor_Run_Hist
    ' and LETTER_Automation_Processor_Status
    
 
    '' deletefullfolder(    sTempFldr )
    
    '' Ok, now we need to get the data xml files ready for the mail room
'    Call ExportXmlForMailRoom
    
    
Block_Exit:
    ' Clean up any word objects..
    If Not oWordApp Is Nothing Then
        If oWordApp.Documents.Count > 0 Then
            For Each oCombinedDoc In oWordApp.Documents
                oCombinedDoc.Close False
            Next
        End If
        oWordApp.Quit
        Set oWordApp = Nothing
    End If
    Call UpdateProcessorStatus("Combine Batches", Me.ThisQueueRunId, 100, , , True, False, lBatchCnt)
    If Not oSummaryRs Is Nothing Then
        If oSummaryRs.State = adStateOpen Then oSummaryRs.Close
        Set oSummaryRs = Nothing
    End If
    If Not oDetailsRs Is Nothing Then
        If oDetailsRs.State = adStateOpen Then oDetailsRs.Close
        Set oDetailsRs = Nothing
    End If
'    CloseAdoRs (oSummaryRS)
'    CloseAdoRs (oDetailsRs)
    Set oAdo = Nothing
'Dim oCombinedDoc As Word.Document
'Dim oWordApp As Word.Application
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetInstanceIdRS(sInstanceId As String) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".GetInstanceIdRs"
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = DataConnString
        .CursorLocation = adUseClient
        .Open
    End With
    
    Set oRs = New ADODB.RecordSet
    With oRs
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        
        Set .ActiveConnection = oCn
        .Open "SELECT * FROM LETTER_Automation_Static_Details WHERE InstanceId = '" & sInstanceId & "' AND LetterBatchId = " & CStr(Me.CurrentBatchId)
        Set .ActiveConnection = Nothing
    End With
    
    
Block_Exit:
    Set GetInstanceIdRS = oRs
    oCn.Close
    Set oCn = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ExportXmlForMailRoom(saryBatchIds() As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oDetailsRs As ADODB.RecordSet
Dim sOutFolder As String
Dim sXmlFileName As String
Dim oFso As Scripting.FileSystemObject
Dim oTxt As Scripting.TextStream
Dim sXmlSummary As String
Dim sXmlDetails As String
Dim sBatchIds As String

    strProcName = ClassName & ".ExportXmlForMailRoom"

    Set oFso = New Scripting.FileSystemObject
    sBatchIds = MultipleValuesToXml("LetterBatchId", saryBatchIds)
'Stop
    sOutFolder = GetSetting("MAILROOM_DATA_PATH")
    If sOutFolder = "" Then
        Stop
    End If
    
    sOutFolder = QualifyFldrPath(sOutFolder)
    
    sOutFolder = sOutFolder & QualifyFldrPath(Format(Now(), "yyyy-mm-dd"))
    

    CreateFolders (sOutFolder)

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetXMLOutputDetails"
        .Parameters.Refresh
        .Parameters("@pBatchidList") = sBatchIds
        Set oDetailsRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
'        sXmlSummary = .Parameters("@pXmlSummaryOut").Value
    End With
    
    If oDetailsRs.EOF And oDetailsRs.BOF Then
        ' nothing to do
        GoTo Block_Exit
    End If
    If oDetailsRs.recordCount < 1 Then
        ' yup, nothing to do
        GoTo Block_Exit
    End If
    
    sXmlFileName = Format(Me.ThisQueueRunId, "0###") & "_Details_" & Format(Now(), "yyyymmdd_hhnnss") & ".xml"
    
    Set oTxt = oFso.CreateTextFile(sOutFolder & sXmlFileName, True, True)
    oTxt.Write Nz(oDetailsRs(0).Value, "")
    oTxt.Close

            '    sXmlFileName = Replace(sXmlFileName, "Summary", "Details")
            '
            '    Set oTxt = oFso.CreateTextFile(sOutFolder & sXmlFileName, True, True)
            '    oTxt.Write oDetailsRs(0).Value
            '    oTxt.Close
            '
            '
    
    ExportXmlForMailRoom = True
    
Block_Exit:
    If Not oDetailsRs Is Nothing Then
        If oDetailsRs.State = adStateOpen Then oDetailsRs.Close
        Set oDetailsRs = Nothing
    End If
    Set oTxt = Nothing
    Set oAdo = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GenerateIndividualLetters() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oTmpltRs As ADODB.RecordSet
Dim dctLtrTemplates As Scripting.Dictionary
Dim lRowsAffected As Long
Dim sTmpFolder As String
Dim oWordApp As Word.Application

    strProcName = ClassName & ".GenerateIndividualLetters"
    Call UpdateProcessorStatus("Mail Merge", Me.ThisQueueRunId, 0, , , , True)
    
        '' First, release anything that is time sensitive and about to expire...
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ReleaseTimeSensitiveLtrs"
        .Parameters.Refresh
        .Parameters("@pAccountId") = 0  ' 0 for all accounts in Letter_Type table..
        .CurrentConnection.CommandTimeout = 300
        .cmd.CommandTimeout = 300
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "There was an error releasing time sensitive letters in " & .sqlString, .Parameters("@pErrMsg").Value
            RaiseEvent ProcessorError("There was an error releasing time sensitive letters in " & .sqlString, DataIssue, strProcName & "." & .sqlString)
            ' note, we will continue to process other items
        End If

    End With
  
        
        '' Next, generate instanceid's
        '' for the 'released' (held = 0) ones in the queue
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AssignInstanceIds"
        .CurrentConnection.CommandTimeout = 300
        .cmd.CommandTimeout = 30
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Parameters("@pAccountId") = 0  ' 0 for all accounts in Letter_Type table..
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "There was an error assigning instance ids in " & .sqlString, .Parameters("@pErrMsg").Value
            RaiseEvent ProcessorError("There was an error assigning instance ids in " & .sqlString, DataIssue, strProcName & "." & .sqlString)
            ' ok, if we don't have any instanceids then there's nothing we can do - therefore we actually
            ' CAN continue to process because if something was fixed it may be sitting there waiting to be generated..
        End If

    End With


    
    '         EXEC usp_LETTER_Automation_AddToLetterTables @pErrMsg OUT
    ' Add them into the Letter tables
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AddToLetterTables"
        .CurrentConnection.CommandTimeout = 300
        .cmd.CommandTimeout = 30
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "There was an error adding letters to the Letter Hdr and Dtl tables in " & .sqlString, .Parameters("@pErrMsg").Value
            
            RaiseEvent ProcessorError("There was an error adding letters to the Letter Hdr and Dtl tables in " & .sqlString, DataIssue, strProcName & "." & .sqlString)
            ' ok, if we don't have any instanceids then there's nothing we can do - therefore we actually
            ' CAN continue to process because if something was fixed it may be sitting there waiting to be generated..
        End If
    End With
    ' Set up our batches
        '' Next, generate batchids
        '' for the 'released' (held = 0) ones in the queue
        '' This will return a recordset of unique BatchId / LetterTypes
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AssignBatchIds"
        .CurrentConnection.CommandTimeout = 300
        .cmd.CommandTimeout = 30
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Parameters("@pAccountId") = 0  ' 0 for all accounts in Letter_Type table..
        Set oTmpltRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "There was an error assigning batch ids in " & .sqlString, .Parameters("@pErrMsg").Value
            RaiseEvent ProcessorError("There was an error assigning batch ids in " & .sqlString, DataIssue, strProcName & "." & .sqlString)
            ' ok, if we don't have any instanceids then there's nothing we can do - therefore we actually
            ' CAN continue to process because if something was fixed it may be sitting there waiting to be generated..
        End If

    End With
    
        ' Because this sp returns 2 recordsets, we are going to capture the first here:
        ' the first one is going to be what we use to get the templates
        ' and we'll build a dictionary out of it (if we need it, I don't think we do actually)\
        
    '' Copy templates to a temp work folder:
    If CopyTemplatesToTempWorkFldr(oTmpltRs, sTmpFolder, dctLtrTemplates) = False Then
        Stop
            LogMessage strProcName, "ERROR", "There was an error copying the template to a temp folder", sTmpFolder
            RaiseEvent ProcessorError("There was an error copying the template to a temp folder: " & sTmpFolder, FileToConvertNotFound, strProcName)
            ' ok, if we don't have any instanceids then there's nothing we can do - therefore we actually
            ' CAN continue to process because if something was fixed it may be sitting there waiting to be generated..
    End If
    
    ' and advance the oRs to the next one
    Set oRs = oTmpltRs.NextRecordset
    
    '' Now do the individual mail merges:
    If PerformIndividualMailMerges(oWordApp, oRs, dctLtrTemplates, sTmpFolder, lRowsAffected) = False Then
        Stop
            LogMessage strProcName, "ERROR", "There was an error during an individual mail merge "
            RaiseEvent ProcessorError("There was an error during an individual mail merge ", DataIssue, strProcName)
            ' ok, if we don't have any instanceids then there's nothing we can do - therefore we actually
            ' CAN continue to process because if something was fixed it may be sitting there waiting to be generated..
    End If
    

    If Not oWordApp Is Nothing Then
        Dim oDoc As Word.Document
        For Each oDoc In oWordApp.Documents
            oDoc.Close False
        Next
        oWordApp.Quit
        Set oWordApp = Nothing
    End If
    
    If Not oTmpltRs Is Nothing Then
        If oTmpltRs.State = adStateOpen Then oTmpltRs.Close
        Set oTmpltRs = Nothing
    End If
    
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If

'    Set rsLetterConfig = GetLetterConfigDetails()
'
'    If rsLetterConfig.RecordCount = 0 Then
'        strErrMsg = "ERROR: Letter configuration parameters is missing"
'        GoTo Block_Err
'    ElseIf rsLetterConfig.RecordCount > 1 Then
'        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
'        GoTo Block_Err
'    Else
'        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
'        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
'        gbVerboseLogging = IIf(rsLetterConfig("VerboseLogging").Value = 0, False, True)
'    End If
'

    ' now what?
    ' now we need to combine them into single documents to send off
    '
'    oRs.MoveFirst
    

    Call UpdateProcessorStatus("Mail Merge", Me.ThisQueueRunId, 100, , , True, False, lRowsAffected)
    
    
    
    GenerateIndividualLetters = True
Block_Exit:
    If Not oTmpltRs Is Nothing Then
        If oTmpltRs.State = adStateOpen Then oTmpltRs.Close
        Set oTmpltRs = Nothing
    End If
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function InsureSqlIsFinished() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO


    strProcName = ClassName & ".InsureSqlIsFinished"
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_Letter_Automation_IsSqlDone"
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Execute
        InsureSqlIsFinished = IIf(.Parameters("@pComplete").Value, True, False)
    End With


Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function RunLoadQueue() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oErr As ADODB.Error
Dim dtSt As Date
Dim lRecordsAdded As Long
Dim sLoadQueueStoredProc As String

    strProcName = ClassName & ".RunWholeProcess"
Debug.Print strProcName

    ' reset our status
    QueueFinishedLoading = False
    sLoadQueueStoredProc = GetSetting("LOAD_Q_SPROC")
    
    Call SetProcessorStatus("Loading Queue")
    
    ' Connect
    ' (on connection complete our sproc is executed)
    If Not coCN Is Nothing Then
        If coCN.State = adStateOpen Then
            coCN.Close
        End If
    End If
    If coCN Is Nothing Then
        Set coCN = New ADODB.Connection
        coCN.CommandTimeout = 3600 ' no timeout seems dangerous huh? I just know the CMS proc takes 20 - 40 minutes so... lets do an hour (3600)
        ' but maybe this should be configurable since CMS is slow right now and it still takes up to 40 minutes..
        coCN.CursorLocation = adUseClient
        coCN.ConnectionTimeout = 3600
    End If


    
    coCN.Open Me.CodeConnString ', , , adAsyncConnect
    
    SleepEvents 1
    
    dtSt = Now()
    
    Do While coCN.State <> adStateOpen
        If DateDiff("s", dtSt, Now()) > 30 Then Exit Do    ' timeout just in case..
        DoEvents
    Loop
    
 
    Set coCmd = New ADODB.Command
    With coCmd
        Set .ActiveConnection = coCN
        .commandType = adCmdStoredProc
        .CommandText = sLoadQueueStoredProc
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = clThisQueueRunId
'        .Execute , , adCmdStoredProc + adExecuteNoRecords + adAsyncExecute
        .Execute
    End With
        
    QueueFinishedLoading = True
    ' The stored proc will be run on the ConnectComplete event of the connection object
    ' Hang out until it's finished..
    Do While QueueFinishedLoading = False
        SleepEvents 120 ' wait 2 minutes.. This is a long proc so...
    Loop
    
    If QueueFinishedLoading = False Or coCN.State <> adStateOpen Then
        Call HandleDbErrors(strProcName)
        GoTo Block_Exit
    End If
 
    ' the number of records we loaded is added upon the coCN.Execute_Completed event because
    ' we close the connection there..
    
    
    
    RunLoadQueue = True
        
Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetQueueRunId(Optional bStartingRun As Boolean) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".GetQueueRunId"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_GetQueueRunId"
        .Parameters.Refresh
'        .Parameters("@pAccountId") = 0  ' 0 not account specific
        If bStartingRun = True Then
            .Parameters("@pStart") = 1
        Else
            .Parameters("@pStart") = 0
        End If
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Call SetErrorOccurred(strProcName, .Parameters("@pErrMsg").Value)
        Else
            GetQueueRunId = CLng("0" & .Parameters("@pQueueRunId").Value)
        End If
    End With


Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function StartProcessing()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".StartProcessing"

'    Runnable_Start
    If Verbose Then LogMessage strProcName, , "Starting to process"

    Start_Runnable
    
    If Verbose Then LogMessage strProcName, , "Finished to processing"


'    '' We want this to be sort of multi-threaded so we are going to use the RUNNABLE.TLB type library
'    '' I THINK we can do that here in MS Access.. I think
'
'   If Not cblnRunning Then
'      cblnRunning = True
'      ' Call the mStart module.  This uses a timer to
'      ' fire the Runnable_Start() implementation,
'      ' which ensures we yield control back to the
'      ' caller before the processing starts.  This
'      ' ensures that the processing runs asynchronously
'      ' to the client.  Easy!!!
'      modMulti_Thread_Support.Start Me
'   Else
'      ' Just checking....
'      FireErrorStr ProcessorAlreadyRunning, "Processor is already running, can't start again", strProcName
'   End If


Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub StopProcessing()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".StopProcessing"

    cblnPaused = False
    cblnRunning = False
    cblnOn = True
    LogMessage strProcName, , "User stopped processing"

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear

    GoTo Block_Exit
End Sub

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function PauseProcessing() As Boolean
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".PauseProcessing"

    cblnRunning = False
    LogMessage strProcName, , "User Paused processing"


Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CancelProcessing() As Boolean
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".CancelProcessing"
    cstrStopReason = "Client Canceled"
    
    cblnOn = False
    cblnRunning = False
    cblnPaused = False

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




''##########################################################
''##########################################################
''##########################################################
'' Auditing / Setup data / interacting with the cTable object
'' as well as any generically private methods
''##########################################################
''##########################################################
''##########################################################



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetRecordset(sSql As String, Optional sTableName As String = csTableName) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetRecordset"

    Set oAdo = New clsADO
    With oAdo
        '''.ConnectionString = GetConnectString(sTableName)
        Stop
'        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With

    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    Set GetRecordset = oRs

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, sSql
    GoTo Block_Exit
End Function

            '
            '
            '''' ##############################################################################
            '''' ##############################################################################
            '''' ##############################################################################
            ''''  Now with locking!
            ''''
            'Private Function GetRecordset(sSql As String, Optional sTableName As String = csTABLENAME) As ADODB.RecordSet
            'On Error GoTo Block_Err
            'Dim strProcName As String
            'Dim oCn As ADODB.Connection
            'Dim oRs As ADODB.RecordSet
            '
            '    strProcName = ClassName & ".GetRecordset"
            '
            '    Set oCn = New ADODB.Connection
            '    With oCn
            '        .ConnectionString = GetConverterConnectionString()
            '        .Open
            '    End With
            '
            '    oCn.BeginTrans
            '    Set oRs = oCn.Execute(sSql)
            '
            '    oCn.CommitTrans
            '
            '
            '    If oRs Is Nothing Then GoTo Block_Exit
            '    If oRs.RecordCount < 1 Then GoTo Block_Exit
            '    Set GetRecordset = oRs
            '
            'Block_Exit:
            '    Set oRs = Nothing
            '    If Not oCn Is Nothing Then
            '        If oCn.State = adStateOpen Then oCn.Close
            '        Set oCn = Nothing
            '    End If
            '    Exit Function
            'Block_Err:
            '    FireError Err, strProcName, sSql
            '    GoTo Block_Exit
            'End Function


            '
            '''' ##############################################################################
            '''' ##############################################################################
            '''' ##############################################################################
            ''''  Now with locking
            ''''
            'Private Function GetRecordsetSP(sSpName As String, Optional sParamString As String = "", Optional sTableName As String = csTABLENAME) As ADODB.RecordSet
            'On Error GoTo Block_Err
            'Dim strProcName As String
            'Dim oCn As ADODB.Connection
            'Dim oCmd As ADODB.Command
            'Dim oRs As ADODB.RecordSet
            'Dim sParams() As String
            'Dim iIdx As Integer
            'Dim sPName As String
            'Dim sPVal As String
            '
            '    strProcName = ClassName & ".GetRecordset"
            '
            '    If sParamString <> "" Then
            '        sParams = Split(sParamString, ",")
            '    End If
            '
            '    Set oCn = New ADODB.Connection
            '    With oCn
            '        .ConnectionString = GetConverterConnectionString()
            '        .Open
            '    End With
            '        '''.ConnectionString = GetConnectString(sTableName)
            '    Set oCmd = New ADODB.Command
            '    With oCmd
            '        .CommandType = adCmdStoredProc
            '        .CommandText = sSpName
            '        If sParamString <> "" Then
            '            For iIdx = 0 To UBound(sParams)
            '                sPName = Split(sParams(iIdx), "=")(0)
            '                sPVal = Split(sParams(iIdx), "=")(1)
            '                .Parameters(sPName) = sPVal
            '            Next
            '        End If
            '        oCn.BeginTrans
            '        Set oRs = .Execute
            '        oCn.CommitTrans
            '    End With
            '
            '    If oRs Is Nothing Then GoTo Block_Exit
            '    If oRs.RecordCount < 1 Then GoTo Block_Exit
            '    Set GetRecordsetSP = oRs
            '
            'Block_Exit:
            '    Set oRs = Nothing
            '    Set oCmd = Nothing
            '    If Not oCn Is Nothing Then
            '        If oCn.State = adStateOpen Then oCn.Close
            '        Set oCn = Nothing
            '    End If
            '    Exit Function
            'Block_Err:
            '    FireError Err, strProcName, sSpName
            '    GoTo Block_Exit
            'End Function
'



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetRecordsetSP(sSpName As String, Optional sParamString As String = "", Optional sTableName As String = csTableName) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sParams() As String
Dim iIdx As Integer
Dim sPName As String
Dim sPVal As String

    strProcName = ClassName & ".GetRecordset"

    If sParamString <> "" Then
        sParams = Split(sParamString, ",")
    End If

    Set oAdo = New clsADO
    With oAdo
        '''.ConnectionString = GetConnectString(sTableName)
Stop
'        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = sSpName
        If sParamString <> "" Then
            For iIdx = 0 To UBound(sParams)
                sPName = Split(sParams(iIdx), "=")(0)
                sPVal = Split(sParams(iIdx), "=")(1)
                .Parameters(sPName) = sPVal
            Next
        End If
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With

    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    Set GetRecordsetSP = oRs

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, sSpName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function ExecuteSQL(sSql As String, Optional sTableName As String = csTableName) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim lRet As Long
Dim sPName As String
Dim sPVal As String

    strProcName = ClassName & ".GetRecordset"


    Set oAdo = New clsADO
    With oAdo
        '''.ConnectionString = GetConnectString(sTableName)
Stop
'        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = sqltext
        .sqlString = sSql

        lRet = .Execute
    End With
    
    
    ExecuteSQL = True
    
Block_Exit:

    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, sSql
    ExecuteSQL = False
    GoTo Block_Exit
End Function





''##########################################################
''##########################################################
''##########################################################
'' Error handling /  stuff that's repeated a lot!
''##########################################################
''##########################################################
''##########################################################
    ''##########################################################
Private Sub HandleDbErrors(sCallingProc As String)
Dim oErr As ADODB.Error

    For Each oErr In coCN.Errors
            LogMessage sCallingProc, "DB ERROR", oErr.Description, oErr.Source
    Next
Set oErr = Nothing
End Sub

    ''##########################################################
Private Sub FireError(oErr As ErrObject, sErrSourceProcName As String, Optional sAdditionalDetails As String, Optional bFatal As Boolean = False)

    cbErrorOccurred = True
    If bFatal = True Then
        cbFatalErrorOccurred = bFatal
    End If
    ReportError oErr, sErrSourceProcName, , sAdditionalDetails
    
    If sAdditionalDetails <> "" Then sAdditionalDetails = vbCrLf & sAdditionalDetails
    
    RaiseEvent ProcessorError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub


    ''##########################################################
Private Sub FireErrorStr(ErrNum As ProcessorErrors, sDesc As String, sSource As String, Optional sAdditionalDetails As String, Optional bFatal As Boolean = False)
Dim oErr As ErrObject

    If bFatal = True Then
        cbFatalErrorOccurred = bFatal
    End If
    
    Set oErr = Err
    With oErr
        .Number = ErrNum
        .Description = sDesc
        .Source = sSource
    End With
    
    Call SendErrorEmail(CStr(ErrNum) & " : " & sDesc & " IN " & sSource & " : " & sAdditionalDetails, CurrentJobId, CurrentBatchId)
    
    FireError oErr, sSource, sAdditionalDetails
End Sub


    ''##########################################################
Private Sub SetErrorOccurred(sSource As String, sDesc As String, Optional sAdditionalInfo As String, Optional bFatal As Boolean = False)
    LogMessage sSource, "ERROR", sDesc, sAdditionalInfo
    cblnErrorOccurred = True
    cbFatalErrorOccurred = bFatal
End Sub

    ''##########################################################
Private Sub SetErrorOccurredObj(sSource As String, oErr As ErrObject, Optional sAdditionalInfo As String, Optional bFatal As Boolean = False)
    LogMessage sSource, "ERROR", oErr.Description, sAdditionalInfo
    cblnErrorOccurred = True
    cbFatalErrorOccurred = bFatal
End Sub

    ''##########################################################
Private Sub RaiseStatusChange()
    Stop
'    RaiseEvent ProcessorStatusChange(coCurBatch.BatchID, coCurJob.JobId, _
        csCurBatchFile, csCurJobFile, coCurBatch.ToType, _
        coCurJob.ToType, csBatchStatus, csJobStatus, BatchRemaining, JobsRemaining)
End Sub




    ''##########################################################
Private Sub SetProcessorStatus(sMsg As String, Optional lRowsAffected As Long = -1)
On Error GoTo Block_Err
Dim strProcName As String
Dim sLogSproc As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetProcessorStatus"
    
    sLogSproc = GetSetting("SetProcessorStatusSproc")
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = sLogSproc
        .Parameters.Refresh
        .Parameters("@pPosition") = sMsg
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        If lRowsAffected > -1 Then
            .Parameters("@pRowsAffected") = lRowsAffected
        End If
        .Execute
    End With
    
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Sub



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub DictAdd(DictionaryToAdd As Scripting.Dictionary, vKeyToAdd As Variant, vItemToAdd As Variant)
On Error GoTo Block_Err
Dim strProcName As String
'Dim oJob As clsJob
Dim oColl As Collection
    strProcName = ClassName & ".DictAdd"
    
    If DictionaryToAdd Is Nothing Then Set DictionaryToAdd = New Scripting.Dictionary

'Debug.Assert CStr("" & vKeyToAdd) <> "20"
    
    If DictionaryToAdd.Exists(vKeyToAdd) = False Then
        Set oColl = New Collection
        oColl.Add vItemToAdd
        DictionaryToAdd.Add vKeyToAdd, oColl
    Else
        If TypeName(DictionaryToAdd.Item(vKeyToAdd)) = "Collection" Then
            Set oColl = DictionaryToAdd.Item(vKeyToAdd)
            oColl.Add vItemToAdd
        Else    ' first one..
            ' maybe something went wrong and it's a clsJob?
            If TypeName(DictionaryToAdd.Item(vKeyToAdd)) = "clsJob" Then
                Set oColl = New Collection
                oColl.Add TypeName(DictionaryToAdd.Item(vKeyToAdd))
                ' make sure it's not the same job:
                If DictionaryToAdd.Item(vKeyToAdd).JobId <> vItemToAdd.JobId Then
                    oColl.Add vItemToAdd
                End If
            Else
                Set oColl = New Collection
                oColl.Add vItemToAdd
            End If
            Set DictionaryToAdd.Item(vKeyToAdd) = oColl

        End If
        
'        Set DictionaryToAdd.Item(vKeyToAdd) = vItemToAdd
    End If

Block_Exit:
    Exit Sub
Block_Err:
    FireError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub


'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################
'
'       Class Init / Term
'
'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################


Private Sub Class_Initialize()
'    Set cdctJobs = New Scripting.Dictionary
'    Set cdctJobsFinished = New Scripting.Dictionary
'    Set cdctJobsFailed = New Scripting.Dictionary
'    Set cdctAllFileBatches = New Scripting.Dictionary
'    Set coCurBatch = New clsBatch
'    Set coCurJob = New clsJob
    
'    MaxTimesRetried = CInt(goSettings.GetSetting("MAX_TIMES_RETRY"))
'    Call Refresh
    
End Sub

Private Sub Class_Terminate()
'    Set cdctJobs = Nothing
'    Set cdctJobsFinished = Nothing
'    Set cdctJobsFailed = Nothing
'    Set cdctAllFileBatches = Nothing
'    Set coCurBatch = Nothing
'    Set coCurJob = Nothing
End Sub

            'Private Sub Runnable_Start()
            '
            '    ProcessAll
            '
            '    RaiseEvent Complete
            'End Sub

Private Sub Start_Runnable()
'Stop
    Set goBOLD_Processor = Me
    RunWholeProcess

    RaiseEvent Complete
    Me.IsOn = False
    cstrStopReason = "Completed"

End Sub
            
            'Private Sub StartAsync()
            ''Dim strProcName As String
            ''
            ''    strProcName = ClassName & ".StartAsync"
            ''
            ''   If Not cblnRunning Then
            ''      cblnRunning = True
            ''      ' Call the mStart module.  This uses a timer to
            ''      ' fire the Runnable_Start() implementation,
            ''      ' which ensures we yield control back to the
            ''      ' caller before the processing starts.  This
            ''      ' ensures that the processing runs asynchronously
            ''      ' to the client.  Easy!!!
            ''      modMulti_Thread_Support.Start Me
            ''   Else
            ''      ' Just checking....
            ''      FireErrorStr ProcessorAlreadyRunning, "Processor is already running, can't start again", strProcName
            ''   End If
            'End Sub
            
            'Private Sub StartNonAsync()
            '   ' Just here to demonstrate what happens if
            '   ' you call a normal AX EXE method
            '   Runnable_Start
            'End Sub
            

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function ShouldWeStopProcessingNow() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".ShouldWeStopProcessingNow"

    ShouldWeStopProcessingNow = cblnPaused Or cblnRunning Or (Not cblnOn)

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    ShouldWeStopProcessingNow = True
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function UnPause() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'Dim oJob As clsJob
Dim oCol As Collection
'Dim oBatch As clsBatch

    strProcName = ClassName & ".UnPause"
    
        '   if we aren't paused, then there's nothing to unpause
    If cblnPaused = False Then GoTo Block_Exit
    
        ' reset our flags
    cblnPaused = False
    cblnRunning = True
    cblnOn = True
    

    Select Case csPausedAt
    Case ClassName & ".ProcessCollection" ' collection: remove all that came BEFORE our last one
            ' then call the proc
        If TypeName(cobjPausedAtObject) = "Collection" Then
            Set oCol = cobjPausedAtObject
Stop
'            Call ProcessCollection(oCol)
        Else
            Stop
        End If
    Case ClassName & ".ProcessAllFilesFoundBatch"
            ' ok, so in this case we have a Batch Object
            ' the ID is of no use to us really since we already have the batch
            ' we do have the last fileprocessed so, we may need to add
            ' an optional param to the function sto 'skip files until' our file..
            ' question is, when using the fso.folders.File method, can we count on them
            ' being alphabetical?
            ' We aren't going to worry about the risk that one or more files have been created
            ' while paused cause I don't really ever see this thing being paused once in production
            ' more for testing / qa (whatever that is! lol!)
        If TypeName(cobjPausedAtObject) = "clsBatch" Then
Stop
'            Set oBatch = cobjPausedAtObject
'            Call ProcessAllFilesFoundBatch(oBatch, csLastFileProcessed)
        End If
    Case ClassName & ".ProcessJob"
        ' It's highly doubtful that it was paused on a single job like this but for good measure (and
        ' because everytime I figure that it's not worth doing something like this, it actually happens)
        
        If TypeName(cobjPausedAtObject) = "clsJob" And cblnPauseIdIsBatch = False Then
Stop
'            Set oJob = cobjPausedAtObject
'            Call ProcessJob(oJob)
        End If
    End Select
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    UnPause = True
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub coCN_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
On Error GoTo Block_Err
Dim strProcName As String
Dim dtSt As Date
Dim sLoadQueueStoredProc As String


    strProcName = ClassName & ".coCN_ConnectComplete"
Debug.Print strProcName


    Do While coCN.State <> adStateOpen
        If DateDiff("s", dtSt, Now()) > 120 Then Exit Do
        DoEvents
    Loop
    dtSt = Now
    sLoadQueueStoredProc = GetSetting("LOAD_Q_SPROC")
    Debug.Print "EXEC " & sLoadQueueStoredProc & " @pQueueRunId = " & CStr(clThisQueueRunId)
    
    SleepEvents 1
'    coCN.Execute "EXEC " & sLoadQueueStoredProc & " @pQueueRunId = " & CStr(clThisQueueRunId), , adCmdStoredProc + adExecuteNoRecords + adAsyncExecute
    
    
    Set coCmd = New ADODB.Command
    With coCmd
        Set .ActiveConnection = coCN
        .commandType = adCmdStoredProc
        .CommandText = sLoadQueueStoredProc
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = clThisQueueRunId
        .Execute , , adCmdStoredProc + adExecuteNoRecords + adAsyncExecute

    End With
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub coCN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.RecordSet, ByVal pConnection As ADODB.Connection)
Dim strProcName As String
    strProcName = ClassName & ".coCN_ExecuteComplete"
    
Debug.Print strProcName
    
    If adStatus <> adStatusOK Then
        Stop
    End If
    
    If coCN.State = adStateExecuting Then
    Stop
    End If
    
    SleepEvents 2
    


    
'    If Not coCmd Is Nothing Then
'        QueueRecordsLoadedThisTime = coCmd.Parameters("@pRecordsLoaded")
'        LogMessage strProcName, , "Loaded " & CStr(QueueRecordsLoadedThisTime) & " claims into the letter queue", Me.ThisQueueRunId
'    Else
'        Call HandleDbErrors(strProcName)
'    End If
    
       Stop
    coCN.Close
    
    Me.QueueFinishedLoading = True
End Sub

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub coCN_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Stop
End Sub


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''



Public Function SetDefaultPrinterToAcrobat(sOrigPrinter, Optional sSetPrinterTo As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oPrinter As Printer

    strProcName = ClassName & ".SetDefaultPrinterToAcrobat"
    
    If sSetPrinterTo = "" Then
        sSetPrinterTo = "Adobe PDF"
    End If
    
    ' First, get the default printer's name so we can return it (and eventually pass it back to this
    ' function to reset..
    
    
    Set oPrinter = Application.Printer
    sOrigPrinter = oPrinter.DeviceName
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 2)

    Set Application.Printer = Application.Printers(sSetPrinterTo)
    
    SetDefaultPrinterToAcrobat = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function SetDefaultPrinterToAcrobatAPI(sOrigPrinter, Optional sSetPrinterTo As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'Dim oPrinter As Printer
Dim lRet As Long

    strProcName = ClassName & ".SetDefaultPrinterToAcrobatAPI"
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 7)
    If sSetPrinterTo = "" Then
        sSetPrinterTo = "Adobe PDF"
    End If
    
    ' First, get the default printer's name so we can return it (and eventually pass it back to this
    ' function to reset..
    
    
'    Set oPrinter = Application.Printer
    'sOrigPrinter = oPrinter.DeviceName
    sOrigPrinter = DefaultPrinterInfo()
    Debug.Print sOrigPrinter
    ' HP_P3015_01_PCL6 on ssconsho01-001 (redirected 2)

    lRet = SetDefaultPrinter(sSetPrinterTo)
    
    SetDefaultPrinterToAcrobatAPI = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function DefaultPrinterInfo() As String
Dim strLPT As String * 255
Dim Result As String
Dim ResultLength As Long
Dim Comma1 As Integer, Comma2 As Integer
Dim Driver As String
Dim Port As String
Dim sPrinter As String

    Call GetProfileStringA("Windows", "Device", "", strLPT, 254)
    
    Result = TrimNull(strLPT)
    ResultLength = Len(Result)

    Comma1 = InStr(1, Result, ",", 1)
    Comma2 = InStr(Comma1 + 1, Result, ",")

'   Gets printer's name
    sPrinter = left(Result, Comma1 - 1)
    DefaultPrinterInfo = sPrinter
'   Gets driver
    Driver = Mid(Result, Comma1 + 1, Comma2 - Comma1 - 1)

'   Gets last part of device line
    Port = Right(Result, ResultLength - Comma2)

    Debug.Print sPrinter
    Debug.Print Driver
    Debug.Print Port

End Function







''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GenerateReprintLetters() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oTmpltRs As ADODB.RecordSet
Dim dctLtrTemplates As Scripting.Dictionary
Dim lRowsAffected As Long
Dim sTmpFolder As String
Dim oWordApp As Word.Application

    strProcName = ClassName & ".GenerateReprintLetters"
    Call UpdateProcessorStatus("Mail Merge Reprint", Me.ThisQueueRunId, 0, , , , True)
  
  
        '' Next, generate instanceid's
        '' for the 'released' (held = 0) ones in the queue
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_UpdateInstanceIdsForReprint"
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Parameters("@pAccountId") = 0  ' 0 for all accounts in Letter_Type table..
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If

    End With


    
    '         EXEC usp_LETTER_Automation_AddToLetterTables @pErrMsg OUT
    ' Add them into the Letter tables
    ' Actually, these should be here already
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AddToLetterTables"
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    ' Set up our batches
        '' Next, generate batchids
        '' for the 'released' (held = 0) ones in the queue
        '' This will return a recordset of unique BatchId / LetterTypes
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = Me.CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_AssignBatchIdsForRepint"
        .Parameters.Refresh
        .Parameters("@pQueueRunId") = Me.ThisQueueRunId
        .Parameters("@pAccountId") = 0  ' 0 for all accounts in Letter_Type table..
        Set oTmpltRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If

    End With
    
        ' Because this sp returns 2 recordsets, we are going to capture the first here:
        ' the first one is going to be what we use to get the templates
        ' and we'll build a dictionary out of it (if we need it, I don't think we do actually)\
        
    '' Copy templates to a temp work folder:
    If oTmpltRs.recordCount = 0 Then
        GenerateReprintLetters = True
        GoTo Block_Exit
    Else
        If CopyTemplatesToTempWorkFldr(oTmpltRs, sTmpFolder, dctLtrTemplates) = False Then
            Stop
        End If
    End If
    
    ' and advance the oRs to the next one
    Set oRs = oTmpltRs.NextRecordset
    
    '' Now do the individual mail merges:
    If PerformIndividualMailMerges(oWordApp, oRs, dctLtrTemplates, sTmpFolder, lRowsAffected, , , True) = False Then
        Stop
    End If
    

    If Not oWordApp Is Nothing Then
        Dim oDoc As Word.Document
        For Each oDoc In oWordApp.Documents
            oDoc.Close False
        Next
        oWordApp.Quit
        Set oWordApp = Nothing
    End If
    
    If Not oTmpltRs Is Nothing Then
        If oTmpltRs.State = adStateOpen Then oTmpltRs.Close
        Set oTmpltRs = Nothing
    End If
    
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If

'    Set rsLetterConfig = GetLetterConfigDetails()
'
'    If rsLetterConfig.RecordCount = 0 Then
'        strErrMsg = "ERROR: Letter configuration parameters is missing"
'        GoTo Block_Err
'    ElseIf rsLetterConfig.RecordCount > 1 Then
'        strErrMsg = "ERROR: more than 1 row of letter configuration parameters returned."
'        GoTo Block_Err
'    Else
'        strOutputLocation = rsLetterConfig("LetterOutputLocation").Value
'        strODCFile = rsLetterConfig("ODC_ConnectionFile").Value
'        gbVerboseLogging = IIf(rsLetterConfig("VerboseLogging").Value = 0, False, True)
'    End If
'

    ' now what?
    ' now we need to combine them into single documents to send off
    '
'    oRs.MoveFirst
    

    Call UpdateProcessorStatus("Mail Merge", Me.ThisQueueRunId, 100, , , True, False, lRowsAffected)
    
    
    
    GenerateReprintLetters = True
Block_Exit:
    If Not oTmpltRs Is Nothing Then
        If oTmpltRs.State = adStateOpen Then oTmpltRs.Close
        Set oTmpltRs = Nothing
    End If
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function