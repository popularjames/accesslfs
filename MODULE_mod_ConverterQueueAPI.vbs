Option Compare Database
Option Explicit



''' Last Modified: 03/10/2015
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 03/10/2015 - KD: Added logging for when code is waiting for a file to complete, but it times out. Also added
'''     optional parameter to specify how many seconds to wait before considering the job as timed out.
'''  - 03/06/2015 - KD: small tweak to have AddJobto queue return false if it times out (when waiting for it to finish)
'''  - 02/27/2015 - KD: Set up for new contract dev (bigsky)
'''  - 10/11/2013 - KD: Added GetConverterConnectionString so we can switch to Big-sky all at once
'''  - 05/29/2013 - Added Cancel batch and added optional param to wait for conversion when adding a record
'''  - 05/08/2012 - updated wait for code
'''  - 04/19/2012 - Created...
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


Private Const ClassName As String = "mod_ConverterQueueAPI"

Public Function GetConverterConnectionString(Optional sNotUsed As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim dtDummyDt As Date
Static sRet As String
Static dLastTimeChecked As Date


    strProcName = ClassName & ".GetConverterConnectionString"
    
    ' we don't want to hit the database every single time..
    ' hence the static variables here..
    If (DateDiff("n", dLastTimeChecked, Now()) >= 30 And sRet <> "") _
        Or (dLastTimeChecked = dtDummyDt) _
        Or sRet = "" Then

            dLastTimeChecked = Now()
    
            Set oAdo = New clsADO
            With oAdo
                .ConnectionString = GetConnectString("v_Code_Database")
                .SQLTextType = StoredProc
                .sqlString = "usp_CONVERTER_Queue_ConnectString"
                .Parameters.Refresh
                .Execute
                If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                    ' default to FLD-009
                    sRet = GetConnectString("ConceptDocTypes")
                    GoTo Block_Exit
                End If
                sRet = .Parameters("@pConnString").Value
            End With
        
    End If
    
Block_Exit:
    Set oAdo = Nothing
    GetConverterConnectionString = sRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    sRet = GetConnectString("ConceptDocTypes")
    GoTo Block_Err
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function WaitForBatchOrJobFinish(Optional lBatchId As Long, Optional lJobId As Long, _
            Optional sReport As String, Optional bTimedOut As Boolean, Optional lSecondsForTimeout As Long = 120) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim dtTimeoutStart As Date

    strProcName = ClassName & ".WaitForBatchOrJobFinish"
    
    dtTimeoutStart = Now
    bTimedOut = True    ' start this way, if we don't we'll change it
    
    
    Do While Not DateDiff("s", dtTimeoutStart, Now) > lSecondsForTimeout
        DoCmd.Echo True, "Converting files... "
        Sleep 500
        DoEvents
        
        If CheckForBatchFinish(lBatchId, lJobId) = True Then
            bTimedOut = False
            WaitForBatchOrJobFinish = True
            GoTo Block_Exit
        End If
    Loop
  
    
Block_Exit:
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    WaitForBatchOrJobFinish = False
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CheckForBatchFinish(Optional lBatchId As Long, Optional lJobId As Long, Optional iStatusId As Integer) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".CheckForBatchFinish"
    
    Set oAdo = New clsADO
    With oAdo
        ''.ConnectionString = GetConnectString("ConceptDocTypes")
        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_IsJobFinished"
        .Parameters.Refresh
        .Parameters("@pBatchId") = lBatchId
        .Parameters("@pJobId") = lJobId
        .Execute
        If .Parameters("@pErrMsg") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error checking for the batch / job completion:" & .Parameters("@pErrMsg"), .Parameters("@pErrMsg")
            GoTo Block_Exit
        End If
        iStatusId = .Parameters("@pStatusNum").Value
    End With
    
    CheckForBatchFinish = IIf(iStatusId = 1, True, False)
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    CheckForBatchFinish = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function CloseBatch(Optional lBatchId As Long, Optional sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".CloseBatch"
    
    Set oAdo = New clsADO
    With oAdo
'''        .ConnectionString = GetConnectString("ConceptDocTypes")
        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_Close_Batch"
        .Parameters.Refresh
        .Parameters("@pBatchId") = lBatchId
        .Execute
        If .Parameters("@pErrMsg") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error closing the batch: " & .Parameters("@pErrMsg"), .Parameters("@pErrMsg")
            GoTo Block_Exit
        End If
    End With
    
    CloseBatch = True
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    CloseBatch = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function AddConverterQueueJob(sFullPathToFileToConvert As String, sToType As String, _
    Optional sOutFolder As String, Optional sOutFileName As String, Optional bSendEmail As Boolean, _
    Optional bOverWriteIfFound As Boolean, Optional bDeleteOrig As Boolean, Optional iNotificationId As Integer, _
    Optional lBatchId As Long, Optional lJobIdReturned As Long, Optional bWaitForConversion As Boolean = False, Optional lSecondsForTimeout As Long = 120) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oAdo As clsADO
Dim oControl As Control
Dim sParams As String
Dim sErr As String
Dim iTryCount As Integer
Dim bTimedOut As Boolean

    strProcName = ClassName & ".AddConverterQueueJob"
    
        '' Make sure we have what is required:
TryAgain:
    If FileExists(sFullPathToFileToConvert) = False Then
        If iTryCount > 2 Then
            LogMessage strProcName, "ERROR", "The File to convert is required!"
            GoTo Block_Exit
        Else
            iTryCount = iTryCount + 1
            Sleep 1500
            GoTo TryAgain
        End If
    End If

    If CStr(sToType) = "" Then
        LogMessage strProcName, "ERROR", "You need to select a format to convert this file to!"
        GoTo Block_Exit
    End If

            '' 20120502 KD: Bug fix to squash .pdf.pdf named files!
    sOutFileName = Replace(sOutFileName, "." & LCase(sToType), "")

    Set oAdo = New clsADO
    With oAdo
        '''.ConnectionString = GetConnectString("ConceptDocTypes")
        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_Add_Job"
        .Parameters.Refresh
        .Parameters("@pFullPathToFileToConvert") = sFullPathToFileToConvert
        .Parameters("@pToType") = sToType
        
        If lBatchId > 0 Then .Parameters("@pBatchId") = lBatchId
        If iNotificationId > 0 Then .Parameters("@pNotifyId") = iNotificationId
        If sOutFolder <> "" Then .Parameters("@pOutFolder") = sOutFolder
        If sOutFileName <> "" Then .Parameters("@pOutFileName") = sOutFileName
        .Parameters("@pSendEmail") = IIf(bSendEmail, 1, 0)
        .Parameters("@pDeleteOrig") = IIf(bDeleteOrig, 1, 0)
        .Parameters("@pOverwriteIfFound") = IIf(bOverWriteIfFound, 1, 0)
        
        .Execute
        
        sErr = CStr("" & .Parameters("@pErrMsg"))
        If sErr <> "" Then
            LogMessage strProcName, "ERROR", "Problem creating Converter Queue job: " & sErr, sErr, True
            GoTo Block_Exit
        End If
        lJobIdReturned = .Parameters("@pJobId")
    End With
    
    If lJobIdReturned > 0 Then AddConverterQueueJob = True

    If bWaitForConversion = True Then
        Call WaitForBatchOrJobFinish(, lJobIdReturned, , bTimedOut, lSecondsForTimeout)
        If bTimedOut = True Then
            LogMessage strProcName, "TIME OUT", "Converter Queue timed out while waiting for a file to be converted", sFullPathToFileToConvert
        End If
        AddConverterQueueJob = Not bTimedOut
    End If
    
    

Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    AddConverterQueueJob = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function AddConverterQueueBatch(sInFolder As String, bAllFilesInDir As Boolean, sToType As String, _
    Optional sOutFolder As String, Optional sFromTypes As String, Optional bSendEmail As Boolean, _
    Optional bOverWriteIfFound As Boolean, Optional bDeleteOrig As Boolean, Optional iNotificationId As Integer, _
    Optional sErrFldr As String, Optional sMoveOrigFileToFld As String, Optional lBatchIdReturned As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oAdo As clsADO
Dim oControl As Control
Dim sParams As String
Dim sErr As String


    strProcName = ClassName & ".AddConverterQueueBatch"
    
        '' Make sure we have what is required:
    If FolderExist(sInFolder) = False Then
        LogMessage strProcName, "ERROR", "The folder to convert is required!"
        GoTo Block_Exit
    End If

    If CStr(sToType) = "" Then
        LogMessage strProcName, "ERROR", "You need to select a format to convert this file to!"
        GoTo Block_Exit
    End If

    Set oAdo = New clsADO
    With oAdo
        '''.ConnectionString = GetConnectString("ConceptDocTypes")
        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_Add_Batch"
        .Parameters.Refresh
        .Parameters("@pInFolder") = sInFolder
        .Parameters("@pAllFilesFoundInDir") = IIf(bAllFilesInDir, 1, 0)
        .Parameters("@pToType") = sToType
        
        If sOutFolder <> "" Then .Parameters("@pOutFolder") = sOutFolder
        If sErrFldr <> "" Then .Parameters("@pErrorFolder") = sErrFldr
        If sMoveOrigFileToFld <> "" Then .Parameters("@pMoveOrigToDir") = sMoveOrigFileToFld
        
        If sFromTypes <> "" Then .Parameters("@pFromType") = sFromTypes
        .Parameters("@pSendEmail") = IIf(bSendEmail, 1, 0)
        If iNotificationId > 0 Then .Parameters("@pNotifyId") = iNotificationId
        .Parameters("@pDeleteOrig") = IIf(bDeleteOrig, 1, 0)
        .Parameters("@pOverwriteIfFound") = IIf(bOverWriteIfFound, 1, 0)
        
        .Execute
        
        sErr = CStr("" & .Parameters("@pErrMsg"))
        If sErr <> "" Then
            LogMessage strProcName, "ERROR", "Problem creating Converter Queue batch: " & sErr, sErr, True
            GoTo Block_Exit
        End If
        lBatchIdReturned = .Parameters("@pBatchId")
    End With
    
    If lBatchIdReturned > 0 Then AddConverterQueueBatch = True

Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    AddConverterQueueBatch = False
    GoTo Block_Exit
End Function

Public Function CancelBatch(lBatchId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim bRet As Boolean
Dim oAdo As clsADO
Dim sErr As String

    strProcName = ClassName & ".CancelBatch"
    
    If lBatchId < 3 Then
        LogMessage strProcName, "ERROR", "Cannot cancel batch with ID's lower than 3!"
        GoTo Block_Exit
    End If
    
    '' we are going to set this to Complete and times retried to 11 so it isn't tried again.
    '' At some point I should put a canceled field in the tables


    Set oAdo = New clsADO
    With oAdo
'''        .ConnectionString = GetConnectString("ConceptDocTypes")
        .ConnectionString = GetConverterConnectionString()
        .SQLTextType = StoredProc
        .sqlString = "usp_CONVERTER_Cancel_Batch"
        .Parameters.Refresh
        
        .Parameters("@pBatchId") = lBatchId
        
        .Execute
        
        sErr = CStr("" & .Parameters("@pErrMsg"))
        If sErr <> "" Then
            LogMessage strProcName, "ERROR", "Problem creating Converter Queue batch: " & sErr, sErr, True
            GoTo Block_Exit
        End If
        
    End With
    
    bRet = True
    
Block_Exit:
    Set oAdo = Nothing
    CancelBatch = bRet
    Exit Function
Block_Err:
    ReportError Err, strProcName
    bRet = False
    GoTo Block_Exit
End Function


Public Function TestWaitForConverterToFinish() As Boolean
Dim sReport As String
Dim bTimedOut As Boolean
Dim lJobId As Long
Dim lBatchId As Long
Dim dStarted As Date

Const sFromFilePath As String = "Y:\Data\CMS\AnalystFolders\KevinD\_ERAC\_FileConversion_Queue\Test\10anweb.xls"
Const sOutFolder As String = "Y:\Data\CMS\AnalystFolders\KevinD\_ERAC\_FileConversion_Queue\Test\"

    If AddConverterQueueJob(sFromFilePath, "PDF", sOutFolder, , True, True, False, 0, 0, lJobId) = False Then
        Stop
    End If
    
    dStarted = Now

    TestWaitForConverterToFinish = WaitForBatchOrJobFinish(lBatchId, lJobId, sReport, bTimedOut)

Debug.Print ProcessTookHowLong(dStarted)
    
    If bTimedOut = True Then
        MsgBox "Timed out!"
    End If
    If sReport <> "" Then
        MsgBox sReport
    End If
End Function