Option Compare Database
Option Explicit



Sub SplitImages()
    'check if the images have finished converting to PDF and update the status for all those in status TOCONVERT
    
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO
    
    Dim rsToConvertImages As ADODB.RecordSet
    Dim rsToSplitImages As ADODB.RecordSet
    Dim rsToInsertImages As ADODB.RecordSet
    Dim rsToInsertCoverSheets As ADODB.RecordSet
    Dim rsToInsertSplits As ADODB.RecordSet
    
    Dim MyAdo As clsADO
    Dim strSQLcmd As String
    Dim strConvertedFile As String
    Dim strToConvertFile As String
    Dim strConvErrFile As String
    Dim strSplitFile As String
    Dim strErrMsg As String
    Dim intSplitResult As Integer
    
    Dim strSQLCode As String
    Dim iRetCd As Integer
    
    Dim bolFileCopy As Boolean
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
'On Error GoTo Err_Handler
    
    ' First check for the files that have been converted from TIF to PDF
    
    
    strSQLcmd = "select distinct OrigCoverSheetNum, SplitPathBase, SplitPathConvErr, SplitPathIn, SplitSourceFile, SplitsTotal, PgCnt = max(SplitEndPg) from v_SCANNING_FastScan_SplitQueue where SplitProcStatusCd = 'TOCONVERT' group by OrigCoverSheetNum, SplitPathBase, SplitPathConvErr, SplitPathIn, SplitSourceFile, SplitsTotal"
    
    Set rsToConvertImages = MyAdo.OpenRecordSet(strSQLcmd, False)
    
    If Not rsToConvertImages Is Nothing Then
    
        If Not (rsToConvertImages.BOF = True And rsToConvertImages.EOF = True) Then
            
            With rsToConvertImages
                
                .MoveFirst
                
                Do While Not .EOF
                    
                    strToConvertFile = !SplitPathBase & !SplitSourceFile & ".TIF"
                    strConvErrFile = !SplitPathConvErr & !SplitSourceFile & ".TIF"
                    strConvertedFile = !SplitPathIn & !SplitSourceFile & ".PDF"
                    
                    If FileExists(strConvertedFile) Then
                        
                       If Not FileLocked(strConvertedFile) Then
                       
                            If Count_PDF_Pages(strConvertedFile) = !PgCnt Then
                            
                                'coversion was sucessfull so update the status to TOSPLIT

                                strErrMsg = ""
                                
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOSPLIT' where OrigCoverSheetNum = '" & !origcoversheetnum & "'"
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                If iRetCd <> !SplitsTotal Then
                                    Stop
                                    strErrMsg = "Image Split Error: Error Updating FastScan_Splits table with TOSPLIT"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                            Else
                                Stop
                                'page count after convert error
                                strErrMsg = "Image Split Error: Incorrect Converted File Page Count"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOERRORCONVERT' where OrigCoverSheetNum = '" & !origcoversheetnum & "'"
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                If iRetCd <> !SplitsTotal Then
                                    Stop
                                    strErrMsg = "Image Split Error: Cannot Update Incorrect Converted File Page Count"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                            End If
                        
                        Else
                            Stop
                            strErrMsg = "Image Split Error: Cannot Lock Converted file: " & strConvertedFile
                            LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                        
                        End If
                        
                    Else
                        
                        'if it is not waiting to be converted
                        If Not FileExists(strToConvertFile) Then
                            
                            'if it ended up in the conversion error folder
                            If FileExists(strConvErrFile) Then
                            
                                Stop
                                
                                strErrMsg = "Image Split Error: Conversion Process Failed"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                    
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOERRORCONVERT' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                
                                If iRetCd <> 1 Then
                                    Stop
                                    strErrMsg = "Image Split Error: Cannot Update Conversion Process Failed"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                                
                            'file is nowhere to be found
                            Else
                            
                                strErrMsg = "Image Split Error: Cannot Find File to Process"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                    
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOERRORFILELOST' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                
                                If iRetCd <> 1 Then
                                    Stop
                                    strErrMsg = "Image Split Error: Cannot Update Cannot Find File to Process"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                                
                            End If
                        End If
                        
                    End If
                
                    .MoveNext
                
                Loop
            
            End With
            
        End If
    
    End If
    
    
    ' Now do the splits for the PDF files
    Stop
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
    strSQLcmd = "select * from v_SCANNING_FastScan_SplitQueue where SplitProcStatusCd = 'TOSPLIT'"
    
    Set rsToSplitImages = MyAdo.OpenRecordSet(strSQLcmd, False)
    
    If Not (rsToSplitImages.BOF = True And rsToSplitImages.EOF = True) Then
        
        With rsToSplitImages
            
            .MoveFirst
            
            Do While Not .EOF
                
                strConvertedFile = !SplitPathIn & !SplitSourceFile & ".PDF"
                strSplitFile = !SplitPathOut & !ToSplitFilename & "." & !SplitFileExt
                
                If FileExists(strConvertedFile) Then
                    
                   If Not FileLocked(strConvertedFile) Then
                   
                        intSplitResult = SplitPDF(strConvertedFile, !SplitStartPg, !SplitEndPg, strSplitFile)
                        
                        If Not intSplitResult Then 'if split failed

                            Stop
                            strErrMsg = "Image Split Error: Acrobat Split Process Failed"
                            LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                            
                            Set MyAdo = New clsADO
                            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                            strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOERRORSPLIT' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                            MyAdo.sqlString = strSQLCode
                            MyAdo.SQLTextType = sqltext
                            iRetCd = MyAdo.Execute
                            
                            If iRetCd <> 1 Then
                                Stop
                                strErrMsg = "Image Split Error: Cannot Update Acrobat Split Process Failed"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                            End If
                            
                        Else 'if split went good
                   
                            If Count_PDF_Pages(strSplitFile) = !SplitPgCnt Then
                            
                                'coversion was sucessfull so update the status to TOINSERT
                               
                                strErrMsg = ""
                                
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOINSERT' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                If iRetCd <> 1 Then
                                    Stop
                                    strErrMsg = "Image Split Error: Error Updating FastScan_Splits table with TOINSERT"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                            
                            Else 'error split pages dont match
                                                                
                                Stop
                                
                                strErrMsg = "Image Split Error: Incorrect Split File Page Count"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                
                                Set MyAdo = New clsADO
                                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'ERRORSPLIT' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                                MyAdo.sqlString = strSQLCode
                                MyAdo.SQLTextType = sqltext
                                iRetCd = MyAdo.Execute
                                
                                If iRetCd <> 1 Then
                                    Stop
                                    strErrMsg = "Image Split Error: Cannot Update Incorrect Split File Page Count"
                                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                End If
                                
                            End If
                    
                        End If
                    
                    Else
                        'error converted file is locked
                        Stop
                        strErrMsg = "Image Split Error:Image Split Error: Split File Is Locked"
                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False

                    End If
                    
                Else
                    'error converted file does not exists
                    Stop
                    strErrMsg = "Image Split Error: Split File Does Not Exist: " & strConvertedFile
                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                        
                End If
            
                .MoveNext
            
            Loop
        
        End With
        
    End If
    
    'archive error original files and delete error splits
    Stop
    
    Dim strOriginalFileFullPath As String
    Dim strOriginalFileErrorFolder As String
    Dim strOriginalFileErrorFullPath As String
    
    Dim strSplitFileToDeleteFullPath As String
    
    Dim strLastOriginal As String
    
    
    Dim rsToRemoveErrors As ADODB.RecordSet
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
    strSQLcmd = "select distinct OrigCoverSheetNum, SplitPathIn, SplitSourceFile, SplitsTotal, SplitPathOut, ToSplitFilename  from v_SCANNING_FastScan_SplitQueue where SplitsToError >=1"
    
    Set rsToRemoveErrors = MyAdo.OpenRecordSet(strSQLcmd, False)
    

    If Not (rsToRemoveErrors.BOF = True And rsToRemoveErrors.EOF = True) Then
    
        With rsToRemoveErrors
            
            .MoveFirst
            
            strLastOriginal = ""
        
            Do While Not .EOF
            
                strOriginalFileFullPath = !SplitPathIn & !SplitSourceFile & ".PDF"
                strOriginalFileErrorFolder = !SplitPathIn & "ERROR"
                strOriginalFileErrorFullPath = strOriginalFileErrorFolder & "\" & !SplitSourceFile & ".PDF"
            
                If !SplitSourceFile <> strLastOriginal Then
                    
                    If FileExists(strOriginalFileFullPath) Then
                    
                        If Not FolderExists(strOriginalFileErrorFolder) Then
                            If Not CreateFolder(strOriginalFileErrorFolder) Then
                                Stop
                                strErrMsg = "Image Split Error: Cannot Create ERROR folder"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                GoTo Err_handler
                            End If
                        End If
                        
                        If Not MoveFile(strOriginalFileFullPath, strOriginalFileErrorFullPath, False) Then
                            Stop
                            strErrMsg = "Image Split Error: Cannot move original image to ERROR folder"
                            LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                        End If
                        
                    End If
                    
                End If
                
                strLastOriginal = !SplitSourceFile
                
                strSplitFileToDeleteFullPath = !SplitPathOut & !ToSplitFilename & ".PDF"
                
                If FileExists(strSplitFileToDeleteFullPath) Then
                    Call DeleteFile(strSplitFileToDeleteFullPath, False)
                End If
            
                .MoveNext

            Loop
            
        End With
        
    End If
    
    'now that the error files have been taken care off we need to update the tables
    Stop
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
    strSQLcmd = "select distinct OrigCoverSheetNum from v_SCANNING_FastScan_SplitQueue where SplitsToError >=1"
    
    Set rsToRemoveErrors = MyAdo.OpenRecordSet(strSQLcmd, False)
    

    If Not (rsToRemoveErrors.BOF = True And rsToRemoveErrors.EOF = True) Then
    
        With rsToRemoveErrors
            
            .MoveFirst
            
       
            Do While Not .EOF
                    'Update fastscan_log coversheet to indicate the auto split process failed
                    strErrMsg = ""
                    
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    strSQLCode = "Update fl set procstatuscd = 'NOMATCH', " & _
                                    " ProcStatusLastUpDt = getdate(), " & _
                                    " ProcStatusLastUserID = 'SPLITPROCESS', " & _
                                    " NoMatchReasonCd = '05F' " & _
                                    " From cms_auditors_claims.dbo.SCANNING_FastScan_Log fl" & _
                                    " Where CoverSheetNum = '" & !origcoversheetnum & "'"
                    MyAdo.sqlString = strSQLCode
                    MyAdo.SQLTextType = sqltext
                    iRetCd = MyAdo.Execute
                    If iRetCd <> 1 Then
                        Stop
                        strErrMsg = "Image Split Error: Cannot Update FastScan Coversheet with Split Failed status"
                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                    End If
                    
                    'update fastscan_splits
                    strErrMsg = ""
                    
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    strSQLCode = "Update fs set SplitProcStatusCd = replace(SplitProcStatusCd,'TO','FAILED') " & _
                                    " From cms_auditors_claims.dbo.SCANNING_FastScan_Splits fs" & _
                                    " Where OrigCoverSheetNum = '" & !origcoversheetnum & "'"
                    MyAdo.sqlString = strSQLCode
                    MyAdo.SQLTextType = sqltext
                    iRetCd = MyAdo.Execute
                    If iRetCd <= 0 Then
                        Stop
                        strErrMsg = "Image Split Error: Cannot Update Split Coversheet with Split Failed status"
                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                    End If
                    
                    .MoveNext
                    
            Loop
            
        End With
        
    End If
    
    
    'archive the original files for all sucessfull splits
    Stop
    
    'Dim strOriginalFileFullPath As String
    Dim strOriginalFileDoneFolder As String
    Dim strOriginalFileDoneFullPath As String
    
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
    strSQLcmd = "select distinct OrigCoverSheetNum, SplitPathIn, SplitSourceFile, SplitsTotal  from v_SCANNING_FastScan_SplitQueue where SplitProcStatusCd = 'TOINSERT' and SplitsTotal = SplitsToInsert"
    
    Set rsToInsertCoverSheets = MyAdo.OpenRecordSet(strSQLcmd, False)
    
    If Not (rsToInsertCoverSheets.BOF = True And rsToInsertCoverSheets.EOF = True) Then

        With rsToInsertCoverSheets
        
            .MoveFirst
            
            strOriginalFileDoneFolder = !SplitPathIn & "DONE"
            If Not FolderExists(strOriginalFileDoneFolder) Then
                If Not CreateFolder(strOriginalFileDoneFolder) Then
                    Stop
                    strErrMsg = "Image Split Error: Cannot Create Original file DONE folder"
                    LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                    GoTo Err_handler
                End If
            End If

            Do While Not .EOF
            
                
                strOriginalFileFullPath = !SplitPathIn & !SplitSourceFile & ".PDF"
                strOriginalFileDoneFullPath = strOriginalFileDoneFolder & "\" & !SplitSourceFile & ".PDF"
                
                
                If Not FileExists(strOriginalFileDoneFullPath) Then
                
                
                    If Not MoveFile(strOriginalFileFullPath, strOriginalFileDoneFullPath, False) Then
                        
                        Stop
                        strErrMsg = "Image Split Error: Cannot Move Split Source File to DONE folder"
                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                        
                        Set MyAdo = New clsADO
                        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                        strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'TOSPLIT' where OrigCoverSheetNum = '" & !origcoversheetnum & "'"
                        MyAdo.sqlString = strSQLCode
                        MyAdo.SQLTextType = sqltext
                        iRetCd = MyAdo.Execute
                        If iRetCd <> !SplitsTotal Then
                            Stop
                            strErrMsg = "Image Split Error: Cannot Update Cannot Move Split Source File to DONE folder"
                            LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                        End If
    
                    End If
            
                Else
                    'just in case
                    
                    Call DeleteFile(strOriginalFileFullPath, False)
                    
                End If
                .MoveNext
            
            Loop
            
        End With
    
    End If
    
    
    'create coversheets for the images that have sucessfull splits
    Stop
    
    
    Dim strSplitDoneFolder As String
    Dim strSplitDoneFullPath As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
    strSQLcmd = "select distinct OrigCoverSheetNum, SplitsTotal from v_SCANNING_FastScan_SplitQueue where SplitProcStatusCd = 'TOINSERT' and SplitsTotal = SplitsToInsert"
    
    Set rsToInsertCoverSheets = MyAdo.OpenRecordSet(strSQLcmd, False)
    
    If Not (rsToInsertCoverSheets.BOF = True And rsToInsertCoverSheets.EOF = True) Then

        With rsToInsertCoverSheets

            .MoveFirst

            Do While Not .EOF
            
            
                 'Moving the inserted image files into the production folders
                    
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_Code_Database")
    
                    strSQLcmd = "select * from v_SCANNING_FastScan_SplitQueue where SplitProcStatusCd = 'TOINSERT' and SplitsTotal = SplitsToInsert and OrigCoverSheetNum = '" & !origcoversheetnum & "'"
    
                    Set rsToInsertSplits = MyAdo.OpenRecordSet(strSQLcmd, False)
    
                    If Not (rsToInsertSplits.BOF = True And rsToInsertSplits.EOF = True) Then
    
                        With rsToInsertSplits
    
                            .MoveFirst
    
                            bolFileCopy = True
    
                            Do While Not .EOF And bolFileCopy
    
                                strSplitFile = !SplitPathOut & !ToSplitFilename & "." & !SplitFileExt
    
                                bolFileCopy = CopyFile(strSplitFile, !SplitInsertPathBase & !ToInsertFileName & "." & !SplitFileExt, False)
                                
                                If bolFileCopy Then
                                    
                                    Set MyAdo = New clsADO
                                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                                    strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set NewCoverSheetNum = '" & !NewCoverSheetNum & "' where OrigCoverSheetNum = '" & !origcoversheetnum & "' and SplitNumber = " & !SplitNumber
                                    MyAdo.sqlString = strSQLCode
                                    MyAdo.SQLTextType = sqltext
                                    iRetCd = MyAdo.Execute
                                    If iRetCd <> 1 Then
                                        Stop
                                        strErrMsg = "Image Split Error: Cannot Update NewCoverSheetNumer"
                                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                    End If
                                    
                                End If
    
                                .MoveNext
    
                            Loop 'Splits
    
                        End With
    
                    End If
                    
                    
            
                    'if error moving the splits
                    If Not bolFileCopy Then
                        Stop
                        strErrMsg = "Could not copy split files to provider folder " & strSplitFile
                        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                    
                    Else
                    
                        'if moving the split files was sucessfull then insert the new coversheets
                        Set myCode_ADO = New clsADO
                        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
                        
                        myCode_ADO.BeginTrans
                        
                        'Save coversheet to FastScan Table
                        Set cmd = New ADODB.Command
                        cmd.ActiveConnection = myCode_ADO.CurrentConnection
                        cmd.commandType = adCmdStoredProc
                        cmd.CommandText = "usp_SCANNING_FastScan_ScanCoverSheet_MassFromSplit"
                        cmd.Parameters.Refresh
                        cmd.Parameters("@pOrigCoverSheetNum").Value = !origcoversheetnum
                        cmd.Parameters("@pUserID").Value = Identity.UserName
                        
                        cmd.Execute
                                                    
                        If strErrMsg <> "" Then 'if stored procedure threw an error
                        
                            Stop
                            myCode_ADO.RollbackTrans
                            
                            strErrMsg = "Image Split Error: New Coversheet Insert Failed" & " " & strErrMsg
                            
                            Set MyAdo = New clsADO
                            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                            strSQLCode = "update cms_auditors_claims.dbo.SCANNING_FastScan_Splits set SplitProcStatusCd = 'ERRORINSERT' where OrigCoverSheetNum = '" & !origcoversheetnum & "'"
                            MyAdo.sqlString = strSQLCode
                            MyAdo.SQLTextType = sqltext
                            iRetCd = MyAdo.Execute
                            If iRetCd <> !SplitsTotal Then
                                strErrMsg = "Image Split Error: Cannot Update New Coversheet Insert Failed"
                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                            End If
                        
                        Else 'if all went good
                        
                            myCode_ADO.CommitTrans
                            
                            rsToInsertSplits.MoveFirst
                            
                            If Not (rsToInsertSplits.BOF = True And rsToInsertSplits.EOF = True) Then
            
                                With rsToInsertSplits
                                
                                    Do While Not .EOF
            
                                        strSplitFile = !SplitPathOut & !ToSplitFilename & "." & !SplitFileExt
                                        strSplitDoneFolder = !SplitPathOut & "Done"
                                        strSplitDoneFullPath = strSplitDoneFolder & "\" & !ToSplitFilename & "." & !SplitFileExt
                                        
                                        If Not FolderExists(strSplitDoneFolder) Then
                                            If Not CreateFolder(strSplitDoneFolder) Then
                                                Stop
                                                strErrMsg = "Image Split Error: Cannot create Split Done Folder"
                                                LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                                GoTo Err_handler
                                            End If
                                        End If
                                        
                                        If Not MoveFile(strSplitFile, strSplitDoneFullPath, False) Then
                                            Stop
                                            strErrMsg = "Image Split Error: Cannot move split file to Done folder"
                                            LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
                                            
                                        End If
            
                                        .MoveNext
            
                                    Loop 'Splits
            
                                End With
            
                            End If
                        
                        
                        End If
                         
                    End If

                .MoveNext

            Loop 'CoverSheets

        End With

    End If
    
    Set myCode_ADO = Nothing
    
    'update the status of the original coversheets in scanning_fastscan_log to queuecomplete for all sucessfull splits
    Stop
    
    strErrMsg = ""
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQLCode = "Update fl set procstatuscd = 'SPLITCOMPLETE', " & _
                    " ProcStatusLastUpDt = getdate(), " & _
                    " ProcStatusLastUserID = 'SPLITPROCESS', " & _
                    " NoMatchReasonCd = null " & _
                    " From cms_auditors_claims.dbo.SCANNING_FastScan_Log fl" & _
                    " join  cms_auditors_code.dbo.v_SCANNING_FastScan_SplitQueue  fs on fl.CoverSheetNum = fs.OrigCoverSheetNum " & _
                    " Where fs.SplitsTotal = SplitsComplete "
    MyAdo.sqlString = strSQLCode
    MyAdo.SQLTextType = sqltext
    iRetCd = MyAdo.Execute
    If iRetCd < 0 Then
        Stop
        strErrMsg = "Image Split Error: Cannot Update Completed Coversheets in Scanning_FastScan_Log"
        LogMessage "mod_ImageSplit.SplitImages", "ERROR", strErrMsg, , False
    End If
    
GoTo Clean_And_Exit
  
Err_handler:

Stop

MsgBox strErrMsg, vbExclamation, "Error in Split Process"

Clean_And_Exit:

Set cmd = Nothing
Set myCode_ADO = Nothing

Set rsToConvertImages = Nothing
Set rsToSplitImages = Nothing
Set rsToInsertImages = Nothing
Set rsToInsertCoverSheets = Nothing
Set rsToInsertSplits = Nothing
    
End Sub





Private Function SplitPDF(PathIn As String, FromPage As Integer, ToPage As Integer, PathOut As String) As Integer
'*************************
'Path should be the complete path to the input file
'From and To Pages are included in the extract
'From and To Pages are incremented by 1 as Adobe starts counting at ZERO.
'*************************

Dim oAcroApp As Object
'Dim AVPageView As Object
Dim oAcroPdDocIn
Dim oAcroPdDocOut

    
On Error Resume Next
    
    SplitPDF = True
    
    FromPage = FromPage - 1
    ToPage = ToPage - 1
      '//-> Set general used Objects
    Set oAcroApp = CreateObject("AcroExch.App")
    'Set oAcroPdDoc = CreateObject("AcroExch.PDDoc")
    
    Set oAcroPdDocIn = CreateObject("AcroExch.PDDoc")
    Set oAcroPdDocOut = CreateObject("AcroExch.PDDoc")
    
    If Err.Number = 0 Then
        oAcroApp.Hide
        
        If oAcroPdDocIn.Open(PathIn) And oAcroPdDocOut.Create Then
            If oAcroPdDocOut.InsertPages(-1, oAcroPdDocIn, FromPage, (ToPage - FromPage) + 1, False) Then
                If Not oAcroPdDocOut.Save(PDSaveFull, PathOut) Then
                    'error, lets give it another try
                    Sleep 2000
                    If Not oAcroPdDocOut.Save(PDSaveFull, PathOut) Then
                        SplitPDF = False
                    End If
                    
                End If
            Else
                SplitPDF = False
                'error
            End If
            'error
        End If
        
        oAcroPdDocIn.Close
        oAcroPdDocOut.Close
        
        Set oAcroPdDocIn = Nothing
        Set oAcroPdDocOut = Nothing
        
        oAcroApp.Exit
        
        Set oAcroApp = Nothing
    Else
    
        SplitPDF = False
        
    End If
'        Set oAcroAvDoc = CreateObject("AcroExch.AVDoc")
'
'      '//-> Open the File via Avdoc and assign JSO
'        If oAcroAvDoc.Open(PathIn, vbNull) Then
'          'Set AVPageView = AVDoc.GetAVPageView
'          Set oAcroPdDoc = oAcroAvDoc.GetPDDoc
'          Set oAcroJSO = oAcroPdDoc.GetJSObject
'          ' gApp.Show   '<==show or comment out for operate visible/invisible
'
'          'FLNew = Replace(Pathin, ".PDF", "") & "_From_" & Right(FromPage + 100001, 5) & "_To_" & Right(ToPage + 100001, 5) & ".pdf"
'          oAcroJSO.extractPages FromPage, ToPage, PathOut
'        End If
'        'MsgBox "Done!"
'    End If
'
'    'Set oAcroJSO = Nothing
'
'    oAcroPdDoc.Close
'
'    oAcroAvDoc.Close (True)
'    oAcroApp.Exit
'
'    Set oAcroJSO = Nothing
'    Set oAcroPdDoc = Nothing
'    Set oAcroAvDoc = Nothing
'    Set oAcroApp = Nothing
'

End Function

Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      'MsgBox "Error #" & str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
   End If
End Function