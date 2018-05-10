Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_Letter_Reconciliation_Tool"


Public Function ProcessFolder(sFldr As String, Optional oForm As Form_frm_LETTER_Reconciliation_Tool) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim objWordField As Word.Field
Dim objWordSection As Word.Section

Dim lSections As Long
Dim i As Integer
Dim oRng As Word.Range
Dim sTxt As String

Dim oRegEx As RegExp
Dim oMatchs As MatchCollection
Dim oMatch As Match
Dim iSubM As Integer
Dim sDocNum As String, iDocNum As Integer
Dim sInstanceId As String
Dim dtStarted As Date

Const lUpdateModulous As Long = 1
Dim lProcessedCnt As Long
Dim sPctDone As String
Dim lMaxCnt As Long
Dim sElapsedTime As String

    strProcName = ClassName & ".ProcessFolder"

    If Not oForm Is Nothing Then
        oForm.Status = ""
        sPctDone = ""
        lProcessedCnt = 0
        lMaxCnt = oForm.MaxToProcess
        dtStarted = Now()
    End If

'    sFldr = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\Scan Ops\Automated Letters\ToPrint\2014-12-03\VADRA"

    Set oDb = CurrentDb
    oDb.Execute "DELETE FROM LETTER_Reconciliation_Tool"
    Set oRs = oDb.OpenRecordSet("SELECT * FROM LETTER_Reconciliation_Tool WHERE 1 = 2")
    
    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sFldr)
    
    Set oWordApp = New Word.Application
'oWordApp.Visible = True
    
    Set oRegEx = New RegExp
    With oRegEx
        .Global = False
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^Re:\s*?(.*?)\s*?\- ([a-zA-Z0-9]+?)$"
    End With
    
    For Each oFile In oFldr.Files
        lProcessedCnt = lProcessedCnt + 1
        DoEvents
        If Not oForm Is Nothing Then
            oForm.CurrentLetterNum = lProcessedCnt

            DoEvents
            If oForm.StopNow = True Then
                GoTo Block_Exit
            End If
            oForm.UpdateStatus
            While oForm.Paused = True
                SleepEvents 2
'                Stop
            Wend
            
        End If
        Set oWordDoc = oWordApp.Documents.Open(oFile.Path)
        sDocNum = left(oFile.Name, 4)
        iDocNum = sDocNum
'Stop
        sInstanceId = Replace(oFile.Path, sFldr & "\", "", , , vbTextCompare)
        sInstanceId = Replace(sInstanceId, sDocNum & "_", "", , , vbTextCompare)
'Stop
'        For Each objWordSection In oWordApp.ActiveDocument.Sections
'            For Each objWordField In objWordSection.Range.Fields
'                Stop
'
'                Debug.Print objWordField.Type
'
'                Stop
'
'            Next
'        Next

        DoEvents

        With oWordDoc
            ' Loop through Story Ranges and update.
            ' Note that this may trigger interactive fields (eg ASK and FILLIN).
            For Each oRng In .StoryRanges
                Do
'                    Stop
'                    Debug.Print oRng.Text
                    sTxt = left(oRng.Text, 2000)
                    Debug.Print sTxt
                    Set oMatchs = oRegEx.Execute(sTxt)
                    
                    If oMatchs.Count = 0 Then
                        Stop ' fix it!
                    End If
                    
                    For Each oMatch In oMatchs
                        Debug.Print oMatch
                        Debug.Print oMatch.SubMatches(0)
                        Debug.Print oMatch.SubMatches(1)
                        oRs.AddNew
                        oRs("DocNum") = iDocNum
                        oRs("DocName") = oFile.Path
                        oRs("InstanceId") = sInstanceId
                        oRs("ProvName") = oMatch.SubMatches(0)
                        oRs("ProvNum") = oMatch.SubMatches(1)
                        oRs.Update
                        GoTo NextDoc
                    Next

'                    oRng.Fields.Unlink
'                    For Each hLink In oRng.Hyperlinks
'                        hLink.Delete
'                    Next
                    Set oRng = oRng.NextStoryRange
                Loop Until oRng Is Nothing

            Next
        End With
NextDoc:
        oWordDoc.Close False
        DoEvents
    Next
    

    
Block_Exit:
    Set oDb = Nothing
    
    Set oRs = Nothing
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing

    If Not oWordApp Is Nothing Then
        For Each oWordDoc In oWordApp.Documents
            oWordDoc.Close False
        Next
        oWordApp.Quit
        Set oWordApp = Nothing
    End If

    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

''''
'''' This needs to do the same thing that the new tool will do
'''' - create a 'done' folder
'''' - prompt the user to choose the printer
'''' - set that as the default printer (keeping track of what WAS the default)
'''' - loop over each letter found in the folder in order
'''' - Record the instance id, printer chosen and person who printed it, start time
'''' - Send the letter to the printer
'''' - record the instanceid, printer chosen and person who printed it, and the stop time
'''' - move the file into the done folder (unless there was an error)
'''' - finish the loop
'''' - on error, plop up a message
'''' - return how many were succesful
'''PrintFolder(Me.txtFolderToProcess, sLetterType, sLtrReqDt, Me)
Public Function PrintFolder(sFldr As String, sLetterType As String, sLtrReqDt As String, Optional oForm As Form_frm_LETTER_Legacy_Print_Tool) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim dtStarted As Date

Const lUpdateModulous As Long = 1
Dim lProcessedCnt As Long
Dim sPctDone As String
Dim lMaxCnt As Long
Dim sElapsedTime As String
Dim sSelectedPrinter As String
Dim strPrinterBefore As String
Dim lNumPrinted As Long

    strProcName = ClassName & ".PrintFolder"

    If Not oForm Is Nothing Then
        oForm.Status = ""
        sPctDone = ""
        lProcessedCnt = 0
        lMaxCnt = oForm.MaxToProcess
        dtStarted = Now()
    End If

    If FolderExists(sFldr) = False Then
        Stop
        GoTo Block_Exit
    End If
    
    sSelectedPrinter = SelectPrinter(, strPrinterBefore)
    
    DoCmd.Hourglass True


' here we need to get the details for the coversheet

    If PrepCoverSheet(sFldr, sLetterType, sLtrReqDt, sSelectedPrinter) = False Then
        Stop
        GoTo Block_Exit
    End If


    If CreatecoverPageAndPrint_4Legacy(sLetterType, sLtrReqDt, sFldr, sSelectedPrinter, oForm, lNumPrinted) = False Then
        Stop
    End If
    PrintFolder = lNumPrinted

    
Block_Exit:
    DoCmd.Hourglass False
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Function PrepCoverSheet(sFldr As String, sLetterType As String, sLtrReqDt As String, sSelectedPrinter As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oFile As Scripting.file
Dim lPageCnt As Long
Dim lThisCnt As Long
Dim lLtrCnt As Long
Dim sInstanceId As String
Dim sDocNum As String, iDocNum As Integer
Dim sFPath As String
Dim iTried As Integer

    strProcName = ClassName & ".PrepCoverSheet"

    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sFldr)
    
    For Each oFile In oFldr.Files
     
        sDocNum = left(oFile.Name, 4)
        If IsNumeric(sDocNum) = False Then
            GoTo NextOne
        End If
        iDocNum = sDocNum
        sFPath = oFile.Path
TryAgain:
        sInstanceId = Replace(sFPath, sFldr & "\", "", , , vbTextCompare)
        sInstanceId = Replace(sInstanceId, sDocNum & "_", "", , , vbTextCompare)
        sInstanceId = left(sInstanceId, InStr(1, sInstanceId, ".doc", vbTextCompare) - 1)
        sInstanceId = Replace(sInstanceId, sLetterType & "-", "")
        sInstanceId = Replace(sInstanceId, "Reprint-", "")
        
'Stop
        lThisCnt = Nz(DLookup("PageCount", "LETTER_Static_Details", "InstanceId = '" & sInstanceId & "'"), 0)
        If lThisCnt = 0 Then
            If iTried > 0 Then
                LogMessage strProcName, "ERROR", "Could not find instanceId for this file", sInstanceId & ": " & oFile.Path, True
                GoTo Block_Exit
            End If
        
            
            If InStr(1, oFile.Name, "reprint", vbTextCompare) > 0 Then
                sFPath = Replace(oFile.Path, "-Reprint-", "_")
                iTried = iTried + 1
                GoTo TryAgain
            End If
        End If
        lPageCnt = lPageCnt + lThisCnt
        lLtrCnt = lLtrCnt + 1

NextOne:
    Next
    
    Call InsertBatchDetails(sFldr, sLetterType, lPageCnt, lLtrCnt, sSelectedPrinter)
    
    PrepCoverSheet = True
    
Block_Exit:
    Set oFile = Nothing
    Set oFldr = Nothing
    Set oFso = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function UpdateBatchDetails(ByVal sOutputFolderPath As String, ByVal sLetterType As String, sPrinterName As String) As Boolean
' sFldr, sLetterType, sLtrReqDt, sSelectedPrinter
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".UpdateBatchDetails"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_UpdateBatchDetails"
        .Parameters.Refresh
        .Parameters("@pLetterType") = sLetterType
        .Parameters("@pOutputFolderPath") = QualifyFldrPath(sOutputFolderPath)
        .Parameters("@pPrinterName") = sPrinterName
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            
        End If
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Function


Private Function InsertBatchDetails(ByVal sOutputFolderPath As String, ByVal sLetterType As String, ByVal lTtlPageCnt As Long, ByVal lLetterCount As Long, ByVal sSelectedPrinter As String) As Boolean
' sFldr, sLetterType, sLtrReqDt, sSelectedPrinter
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".InsertBatchdetails"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_InsertBatchDetails"
        .Parameters.Refresh
        .Parameters("@pLetterType") = sLetterType
        .Parameters("@pOutputFolderPath") = QualifyFldrPath(sOutputFolderPath)
        .Parameters("@pTtlPageCnt") = lTtlPageCnt
        .Parameters("@pLetterCnt") = lLetterCount
        .Parameters("@pPrinter") = sSelectedPrinter
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            
        End If
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Function