Option Compare Database
Option Explicit


Public Sub Wilton_Image_Reconcile()
    Dim fso As New FileSystemObject
    
    Dim strInputFile As String
    Dim strOutPutFile As String
    Dim strPhillyFile As String
    Dim strWiltonFile As String
    
    
    strInputFile = "Y:\RAW\CMS\FROM_PHILLY\DailyScan_Files.txt"
    strOutPutFile = "Y:\RAW\CMS\TO_PHILLY\DailyScan_Reconciled_Files.txt"
    
    Open strInputFile For Input As #1
    Open strOutPutFile For Output As #2
    
    While Not EOF(1)
        Line Input #1, strPhillyFile
        If fso.FileExists(strWiltonFile) Then
            Print #2, strPhillyFile
        End If
    Wend
    Close
    MsgBox "Process complete"
End Sub


Public Sub Philly_Image_Catalog()
    Dim fso As New FileSystemObject
    
    Dim strOutPutFile As String
    Dim strFolderPath As String
    Dim bContinue As Boolean
    Dim OutputFileNum
    
    
    strOutPutFile = "Y:\RAW\CMS\FROM_PHILLY\DailyScan_Files.txt"
    
    OutputFileNum = FreeFile
    Open strOutPutFile For Output As #OutputFileNum
    
    bContinue = True
    strFolderPath = "Y:\Raw\CMS\DailyScans"
    Call ScanFileInFolder(fso.GetFolder(strFolderPath), OutputFileNum, bContinue)
    Close
    MsgBox "Process complete"
End Sub


Private Function ScanFileInFolder(oFolder As Folder, OutPutFile, bContinue As Boolean) As Boolean
    Dim oSubFolder As Folder
    Dim oFile As file
    Dim ScanFile As Boolean
    
    
    On Error GoTo Err_handler
Debug.Print oFolder.Path

    For Each oFile In oFolder.Files
        If UCase(Right(oFile.Name, 4)) = ".TIF" Or UCase(Right(oFile.Name, 4)) = ".PDF" Then
Debug.Print oFile.Name
            Print #OutPutFile, oFile.Name
        ElseIf UCase(oFile.Name) = "THUMBS.DB" Then
            oFile.Delete
        End If
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        bContinue = ScanFileInFolder(oSubFolder, OutPutFile, bContinue)
        If bContinue = False Then Exit For
        DoEvents
        DoEvents
    Next
    
    Set oFile = Nothing
    Set oSubFolder = Nothing
    
    ScanFile = bContinue
    
    Exit Function

Err_handler:
    MsgBox Err.Number & " -- " & Err.Description
    MsgBox oFile.Path
    bContinue = False
    ScanFile = bContinue
End Function


Public Sub GetTempImage(frm As Form)
    Dim db As Database
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    
    Dim strErrMsg As String
    
    Dim strDailyScanFile As String ' this is ad-hoc -- thieu
    Dim strMedicalRecordFile As String ' this is ad-hoc -- thieu
    
    
    
    On Error GoTo Err_handler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    Set db = CurrentDb
    db.Execute ("delete from SCANNING_Image_Retransfer")
    
    Dim strSQL As String
    strSQL = "select * from SCANNING_Image_Log_tmp"
    Set rs = MyAdo.OpenRecordSet(strSQL)
   
       
    If (rs.BOF = True And rs.EOF = True) Then
        MsgBox "Nothing to do"
    Else
        rs.MoveFirst
        With rs
            While Not .EOF
                strDailyScanFile = "E:\DATA\IMAGING\CMS\DAILYSCANS\" & !cnlyProvID & "\" & !ImageName
                strMedicalRecordFile = "E:\DATA\IMAGING\CMS\DAILYSCANS\" & !cnlyProvID & "\" & !ImageName
                
                ' this is ad-hoc claim execution (BEGIN)
                db.Execute ("insert into SCANNING_Image_Retransfer values('" & strMedicalRecordFile & "','" & _
                                strDailyScanFile & "','" & !ImagePath & "')")
                    
                ' update display
                If frm.lstFiles.ListCount > 30 Then
                    frm.lstFiles.RemoveItem (0)
                End If
                frm.lstFiles.AddItem strDailyScanFile
                    
                .MoveNext
            Wend
        End With
    End If
    
    
Exit_Sub:
    Set MyAdo = Nothing
    Set rs = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Sub


Public Sub ResetTempImage(frm As Form)
    Dim db As Database
    Dim rs As DAO.RecordSet
    
    Dim strErrMsg As String
    
    Dim strDailyScanFile As String
    Dim strMedicalRecordFile As String
    
    Dim fso As New FileSystemObject
    Dim f As file
    
    Dim bFileExist As Boolean
    Dim strFileName As String
    
    On Error GoTo Err_handler
    
    
    Set db = CurrentDb
    Set rs = db.OpenRecordSet("SCANNING_Image_Retransfer")
    
       
    If (rs.BOF = True And rs.EOF = True) Then
        MsgBox "Nothing to do"
    Else
        rs.MoveFirst
        With rs
            While Not .EOF
                strDailyScanFile = !DailyScanPath
                strMedicalRecordFile = !MedicalRecordPath
                
                bFileExist = False
                    
                If fso.FileExists(strDailyScanFile & ".tif") Then
                    bFileExist = True
                    strFileName = strDailyScanFile & ".tif"
                ElseIf fso.FileExists(strDailyScanFile & ".pdf") Then
                    bFileExist = True
                    strFileName = strDailyScanFile & ".pdf"
                ElseIf fso.FileExists(strMedicalRecordFile & ".tif") Then
                    strFileName = strDailyScanFile & ".tif"
                    Call fso.MoveFile(strMedicalRecordFile & ".tif", strFileName)
                    Set f = fso.GetFile(strFileName)
                    f.Attributes = f.Attributes Or Archive
                    bFileExist = True
                ElseIf fso.FileExists(strMedicalRecordFile & ".pdf") Then
                    strFileName = strDailyScanFile & ".pdf"
                    Call fso.MoveFile(strMedicalRecordFile & ".pdf", strFileName)
                    Set f = fso.GetFile(strFileName)
                    f.Attributes = f.Attributes Or Archive
                    bFileExist = True
                Else
                    strDailyScanFile = strDailyScanFile & "; NOT EXISTS"
                End If
                
                'If bFileExist Then
                '    Set f = fso.GetFile(strFileName)
                '    f.Attributes = f.Attributes Or Archive
                'End If
                    
                
                ' update display
                If frm.lstFiles.ListCount > 30 Then
                    frm.lstFiles.RemoveItem (0)
                End If
                frm.lstFiles.AddItem strFileName
                    
                .MoveNext
            Wend
        End With
    End If
    
    
Exit_Sub:
'    Set myADO = Nothing
    Set rs = Nothing
    Set fso = Nothing
    Set f = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Sub

Public Function ConvertTimeToString(InTime As Date) As String
    Dim myDate As String
    Dim myHour As Integer
    Dim myMinute As Integer
    Dim mySecond As Integer
    Dim myMillisecond As Double
    Dim myTimeValue As Double
    
    myTimeValue = InTime - DateValue(InTime)
    
    myDate = Format(InTime, "mm-dd-yyyy")
    myHour = Int(myTimeValue * 24)
    '2014:03:12:JS Added the code below to correct bug that was making SQL time like 23:59:59.324 to be converted to next day
    If myHour = -1 Then
        myHour = 23
        myDate = Format(InTime - 1, "mm-dd-yyyy")
    End If
    myMinute = Int(myTimeValue * 24 * 60 - Int(myTimeValue * 24) * 60)
    mySecond = Int(myTimeValue * 24 * 3600 - Int(myTimeValue * 24 * 60) * 60)
    myMillisecond = myTimeValue * 24 * 3600 - Int(myTimeValue * 24 * 3600)
    If Format(myMillisecond * 1000, "#000") = "1000" Then
        mySecond = mySecond + 1
        myMillisecond = 0
        
        If mySecond = 60 Then
            mySecond = 0
            myMinute = myMinute + 1
        End If
        
        If myMinute = 60 Then
            myMinute = 0
            myHour = myHour + 1
        End If
    End If
        
    ConvertTimeToString = myDate & " " & Format(myHour, "#00") & ":" & Format(myMinute, "#00") & ":" & Format(mySecond, "#00") & "." & Format(myMillisecond * 1000, "#000")
End Function


Public Sub ImageRecount()
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    
    Dim strFileName As String
    Dim iPageCnt As Integer
    
    Dim fso As New FileSystemObject
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from SCANNING_Image_Log where ValidationDt = '10/15/2010' and imagepath like '%.PDF' and imagename = '450596MR-101014142312337'"
    Set rs = MyAdo.OpenRecordSet
    Debug.Print rs.recordCount
    
    While Not rs.EOF
        strFileName = rs("LocalPath")
        If Not fso.FileExists(strFileName) Then
            strFileName = rs("ImagePath")
        End If
            
        If fso.FileExists(strFileName) Then
            If UCase(Right(strFileName, 3)) = "PDF" Then
                iPageCnt = Count_PDF_Pages(strFileName)
                If iPageCnt > 0 Then
Debug.Print CStr(iPageCnt) & " - " & rs("ImageName")
                    rs("PageCnt") = iPageCnt
                    rs("PDFCnt") = iPageCnt
                End If
            Else
                iPageCnt = Count_TIF_Pages(strFileName)
                If iPageCnt > 0 Then
                    rs("PageCnt") = iPageCnt
                    rs("TIFCnt") = iPageCnt
                End If
            End If
        End If
        
        rs.Update
        rs.MoveNext
    Wend
    
    Call MyAdo.BatchUpdate(rs)
    
    MsgBox "Process complete"
    
    Set rs = Nothing
    Set MyAdo = Nothing
    Set fso = Nothing
End Sub