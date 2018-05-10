Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private strSessionID As String
Private strCDFilesDestination As String
Private strCDSaveDestination As String
Private strMRLoadDestination As String
Private mbContinue As Boolean

Const CstrFrmAppID As String = "CD_Load"

'---------- MR CD Auto Load Module -----------------
'---------------------------------------------------
'
'       Created by      James Segura
'       on              10/04/2012
'
' Put in Production :
'
'
'


Public Sub RefreshData()



    UpdateStatus "Idle."
    
    
    Me.subQuickImageLog.Form("btnMarkAll").visible = True
    Me.subQuickImageLog.Form("cboReceivedMeth") = "CD"
    Me.subQuickImageLog.Form("cboReceivedMeth").Locked = True
    Me.subQuickImageLog.Form("cboReceivedMeth").visible = True
    Me.subQuickImageLog.Form("cboCarrier").visible = True
    Me.subQuickImageLog.Form("cmdGenerateCoverPage").Caption = "Generate CD Reqs"
    Me.subQuickImageLog.Form("txtBatchID").visible = True
    Me.subQuickImageLog.Form("lblBatchID").visible = True
    Me.subQuickImageLog.Form("lblCnlyProvID").visible = True
    Me.subQuickImageLog.Form("txtCnlyProvID").visible = True
    ClearCDLoadFilesTempTable

    LoadCDFilesList True
    
    CalcTotals
    
End Sub


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub btnBrowseCDSourceFolder_Click()
    Me.txtCDSourceFolder = FolderWithoutFile(BrowseFolder("Select CD Folder", Me.txtCDSourceFolder, msoFileDialogViewList))
    'Me.txtCDSourceFolder = GetSourceFolder(Me.txtCDSourceFolder)
    If Nz(Me.txtCDSourceFolder, "") = "" Then Me.txtCDSourceFolder = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\DailyScans\"
End Sub

Function GetSourceFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetSourceFolder = sItem
Set fldr = Nothing
End Function


Private Sub btnExit_Click()

    If MsgBox("Are you sure want to go back clean all info?", vbQuestion + vbYesNo, "About to purge and start over") = vbNo Then
        Exit Sub
    End If
    
    'clean the cd dest folder
    If FolderExists(strCDFilesDestination) Then
        Clear_All_Files_And_SubFolders_In_Folder strCDFilesDestination, True
    End If
    
    Me.subQuickImageLog.Enabled = True
    Me.subQuickImageLog.Form.Purge_Record
    Me.subQuickImageLog.Form("cboRequestNumber") = ""
    Me.subQuickImageLog.Form("cboRequestNumber").SetFocus
    Me.subQuickImageLog.Form.Refresh_Screen
    Me("tabCtlMain") = 0
    Me.subQuickImageLog.Form("btnLoadMRCD").visible = False
    Me.subQuickImageLog.Form("txtBatchID").visible = False
    Me.subQuickImageLog.Form("cboReceivedMeth") = "CD"
    Me.subQuickImageLog.Form("cboReceivedMeth").Locked = True
    Me.subQuickImageLog.Form("cboReceivedMeth").visible = True
    Me.btnMoveFiles.Enabled = True
    Me.btnMarkNotImage.Enabled = True
    Me.btnOpenLoadedFile.Enabled = True
    Me.btnOpenLoadedFolder.Enabled = True
    Me.btnLoadFilesFromCD.Enabled = True
    Me.btnReload.Enabled = True
    Me.subQuickImageLog.Enabled = True
    Me.tabAuto.visible = False
    Me.txtBatchID = ""
    Me.txtCnlyProvID = ""
    Me.txtCDSourceFolder = ""
    Me.txtZipPassword = ""
    Me.subFileList.SourceObject = ""
    CalcTotals
    Me.TabCtlMain = 0
End Sub

Private Sub btnLoadFilesFromCD_Click()

    Dim UserAnswer As Integer
    Dim fldSource As Folder
    Dim fldDestination As Folder
    Dim fso As FileSystemObject
    Dim Result As Boolean
    
    If Nz(Me.txtCDSourceFolder, "") = "" Then
        MsgBox "No Source Folder for CD MRs!", vbExclamation, "Error: No Source Folder"
        Exit Sub
    End If
    
    
    If Nz(Me.txtCDSourceFolder, "") = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\DailyScans\" Then
        MsgBox "You didn't select a specific CD folder.", vbExclamation, "Error: No Source Folder"
        Exit Sub
    End If

   
'    If Nz(Me.txtZipPassword, "") = "" Then
'        UserAnswer = MsgBox("Are you sure you want to continue without a password?", vbQuestion + vbYesNo, "Password Missing")
'        If UserAnswer = vbNo Then Exit Sub
'    End If
    
    If Nz(Me.txtBatchID, "") = "" Or Nz(Me.txtCnlyProvID, "") = "" Then
        MsgBox "Cannot continue without BatchID and CnlyProvID", vbExclamation, "Error: missing parameters"
        Exit Sub
    End If
    
    
    Set fso = CreateObject("scripting.filesystemobject")
    If fso.FolderExists(strCDFilesDestination) = False Then
        Call CreateFolder(strCDFilesDestination)
        Dim FreeMeg As Long
        FreeMeg = FreeMegaBytes(strCDFilesDestination)
        If FreeMeg < 500 Then
            MsgBox "At least 5GB of free space is needed on " & strCDFilesDestination, vbExclamation, "Error, Can't Continue"
            Exit Sub
        End If
    Else
        UserAnswer = MsgBox("The destination folder: " & strCDFilesDestination & " already exists." & vbNewLine & _
                                "Please choose Yes to delete all its contents and continue. Choose No to stop now", vbYesNo, "Folder exists")
        If UserAnswer = vbYes Then
            UpdateStatus "Cleaning folder", strCDFilesDestination
            Clear_All_Files_And_SubFolders_In_Folder (strCDFilesDestination)
        Else
            Exit Sub
        End If
    End If
    
    DoCmd.Hourglass True
    
    'Move all files into local folder, zip or no zip
    Set fldSource = fso.GetFolder(Me.txtCDSourceFolder)
    Set fldDestination = fso.GetFolder(strCDFilesDestination)
    
    ClearCDLoadFilesTempTable
    LoadCDFilesList True
'    Me.subFileList.Form.RecordSource = ""
    
    Result = True
    
    Call ScanCDFiles(fldSource, fldDestination, Result)
    
    If Not Result Then
        MsgBox "Error ocurred while copying data from disk. Can't Continue!", vbExclamation, "Error"
        ClearCDLoadFilesTempTable
        Me.subFileList.SourceObject = ""
        UpdateStatus "Error Ocurred. Try Again."
        GoTo OuttaHere
    End If
    
   
    'find the claims in the system
    MatchFilesToICN_Click
    
    ' add the claims that are not matched to the list
    AddUnMatchedClaimsToTempTable
    
    'load all file names into the list (temp table)
    LoadCDFilesList
    
    UpdateStatus "Idle."
    
    'caculate total fields
    CalcTotals
    
OuttaHere:

    DoCmd.Hourglass False
    
    Set fso = Nothing
    
End Sub

Private Sub MatchFilesToICN_Click()
    UpdateStatus "Matching files to claims"
    Dim MyAdo As clsADO
    Dim cmd As ADODB.Command

    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = MyAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_SCANNING_Match_CD_FIles"
    cmd.Parameters.Refresh
    cmd.Parameters("@BatchID") = Me.txtBatchID
    cmd.Parameters("@UserID") = Identity.UserName()
    ' execute stored procedure that matches claims in list against ICNs in the request
    cmd.Execute
    
    'reload list to show changes
    LoadCDFilesList
    
    'rename matched file
    Dim rs As ADODB.RecordSet
    Dim NewName As String
    
    Set rs = Me.subFileList.Form.RecordsetClone
    If Not rs Is Nothing Then
        If Not (rs.EOF Or rs.BOF) Then
        With rs
            rs.MoveFirst
            While Not rs.EOF
                'if the file was matched then rename the file name to the icn so it is easier to read in the list
                If !FileStatus = "MATCHED A" And FileWithoutExtension(!FileName) <> !Icn Then
                
                    Call RenameWithICN(strCDFilesDestination, !FileName, !Icn, NewName)
                    
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    MyAdo.sqlString = " UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp " & _
                                        " SET FileName = '" & NewName & "'" & _
                                        " where BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and filename = '" & !FileName & "' and AccountID = " & gintAccountID
                    MyAdo.SQLTextType = sqltext
                    MyAdo.Execute
                    
                End If
            .MoveNext
            Wend
        End With
        End If
    End If
    Set rs = Nothing
    MyAdo.DisConnect
    Set MyAdo = Nothing
    
    UpdateStatus "Done Matching Claims."
End Sub

Private Sub btnManualMatch_Click()

    If Me.lstUnmatchedClaims.ListIndex = -1 Or Me.lstUnmatchedFiles.ListIndex = -1 Then
        
        MsgBox "You must select a file and a claim to match", vbInformation, "Match"
        Exit Sub
        
    End If
    Me.WebBrowser.Object.Navigate "http://localhost/"
    Me.WebBrowser.visible = False
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    Me.WebBrowser.Object.Navigate "about:blank"
    Me.WebBrowser.Refresh
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    Me.WebBrowser.visible = True
    While Me.WebBrowser.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Wend
    Sleep 1000
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    ManualMatch Me.lstUnmatchedFiles.Column(0, Me.lstUnmatchedFiles.ListIndex), Me.lstUnmatchedClaims.Column(0, Me.lstUnmatchedClaims.ListIndex), Me.lstUnmatchedClaims.Column(4, Me.lstUnmatchedClaims.ListIndex)

    TabCtlMain_Change
    
    LoadCDFilesList
End Sub


Private Sub btnMarkNotImage_Click()
    Dim UserAnswer As Integer
    Dim MyAdo As clsADO
    
    If Me.subFileList.Form("FileName") = "---" Then
        MsgBox "You can only use this feature to open existing CD files.", vbExclamation, "Error: No CD Files"
        GoTo OuttaHere
    End If
    
    Set MyAdo = New clsADO
    
    'If the file is an MR
    If Me.subFileList.Form.FileIsMR = 1 Then
        UserAnswer = MsgBox("Are you sure you want to make this file a Not Image?", vbQuestion + vbYesNo, "Switch Image Status")
        If UserAnswer = vbNo Then
            Exit Sub
        End If
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp " & _
                            " SET FileIsMR = 0, FileStatus = 'Copied', ICN = '', ImageName = '', MRValid = 0 " & _
                            " WHERE BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and FileName = '" & Me.subFileList.Form.FileName & "' and AccountID = " & gintAccountID
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute
    'if not
    Else
        UserAnswer = MsgBox("Are you sure you want to make this file an Image?", vbQuestion + vbYesNo, "Switch Image Status")
        If UserAnswer = vbNo Then
            Exit Sub
        End If
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp " & _
                            " SET FileIsMR = 1, FileStatus = 'Not Matched', ICN = '---', ImageName = '---', MRValid = 0 " & _
                            " WHERE BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and FileName = '" & Me.subFileList.Form.FileName & "' and AccountID = " & gintAccountID
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute
    End If
    
OuttaHere:
    MyAdo.DisConnect
    Set MyAdo = Nothing
    LoadCDFilesList
    CalcTotals
End Sub

Private Sub btnMoveFiles_Click()

    Dim UserAnswer As Integer
    Dim fso As FileSystemObject
    Dim Extension As String
    
    If (Me.txtFilesNotMatched > 0 Or Me.txtClaimsNotMatched > 0) Then
        UserAnswer = MsgBox("Are you sure you want to continue without matching all files/claims?", vbYesNo + vbQuestion, "Unmatched Files / Claims")
        If UserAnswer = vbNo Then
            Exit Sub
        End If
    End If
    If Me.txtFilesMatched = 0 Then
        MsgBox "You didn't match any file, nothing to move.", vbInformation, "Error: No Matched files"
        Exit Sub
    End If
    
    UserAnswer = MsgBox("Are you sure you want to move the matched files to be loaded in the system?", vbQuestion + vbYesNo, "Load Files")
    If UserAnswer = vbNo Then
        Exit Sub
    End If
    
    Set fso = New FileSystemObject
    
    If fso.FolderExists(strMRLoadDestination) = False Then
        If Not CreateFolder(strMRLoadDestination) Then
            MsgBox "Error Creating Folder. Please try again.", vbExclamation, "Error"
            Exit Sub
        End If
    End If
    
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    Set rs = Me.subFileList.Form.RecordsetClone
    
    With rs
        .MoveFirst
        While Not .EOF
            If !FileIsMR = 1 And !MRValid = 1 Then
                Select Case !FileType
                    Case "PDF Image"
                        Extension = ".pdf"
                    Case "TIF Image"
                        Extension = ".tif"
                    Case Else
                        Extension = ".pdf"
                End Select
                fso.CopyFile strCDFilesDestination & !FileName, strMRLoadDestination & !ImageName & Extension, True
                
                If FileExists(strMRLoadDestination & !ImageName & Extension) Then
                'updating CD Load table
                    MyAdo.sqlString = "UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp set FileStatus = 'MOVED' where BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and FileName = '" & !FileName & "' and AccountID = " & gintAccountID
                Else
                    MyAdo.sqlString = "UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp set FileStatus = 'ERROR - Not Moved' where BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and FileName = '" & !FileName & "' and AccountID = " & gintAccountID
                End If
                
                MyAdo.SQLTextType = sqltext
                MyAdo.Execute

                
            End If
            .MoveNext
        Wend
    End With
    
    
    'update image_log_temp with pagecount
    MyAdo.sqlString = " UPDATE t2 set t2.PageCnt = t3.filepagecnt " & _
                      " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                      " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                      " ON t1.ScannedDt=t2.ScannedDt " & _
                      " left join cms_auditors_claims.dbo.SCANNING_CDLoad_Temp as t3 " & _
                      " on t1.ICN = t3.ICN and t1.sessionid = t3.batchid " & _
                      " where t1.sessionid = '" & Me.txtBatchID & "' and t3.MRValid = 1 and t3.FileIsMR = 1 and t2.accountid = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    
    'delete from image_log_tmp rows for images not received/not matched
    MyAdo.sqlString = " delete t1  " & _
                      " FROM cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t1 " & _
                      " join cms_auditors_claims.dbo.SCANNING_CDLoad_Temp as t3 " & _
                      " on t1.ICN = t3.ICN " & _
                      " where t1.receivedmeth = 'CD' and t3.batchid = '" & Me.txtBatchID & "' and (t3.MRValid = -1 or t3.FileIsMR = -1) and t1.accountid = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    
    'delete from quick_image_log rows for images not received/matched
    MyAdo.sqlString = " delete t1  " & _
                      " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                      " join cms_auditors_claims.dbo.SCANNING_CDLoad_Temp as t3 " & _
                      " on t1.ICN = t3.ICN " & _
                      " t1.receivedmeth = 'CD' and t1.sessionid = '" & Me.txtBatchID & "' and (t3.MRValid = -1 or t3.FileIsMR = -1) and t2.accountid = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    
    
    If Not FolderExists(strCDSaveDestination) Then
        If Not CreateFolder(strCDSaveDestination) Then
            MsgBox "Error: Cannot create destination folder: " & strCDSaveDestination, vbExclamation, "Error Moving files"
            Exit Sub
        End If
    End If
    
    UpdateStatus "Making a save copy of all files"
    fso.CopyFile strCDFilesDestination & "*.*", strCDSaveDestination, True
    
    Clear_All_Files_And_SubFolders_In_Folder strCDFilesDestination, True
    
    Set rs = Nothing
    MyAdo.DisConnect
    Set MyAdo = Nothing

    UpdateStatus "Done."

    Me.btnExit.SetFocus
    Me.btnLoadFilesFromCD.Enabled = False
    Me.btnMoveFiles.Enabled = False
    Me.tabManual.visible = False
    Me.btnMarkNotImage.Enabled = False
    Me.btnOpenLoadedFile.Enabled = False
    Me.btnOpenLoadedFolder.Enabled = False
    Me.btnReload.Enabled = False
    Me.subQuickImageLog.Enabled = False
    

End Sub

Private Sub btnOpenLoadedFile_Click()
    If Me.subFileList.Form("FileName") = "---" Then
        MsgBox "You can only use this feature to open existing CD files.", vbExclamation, "Error: No CD Files"
        Exit Sub
    End If
    Call ShellWait("cmd /C " & Chr(34) & strCDFilesDestination & Me.subFileList.Form("FileName") & Chr(34), vbMinimizedNoFocus)
End Sub

Private Sub btnOpenLoadedFolder_Click()
    'Call ShellWait("explorer " & Chr(34) & strCDFilesDestination & Chr(34), vbNormalFocus)
    If FolderExists(Me.txtCDSourceFolder) Then
        Call ShellWait("explorer " & Chr(34) & Me.txtCDSourceFolder & Chr(34), vbNormalFocus)
    Else
        MsgBox "Source folder does not exist! Try Again.", vbExclamation, "Error with source folder"
    End If
End Sub

Private Sub btnReload_Click()
    
    Dim Result As Boolean
    Dim fldSource As Folder
    Dim fldDestination As Folder
    Dim fso As FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ClearCDLoadFilesTempTable
    LoadCDFilesList True
    'Me.subFileList.Form.RecordSource = Null
    
    Result = True
  
    Set fldSource = fso.GetFolder(strCDFilesDestination)
    Set fldDestination = fso.GetFolder(strCDFilesDestination)

    
    Call ScanCDFiles(fldSource, fldDestination, Result, False)
    
    If Not Result Then
        MsgBox "Error ocurred while copying data from disk. Can't Continue!", vbExclamation, "Error"
        ClearCDLoadFilesTempTable
        Me.subFileList.SourceObject = ""
        UpdateStatus "Error Ocurred. Try Again."
        GoTo OuttaHere
    End If
    
   
    'load all file names into the list (temp table)
    
    'find the claims in the system
    'MatchClaims
    
    MatchFilesToICN_Click
    
    AddUnMatchedClaimsToTempTable
    
    LoadCDFilesList
    
    CalcTotals
    
    UpdateStatus "Idle."
    
    
'    AddToTempTable Me.txtBatchID, Me.subFileList.Form("FileName"), True
'    MatchFilesToICN_Click
'    UpdateStatus "Idle."

OuttaHere:
    
Exit Sub
    
End Sub

Private Sub btnResetSearch_Click()
    Me.txtMemberName = ""
    Me.txtDOB = ""
    ReloadManualMatchClaimList
End Sub


Private Sub chkHideIsImage_AfterUpdate()
    LoadCDFilesList
End Sub

Private Function ScanCDFiles(oSourceFolder As Folder, oDestinationFolder As Folder, ByRef Result As Boolean, Optional CopyToWorkFolder As Boolean = True) As Boolean
    Dim oSubFolder As Folder
    Dim ZipFileExtractFolder As String
    Dim oFile As file
    Dim fso As FileSystemObject
    Dim PasswordPart As String
    Dim Command As String
    Dim Test1 As Integer
    Dim Test2 As Integer
    Dim FileWorkedOn As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Err_handler
    
    UpdateStatus "Start scanning source folder"
    
    
    Debug.Print oSourceFolder.Path
    
    For Each oFile In oSourceFolder.Files
        If IsZipFile(oFile.Name) Then
        
            'cannot be that there is a zip file in the work folder
            If Not CopyToWorkFolder Then Err.Raise 65000, "Found Zip file in work folder"
            
            Debug.Print "Zip File Name " & oFile.Name
            
            'here is where it unzips
            
            If Nz(Me.txtZipPassword, "") <> "" Then
                PasswordPart = " -s" & Me.txtZipPassword & " "
            Else
                PasswordPart = ""
            End If
            
            UpdateStatus "Unzipping", oFile.Path
            
            Call CreateFolder(strCDFilesDestination & oFile.Name)
            'Chr(34) & "c:\PROGRA~2\winzip\ .exe" & Chr(34) & "
            Command = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\APPEALS\WINZIP\wzunzip.exe" & " -ybc -o " & PasswordPart & """" & oFile.Path & """" & " " & """" & strCDFilesDestination & oFile.Name & """" & ""
            Call ShellWait(Command, vbHide)
            If NumFilesInFolder(strCDFilesDestination & oFile.Name) > 0 Then
                ZipFileExtractFolder = strCDFilesDestination & oFile.Name
                ' warning: recursivity ahead.
                Call ScanCDFiles(fso.GetFolder(ZipFileExtractFolder), oDestinationFolder, Result)
                Clear_All_Files_And_SubFolders_In_Folder ZipFileExtractFolder, True
            Else
                'AddToTempTable Me.txtBatchID, oFile.Name, "ERROR: Can't Unzip"
                MsgBox "Error unziping file: " & oFile.Name, vbInformation, "Can't Unzip"
                GoTo Err_handler
            End If
        
        Debug.Print oFile.Path
        Else
            Debug.Print "No zip File Name " & oFile.Name
            'here is where it copies or converts to pdf while copying
            FileWorkedOn = oFile.Name
            UpdateStatus "Reading file ", FileWorkedOn
            If CopyToWorkFolder Then
                If FileExtension(oFile.Name) = "TIF" Then
                    UpdateStatus "Converting TIF to PDF ", oFile.Name
                    FileWorkedOn = FileWithoutExtension(oFile.Name) & ".pdf"
                    If Not TiffToPdf(oFile.Path, strCDFilesDestination & FileWorkedOn) Then GoTo Err_handler
                Else
                    Call fso.CopyFile(oFile.Path, strCDFilesDestination & oFile.Name, True)
                End If
                
            End If
            AddToTempTable Me.txtBatchID, FileWorkedOn
        End If
    Next
    
    For Each oSubFolder In oSourceFolder.SubFolders
Debug.Print oSourceFolder.SubFolders.Count
        mbContinue = ScanCDFiles(oSubFolder, oDestinationFolder, Result)
Debug.Print oSourceFolder.Path
        If mbContinue = False Then Exit For
        DoEvents
        DoEvents
    Next
    
    Set oFile = Nothing
    Set oSubFolder = Nothing
    Set fso = Nothing
    
    ScanCDFiles = mbContinue
    
    Exit Function

Err_handler:
    mbContinue = False
    ScanCDFiles = mbContinue
    Result = False
End Function


Sub UpdateStatus(Message1 As String, Optional ByVal Message2 As String)
    If Len(Message2) > 80 Then
        Message2 = "..." & Right(Message2, 67)
    End If
    
    Me.txtCurrentStatus = Message1 & " " & Message2 & IIf(Right(Message1, 1) = ".", "", "...")
    
    If Not Right(Message1, 1) = "." Then
        Me.txtCurrentStatus.BackColor = 65535
    Else
        Me.txtCurrentStatus.BackColor = 12632256
    End If
    
    Me.Repaint
End Sub
Function NumFilesInFolder(FolderToCheck As String) As Integer

Dim objFS As New Scripting.FileSystemObject
Dim objFolder As Scripting.Folder

NumFilesInFolder = 0

If objFS.FolderExists(FolderToCheck) Then
    Set objFolder = objFS.GetFolder(FolderToCheck)
    NumFilesInFolder = objFolder.Files.Count
End If

Set objFolder = Nothing
Set objFS = Nothing
End Function


Public Sub ClearCDLoadFilesTempTable()
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "DELETE FROM cms_auditors_claims.dbo.SCANNING_CDLoad_Temp where BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and AccountID = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    Me.txtFilesImage = "--"
    Me.txtFilesNotImage = "--"
    Me.txtFilesMatched = "--"
    Me.txtFilesNotMatched = "--"
    Me.txtFilesNotMatched.ForeColor = 16384
    Me.txtClaimsNotMatched = "--"
    Me.txtClaimsNotMatched.ForeColor = 16384
    Me.txtTotalFiles = "--"
    Me.txtTotalClaims = "--"
    MyAdo.DisConnect
    Set MyAdo = Nothing
End Sub


Sub AddUnMatchedClaimsToTempTable()
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "INSERT INTO cms_auditors_claims.dbo.SCANNING_CDLoad_Temp (FileName , FileType , MRValid, BatchID, FileStatus, UserID, FileIsMR, FilePageCnt, ICN, ImageName, AccountID, AddedDt) " & _
                     " SELECT '---', '---', -1, t1.sessionID, 'UNMATCHED CLAIM', '" & Identity.UserName() & "', -1, 0, t1.ICN, t1.ImageName, " & gintAccountID & ", '" & Now & "'" & _
                     " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                     " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                     " ON t1.ScannedDt=t2.ScannedDt " & _
                     " left join cms_auditors_claims.dbo.SCANNING_CDLoad_Temp as t3 " & _
                     " on t1.ICN = t3.ICN and t1.sessionid = t3.batchid " & _
                     " where t1.sessionid = '" & Me.txtBatchID & "' and isnull(t3.icn,'')='' "

    MyAdo.SQLTextType = sqltext
    MyAdo.Execute

    MyAdo.DisConnect
    Set MyAdo = Nothing
    
End Sub

Private Sub AddToTempTable(BatchID As String, FileName As String, Optional Reprocess As Boolean = False)
    Dim strSQL As String
    Dim FileType As String
    Dim FileIsMR As Integer
    Dim MRValid As Integer
    Dim PageCnt As Integer
    Dim MyAdo As clsADO
    Dim FileStatus As String
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    
    FileStatus = ""
    
    If Not fso.FileExists(strCDFilesDestination & FileName) Then
        MsgBox "Error: File does not exist", vbExclamation, "Error adding file to table"
        Exit Sub
    End If
    
    If Reprocess Then
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "DELETE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp where BatchID = '" & Me.txtBatchID & "' and FileName = '" & FileName & "' and UserID = '" & Identity.UserName() & "' and AccountID = " & gintAccountID
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute
    End If
    
    Select Case FileExtension(FileName)
        Case "DOC", "DOCX", "TXT"
            FileType = "Word"
            MRValid = 0
            FileIsMR = 0
            FileStatus = "Copied"
        Case "XLS", "XLSX"
            FileType = "Excel"
            MRValid = 0
            FileIsMR = 0
            FileStatus = "Copied"
        Case "PDF"
            FileType = "PDF Image"
            PageCnt = Count_PDF_Pages(strCDFilesDestination & FileName)
            If PageCnt <= 0 Then
                FileStatus = "ERROR: Cannot Open"
                MRValid = 0
                FileIsMR = 1
            Else
                FileStatus = "Ready to Match"
                MRValid = 1
                FileIsMR = 1
            End If
        Case "TIF", "TIFF"
            FileType = "TIF Image"
            MRValid = 1
            FileIsMR = 1
            FileStatus = "Ready to Match"
        Case Else
            FileType = "Other"
            MRValid = 0
            FileIsMR = 0
            FileStatus = "Copied"
    End Select
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "INSERT INTO cms_auditors_claims.dbo.SCANNING_CDLoad_Temp (FileName , FileType , MRValid, BatchID, FileStatus, UserID, FileIsMR, FilePageCnt, AccountID, AddedDt) " & _
                        " VALUES ('" & FileName & "', '" & FileType & "', " & MRValid & ", '" & BatchID & "', '" & FileStatus & "', '" & Identity.UserName() & "', " & FileIsMR & ", " & PageCnt & " , " & gintAccountID & ", '" & Now & "' )"
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    
    MyAdo.DisConnect
    Set MyAdo = Nothing

End Sub

Private Function NumberFilesInZip(ZipFile As String)

On Error GoTo Trap_Error

Dim sh As Object, fld As Object
NumberFilesInZip = 0
Set sh = CreateObject("Shell.Application")
Set fld = sh.Namespace(ZipFile)
NumberFilesInZip = Nz(fld.Items.Count, 0)

Trap_Error:

Set fld = Nothing
Set sh = Nothing

On Error GoTo 0

End Function



Sub Clear_All_Files_And_SubFolders_In_Folder(MyPath As String, Optional DeleteFolder As Boolean = False)
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim fso As Object

    Set fso = CreateObject("scripting.filesystemobject")

    If Right(MyPath, 1) = "\" Then
        MyPath = left(MyPath, Len(MyPath) - 1)
    End If

    If fso.FolderExists(MyPath) = False Then
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    fso.DeleteFile MyPath & "\*.*", True
    'Delete subfolders
    fso.DeleteFolder MyPath & "\*.*", True
    
    If DeleteFolder Then
        fso.DeleteFolder MyPath, True
    End If
    
    On Error GoTo 0

End Sub

Private Sub LoadCDFilesList(Optional EmptySQL As Boolean = False)

    Me.subFileList.SourceObject = "frm_SCANNING_CDLoad_SubFileList"
  
    
    Dim MyAdo As clsADO
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = " SELECT a.*, ClmStatus = b.ClmStatus + '-' + c.Clmstatusdesc FROM cms_auditors_claims.dbo.SCANNING_CDLoad_Temp a " & _
                        " LEFT JOIN CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_HDR b with (nolock) ON a.ICN = b.ICN " & _
                        " LEFT JOIN CMS_AUDITORS_CLAIMS.dbo.XREF_ClaimStatus c with (nolock) on b.clmstatus = c.clmstatus " & _
                        " where a.BatchID = '" & Me.txtBatchID & "' and a.UserID = '" & Identity.UserName() & "' and " & IIf(Me.chkHideIsImage, " ABS(a.FileIsMR) = 1 ", " 1=1 ") & " and " & IIf(EmptySQL, " 1=2 ", " 1=1 ") & " order by abs(a.FileIsMR) desc, abs(a.MRValid), a.filename"
    Set Me.subFileList.Form.RecordSet = MyAdo.OpenRecordSet
    MyAdo.DisConnect
    Set MyAdo = Nothing
    
End Sub

Private Sub chkSortClaims_AfterUpdate()
    ReloadManualMatchClaimList
End Sub

Private Sub Form_Load()
    Me.subQuickImageLog.Form("CDSubForm") = 1
    RefreshData
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not Me.btnMoveFiles.Enabled Then Exit Sub
    
    Dim UserAnswer As Integer
    UserAnswer = MsgBox("Are you sure you want to exit the CD MR Load window?" & vbNewLine & vbNewLine & _
                        "All changes will be lost", vbQuestion + vbYesNo, "MR CD Load")
    If UserAnswer = vbNo Then
        Cancel = True
    End If
                        
End Sub

Private Sub lstUnmatchedFiles_Click()
    Dim objIE As InternetExplorer
    Set objIE = Me.WebBrowser.Object
    objIE.Navigate strCDFilesDestination & Me.lstUnmatchedFiles.Column(0, Me.lstUnmatchedFiles.ListIndex)
End Sub
Private Sub TabCtlMain_Change()

    DoCmd.Hourglass False

    Set Me.lstUnmatchedClaims.RecordSet = Nothing
    Set Me.lstUnmatchedFiles.RecordSet = Nothing
    Me.lstUnmatchedClaims = -1
    Me.lstUnmatchedFiles = -1
    
    Select Case TabCtlMain.Value
        Case 2 'Manual
            Me.WebBrowser.Requery
            Me.txtMemberName = ""
            Me.txtDOB = ""
    
            Dim MyAdo As clsADO
            
            Set MyAdo = New clsADO
            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
            MyAdo.sqlString = "SELECT FileName FROM cms_auditors_claims.dbo.SCANNING_CDLoad_Temp where FileIsMR = 1 and MRValid = 0 and BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and left(filestatus,5) <> 'ERROR' order by filename"
            Set Me.lstUnmatchedFiles.RecordSet = MyAdo.OpenRecordSet
            
            ReloadManualMatchClaimList
            
            MyAdo.DisConnect
            Set MyAdo = Nothing
        Case 1 'auto
       
            strCDFilesDestination = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\CD_DVD\_TEMP\" & Identity.UserName() & "\" & Me.txtBatchID & "\"
            strCDSaveDestination = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\CD_DVD\_PROCESSED\" & Me.txtCnlyProvID & "\" & Me.txtBatchID & "\"
            strMRLoadDestination = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\DailyScans\" & Me.txtCnlyProvID & "\"
            If Nz(Me.txtCDSourceFolder, "") = "" Then Me.txtCDSourceFolder = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\MEDICALRECORD_TEMP\DailyScans\"
            AddUnMatchedClaimsToTempTable
            LoadCDFilesList
            CalcTotals
    End Select
    
    CalcTotals

End Sub


Sub ManualMatch(FileName As String, Icn As String, ImageName As String)

    
    Dim NewName As String
    NewName = ""
    Call RenameWithICN(strCDFilesDestination, FileName, Icn, NewName)
    
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = " UPDATE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp " & _
                        " SET FileName = '" & NewName & "', ICN = '" & Icn & "', ImageName = '" & ImageName & "', MRValid = 1, FileStatus = 'Matched M'" & _
                        " where BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and filename = '" & FileName & "' and AccountID = " & gintAccountID & _
                        " DELETE cms_auditors_claims.dbo.SCANNING_CDLoad_Temp WHERE BatchID = '" & Me.txtBatchID & "' and UserID = '" & Identity.UserName() & "' and MRValid = -1 and FileIsMR = -1 and ICN = '" & Icn & "' and AccountID = " & gintAccountID
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
    MyAdo.DisConnect
    Set MyAdo = Nothing
    
End Sub

Private Function RenameWithICN(Folder As String, FileName As String, Icn As String, ByRef NewFileName As String) As Boolean


    Dim Extension As String


    RenameWithICN = False
   
    Extension = FileExtension(FileName)
    NewFileName = Icn & "." & Extension
    
    Name strCDFilesDestination & FileName As strCDFilesDestination & NewFileName
    
    If strCDFilesDestination & FileName = strCDFilesDestination & NewFileName Then
        'nothing to do
        RenameWithICN = True
        Exit Function
    End If
    
    If FileExists(strCDFilesDestination & NewFileName) Then RenameWithICN = True
  
End Function

Sub CalcTotals()
    
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    Dim rs As ADODB.RecordSet
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = " select FilesImage = sum(case when FileIsMR = 1 then 1 else 0 end), FilesNotImage = SUM(case when FileIsMR = 0 then 1 else 0 end), FilesMatched = SUM(case when MRValid = 1 then 1 else 0 end), FilesNotMatched = SUM(case when (mrvalid = 0 and FileisMR = 1) then 1 else 0 end), ClaimsNotMatched = SUM(Case when (FileIsMR = -1 and MRValid = -1) then 1 else 0 end)" & _
                      " from cms_auditors_claims.dbo.SCANNING_CDLoad_Temp a " & _
                      " where batchid = '" & Me.txtBatchID & "' and userid = '" & Identity.UserName() & "'"
    Set rs = MyAdo.OpenRecordSet
        
    If Not rs Is Nothing Then
    
        If rs.recordCount > 0 Then
            Me.txtFilesImage = Nz(rs("FilesImage"), 0)
            Me.txtFilesNotImage = Nz(rs("FilesNotImage"), 0)
            Me.txtFilesMatched = Nz(rs("FilesMatched"), 0)
            Me.txtFilesNotMatched = Nz(rs("FilesNotMatched"), 0)
            Me.txtClaimsNotMatched = Nz(rs("ClaimsNotMatched"), 0)
            Me.txtTotalFiles = Me.txtFilesImage + Me.txtFilesNotImage
            Me.txtTotalClaims = Me.txtFilesMatched + Me.txtClaimsNotMatched
        End If
        
    End If
    
    Set rs = Nothing
    MyAdo.DisConnect
    Set MyAdo = Nothing
    
    If Me.txtFilesNotMatched = 0 Then
        Me.txtFilesNotMatched.ForeColor = 16384
    Else
        Me.txtFilesNotMatched.ForeColor = 255
    End If
    
    If Me.txtClaimsNotMatched = 0 Then
        Me.txtClaimsNotMatched.ForeColor = 16384
    Else
        Me.txtClaimsNotMatched.ForeColor = 255
    End If
    
    
    If Me.TabCtlMain = 1 Then
        If Me.txtFilesNotMatched > 0 And Me.txtClaimsNotMatched > 0 Then Me.tabManual.visible = True Else Me.tabManual.visible = False
    End If
    
End Sub

Sub ReloadManualMatchClaimList()
    Dim strSelectedOrder As String
    
    Select Case Me.chkSortClaims
        Case 1 'ICN
            strSelectedOrder = "ICN"
        Case 2 'Name
            strSelectedOrder = "MemberName"
        Case 3 'DOB
            strSelectedOrder = "DOB"
        Case 4 'ClmFromDt (default)
            strSelectedOrder = "ClmFromDt"
    End Select
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = " SELECT t1.ICN, t1.MemberName, convert(date, t1.DOB), t1.ClmFromDt, t1.imagename " & _
                     " FROM cms_auditors_claims.dbo.SCANNING_Quick_Image_Log as t1 " & _
                     " INNER JOIN cms_auditors_claims.dbo.SCANNING_Image_Log_Tmp as t2 " & _
                     " ON t1.ScannedDt=t2.ScannedDt " & _
                     " left join cms_auditors_claims.dbo.SCANNING_CDLoad_Temp as t3 " & _
                     " on t1.ICN = t3.ICN and t1.sessionid = t3.batchid " & _
                     " where t1.sessionid = '" & Me.txtBatchID & "' and (t3.MRValid = -1 and t3.FileIsMR = -1) " & _
                     " and t1.MemberName like '" & IIf(Nz(Me.txtMemberName, "") = "", "%", Me.txtMemberName & "%") & "' " & _
                     " and convert(date, t1.DOB) like '" & IIf(Nz(Me.txtDOB, "") = "", "%", Me.txtDOB & "%") & "' " & _
                     " order by t1." & strSelectedOrder
    Set Me.lstUnmatchedClaims.RecordSet = MyAdo.OpenRecordSet
    MyAdo.DisConnect
    Set MyAdo = Nothing
End Sub

Private Sub txtDOB_Exit(Cancel As Integer)
    ReloadManualMatchClaimList
End Sub

Private Sub txtMemberName_Exit(Cancel As Integer)
    ReloadManualMatchClaimList
End Sub


Private Function FreeMegaBytes(NetworkShare As String) As Long

Dim curBytesFreeToCaller As Currency
Dim curTotalBytes As Currency
Dim curTotalFreeBytes As Currency


Call GetDiskFreeSpaceEx(NetworkShare, _
curBytesFreeToCaller, _
curTotalBytes, _
curTotalFreeBytes)

'show the results, multiplying the returned
'value by 10000 to adjust for the 4 decimal
'places that the currency data type returns.

'Debug.Print " Free Bytes Available:"
FreeMegaBytes = ((curTotalFreeBytes * 10000) / 1000) / 1000

'Print FreeBytes

End Function






Private Function FileExtension(FileName As String) As String
    If Trim(FileName) = "" Then
        Exit Function
    End If
    FileExtension = UCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".")))
End Function


Private Function FileWithoutExtension(FileName As String) As String
    If Trim(FileName) = "" Then
        Exit Function
    End If
    If FileName <> "---" Then
        FileWithoutExtension = left(FileName, Len(FileName) - (Len(FileName) - InStrRev(FileName, ".")) - 1)
    End If
End Function

Private Function FolderWithoutFile(FullPath As String) As String
    If Trim(FullPath) = "" Then
        Exit Function
    End If
    FolderWithoutFile = left(FullPath, Len(FullPath) - (Len(FullPath) - InStrRev(FullPath, "\")) - 1)
End Function


Function IsZipFile(FileName As String) As Boolean
    Dim Extension As String
    Extension = FileExtension(FileName)
    Select Case Extension
        Case "ZIP"
            IsZipFile = True
        Case "7Z"
            IsZipFile = True
        Case "ZIPX"
            IsZipFile = True
        Case "RAR"
            IsZipFile = True
        Case "ARC"
            IsZipFile = True
        Case Else
            IsZipFile = False
    End Select
End Function


Function BrowseFolder(Title As String, _
        Optional InitialFolder As String = vbNullString, _
        Optional InitialView As Office.MsoFileDialogView = _
            msoFileDialogViewList) As String
    Dim V As Variant
    Dim InitFolder As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .show
        On Error Resume Next
        Err.Clear
        V = .SelectedItems(1)
        If Err.Number <> 0 Then
            V = vbNullString
        End If
    End With
    BrowseFolder = CStr(V)
End Function
