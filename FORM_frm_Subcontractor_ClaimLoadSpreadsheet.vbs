Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean
Dim miError As Long
Dim miFileCnt As Long

Dim mstrFileToLoad As String
Dim mstrStagingFolder As String


Const CstrFrmAppID As String = "SubContractor"


Private Sub LoadSpreadsheet(strpExcelFileToLoad As String)

'    Dim djEng As New DJEC.Engine
'    Dim djConv As New DJEC.Conversion
'    Dim djLog As New DJEC.LogManager
    Dim strCosmosMapName As String
    Dim strSourceFileName As String
    
    strSourceFileName = strpExcelFileToLoad
    strCosmosMapName = "\\ccaintranet.com\dfs-cms-ds\Data\CMS\Cosmos\Subcontractor\ReviewResults_Import.tf.xml"


'    djConv.Load strCosmosMapName
'
''        Select Case Mode
''            Case Is = ""
''                '* do nothing
''            Case Is = "Append"
''                 djConv.Targets(0).OutputMode = omAPPEND
''            Case Is = "DeleteAppend"
''                djConv.Targets(0).OutputMode = omDELAPPEND
''            Case Else
''                MsgBox "Map output mode error"
''
''        End Select
'
'    If strSourceFileName > "" Then
'         djConv.Sources(0).ConnectionInfo.file = strSourceFileName
'         Set djLog = djConv.MessageLog
'         djLog.FileName = strSourceFileName & ".err"
'    End If
'
'     djConv.Run

         
    Dim iErrorCount As Integer
    Dim lRecordCount As Long
    Dim strTargetFolder As String


' 1. Creating an engine object and its members.
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


' 2. Load map
    djConversion.Load (strCosmosMapName)
    
    
    
'3a. SOURCE Connection data
    djConversion.Sources(0).connectioninfo.Database = strSourceFileName
    djConversion.Sources(0).connectioninfo.Table = "Sheet1"
'    Debug.Print djConversion.Sources(0).SpokeTypeName
    Debug.Print djConversion.Sources(0).connectioninfo.Database
'    Debug.Print djConversion.Sources(0).connectioninfo.Table



'' 3b. TARGET Connection data
'   'Connect the target
'    djConversion.Targets(0).Connect
'    Debug.Print djConversion.Targets(0).connectioninfo.Server
'    Debug.Print djConversion.Targets(0).connectioninfo.Database
'    Debug.Print djConversion.Targets(0).connectioninfo.Table

'
'
'' 4. Running map conversion
    djConversion.Run
'    iErrorCount = djConversion.ErrorCount
'    lRecordCount = djConversion.WriteCount


    lblStatus.Caption = "Loaded: " & djConversion.WriteCount & " from file " & mstrFileToLoad
    
    
Exit_Sub:
    Set djEngine = Nothing
    Set djConversion = Nothing
    Set djLog = Nothing
    Exit Sub


End Sub

Private Sub ClearTempTable()
    Dim MyAdo As clsADO

    ' Set error handler
    On Error GoTo Err_handler


    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = " delete from CMS_AUDITORS_Workspace.dbo.ClaimLoad_ReviewResult_temp "
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute


Exit_Sub:
    Set MyAdo = Nothing
    Exit Sub

Err_handler:
    MsgBox "Error in module " & vbCrLf & vbCrLf & Err.Description
    Resume Exit_Sub
End Sub

Private Sub TransferFiles()


'=========================================================================================
' By: Tuan Khong
' Date: 5\10\2011
' This code is to transfer Medical Records via 3 steps below
'=========================================================================================
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rs As ADODB.RecordSet
    Dim rsImageReference As ADODB.RecordSet
    Dim rsImageLog As ADODB.RecordSet
    Dim cmd As ADODB.Command

    Dim strSQL As String
    Dim strFileName As String
    Dim strOldFolder As String
    Dim strNewFolder As String

    Dim strSourceFolder As String
    Dim strSourceImagePath As String

    Dim strDestinationImagePath As String
    
    Dim oSourceFolder As Folder
    Dim oDestinationFolder As Folder

    Dim strErrMsg As String
    Dim lngCounter As Long
    Dim lngCountTotal As Long
    Dim intFilesToTransfer As Integer
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Set error handler
    On Error GoTo Err_handler

 
    ' TEST - to remove
'    intFilesToTransfer = 0
'    If IsNumeric(txtTransferAmount.Value) Then
'        intFilesToTransfer = txtTransferAmount
'    Else
'        MsgBox "Enter numeric value 1 to 500."
'        Exit Sub
'    End If

    mstrStagingFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Subcontractor\" & Format(Date, "yyyymmdd") & "\"
 

'=========================================================================================
'   5\13\2011 Copy files to loading folder
'   TK: Set SOURCE folder for NJPR and Nurse Audit. To dynamically define this in future instead of hardcode.
'=========================================================================================
'    Dim bFileToTransfer As Boolean
'    Set bFileToTransfer = False
    ' Set DESTINATION folder
    
    
    If Not fso.FolderExists(mstrStagingFolder) Then
        Call CreateFolder(mstrStagingFolder)
    End If
    
    ' Exit sub if folder still does not exist
    If Not fso.FolderExists(mstrStagingFolder) Then
        MsgBox "Error creating destination folder: " & _
        mstrStagingFolder & " . No files transferred. ", vbCritical
        Exit Sub
    Else
        Set oDestinationFolder = fso.GetFolder(mstrStagingFolder)
    End If
    
    
    ' Nurse Audit files transfer
    mbContinue = True
    strSourceFolder = "\\ccaintranet.com\DFS-MCR-02\Audits\NurseAudit\EXCEL\"
    If fso.FolderExists(strSourceFolder) Then
        Set oSourceFolder = fso.GetFolder(strSourceFolder)
        Call ScanFile(oSourceFolder, oDestinationFolder)
    Else
        MsgBox "Source folder does not exist: " & _
        strSourceFolder & " . No files transfered from this folder."
    End If

    
'    ' NJPR files transfer
'    mbContinue = True
'    strSourceFolder = "\\ccaintranet.com\DFS-MCR-01\Audits\NJPR\Excel\"
'    If fso.FolderExists(strSourceFolder) Then
'        Set oSourceFolder = fso.GetFolder(strSourceFolder)
'        Call ScanFile(oSourceFolder, oDestinationFolder)
'    Else
'        MsgBox "Source folder does not exist: " & _
'        strSourceFolder & " . No files transfered from this folder."
'    End If


    ' MRC files transfer
    mbContinue = True
    strSourceFolder = "\\ccaintranet.com\DFS-MCR-03\Audits\MRC\Excel\"
    If fso.FolderExists(strSourceFolder) Then
        Set oSourceFolder = fso.GetFolder(strSourceFolder)
        Call ScanFile(oSourceFolder, oDestinationFolder)
    Else
        MsgBox "Source folder does not exist: " & _
        strSourceFolder & " . No files transfered from this folder."
    End If
    

    MsgBox "Done", vbOKOnly
    
    ' TEST
    Exit Sub
'
''=========================================================================================
''1. Insert new MR records to VIANT_EXPORT_MR_Log, set StatusFlag = 0
''=========================================================================================
'
'   ' Insert new MR records to VIANT_EXPORT_MR_Log, StatusFlag = 0 as default
'    Set myCODE_Ado = New clsADO
'    myCODE_Ado.ConnectionString = GetConnectString("v_CODE_Database")
'    myCODE_Ado.SQLstring = " exec usp_Scanning_MoveMR_Log_INSERT "
'    myCODE_Ado.SQLTextType = sqltext
'    myCODE_Ado.Execute
'
'
''=========================================================================================
''2. Copy all medical records from Scanning_NonRecoveryMR_Log table with status = 0
''=========================================================================================
'    ' Select all medical records from  with status = 0
'    Set MYADO = New clsADO
'    MYADO.ConnectionString = GetConnectString("v_DATA_Database")
'
'    ' Select non-processed records from VIANT_EXPORT_MR_Log table
'    strsql = " SELECT top " & intFilesToTransfer & " * FROM CMS_AUDITORS_Claims.dbo.Scanning_NonRecoveryMR_Log t1 WHERE t1.StatusFlag = 0 "
'    Set rs = MYADO.OpenRecordSet(strsql)
'
'    ' Exit if record set is empty
'    If rs.BOF = True And rs.EOF = True Then
'        MsgBox "Nothing to do.", vbOKOnly
'        GoTo Exit_Sub
'    End If
'
'
'    lngCounter = 0
'    lngCountTotal = rs.RecordCount
'    ' Loop record set and copy file to staging area
'    rs.MoveFirst
'    While Not rs.EOF
'        lngCounter = lngCounter + 1
'
'        lblLoading.Caption = "Transfering files " & lngCounter & " of " & lngCountTotal
'        lblLoading.visible = True
'        DoEvents
'
'        ' Getting file & folder info
'        strSourceImagePath = rs!SourceImagePath
'        strDestinationImagePath = Replace(strSourceImagePath, strOldFolder, strNewFolder)
'        strFileName = fso.GetFileName(strDestinationImagePath)
'        mstrStagingFolder = left(strDestinationImagePath, Len(strDestinationImagePath) - Len(strFileName))
'
'        ' Check if source file exists
'        If Not fso.FileExists(strSourceImagePath) Then
'            'lstFiles.AddItem "Image Not Ready;" & strFileName
'            GoTo NextImage
'        End If
'
'        ' Check folder
'        If Not fso.FolderExists(mstrStagingFolder) Then
'            Call CreateFolder(mstrStagingFolder)
'        End If
'
'        ' See if destination file exists
'        If fso.FileExists(mstrStagingFolder) Then
'
'        Else
'            Call fso.CopyFile(strSourceImagePath, strDestinationImagePath, False)
'        End If
'
'NextImage:         rs.MoveNext
'    Wend
'
'
'
'''=========================================================================================
'''3. UPDATE Scanning_NonRecoveryMR_Log, set StatusFlag = 1. Status 1 = File transferred
'''=========================================================================================
'
'    Set rs = MYADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
'    If Not (rs.BOF = True And rs.EOF = True) Then
'        lngCounter = 0
'        lngCountTotal = rs.RecordCount
'        rs.MoveFirst
'        While Not rs.EOF
'            lngCounter = lngCounter + 1
'
'            lblLoading.Caption = "Updating files " & lngCounter & " of " & lngCountTotal
'            lblLoading.visible = True
'            DoEvents
'
'            ' Setting destination file path
'            strSourceImagePath = rs!SourceImagePath
'            strDestinationImagePath = Replace(strSourceImagePath, strOldFolder, strNewFolder)
'
'            ' Update table if image exist
'            If fso.FileExists(strDestinationImagePath) = True Then
'
''                Set rsImageReference = myADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
''                Set rsImageLog = myADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
'
'
'                rs!DestinationImagePath = strDestinationImagePath
'                rs!StatusFlag = 1
'                rs!Notes = "Image transferred on " & Date
'
'                'Debug.Print mstrRemotePath & strFileName
'            End If
'        rs.Update
'        rs.MoveNext
'        Wend
'    End If
'    MYADO.BatchUpdate rs
'
'
'
'
'''=========================================================================================
''' 4. VERIFY new image and DELETE old image StatusFlag = 2
'''=========================================================================================
'    Set rs = MYADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 1 ")
'    If Not (rs.BOF = True And rs.EOF = True) Then
'        lngCounter = 0
'        lngCountTotal = rs.RecordCount
'        rs.MoveFirst
'        While Not rs.EOF
'            lngCounter = lngCounter + 1
'
'            lblLoading.Caption = "Updating files " & lngCounter & " of " & lngCountTotal
'            lblLoading.visible = True
'            DoEvents
'
'            ' Setting destination file path
'            strSourceImagePath = rs!SourceImagePath
'            strDestinationImagePath = rs!DestinationImagePath
'
'            ' Update table if image exist
'            If fso.FileExists(strDestinationImagePath) = True Then
'                rs!StatusFlag = 2
'                rs!Notes = "Image verified on " & Date
'
'                ' Delete old image
'                If fso.FileExists(strSourceImagePath) = True Then
'                    fso.DeleteFile strSourceImagePath, True
'                End If
'
'
'            Else
'                rs!StatusFlag = 1
'                rs!Notes = "Image not verify. Need to re-transfer. " & Date
'            End If
'        rs.Update
'        MYADO.BatchUpdate rs
'        rs.MoveNext
'        Wend
'    End If
'
'
'
''=========================================================================================
'' 5. UPDATE AUDITCLM_References and SCANNING_Image_Log with new file path
''=========================================================================================
''INCOMPLETE
'
''    Update all the 2's
'    myCODE_Ado.SQLstring = " exec usp_Scanning_MoveMR_Log_UPDATE "
'    myCODE_Ado.SQLTextType = sqltext
'    myCODE_Ado.Execute
'
'

''=========================================================================================
'' END usp_Scanning_MoveMR_Log_UPDATE
''=========================================================================================




Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set fso = Nothing
    Set oSourceFolder = Nothing
    Set oDestinationFolder = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    Set rs = Nothing
    Set rsImageReference = Nothing
    Set rsImageLog = Nothing
    
    Exit Sub

Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    Resume Exit_Sub

'' =========================================================================================
'
End Sub



Private Function ScanFile(oSourceFolder As Folder, oDestinationFolder As Folder) As Boolean
    Dim oSubFolder As Folder
    Dim oFile As file
    Dim fso As FileSystemObject
    Dim strSourceFilePath As String
    Dim strDestinationFilePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Err_handler
'test
Debug.Print oSourceFolder.Path
    
    For Each oFile In oSourceFolder.Files
        If UCase(Right(oFile.Name, 4)) = ".xls" Or UCase(Right(oFile.Name, 4)) = ".xlsx" Then
            Debug.Print "File modified on " & oFile.DateLastModified
            Debug.Print "Date is " & Date
            Debug.Print DateDiff("d", oFile.DateLastModified, Date)
            
            If DateDiff("d", oFile.DateLastModified, Date) < 5 Then
                strSourceFilePath = oFile.Path
                strDestinationFilePath = oDestinationFolder.Path & "\" & oSourceFolder.Name & "_" & oFile.Name
                Call fso.CopyFile(strSourceFilePath, strDestinationFilePath, True)
                
                lblStatus.Caption = "Loading file: " & strSourceFilePath
                lblStatus.visible = True
                DoEvents
                DoEvents
            End If

'test
Debug.Print oFile.Path
        End If
    Next
    
    For Each oSubFolder In oSourceFolder.SubFolders
Debug.Print oSourceFolder.SubFolders.Count
        mbContinue = ScanFile(oSubFolder, oDestinationFolder)
Debug.Print oSourceFolder.Path
        If mbContinue = False Then Exit For
        DoEvents
        DoEvents
    Next
    
    Set oFile = Nothing
    Set oSubFolder = Nothing
    Set fso = Nothing
    
    ScanFile = mbContinue
    
    Exit Function

Err_handler:
    MsgBox Err.Number & " -- " & Err.Description
    MsgBox oFile.Path
    mbContinue = False
    ScanFile = mbContinue
End Function


Private Sub cmdCopyFile_Click()

    ' reset list
    lstFilesToLoad.RowSource = ""
    lstFilesToLoad.Requery
    
    ' Display loading
    lblStatus.Caption = "Loading file: " & mstrFileToLoad
    lblStatus.visible = True
    DoEvents
    
    ' Load claims via cosmos
    TransferFiles
    DoEvents
    DoEvents
    DoEvents
    
    
    RefreshList
    lblStatus.Caption = "Complete loading. "
    
    
End Sub

Private Sub cmdLoadAllFiles_Click()
    Dim oFolder As Folder
    Dim oFile As file
    Dim fso As FileSystemObject
    Dim intCounter As Integer
    Dim intTotalFileCount As Integer
    Dim dPercentage As Double
    
    
    ' Clear table first
    ClearTempTable
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set folder
    Set oFolder = fso.GetFolder(mstrStagingFolder)
    intCounter = 1
    intTotalFileCount = lstFilesToLoad.ListCount
    If intTotalFileCount > 0 Then
        For Each oFile In oFolder.Files
            If UCase(Right(oFile.Name, 4)) = ".xls" Or UCase(Right(oFile.Name, 4)) = ".xlsx" Then
                mstrFileToLoad = oFile.Path
                dPercentage = (intCounter / intTotalFileCount) * 100

                lblFileCount.Caption = "Loading file " & intCounter & " of " & intTotalFileCount & ".  " & Chr(13) & _
                                        Round(dPercentage, 0) & "% complete..."
                Debug.Print "Loading file " & intCounter & " of " & intTotalFileCount & ".  " & Chr(13) & _
                                        Round(dPercentage, 0) & "% complete..."
                ' Call Cosmos map to load single file
                LoadSpreadsheet (mstrFileToLoad)
                DoEvents
                DoEvents
                DoEvents
    
            End If
            intCounter = intCounter + 1
        Next
    End If

    GoTo Exit_Sub
    
Exit_Sub:
    Set oFolder = Nothing
    Set oFile = Nothing
    Set fso = Nothing
        
    Exit Sub

Err_handler:
    MsgBox Err.Number & " -- " & Err.Description
    MsgBox oFile.Path
    Resume Exit_Sub
End Sub



Private Sub cmdRefreshList_Click()
    ' Refresh file list
    RefreshList
    
    lblStatus.visible = False
    DoEvents
    
End Sub



Private Sub RefreshList()
    Dim oFolder As Folder
    Dim oFile As file
    Dim fso As FileSystemObject
    Dim strSourceFilePath As String
    Dim strDestinationFilePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' reset list
    lstFilesToLoad.RowSource = ""
    lstFilesToLoad.Requery
    
    ' Set folder
    ' Set oFolder = fso.GetFolder(mstrStagingFolder)
    Set oFolder = fso.GetFolder(mstrStagingFolder)
    
    For Each oFile In oFolder.Files
        If UCase(Right(oFile.Name, 4)) = ".xls" Or UCase(Right(oFile.Name, 4)) = ".xlsx" Then
            strSourceFilePath = oFile.Path
            strDestinationFilePath = oFile.Path
            ' Add file to list
            lstFilesToLoad.AddItem strDestinationFilePath
            lblFileCount.Caption = "Files count: " & lstFilesToLoad.ListCount
            DoEvents
            DoEvents

        End If
    Next
    

    lblFileCount.Caption = "Total count: " & lstFilesToLoad.ListCount
    
    GoTo Exit_Sub
    
Exit_Sub:
    Set oFolder = Nothing
    Set oFile = Nothing
    Set fso = Nothing
        
    Exit Sub

Err_handler:
    MsgBox Err.Number & " -- " & Err.Description
    MsgBox oFile.Path
    Resume Exit_Sub
    
End Sub


Private Sub Form_Load()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' reset display info
    lblStatus.visible = False
    mstrFileToLoad = ""
    lstFilesToLoad.RowSource = ""
    lstFilesToLoad.Requery
    
    mstrStagingFolder = "\\ccaintranet.com\dfs-cms-ds\Raw\CMS\Subcontractor\" & Format(Date, "yyyymmdd") & "\"
    If Not fso.FolderExists(mstrStagingFolder) Then
        Call CreateFolder(mstrStagingFolder)
    End If
    
    ' loading default folder
    RefreshList
    
    
    
End Sub

Private Sub lstFilesToLoad_DblClick(Cancel As Integer)
    Dim strFileName As String
    Dim strImagePath As String
    
    ' Assigning file to load
    mstrFileToLoad = lstFilesToLoad.Value



    ' Display loading
    lblStatus.Caption = "Loading file: " & mstrFileToLoad
    lblStatus.visible = True
    DoEvents
    
    ' Call Cosmos map to load single file
    LoadSpreadsheet (mstrFileToLoad)
    DoEvents

    MsgBox "Done.", vbOKOnly

End Sub
