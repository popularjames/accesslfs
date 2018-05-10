Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean
Dim miError As Long
Dim miFileCnt As Long

Dim mstrLocalHoldPath As String
Dim mstrLocalPath As String
Dim mstrRemotePath As String
Dim mstrHoldImageName As String
Dim mstrLocalImageName As String
Dim mstrRemoteImageName As String
Dim mstrLocalImagePath As String
Dim mstrRemoteImagePath As String

Dim mstrCalledFrom As String

Const CstrFrmAppID As String = "SubContractor"
Private Sub Command0_Click()
    ' Display loading
    lblLoading.visible = True
    
    DoEvents
    ' Transfer file to folder to be process by Cosmos
     TransferFiles
    
    
    lblLoading.visible = False
End Sub


Private Sub TransferFiles()
'=========================================================================================
' By: Tuan Khong
' Date: 9/10/2010
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
    Dim strSourceImagePath As String
    Dim strDestinationImagePath As String
    Dim strDestinationFolder As String
    Dim oFolder As Folder
    Dim strErrMsg As String
    
    Dim lngCounter As Long
    Dim lngCountTotal As Long
    Dim intFilesToTransfer As Integer
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Set error handler
    On Error GoTo Err_handler


 
' Check
    intFilesToTransfer = 0
    If IsNumeric(txtTransferAmount.Value) Then
        intFilesToTransfer = txtTransferAmount
    Else
        MsgBox "Enter numeric value 1 to 500."
        Exit Sub
    End If

    ' Set folder variable to replace
    strOldFolder = "\\ccaintranet.com\dfs-fld-01\Imaging\Client\Out\CMS\MedicalRecords_Current\"
    strNewFolder = "\\cca-audit\dfs-fld-01\Imaging\Client\Out\CMS\MedicalRecords_Archive\"


'=========================================================================================
'1. Insert new MR records to VIANT_EXPORT_MR_Log, set StatusFlag = 0
'=========================================================================================
    
   ' Insert new MR records to VIANT_EXPORT_MR_Log, StatusFlag = 0 as default
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = " exec usp_Scanning_MoveMR_Log_INSERT "
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.Execute
   

'=========================================================================================
'2. Copy all medical records from Scanning_NonRecoveryMR_Log table with status = 0
'=========================================================================================
    ' Select all medical records from  with status = 0
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    ' Select non-processed records from VIANT_EXPORT_MR_Log table
    strSQL = " SELECT top " & intFilesToTransfer & " * FROM CMS_AUDITORS_Claims.dbo.Scanning_NonRecoveryMR_Log t1 WHERE t1.StatusFlag = 0 "
    Set rs = MyAdo.OpenRecordSet(strSQL)
    
    ' Exit if record set is empty
    If rs.BOF = True And rs.EOF = True Then
        MsgBox "Nothing to do.", vbOKOnly
        GoTo Exit_Sub
    End If
    
    
    lngCounter = 0
    lngCountTotal = rs.recordCount
    ' Loop record set and copy file to staging area
    rs.MoveFirst
    While Not rs.EOF
        lngCounter = lngCounter + 1
        
        lblLoading.Caption = "Transfering files " & lngCounter & " of " & lngCountTotal
        lblLoading.visible = True
        DoEvents
        
        ' Getting file & folder info
        strSourceImagePath = rs!SourceImagePath
        strDestinationImagePath = Replace(strSourceImagePath, strOldFolder, strNewFolder)
        strFileName = fso.GetFileName(strDestinationImagePath)
        strDestinationFolder = left(strDestinationImagePath, Len(strDestinationImagePath) - Len(strFileName))
        
        ' Check if source file exists
        If Not fso.FileExists(strSourceImagePath) Then
            'lstFiles.AddItem "Image Not Ready;" & strFileName
            GoTo NextImage
        End If
        
        ' Check folder
        If Not fso.FolderExists(strDestinationFolder) Then
            Call CreateFolder(strDestinationFolder)
        End If
        
        ' See if destination file exists
        If fso.FileExists(strDestinationFolder) Then
            
        Else
            Call fso.CopyFile(strSourceImagePath, strDestinationImagePath, False)
        End If
        
NextImage:         rs.MoveNext
    Wend



''=========================================================================================
''3. UPDATE Scanning_NonRecoveryMR_Log, set StatusFlag = 1. Status 1 = File transferred
''=========================================================================================

    Set rs = MyAdo.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
    If Not (rs.BOF = True And rs.EOF = True) Then
        lngCounter = 0
        lngCountTotal = rs.recordCount
        rs.MoveFirst
        While Not rs.EOF
            lngCounter = lngCounter + 1
            
            lblLoading.Caption = "Updating files " & lngCounter & " of " & lngCountTotal
            lblLoading.visible = True
            DoEvents
            
            ' Setting destination file path
            strSourceImagePath = rs!SourceImagePath
            strDestinationImagePath = Replace(strSourceImagePath, strOldFolder, strNewFolder)
        
            ' Update table if image exist
            If fso.FileExists(strDestinationImagePath) = True Then
            
'                Set rsImageReference = myADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
'                Set rsImageLog = myADO.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 0 ")
                
                
                rs!DestinationImagePath = strDestinationImagePath
                rs!StatusFlag = 1
                rs!Notes = "Image transferred on " & Date

                'Debug.Print mstrRemotePath & strFileName
            End If
        rs.Update
        rs.MoveNext
        Wend
    End If
    MyAdo.BatchUpdate rs




''=========================================================================================
'' 4. VERIFY new image and DELETE old image StatusFlag = 2
''=========================================================================================
    Set rs = MyAdo.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..Scanning_NonRecoveryMR_Log where StatusFlag = 1 ")
    If Not (rs.BOF = True And rs.EOF = True) Then
        lngCounter = 0
        lngCountTotal = rs.recordCount
        rs.MoveFirst
        While Not rs.EOF
            lngCounter = lngCounter + 1
            
            lblLoading.Caption = "Updating files " & lngCounter & " of " & lngCountTotal
            lblLoading.visible = True
            DoEvents
            
            ' Setting destination file path
            strSourceImagePath = rs!SourceImagePath
            strDestinationImagePath = rs!DestinationImagePath
        
            ' Update table if image exist
            If fso.FileExists(strDestinationImagePath) = True Then
                rs!StatusFlag = 2
                rs!Notes = "Image verified on " & Date
                
                ' Delete old image
                If fso.FileExists(strSourceImagePath) = True Then
                    fso.DeleteFile strSourceImagePath, True
                End If
                
                
            Else
                rs!StatusFlag = 1
                rs!Notes = "Image not verify. Need to re-transfer. " & Date
            End If
        rs.Update
        MyAdo.BatchUpdate rs
        rs.MoveNext
        Wend
    End If
    


'=========================================================================================
' 5. UPDATE AUDITCLM_References and SCANNING_Image_Log with new file path
'=========================================================================================
'INCOMPLETE

'    Update all the 2's
    myCode_ADO.sqlString = " exec usp_Scanning_MoveMR_Log_UPDATE "
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.Execute

    

''=========================================================================================
'' END usp_Scanning_MoveMR_Log_UPDATE
''=========================================================================================

       MsgBox "Done", vbOKOnly



Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set fso = Nothing
    Set oFolder = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    Set rs = Nothing
    Set rsImageReference = Nothing
    Set rsImageLog = Nothing
    
    Exit Sub

Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub

'' =========================================================================================
'
End Sub

Private Sub Form_Load()
    lblLoading.visible = False
End Sub
