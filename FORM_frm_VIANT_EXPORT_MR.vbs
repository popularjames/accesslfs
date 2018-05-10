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
' Date: 5/13/2010
' This code is to transfer Medical Records to Viant via 3 steps below
'=========================================================================================
    Dim MyAdo As clsADO
    Dim myCode_ADO As clsADO
    Dim rs As ADODB.RecordSet
    Dim rsViantMR As ADODB.RecordSet
    Dim cmd As ADODB.Command
    Dim strSQL As String
    Dim strFile As String
    Dim oFolder As Folder
    Dim strErrMsg As String
    Dim intCount As Integer
    ' Set error handler
    On Error GoTo Err_handler

'=========================================================================================
'1. Insert new MR records to VIANT_EXPORT_MR_Log, set StatusFlag = 0
'=========================================================================================
    
   ' Insert new MR records to VIANT_EXPORT_MR_Log, StatusFlag = 0 as default
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = " exec usp_VIANT_MR_Log_INSERT "
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.Execute
   
' COMPLETE step 1



'=========================================================================================
'2. Copy all medical records from VIANT_EXPORT_MR_Log table with status = 0
'=========================================================================================
    ' Select all medical records from  with status = 0
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    ' Select non-processed records from VIANT_EXPORT_MR_Log table
    strSQL = " SELECT * FROM CMS_AUDITORS_Claims..VIANT_EXPORT_MR_Log t1 WHERE t1.StatusFlag != 1 "
    Set rs = MyAdo.OpenRecordSet(strSQL)
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Set staging folder location
    ' OLD folder mstrRemotePath = "\\cca-audit\dfs-dc-01\Imaging\Misc\Hold\In\CMS\MedicalRecords\VIANT\"
    ' 9/28 removed: mstrRemotePath = "\\cca-audit\DFS-DC-01\Raw\FS02-J\CMS\Viant\MR_Export_Staging\"
    mstrRemotePath = "\\ccaintranet.com\dfs-cms-ds\Data\CMS\EXPORTS\VIANT\MR_Export_Staging\"
    
    

'    If fso.FolderExists(mstrLocalPath) = False Then
'        strErrMsg = "Local image path '" & mstrLocalPath & "' does not exists or in accessible.  Please check."
'        GoTo Err_handler
'    End If
'

    ' Exit if record set is empty
    If rs.BOF = True And rs.EOF = True Then
        MsgBox "Nothing to do.", vbOKOnly
        GoTo Exit_Sub
    End If
    
    
    ' Loop record set and copy file to staging area
    rs.MoveFirst
    While Not rs.EOF
        ' See if source file exists
        If Not fso.FileExists(rs!ImagePath) Then
            'lstFiles.AddItem "Image Not Ready;" & strFile
            GoTo NextImage
        End If
        
        ' See if destination file exists
        strFile = rs!CnlyClaimNum & "MR_" & fso.GetFileName(rs!ImagePath)
        If fso.FileExists(mstrRemotePath & strFile) Then
            ' lstFiles.AddItem "Image Already Moved;" & strFile
        Else
            Call fso.CopyFile(rs!ImagePath, mstrRemotePath & strFile, False)
            '  lstFiles.AddItem "Copied;" & left(strFile, 15)
        End If
NextImage:         rs.MoveNext
    Wend


' COMPLETE step 2


''=========================================================================================
''3. UPDATE VIANT_EXPORT_MR_Log, set StatusFlag = 1
''=========================================================================================
''INCOMPLETE

    intCount = 0
    Set rsViantMR = MyAdo.OpenRecordSet(" SELECT * FROM CMS_AUDITORS_Claims..VIANT_EXPORT_MR_Log where StatusFlag = 0 ")
    If Not (rsViantMR.BOF = True And rsViantMR.EOF = True) Then
        rsViantMR.MoveFirst
        While Not rsViantMR.EOF
            ' update file path
            strFile = rsViantMR!CnlyClaimNum & "MR_" & fso.GetFileName(rsViantMR!ImagePath)
            ' Update table if image exist
            If fso.FileExists(mstrRemotePath & strFile) = True Then
                rsViantMR!StatusFlag = 1
                rsViantMR!Notes = "Image loaded on " & Date
                intCount = intCount + 1
                'Debug.Print mstrRemotePath & strFile
            End If
        rsViantMR.Update
        rsViantMR.MoveNext
        Wend
    End If
    MyAdo.BatchUpdate rsViantMR

    ' Done
    MsgBox intCount & " images transferred.", vbOKOnly
    

''=========================================================================================
'' END
''=========================================================================================


Exit_Sub:
    Set MyAdo = Nothing
    Set myCode_ADO = Nothing
    Set fso = Nothing
    Set oFolder = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    Set rsViantMR = Nothing
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
