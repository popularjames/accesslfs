Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub cmdArchive_Click()

'Dim sproc As New clsAdoSproc
    
'    If txtFileId = Nz(Me.txtFileId, "") = "" Then
    
'       Me.txtFileId = DLookup("FileId", "RAC_OUTBOUND_StatusRecord")
       
'    End If
    
'    With sproc
'        .ConnectString = GetConnectString("dbo_v_Cms_Auditors_Code")
'        .CommandText = "usp_RAC_Outbound_StatusRecord_MoveToArchive"
'        .Setup
'
'        .AddParam "@FileId", Me.txtFileId
'        .Exec
'        Debug.Print .ReturnValue
'
'    End With'
'
'ExitHere:
'
'    Exit Sub
'
'HandleError:
'
'    GoTo ExitHere
'
End Sub


Private Sub cmdExecute_Click()

'  Dim sproc As New clsAdoSproc
'  Dim cbo As New ComboBox
  
'  Set cbo = Me.cboProcedure
  
'
'    With sproc
'        .ConnectString = GetConnectString(cbo(3))
'        .CommandText = cbo(2)
'        .Setup
'
'        .Exec
'
'    End With

'ExitHere:
'    On Error Resume Next
    
'    Set sproc = Nothing
    
'     Exit Sub
'HandleError:

'    MsgBox "Error: " & CStr(Err.Number) & vbCr & Err.Description
    
'    Debug.Print Err.Number
'    Debug.Print Err.Description
    
'    GoTo ExitHere

End Sub

Private Sub cmdExportMap_Click()
  'On Error GoTo HandleError

    Dim dlg As New clsDialogs
    Dim strOutputFolder As String
    Dim strOutPutFile As String
    Dim fso As Variant
    
    Dim djEngine
    Dim djConversion
    Dim djLog
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set djEngine = CreateObject("DJEC.Engine")
    Set djConversion = CreateObject("DJEC.Conversion")
    Set djLog = CreateObject("DJEC.LogManager")
    
    Me.lblLoading.visible = True
    Me.lblErrors.Caption = ""
    Me.lblRows.Caption = ""
    
    If Nz(Me.cboCosmosMap.Value, "") = "" Then
         MsgBox "Please select a map", vbOKOnly
        GoTo exitHere
    End If
    
    ' Validate folder
    strOutputFolder = Nz(Me.txtTargetFolderName, "")
       
    If strOutputFolder = "" Or Not fso.FolderExists(strOutputFolder) Then
        MsgBox "Invalid folder", vbOKOnly
        GoTo exitHere
    End If
    
            ' Replace any "/" with "\"
    strOutputFolder = Replace(strOutputFolder, "/", "\")

    If Right(strOutputFolder, 1) <> "\" Then
        strOutputFolder = strOutputFolder & "\"
    End If

    

    'test folder "Y:\Raw\FS01-K\CMS\PROVIDERFILES\QA_CHECK\Cosmos Test\"
    
    'SFileName = dlg.OpenPath("Y:\Raw\FS01-K\CMS\PROVIDERFILES\QA_CHECK\Cosmos Test\", txtf)
    


    'Me.lblMap.Caption = SFileName
    
    djEngine.InitializationFile = "C:\Program Files\Pervasive\Cosmos\Common800\dj800.ini"
    
    ' Data from MapFilePath
    djConversion.Load Me.cboCosmosMap.Column(2)
    
    '("Y:\Data\FS01-M\CMS\Cosmos\Providers\CosmosTest.tf.xml")
    
    '"Y:\Data\FS01-M\CMS\Cosmos\Providers\CosmosTest.tf.xml"
    

    strOutPutFile = strOutputFolder & Year(Date) & "-" & Month(Date) & _
                        "-" & Day(Date) & "-Test.txt"
    
   'Set the connection information for the target.
   djConversion.Targets(0).connectioninfo.file = strOutPutFile
   
    
' Run map
    Set djLog = djConversion.MessageLog
    djLog.FileName = strOutPutFile & ".err"
    
    djConversion.Run
    
    Me.lblErrors.Caption = CStr(djConversion.ErrorCount)
    Me.lblRows.Caption = CStr(djConversion.WriteCount)
    
exitHere:
On Error Resume Next

Me.lblLoading.visible = False

    Set djEngine = Nothing
    Set djConversion = Nothing
    Set djLog = Nothing
    
    Exit Sub
HandleError:

    'MsgBox "Error: " & CStr(Err.Number) & vbCr & Err.Description & vbCr & "DJ: " & djLog.LastErrorMessage
    

    Debug.Print djLog.LastErrorMessage
    Debug.Print Err.Number
    Debug.Print Err.Description
    

    GoTo exitHere


End Sub



Private Sub Form_Load()
    Me.lblLoading.visible = False
    Me.lblErrors.Caption = ""
    Me.lblRows.Caption = ""
    
End Sub
