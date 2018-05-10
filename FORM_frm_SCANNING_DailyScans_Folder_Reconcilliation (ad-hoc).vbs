Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mbContinue As Boolean
Dim miFileCnt As Long
Dim miFileDelete As Long
Dim miFileCopy As Long

Dim mstrTargetDir As String
Dim mstrSourceDir As String


Dim mstrCalledFrom As String



Private Sub cmdBrowseSouce_Click()
    Dim oFileDiag As FileDialog
    
    Set oFileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    If Right(Me.txtSourceDir, 1) <> "\" Then
        oFileDiag.InitialFileName = Me.txtSourceDir & "\"
    Else
        oFileDiag.InitialFileName = Me.txtSourceDir
    End If
    oFileDiag.show
    If oFileDiag.SelectedItems.Count > 0 Then
        txtSourceDir = oFileDiag.SelectedItems(1)
    End If
    Set oFileDiag = Nothing
End Sub

Private Sub cmdBrowseTarget_Click()
    Dim oFileDiag As FileDialog
    
    Set oFileDiag = Application.FileDialog(msoFileDialogFolderPicker)
    If Right(Me.txtTargetDir, 1) <> "\" Then
        oFileDiag.InitialFileName = Me.txtTargetDir & "\"
    Else
        oFileDiag.InitialFileName = Me.txtTargetDir
    End If
    oFileDiag.show
    If oFileDiag.SelectedItems.Count > 0 Then
        txtTargetDir = oFileDiag.SelectedItems(1)
    End If
    Set oFileDiag = Nothing
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub CmDRun_Click()
    
    Dim bFileExists As Boolean
    
    Dim fso As FileSystemObject
    Dim oFolder As Folder
    
    Dim bResult As Boolean
    
    
    On Error GoTo Err_handler
    
    ' reset display info
    lstFiles.RowSource = ""
    lstFiles.Requery
    lblFileScanned.Caption = ""
    lblFileCopied.Caption = ""
    lblFileDeleted.Caption = ""
    miFileCnt = 0
    miFileDelete = 0
    miFileCopy = 0
    
    
    
    
    ' check image paths
    Set fso = New FileSystemObject
    mstrSourceDir = txtSourceDir.Value
    mstrTargetDir = txtTargetDir.Value
    
    If Right(mstrSourceDir, 1) <> "\" Then mstrSourceDir = mstrSourceDir & "\"
    If Right(mstrTargetDir, 1) <> "\" Then mstrTargetDir = mstrTargetDir & "\"
    
    If fso.FolderExists(mstrSourceDir) = False Then
        MsgBox "Source directory does not exists!"
        Exit Sub
    End If
    
    If fso.FolderExists(mstrTargetDir) = False Then
        MsgBox "Target directory does not exists!"
        Exit Sub
    End If
    
    
    
    'scanned files
    Set oFolder = fso.GetFolder(mstrSourceDir)
    mbContinue = True
    Call ScanFile(oFolder)
    
    
    If mbContinue = True Then
        MsgBox "Scanning completed"
    Else
        MsgBox "Scanned stopped"
    End If
    
    Call RemoveEmptyFolders(mstrSourceDir, False)
    
    
Exit_Sub:
    Set fso = Nothing
    Set oFolder = Nothing
    Exit Sub
    
Err_handler:
    MsgBox Err.Description
    
    Resume Exit_Sub
End Sub

Private Sub cmsStop_Click()
    mbContinue = False
End Sub

Private Sub ScanFile(oFolder As Folder)
    
    Dim fso As New FileSystemObject
    
    Dim oSubFolder As Folder
    Dim oFile As file
    
    Dim strSourceFile As String
    Dim strTargetFile As String
    Dim strMsg As String
    
    On Error GoTo Err_handler
    
    For Each oFile In oFolder.Files
        strTargetFile = Replace(oFile.Path, mstrSourceDir, mstrTargetDir)

        strSourceFile = oFile.Path

'    Debug.Print strTargetFile
'    Debug.Print strSourceFile

        If fso.FileExists(strTargetFile) Then
            If Me.optFileCompare.Value = 1 Then
                oFile.Delete
                miFileDelete = miFileDelete + 1
                strMsg = "File deleted"
            ElseIf Me.optFileCompare.Value = 2 Then
                oFile.Copy strTargetFile, True
                If fso.FileExists(strTargetFile) Then
                    oFile.Delete
                    strMsg = "File copied"
                    miFileCopy = miFileCopy + 1
                    miFileDelete = miFileDelete + 1
                End If
            End If
        Else
            If oFile.Name Like "########.tif" Or left(oFile.Name, 13) = "SCANNED VALUE" Or InStr(1, ".LNK/.TMP/.RDP", UCase(Right(oFile.Name, 4))) > 0 Then
                oFile.Delete
                miFileDelete = miFileDelete + 1
                strMsg = "File deleted"
            Else
                strMsg = "File not exists"
            End If
        End If
        
        miFileCnt = miFileCnt + 1
        
        lblFileScanned.Caption = "Total scanned: " & miFileCnt
        lblFileDeleted.Caption = "Total deleted: " & miFileDelete
        lblFileCopied.Caption = "Total copied: " & miFileCopy
        
        Me.lstFiles.AddItem (strMsg & ";" & Replace(strSourceFile, ",", "")), 0
        
        
        If Me.lstFiles.ListCount > 100 Then
            Me.lstFiles.RemoveItem (100)
        End If
        DoEvents
        DoEvents
    Next
    
    For Each oSubFolder In oFolder.SubFolders
'Debug.Print oSubFolder.Path
        Call ScanFile(oSubFolder)
        
        If mbContinue = False Then Exit For
        DoEvents
        DoEvents
    Next
    
    Set oFile = Nothing
    Set oSubFolder = Nothing
    Set fso = Nothing
    
   
    Exit Sub

Err_handler:
    MsgBox Err.Number & " -- " & Err.Description & " - File " & strSourceFile
    mbContinue = False
End Sub


Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "FOLDER RECONCILLIATION"
    
    Me.txtSourceDir = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\MEDICALRECORD_TEMP\DailyScans"
    Me.txtTargetDir = "\\ccaintranet.com\DFS-CMS-FLD\Imaging\Client\Out\CMS\MedicalRecords_Current"
    
    lblFileScanned.Caption = ""
    lblFileDeleted.Caption = ""
    lblFileCopied.Caption = ""
    
    lstFiles.RowSource = ""
    lstFiles.Requery
    
    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs()
    End If
End Sub


Private Sub lstFiles_DblClick(Cancel As Integer)
    Dim strFileName As String
    Dim strImagePath As String
    
    strFileName = lstFiles.Column(1)
    strImagePath = Mid(strFileName, 1, Len(strFileName) - InStr(1, StrReverse(strFileName), "\"))
    Shell "explorer.exe " & strImagePath, vbNormalFocus
    Shell "explorer.exe " & strFileName, vbNormalFocus

End Sub
