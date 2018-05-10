Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrFilePath As String
Private strRowSource As String
Private strAppID As String
Private mstrFieldReference As String
Private mstrFieldValue As String
Private mstrAttachmentType As String
Private WithEvents frmAttachmentSelection As Form_frm_AUDITCLM_References_Attachment_Selection
Attribute frmAttachmentSelection.VB_VarHelpID = -1

Const CstrFrmAppID As String = "ProvRef"


Public Property Get FieldReference() As String
    FieldReference = mstrFieldReference
End Property
Property Let FieldReference(data As String)
     mstrFieldReference = data
End Property
Public Property Get FieldValue() As String
    FieldValue = mstrFieldValue
End Property
Property Let FieldValue(data As String)
     mstrFieldValue = data
End Property
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CnlyRowSource(data As String)
     strRowSource = data
End Property
Property Get CnlyRowSource() As String
     CnlyRowSource = strRowSource
End Property
Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
    'Refresh the grid based on the rowsource passed into the form
    Me.RecordSource = CnlyRowSource
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub


Private Sub cmdAttach_Click()
    Dim oFileDiag As Variant
    Dim strLocalFile As String
    
    Set oFileDiag = Application.FileDialog(msoFileDialogOpen)
    oFileDiag.show
    If oFileDiag.SelectedItems.Count > 0 Then
        ' file selected
        strLocalFile = oFileDiag.SelectedItems(1)
    
        mstrAttachmentType = ""
        Set frmAttachmentSelection = New Form_frm_AUDITCLM_References_Attachment_Selection
        ColObjectInstances.Add frmAttachmentSelection.hwnd & ""
        ShowFormAndWait frmAttachmentSelection
        Set frmAttachmentSelection = Nothing
    
        If mstrAttachmentType <> "" Then
            CopyDocument strLocalFile
        End If
        Set oFileDiag = Nothing
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strOwner As String
    Dim iAnswer As Integer
    Dim strSQL As String
    Dim strFileName As String
    Dim strErrMsg As String
    Dim iReturnCd As Integer
    
    Dim MyAdo As clsADO
    Dim fso As FileSystemObject
    
    
    ' make sure that only attached files are allowed to be deleted
    If Me.RefType <> "PROVATTACH" Then
        MsgBox "I'm Sorry.  You can only delete attached files." & strOwner, vbCritical
        Exit Sub
    End If
    
    ' make sure that only the person who attached the file can delete it
    strOwner = Me.LastUpdateUser
    
    If Identity.UserName <> strOwner Then
        MsgBox "I'm Sorry.  You can only delete files that you attached personally.  This file is attached by " & strOwner, vbCritical
        Exit Sub
    End If
    
    strFileName = Me.RecordSet("RefLink")
    iAnswer = MsgBox("Are you sure you want to delete this file?" & strFileName, vbYesNo)
    
    On Error GoTo Err_handler
    
    If iAnswer = vbYes Then
        ' delete XREF entry
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(strFileName) Then
            ' remove reference entry
            Set MyAdo = New clsADO
            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
            strSQL = "delete from PROV_References where CnlyProvID = '" & Me.cnlyProvID & "' and CreateDt = '" & Me.CreateDt & "' and RefLink = '" & strFileName & "'"
            'CurrentDb.Execute (strSQL)
            MyAdo.SQLTextType = sqltext
            MyAdo.sqlString = strSQL
            iReturnCd = MyAdo.Execute
            If iReturnCd = 1 Then
                'delete file
                Call fso.DeleteFile(strFileName, True)
            Else
                Err.Raise 65000, "Can not remove reference"
            End If
        Else
            strErrMsg = "Error: file " & strFileName & " does not exists!   Please notify IT"
            Err.Raise 65000, , strErrMsg
        End If
    End If
    
    Me.Requery
    
Exit_Sub:
    Set fso = Nothing
    Exit Sub
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    GoTo Exit_Sub
End Sub

Private Sub cmdView_Click()
    Dim strFileName As String
    strFileName = Me.RecordSet("RefLink")
    SetFileReadOnly (strFileName)
    If UCase(Right(strFileName, 3)) = "TIF" Then
    
        Shell "explorer.exe " & strFileName, vbNormalFocus
        ' TK: Removed to work with new TS-CMS server
'        If UCase(left(GetPCName(), 9)) = "TS-FLD-03" Then
'            Shell "explorer.exe " & strFileName, vbNormalFocus
'        Else
'            Shell "c:\program files\Common Files\Microsoft Shared\MODI\11.0\mspview.exe " & strFileName, vbNormalFocus
'        End If
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If
End Sub

Private Sub Form_Load()
    
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    mstrFilePath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = '" & Me.frmAppID & "'"), "")
    
    Me.RecordSource = "SELECT * FROM v_PROV_References WHERE 1=2"
    
    If mstrFilePath = "" Then
        Me.cmdAttach.Enabled = False
    End If
    
End Sub
Private Function CopyDocument(strSource As String) As Boolean
    Dim strErrMsg As String
    Dim strdestinationpath As String
    Dim strFileName As String
    Dim strDestinationFile As String
    
    On Error GoTo Err_handler
    
    Dim fso As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    strdestinationpath = mstrFilePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
    strFileName = Right(strSource, InStr(1, StrReverse(strSource), "\") - 1)

    
    If strFileName = "" Then
       Err.Raise 65000, , "File Not Selected"
    End If
    
    strDestinationFile = strdestinationpath & strFileName
    If fso.FolderExists(strdestinationpath) = False Then
        Call CreateFolder(strdestinationpath)
    End If
    If Not fso.FileExists(strDestinationFile) Then
       Call fso.CopyFile(strSource, strDestinationFile, False)
       Call SetFileReadOnly(strDestinationFile)
    Else
       Err.Raise 65000, , "File Already Exists"
    End If
    
    If Not LogDocument(strDestinationFile) Then
        Call fso.DeleteFile(strDestinationFile, True)
        Err.Raise 65000, , "Error Logging Image"
    End If
    
    
    Me.RefreshData
    
Exit_Sub:
    Set fso = Nothing
    Exit Function
    
Err_handler:
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Function

Private Function LogDocument(strFileName As String) As Boolean
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim iResult As Integer
    

    On Error GoTo ErrHandler

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_PROV_References_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_PROV_References_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pCnlyProvID") = Me.FieldValue
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pRefType") = "PROVATTACH"
    cmd.Parameters("@pRefSubType") = mstrAttachmentType
    cmd.Parameters("@pRefLink") = strFileName
    
    iResult = myCode_ADO.Execute(cmd.Parameters)

    'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        LogDocument = False
        'Err.Raise 65000, "SaveData", "Error updating Hours - " & strErrMsg
        MsgBox "SaveData", "Error Logging Image - " & strErrMsg
    Else
        LogDocument = True
    End If
    
Exit_Function:
    Set cmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    LogDocument = False
    Resume Exit_Function
End Function

Private Sub frmAttachmentSelection_AttachmentSelected(strAttachmentType As String)
    mstrAttachmentType = strAttachmentType
End Sub
