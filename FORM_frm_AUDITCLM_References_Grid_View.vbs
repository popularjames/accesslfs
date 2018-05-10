Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private strRowSource As String
Private strAppID As String
Private mstrFieldReference As String
Private mstrFieldValue As String
Private mstrAttachmentType As String
Private mbReferenceUpdate As Boolean

Private WithEvents frmAuditClmReferenceUpdate As Form_frm_AUDITCLM_References_Update
Attribute frmAuditClmReferenceUpdate.VB_VarHelpID = -1
Private frmAuditClmRelatedImageAssign As Form_frm_AuditClm_RelatedImage_Assign
Attribute frmAuditClmRelatedImageAssign.VB_VarHelpID = -1
Private WithEvents frmAuditClmReferenceRelatedQA As Form_frm_AUDITCLM_References_RelatedQA
Attribute frmAuditClmReferenceRelatedQA.VB_VarHelpID = -1
Private WithEvents frmAttachmentSelection As Form_frm_AUDITCLM_References_Attachment_Selection
Attribute frmAttachmentSelection.VB_VarHelpID = -1
Private frmAttachmentComment As Form_frm_AUDITCLM_References_Comment

Const CstrFrmAppID As String = "AuditClmRef"

'Needs to be a table when I am done
Dim strImagePath As String
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



Private Sub cmdAssignRelated_Click()
    'JS 2014/09/16
    'this will allow for an operator to link this image to other claims as a related image
    
    'first make sure the button has been clicked on an image
    If Me.RefType <> "IMAGE" Then
        MsgBox "This option only works on Images!", vbInformation, "ERROR: RefType not IMAGE"
        Exit Sub
    End If
    
    'check that we are not trying to propagate an image that was rejected as a related claim
    
    'xxxxxxx
    If Me.RelatedClaimMatch = 1 Or Me.RelatedClaimMatch = 3 Then
        MsgBox "This related image is in Pending or Rejected state. Cannot Propagate!", vbInformation, "ERROR: Related Image not Confirmed"
        Exit Sub
    End If
    

    Set frmAuditClmRelatedImageAssign = New Form_frm_AuditClm_RelatedImage_Assign
        
    frmAuditClmRelatedImageAssign.CnlyClaimNum = Me.RecordSet("CnlyClaimNum")
    frmAuditClmRelatedImageAssign.ImageCreateDt = Me.RecordSet("CreateDtText")
    frmAuditClmRelatedImageAssign.RefreshData
    ColObjectInstances.Add frmAuditClmRelatedImageAssign, frmAuditClmRelatedImageAssign.hwnd & ""
    frmAuditClmRelatedImageAssign.visible = True
    
End Sub


Private Function getDocID()
    getDocID = left(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 5) & Right(Replace(Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36), "-", ""), 3)
End Function


Private Sub btn_Comment_Click()
'Added by Gautam(3/21/2011): Leave short comments for supporting appeals documentation sent to payers.
        Dim strLocalFile As String
        Set frmAttachmentComment = New Form_frm_AUDITCLM_References_Comment
        ColObjectInstances.Add frmAttachmentComment.hwnd & ""
        ShowFormAndWait frmAttachmentComment
        strLocalFile = frmAttachmentComment.FileName
        
        If frmAttachmentComment.SaveComment Then
            mstrAttachmentType = "ClmSupport"
            CopyDocument strLocalFile
        End If

        Set frmAttachmentComment = Nothing
End Sub



Private Sub cmbRelated_GotFocus()
    If Me.RelatedClaimMatch = 0 Then
        MsgBox "This image is not waiting for MR QA. Try another one."
        Me.cmdView.SetFocus
    End If
End Sub

Private Sub cmdAnnotate_Click()

Dim MyCodeAdo As clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String

Dim strdestinationpath As String
Dim strFileName As String
Dim strDestination As String



Set MyCodeAdo = New clsADO

If Me.RelatedClaimMatch = 1 Then
    MsgBox "Please do the related claim image QA first then try again!", vbExclamation, "Related claim image QA is Pending"
    Exit Sub
End If

If Me.RefSubType <> "OCR" Then
    ErrMsg = "Document copy not created.  Cannot annotate a non-searchable document.  Document Type must be OCR."
    MsgBox ErrMsg, vbCritical, "Error Creating Annotated Image"
    Exit Sub
End If

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
Set cmd = New ADODB.Command
cmd.ActiveConnection = MyCodeAdo.CurrentConnection
cmd.commandType = adCmdStoredProc
cmd.CommandText = "usp_QUEUE_Exception_apply"
cmd.Parameters.Refresh
cmd.Parameters("@pCnlyClaimNum") = Me.RecordSet("CnlyClaimNum")
cmd.Parameters("@pExceptionType") = "EX051"
cmd.Parameters("@pExceptionStatus") = "OPEN"
cmd.Parameters("@pCreateDt") = Now
cmd.Parameters("@pLastUpdate") = Now()
cmd.Parameters("@pUpdateUser") = Identity.UserName
cmd.Execute
spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@pErrMsg")
    MsgBox ErrMsg, vbCritical, "Error Creating Annotated Image"
Else
    
    Dim strLocalFile As String
    
        strLocalFile = Me.RefLink
        mstrAttachmentType = "ANNOT"
        If mstrAttachmentType <> "" Then
            CopyDocument strLocalFile
            
            strdestinationpath = strImagePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
            strFileName = Right(strLocalFile, InStr(1, StrReverse(strLocalFile), "\") - 1)
            strDestination = strdestinationpath & strFileName
            SetAttr strDestination, vbNormal
        End If
        
    End If
Set MyCodeAdo = Nothing
Set cmd = Nothing


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
    If Me.RefType <> "ATTACH" Then
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
            strSQL = "delete from AUDITCLM_References where CreateDt = '" & Me.CreateDt & "' and RefLink = '" & strFileName & "'"
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

Private Sub cmdFaxStatCS_Click()

If CurrentProject.AllForms("frm_Fax_Selection").IsLoaded = True Then
    DoCmd.Close acForm, "frm_Fax_Selection"
End If

DoCmd.OpenForm "frm_Fax_Selection", acNormal, , , , , "CUSTSERV"

End Sub

Private Sub cmdOpenFax_Click()

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim sIcn As String
Dim sClaimNum As String


sIcn = Me.Parent.Icn
sClaimNum = Me.Parent.CnlyClaimNum

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_CUST_SERV_Load_FaxTbl"
                cmd.Parameters.Refresh
                cmd.Parameters("@sClaimNum") = sClaimNum
                cmd.Parameters("@sICN") = sIcn
                cmd.Parameters("@sProvNum") = Me.RecordSet("ProvNum")
                cmd.Parameters("@sClient") = "2"
             '   cmd.Parameters("@ErrMsg") = ""
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")
                
If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, "Customer Service"
Else
    DoCmd.OpenForm "Frm_CUST_SERV_Review_Results_Worktable", , , , , , sClaimNum
End If

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Sub

Private Sub cmdRelated_Click()

    If Me.Parent.Dirty Or Me.Parent.RecordChanged Then
        MsgBox "Please save other pending changes to this claim first and then try again!", vbExclamation, "Other Claim changes need to be saved first"
        Exit Sub
    End If

   If Me.RelatedClaimMatch <> 0 Then
        If frmAuditClmReferenceRelatedQA Is Nothing Then
            Set frmAuditClmReferenceRelatedQA = New Form_frm_AUDITCLM_References_RelatedQA
            ColObjectInstances.Add frmAuditClmReferenceRelatedQA.hwnd & ""
            frmAuditClmReferenceRelatedQA.CurrRelatedQA = Me.RelatedClaimMatch
            frmAuditClmReferenceRelatedQA.RefreshScreen
            ShowFormAndWait frmAuditClmReferenceRelatedQA
            Set frmAuditClmReferenceRelatedQA = Nothing
        Else
            frmAuditClmReferenceRelatedQA.SetFocus
        End If
    Else
        MsgBox "This document does not need to be QA for Related Claim", vbInformation
    End If
    
End Sub

Private Sub cmdUpdate_Click()


    If Me.Parent.Dirty Or Me.Parent.mbRecordChanged Then
        MsgBox "Please save this claim first (there are pending changes) and then try again!", vbExclamation, "Claim changes need to be saved first"
        Exit Sub
    End If

    If Me.RelatedClaimMatch = 1 Then
        MsgBox "Please do the related claim image QA first then try again!", vbExclamation, "Related claim image QA is Pending"
        Exit Sub
    End If

    If UCase(Me.ModType) = "WRONGIMAGE" Then
        MsgBox "Record has been flagged as 'Wrong Image'.  Can not update this record at this time.  Please alert IT if you wish to do so."
        Exit Sub
    End If
    
    If UCase(Me.ModType) = "DUP" Then
        MsgBox "Record has been flagged as 'Dup'.  Can not update this record at this time.  Please alert IT if you wish to do so."
        Exit Sub
    End If
    
    If UCase(Me.RefType) = "IMAGE" Or UCase(Me.RefType) = "esMDSource" Then
        If frmAuditClmReferenceUpdate Is Nothing Then
            Set frmAuditClmReferenceUpdate = New Form_frm_AUDITCLM_References_Update
            ColObjectInstances.Add frmAuditClmReferenceUpdate.hwnd & ""
            frmAuditClmReferenceUpdate.CurrImageType = Me.RefSubType
            frmAuditClmReferenceUpdate.CurrPageCount = Nz(Me.PageCnt, -1)
            frmAuditClmReferenceUpdate.RefreshScreen
            ShowFormAndWait frmAuditClmReferenceUpdate
            Set frmAuditClmReferenceUpdate = Nothing
        Else
            frmAuditClmReferenceUpdate.SetFocus
        End If
    Else
        MsgBox "Update function is only available for images only", vbInformation
    End If

End Sub


Private Sub cmdView_Click()
Dim strFileName As String
Dim dtCreateDt As Date
Dim sMsg As String
Dim sMsgType As String

    'JS 2013/12/31: Added handling for LostImage ModType
    If Me.ModType = "LostImage" Then
        MsgBox "This image was marked as lost, it cannot be viewed.", vbInformation, "Lost Image"
        GoTo Block_Exit
    End If


    '' 20120920 : KD If the file hasn't been copied to it's directory yet then
    '' tell user to wait a little bit
    '' But if the createdt is kind of old and the file doesn't exist - notify the user
    '' to tell support
    
  
    strFileName = Me.RecordSet("RefLink")
    
    If FileExists(strFileName) = False Then
        dtCreateDt = Me.RecordSet("CreateDt").Value
        If DateDiff("h", dtCreateDt, Now()) < 16 Then
            sMsg = "It looks like there is a delay in the file being copied into the location stored in the database. Please give it some time to work it's way through the queue and try again later."
            sMsgType = "WARNING"
        Else
            ' notify user that the document isn't there - tell support that the link is wrong or something
            sMsg = "There is a disconnect between the database and the actual filename.  Please contact support and let them know the ICN and the document type you are trying to open"
            sMsgType = "ERROR"
        End If
        LogMessage TypeName(Me) & ".cmdView_Click", sMsgType, sMsg, strFileName, True
        GoTo Block_Exit
    End If
    
    If Not Me.RecordSet("RefSubType") = "ANNOT" Then
        SetFileReadOnly (strFileName)
    End If
    If UCase(Right(strFileName, 3)) = "TIF" Then
        If UCase(left(GetPCName(), 9)) = "TS-FLD-03" Then
            Shell "explorer.exe " & strFileName, vbNormalFocus
        Else
            '**DPR 2012-12-12 Image Launch
            'Shell "C:\Program Files (x86)\Common Files\microsoft shared\MODI\11.0\mspview.exe " & strFileName, vbNormalFocus
            Shell "C:\Program Files (x86)\IrfanView\i_view32.exe " & strFileName, vbNormalFocus
            
            
        End If
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, TypeName(Me) & ".cmdView_Click"
    GoTo Block_Exit
End Sub




Private Sub Detail_Paint()
    Debug.Print Me.LastUpdateUser
End Sub

Private Sub Form_Load()


    
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    strImagePath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = '" & Me.frmAppID & "'"), "")
    
    If strImagePath = "" Then
        Me.cmdAttach.Enabled = False
    End If
    
    mbReferenceUpdate = False
    
    Me.RecordSource = "select * from v_AUDITCLM_References where 1=2"
    
    
    
'On Error Resume Next
'    CurrentDb.Execute "CREATE UNIQUE INDEX PK_v_Auditclm_References On v_Auditclm_References (CnlyClaimNum, CreateDt)"
    
    
End Sub
Private Function CopyDocument(strSource As String) As Boolean

    Dim strErrMsg As String
    Dim strdestinationpath As String
    Dim strFileName As String
    Dim strDestination As String
    
    On Error GoTo Err_handler
    
    Dim fso As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    strdestinationpath = strImagePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
    strFileName = Right(strSource, InStr(1, StrReverse(strSource), "\") - 1)

    
    If strFileName = "" Then
       Err.Raise 65000, , "File Not Selected"
    End If
    
    strDestination = strdestinationpath & strFileName
    If fso.FolderExists(strdestinationpath) = False Then
        Call CreateFolder(strdestinationpath)
    End If
    If Not fso.FileExists(strDestination) Then
       Call fso.CopyFile(strSource, strDestination, False)
       Call SetFileReadOnly(strDestination)
    Else
       Err.Raise 65000, , "File Already Exists"
    End If
    
    If Not LogDocument(strDestination) Then
        Call fso.DeleteFile(strDestination, True)
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
    myCode_ADO.sqlString = "usp_AUDITCLM_References_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_AUDITCLM_References_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pCnlyClaimNum") = Me.FieldValue
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pRefType") = "ATTACH"
    cmd.Parameters("@pRefSubType") = mstrAttachmentType
    cmd.Parameters("@pRefLink") = strFileName
    
    iResult = myCode_ADO.Execute(cmd.Parameters)

    'Make sure there are no errors
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    If strErrMsg <> "" Then
        LogDocument = False
        'Err.Raise 65000, "SaveData", "Error updating Hours - " & strErrMsg
        MsgBox "SaveData", "Error Logging Image - " & strErrMsg
    Else
        LogDocument = True
    End If
    
Exit_Function:
    Set cmd = Nothing
    Set myCode_ADO = Nothing
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

Private Sub frmAuditClmReferenceUpdate_UpdateReferences(ErrorCode As String, NewImageType As String, NewPageCount As Integer, Comment As String)
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim strUserMsg As String
    Dim iResult As Integer
    

    On Error GoTo ErrHandler

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_SCANNING_Image_Error_Log_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_SCANNING_Image_Error_Log_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pScannedDt") = ConvertTimeToString(Me.CreateDt)
    cmd.Parameters("@pCnlyClaimNum") = Me.RecordSet("CnlyClaimNum")
    cmd.Parameters("@pNewCnlyClaimNum") = ""
    cmd.Parameters("@pNewImageType") = UCase(NewImageType)
    cmd.Parameters("@pNewPageCnt") = NewPageCount
    cmd.Parameters("@pErrorCd") = ErrorCode
    cmd.Parameters("@pComment") = Comment
    
    cmd.Execute
    
    'Make sure there are no errors
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    If strErrMsg <> "" Then
        strErrMsg = "Error updating Image - " & strErrMsg
        GoTo ErrHandler
    End If
    
    'check to see if we need to display a user message
    strUserMsg = cmd.Parameters("@pUserMsg") & ""
    If strUserMsg <> "" Then MsgBox strUserMsg, vbInformation
    
    ' refresh grid box view
    Me.Requery
    'Me.Parent.RefreshMain
    Me.Parent.LoadData
    
Exit_Function:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub

ErrHandler:
    If strErrMsg <> "" Then
        MsgBox strErrMsg
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Function
End Sub


Private Sub frmAuditClmReferenceRelatedQA_UpdateReferences(RelatedQAMatch As Integer)
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim strUserMsg As String
    Dim iResult As Integer
    

    On Error GoTo ErrHandler

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_AUDITCLM_RelatedQAMatch"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_AUDITCLM_RelatedQAMatch"
    cmd.Parameters.Refresh
    
    
    cmd.Parameters("@pCnlyClaimNum") = Me.RecordSet("CnlyClaimNum")
    cmd.Parameters("@pRefLink") = Me.RecordSet("RefLink")
    cmd.Parameters("@pRelatedQAMatch") = RelatedQAMatch

    cmd.Execute
    
    'Make sure there are no errors
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    If strErrMsg <> "" Then
        strErrMsg = "Error updating Image - " & strErrMsg
        GoTo ErrHandler
    End If
    
    'check to see if we need to display a user message
    strUserMsg = cmd.Parameters("@pUserMsg") & ""
    If strUserMsg <> "" Then MsgBox strUserMsg, vbInformation
    
    ' refresh grid box view
    Me.Requery
    
Exit_Function:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub

ErrHandler:
    If strErrMsg <> "" Then
        MsgBox strErrMsg
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Function
End Sub
