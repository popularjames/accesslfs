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
Const CstrFrmAppID As String = "ConceptRef"
Private mstrAttachmentType As String

'' 20120416 KD Added
Private cstrCnlyAttachmentType As String
Private cintTaggedClaimId As Integer
Private cstClientIssueId As String
Private cstrFileNewFileName As String
Private csConceptId As String
Private csPayerNames() As String
Private csPayerNameIds() As String
Private cstrPayerNames As String
Private cstrPayerNameIds As String

Private coEracReqDocType As clsConceptReqDocType
    'Private WithEvents frmAttachmentSelection As Form_frm_AUDITCLM_References_Attachment_Selection
Private WithEvents frmAttachmentSelection As Form_frm_CONCEPT_References_Attachment_Selection
Attribute frmAttachmentSelection.VB_VarHelpID = -1
'' 20120416 KD End Added



'Needs to be a table when I am done
Dim strImagePath As String




'' 20120416 KD added
Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sConceptId As String)
    csConceptId = sConceptId
    Me.txtSelectedId = sConceptId
End Property

Public Property Get ClientIssueId() As String
    If cstClientIssueId = "" Then
        Call mod_Concept_Specific.GetConceptHeaderDetails(Me.IdValue, , , , cstClientIssueId)
    End If
    ClientIssueId = cstClientIssueId
End Property
Property Let ClientIssueId(sClientIssueId As String)
     cstClientIssueId = sClientIssueId
End Property

'' This property will hold the new filename as defined by the naming convention
'' of the Doc Type that it is..  If null, don't rename..
Public Property Get FileNewFileName() As String
    FileNewFileName = cstrFileNewFileName
End Property
Property Let FileNewFileName(strFileNewFileName As String)
     cstrFileNewFileName = strFileNewFileName
End Property
'' 20120416 KD End added

Public Property Get PayerNames() As String
    PayerNames = cstrPayerNames
End Property
Public Property Let PayerNames(strPayerNames As String)
    cstrPayerNames = strPayerNames
End Property


Public Property Get PayerNameIds() As String
    PayerNameIds = cstrPayerNameIds
End Property
Public Property Let PayerNameIds(strPayerNameIds As String)
    cstrPayerNameIds = strPayerNameIds
End Property

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

'?
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



Public Sub PayerChange()
    cmbPayer_Change
End Sub


Private Sub cmbPayer_Change()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_CONCEPT_Main

    strProcName = ClassName & ".cmbPayer_Change"
    
        '' Need to filter or unfilter tagged claims
    
    If cmbPayer.Value = 1000 Then
        ' No filter:
        Me.filter = ""
        Me.FilterOn = False
    Else
        Me.filter = "PayerNameId = " & CStr(cmbPayer.Value)
        Me.FilterOn = True
    End If
    
    
        '  Globally save the selected payer
    If IsSubForm(Me) = True Then
        Set oFrm = Me.Parent
        oFrm.SelectedPayerNameId = Me.cmbPayer.Value
    End If
    
Block_Exit:
    Set oFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub




Private Sub cmdAttach_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFileDiag As Variant
Dim strLocalFile As String

    strProcName = ClassName & ".cmdAttach_Click"
    

    Set oFileDiag = Application.FileDialog(msoFileDialogOpen)
    oFileDiag.show
    If oFileDiag.SelectedItems.Count > 0 Then
        ' file selected
        strLocalFile = oFileDiag.SelectedItems(1)

        LogMessage strProcName, , "User selected a file", strLocalFile

        '' make sure everything is reset
        mstrAttachmentType = ""
        cstrCnlyAttachmentType = ""
        cintTaggedClaimId = 0
        
        '   Set frmAttachmentSelection = New Form_frm_AUDITCLM_References_Attachment_Selection
        Set frmAttachmentSelection = New Form_frm_CONCEPT_References_Attachment_Selection
        frmAttachmentSelection.ConceptID = Me.IdValue
        frmAttachmentSelection.ClientIssueId = Me.ClientIssueId
        frmAttachmentSelection.FilePathSelected = strLocalFile
        frmAttachmentSelection.visible = True
        
        ColObjectInstances.Add frmAttachmentSelection.hwnd & " "
        ShowFormAndWait frmAttachmentSelection
        Set frmAttachmentSelection = Nothing

        LogMessage strProcName, , "User chose type: " & cstrCnlyAttachmentType & " and " & CStr(cintTaggedClaimId) & " as eRacTaggedClaimId"

        If mstrAttachmentType <> "" Then
            CopyDocument strLocalFile
        End If
        Set oFileDiag = Nothing
        
    Else
        LogMessage strProcName, , "User did not select a file", strLocalFile
    End If


Block_Exit:
    If IsOpen("frm_CONCEPT_References_Attachment_Selection") Then
        DoCmd.Close acForm, "frm_CONCEPT_References_Attachment_Selection", acSaveNo
    End If
    Set frmAttachmentSelection = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdEditURL_Click()
   
   Dim bWorked As Boolean
   bWorked = UpdateURL(Me.ConceptID, Me.RefSequence, Me.RowID)
   Me.Requery


Exit_Sub:
    Exit Sub

ErrHandler:
    Resume Exit_Sub
End Sub

Private Sub Command56_Click()
    Stop
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me.ConceptID.Value = Me.Parent.Form.txtConceptID
End Sub
    
    
Private Sub Form_Load()
    
Dim iAppPermission As Integer
Dim sPayers As String


    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then
        Me.RecordSource = ""
        Exit Sub
    End If

    sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
                

    strImagePath = Nz(DLookup("FolderPath", "GENERAL_Attachment_Path", "Appid = '" & Me.frmAppID & "'"), "")
    
    If strImagePath = "" Then
        Me.cmdAttach.Enabled = False
    End If

    If IsSubForm(Me) = True Then
    
        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If
        If sPayers <> "" Then
            sPayers = "1000," & sPayers
            'sRecordSource = sRecordSource & " AND (PayerNameID IN (" & sPayers & ") OR PayerNameID IS NULL ) "
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (" & sPayers & ") ORDER BY PayerName"
        
        Else
            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        End If
    
    
        Me.RecordSource = "SELECT * FROM v_CONCEPT_References WHERE ConceptID = '" & Me.Parent.Form.txtConceptID & "'" & " ORDER BY ConceptId, RefSequence"
    
    Else
        Me.RecordSource = "SELECT * FROM v_CONCEPT_References ORDER BY ConceptId, RefSequence"
        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000"
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If


    
End Sub

'' 20120416 KD: changed this whole function around. See previous version if needed
'' 20120622 KD: changed this whole function AGAIN due to CMS concept per payer changes
Private Function CopyDocument(strSourceFullPath As String) As Boolean
Dim strErrMsg As String
Dim strdestinationpath As String
Dim strFileName As String
Dim fso As Scripting.FileSystemObject
Dim intRecordCOunt As Integer
Dim strProcName As String
Dim sFileExtension As String
Dim sDestFileName As String
Dim sPayerFiles() As String
Dim iPayerIdx As Integer
Dim sThisPayerName As String
Dim sThisPayerNameID   As String
Dim oPayer As clsConceptPayerDtl

On Error GoTo Err_handler
    
    strProcName = ClassName & ".CopyDocument"
    DoCmd.Hourglass True
    
    Set fso = New Scripting.FileSystemObject
    
    strdestinationpath = strImagePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
    strFileName = Right(strSourceFullPath, InStr(1, StrReverse(strSourceFullPath), "\") - 1)
    
    ' We need to know if we have to convert this puppy, so we get the file extension
    '' for use later
    ' Get the file extension:
    sFileExtension = FileExtension(strFileName)
 
    '' First, we need to see if we need to rename the file based on the naming conventions?
    If Me.FileNewFileName <> "" Then
        ' Yes, rename it..
        ' rename it and add the file extension back (FileNewFileName isn't going to have it)
        
        sPayerFiles = Split(Me.FileNewFileName, ",")
        If UBound(csPayerNameIds) < 0 Then
            ' not a payer specific document...
            If SaveDoc(strSourceFullPath, strFileName, strdestinationpath, sFileExtension, "", "") = False Then
                Stop
            End If
            GoTo NextOne
        End If
'        For iPayerIdx = 0 To UBound(sPayerFiles)
        For iPayerIdx = 0 To UBound(csPayerNameIds)
            Set oPayer = New clsConceptPayerDtl

            If oPayer.LoadFromConceptNPayer(csConceptId, CLng(csPayerNameIds(iPayerIdx))) = False Then
                Stop
            End If
        
            '        20120622: If there is a comma in the filename then we need to make Payer Specific copies of the file
            If UBound(csPayerNames) > -1 Then
                sThisPayerName = csPayerNames(iPayerIdx)
            Else
                sThisPayerName = ""
            End If
            If UBound(csPayerNameIds) > -1 Then
                sThisPayerNameID = csPayerNameIds(iPayerIdx)
            Else
                sThisPayerNameID = ""
            End If
            
            If LCase(Right(Me.FileNewFileName, Len(sFileExtension) + 1)) <> "." & LCase(sFileExtension) Then
'Stop
                strFileName = Me.FileNewFileName & "." & sFileExtension
            Else
                strFileName = Me.FileNewFileName
            End If
            
            '' Later, if we have sDestFileName then we use that. otherwise filename remains unchanged
            sDestFileName = Me.FileNewFileName & "." & sFileExtension
            strdestinationpath = strImagePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
    
            strdestinationpath = strdestinationpath & UCase(sThisPayerName)
''            coEracReqDocType.


            strFileName = coEracReqDocType.ParseFileName(csConceptId, oPayer.ClientIssueId, "", strFileName, oPayer.PayerName, oPayer)
            '' NOTE: this is now withOUT the file extension..
'Stop

            Call CreateFolders(strdestinationpath)
            If SaveDoc(strSourceFullPath, strFileName & "." & sFileExtension, strdestinationpath & "\", sFileExtension, sThisPayerName, sThisPayerNameID) = False Then
                LogMessage strProcName, "ERROR", "There was an error saving the document!" & strSourceFullPath
                
            End If
            
            
'Stop    ' ok, we need to store this in the database.. , each instance I suppose
            ' sPayerFiles(iPayerIdx)
'
'            If Not AddAttachedDocumentToDb(strDestinationPath & strFileName, strDestinationPath, strFileName, intRecordCOunt, sThisPayerName, CInt(sThisPayerNameId)) Then
'                LogMessage strProcName, "WARNING", "Was unable to link the file in the database!"
'                    '    If Not LogDocument(strDestinationPath & strFileName, strDestinationPath, strFileName, intRecordCOunt) Then
'                Call fso.DeleteFile(strDestinationPath & strFileName)
'                Err.Raise 65000, , "Error Logging Image"
'            End If
            
NextOne:
        Next
        
    Else
        sThisPayerName = Me.PayerNames
        sThisPayerNameID = Me.PayerNameIds
        
        If SaveDoc(strSourceFullPath, strFileName, strdestinationpath, sFileExtension, sThisPayerName, sThisPayerNameID) = False Then
            Stop
        End If
    End If
    
    Me.RefreshData
    
Exit_Sub:
    DoCmd.Hourglass False

    Set fso = Nothing
    Exit Function
    
Err_handler:
    DoCmd.Hourglass False
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
    GoTo Exit_Sub
End Function


Private Function SaveDoc(strSourceFullPath As String, strFileName As String, strdestinationpath As String, sFileExtension As String, _
        sThisPayerName As String, sThisPayerNameID As String) As Boolean
Dim strErrMsg As String
Dim fso As Scripting.FileSystemObject
Dim intRecordCOunt As Integer
Dim strProcName As String
Dim sPayerFiles() As String
Dim iPayerIdx As Integer

On Error GoTo Block_Err
    
    strProcName = ClassName & ".SaveDoc"
    
    '' ok so we should have:
    ' strFileName (filename AND extension, just no path)
    '   (this is separated in case we are changing the file's name)
    ' strDestnationPath is where it should be saved to
    ' sFileExtension (what it SHOULD be saved as..)
    
    Set fso = New Scripting.FileSystemObject

    If coEracReqDocType Is Nothing Then
        If mstrAttachmentType <> "" Then
            Set coEracReqDocType = New clsConceptReqDocType
            If coEracReqDocType.LoadFromId(CInt(mstrAttachmentType)) = False Then
                LogMessage strProcName, "ERROR", "Could not load the doc type with id: " & mstrAttachmentType
                GoTo Block_Exit
            End If
        Else
            LogMessage strProcName, "ERROR", "Seems to be a disconnect in the Required Document Type to Attach sub"
            GoTo Block_Exit
        End If
    End If
    
    ' If the file extension doesn't match what we are supposed to have then we need to put it in the convert queue
    If UCase(sFileExtension) <> UCase(coEracReqDocType.SendAsFileType) Then
        LogMessage strProcName, , "File needs to be converted", sFileExtension & " to " & coEracReqDocType.SendAsFileType
        ' Put it in the Conversion Queue
        '' Ok, so the conversion queue changed so it'll just take care of it.. so later, when we add to the database it's going
        '' to be added to the conversion queue there too.. it'll be renamed or converted or moved or whatever needs to happen
    End If
    
    
    If strFileName = "" Then
       Err.Raise 65000, , "File Not Selected"
    End If
    
    If fso.FolderExists(strdestinationpath) = False Then
        Call CreateFolder(strdestinationpath)
    End If
    If Not fso.FileExists(strdestinationpath & strFileName) Then
        LogMessage strProcName, , "Copying file", strSourceFullPath & " to " & strdestinationpath
        Call fso.CopyFile(strSourceFullPath, strdestinationpath & strFileName, False)
    Else
        LogMessage strProcName, "WARNING", "File being copied already resides in destination", strdestinationpath & strFileName
        Err.Raise 65000, , "File Already Exists"
        GoTo Block_Exit
    End If
    
    If Me.RecordSet.EOF = True Then
        intRecordCOunt = 1
    Else
        intRecordCOunt = Me.txtRecordCount + 1
    End If
    
    'edit usp in logdocument
    LogMessage strProcName, , "About to attach this document in CONCEPT_References table"
    If Not AddAttachedDocumentToDb(strdestinationpath & strFileName, strdestinationpath, strFileName, intRecordCOunt, _
            coEracReqDocType.DocTypeId, sThisPayerName, CLng(CStr("0" & sThisPayerNameID)), strSourceFullPath) Then
        LogMessage strProcName, "WARNING", "Was unable to link the file in the database!"
            '    If Not LogDocument(strDestinationPath & strFileName, strDestinationPath, strFileName, intRecordCOunt) Then
        Call fso.DeleteFile(strdestinationpath & strFileName)
        Err.Raise 65000, , "Error Logging Image"
        SaveDoc = False
    End If

    SaveDoc = True

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    SaveDoc = False
    GoTo Block_Exit
End Function

''''' 20120416 KD: changed this whole function around. See previous version if needed
'''Private Function CopyDocument_LEGACY(strSource As String) As Boolean
'''Dim strErrMsg As String
'''Dim strDestinationPath As String
'''Dim strFileName As String
'''Dim fso As Scripting.FileSystemObject
'''Dim intRecordCOunt As Integer
'''Dim strProcName As String
'''Dim sFileExtension As String
'''Dim sDestFileName As String
'''On Error GoTo Err_Handler
'''
'''    strProcName = ClassName & ".CopyDocument_LEGACY"
'''
'''    Set fso = New Scripting.FileSystemObject
'''
'''    strDestinationPath = strImagePath & "\" & mstrFieldReference & "\" & mstrFieldValue & "\"
'''    strFileName = Right(strSource, InStr(1, StrReverse(strSource), "\") - 1)
'''
'''    ' We need to know if we have to convert this puppy, so we get the file extension
'''    '' for use later
'''    ' Get the file extension:
'''    Call PathInfoFromPath(strFileName, , , sFileExtension)
'''
'''    '' First, we need to see if we need to rename the file based on the naming conventions?
'''    If Me.FileNewFileName <> "" Then
'''        ' Yes, rename it..
'''        ' rename it and add the file extension back (FileNewFileName isn't going to have it)
'''
'''        strFileName = Me.FileNewFileName & "." & sFileExtension
'''        '' Later, if we have sDestFileName then we use that. otherwise filename remains unchanged
'''        sDestFileName = Me.FileNewFileName & "." & sFileExtension
'''    End If
'''
'''    If coEracReqDocType Is Nothing Then
'''        If mstrAttachmentType <> "" Then
'''            Set coEracReqDocType = New clsConceptReqDocType
'''            If coEracReqDocType.LoadFromID(CInt(mstrAttachmentType)) = False Then
'''                LogMessage strProcName, "ERROR", "Could not load the doc type with id: " & mstrAttachmentType
'''                GoTo Exit_Sub
'''            End If
'''        Else
'''            LogMessage strProcName, "ERROR", "Seems to be a disconnect in the Required Document Type to Attach sub"
'''            GoTo Exit_Sub
'''        End If
'''    End If
'''
'''    ' If the file extension doesn't match what we are supposed to have then we need to put it in the convert queue
'''    If UCase(sFileExtension) <> UCase(coEracReqDocType.SendAsFileType) Then
'''        LogMessage strProcName, , "File needs to be converted", sFileExtension & " to " & coEracReqDocType.SendAsFileType
'''        ' Put it in the Conversion Queue
'''        '' Ok, so the conversion queue changed so it'll just take care of it.. so later, when we add to the database it's going
'''        '' to be added to the conversion queue there too.. it'll be renamed or converted or moved or whatever needs to happen
'''    End If
'''
'''
'''    If strFileName = "" Then
'''       Err.Raise 65000, , "File Not Selected"
'''    End If
'''
'''    If fso.FolderExists(strDestinationPath) = False Then
'''        Call CreateFolder(strDestinationPath)
'''    End If
'''    If Not fso.FileExists(strDestinationPath & strFileName) Then
'''        'try this to set a new file name
'''        LogMessage strProcName, , "Copying file", strSource & " to " & strDestinationPath
'''        Call fso.CopyFile(strSource, strDestinationPath & sDestFileName, False)
'''    Else
'''        LogMessage strProcName, "WARNING", "File being copied already resides in destination", strDestinationPath & strFileName
'''        Err.Raise 65000, , "File Already Exists"
'''    End If
'''
'''    If Me.Recordset.EOF = True Then
'''        intRecordCOunt = 1
'''    Else
'''        intRecordCOunt = Me.txtRecordCount + 1
'''    End If
'''
'''    'edit usp in logdocument
'''    LogMessage strProcName, , "About to attach this document in CONCEPT_References table"
'''    If Not AddAttachedDocumentToDb(strDestinationPath & strFileName, strDestinationPath, strFileName, intRecordCOunt) Then
'''        LogMessage strProcName, "WARNING", "Was unable to link the file in the database!"
'''            '    If Not LogDocument(strDestinationPath & strFileName, strDestinationPath, strFileName, intRecordCOunt) Then
'''        Call fso.DeleteFile(strDestinationPath & strFileName)
'''        Err.Raise 65000, , "Error Logging Image"
'''    End If
'''
'''    Me.RefreshData
'''    CopyDocument_LEGACY = True
'''Exit_Sub:
'''
'''    Set fso = Nothing
'''    Exit Function
'''
'''Err_Handler:
'''    If strErrMsg = "" Then strErrMsg = Err.Description
'''    MsgBox "Error in module " & vbCrLf & vbCrLf & strErrMsg
'''    CopyDocument_LEGACY = False
'''    GoTo Exit_Sub
'''End Function

Private Function LogDocument(strPathFileName As String, strPath As String, strFileName As String, intSequence) As Boolean
    Dim myCode_ADO As clsADO
    Dim colPrms As ADODB.Parameters
    Dim prm As ADODB.Parameter
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
    On Error GoTo ErrHandler
    Dim cmd As ADODB.Command
    

'ALTER Procedure [dbo].[usp_CONCEPT_References_Insert]
'    @pCnlyClaimNum varchar(30),
'    @pCreateDt datetime,
'    @pRefType varchar(20),
'    @pRefSubType varchar(20),
'    @pRefLink varchar(1000),
'    @pErrMsg varchar(255) output
'as

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_CONCEPT_References_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_CONCEPT_References_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pConceptID") = Me.FieldValue
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pRefType") = "DOC"
    '' 20120416 KD COMEBACK: this is only needed during the portion where we switch over to the new attachment stuff.
    '' actually, I don't think this is going to be used anymore
    If IsNumeric(mstrAttachmentType) Then
        cmd.Parameters("@pRefSubType") = cstrCnlyAttachmentType
    Else
        cmd.Parameters("@pRefSubType") = mstrAttachmentType
    End If
    '' 20120416 KD End
    cmd.Parameters("@pRefLink") = strPathFileName
    'New Fields
    cmd.Parameters("@pRefPath") = strPath
    cmd.Parameters("@pRefFileName") = strFileName
    cmd.Parameters("@pRefSequence") = intSequence
    cmd.Parameters("@pRefURL") = ""
    'New Fields 9/17/09
    cmd.Parameters("@pRefDesc") = ""
    cmd.Parameters("@pRefOnReport") = ""
    cmd.Parameters("@pURLOnReport") = ""
    
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


'' 20120416 KD Added function:


Private Function AddAttachedDocumentToDb(strPathFileName As String, strPath As String, strFileName As String, intSequence As Integer, _
            lDocType As Long, Optional sThisPayerName As String, Optional lThisPayerNameId As Long, Optional sOrigFileName As String) As Boolean
Dim strProcName As String
On Error GoTo Block_Err
Dim myCode_ADO As clsADO
Dim colPrms As ADODB.Parameters
Dim prm As ADODB.Parameter
Dim LocCmd As New ADODB.Command
Dim iResult As Integer
Dim iRowID As Long
Dim strErrMsg As String
Dim cmd As ADODB.Command
Dim iJobId As Integer

    strProcName = ClassName & ".AddAttachedDocumentToDb"
                
    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_CONCEPT_References_Insert_NEW_PayerDtl"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_CONCEPT_References_Insert_NEW_PayerDtl"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pConceptID") = Me.FieldValue
    cmd.Parameters("@pPayerNameId") = lThisPayerNameId
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pRefType") = "DOC"
    
    '' KD COMEBACK this is only needed as we switch to the new attachment id's
    If IsNumeric(mstrAttachmentType) Then
        cmd.Parameters("@pRefSubType") = cstrCnlyAttachmentType
    Else
        cmd.Parameters("@pRefSubType") = mstrAttachmentType
    End If
    cmd.Parameters("@pRefLink") = strPathFileName
    
    cmd.Parameters("@pRefPath") = strPath
    cmd.Parameters("@pRefFileName") = strFileName
    cmd.Parameters("@pRefSequence") = intSequence
    cmd.Parameters("@pRefURL") = ""
    
    cmd.Parameters("@pRefDesc") = ""
    cmd.Parameters("@pRefOnReport") = ""
    cmd.Parameters("@pURLOnReport") = ""
    cmd.Parameters("@pEracTaggedClaimId") = cintTaggedClaimId
    cmd.Parameters("@pDocTypeId") = lDocType
    If sOrigFileName <> "" Then
        cmd.Parameters("@pOrigDocName") = GetFileName(sOrigFileName)
    Else
        cmd.Parameters("@pOrigDocName") = GetFileName(strPathFileName)
    End If
    
    LogMessage strProcName, , "Executing " & myCode_ADO.sqlString
    
    iResult = myCode_ADO.Execute(cmd.Parameters)

        'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        AddAttachedDocumentToDb = False
            'Err.Raise 65000, "SaveData", "Error updating Hours - " & strErrMsg
        LogMessage strProcName, "WARNING", "Error in proc: " & strErrMsg, , True
            'MsgBox "SaveData", "Error Logging Image - " & strErrMsg
    Else
        AddAttachedDocumentToDb = True
        
        iRowID = Nz(cmd.Parameters("@pRowId").Value, 0)
        
            '' Now, add it to the Converter Queue
        ' should we do this if the concept hasn't been converted?
        If IsConceptPayerType(Me.FieldValue) = True Then
'            If AddAttachedDocToConversionQueue(Me.FieldValue, iRowID, strPath, strFileName, sThisPayerName, False) = 0 Then
'                LogMessage strProcName, "WARNING", "Didn't add to converter queue!"
'            End If
        End If
    End If
    
Block_Exit:
    Set cmd = Nothing
    Set myCode_ADO = Nothing
    Exit Function

Block_Err:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    AddAttachedDocumentToDb = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Private Function UpdateSequence(ConceptID As String, intSequence As Integer, intRowID As Integer) As Boolean
    Dim myCode_ADO As clsADO
    Dim colPrms As ADODB.Parameters
    Dim prm As ADODB.Parameter
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
    On Error GoTo ErrHandler
    Dim cmd As ADODB.Command
    

'ALTER Procedure [dbo].[usp_CONCEPT_References_Insert]
'    @pCnlyClaimNum varchar(30),
'    @pCreateDt datetime,
'    @pRefType varchar(20),
'    @pRefSubType varchar(20),
'    @pRefLink varchar(1000),
'    @pErrMsg varchar(255) output
'as

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_CONCEPT_References_UpDate"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_CONCEPT_References_Update"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pConceptID") = Me.FieldValue
    cmd.Parameters("@pCreateDt") = Null
    cmd.Parameters("@pRefType") = Null
    cmd.Parameters("@pRefSubType") = Null
    cmd.Parameters("@pRefLink") = Null
    'New Fields
    cmd.Parameters("@pRefPath") = Null
    cmd.Parameters("@pRefFileName") = Null
    cmd.Parameters("@pRefSequence") = intSequence
    cmd.Parameters("@pRefURL") = Null
    cmd.Parameters("@pRowID") = intRowID
    iResult = myCode_ADO.Execute(cmd.Parameters)


        UpdateSequence = True

    
Exit_Function:
    Set cmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    UpdateSequence = False
    Resume Exit_Function
End Function


Private Function DeleteDocumentation(ConceptID As String, intSequence As Integer, intRowID As Integer) As Boolean
Dim myCode_ADO As clsADO
Dim colPrms As ADODB.Parameters
Dim prm As ADODB.Parameter
Dim LocCmd As New ADODB.Command
Dim iResult As Integer
Dim strErrMsg As String
On Error GoTo ErrHandler
Dim cmd As ADODB.Command
    

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_CONCEPT_References_Delete"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_CONCEPT_References_Delete"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pRowID") = intRowID
    iResult = myCode_ADO.Execute(cmd.Parameters)
    
    DeleteDocumentation = True

    
Exit_Function:
    Set cmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    DeleteDocumentation = False
    Resume Exit_Function
End Function

Private Sub cmdMoveUp_Click()

Dim strSQL As String
Dim rs As DAO.RecordSet
Dim intCurrentRowID As Integer
Dim intTotalRecordCount As Integer
Dim intOldSequence As Integer
Dim intNewSequence As Integer

Dim bWorked As Boolean
Dim intRowID As Integer


strSQL = Me.RecordSource
Set rs = CurrentDb.OpenRecordSet(strSQL)
rs.MoveLast

intTotalRecordCount = Me.txtRecordCount


intCurrentRowID = Me.RowID
intOldSequence = Me.RefSequence

If intOldSequence > 1 Then

    intNewSequence = intOldSequence - 1
    
    rs.MoveFirst
    
    Do While Not rs.EOF
        If rs!RefSequence = intOldSequence Then
            bWorked = UpdateSequence(rs!ConceptID, 1001, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    Do While Not rs.EOF
        If rs!RefSequence = intNewSequence Then
            bWorked = UpdateSequence(rs!ConceptID, intOldSequence, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    Do While Not rs.EOF
        If rs!RefSequence = 1001 Then
            bWorked = UpdateSequence(rs!ConceptID, intNewSequence, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    

End If

        
'Me.RecordSource = "SELECT * FROM v_CONCEPT_References WHERE ConceptID = '" & Me.Parent.Form.txtConceptID & "'" & " ORDER BY ConceptId, RefSequence"

Me.Requery

End Sub


Private Sub cmdMoveDown_Click()

Dim strSQL As String
Dim rs As DAO.RecordSet
Dim intCurrentRowID As Integer
Dim intTotalRecordCount As Integer
Dim intOldSequence As Integer
Dim intNewSequence As Integer

Dim bWorked As Boolean
Dim intRowID As Integer


strSQL = Me.RecordSource
Set rs = CurrentDb.OpenRecordSet(strSQL)
rs.MoveLast

intTotalRecordCount = rs.recordCount
intCurrentRowID = Me.RowID
intOldSequence = Me.RefSequence



If intOldSequence < intTotalRecordCount Then


    intNewSequence = intOldSequence + 1
    
    rs.MoveFirst
    Do While Not rs.EOF
        If rs!RefSequence = intOldSequence Then
            bWorked = UpdateSequence(rs!ConceptID, 1001, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    Do While Not rs.EOF
        If rs!RefSequence = intNewSequence Then
            bWorked = UpdateSequence(rs!ConceptID, intOldSequence, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    Do While Not rs.EOF
        If rs!RefSequence = 1001 Then
            bWorked = UpdateSequence(rs!ConceptID, intNewSequence, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
End If

        
'Me.RecordSource = "SELECT * FROM v_CONCEPT_References WHERE ConceptID = '" & Me.Parent.Form.txtConceptID & "'" & " ORDER BY ConceptId, RefSequence"
Me.Requery

End Sub



Private Sub cmdDeleteRecord_Click()
On Error GoTo Block_Err
Dim strProcName As String
' TK added 2/24/2010
Dim MyAdo As clsADO
Dim rst As ADODB.RecordSet
Dim oRS2 As ADODB.RecordSet
Dim strSQL As String
Dim fso As New FileSystemObject
Dim strFilePath As String
Dim bWorked As Boolean
     
     
     strProcName = ClassName & ".cmdDeleteRecord_Click"
     
    ' TK added 2/24/2010
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
    strSQL = ""
    strSQL = strSQL & "select OGRefLink AS reflink from v_CONCEPT_References "
    strSQL = strSQL & " where ConceptID = '" & Me.ConceptID & "'"
    strSQL = strSQL & " and RowID = " & Me.RowID
    strSQL = strSQL & " and RefSequence = " & CStr("" & Me.RefSequence)
'    Set rst = myado.OpenRecordSet(StrSQL)
    MyAdo.SQLTextType = sqltext
    MyAdo.sqlString = strSQL
    Set rst = MyAdo.ExecuteRS
    strSQL = ""

    
        ' Delete file TK added 2/24/2010
    rst.MoveFirst
    strFilePath = rst!RefLink
    If fso.FileExists(strFilePath) = False Then
        LogMessage TypeName(Me) & ".cmdDeleteRecord_Click", "ERROR", "File to delete not found!", strFilePath, False
        MsgBox "This file may have already been deleted.. Removing it from the database", vbOKOnly, "Already deleted"
        bWorked = True
    Else
        If MsgBox("Are you sure you want to delete this file? " & strFilePath, vbYesNo) = vbNo Then
            GoTo Block_Exit
        End If
        
            ' Delete the files, if that works then we can delete the records
        bWorked = DeleteFile(strFilePath, False)
        If bWorked = False Then
            Stop
        End If
        
            '' we aren't capturing this because it may not have been converted..
        Call DeleteConvertedFile(Me.RowID)
                    

    
    End If
    rst.Close
    
    '' If it worked, then we can remove from the database
    
    If bWorked = False Then
        LogMessage strProcName, "ERROR", "There was a problem deleting a file"
        GoTo Block_Exit
    End If
    
        '' this call deletes it from the database
    If DeleteDocumentation(Me.ConceptID, Me.RefSequence, Me.RowID) = False Then

        LogMessage strProcName, "ERROR", "There was a problem removing it from the database"
        GoTo Block_Exit
    End If
    
    
    
        '' Fix the sequence numbers
Dim rs As DAO.RecordSet
Dim intCurrentRowID As Integer
Dim intTotalRecordCount As Integer
Dim intOldSequence As Integer
Dim intNewSequence As Integer
Dim intRowID As Integer

    
    intCurrentRowID = Me.RowID
    intTotalRecordCount = Me.txtRecordCount
    intOldSequence = Me.RefSequence


    intNewSequence = intOldSequence - 1
    
    strSQL = Me.RecordSource
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    Do While Not rs.EOF
        If rs!RefSequence > intOldSequence Then
            bWorked = UpdateSequence(rs!ConceptID, rs!RefSequence - 1, rs!RowID)
        End If
        rs.MoveNext
    Loop
    rs.Close
    
Block_Exit:
    ' reset / refresh the form
    Me.RecordSource = "SELECT * FROM v_CONCEPT_References WHERE ConceptID = '" & Me.Parent.Form.txtConceptID & "'" & " ORDER BY ConceptId, RefSequence"
    
    Call cmbPayer_Change
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub





Private Function UpdateURL(ConceptID As String, intSequence As Integer, intRowID As Integer) As Boolean
    Dim myCode_ADO As clsADO
    Dim colPrms As ADODB.Parameters
    Dim prm As ADODB.Parameter
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
    On Error GoTo ErrHandler
    Dim cmd As ADODB.Command
    Dim strURL As String
    
    strURL = InputBox("Enter the Hyperlink", "Add Link", "http://")
    If Nz(strURL) = "" Then
        Err.Raise 65000, , "Action Cancelled URL Not Specified"
    End If

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_CONCEPT_References_UpDate"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_CONCEPT_References_Update"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pConceptID") = Me.FieldValue
    cmd.Parameters("@pCreateDt") = Null
    cmd.Parameters("@pRefType") = Null
    cmd.Parameters("@pRefSubType") = Null
    cmd.Parameters("@pRefLink") = Null
    'New Fields
    cmd.Parameters("@pRefPath") = Null
    cmd.Parameters("@pRefFileName") = Null
    cmd.Parameters("@pRefSequence") = intSequence
    cmd.Parameters("@pRefURL") = strURL
    cmd.Parameters("@pRowID") = intRowID
    iResult = myCode_ADO.Execute(cmd.Parameters)


    UpdateURL = True

    
Exit_Function:
    Set cmd = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    UpdateURL = False
    Resume Exit_Function
End Function


Private Sub frmAttachmentSelection_AttachmentSelected(strAttachmentType As String, sCnlyDocTypeID As String)
    mstrAttachmentType = strAttachmentType
    cstrCnlyAttachmentType = sCnlyDocTypeID '' 20120416 KD added
End Sub

'' 20120416 KD added the following...
Private Sub frmAttachmentSelection_EracRequiredDocTypeFound(oReqdDocType As clsConceptReqDocType)
    Set coEracReqDocType = oReqdDocType
End Sub

Private Sub frmAttachmentSelection_NewNameOfFileGenerated(sNewFileName As String)
    Me.FileNewFileName = sNewFileName
End Sub

Private Sub frmAttachmentSelection_PayersSelected(sPayerNameIds As String, sPayerNames As String)
    csPayerNames = Split(sPayerNames, ",")
    csPayerNameIds = Split(sPayerNameIds, ",")
    
    Me.PayerNameIds = sPayerNameIds
    Me.PayerNames = sPayerNames
    
End Sub

Private Sub frmAttachmentSelection_TaggedClaimSelected(intEracTaggedClaimId As Integer)
    cintTaggedClaimId = intEracTaggedClaimId
End Sub
'' 20120416 KD End added
