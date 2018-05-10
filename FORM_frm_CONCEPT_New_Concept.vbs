Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mbRecordChanged As Boolean

Private mbInsert As Boolean
Private mstrUserProfile As String
Private miAppPermission As Integer
Private mbAllowView As Boolean
Private mbAllowChange As Boolean
Private mbAllowDelete As Boolean
Private mbAllowAdd As Boolean
Private mbLocked As Boolean
Private mstrConceptID As String
Private mstrClientIssueNum As String
Private mrsConceptHdr As ADODB.RecordSet
Private mrsConceptPayerDtl As ADODB.RecordSet

Public Event ConceptSaved(strNewConceptId As String)


Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

''' HISTORY:
'' 04/25/2012   KD: various modifications.. (sorry for the lack of detail!!!)
'' 03/12/2012   KD: Added: LockFieldsIfPkgCreated,
''      modified: cmdRequestClientIssueId_Click, added call to LockFieldsIfPkgCreated() in RefreshData in order
''      to lock certain fields that should not change once a PackageID has been requested.
''
''


Const CstrFrmAppID As String = "ConceptHdr"
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property
Property Set ConceptRecordSource(data As ADODB.RecordSet)
     Set mrsConceptHdr = data
End Property
Property Get ConceptRecordSource() As ADODB.RecordSet
     Set ConceptRecordSource = mrsConceptHdr
End Property
Property Let Insert(data As Boolean)
    mbInsert = data
End Property
Property Get Insert() As Boolean
    Insert = mbInsert
End Property
Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property
Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property
Property Let FormConceptID(data As String)
    mstrConceptID = data
End Property
Property Get FormConceptID() As String
    FormConceptID = mstrConceptID
End Property


    ''' KD: I typically use this for debugging / logging so that I know (from a log file)
    ''' where the procedure is in the stack at any given time..
Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Sub RefreshData()
    'Refresh the main form
    
    On Error GoTo ErrHandler
    
    
    Dim strSQL As String
    Dim ctl As Variant
        
    Dim iAppPermission As Integer
    Me.cmdSave.SetFocus

    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    miAppPermission = GetAppPermission(Me.frmAppID)
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowView = (miAppPermission And gcAllowView)
    
    If mbAllowChange = True Then
        Me.ConceptDesc.SetFocus
        Me.cmdSave.Enabled = True
    Else
        Me.ConceptDesc.SetFocus
        Me.cmdSave.Enabled = False
    End If
        
    Me.Caption = "Concept: " & Me.FormConceptID
    
    Set Me.RecordSet = Nothing
    Set Me.RecordSet = mrsConceptHdr
        
    'Loop through the controls setting their control source to the recordset
    For Each ctl In Me.Controls
        If ctl.Tag <> "" Then
            If InStr(1, ctl.Tag, ".", vbTextCompare) > 0 Then
                Select Case UCase(left(ctl.Tag, InStr(1, ctl.Tag, ".", vbTextCompare) - 1))
                Case "CONCEPT_HDR"  '   , "BOTH"
                    If isField(mrsConceptHdr, ctl.Name) = True Then
                        Me.Controls(ctl.Name).ControlSource = mrsConceptHdr.Fields(ctl.Name).Name
                    End If
'                Case "CONCEPT_PAYER_DTL"    ' not going to do this now because we will take care of it on save
                                            ' and, we don't want Concept_HDr to be bound to the controls that are both because
                                            ' the ultimate goal is to have Concept_Hdr null for the stuff that goes into
                                            ' the payer specific stuff.
'    Stop
    '                If isField(mrsConceptPayerDtl, ctl.Name) = True Then
    '                    Me.Controls(ctl.Name).ControlSource = mrsConceptPayerDtl.Fields(ctl.Name).Name
    '                End If
                Case Else
                    Me.Controls(ctl.Name).ControlSource = ""
                End Select
            End If
        End If
    Next

    '' Only this one - for new concepts:
    If IsNull(Me.Auditor) Then
        Me.Auditor = Identity.UserName()
        Me.ConceptStatus = 100  ' 100 = new concept
    End If

'    LockFieldsIfPkgCreated

Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_Concept_hdr : RefreshMain"
End Sub


Private Sub SaveData()
Dim bResult As Boolean
Dim tErrTxt As String
Dim strProcName As String
Dim oConcept As clsConcept
Dim sErrMsg As String
Dim oPayerFrm As Form_frm_PAYERNAMES
On Error GoTo Block_Err
    
    strProcName = ClassName & ".SaveData"
    
'    If mbRecordChanged = False And Me.Dirty = False Then
'        MsgBox "There are no changes to save."
'        Exit Sub
'    End If
    
    'Alex C 09062011 - Added ConceptRationale to list of required field checks.  Changed error message to identify which fields are missing
    'on this record
    tErrTxt = ""
    If Trim(Nz(Me.ConceptDesc)) = "" Then
        tErrTxt = "Issue Name"
    End If
    
    If Trim(Nz(Me.ConceptLogic)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Issue Description"
    End If
    
    If Trim(Nz(Me.ConceptRationale)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Rationale"
    End If
    '' ConceptCatId
    If Trim(Nz(Me.ConceptCatId)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Concept Category"
    End If
    
    '' ConceptCatId
    If Trim(Nz(Me.BudgetGroup)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Budget Group"
    End If
    

    '' ConceptCatId
    If Trim(Nz(Me.ContractId)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Contract Id"
    End If
    
    'Are there any missing fields? Tell the user and exit without saving
    If Len(tErrTxt) > 0 Then
        MsgBox "Data must be entered for these field(s) before the issue can be saved; " + tErrTxt, vbOKOnly + vbExclamation
        Exit Sub
    End If
 

    If Me.ckCreateWithoutPayers = False Then

       Set oPayerFrm = Me.sfrmPayers.Form
       
       ' 20120616: KD Need to have selected the payers as of 06/19/2012
       If oPayerFrm.GetSelectedPayerNameIDs = "" Then
           MsgBox "You must select the appropriate payers for this concept", vbOKOnly + vbExclamation, "ERROR! Incomplete!"
           GoTo Block_Exit
       End If
    End If
    
    'Inserting, set some recordset values
    If Me.Insert Then
        mrsConceptHdr.MoveFirst
        mrsConceptHdr.Fields("AccountID") = gintAccountID
        mrsConceptHdr.Fields("ConceptID") = Me.FormConceptID
        mrsConceptHdr.Fields("NextContract") = 0
        mrsConceptHdr.Fields("RefreshFinancials") = 0
    End If
    
    mrsConceptHdr.Fields("LastUpDt") = Now
    mrsConceptHdr.Fields("LastUpUser") = Identity.UserName
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    bResult = myCode_ADO.Update(mrsConceptHdr, "usp_CONCEPT_Hdr_Apply")
    
    Stop
    
        
    
    '' Now, put the payer records in there
    '' We need the actual concept id now though..
    If bResult = True Then
        ' get the concept id
        
        With myCode_ADO
            .ConnectionString = GetConnectString("V_DATA_DATABASE")
            .SQLTextType = sqltext
            .sqlString = "SELECT * FROM CONCEPT_Hdr WHERE ConceptId = ( SELECT MAX(ConceptId) FROM CONCEPT_Hdr WHERE COntractId = " & CStr(Nz(Me.ContractId, 100)) & ") "
            Set mrsConceptHdr = .ExecuteRS
            If .GotData = True Then
                Me.ConceptID = mrsConceptHdr("ConceptID").Value
            Else

            End If
        End With
    End If

    

    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.ConceptID) = False Then

        LogMessage strProcName, "ERROR", "There was a problem creating the concept object. Please close Concept Management form, reopen and try again. If you get this message again, please contact support!", Me.ConceptID, True
        GoTo Block_Exit
    End If
    
    ' We need an editable recordset now so reload our recordset this now that we have the concept id..
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select * from Concept_hdr where ConceptID = '" & Me.ConceptID & "'"
    Set mrsConceptHdr = MyAdo.OpenRecordSet
    Set MyAdo = Nothing

    Dim oRs As RecordSet
    
    If Me.ckCreateWithoutPayers = True Then
        ' We need to assign a Client Issue Number:
        If mod_Concept_Specific.IssueClientIssueNumToHdr(oConcept, sErrMsg) = "" Then
            LogMessage strProcName, "ERROR", "There was a problem getting the clientissue num!"
        End If
    Else
        Set oRs = sfrmPayers.Form.RecordSet
        With oRs
            .MoveFirst
            
            While Not .EOF
                ' skip the "All"
                If oRs("PayerNameID").Value > 1000 Then
                    Call GetPayerDtlRS
                    If oRs("Selected").Value = True Then
                        LogMessage strProcName, , "Payer ID: " & CStr("" & oRs("PayerNameID").Value)
                        Call InsertDetailForPayer(Me, mrsConceptPayerDtl, oRs("PayerNameId").Value, mrsConceptHdr, True)
                        ' Assign the Client Issue ID:
                        If mod_Concept_Specific.IssueClientIssueNum(oConcept, oRs("PayerNameID").Value, sErrMsg) = "" Then
                            LogMessage strProcName, "ERROR", "There was a problem getting the clientissue num!"
                        End If
                        
                    End If
                End If
                .MoveNext
            Wend
            
        End With
    End If
    
    '' Now, to cover the stuff that's BOTH we update the CONCEPT_Hdr table (AGAIN) - I know this is a stupid way to do this!
    ' but I had like 3 business days to convert the entire concept submission process!
    ''
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    bResult = myCode_ADO.Update(mrsConceptHdr, "usp_CONCEPT_Hdr_Apply")
    
    
    If bResult Then
        MsgBox "Record Saved", vbOKOnly
        RaiseEvent ConceptSaved(oConcept.ConceptID)
        If IsSubForm(Me) = False Then
            Me.Dirty = False
            mbRecordChanged = False
            DoCmd.Close acForm, Me.Name
        Else
            Me.Dirty = False
            mbRecordChanged = False
        End If
       
    Else
        Err.Raise 65000, , "Error Saving Record"
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_Concept_hdr : RefreshMain"
    GoTo Block_Exit
End Sub


'Private Sub InsertDetailForPayer(intPayerNameId As Integer)
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim bResult As Boolean
'Dim oCtl As Control
'
'
'    strProcName = ClassName & ".InsertDetailForPayer"
'
'    Call GetPayerDtlRS
'
'    If mrsConceptPayerDtl.RecordCount < 1 Then
'        mrsConceptPayerDtl.AddNew
'    Else
'        Stop    ' problem!
'    End If
'
'    'Loop through the controls setting their control source to the recordset
'    For Each oCtl In Me.Controls
'        If oCtl.Tag <> "" Then
''Debug.Assert oCtl.Name <> "ConceptStatus"
'            If InStr(1, oCtl.Tag, ".", vbTextCompare) > 0 Then
'                Select Case UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
'                Case "CONCEPT_HDR"
''                    Stop    ' do nothing
'
'                Case "CONCEPT_PAYER_DTL"
'
'                    If isField(mrsConceptPayerDtl, oCtl.Name) = True Then
'                        mrsConceptPayerDtl.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name).Value
'                    Else
'                        Stop
'                    End If
'                Case "BOTH"
'                    If isField(mrsConceptPayerDtl, oCtl.Name) = True Then
'                        mrsConceptPayerDtl.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name).Value
'                    Else
'                        Stop
'                    End If
'                    If mrsConceptHdr Is Nothing Then
'                        Stop    ' should never get here
'                    End If
'                    If mrsConceptHdr.RecordCount < 1 Then
'                        Stop    ' should never get here
'                    End If
'
'                    If isField(mrsConceptHdr, oCtl.Name) = True Then
'                        mrsConceptHdr.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name).Value
'                    End If
'                End Select
'            End If
'        End If
'    Next
'
'    mrsConceptPayerDtl("PayerNameId") = intPayerNameId
'    mrsConceptPayerDtl("ConceptID") = Me.ConceptID
''    mrsConceptPayerDtl("ConceptIdPayerId_RowId") = 0    ' not used
'
'    mrsConceptPayerDtl.Update
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("V_CODE_DATABASE")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_CONCEPT_PAYER_Dtl_Apply"
'        bResult = .Update(mrsConceptPayerDtl, "usp_CONCEPT_PAYER_Dtl_Apply")
'        If bResult = False Then
'            Stop
'        End If
'    End With
'
'
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Sub

Private Function GetPayerDtlRS() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".GetPayerDtlRS"

'    sSql = "SELECT * FROM CONCEPT_PAYER_Dtl WHERE ConceptId = '" & Me.ConceptID & "' "

    sSql = "SELECT * FROM CONCEPT_PAYER_Dtl WHERE 1 = 2"

    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("V_DATA_DATABASE")
    oCn.Open
    
    Set mrsConceptPayerDtl = New ADODB.RecordSet
    mrsConceptPayerDtl.CursorLocation = adUseClientBatch
    mrsConceptPayerDtl.CursorType = adOpenStatic
    mrsConceptPayerDtl.LockType = adLockBatchOptimistic
    
    mrsConceptPayerDtl.Open sSql, oCn
    Set mrsConceptPayerDtl.ActiveConnection = Nothing


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function

Private Sub cmdConceptSQL_Click()
    If Me.ConceptID <> "" Then
        Call GetConceptStoredProcSQL(Me.ConceptID)
    End If
End Sub



Private Sub RenameConceptFiles(rst As ADODB.RecordSet, strFilePath As String, strConceptID As String)
' TK: renaming staging files according to conceptID format

    Dim intCount As Integer
    Dim fso As Variant
    Dim strOriginalFile As String
    Dim strRenamedFile As String

    
    'intCount = rst.RecordCount - 2
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If rst.EOF = True And rst.BOF = True Then
        MsgBox "No record for RenameConceptFiles "
        Exit Sub
    End If
    
    
    With rst
        .MoveFirst
        intCount = 1
        Do While intCount <= (.recordCount - 2)
            'strrefilename = rst.Fields("reffilename")
            Debug.Print "Count = " & intCount
            Debug.Print "RefSequence = " & rst.Fields("RefSequence")
            strOriginalFile = strFilePath & "\" & rst.Fields("RefFileName")
            strRenamedFile = strFilePath & "\" & strConceptID & "_" & rst.Fields("RefSequence") & Right(rst.Fields("RefFileName"), 4)
            fso.MoveFile strOriginalFile, strRenamedFile
            
            intCount = intCount + 1
            .MoveNext
        Loop
    End With
        
        
    
    
    

End Sub

Private Function ExportRsToExcel(rst As ADODB.RecordSet, sExcelFileAndPath As String) As Boolean
'Function to export recordset to excel file
    Dim fso As Variant
    Dim cie As clsImportExport
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set cie = New clsImportExport


    If rst.recordCount > 65535 Then
        MsgBox "Warning: Your recordset contains more than 65535 rows, the maximum number of rows allowed in Excel.  " & _
        Trim(str(rst.recordCount - 65535)) & " rows will not be displayed.", vbCritical
    End If


    If fso.FileExists(sExcelFileAndPath) Then
        fso.DeleteFile sExcelFileAndPath
        'MsgBox "file deleted", vbOKOnly
    End If

    
    With cie
        .ExportExcelRecordset rst, sExcelFileAndPath, True
    End With

    ExportRsToExcel = True
    
exitHere:
    Set cie = Nothing
    Exit Function
    
HandleError:
    ExportRsToExcel = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function


Private Sub ckCreateWithoutPayers_AfterUpdate()
    Me.sfrmPayers.visible = False
End Sub


'''
'''Private Sub cmdRunReport_Click()
'''Dim tErrTxt As String
'''Dim oConcept As clsConcept
'''Dim sPromptMsg As String
'''    'Alex C 09062011 - Changed error message to identify which fields are missing on this record
'''    tErrTxt = ""
'''    If Trim(Nz(Me.ConceptDesc)) = "" Then
'''        tErrTxt = "Issue Name"
'''    End If
'''
'''    If Trim(Nz(Me.ConceptLogic)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Issue Description"
'''    End If
'''
'''    If Trim(Nz(Me.ConceptReferences)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "References"
'''    End If
'''
'''    If Trim(Nz(Me.Comments)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Detailed Explanation of References"
'''    End If
'''
'''    If Trim(Nz(Me.ReferralFlag)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Referral"
'''    End If
'''
'''    If Trim(Nz(Me.OpportunityType)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Overpayment or Underpayment"
'''    End If
'''
'''    If Trim(Nz(Me.ReviewType)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Review Type"
'''    End If
'''
'''    If Trim(Nz(Me.DataType)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Data Type"
'''    End If
'''
'''    If Trim(Nz(Me.ProviderTypeID)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Provider Type"
'''    End If
'''
'''    If Trim(Nz(Me.ErrorCode)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Error Code"
'''    End If
'''
'''    If Trim(Nz(Me.ConceptPriority)) = "" Then
'''        If Len(tErrTxt) > 0 Then
'''            tErrTxt = tErrTxt + ", "
'''        End If
'''        tErrTxt = tErrTxt + "Priority"
'''    End If
'''
'''    'Are there any missing fields? Tell the user and exit without saving
'''    If Len(tErrTxt) > 0 Then
'''        MsgBox "Data must be entered for these field(s) before the report can be run; " + tErrTxt, vbOKOnly + vbExclamation
'''        Exit Sub
'''    End If
'''
''''    If GetUserProfile = "CM_Admin" Then
'''
'''
'''    Set oConcept = New clsConcept
'''    If oConcept.LoadFromID(Me.FormConceptID) = False Then
'''        ' hmm.. weird
'''        LogMessage ClassName & ".cmdRunReport", "ERROR", "Could not load the concept object!?!?!", Me.FormConceptID
'''        DoCmd.OpenReport "rpt_CONCEPT_New_Issue", acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
'''    Else
'''
'''        If oConcept.NIRF_Exists = True Then
'''            sPromptMsg = "There is an existing NIRF. Do you want to replace it (Yes) or just view it (No)?"
'''        Else
'''            sPromptMsg = "Do you want to save this as a PDF for submission to CMS?"
'''        End If
'''
'''
'''        If MsgBox(sPromptMsg, vbYesNo, "Save?") = vbYes Then
'''            If CreatePackageNirf(Me.FormConceptID, True, True) = False Then
'''                LogMessage TypeName(Me) & ".cmdRunReport", "ERROR", "There was an error converting to PDF, opening as an MS Access report - please print to PDF", , True
'''                DoCmd.OpenReport "rpt_CONCEPT_New_Issue", acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
'''            End If
'''        Else
'''            DoCmd.OpenReport "rpt_CONCEPT_New_Issue", acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
'''        End If
'''
'''    End If
'''
'''
'''
'''
'''End Sub


Private Sub cmdSave_Click()
Dim oForm As Form_frm_PAYERNAMES

    ' Make sure at least 1 payer is selected (and not JUST all)
    Set oForm = Me.sfrmPayers.Form
    If Me.ckCreateWithoutPayers = False Then
        If oForm.AtLeastOnePayerSelected = False Then
            MsgBox "At least 1 Payer must be selected! Please select the appropriate payer(s) and click save when ready!", vbCritical, "No Payers selected!"
            Me.sfrmPayers.SetFocus
            Exit Sub
        End If
    End If
    
    SaveData
End Sub





Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub

Private Sub Form_Load()
Dim oFrm As Form_frm_PAYERNAMES
    
    Set mrsConceptHdr = New ADODB.RecordSet '' 20120416 KD: Early bound thanks!!
    
    If IsSubForm(Me) = False Then
        Me.FormConceptID = "NEW"
       
        If Me.FormConceptID <> "" Then
            Me.Insert = True
            Set MyAdo = New clsADO
            MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
            MyAdo.sqlString = "select * from Concept_hdr where ConceptID = '" & Me.FormConceptID & "'"
            Set mrsConceptHdr = MyAdo.OpenRecordSet
            Set MyAdo = Nothing
        End If
        RefreshData
    End If
    DoCmd.SetWarnings False

    Set oFrm = Me.sfrmPayers.Form
    oFrm.ShowPastPayers = False


End Sub



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
    
    cmd.Parameters("@pConceptID") = Me.FormConceptID
    cmd.Parameters("@pCreateDt") = Now
    cmd.Parameters("@pRefType") = "DOC"
    cmd.Parameters("@pRefSubType") = "ATTACH"
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



Private Sub Form_Unload(Cancel As Integer)
If mbAllowChange Then
    If Not (mbRecordChanged = False And Me.Dirty = False) Then
        If MsgBox("Record has changed. Would you like to save changes to Concept - " & Me.FormConceptID & "?", vbYesNo + vbQuestion) = vbYes Then
            SaveData
        End If
    End If
End If
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "ADO ERROR"
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "ADO ERROR"
End Sub



Private Function IsFileOpen(FileName As String)
    Dim iFileNum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFileNum = FreeFile()
    Open FileName For Input Lock Read As #iFileNum
    Close iFileNum
    iErr = Err
    On Error GoTo 0
     
    Select Case iErr
    Case 0:    IsFileOpen = False
    Case 53:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: Error iErr
    End Select
     
End Function


Private Sub Pause(Duration As Integer)
    Dim Current As Double
    Current = Timer
    
    Do Until Timer - Current >= (Duration / 1000)
        'Debug.Print (Timer - Current)
        DoEvents
    Loop
    
End Sub



''' This will lock the fields
'''
Private Sub LockFieldsIfPkgCreated()
On Error GoTo Block_Err
Dim strProcName As String
Dim bEnabled As Boolean
Dim oControl As Control
Dim oConcept As clsConcept

    strProcName = ClassName & ".LockFieldsIfPkgCreated"
    
'    ClientIssueNum.SetFocus
'    If Nz(Me.ClientIssueNum.Text, "") <> "" Then
'        bEnabled = False
'    Else
'        bEnabled = True
'    End If
    Me.ConceptDesc.SetFocus
    
    
    
    Me.ReviewType.Enabled = True
    Me.DataType.Enabled = True
    Me.ErrorCode.Enabled = True
    Me.ErrorCode2.Enabled = True
    Me.ProviderTypeID.Enabled = True
    
    
    Me.ClientIssueNum.Locked = Not bEnabled
    If bEnabled = False Then
        Me.ClientIssueNum.BackColor = Me.Detail.BackColor
    Else
        Me.ClientIssueNum.BackColor = 16777215  '' White
    End If

    Set oConcept = New clsConcept
    If IsNull(Me.ConceptID) = False Then
        Call oConcept.LoadFromId(Me.ConceptID)
    End If

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub
