Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' HISTORY:
'' 09/12/2012   KD: made a lot of changes mostly submission and validation related
'' 09/10/2012   KD: Locked Create NIRF & Sample Claims Docs
''              - Fixed some stuff with the sample claims doc
''              - fixed the NIRF (even though that didn't change this code..)
'' 08/22/2012   KD: cleanup some junk..
'' 07/10/2012   KD: MASSIVE changes due to CMS concept submission changes to payer specific
'' 04/25/2012   KD: various modifications.. (sorry for the lack of detail!!!)
'' 03/12/2012   KD: Added: LockFieldsIfPkgCreated,
''      modified: cmdRequestClientIssueId_Click, added call to LockFieldsIfPkgCreated() in RefreshData in order
''      to lock certain fields that should not change once a PackageID has been requested.
''
''

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
Private mrsPayerDtl As ADODB.RecordSet
Private mrsThisConceptPayers As ADODB.RecordSet

Public Event ConceptSaved()
Public Event ConceptIdChanged(sNewConceptId As String)
Public Event PayerNameIdChanged(lNewPayerNameId As Long)

Private Const cs_ADR_NOT_SELECTED_MSG  As String = "-- NOT SELECTED --"

Private clngPayerNameId As Long
Private cblnIsPayerSetToAll As Boolean
Private cblnOnly1Payer As Boolean   ' we need this so we can show the only payer detail when concept header level detail is selected..
                                    ' this gets a little message I do admit but it's nice for the user to not have to select
                                    ' the payer and back to the concept in order to see everything

Private cdctPrevVals As Scripting.Dictionary

Private cstrResubmitPkgFldr As String

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Private Const csSubmitBtnText As String = "Create Package"
Private Const csReSubmitBtnText As String = "Re-Package"

Private cintCurPage As Integer
Private coPayers As Collection

Private coCurConcept As clsConcept

Private Const CstrFrmAppID As String = "ConceptHdr"
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
Public Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property
Public Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property
Public Property Let FormConceptID(data As String)
    If data <> mstrConceptID Then
        RaiseEvent ConceptIdChanged(data)
    End If
    Me.txtSelectedId = data
    mstrConceptID = data
End Property
Public Property Get FormConceptID() As String
    If mstrConceptID = "" Then
        If Me.visible = True Then
'            LogMessage ClassName & ".FormConceptID", "Out of Synch", "Somehow Concept Management got out of synch. Please close Concept Mgmt and reload to resolve", , True
        End If
    Else
        FormConceptID = mstrConceptID
    End If
End Property


Public Property Let CurrentPageSelected(iPageToSel As Integer)
    cintCurPage = iPageToSel
    If Me.TabCtl128 Is Nothing Then
        Stop
'    Else
'        Stop
    End If
    If Me.TabCtl128 <> iPageToSel Then
        Me.TabCtl128 = iPageToSel
    End If
    giHdrFormSelectedPage = iPageToSel
End Property
Public Property Get CurrentPageSelected() As Integer
    CurrentPageSelected = cintCurPage
End Property

' Unbound now so me.Dirty isn't going to work, therefore:
Public Property Get IsDirty() As Boolean
    IsDirty = DidAnythingChange()
End Property


Public Property Get PayerNameId() As Long
    If clngPayerNameId = 0 Then
        clngPayerNameId = Nz(Me.cmbPayer, 1000)
    End If
    PayerNameId = clngPayerNameId
End Property
Public Property Let PayerNameId(lngPayerNameId As Long)
    If lngPayerNameId <> clngPayerNameId Then
        RaiseEvent PayerNameIdChanged(lngPayerNameId)
    End If

    clngPayerNameId = lngPayerNameId
End Property


Public Property Get IsPayerSetToAll() As Boolean
'    IsPayerSetToAll = True
    If Nz(cmbPayer, 1000) = 1000 Or cmbPayer = 0 Then
        IsPayerSetToAll = True
    Else
        IsPayerSetToAll = False
    End If
End Property
'Public Property Let IsPayerSetToAll(blnIsPayerSetToAll As Boolean)
'    cblnIsPayerSetToAll = blnIsPayerSetToAll
'End Property



''' KD: I typically use this for debugging / logging so that I know (from a log file)
''' where the procedure is in the stack at any given time..
Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get IsConceptPayerSpecific() As Boolean
    ' false = old concept
    ' true = new as of 2/16/2012 -> whenever they change this again
Dim bConceptIsNew As Boolean

        ' enable if: This concept has NO payers
        ' AND it hasn't been submitted yet
    If mrsThisConceptPayers Is Nothing Then
        bConceptIsNew = False
        GoTo Block_Exit
    End If
    
        
        ' enable if: This concept has NO payers
        ' AND it hasn't been submitted yet
    If mrsThisConceptPayers.recordCount < 2 Then
        bConceptIsNew = False
    Else
        bConceptIsNew = True
    End If
    
Block_Exit:
    IsConceptPayerSpecific = bConceptIsNew
End Property


Public Sub RefreshData()     'Refresh the main form
On Error GoTo ErrHandler
Dim strSQL As String
Dim ctl As Variant
Dim iAppPermission As Integer
Dim bLoadConcept As Boolean

    Me.cmdSave.SetFocus
    cmdSubmit.Enabled = False

    CurrentPageSelected = giHdrFormSelectedPage
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
    miAppPermission = GetAppPermission(Me.frmAppID)
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowView = (miAppPermission And gcAllowView)
    
    If mbAllowChange = True Then
        Me.LastUpDt.SetFocus
'        Me.ConceptDesc.SetFocus
        Me.cmdSave.Enabled = True
    Else
'        Me.ConceptDesc.SetFocus
        Me.LastUpDt.SetFocus
        Me.cmdSave.Enabled = False
    End If
        
    Me.Caption = "Concept: " & Me.FormConceptID
    

    
    Call GetPayersForConceptRS
    Call GetPayerDtlRS
    Call GetHdrRS

    
    Call GetOrSetControlValues(True)
    Call EnableDisableConvertWizard
    Call SetupPrevValsDict

    Call FilterPayerCombo
    Call EnableAndRecolorFields
    
    Call CreateConceptObjIfNeeded
    
        '' Set the time that we refreshed the data here so we don't repeatedly refresh for a single action..
    If IsSubForm(Me) = True Then
        If Me.Parent.Name = "frm_CONCEPT_Main" Then
            Call Me.Parent.SetSubformRefreshTime(Me.Name)
            Me.TabCtl128 = CurrentPageSelected
        End If
    End If

'    Select Case Me.ReviewType
'    Case "A", "S"
'Stop
        Me.RefreshFinancials.visible = True
'    Case Else
'        Me.RefreshFinancials.visible = False
        
'Stop
'    End Select
    
    Call RefreshADRLetterCombo

    LockFieldsIfPkgCreated

Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_Concept_hdr : RefreshMain"
End Sub

Public Sub AdrLetterNotSelectedYet(Optional bSet As Boolean = False)
On Error GoTo Block_Err
Dim strProcName As String
Dim bAutomated As Boolean

    strProcName = ClassName & ".AdrLetterNotSelectedYet"
    
    If coCurConcept Is Nothing Then
        GoTo Block_Exit
    End If
    bAutomated = IIf(coCurConcept.GetField("ReviewType") = "A", True, False)
    
    If bSet = False And bAutomated = False Then
        Me.lblAdrLetter.visible = True
        Me.txtADRLetter.visible = True
        Me.cmdADRLetterSelect.visible = True
        Me.cmdADRLetterView.visible = True
    
        Me.txtADRLetter = cs_ADR_NOT_SELECTED_MSG
        Me.txtADRLetter.BackColor = RGB(255, 0, 0)
        Me.txtADRLetter.ForeColor = RGB(255, 255, 0)
        Me.txtADRLetter.BorderColor = RGB(255, 0, 0)
        Me.txtADRLetter.FontWeight = 700
        Me.txtADRLetter.TextAlign = 2
'        Stop
        Me.txtADRLetter.BorderWidth = 3
    
        Me.cmdADRLetterView.Enabled = False
    
    Else
        If bAutomated = False Then
            Me.lblAdrLetter.visible = True
            Me.txtADRLetter.visible = True
            Me.cmdADRLetterSelect.visible = True
            Me.cmdADRLetterView.visible = True
        Else
            Me.lblAdrLetter.visible = False
            Me.txtADRLetter.visible = False
            Me.cmdADRLetterSelect.visible = False
            Me.cmdADRLetterView.visible = False
        End If

        Me.txtADRLetter.BorderColor = RGB(0, 0, 0)
        Me.txtADRLetter.BackColor = RGB(255, 255, 255)
        Me.txtADRLetter.ForeColor = RGB(0, 0, 0)
        Me.txtADRLetter.FontWeight = 400
        Me.txtADRLetter.TextAlign = 1
'        Stop
        Me.txtADRLetter.BorderWidth = 0
        Me.cmdADRLetterView.Enabled = True


    End If
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub RefreshADRLetterCombo()
On Error GoTo Block_Err
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim strProcName As String
Dim sSql As String
    
    strProcName = ClassName & ".RefreshADRLetterCombo"
    
    
    sSql = "SELECT TOP 1 LetterType FROM CONCEPT_Letters WHERE ConceptId = '" & Me.FormConceptID & "'"
    If Me.IsConceptPayerSpecific = True Then
        If Me.IsPayerSetToAll = False Then
            'sSql = sSql & " AND payerNameId = " & CStr("")
        Else
            sSql = sSql & " AND payerNameId = " & CStr(Me.PayerNameId)
        End If
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Call AdrLetterNotSelectedYet
        Else
            Call AdrLetterNotSelectedYet(True)
            Me.txtADRLetter.Locked = False
            Me.txtADRLetter = oRs("LetterType").Value
            Me.txtADRLetter.Locked = True
        End If
    End With
    
    
Block_Exit:
    Set oAdo = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub PayerChange()
    cmbPayer_Change
End Sub


Private Sub CreateConceptObjIfNeeded()
'On Error Resume Next
On Error GoTo Block_Err
Dim strProcName As String
Dim bLoadConcept As Boolean
    strProcName = ClassName & ".CreateConceptObjIfNeeded"
    

    If coCurConcept Is Nothing Then
        bLoadConcept = True
    Else
        If Not mrsConceptHdr.EOF Then
            mrsConceptHdr.MoveFirst
            If coCurConcept.ConceptID <> mrsConceptHdr("ConceptID") Then bLoadConcept = True
        End If
    End If
    

    If Not mrsConceptHdr.EOF Then
       If bLoadConcept = True Then
           Set coCurConcept = New clsConcept
           If coCurConcept.LoadFromId(mrsConceptHdr("ConceptID")) <> True Then
    '           Stop
           End If
       End If
    End If
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub

''   Note: bSetFormValues = True means that we are Setting the form.control values FROM the recordsets
''      NOT setting the recordset from the form controls - that's = False!
Private Function GetOrSetControlValues(Optional bSetFormValues As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim bPayerChanges  As Boolean
Dim bResult As Boolean
Dim sUseDt As String
Dim dUseDt As Date
    
    strProcName = ClassName & ".GetOrSetControlValues"
    
    Call CreateConceptObjIfNeeded
    
        '' KD : OLD Legacy way for older concepts that aren't payer specific (sent straight to CMS for approval)
    If mrsPayerDtl.recordCount < 1 Then
            'Loop through the controls setting their control source to the recordset
        For Each oCtl In Me.Controls

            If Right(oCtl.Tag, 1) = "R" Then    ' had to change this ever so slightly (look at the right 1 char)
                If isField(mrsConceptHdr, oCtl.Name) = True Then
                    If bSetFormValues = True Then
                            Me.Controls(oCtl.Name).ControlSource = mrsConceptHdr.Fields(oCtl.Name).Name
                    Else
                        ' 20130918 KD: This is me being lazy
                        If oCtl.Name = "RepriceFlag" Then
                            mrsConceptHdr(oCtl.Name) = Nz(oCtl.Value, False)
                        Else
                            mrsConceptHdr(oCtl.Name) = oCtl.Value
                        End If

                    End If
                End If
            End If
        Next
        If bSetFormValues = False Then
            mrsConceptHdr.Update
        End If
        
    Else
    
            'Loop through the controls setting their VALUE to the recordset value (i.e. NOT bound!)
        mrsConceptHdr.MoveFirst
        mrsPayerDtl.MoveFirst
 
        For Each oCtl In Me.Controls
            If oCtl.Tag <> "" Then
'Debug.Assert oCtl.Name <> "RepriceFlag"
                If InStr(1, oCtl.Tag, ".", vbTextCompare) > 0 Then
                    Me.Controls(oCtl.Name).ControlSource = ""
                        '' If "All payers" are selected then we'll have some to aggregate:
                    
                    If Me.IsPayerSetToAll = True Then
                        Select Case UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
                        Case "CONCEPT_HDR"
                            If isField(mrsConceptHdr, oCtl.Name) = True Then

                                If bSetFormValues = True Then
                                    ' 20130918 KD: This is me being lazy
                                    If oCtl.Name = "RepriceFlag" Then
                                        Me.Controls(oCtl.Name) = Nz(mrsConceptHdr.Fields(oCtl.Name).Value, False)
                                    Else
                                        Me.Controls(oCtl.Name) = mrsConceptHdr.Fields(oCtl.Name).Value
                                    End If
                                
                                Else
                                    ' 20130918 KD: This is me being lazy
                                    If oCtl.Name = "RepriceFlag" Then
                                        mrsConceptHdr.Fields(oCtl.Name).Value = Nz(Me.Controls(oCtl.Name), False)
                                    Else
                                        mrsConceptHdr.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name)
                                    End If

                                End If
                            End If
                        Case "CONCEPT_PAYER_DTL"    ' not going to do this now because we will take care of it on save
                                                    ' and, we don't want Concept_HDr to be bound to the controls that are both because
                                                    ' the ultimate goal is to have Concept_Hdr null for the stuff that goes into
                                                    ' the payer specific stuff.
                            If isField(mrsPayerDtl, oCtl.Name) = True Then
                                If bSetFormValues = True Then
                                    
                                    Me.Controls(oCtl.Name) = coCurConcept.AggregateFields(oCtl.Name)
                                    
                                Else
                                    ' no payer changes are allowed yet (when drop down is set to concept header level)
'                                    mrsPayerDtl.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name)
'                                    bPayerChanges = True
                                End If

                            End If
                        Case "BOTH"
                                '' Aggregation functions here
                            If bSetFormValues = True Then

                            
                                If (TypeName(oCtl) = "Textbox" Or TypeName(oCtl) = "Combobox") And InStr(1, oCtl.Name, "date", vbTextCompare) < 1 Then
'                                    Me.Controls(oCtl.Name) = " Aggregated value "
                                    If oCtl.Name = "DateSubmitted" Then

                                        If IsDate(coCurConcept.AggregateFields(oCtl.Name)) = True Then
                                            dUseDt = coCurConcept.AggregateFields(oCtl.Name)
                                            sUseDt = DatePart("m", dUseDt) & "/" & DatePart("d", dUseDt) & "/" & DatePart("yyyy", dUseDt)
                                            Me.Controls(oCtl.Name) = CDate(sUseDt)
                                            Stop
                                            
                                        Else
                                            Stop
                                        End If


                                    Else
                                        Me.Controls(oCtl.Name) = coCurConcept.AggregateFields(oCtl.Name)
                                    End If
                                End If
                            
                            
                            Else
                                ' since this is both, but we are on the header level, then we need to save to the header only
                                mrsConceptHdr.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name)
'                                bPayerChanges = false
                            End If
                        End Select

                    Else    ' an individual payer is selected

                        Select Case UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
                        Case "CONCEPT_HDR"  '   , "BOTH"
                            If isField(mrsConceptHdr, oCtl.Name) = True Then
                                If bSetFormValues = True Then
                                    Me.Controls(oCtl.Name) = mrsConceptHdr.Fields(oCtl.Name).Value
                                Else
                                    mrsConceptHdr.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name)
                                End If
                            End If
                                ' Since an individual payer is selected, we just want that ones value
                        Case "CONCEPT_PAYER_DTL", "BOTH"    ' not going to do this now because we will take care of it on save
                                                    ' and, we don't want Concept_HDr to be bound to the controls that are both because
                                                    ' the ultimate goal is to have Concept_Hdr null for the stuff that goes into
                                                    ' the payer specific stuff.
                            If isField(mrsPayerDtl, oCtl.Name) = True Then
                                If bSetFormValues = True Then
                                    If oCtl.Name = "DateSubmitted" Then
                                        If IsDate(mrsPayerDtl.Fields(oCtl.Name).Value) = True Then
                                            dUseDt = mrsPayerDtl.Fields(oCtl.Name).Value
                                            sUseDt = DatePart("m", dUseDt) & "/" & DatePart("d", dUseDt) & "/" & DatePart("yyyy", dUseDt)
                                            Me.Controls(oCtl.Name) = CDate(sUseDt)
                                        End If
                                    Else
                                        Me.Controls(oCtl.Name) = mrsPayerDtl.Fields(oCtl.Name).Value
                                    End If
                                
                                Else
                                    
                                    If oCtl.Name = "DateSubmitted" Then
                                        If IsDate(Me.Controls(oCtl.Name)) = True Then
                                            dUseDt = Me.Controls(oCtl.Name)
                                            sUseDt = DatePart("m", dUseDt) & "/" & DatePart("d", dUseDt) & "/" & DatePart("yyyy", dUseDt)
                                            mrsPayerDtl.Fields(oCtl.Name).Value = CDate(sUseDt)
                                        End If
                                    Else
                                        mrsPayerDtl.Fields(oCtl.Name).Value = Me.Controls(oCtl.Name)
                                    End If

                                    bPayerChanges = True
                                End If
                            End If
                        
                        
                        End Select

                    End If
                
                End If
            End If
        Next
        
    End If
    
        ' Now, do we need to update the database or were we just setting the form values?
    If bSetFormValues = False Then
        mrsConceptHdr.Update
        
        Dim oAdo As clsADO
        Set oAdo = New clsADO
        oAdo.ConnectionString = GetConnectString("v_CODE_Database")
        bResult = oAdo.Update(mrsConceptHdr, "usp_CONCEPT_Hdr_Apply")
        If bPayerChanges = True Then
            mrsPayerDtl.Update
'Stop
            Set oAdo = New clsADO
            oAdo.ConnectionString = GetConnectString("v_CODE_Database")
            bResult = oAdo.Update(mrsPayerDtl, "usp_CONCEPT_PAYER_Dtl_Apply")

        End If
    Else
        bResult = True
    End If
        
    
    Call EnableDisableConvertWizard

    Call FlipNotesSubform

    GetOrSetControlValues = bResult
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GetOrSetControlValues = False
    GoTo Block_Exit ' or perhaps resume next if we want other controls to get their values?
End Function


Private Sub SaveData()
Dim bResult As Boolean
Dim tErrTxt As String
Dim ofrmProdNotes As Form_frm_CONCEPT_Production_Notes
Dim strProcName As String

    strProcName = ClassName & ".SaveData"
    On Error GoTo ErrHandler
    
    Set ofrmProdNotes = Me.sfrm_CONCEPT_Production_Notes.Form
    
'    If mbRecordChanged = False And Me.IsDirty = False Then
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
        
    '' Budget Group
    If Trim(Nz(Me.BudgetGroup)) = "" Then
        If Len(tErrTxt) > 0 Then
            tErrTxt = tErrTxt + ", "
        End If
        tErrTxt = tErrTxt + "Budget Group"
    End If
        
        
    'Are there any missing fields? Tell the user and exit without saving
    If Len(tErrTxt) > 0 Then
        MsgBox "Data must be entered for these field(s) before the issue can be saved; " + tErrTxt, vbOKOnly + vbExclamation
        GoTo Block_Exit
    End If
    
    ' Finally, we can't allow them to change the status to 380 (posted to website) if the LCD change flag
    '' is checked (and it hasn't been released yet)
Dim sMsg As String
    If StatusChangeValidationFailed(sMsg) = True Then
        Stop
        LogMessage strProcName, "VALIDATION ERROR", sMsg, , True, Me.FormConceptID

        ' we dont need to set it back, something else could have changed too..
        Stop
        
        GoTo Block_Exit
    End If

    
    'Inserting, set some recordset values
    mrsConceptHdr.MoveFirst
    
    If Me.Insert Then
        mrsConceptHdr.Fields("AccountID") = gintAccountID
        mrsConceptHdr.Fields("ConceptID") = Me.FormConceptID
    End If
    
    mrsConceptHdr.Fields("LastUpDt") = Now
    mrsConceptHdr.Fields("LastUpUser") = Identity.UserName

        
    bResult = GetOrSetControlValues(False)
    Call SetupPrevValsDict
    Call FilterPayerCombo
    
    If bResult Then
        MsgBox "Record Saved", vbOKOnly
        RaiseEvent ConceptSaved
        If IsSubForm(Me) = False Then
'            Me.IsDirty = False
            mbRecordChanged = False
            DoCmd.Close acForm, Me.Name
        Else
'            Me.IsDirty = False
            mbRecordChanged = False
        End If
       
    Else
        Err.Raise 65000, , "Error Saving Record"
    End If
    
    If ofrmProdNotes.IsDirty = True Then
        ofrmProdNotes.SaveData
    End If
    
Block_Exit:
    Set ofrmProdNotes = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_Concept_hdr : RefreshMain"
End Sub


Private Sub AdditionalInfo_Change()
    mbRecordChanged = True
End Sub

Private Sub AdditionalInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("AdditionalInfo")
End Sub

Private Sub Auditor_Change()
    mbRecordChanged = True
End Sub

Private Function LoadAdjOutcomeCombo() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim sList As String

    strProcName = ClassName & ".LoadAdjOutcomeCombo"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_DesiredOutcome"
        '.Parameters.Refresh
        Set oRs = .ExecuteRS
    End With
    
    While Not oRs.EOF
        sList = sList & oRs("DesiredAdjOutcome").Value & ";"
        oRs.MoveNext
    Wend
    
    If sList <> "" Then
        sList = left(sList, Len(sList) - 1)
    End If
    Me.DesiredAdjOutcome.RowSource = sList
    
    LoadAdjOutcomeCombo = True
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit

End Function

Private Sub cboConceptStatus_READONLY_Change()
    MsgBox "Concept status has been moved above the tabs. This is here for reference only - eventually to be removed!", vbOKOnly, "Use the other Drop down!"
End Sub

Private Sub cboConceptStatus_READONLY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "Concept status has been moved above the tabs. This is here for reference only - eventually to be removed!", vbOKOnly, "Use the other Drop down!"
End Sub

Private Sub cmbPayer_Change()
On Error GoTo Block_Err
Dim strProcName As String
Dim iNewPayerId As Integer
Dim oFrm As Form_frm_CONCEPT_Main
Dim sPayerName As String
Dim oFrmPg As Form_frm_CONCEPT_Production_Notes

    strProcName = ClassName & ".cmbPayer_Change"
    
    If Me.cmbPayer = 1000 Then  ' all
        lblPayer.FontBold = True
    Else
        lblPayer.FontBold = False
    End If
    
    Me.CurrentPageSelected = Me.TabCtl128
    If Me.PayerNameId <> cmbPayer.Value Then
        RaiseEvent PayerNameIdChanged(cmbPayer.Value)
    End If
    If DidAnythingChange() = True Then
        mbRecordChanged = True
    End If
    Me.PayerNameId = cmbPayer.Value
    
    
    
    ' Did anything change for this payer, if so, prompt to change..
    If mbRecordChanged = True Then
        iNewPayerId = Me.cmbPayer
        sPayerName = Me.cmbPayer.Text
        If MsgBox("It appears that you've edited something for this payer (" & sPayerName & "). Do you wish to save first? Please note that if you select NO, " & _
                " your changes will be lost", vbYesNo, "Save or Discard?") = vbYes Then
            Call SaveData
            RefreshData
            Me.cmbPayer = iNewPayerId
        Else
            ' loose changes! continue!
        End If
        mbRecordChanged = False
    End If
    
        '  Globally save the selected payer
    If IsSubForm(Me) = True Then
        Set oFrm = Me.Parent
        oFrm.SelectedPayerNameId = Me.cmbPayer.Value
    End If
    
    ' Ok, now we need to update our mrsPayerDtl recordset and reload those values
    Call GetPayerDtlRS
    
    Call GetOrSetControlValues(True)
    Call SetupPrevValsDict
    Call EnableAndRecolorFields
    Call LockFieldsIfPkgCreated
    Me.TabCtl128 = CurrentPageSelected
    
    If CurrentPageSelected = Me.pgProductionNotes.PageIndex Then
        
        Set oFrmPg = Me.sfrm_CONCEPT_Production_Notes.Form
        oFrmPg.RefreshData
        Set oFrmPg = Nothing
    End If
    
Block_Exit:
    Set oFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub

Private Sub cmdADRLetterSelect_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_CONCEPT_ADR_Letter_Selection

    strProcName = ClassName & ".cmdADRLetterSelect_Click"
    
'LogMessage strProcName, "NOT READY YET!", "Sorry, this feature isn't 100% yet - very soon though!!!", , True, Me.FormConceptID
'GoTo Block_Exit
'
    Set oFrm = New Form_frm_CONCEPT_ADR_Letter_Selection
    ' set it up:
    With oFrm
        .ConceptID = Me.FormConceptID
        .PayerNameId = Me.PayerNameId
        If Nz(Me.txtADRLetter, "") <> "" Then
            .SelectedId = Me.txtADRLetter
        End If
    End With
    
    Call KDShowFormAndWait(oFrm)
    
    If oFrm.Canceled = True Then
        GoTo Block_Exit
    End If
    
    If oFrm.SelectedId <> "" Then
        ' Save the thing!
'' kd comeback
        Call SaveSelectedAdrLetter(oFrm.SelectedId)
        
        
        ' then refresh our text box:
        Call Me.RefreshADRLetterCombo
    End If
    
Block_Exit:
    If Not oFrm Is Nothing Then
        DoCmd.Close acForm, oFrm.Name
    End If
    Set oFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdADRLetterView_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oWordApp As Word.Application
Dim oWordDoc As Word.Document
Dim sDocPath As String
Dim sTempLoc As String

    strProcName = ClassName & "cmdADRLetterView_Click"
    If Nz(Me.txtADRLetter, "") = "" Then
        LogMessage strProcName, "ERROR", "Could not determine ADR LetterType", , True, Me.FormConceptID
        GoTo Block_Exit
    End If
    
    
'    sSql = "SELECT LetterType, LetterDesc, TemplateLoc,( " & _
'    " SELECT TOP 1  " & _
'    " RefLink FROM AUDITCLM_References R  " & _
'    " WHERE R.RefType = 'LETTER' " & _
'    " AND r.RefSubType = LT.LetterType ORDER BY r.CreateDt DESC " & _
'    ") As SampleDocPath FROM Letter_Type  LT  " & _
'    " Where AccountID = 1 And ADR = 1 AND LT.LetterType = '" & Me.txtADRLetter & "'"
    

    sSql = "SELECT LetterType, LetterDesc, TemplateLoc As SampleDocPath FROM Letter_Type  LT  " & _
    " Where AccountID = 1 And ADR = 1 AND LT.LetterType = '" & Me.txtADRLetter & "'"
        
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_data_database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
    sDocPath = Nz(oRs("SampleDocPath").Value, "")
    
    If sDocPath = "" Then
        LogMessage strProcName, "ERROR", "Could not determine sample document path for this concept", , True, Me.FormConceptID
        GoTo Block_Exit
    End If
    

    sTempLoc = GetUniqueFilename(, , FileExtension(sDocPath))
    If CopyFile(sDocPath, sTempLoc, False) = False Then
        LogMessage strProcName, "ERROR", "Could not copy the sample document to your temp folder", sTempLoc, True, Me.FormConceptID
        GoTo Block_Exit
    End If
    
    Set oWordApp = New Word.Application
        ' open it read only!!!!
    Set oWordDoc = oWordApp.Documents.Open(sTempLoc, , True)
    oWordApp.visible = True
    
    oWordApp.Activate
    
    Call AppActivate(oWordApp.ActiveDocument.Name, False)
    Call SendKeys("^a", False)
    Sleep 500
    Call SendKeys("%{F9}", False)
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    ' No need to .Quit since we made it visible:
    oWordApp.Activate
    Set oWordDoc = Nothing
    Set oWordApp = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdConceptSQL_Click()
    If Me.ConceptID <> "" Then
        Call GetConceptStoredProcSQL(Me.ConceptID)
    End If
End Sub


Private Sub cmdConvertToPayerConcept_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim ofPayers As Form_frm_Prompt_For_Payers
Dim saryPayerNames() As String
Dim saryPayerNIDs() As String
Dim sPayrNIds As String
Dim iPayerIdx As Integer
Dim sThisPayerName As String
Dim lThisPayerId As Long
Dim oConceptHdrCopyRS As ADODB.RecordSet
Dim oPayerDtlRS As ADODB.RecordSet
Dim oCn As ADODB.Connection
Dim sErrMsg As String
Dim sClientIssueNumToUse As String


    strProcName = ClassName & ".cmdConvertToPayerConcept_Click"
    
    Call CreateConceptObjIfNeeded
    LogMessage strProcName, , "User clicked conversion wizard for concept '" & coCurConcept.ConceptID & "'", , , Me.FormConceptID
    
'Stop
    If PromptUserForPayers("Please select the payers for this concept (" & coCurConcept.ConceptID & ")", coCurConcept, , sPayrNIds, saryPayerNames) = False Then
        LogMessage strProcName, , "User canceled or there was an error when prompting for payers", , , Me.FormConceptID
        GoTo Block_Exit
    End If

    sPayrNIds = Replace(sPayrNIds, "1000,", "")
    saryPayerNIDs = Split(sPayrNIds, ",")

        '' Get Our Concept dataset
    Set oConceptHdrCopyRS = mrsConceptHdr.Clone
    Set oConceptHdrCopyRS.ActiveConnection = Nothing
    oConceptHdrCopyRS.MoveFirst ' should be 1 record anyway...
 
    Call CreateConceptObjIfNeeded
    
        '' open our connection
    Set oCn = New ADODB.Connection
    oCn.Open GetConnectString("v_DATA_DATABASE")

        ' create each of the payer dtl records and save to DB...
    For iPayerIdx = 0 To UBound(saryPayerNIDs)
        lThisPayerId = CLng(saryPayerNIDs(iPayerIdx))
        sThisPayerName = GetPayerNameFromID(lThisPayerId)

        If lThisPayerId <> 1000 Then

                ' Now set up our editable, detached recordset
            Set oPayerDtlRS = New ADODB.RecordSet
            With oPayerDtlRS
                .CursorLocation = adUseClientBatch
                .CursorType = adOpenKeyset
                .LockType = adLockBatchOptimistic
                .ActiveConnection = oCn
                .Open "SELECT * FROM CONCEPT_PAYER_Dtl WHERE ConceptId = '" & Me.ConceptID & "' AND PayerNameID = " & CStr(lThisPayerId)
                Set .ActiveConnection = Nothing
            End With
    
    
            '' Odd situation, if there's only 1 payer, AND there's already a Client Issue number then
            '' we want to use that for the 1 payer.
            
            If UBound(saryPayerNIDs) = 0 And Nz(Me.ClientIssueNum, "") <> "" Then
                sClientIssueNumToUse = Me.ClientIssueNum
            End If
            
            '' Save the
            Call InsertDetailForPayer(Me, oPayerDtlRS, lThisPayerId, oConceptHdrCopyRS, False, sClientIssueNumToUse)
            
            If sClientIssueNumToUse = "" Then
                If mod_Concept_Specific.IssueClientIssueNum(coCurConcept, lThisPayerId, sErrMsg) = "" Then
                    LogMessage strProcName, "ERROR", "There was a problem generating the client issue number for this payer: " & sThisPayerName, sThisPayerName, True, Me.FormConceptID
                End If
            End If
        End If  ' end skipping all
    Next
    
        '' Now, we need to null out the Claim Hdr details..
    Call NullOutConceptHdrDueToPayerDtl(Me, oConceptHdrCopyRS)
    
    ' loop through the attached documents and prompt for the payers where the doc type is a payer specific type
    If PromptForPayerOnConceptReferences(coCurConcept, sPayrNIds) = False Then
        LogMessage strProcName, "ERROR", "There was a problem while converting the attached documents to Payer specific", sPayrNIds, , Me.FormConceptID
        GoTo Block_Exit
    End If
    
    ' loop through states maybe? Codes, what else?
    If ConceptConvertConceptDtlCodes(coCurConcept, saryPayerNIDs) = False Then
        LogMessage strProcName, "ERROR", "There was a problem while converting the Codes to payer specific", , , Me.FormConceptID
        GoTo Block_Exit
    End If

    If ConceptConvertConceptStates(coCurConcept, saryPayerNIDs) = False Then
        LogMessage strProcName, "ERROR", "There was a problem while converting the attached documents to Payer specific", , , Me.FormConceptID
    End If
    
    
    ' Tracking.. Not sure how to really do this.
    ' if there's only 1 payer then no big deal
    ' if > 1 then duplicate the entries?
    
    If ConceptConvertTracking(coCurConcept, saryPayerNIDs) = False Then
        LogMessage strProcName, "ERROR", "There was a problem while converting the tracking notes to payer specific", , , Me.FormConceptID
    End If
    
   
    If ConceptConvertTaggedClaims(coCurConcept) = False Then
        LogMessage strProcName, "ERROR", "There was a problem while converting the tracking notes to payer specific", , , Me.FormConceptID
    End If
        
    
    ' done ??
    Call Me.RefreshData
    MsgBox "Complete!"
    Stop
Block_Exit:

    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub




Private Sub cmdCopyConcept_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sNewConceptId As String
Dim oMainFrm As Form_frm_CONCEPT_Main

    strProcName = ClassName & ".cmdCopyConcept_Click"
    
'    MsgBox "This has been disabled for now due to the massive Concept Submission changes mandated by CMS", vbInformation, "Not enabled for now!"
'    GoTo Block_Exit
    
    If MsgBox("Are you sure you wish to create a duplicate of this concept?", vbYesNo, "Duplicate concept?") = vbNo Then
        GoTo Block_Exit
    End If
    
    sNewConceptId = NewConceptBasedOnExisting(Me.ConceptID)
    If sNewConceptId <> "" Then
        ' now we need to search for it and reload the whole thign
        If IsSubForm(Me) Then
            Set oMainFrm = Forms("frm_CONCEPT_Main")
            oMainFrm.txtSearchBox = sNewConceptId
            oMainFrm.ckExpandSearch = False
            oMainFrm.ckIncludeCodes = False
            
            Call oMainFrm.RefreshData
            
            Set oMainFrm = Nothing
            
        End If
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.ConceptID
    GoTo Block_Exit
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


Private Sub cmdCreateNIRF_For_Submit_Click()
On Error GoTo Block_Err
Dim tErrTxt As String
Dim oConcept As clsConcept
Dim iPayerNameId As Integer
Dim sPromptMsg As String
Dim sPayerName As String
Dim strProcName As String

Dim coPayers As Collection
Dim vThisPayerId As Variant
Dim saryPayers() As String
Dim iPayerIdx As Integer
Dim bOpenExplorer As Boolean

    strProcName = ClassName & ".cmdCreateNIRF_For_Submit_Click"
    Call StartMethod
    ' 20120801 KD: Skipping the validation for this stuff because it's not just at the
    ' concept level, we have to check each payer and all kinds of other stuff..
    
    

    iPayerNameId = Nz(Me.cmbPayer, 0)
    If iPayerNameId = 0 Then
        LogMessage strProcName, "ERROR", "No payer detected!!!! Please use the 'Run NIRF Report button for older concepts!", , True, Me.FormConceptID
        GoTo Block_Exit
    End If

    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.FormConceptID) = False Then
        ' hmm.. weird
        LogMessage ClassName & ".cmdRunReport", "ERROR", "Could not load the concept object!?!?!", Me.FormConceptID, , Me.FormConceptID
        DoCmd.OpenReport "rpt_CONCEPT_New_Issue", acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
        GoTo Block_Exit
    Else
 
        sPayerName = GetPayerNameFromID(Me.cmbPayer)
        
        If Me.IsPayerSetToAll = True Then
            If MsgBox("Please note that because you have the Payer set to 'Concept Header Fields', this will create the NIRF for each related payer", vbOKCancel, "Create NIRF for ALL payers?") = vbCancel Then
                GoTo Block_Exit
            End If
            Set coPayers = mod_Concept_Specific.GetRelatedPayerNameIDs(Me.FormConceptID)
        Else
            If MsgBox("Are you sure you want to create the NIRF for " & sPayerName & "?", vbOKCancel, "Create NIRF for " & sPayerName) = vbCancel Then
                GoTo Block_Exit
            End If
            Set coPayers = New Collection
            coPayers.Add Me.cmbPayer.Value
        End If

        
        bOpenExplorer = False
        
        For Each vThisPayerId In coPayers
            
            If CreatePackageNirf(Me.FormConceptID, CInt(vThisPayerId), bOpenExplorer, True) = False Then
                LogMessage TypeName(Me) & ".cmdRunReport", "ERROR", "There was an error converting to PDF, opening as an MS Access report - please print to PDF", , True, Me.FormConceptID
                DoCmd.OpenReport "rpt_CONCEPT_New_Issue", acViewPreview, , "ConceptID = '" & Me.FormConceptID & "' AND PayerNameID = " & CStr(vThisPayerId)
            Else
                LogMessage strProcName, , "Nirf creation successful", "PayeriD: " & CLng(vThisPayerId), , Me.FormConceptID
            End If
            bOpenExplorer = False
        Next
            
    End If
    
    
Block_Exit:
    MsgBox "Finished", vbOKOnly, "Complete"
    Call FinishMethod
'    LogMessage strProcName, , "Nirf creation successful", "PayeriD: " & CLng(vThisPayerId)
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub

Private Sub cmdCreateNIRF_n_ClaimDoc_Click()
Dim sMsg As String

    sMsg = "This will populate the submit date and create the NIRF using that date!" & vbCrLf & "Are you sure you wish to proceed?"
    
    If MsgBox(sMsg, vbYesNo, "This will set the Submit date!") = vbNo Then
        Exit Sub
    End If
    
    LogMessage ClassName & ".cmdCreateNirf_n_ClaimDoc_Click", , "1. Create Nirf and ClaimDoc main sub", , , Me.FormConceptID
    cmdCreateNIRF_For_Submit_Click
    LogMessage ClassName & ".cmdCreateNirf_n_ClaimDoc_Click", , "2. Create Nirf and ClaimDoc main sub", , , Me.FormConceptID
    cmdCreateSampleClaimDocs_Click
    LogMessage ClassName & ".cmdCreateNirf_n_ClaimDoc_Click", , "3. Create Nirf and ClaimDoc main sub", , , Me.FormConceptID
End Sub

Private Sub cmdCreateSampleClaimDocs_Click()
    Call mod_Concept_Specific.CreateSampleClaimsDoc(coCurConcept, 1000)
    MsgBox "Finished!"
End Sub

Private Sub cmdEditNirf_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sFilter As String

'Dim oFrm As Form_frm_CONCEPT_NIRF_EDITOR_New

    strProcName = ClassName & ".cmdEditNirf_Click"
    
    If Me.IsConceptPayerSpecific = True Then
        If Me.IsPayerSetToAll = True Then
            LogMessage strProcName, "USER NOTIFICATION", "This concept is Payer Specific, Please select the payer that you wish to edit the NIRF for", , True, Me.FormConceptID
            GoTo Block_Exit
        End If
    End If
    
    
'    Set oFrm = New Form_frm_CONCEPT_NIRF_EDITOR_New
'    oFrm.OpenArgs = ""
    If Me.FormConceptID <> "" Then
        sFilter = "ConceptId = '" & Me.FormConceptID & "' "
        If Nz(Me.PayerNameId, 1000) > 1000 Then
            sFilter = sFilter & " AND PayerNameId = " & CStr(Me.PayerNameId)
        End If
        
        DoCmd.OpenForm "frm_CONCEPT_NIRF_EDITOR_New", acNormal, , , , , sFilter
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdFinalizePkg_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim colPayers As Collection
Dim sToAddress As String
Dim vPayerNameID As Variant
Dim sThisUser As String
Dim sPayerList As String
Dim oPayer As clsConceptPayerDtl
Dim iPayers As Integer

    strProcName = ClassName & ".cmdFinalizePkg_Click"
    
  
    
    '' This needs to:
    '' - Check that the package has been created


    
    '' Then it can send IT an email
    
    sToAddress = "Kevin.Dearing@connolly.com;Tuan.Khong@connolly.com;"


    '' Zip the stuff up...
    If Me.cmbPayer = 1000 Or Me.cmbPayer = 0 Then
        ' All payers:
        '' How do I get all payers for this concept again? I know I have a function somewhere..
        ' maybe there's a property..  I guess it should really be a property of
        ' the concept class
        Set colPayers = GetRelatedPayerNameIDs(Me.FormConceptID)
        
        For Each vPayerNameID In colPayers
            ' Need to check that the concept's status is good to go..
            Me.PayerNameId = CLng(vPayerNameID)
            
            Set oPayer = New clsConceptPayerDtl
            If oPayer.LoadFromConceptNPayer(coCurConcept.ConceptID, Me.PayerNameId) = False Then
                LogMessage strProcName, "ERROR", "Could not load payer object", coCurConcept.ConceptID & " Payer: " & CStr(Me.PayerNameId), , Me.FormConceptID
                GoTo NextOne
            End If
            
            If oPayer.EffectiveDate > Now() Or oPayer.EndDate < Now() Then
                LogMessage strProcName, "WARNING", "This payer '" & oPayer.PayerName & "' is not valid at this time!", Format(oPayer.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayer.EndDate, "mm/dd/yyyy"), True, Me.FormConceptID
            Else
                sPayerList = sPayerList & GetPayerNameFromID(Me.PayerNameId) & ", "
                iPayers = iPayers + 1
                Call MarkConceptAsSubmitted(coCurConcept.ConceptID, Me.PayerNameId)
            End If
NextOne:                                    '' log the details to the DB:
        Next
        
        sPayerList = left(sPayerList, Len(sPayerList) - 2)  ' remove final comma + space
    Else
    
        ' 1 at a time:
        Me.PayerNameId = CLng(vPayerNameID)
        
        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(coCurConcept.ConceptID, Me.PayerNameId) = False Then
            LogMessage strProcName, "ERROR", "Could not load concept object for " & coCurConcept.ConceptID & " and payer: " & CStr(Me.PayerNameId), , , Me.FormConceptID
            GoTo Block_Exit
        End If
        
        If oPayer.EffectiveDate > Now() Or oPayer.EndDate < Now() Then
            LogMessage strProcName, "WARNING", "This payer '" & oPayer.PayerName & "' is not valid at this time!", Format(oPayer.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayer.EndDate, "mm/dd/yyyy"), True, Me.FormConceptID
            GoTo Block_Exit
        Else
            sPayerList = GetPayerNameFromID(Me.PayerNameId)
            iPayers = 1
            Call MarkConceptAsSubmitted(coCurConcept.ConceptID, Me.PayerNameId)
        End If
        
    End If
    
    If iPayers < 1 Then
        LogMessage strProcName, "ERROR", "This payer is not valid to submit a concept to at this time!", , True, Me.FormConceptID
    Else
        
        sMsg = Me.Auditor & " has clicked Finalized Concept '" & Me.FormConceptID & "' and payer: " & sPayerList & "." & vbCrLf & vbCrLf & "Please submit it via NDM or via the prescribed method for the payer(s) involved!"
        
    
        sThisUser = Identity.UserName()
        sThisUser = sThisUser & "@connolly.com"
    
            ' how about sending the email
        SendsqlMail "[CONCEPT MGMT] Concept Submission: " & Me.FormConceptID & " : " & sPayerList, sToAddress & Me.Auditor & "@connolly.com;", "Kenneth.Turturro@connolly.com;" & sThisUser, "", sMsg
'        SendsqlMail "[CONCEPT MGMT] Concept Submission: " & Me.FormConceptID & " : " & sPayerList, sToAddress, "", "", sMsg
    
        LogMessage strProcName, "CONFIRMATION", "The concept has just been submitted for payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID
    End If
    
Block_Exit:
    Set oPayer = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub

Private Sub cmdMarkAsSent_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sToAddress As String
Dim sMsg As String
Dim sThisUser As String

    strProcName = ClassName & ".cmdMarkAsSent_Click"
    
    ' first, make sure a single payer is selected (unless there IS only 1 payer for this concept..
    If PayerNameId = 1000 Then
        If coCurConcept.ConceptPayers.Count > 1 Then
            MsgBox "Please select the payer to mark as submitted!", vbInformation, "Select payer first!"
            GoTo Block_Exit
        Else
            ' Get the actual payernameid
            PayerNameId = coCurConcept.ConceptPayers.Item(1).PayerNameId
        End If
    End If
    

    sToAddress = "Kevin.Dearing@connolly.com;Tuan.Khong@connolly.com;"
    If Right(sToAddress, 1) <> ";" Then sToAddress = sToAddress & ";"
    
    ' now just do it:
    Call StartMethod
    
    Call IT_Mark_Concept_as_Sent_via_NDM
    
'    Call PrepConceptSubmitEmail(coCurConcept, PayerNameId)
    
    sThisUser = Identity.UserName()
    
    
    '' Ok now we need to generate the email to CMS from a template:
    
'
'    sMsg = Identity.UserName() & " has sent the concept '" & Me.FormConceptID & "' to payer: " & GetPayerNameFromID(PayerNameId) & "." & vbCrLf & vbCrLf
'
'
'     sThisUser = sThisUser & "@connolly.com"
'
'         ' how about sending the email
'     SendsqlMail "[CONCEPT MGMT] Concept Was Sent: " & Me.FormConceptID & " : " & GetPayerNameFromID(PayerNameId), sToAddress & Me.Auditor & "@connolly.com;", "Kenneth.Turturro@connolly.com;" & sThisUser, "", sMsg
'
    
     LogMessage strProcName, "CONFIRMATION", "The concept has just been marked as sent to the payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID
    
    
Block_Exit:
    Call FinishMethod
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub

Private Sub cmdQA_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim strConceptFolder As String
Dim oFso As Scripting.FileSystemObject
Dim sThisPayerName As String
Dim oFrmChecklist As Form_frm_GENERAL_Checklist
Dim oRs As ADODB.RecordSet
Dim oReqDoc As clsConceptReqDocType
Dim oReqObj As clsEracRequirementRule
Dim oPayer As clsConceptPayerDtl
Dim oAdo As clsADO
Dim oRsAddtnlReq As ADODB.RecordSet

    strProcName = ClassName & ".cmdQA_Click"
    
  
    If Me.PayerNameId = 1000 Then
        LogMessage strProcName, "USER MESSAGE", "In this version, please check each payer specifically (you have it set to All payers right now..)'", , True, Me.FormConceptID
        
        GoTo Block_Exit
    End If
     
    Set oPayer = New clsConceptPayerDtl
    If oPayer.LoadFromConceptNPayer(Me.FormConceptID, Me.PayerNameId) = False Then
        Stop
    End If
    
    If oPayer.EffectiveDate > Now() Or oPayer.EndDate < Now() Then
        If oPayer.PayerStatusNum = "990" Then    ' Void
            LogMessage strProcName, "WARNING", "This payer has a VOID status!", oPayer.PayerName, , Me.FormConceptID
            GoTo Block_Exit
        Else
            LogMessage strProcName, "ERROR", "This payer is not valid at the current time! IT should probably be voided out!", Format(oPayer.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayer.EndDate, "mm/dd/yyyy"), True, Me.FormConceptID
            GoTo Block_Exit
        End If
    End If
    
    
    
    '' First, double check that the package has been created.
    If Me.FormConceptID = "" And Me.ConceptID <> "" Then
        Me.FormConceptID = Me.ConceptID
    End If
    If mod_Concept_Specific.WasPackageCreated(Me.FormConceptID, Me.PayerNameId) = False Then
        LogMessage strProcName, "WARNING", "The package was not yet created!", Me.FormConceptID & " : " & GetPayerNameFromID(Me.PayerNameId), True, Me.FormConceptID
        GoTo Block_Exit
    End If
   
    
    sThisPayerName = GetPayerNameFromID(Me.PayerNameId)
    
    MarkConceptAsQAd Me.FormConceptID, Me.PayerNameId
    
    strConceptFolder = csCONCEPT_SUBMISSION_WORK_FLDR & Me.FormConceptID & "\" ' ANd the payername if not 1000
    
    If Me.PayerNameId <> 1000 Then
        strConceptFolder = strConceptFolder & sThisPayerName & "\"
    End If
    
 
        '' Now, gather our details
    Set oReqObj = coCurConcept.RequirementRuleObj
    Set oRs = New ADODB.RecordSet
    
    With oRs
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        Set .ActiveConnection = Nothing

        .Fields.Append "PayerName", adLongVarWChar, 1
        .Fields.Append "RequiredDocType", adLongVarWChar, 1
        .Fields.Append "Notes", adLongVarWChar, 1
        .Open
    End With
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM ADMIN_ConceptMgmt_Config WHERE Name = 'ValidationItem' AND Active <> 0 ORDER BY [Value]"
        Set oRsAddtnlReq = .ExecuteRS
        If .GotData = True Then
            While Not oRsAddtnlReq.EOF
                oRs.AddNew
                oRs("PayerName") = sThisPayerName
                oRs("RequiredDocType") = ""
                oRs("Notes") = oRsAddtnlReq("Value").Value
                oRs.Update
                oRsAddtnlReq.MoveNext
            Wend
        End If
    End With
    
    If Not oReqObj Is Nothing Then
        For Each oReqDoc In oReqObj.RequiredDocs
        
            oRs.AddNew
            oRs("PayerName") = sThisPayerName
            oRs("RequiredDocType") = oReqDoc.DocName
            If oReqDoc.NumPerPayer > 1 Then
                oRs("Notes") = "There should be " & CStr(oReqDoc.NumPerPayer) & " of these for each payer"
            End If
            oRs.Update
            
        Next
    End If
        ' show the checklist..
    If IsOpen("frm_GENERAL_Checklist") = True Then
        DoCmd.Close acForm, "frm_GENERAL_Checklist", acSaveNo
    End If
    

    DoCmd.OpenForm "frm_GENERAL_Checklist", acNormal
    Sleep 500
    
        '' This is going to open a list of documents that are required
    Set oFrmChecklist = New Form_frm_GENERAL_Checklist

    Set oFrmChecklist = Application.Forms("frm_GENERAL_Checklist")
    oFrmChecklist.ListRecordset = oRs
    oFrmChecklist.Caption = "Required documents:"
    
    
    
        '' then open windows explorer to the erac submit folder (where the package was created)
    Set oFso = New Scripting.FileSystemObject
    
        ' Exit if folder doesn't exist
    Dim iTimesTried As Integer
    
TryParentFldr:
    If Not oFso.FolderExists(strConceptFolder) Then
        LogMessage strProcName, "WARNING", "Folder does not exist, was the package actually created?", strConceptFolder, True, Me.FormConceptID
        strConceptFolder = ParentFolderPath(strConceptFolder)
        If iTimesTried < 2 Then
            iTimesTried = iTimesTried + 1
            GoTo TryParentFldr
        Else
            GoTo Block_Exit
        End If

    Else
        Shell "explorer " & strConceptFolder, vbNormalFocus
    End If
        
    
    
Block_Exit:
    If Not oRsAddtnlReq Is Nothing Then
        If oRsAddtnlReq.State = adStateOpen Then oRsAddtnlReq.Close
        Set oRsAddtnlReq = Nothing
    End If
    Set oAdo = Nothing
    Set oPayer = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oFso = Nothing
    Set oFrmChecklist = Nothing
    Set oReqDoc = Nothing
    Set oReqObj = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub

Private Sub cmdRefresh_Click()

    ' undo sort of..
    If IsSubForm(Me) = True Then
        mbRecordChanged = False
        Call Me.Parent.RefreshData
    Else
        mbRecordChanged = False
        Me.RefreshData
    End If
    

End Sub

Private Sub cmdRegenEmail_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim sMsg As String

'    MsgBox "This should not be used anymore - I don't think anyway!"
'    GoTo Block_Exit

    strProcName = ClassName & ".cmdRegenEmail_Click"
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.ConceptID) = False Then
        LogMessage strProcName, "ERROR", "There was a problem creating the concept object. Please close Concept Management form, reopen and try again. If you get this message again, please contact support!", Me.ConceptID, True, Me.ConceptID
        GoTo Block_Exit
    End If
    
   
    
    ' make sure it's considered submitted already..
    If oConcept.AlreadySubmitted(clngPayerNameId, sMsg) < CDate("1/2/1900") Then
        LogMessage strProcName, "ERROR", "This concept has not been submitted yet.. Please click the 'Submit Package' button", sMsg & " " & Me.ConceptID, True, Me.ConceptID
        GoTo Block_Exit
    End If
    
    If PrepConceptSubmitEmail(oConcept, clngPayerNameId, cstrResubmitPkgFldr, sMsg) = False Then
        LogMessage strProcName, "ERROR", "There was a problem recreating the email!", sMsg & " " & oConcept.ConceptID, True, oConcept.ConceptID
        GoTo Block_Exit
    End If

Block_Exit:
    Set oConcept = Nothing
    DoCmd.Hourglass False
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , oConcept.ConceptID
    Err.Clear
    GoTo Block_Exit
End Sub



Private Sub cmdReZip_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim vPayerNameID As Variant


    strProcName = ClassName & ".cmdReZip_Click"
    DoCmd.Hourglass True
    
    Call GetPayerCollection
    
    For Each vPayerNameID In coPayers
        If ZipConceptSubmitPackage(coCurConcept, CLng(vPayerNameID)) = "" Then
Stop
        End If
    Next
    
    
    MsgBox "Finished 'Rezipping'", vbOKOnly
    
Block_Exit:
    DoCmd.Hourglass False

    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub




'''Private Sub cmdRequestClientIssueId_Click()
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim oConcept As clsConcept
'''Dim iSelectedPayerID As Integer
'''
'''    strProcName = ClassName & ".cmdRequestClientIssueId_Click"
'''
'''    Set oConcept = New clsConcept
'''    If oConcept.LoadFromID(Me.ConceptID) = False Then
'''        LogMessage strProcName, "ERROR", "There was a problem creating the concept object. Please close Concept Management form, reopen and try again. If you get this message again, please contact support!", Me.ConceptID, True
'''        GoTo Block_Exit
'''    End If
'''
'''    iSelectedPayerID = Me.cmbPayer
'''
'''    If iSelectedPayerID = 0 Or iSelectedPayerID = 1000 Then
'''        LogMessage strProcName, , "Please make sure you select the payer from the dropdown and try again"
'''        GoTo Block_Exit
'''    End If
'''
'''    If IssueClientIssueNum(oConcept) = "" Then
'''        LogMessage strProcName, "ERROR", "There was a problem generating the Client Issue ID", oConcept.ConceptID, True
'''        GoTo Block_Exit
'''    End If
'''
'''    Call RefreshData
'''
'''Block_Exit:
'''    Set oConcept = Nothing
'''    DoCmd.Hourglass False
'''
'''    Exit Sub
'''Block_Err:
'''    ReportError Err, strProcName
'''    Err.Clear
'''    GoTo Block_Exit
'''End Sub


Private Sub cmdRunReport_Click()
Dim tErrTxt As String
Dim oConcept As clsConcept
Dim iPayerNameId As Integer
Dim sPromptMsg As String
Const sPayerReportName As String = "rpt_CONCEPT_New_Issue_2014"
Const sCmsOnlyReportName As String = "rpt_CONCEPT_New_Issue_CMS_Only_2014"


    iPayerNameId = Nz(Me.cmbPayer, 0)
    If iPayerNameId = 0 Then
        Stop ' hammer time! Problem - all would  be a nice default don'tcha think?
    End If

    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.FormConceptID) = False Then
        ' hmm.. weird
Stop
        LogMessage ClassName & ".cmdRunReport", "ERROR", "Could not load the concept object!?!?!", Me.FormConceptID, , Me.FormConceptID
        DoCmd.OpenReport sPayerReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
    Else
    

        If Me.IsPayerSetToAll = True Then
            If Me.IsConceptPayerSpecific = True Then
                DoCmd.OpenReport sPayerReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
            Else
'Stop
                DoCmd.OpenReport sCmsOnlyReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
                
            End If
        Else
            If Me.IsConceptPayerSpecific = False Then
                DoCmd.OpenReport sCmsOnlyReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "'"
            Else
'Stop
' 2/11/2014                DoCmd.OpenReport sCmsOnlyReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "' "
                DoCmd.OpenReport sPayerReportName, acViewPreview, , "ConceptID = '" & Me.FormConceptID & "' AND PayerNameId = " & Me.PayerNameId
            End If
            
        End If

    
    End If
    
    
    
    
End Sub

Private Sub cmdSave_Click()
Dim iPayerSelection As Integer

'    iPayerSelection = cmbPayer.ListIndex
'    iPayerSelection = cmbPayer.ItemData(Abs(cmbPayer.ColumnHeads) + 1)
    iPayerSelection = cmbPayer.Value
    
    SaveData
    cmbPayer.SetFocus
'    cmbPayer.ListIndex = iPayerSelection
    cmbPayer.Value = iPayerSelection
    cmbPayer_Change
End Sub


Private Sub cmdSubmit_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim bReady As Boolean
Dim oConcept As clsConcept
Dim oRs As ADODB.RecordSet
Dim oMainFrm As Form_frm_CONCEPT_Main
Dim iPayerNameId As Integer

    strProcName = ClassName & ".cmdSubmit_Click"
    
    If Me.cmdSubmit.Caption = csReSubmitBtnText Then
        Stop
        Call ResubmitPackage
        GoTo Block_Exit
    End If
    
        '' Submit (not Re-Submit)
    Call CreateSubmitPackage
    GoTo Block_Exit


    
Block_Exit:
    Set oMainFrm = Forms("frm_CONCEPT_Main")
    oMainFrm.RefreshData
    MsgBox "Finished!"
    Set oMainFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub


Private Function ResubmitPackage() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim colPayers As Collection
Dim sToAddress As String
Dim vPayerNameID As Variant
Dim sThisUser As String
    
    strProcName = ClassName & ".ResubmitPackage"
    
    sToAddress = "Kevin.Dearing@connolly.com;"


    '' Zip the stuff up...
    If Me.cmbPayer = 1000 Or Me.cmbPayer = 0 Then
        ' All payers:
        '' How do I get all payers for this concept again? I know I have a function somewhere..
        ' maybe there's a property..  I guess it should really be a property of
        ' the concept class
        Set colPayers = GetRelatedPayerNameIDs(Me.FormConceptID)
        
        For Each vPayerNameID In colPayers
            ' Need to check that the concept's status is good to go..
            Me.PayerNameId = CLng(vPayerNameID)
            Call ZipConceptSubmitPackage(coCurConcept, CLng(vPayerNameID), True)
                '' log the details to the DB:
            Call MarkConceptAsSubmitted(Me.FormConceptID, CLng(vPayerNameID), True)

        Next
        
    Else
        ' 1 at a time:
        Me.PayerNameId = CLng(vPayerNameID)
        Call ZipConceptSubmitPackage(coCurConcept, Me.cmbPayer, True)
            '' log the details to the DB:
        Call MarkConceptAsSubmitted(Me.FormConceptID, CLng(Me.cmbPayer), True)

    End If
    
    sMsg = Me.Auditor & " has clicked submit for concept '" & Me.FormConceptID & "' and payer: " & GetPayerNameFromID(Me.cmbPayer) & "." & vbCrLf & vbCrLf & "Please submit it via NDM or by the appropriate means for this payer!"
    

    sThisUser = Identity.UserName()
    sThisUser = sThisUser & "@connolly.com"

    ' This is until I can do the real deal..
    ' Not going to validate here, we're just going to send dataservices an email and
    ' copy ken
'Stop
        ' how about sending the email
    SendsqlMail "[CONCEPT MGMT] Concept Re-Submission: " & Me.FormConceptID & " : " & GetPayerNameFromID(Me.PayerNameId), sToAddress & Me.Auditor & "@connolly.com;", "Kenneth.Turturro@connolly.com;" & sThisUser, "", sMsg

    LogMessage strProcName, "CONFIRMATION", "The concept has just been Re-submitted for payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Function

'
'Private Function ResubmitPackage() As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim bReady As Boolean
'Dim oConcept As clsConcept
'Dim oRs As ADODB.Recordset
'Dim oMainFrm As Form_frm_CONCEPT_Main
'Dim iPayerNameID As Integer
'
'    strProcName = ClassName & ".ResubmitPackage"
'
'
'
'    Set oConcept = New clsConcept
'    If oConcept.LoadFromID(Me.ConceptID) = False Then
'        ' not ready - something went wrong
'        LogMessage strProcName, "ERROR", "Could not load the concept object! Not ready to submit", Me.ConceptID
'        cmdSubmit.Enabled = False
'        GoTo Block_Exit
'    End If
'
'
'    If oConcept.ClientIssueId(Me.PayerNameId) = "" Then
'         Call mod_Concept_Specific.IssueClientIssueNum(oConcept)
'    End If
'
'    ' for now, we aren't going to re-validate the concept when we are re-submitting
'    ' let's just check that it was indeed submitted to CMS first
'    If oConcept.AlreadySubmitted(Me.PayerNameId) <= CDate("1/1/1900") Then
'        LogMessage strProcName, "ERROR", "This concept was not submitted yet so we cannot RE-submit it. It will have to be submitted as a payer specific concept", oConcept.ConceptID, True
'        GoTo Block_Exit
'    End If
'
''    Set oRs = New ADODB.Recordset
'''    bReady = oConcept.ValidateForSubmission(oRs, Me.cmbPayer)
'''
'''    If bReady = False Then
'''        ' not ready - something went wrong
'''        LogMessage strProcName, "ERROR", "Concept failed validation!", Me.ConceptID, True
'''
'''        LogActionToHistory Me.ConceptID, ValidatedConcept, "Failed", , , Me.ConceptID
'''
'''        cmdSubmit.Enabled = False
'''        cmdRegenEmail.Enabled = False
'''            ' Show the user the problem.. :D
'''        Call ValidateForSubmission(Me.cmbPayer)
'''        GoTo Block_Exit
'''    End If
'
''    LogActionToHistory oConcept.ConceptID, 0, ValidatedConcept, "Success", , , Me.ConceptID
'
''    iPayerNameID = Nz(Me.cmbPayer, 0)
''    If iPayerNameID = 0 Then
''Stop ' problem!!!
''    End If
'
'Stop
'    ' Create a new NIRF
'    If oConcept.NIRF_Exists = False Then
'        LogMessage strProcName, "ERROR", "No initial NIRF was found! Was this indeed submitted already?", oConcept.ConceptID, True
'        GoTo Block_Exit
'    Else
'        If MsgBox("Do you want to create a new NIRF (Yes) or use the existing one (No)?", vbYesNo, "Create a new NIRF?") = vbYes Then
'            If CreatePackageNirf(oConcept.ConceptID, 1000, False, False, , False) = False Then
'                LogMessage strProcName, "ERROR", "There was a problem creating the NIRF for concept: " & oConcept.ConceptID, , True
'                LogActionToHistory oConcept.ConceptID, NirfCreated, "Failure creating nirf", , , oConcept.ClientIssueId(Me.PayerNameId)
'                GoTo Block_Exit
'            End If
'        End If
'    End If
'
'    ' Create a Resubmit package folder
'Stop
'    Call CreateResubmitFolder_n_Package(oConcept, cstrResubmitPkgFldr)
'
'    ' Ok, we're good to go..
'Stop
'    cmdRegenEmail_Click
'
''    Call SubmitConcept(oConcept, iPayerNameID, , bReady)
'
'    '' Finally, refresh the data
'    ResubmitPackage = True
'    Set oMainFrm = Forms("frm_CONCEPT_Main")
'    oMainFrm.RefreshData
'Block_Exit:
'    Set oMainFrm = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Function

Private Sub cmdTest_Click()
Stop
End Sub

Private Sub cmdValidate_Click()
    ' Are we submitting all payers or just 1?
    Call ValidateForSubmission(Me.cmbPayer)
    
End Sub


Private Sub cmdViewPackage_Click()
    Dim strConceptFolder As String
    Dim fso As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    
        ''    strConceptFolder = "\\cca-audit\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\" & Me.FormConceptID & "\"
    strConceptFolder = csCONCEPT_SUBMISSION_SAVE_FLDR & Me.FormConceptID & "\"
    
    ' Exit if folder doesn't exist
    If Not fso.FolderExists(strConceptFolder) Then
        MsgBox "Folder does not exist", vbOKOnly
        Exit Sub
    Else
        Shell "explorer " & strConceptFolder, vbNormalFocus
    End If
    

        
End Sub






Private Sub Comments_Change()
    mbRecordChanged = True
End Sub

Private Sub Comments_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("Comments")
End Sub

Private Sub ConceptCatId_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptCatId_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frm_Concept_Categories", acNormal, , , acFormEdit, acDialog
    Me.ConceptCatId.Requery
End Sub

Private Sub ConceptClaimCount_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptClaimCount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptClaimCount")
End Sub

Private Sub ConceptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptDesc")
    mbRecordChanged = True
End Sub

Private Sub ConceptIndicator_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptLevel_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptLogic_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptLogic")
    mbRecordChanged = True
End Sub

Private Sub ConceptPotentialValue_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptPotentialValue_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptPotentialValue")
End Sub

Private Sub ConceptPriority_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptRationale_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptRationale_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptRationale")
End Sub

Private Sub ConceptReferences_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptReferences_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("ConceptReferences")
End Sub

Private Sub ConceptSource_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptSQL_Change()
    mbRecordChanged = True
End Sub

Private Sub ConceptStatus_Change()
Dim bPrevState As Boolean

    bPrevState = mbRecordChanged
    
    '' Check a few things to make sure they are allowed to do this
    
    If AllowedToChangeConceptStatus = False Then
        Call RollbackStatusToLoadedValue
        mbRecordChanged = bPrevState    ' set it back to whatever it was before this happened.. (not dirty Just because of this since we rolled it back)
    Else
        mbRecordChanged = True
    End If
    
End Sub

Private Sub CopyOfReference_Change()
    mbRecordChanged = True
End Sub

Private Sub CopyOfReference_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("CopyOfReference")
End Sub

Private Sub CreateDate_Change()
    mbRecordChanged = True
End Sub

Private Sub CreateDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("CreateDate")
End Sub

Private Sub DataType_Change()
    mbRecordChanged = True
End Sub

Private Sub DateSubmitted_Change()
    mbRecordChanged = True
End Sub

Private Sub DateSubmitted_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("DateSubmitted")
End Sub

Private Sub DesiredAdjOutcome_Change()
    mbRecordChanged = True
End Sub

Private Sub DesiredAdjOutcome_DblClick(Cancel As Integer)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".DesiredAdjOutcome_DblClick"
    
    ' Going to open a modal form to add one, then reload the combo box..
    DoCmd.OpenForm "frm_CONCEPT_Xref_Add_DesiredOutcome", acNormal, , , acFormAdd, acDialog
    Call LoadAdjOutcomeCombo
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub DetailedReviewRationale_Change()
    mbRecordChanged = True
End Sub

Private Sub DetailedReviewRationale_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("DetailedReviewRationale")
End Sub

Private Sub EditParameters_Change()
    mbRecordChanged = True
End Sub

Private Sub EditParameters_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("EditParameters")
End Sub

Private Sub ErrorCode_Change()
    mbRecordChanged = True
End Sub

Private Sub ErrorCode2_Change()
    mbRecordChanged = True
End Sub

Private Sub Form_Current()
Debug.Print ClassName & "._Current"
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub

Private Sub Form_Load()

    Call LoadAdjOutcomeCombo

        '' NOTE: This is no longer being used for New concepts!
    
    Set mrsConceptHdr = New ADODB.RecordSet '' 20120416 KD: Early bound thanks!!!
    
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
        
                    
            ' Let's set the payer to all right from the get go
            ' ok, but ONLY for new ones..
        If IsConceptPayerType(Me.FormConceptID) = True Then
            Me.cmbPayer = 1000
        Else
            Me.cmbPayer = 0
'            me.cmbPayer.Enabled = False
            Call EnableDisableConvertWizard
        End If
    

        
        RefreshData
    End If
    DoCmd.SetWarnings False

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
Debug.Print Me.Name & ".Form_Unload"
    If mbAllowChange Then
        If Not (mbRecordChanged = False And Me.IsDirty = False) Then
            If MsgBox("Record has changed. Would you like to save changes to Concept - " & Me.FormConceptID & "?", vbYesNo + vbQuestion) = vbYes Then
                SaveData
            End If
        End If
    End If
'Stop

End Sub

Private Sub Hyperlink_Change()
    mbRecordChanged = True
End Sub

Private Sub Hyperlink_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("Hyperlink")
End Sub

Private Sub LiabilitySource_Change()
    mbRecordChanged = True
End Sub

Private Sub LOB_Change()
    mbRecordChanged = True
End Sub

Private Sub MedicalRecords_Change()
    mbRecordChanged = True
End Sub

Private Sub MedicalRecords_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("MedicalRecords")
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "ADO ERROR"
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "ADO ERROR"
End Sub

'''
'''
'''Private Sub Command225_Click()
'''
'''
'''    Stop
'''
'''On Error GoTo Err_Command225_Click
'''
'''
'''    Screen.PreviousControl.SetFocus
'''    DoCmd.FindNext
'''
'''Exit_Command225_Click:
'''    Exit Sub
'''
'''Err_Command225_Click:
'''    MsgBox Err.Description
'''    Resume Exit_Command225_Click
'''
'''End Sub


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


        '' SPecial case for Tuan and myself..
    Select Case LCase(Identity.UserName)
    Case "kevin.dearing", "tuan.khong"
        If gbRecordingVideo = False Then
            bEnabled = True
        Else
            bEnabled = False
        End If
        
        cmdReZip.Enabled = bEnabled
        cmdReZip.visible = bEnabled
        cmdCreateSampleClaimDocs.visible = bEnabled
        cmdCreateSampleClaimDocs.Enabled = bEnabled
        cmdCreateNIRF_For_Submit.visible = bEnabled
        cmdCreateNIRF_For_Submit.Enabled = bEnabled

        cmdTest.visible = bEnabled
        cmdTest.Enabled = bEnabled

        Me.cmdMarkAsSent.visible = bEnabled
        Me.cmdMarkAsSent.Enabled = bEnabled

        Me.cmdFinalizePkg.Enabled = bEnabled
        Me.cmdQA.Enabled = bEnabled

        If gbRecordingVideo = False Then GoTo Block_Exit
    Case Else
            ' Reset everything
        Me.cmdSubmit.Caption = csSubmitBtnText
        Me.cmdRegenEmail.visible = False
        cmdReZip.Enabled = False
        cmdReZip.visible = False
        cmdCreateSampleClaimDocs.visible = False
        cmdCreateSampleClaimDocs.Enabled = False
        cmdCreateNIRF_For_Submit.visible = False
        cmdCreateNIRF_For_Submit.Enabled = False
        
        Me.cmdFinalizePkg.Enabled = False
        Me.cmdQA.Enabled = False
        
        cmdTest.visible = False
        cmdTest.Enabled = False
    End Select
    
   
    
    ClientIssueNum.SetFocus
    If Nz(Me.ClientIssueNum.Text, "") <> "" Then
        bEnabled = False
    Else
        bEnabled = True
    End If
    Me.ConceptID.SetFocus
    
    
    Me.ReviewType.Enabled = True
    Me.DataType.Enabled = True
'    Me.ErrorCode.Enabled = True
'    Me.ErrorCode2.Enabled = True
    Me.ProviderTypeID.Enabled = True
    
    
    Me.ClientIssueNum.Locked = Not bEnabled
    If bEnabled = False Then
        Me.ClientIssueNum.BackColor = Me.Detail.BackColor
    Else
        Me.ClientIssueNum.BackColor = 16777215  '' White
    End If
    
'    Me.cmdRequestClientIssueId.Enabled = bEnabled
    Set oConcept = New clsConcept
    If Nz(Me.ConceptID, "") = "" Then
        GoTo Block_Exit
    End If
    
    
    If oConcept.LoadFromId(Me.ConceptID) = False Then
        Stop    ' hammer time
    Else
        ' Ok, if this is an old style concept AND it's already been submitted
        ' then we'll change the Submit button to a Re-Submit button
        If Me.IsConceptPayerSpecific = True Then
            If oConcept.SubmitTrackedDate(Me.PayerNameId) > CDate("1/1/1900") Then
                        ' New, payer specific
                        ' AND it's been submitted
                Me.cmdRegenEmail.Enabled = True
                '20120926: Handling resubmissions on a case by case basis now..
                
'                Me.cmdSubmit.Caption = csReSubmitBtnText
                Me.cmdSubmit.Enabled = False

                Me.cmdFinalizePkg.Enabled = False
                Me.cmdQA.Enabled = True

                Me.cmdCreateNIRF_n_ClaimDoc.Enabled = False
                Me.cmdCreateNIRF_n_ClaimDoc.visible = False
                
            ElseIf oConcept.PreviouslyPassedValidation(Me.PayerNameId) = True Then
                
                    ' New, Payer specific
                    '  HAS passed validation
                    
                If oConcept.AlreadySubmitted(Me.PayerNameId) > CDate("1/1/1900") Then
                        ' AND it's already been submitted
                        
                        '20120926: Handling resubmissions on a case by case basis now..
                    'Me.cmdRegenEmail.Enabled = True
                    'Me.cmdSubmit.Caption = csReSubmitBtnText
                    Me.cmdSubmit.Enabled = False
                    Me.cmdCreateNIRF_n_ClaimDoc.Enabled = False
                    Me.cmdFinalizePkg.Enabled = False
                    Me.cmdQA.Enabled = True
                Else
                        ' has NOT been submitted yet
                    Me.cmdRegenEmail.Enabled = False
                    Me.cmdSubmit.Caption = csSubmitBtnText
                    Me.cmdSubmit.Enabled = True
                    Me.cmdCreateNIRF_n_ClaimDoc.Enabled = True
                    
                    Me.cmdFinalizePkg.Enabled = True
                    Me.cmdQA.Enabled = True
                    
                End If
            ElseIf oConcept.PreviouslyPassedValidation(Me.PayerNameId) = False Then
                    '' needs to pass validation first
                Me.cmdRegenEmail.Enabled = False
                Me.cmdSubmit.Caption = csSubmitBtnText
                Me.cmdCreateNIRF_n_ClaimDoc.Enabled = False


            End If
        Else
            ' an old concept, that's been submitted and / or is not subject to the validation process
'                Me.cmdRegenEmail.Enabled = True
                'Me.cmdSubmit.Caption = csReSubmitBtnText
                Me.cmdSubmit.Caption = csSubmitBtnText
                
                Me.cmdSubmit.Enabled = True
                Me.cmdCreateNIRF_n_ClaimDoc.Enabled = False
        End If
    End If

  
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.ConceptID
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Sub




Private Sub ValidateForSubmission(lngPayerNameId As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oConcept As clsConcept
Dim dSubmitDate As Date
Dim sMsg As String
Dim oRs As ADODB.RecordSet
Dim oCheckListFrm As Form_frm_CONCEPT_Validation_Checklist
Dim bReady As Boolean

    strProcName = ClassName & ".ValidateForSubmission"
    DoCmd.Hourglass True
    DoCmd.Echo True, "Validating..."

    ' First thing is that we have to see if it's a payer specific concept:
    If IsConceptPayerType(Me.ConceptID) = False Then
        LogMessage strProcName, "ERROR", "This concept must be converted to a payer specific concept before it can be validated (or submitted)", Me.ConceptID, True, Me.ConceptID
        GoTo Block_Exit
    End If

    ' First, has it been submitted already?
    Set oRs = New ADODB.RecordSet
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.ConceptID) = False Then
        LogMessage strProcName, "ERROR", "Could not load concept object", Me.ConceptID, True, Me.ConceptID
        GoTo Block_Exit
    End If
    

    dSubmitDate = oConcept.AlreadySubmitted(lngPayerNameId)
    If dSubmitDate <> CDate("1/1/1900") Then
        LogMessage strProcName, "WARNING", "This Concept, " & Me.ConceptID & " is in the records as having been submitted on " & Format(dSubmitDate, "mm/dd/yyyy"), , True, Me.ConceptID
'        cmdRegenEmail.Enabled = True
        GoTo Block_Exit
    End If
    
        ' Now just validate everything:
    bReady = oConcept.ValidateForSubmission(oRs, lngPayerNameId)
    
    DoCmd.OpenForm "frm_CONCEPT_Validation_Checklist", acNormal, , , , acWindowNormal
    
    DoCmd.Hourglass False
    DoCmd.Echo True, "ready..."
    Set oCheckListFrm = Forms("frm_CONCEPT_Validation_Checklist")
    oCheckListFrm.ConceptID = Me.ConceptID
    oCheckListFrm.ValidationReport = oRs
    oCheckListFrm.visible = True
    oCheckListFrm.ShowReport oRs
    
    If oCheckListFrm.ValidationFailed = False Then
        ' it's fine..
        LogMessage strProcName, , "Validation success", Me.ConceptID, , Me.ConceptID
        Me.cmdSubmit.Enabled = True
'        cmdRegenEmail.Enabled = True
        
        Call SaveValidationHist(Me.ConceptID, lngPayerNameId, False, sMsg)
    Else
        Call SaveValidationHist(Me.ConceptID, lngPayerNameId, True, sMsg)
    End If
    

Block_Exit:
    DoCmd.Hourglass False
    DoCmd.Echo True, "Ready..."
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.ConceptID
    Err.Clear
    GoTo Block_Exit
End Sub


Private Function GetHdrRS() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".GetHdrRS"

    sSql = "EXEC usp_ConceptHdr_Get @pConceptId = '" & Me.FormConceptID & "'"

    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("v_Code_DATABASE")
    oCn.Open
    
    Set mrsConceptHdr = New ADODB.RecordSet
    mrsConceptHdr.CursorLocation = adUseClientBatch
    mrsConceptHdr.CursorType = adOpenStatic
    
    mrsConceptHdr.LockType = adLockBatchOptimistic
    
    mrsConceptHdr.Open sSql, oCn
    Set mrsConceptHdr.ActiveConnection = Nothing


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Function



Private Function GetPayerDtlRS() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".GetPayerDtlRS"

    If Me.IsPayerSetToAll = True Or Me.PayerNameId = 0 Then
        sSql = "SELECT * FROM CONCEPT_PAYER_Dtl WHERE ConceptId = '" & Me.ConceptID & "' AND PayerNameID <> 1000 "  '' 1000 = all
    Else
        sSql = "SELECT * FROM CONCEPT_PAYER_Dtl WHERE ConceptId = '" & Me.ConceptID & "' AND PayerNameId = " & Me.PayerNameId
    End If


    ' usp_ConceptPayerDtl_Get
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("v_DATA_DATABASE")
    oCn.Open
    
    Set mrsPayerDtl = New ADODB.RecordSet
    mrsPayerDtl.CursorLocation = adUseClientBatch
    mrsPayerDtl.CursorType = adOpenStatic
    mrsPayerDtl.LockType = adLockBatchOptimistic
    
    mrsPayerDtl.Open sSql, oCn
    Set mrsPayerDtl.ActiveConnection = Nothing


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Function




Private Function GetPayersForConceptRS() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim sSql As String
Dim iRealCount As Integer


    strProcName = ClassName & ".GetPayersForConceptRS"

    sSql = "EXEC usp_CONCEPT_PayerNames_by_ConceptId @pConceptID = '" & Me.FormConceptID & "'"
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("V_CODE_DATABASE")
    oCn.Open
    
    Set mrsThisConceptPayers = New ADODB.RecordSet
    mrsThisConceptPayers.CursorLocation = adUseClientBatch
    mrsThisConceptPayers.CursorType = adOpenStatic
    mrsThisConceptPayers.LockType = adLockBatchOptimistic
    
    mrsThisConceptPayers.Open sSql, oCn
    Set mrsThisConceptPayers.ActiveConnection = Nothing


    If Not mrsThisConceptPayers Is Nothing Then
        
        '' KD: Note: this may have 2, 1000 being one of them..
        ' so we can't really just look at the recordcount.
        mrsThisConceptPayers.MoveFirst
        While Not mrsThisConceptPayers.EOF
            If mrsThisConceptPayers("PayerNameId").Value <> 1000 Then
                iRealCount = iRealCount + 1
            End If
            mrsThisConceptPayers.MoveNext
        Wend
        mrsThisConceptPayers.MoveFirst
        
        If iRealCount = 1 Then
            cblnOnly1Payer = True
        Else
            cblnOnly1Payer = False
        End If
    Else
        cblnOnly1Payer = False  ' not even any payers
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Function


Private Function FilterPayerCombo() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim bJustAll As Boolean
Dim sFilter As String
Dim sOrigRowSourceWoWhere As String
Dim oRegEx As RegExp

    strProcName = ClassName & ".FilterPayerCombo"
    
    Call GetPayersForConceptRS
    Set Me.cmbPayer.RecordSet = mrsThisConceptPayers

    Me.cmbPayer.Requery
    
    Me.cmbPayer.ColumnCount = 2
    Me.cmbPayer.ColumnWidths = "0"";2"";"
    
    Me.cmbPayer = 1000  '' all
    FilterPayerCombo = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    FilterPayerCombo = False
    GoTo Block_Exit
End Function



Private Sub SetupPrevValsDict()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control

    strProcName = ClassName & ".SetupPrevValsDict"
    
    Set cdctPrevVals = New Scripting.Dictionary

    For Each oCtl In Me.Controls
        If InStr(1, CStr("" & oCtl.Tag), ".", vbTextCompare) > 0 Then

            Select Case UCase(TypeName(oCtl))
            Case "TEXTBOX", "COMBOBOX"
                If cdctPrevVals.Exists(oCtl.Name) Then
                    cdctPrevVals.Item(oCtl.Name) = oCtl.Value
                    Stop ' shouldn't get here
                Else
                    cdctPrevVals.Add oCtl.Name, oCtl.Value
                End If
            Case "CHECKBOX"
                If cdctPrevVals.Exists(oCtl.Name) Then
                    cdctPrevVals.Item(oCtl.Name) = Nz(oCtl.Value, False)
                    Stop ' shouldn't get here
                Else
                    cdctPrevVals.Add oCtl.Name, Nz(oCtl.Value, False)
                End If
            Case Else
                Debug.Print "Control is a " & TypeName(oCtl)
                
                Stop
            End Select
        End If
        
    Next
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub



Private Function DidAnythingChange(Optional bLookForHeaderToo As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim bRetVal As Boolean


    strProcName = ClassName & ".DidAnythingChange"
    
    If cdctPrevVals Is Nothing Then
        bRetVal = False
        GoTo Block_Exit
    End If

    For Each oCtl In Me.Controls
        If InStr(1, CStr("" & oCtl.Tag), ".", vbTextCompare) > 0 Then
            If InStr(1, CStr("" & oCtl.Tag), "CONCEPT_PAYER_Dtl", vbTextCompare) > 0 Or bLookForHeaderToo Then
                Select Case UCase(TypeName(oCtl))
                Case "TEXTBOX", "COMBOBOX"
                    If cdctPrevVals.Exists(oCtl.Name) Then
                        If oCtl.Value <> cdctPrevVals.Item(oCtl.Name) Then
                            bRetVal = True
                            ' no need to do any more!
                            Debug.Print "Field change detected: " & oCtl.Name & " from: " & cdctPrevVals.Item(oCtl.Name) & " to " & oCtl.Value
                        End If
                    Else
                        Stop ' uh, we should have everything here already!
                         
                    End If
                Case Else
                    Debug.Print "Control is a " & TypeName(oCtl)
                End Select
            End If
        End If
        
    Next
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    DidAnythingChange = True    ' err on the side of caution!
    GoTo Block_Exit
End Function


Private Sub EnableAndRecolorFields()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim bRetVal As Boolean
Dim bPayerLevelDetail As Boolean
Dim bCHasPayerInfo As Boolean

    '' White background color = 16777215
    
Const dWhite As Double = 16777215
Const dConceptColor As Double = 11468799    ' yellowish
Const dPayerColor As Double = 14876637      ' greenish
Const dBothColor As Double = 14474488
Const dGrey As Double = 12632256
Dim oPrevControl As Control

Dim sThisControlSource As String

    strProcName = ClassName & ".EnableAndRecolorFields"
   
    ' First, if this is NOT a new concept (i.e. not at payer level) then
    ' we don't do anything..
    ' well, except set back to normal:
    
    If mrsThisConceptPayers.recordCount > 1 Then    ' this will always have 1 (all / concept header level)
        bCHasPayerInfo = True
    Else
        bCHasPayerInfo = False
    End If
    
    bPayerLevelDetail = Not Me.IsPayerSetToAll
    
        ' get what currently has the focus:
    Set oPrevControl = Me.ActiveControl
    Me.cmdSave.SetFocus

   
    For Each oCtl In Me.Controls
        Select Case UCase(TypeName(oCtl))
        Case "TEXTBOX", "COMBOBOX"
        

                ' If the tag isn't set then it's not something we use for data right now.
            If CStr("" & oCtl.Tag) <> "" Then

                If bCHasPayerInfo = False Then  ' reset stuff:
                    'If left(oCtl.Name, 3) <> "txt" And left(oCtl.Name, 4) <> "Text" And CStr("" & oCtl.Tag) = "" Then
                    If CStr("" & oCtl.Tag) = "" Then
                        oCtl.Locked = True
                    Else
                        oCtl.Locked = False
                    End If
                    oCtl.BackColor = dWhite
                Else
                    ' the legacy code used to use R for something or other.. I didn't change any of that really
                    ' just added the tablename (or BOTH) then a dot then the R
                    If InStr(1, CStr("" & oCtl.Tag), ".", vbTextCompare) > 0 Then
                        sThisControlSource = UCase(left(oCtl.Tag, InStr(1, oCtl.Tag, ".", vbTextCompare) - 1))
                    Else
                        sThisControlSource = UCase(oCtl.Tag)
                    End If
                    
                        ' enable or disable controls and recolor them
                    If bPayerLevelDetail = True Then

                        Select Case sThisControlSource
                        Case "CONCEPT_HDR"
                            oCtl.Locked = True
                            oCtl.BackColor = dGrey
                            
                        Case "CONCEPT_PAYER_DTL"
                            If UCase(oCtl.Name) <> "CLIENTISSUENUM" Then
                                oCtl.Locked = False
                                oCtl.BackColor = dWhite
                            End If
                        Case "BOTH" ' because we're at detail level, it needs to be enabled
                            oCtl.Locked = False
                            oCtl.BackColor = dWhite
                        Case Else

                        End Select
                    Else    ' Concept Header stuff
                        Select Case sThisControlSource
                        Case "CONCEPT_HDR"
                            oCtl.Locked = False    ' maybe just locked
                            oCtl.BackColor = dWhite
                            
                        Case "CONCEPT_PAYER_DTL"
                            oCtl.Locked = True
                            oCtl.BackColor = dGrey
                        Case "BOTH" ' because we're at detail level, it needs to be enabled
                            oCtl.Locked = False
                            oCtl.BackColor = dBothColor
                        Case Else
                            'Stop    ' shouldn't get here pal!
                            ' Right now, this is only the audit fields last update date and last update user so we're going to leave this be
                        End Select
                    End If

                End If
            
                
            End If
        Case Else
        
        End Select
        
    Next
    
   ' reset the focus:
   On Error Resume Next ' in case the control was a cmd button or what have you
    If oPrevControl.Locked = False Then
        oPrevControl.SetFocus
    End If
        
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub


Private Function IsControl(ByVal sNameToCheck As String) As Boolean
Static dctControlNames As Scripting.Dictionary
Dim oCtl As Control

    sNameToCheck = UCase(sNameToCheck)
    
    If dctControlNames Is Nothing Then
        Set dctControlNames = New Scripting.Dictionary
        For Each oCtl In Me.Controls
            dctControlNames.Add UCase(CStr("" & oCtl.Name)), TypeName(oCtl)
        Next
    End If

    IsControl = IIf(dctControlNames.Exists(sNameToCheck), True, False)

End Function

Private Sub IsLockedMsg(sControlName As String)
Dim oCtl As Control

    If IsControl(sControlName) = False Then
        Exit Sub    ' nothing to do - must be a coding error or something
    End If
    Set oCtl = Me.Controls(sControlName)
    
    If oCtl.Locked = True Then
        If Me.IsPayerSetToAll = True Then
            MsgBox "Please select a payer", vbCritical, "This is Payer data!"
        Else
            MsgBox "This is Concept Header detail, please select that from the Payer detail to edit!"
        End If
    End If

End Sub


Private Sub EnableDisableConvertWizard()
Dim bEnable As Boolean


    ' enable if: This concept has NO payers
    ' AND it hasn't been submitted yet
    
    If IsConceptPayerSpecific = False Then
        If CStr(Nz(Me.DateSubmitted, "1/1/1900") = "1/1/1900") Then
            bEnable = True
        Else
            bEnable = False
        End If
    End If
    
    If bEnable = False Then
        Me.txtHilight.BackStyle = 0 ' Transparent
    Else
        Me.txtHilight.BackStyle = 1 ' normal
    End If


Block_Exit:

    cmdConvertToPayerConcept.Enabled = bEnable

    Exit Sub
End Sub


Private Sub FlipNotesSubform()
Dim sPayerIdsForThisConcept As String


    If IsConceptPayerSpecific = False Then
        Me.ctl_frm_CONCEPT_Tracking.SourceObject = "frm_CONCEPT_Tracking_Orig"
        
        Me.ctl_frm_CONCEPT_Tracking.LinkMasterFields = ""
        Me.ctl_frm_CONCEPT_Tracking.LinkChildFields = ""
        
        Me.ctl_frm_CONCEPT_Tracking.LinkMasterFields = "ConceptId"
        Me.ctl_frm_CONCEPT_Tracking.LinkChildFields = "ConceptId"
        
        Me.ctl_frm_CONCEPT_Tracking.Form.filter = ""
        Me.ctl_frm_CONCEPT_Tracking.Form.FilterOn = False
        
    Else
        Me.ctl_frm_CONCEPT_Tracking.SourceObject = "frm_CONCEPT_Tracking"


        Me.ctl_frm_CONCEPT_Tracking.LinkMasterFields = "ConceptId"
        Me.ctl_frm_CONCEPT_Tracking.LinkChildFields = "ConceptId"
            
            ' but we need to filter it (unless all is selected)
        If Nz(Me.cmbPayer.Value, 1000) = 1000 Then
            Me.ctl_frm_CONCEPT_Tracking.Form.filter = ""
            Me.ctl_frm_CONCEPT_Tracking.Form.FilterOn = False
        Else
            Me.ctl_frm_CONCEPT_Tracking.Form.filter = "PayerNameId = " & CStr(Nz(Me.cmbPayer.Value, 1000))
            Me.ctl_frm_CONCEPT_Tracking.Form.FilterOn = True
        End If
        '' Also, we need to filter the drop down so they can only add payers that are associated with this conept
        sPayerIdsForThisConcept = mod_Concept_Specific.GetRelatedPayerNameIDsForFilter(Me.ConceptID)
        Me.ctl_frm_CONCEPT_Tracking.Form.PayersForThisConcept = sPayerIdsForThisConcept
    End If

End Sub


Private Sub NextContract_Click()
    mbRecordChanged = True
End Sub

Private Sub OpportunityType_Change()
    mbRecordChanged = True
End Sub

Private Sub OtherDocumentation_Change()
    mbRecordChanged = True
End Sub

Private Sub OtherDocumentation_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("OtherDocumentation")
End Sub

Private Sub ProviderTypeID_Change()
    mbRecordChanged = True
End Sub

Private Sub ReferralEntity_Change()
    mbRecordChanged = True
End Sub

Private Sub ReferralFlag_Change()
    mbRecordChanged = True
End Sub

Private Sub RepriceFlag_Click()
    mbRecordChanged = True
End Sub

Private Sub ReviewType_Change()
    mbRecordChanged = True
End Sub

Private Sub SampleOfClaims_Change()
    mbRecordChanged = True
End Sub

Private Sub SampleOfClaims_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IsLockedMsg("SampleOfClaims")
End Sub

Private Sub TempSubmitProcess(Optional bResubmitting As Boolean)
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim colPayers As Collection
'Const sToAddress As String = "Andrew.Lauer@connolly.com;Damon.Ramaglia@connolly.com;Gautam.Malhotra@connolly.com;James.Segura@connolly.com;Tuan.Khong@connolly.com;Kevin.Dearing@connolly.com"
Dim sToAddress As String
Dim vPayerNameID As Variant

    strProcName = ClassName & ".TempSubmitProcess"
    
        
    
    sToAddress = "Kevin.Dearing@connolly.com;Tuan.Khong@connolly.com;"


    '' Zip the stuff up...
    If Me.cmbPayer = 1000 Or Me.cmbPayer = 0 Then
        ' All payers:
        '' How do I get all payers for this concept again? I know I have a function somewhere..
        ' maybe there's a property..  I guess it should really be a property of
        ' the concept class
        Set colPayers = GetRelatedPayerNameIDs(Me.FormConceptID)
        
        For Each vPayerNameID In colPayers
            ' Need to check that the concept's status is good to go..
            Me.PayerNameId = CLng(vPayerNameID)
            Call ZipConceptSubmitPackage(coCurConcept, CLng(vPayerNameID), bResubmitting)
                '' log the details to the DB:
            Call MarkConceptAsSubmitted(Me.FormConceptID, CLng(vPayerNameID), bResubmitting)

        Next
        
    Else
        ' 1 at a time:
        Me.PayerNameId = CLng(vPayerNameID)
        Call ZipConceptSubmitPackage(coCurConcept, Me.cmbPayer, bResubmitting)
            '' log the details to the DB:
        Call MarkConceptAsSubmitted(Me.FormConceptID, CLng(Me.cmbPayer), bResubmitting)

    End If
    
    sMsg = Me.Auditor & " has clicked submit for concept '" & Me.FormConceptID & "' and payer: " & GetPayerNameFromID(Me.cmbPayer) & "." & vbCrLf & vbCrLf & "Please submit it via NDM or by the appropriate means for this payer!"
    
    
    
    Dim sThisUser As String
    sThisUser = Identity.UserName()
    sThisUser = sThisUser & "@connolly.com"

    ' This is until I can do the real deal..
    ' Not going to validate here, we're just going to send dataservices an email and
    ' copy ken
'Stop
        ' how about sending the email
    SendsqlMail "[CONCEPT MGMT] Concept Submission: " & Me.FormConceptID & " : " & GetPayerNameFromID(Me.PayerNameId), sToAddress & Me.Auditor & "@connolly.com;", "Kenneth.Turturro@connolly.com;" & sThisUser, "", sMsg

    LogMessage strProcName, "CONFIRMATION", "The concept has just been submitted for payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub



Private Sub CreateSubmitPackage()
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim colPayers As Collection
Dim sToAddress As String
Dim vPayerNameID As Variant
Dim sThisUser As String
Dim oPayer As clsConceptPayerDtl
Dim iPayersCreated As Integer
Dim sPayerList As String

    strProcName = ClassName & ".CreateSubmitPackage"
    
    sToAddress = "Kevin.Dearing@connolly.com;"


    '' Zip the stuff up...
    If Me.cmbPayer = 1000 Or Me.cmbPayer = 0 Then
        ' All payers:
        '' How do I get all payers for this concept again? I know I have a function somewhere..
        ' maybe there's a property..  I guess it should really be a property of
        ' the concept class
        Set colPayers = GetRelatedPayerNameIDs(Me.FormConceptID)
        
        For Each vPayerNameID In colPayers
            ' Need to check that the concept's status is good to go..
            
            Me.PayerNameId = CLng(vPayerNameID)
            
            
            Set oPayer = New clsConceptPayerDtl
            If oPayer.LoadFromConceptNPayer(coCurConcept.ConceptID, Me.PayerNameId) = False Then
                LogMessage strProcName, "ERROR", "There was a problem loading the Payer for " & coCurConcept.ConceptID & " and payer: " & CStr(Me.PayerNameId), , , Me.FormConceptID
                GoTo Block_Exit
            End If
            
            If oPayer.PayerStatusNum = "990" Then  ' void
                LogMessage strProcName, "WARNING", "This payer has a VOID status.. Not creating package!", oPayer.PayerName, , Me.FormConceptID
            Else
                If oPayer.EffectiveDate > Now() Or oPayer.EndDate < Now() Then
                    LogMessage strProcName, "ERROR", "This payer, '" & oPayer.PayerName & "' is not valid at this time!", Format(oPayer.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayer.EndDate, "mm/dd/yyyy"), True, Me.FormConceptID
                Else
                    iPayersCreated = iPayersCreated + 1
                    Call ZipConceptSubmitPackage(coCurConcept, CLng(vPayerNameID))
                    sPayerList = sPayerList & GetPayerNameFromID(CLng(vPayerNameID)) & ", "
                End If
                
            End If

        Next
        
        If Right(sPayerList, 2) = ", " Then
            '' Remove final comma + space
            sPayerList = left(sPayerList, Len(sPayerList) - 2)
        End If
        
    Else
        ' 1 at a time:
        Dim lThisPayer As Long
        
        If Not IsEmpty(vPayerNameID) Then
            Me.PayerNameId = CLng(vPayerNameID)
        End If
        lThisPayer = Me.PayerNameId

        Set oPayer = New clsConceptPayerDtl
        If oPayer.LoadFromConceptNPayer(coCurConcept.ConceptID, Me.PayerNameId) = False Then
            LogMessage strProcName, "ERROR", "There was a problem loading the Payer for " & coCurConcept.ConceptID & " and payer: " & CStr(Me.PayerNameId), , , Me.FormConceptID
            GoTo Block_Exit
        End If
        
        If oPayer.PayerStatusNum = "990" Then  ' void
            LogMessage strProcName, "WARNING", "This payer has a VOID status.. Not creating package!", oPayer.PayerName, True, Me.FormConceptID
            GoTo Block_Exit
        Else
                
            If oPayer.EffectiveDate > Now() Or oPayer.EndDate < Now() Then
                LogMessage strProcName, "ERROR", "This payer, '" & oPayer.PayerName & "' is not valid at this time!", Format(oPayer.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayer.EndDate, "mm/dd/yyyy"), True, Me.FormConceptID
            
            Else
                sPayerList = GetPayerNameFromID(lThisPayer)
                iPayersCreated = iPayersCreated + 1
                Call ZipConceptSubmitPackage(coCurConcept, Me.cmbPayer)
                    '' log the details to the DB:
'                Call MarkConceptAsSubmitted(Me.FormConceptID, CLng(Me.cmbPayer))
                
            End If
        End If

    End If
    
    If iPayersCreated > 0 Then
        
        sMsg = Me.Auditor & " has generated the package for concept '" & Me.FormConceptID & "' and payer(s): " & sPayerList & "." & vbCrLf & vbCrLf
        
        sMsg = sMsg & vbCrLf & vbCrLf & "NOTE: The next step is for the auditor to QA the package and make sure that all of the proper documents exist and are correctly filled out! Please see the video: CMS - Concept Management - How to QA a concept package! for details!"
        
        sThisUser = Identity.UserName()
        sThisUser = sThisUser & "@connolly.com"
    
        ' This is until I can do the real deal..
        ' copy ken

            ' how about sending the email
        SendsqlMail "[CONCEPT MGMT] Concept Package Creation: " & Me.FormConceptID & " : " & sPayerList, sToAddress & Me.Auditor & "@connolly.com;", "Kenneth.Turturro@connolly.com;" & sThisUser, "", sMsg
    
        LogMessage strProcName, "CONFIRMATION", "The concept package has just been created for payer: " & sPayerList, , True, Me.FormConceptID
    End If

Block_Exit:
    Set oPayer = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub


Private Sub IT_Mark_Concept_as_Sent_via_NDM()
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim sSubmitAuditorEmail As String

Const sToAddress As String = "Gautam.Malhotra@connolly.com;Tuan.Khong@connolly.com;Kevin.Dearing@connolly.com;"

    strProcName = ClassName & ".IT_Mark_Concept_as_Sent_via_NDM"

    ' Log it as sent,
    sMsg = Me.Auditor & "'s Concept: '" & Me.FormConceptID & "' and payer: " & GetPayerNameFromID(Me.PayerNameId) & " has been physically sent to the payer by " & _
        Identity.UserName() & vbCrLf & vbCrLf & "Please submit it via NDM or the appropriate means for this payer!"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkSent"
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.FormConceptID
        .Parameters("@pPayerNameID") = Me.PayerNameId
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, Me.FormConceptID
            GoTo Block_Exit
        End If
        sSubmitAuditorEmail = .Parameters("@pSubmitEmail").Value
    End With
    

        ' Send an email to Ken and the auditor that it was indeed sent
    SendsqlMail "[CONCEPT MGMT] Concept Sent: " & Me.FormConceptID & " : " & GetPayerNameFromID(Me.PayerNameId), sSubmitAuditorEmail, "Kenneth.Turturro@connolly.com;" & sToAddress, "", sMsg
    
    '' Then, we need to generate the canned email to send to us which can then be fwd to the payer
    
    LogMessage strProcName, "NOTE TO USER", "Concept has been sent to payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub



Private Sub SendCMSEmailToKen()
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim sSubmitAuditorEmail As String

Const sToAddress As String = "Kenneth.Turturro@connolly.com;"
Const sCCAddress As String = "Gautam.Malhotra@connolly.com;Tuan.Khong@connolly.com;Kevin.Dearing@connolly.com;"

    strProcName = ClassName & ".SendCMSEmailToKen"

    ' Log it as sent,
    sMsg = Me.Auditor & " has clicked submit for concept '" & Me.FormConceptID & "' and payer: " & GetPayerNameFromID(Me.PayerNameId) & "." & vbCrLf & vbCrLf & "Please submit it via NDM or the appropriate means for this payer!"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkSent"
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.FormConceptID
        .Parameters("@pPayerNameID") = Me.PayerNameId
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, Me.FormConceptID
            GoTo Block_Exit
        End If
        sSubmitAuditorEmail = .Parameters("@pSubmitEmail").Value
    End With
    

        ' Send an email to Ken and the auditor that it was indeed sent
    SendsqlMail "[CONCEPT MGMT] Concept Sent: " & Me.FormConceptID & " : " & GetPayerNameFromID(Me.PayerNameId), sSubmitAuditorEmail, "Kenneth.Turturro@connolly.com;" & sToAddress, "", sMsg
    
    '' Then, we need to generate the canned email to send to us which can then be fwd to the payer
    
    LogMessage strProcName, "NOTE TO USER", "Concept has been sent to payer: " & GetPayerNameFromID(Me.PayerNameId), , True, Me.FormConceptID

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Sub




Public Sub SendsqlMail(ByVal Subject As String, ByVal ToList As String, ByVal CCList As String, ByVal BCCList As String, ByVal Body As String)
    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    Subject = Replace(Subject, "'", "''")
    Body = Replace(Body, "'", "''")

    MyAdo.sqlString = "EXEC Cnly.Mail.SqlNotifySend  @Subject= '" & Subject & "', @ToList ='" & ToList & "', @CCList = '" & CCList & "',@BCCList = '" & BCCList & "',@Body = '" & Body & "',@Result = NULL,@Output = NULL, @Priority = 0, @format = 0"
    MyAdo.SQLTextType = sqltext
    MyAdo.Execute
End Sub

Private Sub TabCtl128_Change()
Dim oFrm As Form_frm_CONCEPT_Production_Notes

    CurrentPageSelected = Me.TabCtl128
    If Me.TabCtl128 = Me.pgProductionNotes.PageIndex Then
        Set oFrm = Me.sfrm_CONCEPT_Production_Notes.Form
        Call oFrm.RefreshData
        Set oFrm = Nothing
    End If

End Sub

Private Sub TabCtl128_Click()
Dim oFrm As Form_frm_CONCEPT_Production_Notes
    CurrentPageSelected = Me.TabCtl128
    If Me.TabCtl128 = Me.pgProductionNotes.PageIndex Then
        Set oFrm = Me.sfrm_CONCEPT_Production_Notes.Form
        Call oFrm.RefreshData
        Set oFrm = Nothing
    End If
End Sub

Private Sub Text241_Change()
mbRecordChanged = True
End Sub

Private Sub Text243_Change()
    mbRecordChanged = True
End Sub





Private Function GetPayerCollection() As Collection
On Error GoTo Block_Err
Dim strProcName As String
Dim sPayerName As String

    strProcName = ClassName & ".GetPayerCollection"

    
    sPayerName = GetPayerNameFromID(Me.cmbPayer)
    
    If Me.IsPayerSetToAll = True Then
        Set coPayers = Nothing  ' make sure it's refreshed..
        Set coPayers = mod_Concept_Specific.GetRelatedPayerNameIDs(Me.FormConceptID)
    Else
'        If MsgBox("Are you sure you want to create the NIRF for " & sPayerName & "?", vbOKCancel, "Create NIRF for " & sPayerName) = vbCancel Then
'            GoTo Block_Exit
'        End If
        Set coPayers = New Collection
        coPayers.Add Me.cmbPayer.Value
    End If


    
Block_Exit:
    DoCmd.Hourglass False
    Set GetPayerCollection = coPayers
    Exit Function
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    Err.Clear
    GoTo Block_Exit
End Function


'' Currently, we do not want to let users save the concept with a status of 380 (Posted to website) if
'' the LCD
' Returns true if the concept's lcd change flag is ticked and it hasn't been manually released
'  but ONLY if the user is attempting to change the status to 380 (Posted to website)
Public Function StatusChangeValidationFailed(Optional sMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".StatusChangeValidation"
 
    '' Check the concept status - if that didn't change, who cares!
    If Me.ConceptStatus = cdctPrevVals.Item("ConceptStatus") Then
        ' no change!
        StatusChangeValidationFailed = False ' it didn't fail validation - we're good to go
        GoTo Block_Exit
    End If
 
    ' Finally, we can't allow them to change the status to 380 (posted to website) if the LCD change flag
    '' is checked (and it hasn't been released yet)
    
    If IsConceptPayerSpecific() = True Then
        If IsPayerSetToAll() = True Then
            ' All payers need to be validated..
            ' wait, if it's a payer specific concept, and all payers is selected, then they can't update the status since that's a payer specific field
            ' I should put some code here just to check anyway
            Stop
            
            ' pointless really.. but I guess I should check the manual release date to see if it's < than the last rundate
            ' if so, we're good, else return true to indicate failure
            
        Else    ' 1 payer at a time

                ' check the manual release date to see if it's > than the last rundate
                ' if so, we're good, else return true to indicate failure
            StatusChangeValidationFailed = IsLcdChangeFlagChecked(sMsg)
            Stop
                    
        End If
    Else
        ' just look at the concept status - not
        If IsPayerSetToAll() = True Then
            ' good, we already know that the status changed so just check if the flag is ticked and the manual releasedate < last run
            ' if so, return true if the releasedate > then return false - we're good.
            Stop
            StatusChangeValidationFailed = IsLcdChangeFlagChecked(sMsg)
Stop
        Else
            ' then stop the show!!
            ' huh? this is not a payer specific concept.. how can the payer be set to a specific payer.
            Stop
        End If
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Function IsLcdChangeFlagChecked(Optional sReasonMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String


    strProcName = ClassName & ".IsLcdChangeFlagChecked"
    
    sSql = "SELECT * FROM CONCEPT_Lcd WHERE ConceptId = '" & Me.FormConceptID & "' "
    If Me.IsConceptPayerSpecific = False Then
        ' nothing to add. we don't need to add payernameid to the where clause because
        ' any of them will kill us
    Else
        sSql = sSql & " AND (PayerNameId = 1000 or PayerNameId = " & CStr(Me.PayerNameId) & " ) "
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        
        Set oRs = .ExecuteRS
        If .GotData = False Then
            ' no problem, no LCD's for this one..
            IsLcdChangeFlagChecked = False
            GoTo Block_Exit
        End If
    End With
    
    If oRs("ChangeFlag").Value <> 0 Then
        If oRs("LastReportDt").Value > Nz(oRs("ReleaseFlagDt").Value, CDate("1/1/1900")) Then
            IsLcdChangeFlagChecked = True
            sReasonMsg = "The LCD Change flag is checked indicating that something changed in the LCD database with an LCD that this concept is associated with.  " & _
                "Please review the LCD changes and release the Change Flag!"
Stop
        Else
            ' not a problem, manually released after the last LCD database import
            IsLcdChangeFlagChecked = False
        End If
    Else
        ' no problem - not changed
        IsLcdChangeFlagChecked = False
        GoTo Block_Exit
    End If
    
    
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function AllowedToChangeConceptStatus() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iOldConceptStatus As Integer
Dim bUpdatable As Boolean


    strProcName = ClassName & ".AllowedToChangeConceptStatus"
    
    bUpdatable = True
    '' Automated concepts don't need ADR letters:
    If coCurConcept.GetField("ReviewType") = "A" Then
        ' all is well automated concepts don't need an ADR letter
        bUpdatable = True
        GoTo Block_Exit
    End If

Stop
' we should probably allow managers to edit them no matter what?
    If IsMgrOrDS = True Then
        bUpdatable = True
        GoTo Block_Exit
    End If

    iOldConceptStatus = CInt(cdctPrevVals.Item("ConceptStatus"))
    
'    Select Case iOldConceptStatus
'    Case 380, 381, Is > 381
'        bUpdatable = False
'
'    Case Else
'        bUpdatable = True
'    End Select
    
    If bUpdatable = False Then
        LogMessage strProcName, "USER WARNING", "You may not change the status of this concept at this time!" & vbCrLf & "Please have a manager or Data Services change it if that is the correct thing to do", , True, Me.FormConceptID
        GoTo Block_Exit
    End If
    
    ' Not going to deal with this yet.. We will need a table to drive this stuff..
    '' well, then again, I guess we should make sure that they've picked the adr letter..
    Select Case Me.ConceptStatus
    Case Is > 379
        If Nz(Me.txtADRLetter, cs_ADR_NOT_SELECTED_MSG) = cs_ADR_NOT_SELECTED_MSG Then
            LogMessage strProcName, "USER ERROR!", "In order to save the concept with this status an ADR letter needs to be selected first!", , True, Me.FormConceptID
            bUpdatable = False
        Else
            bUpdatable = True
        End If
    End Select
    
Block_Exit:
    AllowedToChangeConceptStatus = bUpdatable
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Sub RollbackStatusToLoadedValue()
On Error GoTo Block_Err
Dim strProcName As String
Dim iOldConceptStatus As Integer

    strProcName = ClassName & ".RollbackStatusToLoadedValue"

    If cdctPrevVals.Exists("ConceptStatus") = False Then
Stop
    End If
    iOldConceptStatus = CInt(cdctPrevVals.Item("ConceptStatus"))
    Me.ConceptStatus = iOldConceptStatus
    Stop
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub SaveSelectedAdrLetter(sSelectedLetterType As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String

    strProcName = ClassName & ".SaveSelectedAdrLetter"
    
    sSql = "usp_LETTER_SaveAdrLetterType"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = sSql
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.FormConceptID
        .Parameters("@pPayerNameId") = Nz(Me.PayerNameId, 999)
        .Parameters("@pLetterType") = sSelectedLetterType
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
    Stop
        End If
        
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
