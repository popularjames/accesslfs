Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents coRs As ADODB.RecordSet
Attribute coRs.VB_VarHelpID = -1
Private cbIsDirty As Boolean

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = cbIsDirty
End Property
Public Property Let IsDirty(bIsDirty As Boolean)
    cbIsDirty = bIsDirty
End Property




Private Sub CmdCancel_Click()

    If Me.IsDirty = True Then
        If MsgBox("The data has changed but you did not save it. " & vbCrLf & "Save?", vbYesNo, "Save?") = vbYes Then
            Call cmdSave_Click
        End If
    End If
    
    If IsOpen("frm_CONCEPT_NIRF_Samples_Manual_Edit") Then
    
    End If
    If IsOpen("frm_CONCEPT_NIRF_Universe_Manual_Edit") Then

    End If
    
    
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub cmdEditSampleDetails_Click()
Dim sOpenArgs As String
'    MsgBox "Not yet finished! Stay tuned!"
'
'    Exit Sub

    If Nz(Me.OpenArgs, "") = "" Then
        sOpenArgs = " CONCEPTID = '" & Me.txtConceptID & "' AND PayerNameId = " & Nz(Me.PayerNameId, 1000)
    Else
        sOpenArgs = Me.OpenArgs
    End If
    
    DoCmd.OpenForm "frm_CONCEPT_NIRF_Sample_Manual_Edit", acNormal, , , acFormEdit, , sOpenArgs

End Sub

Private Sub cmdEditUniverse_Click()
Dim sOpenArgs As String

    If Nz(Me.OpenArgs, "") = "" Then
        sOpenArgs = " CONCEPTID = '" & Me.txtConceptID & "' AND PayerNameId = " & Nz(Me.PayerNameId, 1000)
    Else
        sOpenArgs = Me.OpenArgs
    End If

    DoCmd.OpenForm "frm_CONCEPT_NIRF_Universe_Manual_Edit", acNormal, , , acFormEdit, , sOpenArgs

End Sub

Private Sub cmdRevert_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_CONCEPT_NIRF_Editor_Revert_List
Dim bFormClosed   As Boolean
Dim strFormName As String
    
    strProcName = ClassName & ".cmdRevert_Click"
    
    If Nz(Me.txtManualEditId, 0) = 0 Then
        LogMessage strProcName, "USER NOTICE", "There are no previous edits for this NIRF.", , True, Me.txtConceptID.Value

        GoTo Block_Exit
    End If
    
    Set oFrm = New Form_frm_CONCEPT_NIRF_Editor_Revert_List
    ColObjectInstances.Add oFrm, oFrm.hwnd & ""
    oFrm.ConceptID = Me.txtConceptID.Value
    oFrm.PayerNameId = Me.PayerNameId
    oFrm.RefreshData
    
     strFormName = oFrm.Name
     oFrm.visible = True

     Do
        'Is it still Open?
        If IsLoaded(strFormName) Then
            DoEvents
            Wait 1
        ElseIf oFrm.visible = False Then
            bFormClosed = True
        Else
            bFormClosed = True
        End If
        
        If oFrm.AutoIdSelected <> 0 Then
            bFormClosed = True
        End If
        If oFrm.Canceled = True Then
            bFormClosed = True
        End If
        
     Loop Until bFormClosed = True
    
'    ShowFormAndWait oFrm
    
    Call RefreshData
    
'    Set frmCalendar = New Form_frm_GENERAL_Calendar
'    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
'    frmCalendar.DatePassed = Nz(Me.txtThroughDate, Date)
'    frmCalendar.RefreshData

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim oAdo As clsADO
Dim sXmlParam As String
Dim sXml As String

    strProcName = ClassName & ".cmdSave_Click"
    
    ' If we already have a ManualEditId then we need to UPDATE, or, should we archive, delete then insert?
    ' we'll let the stored proc decide that but I think we are always going to insert
    ' and I'll have to change the view to get the use the most recent one (but then again, we're giving them a
    ' revert option..
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_NIRF_Save_Manual_Changes"
        .Parameters.Refresh
        
    End With
    
    For Each oCtl In Me.Controls
        If oCtl.Tag <> "" Then
            sXmlParam = sXmlParam & "p" & oCtl.Tag & "=" & Nz(oCtl.Value, "") & "|"
        End If
    Next
    
    sXml = BuildXmlParams(sXmlParam)
    oAdo.Parameters("@pXmlParams") = sXml
Debug.Print sXml
    
    oAdo.Execute
    If Nz(oAdo.Parameters("@pErrMsg"), "") <> "" Then
        LogMessage strProcName, "ERROR", "There was an error saving the NIRF", oAdo.Parameters("@pErrMsg").Value, True, Me.txtConceptID
    Else
        DoCmd.Close acForm, Me.Name
    End If
    
    Me.IsDirty = False
    
Block_Exit:
    Set oCtl = Nothing
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim sFilter As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".RefreshData"
    
    Me.RecordSource = ""
    Set Me.RecordSet = Nothing
    
'    Me.OpenArgs = "CM_C1447"
'    sFilter = " WHERE ConceptId = 'CM_C2055' " ' AND PayerNameId = 1001 "

    If Me.OpenArgs <> "" Then
        sFilter = "WHERE " & Me.OpenArgs
    End If
    
    
    sSql = "SELECT * FROM v_CONCEPT_NIRF_ALL " & sFilter
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("v_Code_Database")
    oCn.CursorLocation = adUseClientBatch
    oCn.Open
    
    Set coRs = Nothing
    Set coRs = New ADODB.RecordSet
    
    coRs.CursorLocation = adUseClientBatch
    coRs.CursorType = adOpenKeyset
    coRs.LockType = adLockBatchOptimistic
    coRs.Open sSql, oCn
    
    ' disconnect:
    Set coRs.ActiveConnection = Nothing
    
'

'

'    'Loop through the controls setting their control source to the recordset
'    For Each ctl In Me.Controls
'    'MsgBox ctl.Name, vbOKOnly
'        If ctl.Tag = "R" Then
'             Me.Controls(ctl.Name).ControlSource = mrsCollCnlyAdjustment.Fields(ctl.Name).Name
'        End If
'    Next
        
    
'    Stop
    
    
'    Debug.Print TypeName(oCtl.co)
'    Set Me.Recordset = Nothing
'    Set Me.Recordset = coRS
    For Each oCtl In Me.Controls
        If oCtl.Tag <> "" Then
            If isField(coRs, oCtl.Tag) = True Then
                Me.Controls(oCtl.Name).ControlSource = ""   'oCtl.Tag
                oCtl.Value = coRs(oCtl.Tag).Value

            End If
        End If
    Next
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub coRS_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.RecordSet)
    Me.IsDirty = True
End Sub


Private Sub coRS_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.RecordSet)
    Me.IsDirty = True
End Sub



Private Sub coRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.RecordSet)
    Me.IsDirty = True
End Sub

Private Sub coRS_WillChangeRecordset(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.RecordSet)
    Me.IsDirty = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Stop
End Sub

Private Sub Form_DataChange(ByVal Reason As Long)
Stop
End Sub

Private Sub Form_DataSetChange()
Stop
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.IsDirty = True
End Sub

Private Sub Form_Load()

    '' need to bind this to the recordset..
    Call RefreshData
    
End Sub



Private Sub txt01_NameOfRac_Change()
    Me.IsDirty = True
End Sub



Private Sub txt02a_DateSubmitted_Change()
    Me.IsDirty = True
End Sub


Private Sub txt02b_ResubmittedDt_Change()
    Me.IsDirty = True
End Sub


Private Sub txt03a_POCName_Change()
    Me.IsDirty = True
End Sub



Private Sub txt03b_Email_Change()
    Me.IsDirty = True
End Sub



Private Sub txt03c_Phone_Change()
    Me.IsDirty = True
End Sub


Private Sub txt04_IssueName_Change()
    Me.IsDirty = True
End Sub


Private Sub txt05_IssueDesc_Change()
    Me.IsDirty = True
End Sub



Private Sub txt07_OpportunityType1_Change()
    Me.IsDirty = True
End Sub



Private Sub txt08_ReviewType1_Change()
    Me.IsDirty = True
End Sub



Private Sub txt09_LiabilitySource_Change()
    Me.IsDirty = True
End Sub



Private Sub txt10_ConceptIndicator_Change()
    Me.IsDirty = True
End Sub



Private Sub txt11_ErrorCode_Change()
    Me.IsDirty = True
End Sub


Private Sub txt12_ProviderDesc_Change()
    Me.IsDirty = True
End Sub



Private Sub txt13_ConceptReferences_Change()
    Me.IsDirty = True
End Sub



Private Sub txt14_Comments_Change()
    Me.IsDirty = True
End Sub



Private Sub txt15_StateList_Change()
    Me.IsDirty = True
End Sub



Private Sub txt16_StateList_Change()
    Me.IsDirty = True
End Sub


Private Sub txt17_Hyperlink_Change()
    Me.IsDirty = True
End Sub



Private Sub txt23a_ReferralFlag1_Change()
    Me.IsDirty = True
End Sub



Private Sub txt23b_ReferralEntity_Change()
    Me.IsDirty = True
End Sub



Private Sub txt24_ClientIssueNum_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25a_CopyOfReferences_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25b_ConceptRationale_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25c_DetailedReviewRational_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25d_MedicalRecords_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25e_SampleOfClaims_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25f_EditParameters_Change()
    Me.IsDirty = True
End Sub


Private Sub txt25g_OtherDocuments_Change()
    Me.IsDirty = True
End Sub


Private Sub txt26_AdditionalInfo_Change()
    Me.IsDirty = True
End Sub
