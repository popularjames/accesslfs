Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const CstrFrmAppID As String = "NextContract"
Private csConceptId As String
Private cblnFilterApplied As Boolean
Private cbDirty As Boolean



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get FormConceptID() As String
    FormConceptID = IdValue
End Property
Public Property Let FormConceptID(sConceptId As String)
    IdValue = sConceptId
End Property


Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    
    Call Me.RefreshData

Block_Exit:
    Exit Property
End Property



Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Sub PayerChange()
    cmbPayer_Change
End Sub


Private Sub FilterByPayer()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_CONCEPT_Main

    strProcName = ClassName & ".FilterByPayer"
    
    
    If Nz(cmbPayer.Value, 1000) = 1000 Then
        ' No filter:
        Me.filter = ""
        Me.FilterOn = False
    Else
        Me.filter = "PayerNameId = " & CStr(cmbPayer.Value)
        Me.FilterOn = True
    End If
    
    If IsSubForm(Me) = True Then
        Set oFrm = Me.Parent
        oFrm.SelectedPayerNameId = Nz(Me.cmbPayer.Value, 1000)
    End If
Block_Exit:
    Set oFrm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

Private Sub cmbPayer_Change()
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".cmbPayer_Change"
    
        '' Need to filter or unfilter tagged claims
    Call FilterByPayer
'    Call LCDListRefresh
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub


Private Sub cmdSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdSave_Click"
    
    ' if something has changed...
    '
    If cbDirty = False Then
        MsgBox "Nothing seems to have changed.. Please check your values and try again", vbOKOnly
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_Concept_ValidateTab_UpdateOrAdd"
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.FormConceptID
        .Parameters("@pPayerNameId") = Nz(Me.cmbPayer, 1000)
        .Parameters("@pApproved") = Me.fraApprovedStatus
        .Parameters("@pStatusChange") = Me.fraStatusChange
        .Parameters("@pSQLChange") = Me.fraSqlChange
        .Parameters("@pContentUpdate") = Me.fraContentUpdate
        .Parameters("@pNotes") = Me.txtNotes
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error saving the details!", .Parameters("@pErrMsg").Value, True, Me.FormConceptID
            GoTo Block_Exit
        End If
    End With
  
    Call RefreshData
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsSubForm(Me) = True Then
        Me.ConceptID.Value = Me.Parent.Form.txtConceptID
        Me.FormConceptID = Me.ConceptID.Value
    End If
End Sub


Private Sub Form_Load()
Dim iAppPermission As Integer
Dim sRecordSource As String
Dim sPayers As String
'
'    Call Account_Check(Me)
'    iAppPermission = UserAccess_Check(Me)
'    If iAppPermission = 0 Then Exit Sub
    If IsSubForm(Me) = True Then
'        Me.ConceptID.Value = Me.Parent.Form.txtConceptID
        Me.FormConceptID = Me.Parent.Form.txtConceptID
    End If
    
'    Me.ConceptID = Me.Parent.Form.txtConceptID

    Call RefreshData
    
    If IsSubForm(Me) = True Then
        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If

'        sRecordSource = "SELECT * from CONCEPT_Lcd WHERE ConceptID = '" & _
'                Me.Parent.Form.txtConceptID & "' "
'
        sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
        If sPayers <> "" Then
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (1000," & sPayers & ") ORDER BY PayerName"
        Else
            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        End If
'
'        Me.RecordSource = sRecordSource
'    Else
'        Me.RecordSource = "SELECT * FROM CONCEPT_Lcd "
'        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
'
'
    End If
End Sub


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sRecordSource As String
Dim sPayers As String
Dim iAppPermission As Integer
Dim oRs As ADODB.RecordSet
Dim oCn As ADODB.Connection
Dim sSql As String
Dim oCtl As Control


    strProcName = ClassName & ".RefreshData"

    Set oCn = New ADODB.Connection
    With oCn
        .CursorLocation = adUseClientBatch
        .ConnectionString = GetConnectString("v_Data_database")
        .Open
    End With

    sSql = "SELECT * FROM CONCEPT_Valdiation WHERE ConceptId = '" & Me.FormConceptID & "' "
    
    If Me.cmbPayer > 1000 Then
        sSql = sSql & " AND PayerNameId = " & CStr(Me.cmbPayer)
    End If

    Set oRs = New ADODB.RecordSet
    With oRs
        .CursorLocation = adUseClientBatch
        .LockType = adLockBatchOptimistic
        .ActiveConnection = oCn
        .Open sSql
    End With

    ' detach:
    Set oRs.ActiveConnection = Nothing
    
    If oRs.recordCount > 0 Then
        For Each oCtl In Me.Controls
            If oCtl.Tag <> "" Then
                oCtl = oRs(oCtl.Properties("Tag").Value).Value
            End If
        Next
    
    End If
    
    Call FilterByPayer
    RefreshData = True
    cbDirty = False
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
'
'Private Sub ReportDt_DblClick(Cancel As Integer)
'On Error GoTo Block_Err
'Dim strProcName As String
'
'    strProcName = ClassName & ".ReportDt_DblClick"
'
'    DoCmd.OpenForm "frm_REPORT_Main", acNormal, , , , , "Report = R1023"
'
'
'Block_Exit:
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    If cbDirty = True Then
        If MsgBox("Do you wish to save your changes first?", vbYesNo, "Save Next Contract changes?") = vbYes Then
            cmdSave_Click
        End If
    End If
End Sub

Private Sub fraNextContract_AfterUpdate()
    cbDirty = True
End Sub

Private Sub fraNextContract_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cbDirty = True
End Sub

Private Sub fraApprovedStatus_AfterUpdate()
    cbDirty = True
End Sub


Private Sub fraContentUpdate_AfterUpdate()
    cbDirty = True
End Sub


Private Sub fraSqlChange_AfterUpdate()
    cbDirty = True
End Sub

Private Sub fraStatusChange_AfterUpdate()
    cbDirty = True
End Sub



Private Sub txtNotes_Change()
    cbDirty = True
End Sub
