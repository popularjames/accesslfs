Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const CstrFrmAppID As String = "ConceptStatusHist"
Private csConceptId As String
Private cblnFilterApplied As Boolean

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get FormConceptID() As String
    FormConceptID = IdValue
End Property
Public Property Let FormConceptID(sConceptId As String)
    IdValue = sConceptId
    Me.txtSelectedId = sConceptId
End Property


Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    Me.txtSelectedId = csConceptId
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
    
'    Call ErrorCodeListRefresh
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

'
'
'Private Sub cmdAddErrorCode_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'
'    strProcName = ClassName & ".cmdAddErrorCode_Click"
'
''Me.frmConceptID = "CM_C0834"
'
'
'    '' Here we need to insert the ErrorCode and the states for the given payer.
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_Code_Database")
'        .SQLTextType = StoredProc
'        .sqlString = "usp_CONMGNT_Add_ErrorCode_To_Concept"
'        .Parameters.Refresh
'        .Parameters("@pConceptId") = Me.FormConceptID
'        If Nz(Me.cmbPayer.Value, 999) <> 999 Then
'            .Parameters("@pPayerNameId") = Me.cmbPayer.Value
'        Else
''Stop
'        End If
'        .Parameters("@pErrorCode") = Me.cmbErrorCodes
'
'        .Execute
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            LogMessage strProcName, "ERROR", "There was a problem when trying to insert the Error Code data!", .Parameters("@pErrMsg").Value, True, Me.ConceptId
'        Else
'            Call RefreshData
'        End If
'    End With
'
'
'Block_Exit:
'    Set oAdo = Nothing
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub
'
'Private Sub cmdDel_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'
'    strProcName = ClassName & ".cmdDel_Click"
'
'    Debug.Print "Error Code to delete is: " & Me.ErrorCodeID
'    Debug.Print "for concept " & Me.ConceptId
'    Debug.Print "Payer: " & Me.PayerNameID
'
'
'
'    '' Here we need to insert the ErrorCode for the given payer.
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_Code_Database")
'        .SQLTextType = StoredProc
'        .sqlString = "usp_CONMGNT_Del_ErrorCode_From_Concept"
'        .Parameters.Refresh
'        .Parameters("@pConceptId") = Me.FormConceptID
'        If Nz(Me.cmbPayer.Value, 999) <> 999 Then
'            .Parameters("@pPayerNameId") = Me.cmbPayer.Value
'        End If
'        .Parameters("@pErrorCode") = Me.ErrorCodeID
'
'        .Execute
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            LogMessage strProcName, "ERROR", "There was a problem when trying to delete the Error Code: " & Me.ErrorCodeID & " data!", .Parameters("@pErrMsg").Value, True, Me.ConceptId
'        Else
'            Call RefreshData
'        End If
'    End With
'
'
'Block_Exit:
'    Set oAdo = Nothing
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

'Private Sub cmdRelease_Click()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'
'    strProcName = ClassName & ".cmdRelease_Click"
'
'    Debug.Print "LCD to delete is: " & Me.ErrorCodeID
'    Debug.Print "for concept " & Me.ConceptId
'    Debug.Print "Payer: " & Me.PayerNameID
'
''Stop
'
'    '' Here we need to insert the LCD and the states for the given payer.
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_Code_Database")
'        .SQLTextType = StoredProc
'        .sqlString = "usp_CONMGNT_LCD_RelaseChgFlag"
'        .Parameters.Refresh
'        .Parameters("@pConceptId") = Me.FormConceptID
'        If Nz(Me.PayerNameID.Value, 999) <> 999 Then
'            .Parameters("@pPayerNameId") = Me.PayerNameID.Value
'        End If
'        .Parameters("@pLCD") = Me.ErrorCodeID
'
'        .Execute
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            LogMessage strProcName, "ERROR", "There was a problem when trying to release the LCD: " & Me.ErrorCodeID, .Parameters("@pErrMsg").Value, True, Me.ConceptId
'        Else
'            Call RefreshData
'        End If
'    End With
'
'
'Block_Exit:
'    Set oAdo = Nothing
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub

Private Sub cmdRelease_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.cmdRelease.StatusBarText = "Last released: " & Nz(Me.ReleaseFlagDt, "(not yet)")
'    Debug.Print Me.cmdRelease.TabIndex
    ' this doesn't work because the correct record needs to be selected.. i.e. it will only work if that row in the form is selected
'    Me.cmdRelease.ControlTipText = "Last released: " & Nz(Me.ReleaseFlagDt, "(not yet)")
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
    
'    If IsSubForm(Me) = True Then
'        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
'        If Trim(lblPayersNote.Caption) = "" Then
'            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
'        End If
'
'        sRecordSource = "SELECT * from CONCEPT_Lcd WHERE ConceptID = '" & _
'                Me.Parent.Form.txtConceptID & "' "
'
'        sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
'        If sPayers <> "" Then
'            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (1000," & sPayers & ") ORDER BY PayerName"
'        Else
'            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
'            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
'        End If
'
'        Me.RecordSource = sRecordSource
'    Else
'        Me.RecordSource = "SELECT * FROM CONCEPT_Lcd "
'        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
'
'
'    End If
End Sub

'Private Function ErrorCodeListRefresh()
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sSql As String
'
'    strProcName = ClassName & ".ErrorCodeListRefresh"
'
'    sSql = "SELECT * FROM v_CONCEPT_ErrorCodes_SelectList"
'
'    If Nz(cmbPayer.Value, 1000) <> 1000 Then
'        sSql = sSql & " WHERE PayerNameId = " & CStr(cmbPayer.Value)
'    Else
'
'    End If
'
'    Me.cmbErrorCodes.RowSource = sSql
'    Me.cmbErrorCodes.Requery
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sRecordSource As String
Dim sPayers As String
Dim iAppPermission As Integer

    strProcName = ClassName & ".RefreshData"


    
    If IsSubForm(Me) = True Then
        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If
        
        sRecordSource = "SELECT * FROM v_CONCEPT_StatusHist WHERE ConceptID = '" & _
                Me.Parent.Form.txtConceptID & "' ORDER BY StatusChangeDt DESC "
                
        sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
        If sPayers <> "" Then
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (1000," & sPayers & ") ORDER BY PayerName"
        Else
            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        End If
    
        Me.RecordSource = sRecordSource
    Else
        Me.RecordSource = "SELECT * FROM v_CONCEPT_StatusHist ORDER BY StatusChangeDt DESC "
        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
    End If
    
    Call FilterByPayer
'    Call ErrorCodeListRefresh
    RefreshData = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub ReportDt_DblClick(Cancel As Integer)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".ReportDt_DblClick"
    
    DoCmd.OpenForm "frm_REPORT_Main", acNormal, , , , , "Report = R1023"

    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
