Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "ConceptHdr"

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

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
        Me.filter = "PayerNameId = " & CStr(cmbPayer.Value) & " OR NZ(PayerNameId, 1000) = 1000 "
        Me.FilterOn = True
    End If
    
    
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

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me.ConceptID.Value = Me.Parent.Form.txtConceptID
End Sub

'Public Property Get PayerNameId() As Integer
'Dim oParFrm As Form_frm_CONCEPT_Main
'
'    If IsSubForm(Me) = True Then
'
'    End If
'End Property

Private Sub Form_Load()
Dim iAppPermission As Integer
Dim sRecordSource As String
Dim sPayers As String

    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    Me.txtSelectedId = Me.Parent.Form.txtConceptID
    
    sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
            
    If IsSubForm(Me) = True Then
        'Me.ConceptID.ColumnHidden = True
        sRecordSource = "SELECT ConceptId, PayerNameID, CodeTypeId, Code, Reference, Comments from CONCEPT_Dtl_Codes WHERE ConceptID = '" & Me.Parent.Form.txtConceptID & "' "

        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If
        If sPayers <> "" Then
            sPayers = "1000," & sPayers
            sRecordSource = sRecordSource & " AND (PayerNameID IN (" & sPayers & ") OR PayerNameID IS NULL ) "
            Me.cmbPayerNameId.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (" & sPayers & ") ORDER BY PayerName"
            Me.cmbPayer.RowSource = Me.cmbPayerNameId.RowSource
        
        Else
            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
            Me.cmbPayerNameId.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
            Me.cmbPayer.RowSource = Me.cmbPayerNameId.RowSource
        End If
        
        Me.RecordSource = sRecordSource
        
            
    Else
        
        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If

    
        Me.RecordSource = "SELECT  ConceptId, PayerNameID, CodeTypeId, Code, Reference, Comments from CONCEPT_Dtl_Codes "
        Me.cmbPayerNameId.RowSource = " SELECT PAyerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
    
    
End Sub
