Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "ConceptState"

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
        Me.filter = "PayerNameId = " & CStr(cmbPayer.Value)
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

Private Sub Form_Load()
On Error GoTo Block_Err
Dim iAppPermission As Integer
Dim sRecordSource As String
Dim sPayers As String

    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    
    If IsSubForm(Me) = True Then
        lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptID))
        If Trim(lblPayersNote.Caption) = "" Then
            lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
        End If
        
        sRecordSource = "SELECT ConceptID, nz(CONCEPT_Dtl_State.PayerNameID,999) AS PayerNameID, ConceptState, Reference, ClaimCount, ClaimCountSample, " & _
                " ClaimValue, ClaimValueState, Comments from CONCEPT_Dtl_State WHERE ConceptID = '" & _
                Me.Parent.Form.txtConceptID & "' AND ClaimValue <> 0 "
                
        sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
        If sPayers <> "" Then
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (1000," & sPayers & ") ORDER BY PayerName"
        Else
            ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
            Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        End If
    
        Me.RecordSource = sRecordSource
    Else
Block_Err:
        Me.RecordSource = "SELECT * FROM CONCEPT_Dtl_State "
        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
'    Me.RecordsetClone.MoveLast
    
End Sub
