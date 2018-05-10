Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private coRs As ADODB.RecordSet
Private csConceptId As String
Private clPayerNameId As Long
Private cbIsDirty As Boolean

Private WithEvents coParentForm As Form_frm_CONCEPT_Hdr
Attribute coParentForm.VB_VarHelpID = -1

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = cbIsDirty
End Property
Public Property Let IsDirty(bIsDirty As Boolean)
    cbIsDirty = bIsDirty
    If bIsDirty = True Then
        coParentForm.RecordChanged = True
    End If
End Property


Public Property Get FormPayerNameID() As Long

    FormPayerNameID = clPayerNameId
End Property
Public Property Let FormPayerNameID(lPayerNameId As Long)
    clPayerNameId = lPayerNameId
End Property


Public Property Get FormConceptID() As String
    FormConceptID = IdValue
End Property
Public Property Let FormConceptID(sConceptId As String)
    IdValue = sConceptId
End Property

' frmAppID
Public Property Get frmAppID() As String
    frmAppID = 1
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


Private Sub cmdAddNote_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdAddNote_Click"
    
    If coRs Is Nothing Then
       Stop
    End If
    
    coRs.AddNew
    
    coRs("ConceptId") = Me.FormConceptID
    coRs("PayerNameId") = Me.FormPayerNameID
    coRs("ModifyDt") = Now()
    coRs("ModifyUserName") = Identity.UserName
    coRs("Notes") = Me.txtAddNoteText
    
    If Nz(Me.ckAddApproved, -2) = -2 Then
        Stop ' nulls are good today, not going to update anything
        ' well, we are going to take the latest value I think and carry that forward
    Else
        Select Case Me.ckAddApproved
        Case 0  ' not approved
            coRs("Approved") = 0
            coRs("ApprovalUser") = Identity.UserName
'            coRS("ApprovalDate") = vbNull
            
        Case Else   ' approved
            coRs("Approved") = 1
            coRs("ApprovalUser") = Identity.UserName
            coRs("ApprovalDate") = Now()
            
        End Select
    
    
    End If
    coRs("ConceptStatusChange") = IIf(Me.ckAddStatusChg, 1, 0)
    coRs("SQLChangeRequired") = IIf(Me.ckAddSqlChange, 1, 0)
    coRs("ContentUpdate") = IIf(Me.ckAddContentUpdate, 1, 0)
    
    
    coRs.Update
    coRs.MoveLast
    
    Me.txtAddNoteText.SetFocus
    Me.txtAddNoteText = ""
    
'    Me.IsDirty = True
    Call SaveData
'    Me.RecordSource = ""
'    Set Me.Recordset = coRs
'    MsgBox "Your note will be saved and will appear when you save the concept", vbInformation, "Will be saved when you save the concept!"
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub coParentForm_ConceptIdChanged(sNewConceptId As String)
    Me.FormConceptID = sNewConceptId
    Call RefreshData

End Sub

Private Sub coParentForm_PayerNameIdChanged(lNewPayerNameId As Long)
    Me.FormPayerNameID = lNewPayerNameId
    Call RefreshData
End Sub

Private Sub Form_Close()
    Set coParentForm = Nothing
    If Not coRs Is Nothing Then
        If coRs.State = adStateOpen Then coRs.Close
        Set coRs = Nothing
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
    
    If IsSubForm(Me) = True Then
        Set coParentForm = Me.Parent.Form

        Me.FormConceptID = coParentForm.FormConceptID
        Me.FormPayerNameID = coParentForm.PayerNameId
    End If
    
    RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function SaveData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection

    strProcName = ClassName & ".SaveData"
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_Data_Database")
        .CursorLocation = adUseClientBatch
        .Open
    End With
    
    Set coRs.ActiveConnection = oCn
    coRs.UpdateBatch
    
    Set coRs.ActiveConnection = Nothing
    
    
    SaveData = True
    
    RefreshData
Block_Exit:
    If Not oCn Is Nothing Then
        If oCn.State = adStateOpen Then oCn.Close
        Set oCn = Nothing
    End If
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".RefreshData"
    
    sSql = "SELECT ValidationId, ConceptId, PayerNameId, Approved, ApprovalUser, ApprovalDate, ConceptStatusChange, SQLChangeRequired, ContentUpdate, Notes, ModifyUserName, ModifyDt FROM CONCEPT_Validation"
    
    If Me.FormConceptID <> "" Then
        sSql = sSql & " WHERE ConceptId = '" & Me.FormConceptID & "' "
        If Me.FormPayerNameID <> 0 Then
            If coParentForm.IsPayerSetToAll = False Then
                sSql = sSql & " AND PayerNameId = " & CStr(Me.FormPayerNameID)
            End If
        End If
    End If
    
    sSql = sSql & " ORDER BY ModifyDt DESC "
    
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_Data_Database")
        .CursorLocation = adUseClientBatch
        .Open
    End With
    
    If Not coRs Is Nothing Then
        If coRs.State = adStateOpen Then coRs.Close
        Set coRs = Nothing
    End If
    
    Set coRs = New ADODB.RecordSet
    With coRs
        Set .ActiveConnection = oCn
        .Open sSql, oCn, adOpenUnspecified, adLockBatchOptimistic
        Set .ActiveConnection = Nothing
    End With
    
    Me.RecordSource = ""
    Set Me.RecordSet = coRs
    
    RefreshData = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub Form_Resize()
'    ResizeControls Me.Form
End Sub
