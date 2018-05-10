Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Event AutoIdWasPicked(lAutoIdSelected As Long)

Public Event Canceled()

Private cbCanceled As Boolean
Private csConceptId As String
Private ciPayerNameId As Integer
Private clManualEditIdSelected As Long

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get ManualEditIdSelected() As Long
    ManualEditIdSelected = clManualEditIdSelected
End Property
Public Property Let ManualEditIdSelected(lManualEditIdSelected As Long)
    clManualEditIdSelected = lManualEditIdSelected
End Property

Public Property Get PayerNameId() As Integer
    PayerNameId = ciPayerNameId
End Property
Public Property Let PayerNameId(iPayerNameId As Integer)
    ciPayerNameId = iPayerNameId
End Property

Public Property Get Canceled() As Boolean
    Canceled = cbCanceled
End Property
Public Property Let Canceled(bCanceled As Boolean)
    cbCanceled = bCanceled
End Property

Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
End Property



Public Function RefreshData() As Boolean
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
'Me.ConceptID = "CM_C2052"
'Me.PayerNameID = 1008

    sSql = "SELECT ManualEditId, ArchiveDt, EditUser FROM CONCEPT_NIRF_Sample_Edits_Groups " & _
        " WHERE ConceptId = '" & Me.ConceptID & "' AND PayerNameId = " & CStr(Me.PayerNameId) & _
        " ORDER BY ManualEditId DESC"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("CMS_AUDITORS_WORKSPACE")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With

    Set Me.RecordSet = oRs
    
    Me.InsideHeight = 6000

End Function

Private Sub CmdCancel_Click()
    Me.Canceled = True
    Me.visible = False
    RaiseEvent Canceled
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
    Call SelectedRow
End Sub

Private Sub Form_Activate()
Call RefreshData
End Sub

Private Sub Form_Close()
    RemoveObjectInstance Me
End Sub

Private Sub SelectedRow()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SelectedRow"
    
    
' Ok, here, capture the row
    Me.ManualEditIdSelected = Me.txtManualEditId.Value
    
    
    ' now, call our revert sproc:
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_NIRF_Editor_Samples_Revert_To_Id"
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.ConceptID
        .Parameters("@pPayerNameId") = Me.PayerNameId
        .Parameters("@pSelectedManualEditId") = Me.ManualEditIdSelected
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem reverting back to a version of an edited NIRF", .Parameters("@pErrMsg").Value, True, Me.ConceptID
            
        Else
            ' all good to go -
            RaiseEvent AutoIdWasPicked(Me.txtManualEditId.Value)
            Me.visible = False
        End If
    End With
    
    ' and finally raise the event

    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Load()
    Me.InsideHeight = 6000
End Sub

Private Sub Form_ViewChange(ByVal Reason As Long)
Stop
End Sub

Private Sub txtArchiveDt_DblClick(Cancel As Integer)
    Call SelectedRow
End Sub



Private Sub txtEditUser_DblClick(Cancel As Integer)
    Call SelectedRow
End Sub

Private Sub txtManualEditId_DblClick(Cancel As Integer)
    Call SelectedRow
End Sub
