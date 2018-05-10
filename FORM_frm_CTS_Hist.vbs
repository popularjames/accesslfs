Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private rstAuditClm_Dtl As ADODB.RecordSet
Private rstAuditClm_Dtl_Clone As ADODB.RecordSet
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Const CstrFrmAppID As String = "AuditClmDtl"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Set DtlRecordSource(data As ADODB.RecordSet)
    Set rstAuditClm_Dtl = data
    Set Me.RecordSet = data
End Property

Property Get DtlRecordSource() As ADODB.RecordSet
     Set DtlRecordSource = rstAuditClm_Dtl
End Property

Public Sub RefreshData()
    
    Exit Sub

exitHere:
    Exit Sub

ErrHandler:
    'strError = Err.Description
    'MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub
'
'Private Sub Adj_HCPCS_AfterUpdate()
'    If Nz(Adj_HCPCS, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_HCPCS = Me!Adj_HCPCS.OldValue
'    End If
'End Sub
'
'Private Sub Adj_ProjectedSavings_AfterUpdate()
'    If Nz(Adj_ProjectedSavings, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_ProjectedSavings = Me!Adj_ProjectedSavings.OldValue
'    End If
'End Sub
'
'Private Sub Adj_Mod01_AfterUpdate()
'    If Nz(Adj_Mod01, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_Mod01 = Me!Adj_Mod01.OldValue
'    End If
'End Sub
'Private Sub Adj_Mod02_AfterUpdate()
'    If Nz(Adj_Mod02, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_Mod02 = Me!Adj_Mod02.OldValue
'    End If
'End Sub
'Private Sub Adj_Mod03_AfterUpdate()
'    If Nz(Adj_Mod03, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_Mod03 = Me!Adj_Mod03.OldValue
'    End If
'End Sub
'Private Sub Adj_Mod04_AfterUpdate()
'    If Nz(Adj_Mod04, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_Mod04 = Me!Adj_Mod04.OldValue
'    End If
'End Sub
'Private Sub Adj_Mod05_AfterUpdate()
'    If Nz(Adj_Mod05, "") <> "" And Nz(Me.cmoIndicator, "") <> "Y" Then
'        MsgBox "Cannot adjust this line without Adj_Indicator as Y", vbCritical
'        Me!Adj_Mod05 = Me!Adj_Mod05.OldValue
'    End If
'End Sub
'
'
'
'Private Sub Adj_Units_BeforeUpdate(Cancel As Integer)
'    If Me.cmoIndicator <> "Y" Then
'        MsgBox "Cannot Adjust until Adj_Ind=Y", vbCritical
'        Me.Undo
'        Exit Sub
'    End If
'
'    If MsgBox("Recalculate Overpayment?", vbYesNo) = vbYes Then
'        If Nz(Me.Units > 0) Then
'            Me.Adj_ProjectedSavings = Nz(Me.LnReimbAmt, 0) - ((Nz(Me.LnReimbAmt, 0) / Nz(Me.Units, 0)) * Nz(Me.Adj_Units, 0))
'        End If
'    End If
'End Sub
'
'Private Sub cboConceptCd_BeforeUpdate(Cancel As Integer)
'
'    If Me.cmoIndicator <> "Y" Then
'        MsgBox "Cannot Adjust until Adj_Ind=Y", vbCritical
'        Me.Undo
'        Exit Sub
'    End If
'
'    'Clear Header Concept CD if it's being added at the Line Level.
'
'    If Me.Parent.Controls("Adj_ConceptID").Value <> "" Then
'        If MsgBox("Changing the Concept Code at the detail level will clear header level codes." & vbCr & vbCr & "Do you want to continue?", vbYesNo) = vbYes Then
'            Me.Parent.RecordSet.Fields("Adj_ConceptId").Value = ""
'            Me.Parent.RecordSet("Adj_ProjectedSavings").Value = Null
'            Me.Parent.RecordSet.UpdateBatch adAffectAllChapters
'        Else
'           Me.cboConceptCd.Undo
'           Cancel = True
'        End If
'    End If
'End Sub


'DPR - 032610
'FOR OUTPATIENT COMPLEX REVIEW CLAIMS
'NEED TO MAKE IT SO YOU CAN INDICATE A LINE, BUT NOT CHANGE THE CONCEPT
'Private Sub cmoIndicator_BeforeUpdate(Cancel As Integer)
'
'    'Clear Header Concept CD if it's being added at the Line Level.''
'
'    If Me.Parent.Controls("Adj_ConceptID").Value <> "" And Nz(Me.Adj_Ind, "") = "y" Then
'        If MsgBox("Changing the Concept Code at the detail level will clear header level codes." & vbCr & vbCr & "Do you want to continue?", vbYesNo) = vbYes Then
'            Me.Parent.Recordset.Fields("Adj_ConceptId").Value = ""
'            Me.Parent.Recordset("Adj_ProjectedSavings").Value = Null
'            Me.Parent.Recordset.UpdateBatch adAffectAllChapters
'        Else
'           Me.cboConceptCd.Undo
'           Cancel = True
'        End If
'    End If
'End Sub

'Private Sub Form_AfterUpdate()
'
'    '****Calculate total for header
'    Dim curClaimTotal As Currency
'    Dim rst As ADODB.RecordSet
'
'    Set rst = Me.RecordsetClone '*JC using the clone so we're not changing the current row in the form.
'
'    rst.MoveFirst
'
'    Do While rst.EOF = False
'        rst.Find "Adj_ind <> ''", , adSearchForward
'
'        If rst.EOF = False Then
'            curClaimTotal = curClaimTotal + Nz(rst("adj_ProjectedSavings"), 0)
'            rst.MoveNext
'        Else
'            Exit Do
'        End If
'    Loop
'
'    Me.Parent.RecordSet("Adj_ProjectedSavings").Value = curClaimTotal
'    Me.Parent.RecordSet("Adj_ReimbAmt").Value = Me.Parent.RecordSet("ReimbAmt").Value - curClaimTotal
'    Me.Parent.RecordSet.UpdateBatch adAffectAllChapters
'
'    Set rstAuditClm_Dtl_Clone = Me.RecordsetClone
'
'exitHere:
'    Set rst = Nothing
'End Sub


'
'Private Sub Form_Dirty(Cancel As Integer)
'    If IsSubForm(Me) Then
'        Me.Parent.RecordChanged = True
'    End If
'End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    'iAppPermission = UserAccess_Check(Me)
End Sub
