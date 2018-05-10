Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrCnlyClaimNum As String
Private mrsAuditClmHdrAdditionalInfo As ADODB.RecordSet

Const CstrFrmAppID As String = "AuditClm"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Set AdditionalHdrInfoRecordSource(data As ADODB.RecordSet)
    Set mrsAuditClmHdrAdditionalInfo = data
    Set Me.RecordSet = data
End Property


Private Sub Form_AfterUpdate()
    Me.RefreshData
End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
    If IsSubForm(Me) Then
        Me.CnlyClaimNum = Me.Parent.CnlyClaimNum
    End If
End Sub


Private Sub Form_Close()
    DoCmd.SetWarnings True
End Sub

Private Sub Form_Delete(Cancel As Integer)
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
        If IsSubForm(Me) Then
            Me.AllowAdditions = True
            Me.Parent.RecordChanged = True
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    iAppPermission = UserAccess_Check(Me)

    If IsSubForm(Me) Then
        Me.CnlyClaimNum.ColumnHidden = True
    Else
        Me.CnlyClaimNum.ColumnHidden = False
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.SetWarnings False
    Me.RefreshData
End Sub

Private Sub PayeeNum_AfterUpdate()
    Me.PayeeName = Me.PayeeNum.Column(1, Me.PayeeNum.ListIndex)
End Sub


Private Sub Form_Dirty(Cancel As Integer)
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub

Public Sub RefreshData()
    If IsSubForm(Me) Then
        If Not (mrsAuditClmHdrAdditionalInfo Is Nothing) Then
            If Me.RecordSet.recordCount >= 1 Then
                Me.AllowAdditions = False
            Else
                Me.AllowAdditions = True
            End If
        End If
    End If

exitHere:
    Exit Sub

ErrHandler:
    'strError = Err.Description
    'MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub
