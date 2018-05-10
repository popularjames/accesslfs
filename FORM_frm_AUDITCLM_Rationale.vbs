Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AuditClmRationale"

Private rsAuditClmHdr As ADODB.RecordSet

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Set HdrRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmHdr = data
End Property

Private Sub Adj_Rationale_AfterUpdate()
    rsAuditClmHdr.Fields("Adj_Rationale") = Adj_Rationale
    Me.Parent.RecordChanged = True
End Sub

Public Sub RefreshData()
    Adj_Rationale = rsAuditClmHdr.Fields("Adj_Rationale")
End Sub
