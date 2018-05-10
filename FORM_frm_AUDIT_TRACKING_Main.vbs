Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mbDetailFormLoaded As Boolean
Private mstrAuditTableName As String
Private mstrAuditKey As String
Private mstrFrmAppID As String

Public Property Let frmAppID(data As String)
    mstrFrmAppID = data
    Call UserAccess_Check(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = mstrFrmAppID
End Property

Public Sub DetailFormLoaded()
    mbDetailFormLoaded = True
End Sub

Public Property Let AuditTableName(data As String)
    mstrAuditTableName = data
End Property

Public Property Let AppTitle(data As String)
    Me.Caption = data
End Property

Public Property Let AuditKey(data As String)
    mstrAuditKey = data
End Property

Public Sub RefreshData()
    Me.frm_AUDIT_TRACKING_Grid_View.Form.AuditTableName = mstrAuditTableName
    Me.frm_AUDIT_TRACKING_Grid_View.Form.AuditKey = mstrAuditKey
    Me.frm_AUDIT_TRACKING_Grid_View.Form.RefreshData
    RefreshDetail
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Audit Tracking Form"
    
    Call Account_Check(Me)
    
    If Me.frmAppID <> "" Then
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
    
    If IsSubForm(Me) Then
        Me.cmdExit.visible = False
    End If
    Me.frm_AUDIT_TRACKING_Detail_View.SourceObject = "frm_AUDIT_TRACKING_Detail_View"
End Sub

Public Sub RefreshDetail()
    If mbDetailFormLoaded Then
        If txtSQLSource & "" <> "" Then
            Me.frm_AUDIT_TRACKING_Detail_View.Form.RecordSQL = txtSQLSource
            Me.frm_AUDIT_TRACKING_Detail_View.Form.RefreshData
        End If
    End If
End Sub
