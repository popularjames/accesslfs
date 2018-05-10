Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private strCnlyClaimNum As String
Private intAuditID As Integer
Private rsAuditClmClaimsPlus As ADODB.RecordSet

Const CstrFrmAppID As String = "AuditClmPlus"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Set ClaimsPlusRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmClaimsPlus = data
End Property

Property Get ClaimsPlusRecordsource() As ADODB.RecordSet
     Set ClaimsPlusRecordsource = rsAuditClmClaimsPlus
End Property

'Main property of form.  This drives everything that this object is based on
Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property

Private Sub ClaimSrc_AfterUpdate()
    'SaveData
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("CnlyClaimNum") = Me.CnlyClaimNum
    End If
    
    rsAuditClmClaimsPlus("ClaimSrc") = Me.ClaimSrc
    rsAuditClmClaimsPlus("ClaimSrcRootTxt") = Me.ClaimSrc.Column(2)
    rsAuditClmClaimsPlus("ClaimSrcTxt") = Me.ClaimSrc.Column(1)
    
    FormIsDirty
End Sub

Private Sub cmbProject_AfterUpdate()
    'SaveData
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("cnlyClaimNum") = Me.CnlyClaimNum
    End If
    rsAuditClmClaimsPlus("ProjectID") = Me.cmbProject
    FormIsDirty
End Sub

Private Sub Form_Close()
    'SaveData
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Claim Plus Data Entry From"
    
    iAppPermission = UserAccess_Check(Me)
End Sub

Public Sub RefreshData()
    RefreshConnollyReason
    RefreshClientReason
    RefreshRootCause
    RefreshClaimSource
    RefreshProject
    
    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
       rsAuditClmClaimsPlus.MoveFirst
        Me.pwTranAmt = Nz(rsAuditClmClaimsPlus("GrossAmt"), 0)
    End If
    
End Sub

Private Sub FormIsDirty()
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub

Private Sub RefreshConnollyReason()
    
    'TO DO - Damon Add Error Handling
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    
    intAuditID = DLookup("Adj_AuditNum", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'")
    
    strSQL = "SELECT    CC.TranCode,      "
    strSQL = strSQL & " mid(CC.TranCodeText,1,50), "
    strSQL = strSQL & " (CC.TranCode & Space(5) & '(' & mid(CC.TranCodeText,1,50) & ')') AS TranText "
    strSQL = strSQL & " FROM vCPpClaimsCodes CC LEFT JOIN (select visibilityid, trancode, trantype  from vCpuAuditsClaimCodeVisibility where NZ(auditid, 0) = " & intAuditID
    strSQL = strSQL & " and trantype = 0 ) CV ON CC.TranCode = CV.TranCode       "
    strSQL = strSQL & " AND CC.TranType = CV.TranType where     CV.VisibilityID is NULL And CC.TranType=0 "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
    If Not rst.EOF Then
     rst.MoveLast
     rst.MoveFirst
     Set Me.TranCode.RecordSet = rst
    End If
    
    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.MoveFirst
        Me.TranCode = Me.ClaimsPlusRecordsource.Fields("ClaimCode")
    End If
    
End Sub

Private Sub RefreshClientReason()
    
    'TO DO - Damon Add Error Handling
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    
    intAuditID = DLookup("Adj_AuditNum", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'")
    
    strSQL = " SELECT   CLTC.TranCode, CLTC.TranCodeText, (CLTC.TranCode & space(3) & '(' & CLTC.TranCodeText & ')') As TranText "
    strSQL = strSQL & " FROM vCpuClaimsLedgerTranTypes CLT "
    strSQL = strSQL & " INNER JOIN vCpuClaimsLedgerTranCodesClient CLTC ON CLT.TranType = CLTC.TranType "
    strSQL = strSQL & " Where AuditID = " & intAuditID & " and CLTC.TranType = 0 ORDER BY CLTC.TranCode "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
    
    If Not rst.EOF Then
        rst.MoveLast
        rst.MoveFirst
        Set Me.TranCode.RecordSet = rst
    End If

    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.MoveFirst
        'TODO Assign Correct Value
    End If
End Sub

Private Sub RefreshRootCause()
    
    'TO DO - Damon Add Error Handling
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    
    intAuditID = DLookup("Adj_AuditNum", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'")
    
    strSQL = " SELECT  CC.TranCode, CC.TranCodeText, CC.RootName, (CC.TranCode & Space(5) & '(' & CC.TranCodeText & ')') AS TranText "
    strSQL = strSQL & " FROM vCPpClaimsCodes CC LEFT JOIN (select visibilityid, trancode, trantype from vCPuAuditsClaimCodeVisibility "
    strSQL = strSQL & " where auditid=" & intAuditID & " and trantype=20) CV     ON CC.TranCode = CV.TranCode AND CC.TranType = CV.TranType WHERE ((CC.TranType) = 20) AND (CV.VisibilityID is Null) ORDER BY CC.TranCode "
        
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
    
    If Not rst.EOF Then
        rst.MoveLast
        rst.MoveFirst
        Set Me.RootCause.RecordSet = rst
    End If

    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.MoveFirst
        Me.RootCause = Me.ClaimsPlusRecordsource.Fields("RootCause")
    End If


End Sub

Private Sub RefreshClaimSource()
    
    'TO DO - Damon Add Error Handling
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    
    intAuditID = DLookup("Adj_AuditNum", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'")
    
    strSQL = " SELECT CC.TranCode, mid(CC.TranCodeText,1,50), CC.RootName, (CC.TranCode & Space(5) & '(' & mid(CC.TranCodeText,1,50) & ')') AS TranText "

    strSQL = strSQL & " FROM vCPpClaimsCodes CC "
    strSQL = strSQL & " LEFT JOIN (select visibilityid, trancode, trantype from vCPuAuditsClaimCodeVisibility where auditid=" & intAuditID & " and trantype=10) CV     "
    strSQL = strSQL & " ON CC.TranCode = CV.TranCode AND CC.TranType = CV.TranType WHERE ((CC.TranType) = 10) AND (CV.VisibilityID is Null) "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
   
    If Not rst.EOF Then
        rst.MoveLast
        rst.MoveFirst
        Set Me.ClaimSrc.RecordSet = rst
    End If
    
    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.MoveFirst
        Me.ClaimSrc = Me.ClaimsPlusRecordsource.Fields("ClaimSrc")
    End If
    
End Sub

Private Sub RefreshProject()
    
    'TO DO - Damon Add Error Handling
    Dim strSQL As String
    Dim rst As DAO.RecordSet
    
    intAuditID = DLookup("Adj_AuditNum", "AUDITCLM_Hdr", "cnlyClaimNum = '" & Me.CnlyClaimNum & "'")
    
    strSQL = " SELECT ProjectId, ProjectName, ProjectDesc FROM vCPuAuditsProjects p "
    strSQL = strSQL & "where p.AuditId =" & intAuditID & " "
    strSQL = strSQL & " ORDER BY ProjectName "
        
        
    Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
    
    If Not rst.EOF Then
        rst.MoveLast
        rst.MoveFirst
        Set Me.cmbProject.RecordSet = rst
    End If

    If Not (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        'TODO Assign Correct Value
        rsAuditClmClaimsPlus.MoveFirst
    End If


End Sub

Private Sub pwTranAmt_AfterUpdate()
    'SaveData
    
    If Not IsNumeric(Me.pwTranAmt) Then
        Exit Sub
    End If
    
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("CnlyClaimNum") = Me.CnlyClaimNum
    End If
    rsAuditClmClaimsPlus("GrossAmt") = Me.pwTranAmt
    FormIsDirty
End Sub

Private Sub RootCause_AfterUpdate()
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("CnlyClaimNum") = Me.CnlyClaimNum
    End If
    rsAuditClmClaimsPlus("RootCause") = Me.RootCause
    FormIsDirty
End Sub

Private Sub TranClientCode_AfterUpdate()
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("CnlyClaimNum") = Me.CnlyClaimNum
    End If
    rsAuditClmClaimsPlus("ClaimCodeClient") = Me.TranClientCode
    rsAuditClmClaimsPlus("ClaimCodeClientText") = Me.TranClientCode.Column(1)
    FormIsDirty
End Sub

Private Sub TranCode_AfterUpdate()
    If (rsAuditClmClaimsPlus.BOF And rsAuditClmClaimsPlus.EOF) Then
        rsAuditClmClaimsPlus.AddNew
        rsAuditClmClaimsPlus("CnlyClaimNum") = Me.CnlyClaimNum
    End If
    rsAuditClmClaimsPlus("ClaimCode") = Me.TranCode
    rsAuditClmClaimsPlus("ClaimCodeText") = Me.TranCode.Column(1)
    
    FormIsDirty
End Sub
