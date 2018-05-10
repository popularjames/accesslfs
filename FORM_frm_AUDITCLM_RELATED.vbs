Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AuditClm"

Private strCnlyClaimNum As String
Property Let CnlyClaimNum(data As String)
     strCnlyClaimNum = data
End Property
Property Get CnlyClaimNum() As String
     CnlyClaimNum = strCnlyClaimNum
End Property




Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
    
End Property
Public Sub RefreshData()
    Dim MyAdo As clsADO
    Dim strSQL As String
    Dim rst As ADODB.RecordSet
    
    
    
On Error GoTo ErrHandler
    
    'Me.cnlyclaimnum = "072787446970000012265570118003"
   
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    strSQL = " SELECT ii.datatype, ii.ICN, dd.linenum, dd.hcpcscd  , dd.units,  dd.RevCd ,  dd.LnClmFromDt , dd.LnClmThruDt , ii.SchedPmtDt , dd.LnReimbAmt , dd.LnAllowedAmt ,  xx.Adj_Ind , xx.Adj_ConceptID  from AUDITCLM_Hdr hh with (nolock)   "
    strSQL = strSQL & " left JOIN CMS_Data_NCH.dbo.INT_HDR ii with(nolock)    "
    strSQL = strSQL & "  on ii.can = hh.CAN  and hh.bic = ii.bic   "
    strSQL = strSQL & " and hh.clmfromDt between ii.ClmFromDt and ii.ClmThruDt  JOIN CMS_Data_NCH.dbo.INT_dtl dd with(nolock)   "
    strSQL = strSQL & " on ii.cnlyCLaimNum = dd.cnlyClaimNum  "
    strSQL = strSQL & " left join AUDITCLM_Dtl xx with(nolock)   on xx.CnlyClaimNum = dd.CnlyClaimNum and xx.LineNum = dd.LineNum  "
    
    strSQL = strSQL & " where hh.CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
    
    MyAdo.sqlString = strSQL
    Set rst = MyAdo.OpenRecordSet
    
    Me.frm_GENERAL_Datasheet_ADO.Form.InitDataADO rst, "AuditClm_Hdr"
    Set Me.frm_GENERAL_Datasheet_ADO.Form.RecordSet = rst

    MyAdo.DisConnect
   Set MyAdo = Nothing


Exit Sub
ErrHandler:
    MsgBox Err.Description


End Sub
