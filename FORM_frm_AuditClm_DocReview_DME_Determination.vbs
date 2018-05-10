Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private strCnlyClaimNum As String
Const CstrFrmAppID As String = "DMEDetermination"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property


Public Sub DMEDetermatination_RefreshData(strCnlyClaimNum)

    Me.Repaint

    Dim rsDMEDetermination As ADODB.RecordSet
    Dim MyAdo As clsADO
    'Dim strCnlyClaimNum As String
    
    'strCnlyclaimNum = Me.Parent.txtCnlyClaimNum
    
    Set Me.RecordSet = Nothing
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
   
       'BEGIN  3/10/2014 KCF - Which determination to choose
   If Me.Parent.Name = "frm_AuditClm_DocREview_DME" Then
        MyAdo.sqlString = "select * from CMS_AUDITORS_CODE.dbo.udf_AuditClm_DocReview_DME_Determination('" & strCnlyClaimNum & "')"
    ElseIf Me.Parent.Name = "Frm_AuditClm_RulesEngine" Then
        MyAdo.sqlString = "Select * from cms_Auditors_Code.dbo.udf_AuditClm_RulesEngine_Determination('" & strCnlyClaimNum & "')"
    End If

    Set rsDMEDetermination = MyAdo.OpenRecordSet
    
    If Not (rsDMEDetermination.EOF Or rsDMEDetermination.BOF) Then
        Set Me.RecordSet = rsDMEDetermination
        Me.txtDetermination.visible = True
        'Me.chkDetermination_Check.visible = True
        Me.txtDetermination_Desc.visible = True
        Me.txtDetermination_Title = "Final Determination"
    Else
        Me.txtDetermination_Title = "There is not enough information"
    End If
    
    MyAdo.DisConnect
    
    Set MyAdo = Nothing
    
End Sub

Private Sub Form_Close()
    Me.Form.RecordSource = ""
End Sub

Private Sub Form_Load()
DMEDetermatination_RefreshData (strCnlyClaimNum)
'MsgBox (strCnlyClaimNum)
End Sub

Private Sub Form_Open(Cancel As Integer)
DMEDetermatination_RefreshData (strCnlyClaimNum)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Me.Form.RecordSource = ""
End Sub
