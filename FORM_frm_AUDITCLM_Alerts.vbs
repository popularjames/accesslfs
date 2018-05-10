Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130205 KD Fixed, come on guys!  We're SUPPOSED to be programmers right?


Private strCnlyClaimNum As String
Const CstrFrmAppID As String = "Alerts"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property


Public Sub RefreshData()

    Me.Repaint
    

    If Nz(Me.CnlyClaimNum, "") = "" Then 'dont waste time and resources if there is no claim num
        Exit Sub
    End If

    'Dim myCode_Ado As clsADO
    'Dim cmd As ADODB.Command
    'Dim strErrMsg As String
    'Dim iResult As Integer
    Dim rsAlerts As ADODB.RecordSet
    Dim MyAdo As clsADO
    
    Set Me.RecordSet = Nothing
    
    'Set myCode_Ado = New clsADO
    'myCode_Ado.SQLTextType = StoredProc
    'myCode_Ado.ConnectionString = GetConnectString("v_CODE_Database")
    'myCode_Ado.SQLstring = "usp_AUDITCLM_CheckAlerts"
    
    'Set cmd = New ADODB.Command
    'cmd.ActiveConnection = myCode_Ado.CurrentConnection
    'cmd.CommandType = adCmdStoredProc
    'cmd.CommandText = "usp_AUDITCLM_CheckAlerts"
    'cmd.Parameters.Refresh
    
    'cmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
    'cmd.Parameters("@pCurrentUser") = Identity.Username

    'iResult = myCode_Ado.Execute(cmd.Parameters)
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    'MYADO.SQLstring = "select RowNumber= ROW_NUMBER() OVER (ORDER BY Alert_ID), AlertDesc from CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_Alerts where CnlyClaimNum = '" & Me.CnlyClaimNum & "' AND UserName = '" & Identity.Username & "' ORDER BY 1"
    MyAdo.sqlString = "select * from CMS_AUDITORS_CODE.dbo.udf_CheckAlerts('" & Me.CnlyClaimNum & "','" & Identity.UserName & "')"
    Set rsAlerts = MyAdo.OpenRecordSet
    
    If Not (rsAlerts.EOF Or rsAlerts.BOF) Then
        Set Me.RecordSet = rsAlerts
        Me.txtAlert.visible = True
        Me.Alert_ID.visible = True
        Me.txtAlertsTitle = "This claim has " & rsAlerts.recordCount & " alert(s):"
    Else
        Me.txtAlertsTitle = "There are no Alerts for this claim."
    End If
    
    'Set Me.Form.Recordset = myAuditClaim
    MyAdo.DisConnect
    
 
    Set MyAdo = Nothing
    
    DivertFocus.SetFocus

    'If Not (rsAlerts.EOF Or rsAlerts.BOF) Then
        'Me.Recordset = myado.OpenRecordSet
    'End If

    'myCode_Ado.SQLTextType = sqltext
    'myCode_Ado.SQLstring = "exec usp_AUDITCLM_CheckAlerts '" & CurrCnlyClaimNum & "'"
    'iResult = myCode_Ado.Execute
    
End Sub




Private Sub Form_Error(DataErr As Integer, Response As Integer)
    ''' 20130205 KD: So microsoft in their infinite wisdom (ok, so I can't spell)
    ''' decided that when you bind a subform to an ADO recordset, and the main form is
    ''' minimized, the subform's _Unload and then _Close events fire.
    ''' upon restoring (or maximizing) the main form, it's Resize event fires
    ''' then the subform's Form_Error fires with error 3131 (Syntax error in From clause)
    ''' so the below code "eats" the error,
    ''' then we set the main form's timer to fire code to Reload the subform..
    If DataErr = 3131 Then
        Err.Clear
        Response = acDataErrContinue
        If IsSubForm(Me) = True Then
            Me.Parent.Form.Controls("lstTabs").Selected(0) = True
            Me.CnlyClaimNum = Me.Parent.Form.CnlyClaimNum
            Me.Parent.Form.TimerInterval = 500
        End If
    End If
End Sub


Private Sub Form_Open(Cancel As Integer)
    RefreshData
End Sub

Private Sub txtAlertsTitle_DblClick(Cancel As Integer)
    RefreshData
End Sub


 
