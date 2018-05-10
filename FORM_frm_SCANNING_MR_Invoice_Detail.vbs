Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:           Form_frm_SCANNING_MR_Invoice_Detail
' Author:       Barbara Dyroff
' Date:         2010-02-03
' Description:
'   Display Invoice Details for Medical Record Scanning for a given Provider and Invoice.
'
' Modification History:
'
'20100706 Added ReceivedMeth to display by Rob Hall
' =============================================

Private MyAdo As clsADO

Private strMRInvDtlRecSQL As String
Private rsMRInvDtl As ADODB.RecordSet

'Retrieve the detail data using the ADO Class.
Public Property Let RecordSQL(data As String)
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    strMRInvDtlRecSQL = data
    MyAdo.sqlString = strMRInvDtlRecSQL
    Set rsMRInvDtl = MyAdo.OpenRecordSet
    
    Set MyAdo = Nothing
End Property

Public Sub RefreshData()
    
    Set Me.RecordSet = rsMRInvDtl
    
End Sub

Private Sub Form_Load()
    Me.Parent.DetailFormLoaded
End Sub
