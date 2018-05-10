Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdLoadProviders_Click()
    Me.Refresh
    
    Dim myCode_ADO As New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.sqlString = "usp_PROV_Import_Missing_Providers"
    myCode_ADO.Execute
    
    RefreshData
    
    Set myCode_ADO = Nothing
End Sub

Private Sub cmdRefresh_Click()
    RefreshData
End Sub

Private Sub Form_Load()
    Dim myCode_ADO As New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.sqlString = "usp_PROV_Identify_Missing_Providers"
    myCode_ADO.Execute
    
    RefreshData
    
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    
    Set myCode_ADO = Nothing
    
End Sub

Private Sub RefreshData()
    Me.RecordSource = "select * from PROV_Missing where Processed = 0 order by Priority"
    Me.Requery
End Sub
