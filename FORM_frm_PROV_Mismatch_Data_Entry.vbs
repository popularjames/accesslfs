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
    myCode_ADO.sqlString = "usp_PROV_Import_Mismatch_Providers"
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
    myCode_ADO.sqlString = "usp_PROV_Identify_Mismatch_Providers"
    myCode_ADO.Execute
    
    RefreshData
    
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    
    Set myCode_ADO = Nothing
    
    If Me.RecordSet.recordCount = 0 Then
        MsgBox "Nothing to be done."
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    
End Sub

Private Sub RefreshData()
    Me.RecordSource = "select * from PROV_Mismatch where Processed = 0 order by Priority"
    Me.Requery
End Sub

Private Sub MailState_AfterUpdate()
    Me.MailState.DefaultValue = Me.MailState.Value & ""
End Sub
