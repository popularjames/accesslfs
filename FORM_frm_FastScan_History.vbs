Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "FastScanHistory"

Public Event RefreshScreen()

Private mstrCoverSheetNum As String


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let OpenCoverSheetNum(data As String)
     mstrCoverSheetNum = data
End Property

Property Get OpenCoverSheetNum() As String
     OpenCoverSheetNum = mstrCoverSheetNum
End Property


Public Sub RefreshScreen()

Dim strError As String

    Dim MyAdo As clsADO
    Dim strSQL As String
    Dim rst As ADODB.RecordSet
    
    
    
On Error GoTo ErrHandler
    
    'Me.cnlyclaimnum = "072787446970000012265570118003"
   
    'loading data on the main form
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")
    
    strSQL = " SELECT * from FastScanMaint.v_FastScan_History_v2 where ResultLevel = 'H' "
    strSQL = strSQL & " and CoverSheetNum = '" & Me.OpenCoverSheetNum & "'"
    
    MyAdo.sqlString = strSQL
    Set rst = MyAdo.OpenRecordSet
    
    Set Me.RecordSet = rst

    MyAdo.DisConnect
    Set MyAdo = Nothing
    
    
    'loading data in the subform
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_GlobalCnlyMCR_Database")

    strSQL = " SELECT * from FastScanMaint.v_FastScan_History_V2 where ResultLevel = 'D' "
    strSQL = strSQL & " and CoverSheetNum = '" & Me.OpenCoverSheetNum & "'"
    strSQL = strSQL & " order by LogDate, CoverSheetNum, SplitCoverSheetNum "

    MyAdo.sqlString = strSQL
    Set rst = MyAdo.OpenRecordSet

    Set Me.subfrm_FastScan_History_Results.Form.RecordSet = rst

    MyAdo.DisConnect
    Set MyAdo = Nothing

    Me.Requery
    Me.subfrm_FastScan_History_Results.Form.Requery
    
    If Not Me.subfrm_FastScan_History_Results.Form.RecordSet Is Nothing Then
        If Not (Me.subfrm_FastScan_History_Results.Form.RecordSet.EOF And Me.subfrm_FastScan_History_Results.Form.RecordSet.BOF) Then
            Me.subfrm_FastScan_History_Results.Form.RecordSet.MoveLast
        End If
    End If

    
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
    
End Sub





Private Sub cmdOk_Click()
    
    
    DoCmd.Close acForm, Me.Name

    
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    
    
    Me.Caption = "FastScan History"
    
'    Dim iAppPermission As Integer
'
'    Call Account_Check(Me)
'    iAppPermission = UserAccess_Check(Me)
    
        
End Sub
