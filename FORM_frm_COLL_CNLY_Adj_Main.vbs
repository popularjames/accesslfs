Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_COLL_CNLY_Adj_Main
' Author:      Damon
' Create Date:
' Description:
'      Maintain Connolly manual adjustments for the given claim. List the current
' Connolly Adjustments.
'
'
' Modification History:
'   2012-12-26 by BJD to add additional business rules for initial deployment.
'
'
' =============================================

Const CstrFrmAppID As String = "LdgrCnlyM"
Private miAppPermission As Integer
Private mbAllowChange As Boolean
Private mbAllowAdd As Boolean
Private mbAllowDelete As Boolean

Private WithEvents frmCOLLCNLYAdj As Form_frm_COLL_CNLY_Adj
Attribute frmCOLLCNLYAdj.VB_VarHelpID = -1
Private mstrCnlyClaimNum As String
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Public Property Get CnlyClaimNum() As String
    CnlyClaimNum = mstrCnlyClaimNum
End Property
Public Property Let CnlyClaimNum(data As String)
    mstrCnlyClaimNum = data
End Property
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Sub RefreshData()
    On Error GoTo ErrHandler
    Dim strSQL As String
    Dim rst As ADODB.RecordSet
    
    'Creating a new instance of ADO-class variable
    Set MyAdo = New clsADO
    
    'Making a Connection call to SQL database?
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    'Setting a string to a SQL query statement, depending on ID
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM COLL_CNLY_Adj "
    strSQL = strSQL & " where cnlyCLaimNum = '" & Me.CnlyClaimNum & "'"
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    MyAdo.sqlString = strSQL
    
    Set Me.lstManualCollections.RecordSet = Nothing
    Me.lstManualCollections.RowSource = vbNullString
    'Set our listbox columns to be the same as letter_selection_temp
    Me.lstManualCollections.ColumnCount = CurrentDb.TableDefs("COLL_CNLY_Adj").Fields.Count
    'Setting the list record set equal to the specify ADO-class record set
    Set Me.lstManualCollections.RecordSet = MyAdo.OpenRecordSet()
  
    Me.lblAppTitle.Caption = "MANUAL ADJUSTMENTS: " & Nz(mstrCnlyClaimNum, "")
       
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Sub


Private Sub cmdNew_Click()

    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
  
    If miAppPermission = 0 Then Exit Sub
    
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    'mbAllowDelete = (miAppPermission And gcAllowDelete)
    
    If ((mbAllowChange = False) And (mbAllowAdd = False)) Then
        MsgBox "You do not have permission to add or update records", vbOKOnly + vbInformation, "Connolly Adjustment Security Info"
        Exit Sub
    ElseIf mbAllowAdd = False Then
        MsgBox "You do not have permission to add records", vbOKOnly + vbInformation, "Connolly Adjustment Security Info"
        Exit Sub
    End If

    Set frmCOLLCNLYAdj = New Form_frm_COLL_CNLY_Adj
    
    frmCOLLCNLYAdj.Insert = True
    frmCOLLCNLYAdj.FormCnlyClaimNum = Me.CnlyClaimNum
    frmCOLLCNLYAdj.RefreshData
    ShowFormAndWait frmCOLLCNLYAdj
    Set frmCOLLCNLYAdj = Nothing
    Me.RefreshData
End Sub



Private Sub lstManualCollections_DblClick(Cancel As Integer)
  
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
  
    If miAppPermission = 0 Then Exit Sub

    Set frmCOLLCNLYAdj = New Form_frm_COLL_CNLY_Adj
       
    frmCOLLCNLYAdj.FormCnlyClaimNum = ""
    frmCOLLCNLYAdj.FormCnlyARCollID = Me.lstManualCollections.Column(GetColumnPosition(Me.lstManualCollections, "CnlyARCollID"))
    
    frmCOLLCNLYAdj.RefreshData
    
    ShowFormAndWait frmCOLLCNLYAdj
    Set frmCOLLCNLYAdj = Nothing
    
    Me.RefreshData
 
End Sub
