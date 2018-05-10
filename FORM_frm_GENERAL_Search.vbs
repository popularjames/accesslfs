Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frm As Form_frm_GENERAL_Filter
Attribute frm.VB_VarHelpID = -1
Private mvSql As CnlyScreenSQL
Private strGridSource As String
Private strCriteria As String
Private strAppID As String

'Type of search being performed
Property Let frmAppID(data As String)
    strAppID = data
End Property

Property Get frmAppID() As String
    frmAppID = strAppID
End Property

'Sets what to query the data against
Property Let GridSource(data As String)
    strGridSource = data
End Property

Property Get GridSource() As String
    GridSource = strGridSource
End Property

'This is the query criteria
Property Let Criteria(data As String)
    strCriteria = data
End Property

Property Get Criteria() As String
    Criteria = strCriteria
End Property

Private Sub cboSearch_Click()
    'Update the properties based on what is entered by the user
    Me.GridSource = Me.cboSearch
    Me.Criteria = ""
    Me.frmAppID = DLookup("SearchType", "GENERAL_Search", "SQLFrom = '" & Me.cboSearch & "'")
    
    'TL add account ID logic
    RefreshListBox "SELECT CriteriaID, UserID,  Description FROM CRITERIA_hdr WHERE SQLFrom = '" & Me.GridSource & "' and AccountID = " & gintAccountID, Me.lstSearch, "", ""
End Sub

Private Sub cmdClearFilter_Click()
    Me.Criteria = ""
    RefreshMain
    CmDRun_Click
End Sub

Private Sub CmDRun_Click()
    Dim strSQL As String
    Dim strError As String
    
    On Error GoTo ErrHandler
    
    'Build the SQL string and query the data
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM " & strGridSource
    'Check to see if the user had built on criteria
    'TL add account id logic
    If Trim(strCriteria) <> "" Then
        strSQL = strSQL & " WHERE  " & strCriteria & " and AccountID = " & gintAccountID
    Else
        strSQL = strSQL & " WHERE AccountID = " & gintAccountID
    End If
    
    'Refresh the grid based on the rowsource passed into the form
    Me.frm_GENERAL_Datasheet.Form.InitData GridSource, 3
    Me.frm_GENERAL_Datasheet.Form.RecordSource = strSQL


    Dim ctl As Control
     
    'Loop through the controls and size them correctly.
    For Each ctl In Me.frm_GENERAL_Datasheet.Form.Controls
      If ctl.ControlType = acTextBox Then
          ctl.ColumnWidth = -2
      End If
   Next

exitHere:

Exit Sub
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub

Private Sub cmdSimple_Click()
    'Dim frm As Form
    'This launches the Query builder form
    mvSql.From = Me.GridSource
    Set frm = New Form_frm_GENERAL_Filter
   
    With frm
        .visible = True
        .SQL = mvSql
        .CalledBy = Me.Name
        .Setup
        .Modal = False
    End With
End Sub

Private Sub Form_Close()
    On Error Resume Next
        'Instanced form, remove from collection
        RemoveObjectInstance Me
    Set frm = Nothing
End Sub

Public Sub RefreshMain()
    'If nothing is specified on load, we default the form to a Claims search
    If Me.frmAppID = "" Then
        Me.frmAppID = "AUDITCLM"
        RefreshComboBox "SELECT SQLFrom ,SearchName  FROM GENERAL_Search", Me.cboSearch, "AUDITCLM_Hdr", "SQLFrom"
        Me.GridSource = Me.cboSearch
        Me.Criteria = "1=1"
    Else
        RefreshComboBox "SELECT SQLFrom ,SearchName  FROM GENERAL_Search", Me.cboSearch, Me.GridSource, "SQLFrom"
        Me.Criteria = "1=1"
    End If
    
    'TL add account ID logic
    RefreshListBox "SELECT CriteriaID, UserID,  Description FROM CRITERIA_hdr WHERE SQLFrom = '" & Me.GridSource & "' and AccountID = " & gintAccountID, Me.lstSearch, "", ""
    
    Me.lblSend.Caption = Me.GridSource
    Me.lblReturn.Caption = Me.Criteria
End Sub

Private Sub Form_Load()
    Call Account_Check(Me)
    
    'Main refresh based on properties
    'All forms should have one of these
    Me.frm_GENERAL_Datasheet.Form.RecordSource = ""
    RefreshMain
End Sub

Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub frm_UpdateSql()
    'Public event called by the query builder form to set the SQL of the search
    mvSql = frm.SQL
    Me.Criteria = frm.SQL.WherePrimary
    Me.lblReturn.Caption = Me.Criteria
End Sub

Private Sub lstSearch_DblClick(Cancel As Integer)
    'Apply a filter that has previously been saved.
    If Me.lstSearch.ListIndex > -1 Then
        
        Me.Criteria = DLookup("SQLWHERE", "Criteria_Hdr", "CriteriaID = " & Me.lstSearch & "")
        CmDRun_Click
        
        Me.lblReturn.Caption = Me.Criteria
    End If

End Sub
