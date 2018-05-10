Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboReason_AfterUpdate()
    activateOtherSection
End Sub

Private Sub activateOtherSection()
    If cboReason.Value = "Other" Then
        lblOtherReason.visible = True
        txtOtherReason.visible = True
        lblDaysExtend.visible = True
        cboDaysExtend.visible = True
    Else
        lblOtherReason.visible = False
        txtOtherReason.visible = False
        lblDaysExtend.visible = False
        cboDaysExtend.visible = False
    End If
End Sub

Private Sub cmdClear_Click()

    txtSearchValue.Value = ""
    claimsLookup "DELETE", "", "" 'MG delete table record based on their session id
    clearScreen

End Sub

Private Sub clearScreen()
    'Dim lstBoxCount As Integer
    'lstBoxCount = lstSelectedClaims.ListCount
    
    'Dim i As Integer
    'For i = 0 To lstBoxCount - 1
    'For i = 1 To lstSelectedClaims.ListCount - 1
        'Remove an item from the ListBox.
    '    lstSelectedClaims.RemoveItem i
    'Next i
    lstSelectedClaims.RowSource = ""
    createHeaderInListBox
    lblSaveConfirmation.Caption = ""
    txtOtherReason.Value = ""
    cboDaysExtend.Value = "7"
    
End Sub

Private Sub cmdClearAllClaims_Click()
    clearScreen
End Sub

Private Sub cmdImportSpreadsheet_Click()

End Sub

Private Sub cmdSave_Click()
    'MsgBox lstSelectedClaims.ListCount
    If lstSelectedClaims.ListCount > 1 Then
        'MsgBox lstSelectedClaims.Column(0, 1) 'get first row request number
        'MsgBox lstSelectedClaims.Column(1, 1) 'get first row cnly claim num
        
        'Ensure other reasons are not null
        Me.txtOtherReason = Nz(Me.txtOtherReason, "")
        
        'MsgBox cboReason.Value
        
        'MsgBox "day extend = " & cboReason.Value 'MG to get the value, check the property of the bound column. It starts at column 0!
        Dim DaysExtend As Integer
        DaysExtend = 0 'MG SQL sp will calculate days extend
        If cboReason.Value = "Other" Then
            DaysExtend = cboDaysExtend.Value 'For use if user select other reasons
        End If
        
        'MG VBA listbox index starts at 0
        'index starts at 1 because 0 is the header row
        Dim i As Integer
        For i = 1 To lstSelectedClaims.ListCount - 1
            'MsgBox "cnlyClaimNum = " & lstSelectedClaims.Column(1, i)
            'MsgBox "instanceID = " & lstSelectedClaims.Column(0, i)
            'MsgBox i
            
            cmdSaveClaims lstSelectedClaims.Column(0, i), lstSelectedClaims.Column(1, i), Me.cboReason.Value, Me.txtOtherReason, DaysExtend
        Next i
        
        lblSaveConfirmation.Caption = "Saved on " & Now
    End If

End Sub

Function cmdSaveClaims(InstanceId As String, CnlyClaimNum As String, ReasonDesc As String, Note As String, DaysExtend As Integer)

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_PROV_MR_Extension_Claims_Process_v2"
    cmd.Parameters.Refresh
    cmd.Parameters("@pInstanceID") = InstanceId
    cmd.Parameters("@pCnlyClaimNum") = CnlyClaimNum
    cmd.Parameters("@pReasonDesc") = ReasonDesc
    cmd.Parameters("@pNote") = Note
    cmd.Parameters("@pDaysExtend") = DaysExtend
    cmd.Execute
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
End Function

Function claimsLookup(searchField As String, searchValue As String, excludeMRReceived As String)
    
    If Nz(excludeMRReceived, "") = -1 Then
        excludeMRReceived = "Y"
    Else
        excludeMRReceived = "N"
    End If

    'MsgBox excludeMRReceived

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_PROV_MR_Extension_Claims_Lookup_v2"
    cmd.Parameters.Refresh
    cmd.Parameters("@pSearchField") = searchField
    cmd.Parameters("@pSearchValue") = searchValue
    cmd.Parameters("@pExcludeMRReceived") = excludeMRReceived
    cmd.Execute
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
    'MG refresh data sheet
    Dim sqlString As String
    sqlString = " SessionID = " & Chr(34) & Identity.UserName & Chr(34)
    frm_PROV_MR_Extension_Lookup_subform.Form.filter = sqlString
    frm_PROV_MR_Extension_Lookup_subform.Form.FilterOn = True
    frm_PROV_MR_Extension_Lookup_subform.Form.Requery
    frm_PROV_MR_Extension_Lookup_subform.Form.Refresh
    
End Function

'Public g_sqlString As String

Private Sub cmdSearch_Click()
    'MsgBox chkExcludeMRReceived
    
    clearScreen
    
    'MG only execute function is something is typed in
    If Trim(Me.txtSearchValue) <> "" Then
        claimsLookup Me.cboSearchField.Value, Me.txtSearchValue.Value, Me.chkExcludeMRReceived.Value
    End If
    
    'MG check for filter
    'If (Len(Me.txtSearchValue.Value) > 0) Then
    '    'mg show only filtered claims
    '    frm_PROV_MR_Extension_Lookup_subform.Form.filter = sqlString
    '    frm_PROV_MR_Extension_Lookup_subform.Form.FilterOn = True
    '    'frm_subform_v_PROV_MR_Extension_Lookup.Form.filter
    'Else
    '    'mg show all claims
    '    frm_PROV_MR_Extension_Lookup_subform.Form.FilterOn = False
    'End If
    
    'refresh sql

    'MG 06-18-2013 filter based on session ID
    'Dim sqlString As String
    'sqlString = " SessionID = " & Chr(34) & Identity.UserName & Chr(34)
               
    'MG refresh data sheet
    'frm_PROV_MR_Extension_Lookup_subform.Form.filter = sqlString
    'frm_PROV_MR_Extension_Lookup_subform.Form.FilterOn = False
    'frm_PROV_MR_Extension_Lookup_subform.Form.Requery
    'frm_PROV_MR_Extension_Lookup_subform.Form.Refresh
        
End Sub


Private Sub cmdSelectAll_Click()

On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim recordCount As Integer
    Dim index As Integer
    
    Dim db As Database
    Dim rs As DAO.RecordSet
    
    strSQL = "SELECT RequestNumber,CnlyClaimNum FROM PROV_MR_Extension_Lookup_v2 WHERE sessionID='" & Identity.UserName & "'"
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(strSQL)
        
    rs.MoveLast
    recordCount = rs.recordCount 'get record count
    rs.MoveFirst
    
    'For index = 1 To recordCount
    For index = 0 To recordCount - 1
        'MsgBox rs.Fields(1)
        Me.lstSelectedClaims.AddItem rs.Fields(0) & ";" & rs.Fields(1)
        rs.MoveNext
    Next index
   
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical
    
End Sub




Private Sub cmdSpreadsheetView_Click()

    DoCmd.OpenForm "frm_Prov_MR_Extension_Batch"
    
End Sub

Private Sub Form_Load()

    'MG 10/17/2013 check user access. Only display import spreadsheet button for admin only because this button is very dangerous when abused
    Dim accessType As String
    accessType = DLookup("AccessType", "PROV_MR_Extension_Users", "userID='michael.guan'")
    
    If accessType = "admin" Then
        cmdSpreadsheetView.visible = True
    Else
        cmdSpreadsheetView.visible = False
    End If
    
    
    'MG add default value for exclude mr received filter
    Me.chkExcludeMRReceived.Value = 0
    
    'MG clear previous temporary record lookup
    claimsLookup "DELETE", "", "" 'MG delete table record based on their session id
    
    'MG add header to list box
    clearScreen
        
    'MG populate days extend in cbo box
    Dim i As Integer
    Dim dayExtendValues As String
    For i = 7 To 45
         dayExtendValues = dayExtendValues & i & ";"
    Next i
    cboDaysExtend.RowSource = dayExtendValues
    
    activateOtherSection
    
    'MG refresh data sheet
    'frm_PROV_MR_Extension_Lookup_subform.Form.Requery
    'frm_PROV_MR_Extension_Lookup_subform.Form.Refresh
    
    
End Sub



Private Sub createHeaderInListBox()
    'List box was created for users to confirm that selected claims are what they want to grant extension to
    Me.lstSelectedClaims.AddItem "RequestNumber,CnlyClaimNum"
End Sub



Private Sub lstSelectedClaims_DblClick(Cancel As Integer)
    'MG get value from selected row
    'MsgBox "row index = " & lstSelectedClaims.ItemsSelected.Item(0)
    'MsgBox lstSelectedClaims.Column(0, lstSelectedClaims.ItemsSelected.Item(0))
    
    'Dim instanceIDHighlighted As String
    'Dim cnlyClaimNumHighlighted As String
    
    'instanceIDHighlighted = lstSelectedClaims.Column(0, lstSelectedClaims.ItemsSelected.Item(0))
    'cnlyClaimNumHighlighted = lstSelectedClaims.Column(1, lstSelectedClaims.ItemsSelected.Item(0))
    
    'MsgBox instanceIDHighlighted
    'MsgBox cnlyClaimNumHighlighted
    If lstSelectedClaims.ItemsSelected.Count > 0 Then
        lstSelectedClaims.RemoveItem (lstSelectedClaims.ItemsSelected.Item(0))
    End If
    
End Sub
