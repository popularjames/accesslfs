Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private WithEvents coMainGrid As Form_frm_GENERAL_Datasheet_ADO
Attribute coMainGrid.VB_VarHelpID = -1
Private WithEvents coProvAddr As Form_frm_GENERAL_Datasheet_ADO
Attribute coProvAddr.VB_VarHelpID = -1
Private WithEvents coAdo As clsADO
Attribute coAdo.VB_VarHelpID = -1

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property



Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim iInt As Integer


    strProcName = ClassName & ".RefreshData"
    
    If Me.lvwBuiltFilter.ListItems.Count < 1 Then
        Me.cmbAndOrNot.visible = False
    Else
        Me.cmbAndOrNot.visible = True
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ManualFilterByList"
        .Parameters.Refresh
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    Me.cmbFilterByField.ColumnCount = oRs.Fields.Count
    Me.cmbFilterByField.ColumnWidths = "0;3"""
    
    For iInt = 3 To oRs.Fields.Count
        Me.cmbFilterByField.ColumnWidths = Me.cmbFilterByField.ColumnWidths & ";0"
    Next
    
    Set Me.cmbFilterByField.RecordSet = oRs

    Call SetRequiredValEntryMode
        
    Call EnableDisableAndHideControls
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub SetRequiredValEntryMode()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSprocName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".SetRequiredValEntryMode"
    
    If Nz(Me.cmbFilterByField.Value, 0) = 0 Then
        ' nothing to do
        Me.txtRequiredVal.visible = False
        Me.cmbRequiredVal.visible = False
        
        GoTo Block_Exit
    End If
    
    If Me.cmbFilterByField.ListIndex > -1 Then
        sSprocName = Nz(Me.cmbFilterByField.Column(3, Me.cmbFilterByField.ListIndex), "")
        If sSprocName <> "" Then
            Me.txtRequiredVal = ""
            Me.txtRequiredVal.visible = False
            Set oRs = GetListRS(sSprocName, Me.cmbRequiredVal)
            Set Me.cmbRequiredVal.RecordSet = oRs
            Me.cmbRequiredVal.visible = True
        Else
            Me.txtRequiredVal = ""
            Set Me.cmbRequiredVal.RecordSet = Nothing
            Me.cmbRequiredVal.visible = False
            Me.txtRequiredVal.visible = True
        End If
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function GetListRS(sSprocName As String, oComboBox As ComboBox) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim iFlds As Integer

    strProcName = ClassName & ".GetListRS"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = sSprocName
        .Parameters.Refresh
        If .Parameters.Count > 2 Then
            .Parameters("@pAccountId") = gintAccountID
        End If
        Set oRs = .ExecuteRS
        If .GotData = False Or Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    '' Standard first column is bound and should be invisible:
    oComboBox.ColumnWidths = "0;2.5"""
    oComboBox.ColumnCount = oRs.Fields.Count
    For iFlds = 2 To oRs.Fields.Count - 1
        oComboBox.ColumnWidths = oComboBox.ColumnWidths & ";0"
    Next
    
    Set GetListRS = oRs
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'Private Sub cmbFilterByField_AfterUpdate()
'    Call SetRequiredValEntryMode
'End Sub

Private Sub cmbFilterByField_Change()
    Call SetRequiredValEntryMode
End Sub

Private Sub cmdAddCriteria_Click()
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLV As ListView
Dim oLV As Object
Dim oLI As ListItem
Dim iCnt As Integer

    strProcName = ClassName & ".cmdAddCriteria_Click"
    
    Set oLV = Me.lvwBuiltFilter
    iCnt = oLV.ListItems.Count + 1
    
    Set oLI = oLV.ListItems.Add(, , Format(iCnt, "0##"))
    If Me.cmbAndOrNot.visible = True Then
        oLI.SubItems(1) = Nz(Me.cmbAndOrNot, "AND")
    End If
    oLI.SubItems(2) = Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "")
    
    oLI.SubItems(3) = Nz(Me.cmbOperator, "=")
    
    If Me.cmbRequiredVal.visible = True Then
        oLI.SubItems(4) = Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")
        oLI.SubItems(5) = Nz(Me.cmbRequiredVal.Column(1, Me.cmbRequiredVal.ListIndex), "")
    Else
        oLI.SubItems(4) = ""    '   Nz(Me.cmbFilterByField.Column(3, Me.cmbFilterByField.ListIndex), "")
        oLI.SubItems(5) = Nz(Me.txtRequiredVal, "")
        
    End If
    
'    oLI.SubItems(6) = Nz(Me.cmbRequiredVal.Column(6, Me.cmbRequiredVal.ListIndex), "")
    oLI.SubItems(6) = Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "")
    
    Me.cmbAndOrNot.visible = True
    
    cmdViewCollectiveSample_Click
    
    RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub EnableDisableAndHideControls()
    If Me.cmbAndOrNot.visible = False Then
        Me.cmbAndOrNot = "AND"
    End If
    If Me.lvwBuiltFilter.ListItems.Count > 0 Then
        Me.cmdEditSelected.Enabled = True
        Me.cmdRemoveSelected.Enabled = True
        Me.cmdSave.Enabled = True
        Me.cmbAndOrNot.visible = True
    Else
        Me.cmdEditSelected.Enabled = False
        Me.cmdRemoveSelected.Enabled = False
        Me.cmdSave.Enabled = False
        Me.cmbAndOrNot.visible = False
    End If
End Sub

Private Sub cmdEditSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem

    strProcName = ClassName & ".cmdEditSelected_Click"
    
    Set oLI = Me.lvwBuiltFilter.SelectedItem
        
    '' Need to load our combo boxes..
    If oLI.SubItems(1) <> "" Then
        Call SelectComboBoxByVal(Me.cmbAndOrNot, oLI.SubItems(1))
    End If
    
    Call SelectComboBoxByVal(Me.cmbFilterByField, oLI.SubItems(2))

    Me.cmbOperator = oLI.SubItems(3)
    
    Call SetRequiredValEntryMode
    
    
    If Me.txtRequiredVal.visible = True Then
        Me.txtRequiredVal = oLI.SubItems(5)
    Else
        Call SelectComboBoxByVal(Me.cmbRequiredVal, oLI.SubItems(5))
    End If
    
    '' Note, if this is the first one then we need to remove the 'Inclusion' value on the next, and renumber
    
    lvwBuiltFilter.ListItems.Remove (oLI.index)
    Call RenumberItems
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub RenumberItems()
On Error GoTo Block_Err
Dim strProcName As String
Dim iIdx As Integer
Dim oLI As ListItem
Dim oLV As ListView

    strProcName = ClassName & ".RenumberItems"
    For iIdx = 1 To Me.lvwBuiltFilter.ListItems.Count
        Set oLI = Me.lvwBuiltFilter.ListItems(iIdx)
        oLI.Text = Format(iIdx, "0##")
        If iIdx = 1 Then
            If oLI.SubItems(1) <> "" Then
                oLI.SubItems(1) = ""
            End If
        End If
    Next
    
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub SelectComboBoxByVal(oCmb As ComboBox, sDesiredVal As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim sThisVal As String
Dim iIdx As Integer

    strProcName = ClassName & ".SelectComboBoxByVal"
    For iIdx = 0 To oCmb.ListCount
        If Trim(LCase(oCmb.Column(1, iIdx))) = LCase(sDesiredVal) Then

            oCmb = oCmb.Column(oCmb.BoundColumn - 1, iIdx)
            Exit For
        Else
'            Stop
        End If
    Next
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdRemoveSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem

    strProcName = ClassName & ".cmdRemoveSelected_Click"
    
     Set oLI = Me.lvwBuiltFilter.SelectedItem
    lvwBuiltFilter.ListItems.Remove (oLI.index)
    Call RenumberItems
    Call cmdViewCollectiveSample_Click
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdSave_Click"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_InsManualFilter"
        .Parameters.Refresh
        .Parameters("@pAccountId") = 0
        .Parameters("@pOption") = ""
        .Parameters("@pInclusion") = ""
        .Parameters("@pFilterBy") = ""
        .Parameters("@pOperator") = ""
        .Parameters("@pRequiredValIdx") = ""
        .Parameters("@pRequiredVal") = ""
        
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        
    End With
    
    
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim sWhere As String
Dim sFrom As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdSearch_Click"
    
    ' Make sure there's some kind of criteria first:
    If Nz(Me.cmbFilterByField, "") = "" Then
        Stop
    End If
    If Me.cmbRequiredVal.visible = False And Me.txtRequiredVal.visible = False Then
        Stop
    End If
    
    
    sSql = "SELECT TOP 1000 * "
    sFrom = " FROM " & Nz(Me.cmbFilterByField.Column(4, Me.cmbFilterByField.ListIndex), "")
    '    sFrom = " FROM MCRRestricted.AUDITCLM_Hdr "
    
    If Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "") = "" Then
        sWhere = " WHERE " & Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "") & " " & Me.cmbOperator & " '"
    Else
        sWhere = " WHERE " & Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "") & " " & Me.cmbOperator & " '"
    End If
    If Me.cmbRequiredVal.visible = False Then
        sWhere = sWhere & Trim(Me.txtRequiredVal) & "' "
    Else
        sWhere = sWhere & Trim(Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")) & "' "
    End If
    
    sSql = sSql & sFrom & sWhere
    
Debug.Print sSql
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With
    
    Me.oMainGrid.Form.InitDataADO oRs, ""
    Set Me.oMainGrid.Form.RecordSet = oRs
    
    Set coMainGrid = Me.oMainGrid.Form
    coMainGrid.InitDataADO oRs, ""
    coMainGrid.AllowFilters = True
    'Loop through the controls and size them correctly.
Dim oCtl As Control
    For Each oCtl In coMainGrid.Controls
      If oCtl.ControlType = acTextBox Then
          oCtl.ColumnWidth = -2
      End If
   Next
   coMainGrid_Current
   
    sSql = "SELECT COUNT(1) As Cnt " & sFrom & sWhere
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With
    Me.lblRecordCount.Caption = "Full Filter Record Count: " & Format(CStr(oRs("Cnt").Value), "###,###,###,###")
    

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing

    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub cmdViewCollectiveSample_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim sFrom As String
Dim sWhere As String
Dim iParens As Integer
Dim oLV As ListView
Dim oLItem As ListItem
Dim iItem As Integer
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".cmdViewCollectiveSample_Click"
    DoCmd.Hourglass True
  
    
    sSql = "SELECT TOP 1000 * "
    sFrom = " FROM v_LETTER_Automation_ManualOverrideSample "
    sWhere = " WHERE ("
'Stop
    For iItem = 1 To Me.lvwBuiltFilter.ListItems.Count
        Set oLItem = Me.lvwBuiltFilter.ListItems(iItem)
        
        If Nz(oLItem.SubItems(1), "") <> "" Then
            sWhere = sWhere & " " & oLItem.SubItems(1) & " "
        End If
        
        sWhere = sWhere & "( "
        
        If Nz(oLItem.SubItems(6), "") = "" Then
            sWhere = sWhere & oLItem.SubItems(2) & " " & oLItem.SubItems(3) & " '"
        Else
            sWhere = sWhere & oLItem.SubItems(6) & " " & oLItem.SubItems(3) & " '"
        End If
        If Nz(oLItem.SubItems(4), "") = "" Then
            sWhere = sWhere & Trim(oLItem.SubItems(5)) & "' "
        Else
            sWhere = sWhere & Trim(oLItem.SubItems(4)) & "' "
        End If
        
        sWhere = sWhere & " ) "
    Next
    
    sWhere = sWhere & ")"
    
   
    
    sSql = sSql & sFrom & sWhere
    
Debug.Print sSql
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With
    
'    Me.oMainGrid.Form.InitDataADO oRs, ""
'    Set Me.oMainGrid.Form.Recordset = oRs
    
    Set coMainGrid = Me.oMainGrid.Form
    Set coMainGrid.RecordSet = oRs
    coMainGrid.InitDataADO oRs, ""
    coMainGrid.AllowFilters = True
    'Loop through the controls and size them correctly.
Dim oCtl As Control
    For Each oCtl In coMainGrid.Controls
      If oCtl.ControlType = acTextBox Then
          oCtl.ColumnWidth = -2
      End If
   Next
   coMainGrid_Current
   
    sSql = "SELECT COUNT(1) As Cnt " & sFrom & sWhere
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With
    Me.lblRecordCount.Caption = "Full Filter Record Count: " & Format(CStr(oRs("Cnt").Value), "###,###,###,##0")
    
    
    
Block_Exit:
    DoCmd.Hourglass False
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub coMainGrid_Current()
    Set coAdo = New clsADO

    If coMainGrid Is Nothing Then
        GoTo Block_Exit
    End If
    If coMainGrid.RecordSource = "" Then
'        Stop
        GoTo Block_Exit
    End If

    ' Need to load the oProvAddresses grid with the first
    ' provider
    Call LoadAddressGrid

'    Me.txtConceptId = Nz(oMainGrid.Controls("ConceptID"), "")
'    Me.txtNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'    mNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'    'Refresh the tabs to ensure the main form is in sync with the other forms.
'
'    coAdo.ConnectionString = GetConnectString("v_Data_Database")
'
'    coAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.txtConceptId & "'"
'    Set mrsConcept = coAdo.OpenRecordSet()
'
'    coAdo.sqlString = " SELECT * from Note_Detail WHERE NoteID = '" & Me.txtNoteID & "'"
'    Set mrsNotes = coAdo.OpenRecordSet()

'    ' if it's the same concept, no need to "click" the tab
'    If Not frmConceptHdr Is Nothing Then
'        If Me.txtConceptId <> frmConceptHdr.FormConceptID Then
'            lstTabs_Click
'        Else
''            Stop
'            lstTabs_Click
'        End If
'    Else
'        lstTabs_Click
'    End If
Block_Exit:

End Sub

Private Sub LoadAddressGrid()
On Error GoTo Block_Err
Dim strProcName As String
Dim sCnlyProvId As String
Dim oRs As ADODB.RecordSet
Dim oFld As ADODB.Field
Dim oAdo As clsADO
Dim sSql As String

    strProcName = ClassName & ".LoadAddressGrid"
    
    Set oRs = coMainGrid.RecordSet.Clone
    
    For Each oFld In oRs.Fields
        If LCase(oFld.Name) = "cnlyprovid" Then
            sCnlyProvId = oRs("CnlyProvId").Value
            Exit For
        End If
    Next
    
    If sCnlyProvId = "" Then
        GoTo Block_Exit
    End If
    
    Set oRs = Nothing
    
    sSql = "SELECT * FROM v_LETTER_Automation_ManualOverrideProvSample WHERE CnlyProvId = '" & sCnlyProvId & "'"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
'    Me.oProvAddresses.Form.InitDataADO oRs, ""
'    Set Me.oProvAddresses.Form.Recordset = oRs
    
    Set coProvAddr = Me.oProvAddresses.Form
    coProvAddr.InitDataADO oRs, ""
    Set coProvAddr.RecordSet = oRs
    coProvAddr.AllowFilters = True
    'Loop through the controls and size them correctly.
Dim oCtl As Control
    For Each oCtl In coProvAddr.Controls
      If oCtl.ControlType = acTextBox Then
          oCtl.ColumnWidth = -2
      End If
   Next
   coProvAddr_Current
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Command14_Click()
    Call SetRequiredValEntryMode
    Dim oLV As ListView
    
Stop
End Sub



Private Sub coProvAddr_Current()
    

    If coProvAddr Is Nothing Then
        GoTo Block_Exit
    End If
    If coProvAddr.RecordSource = "" Then
'        Stop
        GoTo Block_Exit
    End If



'    Me.txtConceptId = Nz(oMainGrid.Controls("ConceptID"), "")
'    Me.txtNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'    mNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
'    'Refresh the tabs to ensure the main form is in sync with the other forms.
'
'    coAdo.ConnectionString = GetConnectString("v_Data_Database")
'
'    coAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.txtConceptId & "'"
'    Set mrsConcept = coAdo.OpenRecordSet()
'
'    coAdo.sqlString = " SELECT * from Note_Detail WHERE NoteID = '" & Me.txtNoteID & "'"
'    Set mrsNotes = coAdo.OpenRecordSet()

'    ' if it's the same concept, no need to "click" the tab
'    If Not frmConceptHdr Is Nothing Then
'        If Me.txtConceptId <> frmConceptHdr.FormConceptID Then
'            lstTabs_Click
'        Else
''            Stop
'            lstTabs_Click
'        End If
'    Else
'        lstTabs_Click
'    End If
Block_Exit:

End Sub

Private Sub Form_Load()
    RefreshData
End Sub


Private Sub Form_Resize()
Dim dGridSpacer As Double

    ' Very basic stuff here..
    ' only going to do the main grid and the provider addresses grid
    If Me.InsideWidth < 12840 Then
        Exit Sub
    End If
    
    dGridSpacer = (Me.oMainGrid.left * 2)
    
    Me.oMainGrid.Width = Me.InsideWidth - dGridSpacer
    Me.oProvAddresses.Width = Me.InsideWidth - dGridSpacer

    ' Alright, screw it, it looks un finished so we'll do the top stuff too
    
    Me.fraRightButtons.left = Me.InsideWidth - Me.fraRightButtons.Width - dGridSpacer
    Me.lvwBuiltFilter.Width = Me.fraRightButtons.left - (dGridSpacer * 2)
    
    Me.cmdViewCollectiveSample.left = Me.fraRightButtons.left + dGridSpacer
    Me.cmdEditSelected.left = Me.cmdViewCollectiveSample.left
    Me.cmdRemoveSelected.left = Me.cmdViewCollectiveSample.left
    Me.cmdSave.left = Me.cmdViewCollectiveSample.left
    Me.CmdCancel.left = Me.cmdViewCollectiveSample.left
    
    Me.fraUpperRightButtons.Width = Me.InsideWidth - (Me.fraUpperLeftButtons.Width + Me.fraUpperLeftButtons.left) - dGridSpacer
    
    
End Sub
