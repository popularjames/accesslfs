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

Private coManFilter As clsFilter
Private coCurFilterOpt As clsFilterOption

Private cbEditing As Boolean


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get FilterName() As String
    FilterName = Me.txtFilterName
End Property
Public Property Let FilterName(sFilterName As String)
    Me.txtFilterName = sFilterName
End Property

'Public Property Get FilterId() As Long
'
'End Property
Public Property Let FilterID(lFilterId As Long)
    Set coManFilter = New clsFilter
    coManFilter.LoadFromId (lFilterId)
End Property



Public Property Get FilterObject() As clsFilter
    Set FilterObject = coManFilter
End Property
Public Property Let FilterObject(oFilter As clsFilter)
    Set coManFilter = oFilter
    Me.FilterName = coManFilter.FilterName
End Property


Public Function Initialize() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Initialize"
    
    cbEditing = True
    ' bottom line, populate everything from the global coManFilter
    If coManFilter Is Nothing Then
        GoTo Block_Exit
    End If
    If coManFilter.Id = 0 Then
        Stop
        GoTo Block_Exit
    End If
    
    For Each coCurFilterOpt In coManFilter.FilterOptions
        Call AddOptionToListView(coCurFilterOpt)
    Next

    Call Save_Backup
    
    ' then we want the LAST one to load into the controls
    ' so select the last one entered
'    Stop
    If SelectLastItem(Me.lvwBuiltFilter) = True Then
        Me.TimerInterval = 300  ' once the form is visble we can edit the last item
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function Save_Backup() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".Save_Backup"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_AUtomation_ManualFilterEditBkup"
        .Parameters.Refresh
        .Parameters("@pManFilterid") = coManFilter.Id
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    Save_Backup = True
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function Edit_Rollback() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".Edit_Rollback"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_AUtomation_ManualFilterEditRollback"
        .Parameters.Refresh
        .Parameters("@pManFilterid") = coManFilter.Id
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    Edit_Rollback = True
    
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

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


'''
''' This sub basically just sets some control visibility based on the current state
'''
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

'''
''' Generic function to load a combo box with the results from the Sproc passed in
'''
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


Private Sub cmbFilterByField_AfterUpdate()
    Call SetRequiredValEntryMode
End Sub

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
Dim lNewId As Long
Dim bSave As Boolean


    strProcName = ClassName & ".cmdAddCriteria_Click"
    
    If Not coCurFilterOpt Is Nothing Then
        If coCurFilterOpt.Save = False Then
            Call coManFilter.RemoveFilterOption(coCurFilterOpt)
        End If
    End If
    Set coCurFilterOpt = Nothing
    
    ' If the current one is set then we should be able to use it..
'    If coCurFilterOpt Is Nothing Then
'Stop
MakeNew:
        Set coCurFilterOpt = coManFilter.NewFilterOption(CLng(Me.cmbFilterByField))
'        Set coCurFilterOpt = New clsFilterOption
'        coCurFilterOpt.ManFilterID = coManFilter.ManFilterID
'
'        coCurFilterOpt.UnderlyingOptionId = CLng(Me.cmbFilterByField)
'        lNewId = coCurFilterOpt.NewId()
'        coCurFilterOpt.LoadFromID (lNewId)
'        coManFilter.AddFilterOption coCurFilterOpt
'    Else
'        If coCurFilterOpt.UnderlyingOptionId <> CLng(Me.cmbFilterByField) Then
''            Stop
'            If coCurFilterOpt.Save = False Then
'                ' need to remove this one
'                Call coManFilter.RemoveFilterOption(coCurFilterOpt)
'            End If
'            GoTo MakeNew
'        End If
'        bSave = True
''Stop
'    End If
    
    coCurFilterOpt.Save = True

    
    Set oLV = Me.lvwBuiltFilter
    iCnt = oLV.ListItems.Count + 1
    
    coCurFilterOpt.OptionID = Format(iCnt, "000")
    
    
    Set oLI = oLV.ListItems.Add(, , Format(coCurFilterOpt.OptionID, "0##"))
'    oLI.Tag = CStr(coManFilter.ID)
    oLI.Tag = CStr(coCurFilterOpt.Id)
    
    If Me.cmbAndOrNot.visible = True Then
        oLI.SubItems(1) = Nz(Me.cmbAndOrNot, "AND")
        coCurFilterOpt.Inclusion = oLI.SubItems(1)
    End If
    
    oLI.SubItems(2) = Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "")
    coCurFilterOpt.FilterBy = oLI.SubItems(2)
    
    oLI.SubItems(3) = Nz(Me.cmbOperator, "=")
    coCurFilterOpt.Operator = oLI.SubItems(3)
    
    If Me.cmbRequiredVal.visible = True Then
        oLI.SubItems(4) = Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")
        oLI.SubItems(5) = Nz(Me.cmbRequiredVal.Column(1, Me.cmbRequiredVal.ListIndex), "")
        
    Else
        oLI.SubItems(4) = ""    '   Nz(Me.cmbFilterByField.Column(3, Me.cmbFilterByField.ListIndex), "")
        oLI.SubItems(5) = Nz(Me.txtRequiredVal, "")
    End If

    coCurFilterOpt.RequiredValueIdx = oLI.SubItems(4)
    coCurFilterOpt.RequiredValue = oLI.SubItems(5)
    
'    oLI.SubItems(6) = Nz(Me.cmbRequiredVal.Column(6, Me.cmbRequiredVal.ListIndex), "")
    oLI.SubItems(6) = Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "")
        
    Me.cmbAndOrNot.visible = True
    
    coCurFilterOpt.SaveNow

    cmdViewCollectiveSample_Click
    
    RefreshData
    
    Me.cmbFilterByField = ""
    Me.cmbRequiredVal = ""
    Me.txtRequiredVal = ""
    
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

Private Sub CmdCancel_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFltrOpt As clsFilterOption

    strProcName = ClassName & ".cmdCancel_Click"
    
    If cbEditing = True Then
        '' Roll back
        Call Edit_Rollback
        GoTo Block_Exit
    End If
    
    If coManFilter Is Nothing Then
        Stop
        GoTo Block_Exit
    End If
        
    For Each oFltrOpt In coManFilter.FilterOptions
        oFltrOpt.Delete
    Next
    
    coManFilter.Delete
    
    
Block_Exit:
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

'''
''' loads the selected filter option into the controls for editing
''' Note: Removes the list item (not from the database though..)
'''
Private Sub cmdEditSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oFltrOpt As clsFilterOption

    strProcName = ClassName & ".cmdEditSelected_Click"
    
    Set oLI = Me.lvwBuiltFilter.SelectedItem
    cbEditing = True
        
    Set oFltrOpt = New clsFilterOption
    If oFltrOpt.LoadFromId(oLI.Tag) = False Then
        Stop
    End If
    
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
    ' Also need to remove it from the database:
    
    Call coManFilter.RemoveFilterOption(oFltrOpt)
    Call RenumberItems
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


'''
''' Simply makes sure that the "optionid" is numbered from 1 to xxx
'''
Private Sub RenumberItems()
On Error GoTo Block_Err
Dim strProcName As String
Dim iIdx As Integer
Dim oLI As ListItem
Dim oLV As ListView
Dim oOption As clsFilterOption


    strProcName = ClassName & ".RenumberItems"
    For iIdx = 1 To Me.lvwBuiltFilter.ListItems.Count
        Set oLI = Me.lvwBuiltFilter.ListItems(iIdx)
        
        oLI.Text = Format(iIdx, "0##")
        
        Set oOption = GetObjectFromListItem(oLI)
        oOption.OptionID = oLI.Text
        oOption.SaveNow
        
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
    
    '' Also have to delete this from the database..
    
    Set oLI = Me.lvwBuiltFilter.SelectedItem
    Call GetObjectFromListItem(oLI)
    
    lvwBuiltFilter.ListItems.Remove (oLI.index)
    ' Also need to remove it from the database:
    Call coManFilter.RemoveFilterOption(coCurFilterOpt)
    
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
Dim oLV As ListView
Dim oLI As ListItem
Dim oFltrOpt As clsFilterOption

    strProcName = ClassName & ".cmdSave_Click"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ManualFilterOptionSave"
        .Parameters.Refresh
    End With
    
    
    For Each oLI In Me.lvwBuiltFilter.ListItems
        Set oFltrOpt = New clsFilterOption
        
        If oFltrOpt.LoadFromId(CLng(oLI.Tag)) = False Then
            Stop
        End If
        
'Stop
        With oAdo
            .Parameters.Refresh
            .Parameters("@pRID") = oFltrOpt.Id
            .Parameters("@pManFilterId") = oFltrOpt.ManFilterID
            .Parameters("@pUnderlyingOptionId") = oFltrOpt.UnderlyingOptionId
            .Parameters("@pOptionId") = oFltrOpt.OptionID
            .Parameters("@pInclusion") = oFltrOpt.Inclusion
            .Parameters("@pFilterBy") = oFltrOpt.FilterBy
            .Parameters("@pOperator") = oFltrOpt.Operator
            .Parameters("@pRequiredValueIdx") = oFltrOpt.RequiredValueIdx
            .Parameters("@pRequiredValue") = oFltrOpt.RequiredValue
            
            .Execute
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                Stop
            End If
            
        End With
    Next
    
    coManFilter.FilterName = Me.txtFilterName
    
    coManFilter.Active = True
    
    coManFilter.SaveNow
    If cbEditing = True Then
        ' get rid of the backup stuff..
        
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    
Block_Exit:
    
    
    Set oLI = Nothing
    Set oFltrOpt = Nothing
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

  
    If Not coCurFilterOpt Is Nothing Then
        If coCurFilterOpt.Save = False Then
            Call coManFilter.RemoveFilterOption(coCurFilterOpt)
        End If
    End If
    Set coCurFilterOpt = Nothing

    ' If our global object is nothing then we need to make a new one
'    If coCurFilterOpt Is Nothing Then
        If LoadCurrentFilterOption(True) = False Then
            GoTo Block_Exit
        End If
'    Else
'        ' we need to make sure it's set properly
'        If LoadCurrentFilterOption(False) = False Then
'            GoTo Block_Exit
'        End If
'    End If
    
    coCurFilterOpt.Save = False
    
    sWhere = coCurFilterOpt.GetWhereClause
    
    
    
    sSql = "SELECT TOP 1000 * "
    If Me.cmbFilterByField.Column(4, Me.cmbFilterByField.ListIndex) = "" Then
        sFrom = " FROM AUDITCLM_Hdr "
    
    Else
        sFrom = " FROM " & Nz(Me.cmbFilterByField.Column(4, Me.cmbFilterByField.ListIndex), "")
    
    End If
    
    '    sFrom = " FROM MCRRestricted.AUDITCLM_Hdr "
    
'    If Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "") = "" Then
'        sWhere = " WHERE " & Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "") & " " & Me.cmbOperator & " '"
'    Else
'        sWhere = " WHERE " & Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "") & " " & Me.cmbOperator & " '"
'    End If
'    If Me.cmbRequiredVal.visible = False Then
'        sWhere = sWhere & Trim(Me.txtRequiredVal) & "' "
'    Else
'        sWhere = sWhere & Trim(Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")) & "' "
'    End If
    
    sSql = sSql & sFrom & " WHERE " & sWhere
    
Debug.Print sSql
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            If .CurrentConnection.Errors.Count > 0 Then
                LogMessage strProcName, "ERROR", "There was an error: ", .CurrentConnection.Errors(0).Description
                GoTo Block_Exit
            Else
                LogMessage strProcName, "USER NOTICE", "No records found!", sWhere, True
            End If
'            Stop
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
   
    sSql = "SELECT COUNT(1) As Cnt " & sFrom & " WHERE " & sWhere
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

Private Function LoadCurrentFilterOption(Optional bIsNew As Boolean = False) As Boolean
Dim oLV As Object
Dim iCnt As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim lNewId As Long

    strProcName = ClassName & ".LoadCurrentFilterOption"
'Stop
    
    If bIsNew = True Then
        
        Set coCurFilterOpt = coManFilter.NewFilterOption(CLng(Me.cmbFilterByField))
        If coCurFilterOpt Is Nothing Then
            ' error
            Stop
        End If
    Else
        Stop
    End If
   
    Set oLV = Me.lvwBuiltFilter
    iCnt = oLV.ListItems.Count + 1
    
    coCurFilterOpt.OptionID = Format(iCnt, "000")
    
    
    If Me.cmbAndOrNot.visible = True Then
        coCurFilterOpt.Inclusion = Nz(Me.cmbAndOrNot, "AND")
    End If
    
    coCurFilterOpt.FilterBy = Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "")
    
    coCurFilterOpt.Operator = Nz(Me.cmbOperator, "=")
    
    If Me.cmbRequiredVal.visible = True Then
        coCurFilterOpt.RequiredValueIdx = Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")
        coCurFilterOpt.RequiredValue = Nz(Me.cmbRequiredVal.Column(1, Me.cmbRequiredVal.ListIndex), "")
        
    Else
        coCurFilterOpt.RequiredValueIdx = ""    '   Nz(Me.cmbFilterByField.Column(3, Me.cmbFilterByField.ListIndex), "")
        coCurFilterOpt.RequiredValue = Nz(Me.txtRequiredVal, "")
    End If

   
    coCurFilterOpt.SaveNow
    
    LoadCurrentFilterOption = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

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
Dim lIdToLoad As Long



    strProcName = ClassName & ".cmdViewCollectiveSample_Click"
    DoCmd.Hourglass True
    ' This should refresh the coManFilter
    ' then we should be good
    
    If coManFilter Is Nothing Then
        Stop
    End If
    
        ' Force it to refresh
    lIdToLoad = coManFilter.Id
    Set coManFilter = New clsFilter
    coManFilter.LoadFromId (lIdToLoad)
    
    
    sSql = "SELECT TOP 1000 * "
'    sFrom = " FROM v_LETTER_Automation_ManualOverrideSample "
    
    sFrom = coManFilter.GetFilterFromClause
    
    sWhere = coManFilter.GetFilterWhereClause
    
'    sWhere = " WHERE ("
''Stop
'    For iItem = 1 To Me.lvwBuiltFilter.ListItems.Count
'        Set oLItem = Me.lvwBuiltFilter.ListItems(iItem)
'
'        If Nz(oLItem.SubItems(1), "") <> "" Then
'            sWhere = sWhere & " " & oLItem.SubItems(1) & " "
'        End If
'
'        sWhere = sWhere & "( "
'
'        If Nz(oLItem.SubItems(6), "") = "" Then
'            sWhere = sWhere & oLItem.SubItems(2) & " " & oLItem.SubItems(3) & " '"
'        Else
'            sWhere = sWhere & oLItem.SubItems(6) & " " & oLItem.SubItems(3) & " '"
'        End If
'        If Nz(oLItem.SubItems(4), "") = "" Then
'            sWhere = sWhere & Trim(oLItem.SubItems(5)) & "' "
'        Else
'            sWhere = sWhere & Trim(oLItem.SubItems(4)) & "' "
'        End If
'
'        sWhere = sWhere & " ) "
'    Next
'
'    sWhere = sWhere & ")"
    
   
    
    sSql = sSql & sFrom & sWhere
    
Debug.Print sSql
Dim sConnStr As String


    If InStr(1, sFrom, "v_LETTER_Automation", vbTextCompare) > 1 Then
        sConnStr = CodeConnString
    ElseIf InStr(1, sFrom, "usp_LETTER_Automation", vbTextCompare) > 1 Then
        sConnStr = CodeConnString
    Else
        sConnStr = DataConnString
    End If
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = sConnStr

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
        .ConnectionString = sConnStr
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
    
    If oRs.recordCount = 0 Or (oRs.EOF And oRs.BOF) Then
        ' nothing to load - move on jerky!
        GoTo Block_Exit
    End If
    
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
    If Me.OpenArgs <> "" Then
Stop
        Set coManFilter = New clsFilter
        coManFilter.LoadFromId (CLng(Me.OpenArgs))
        ' then we want the LAST one to load into the controls
        ' so select the last one entered
        If SelectLastItem(Me.lvwBuiltFilter) = True Then
            Call cmdEditSelected_Click
        End If
        
    End If

    RefreshData
End Sub

Private Function SelectLastItem(oLV As Object) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oLItm As ListItem
'dim oLItm as Object
    strProcName = ClassName & ".SelectLastItem"
    
    If oLV.ListItems.Count < 1 Then GoTo Block_Exit
    
    ' deselect all
    For Each oLItm In oLV.ListItems
        oLItm.Selected = False
    Next
    oLV.ListItems(oLV.ListItems.Count).Selected = True
    
    SelectLastItem = True
    
Block_Exit:
    Set oLItm = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

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



Public Function GetObjectFromCurrentControls() As clsFilterOption
On Error GoTo Block_Err
Dim strProcName As String
Dim oLV As Object
Dim oLI As ListItem
Dim iCnt As Integer
Dim lNewId As Long


    strProcName = ClassName & ".GetObjectFromCurrentControls"
    
    ' If the current one is set then we should be able to use it..
    If coCurFilterOpt Is Nothing Then
'Stop
MakeNew:
        Set coCurFilterOpt = coManFilter.NewFilterOption(CLng(Me.cmbFilterByField))
    Else
        If coCurFilterOpt.UnderlyingOptionId <> CLng(Me.cmbFilterByField) Then
            Stop
            GoTo MakeNew
        End If
    End If

    

    
    Set oLV = Me.lvwBuiltFilter
    iCnt = oLV.ListItems.Count + 1
    
    coCurFilterOpt.OptionID = Format(iCnt, "000")
    
    oLI.Tag = CStr(coManFilter.Id)
    
    If Me.cmbAndOrNot.visible = True Then
        coCurFilterOpt.Inclusion = Nz(Me.cmbAndOrNot, "AND")
    End If

    coCurFilterOpt.FilterBy = Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "")
    
    coCurFilterOpt.Operator = Nz(Me.cmbOperator, "=")
    
    If Me.cmbRequiredVal.visible = True Then
        coCurFilterOpt.RequiredValueIdx = Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")
        coCurFilterOpt.RequiredValue = Nz(Me.cmbRequiredVal.Column(1, Me.cmbRequiredVal.ListIndex), "")
    Else
        coCurFilterOpt.RequiredValueIdx = ""
        coCurFilterOpt.RequiredValue = Nz(Me.txtRequiredVal, "")
    End If

    Set GetObjectFromCurrentControls = coCurFilterOpt
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function GetObjectFromListItem(oLI As ListItem) As clsFilterOption
On Error GoTo Block_Err
Dim strProcName As String
Dim oLV As Object
Dim iCnt As Integer
Dim lNewId As Long


    strProcName = ClassName & ".GetObjectFromListItem"

    If Nz(oLI.Tag, "") = "" Then
        Stop
    End If

    ' If the current one is set then we should be able to use it..
'    If coCurFilterOpt Is Nothing Then
'Stop
MakeNew:
        Set coCurFilterOpt = New clsFilterOption 'coManFilter.NewFilterOption(CLng(Me.cmbFilterByField))
        coCurFilterOpt.LoadFromId (CLng(oLI.Tag))
'        GoTo Block_Exit
'    Else
'        If coCurFilterOpt.UnderlyingOptionId <> CLng(Me.cmbFilterByField) Then
'            Stop
'            GoTo MakeNew
'        End If
'    End If
    

    
    Set oLV = Me.lvwBuiltFilter
'
'    If coCurFilterOpt.LoadFromID(CLng(oLI.Tag)) = False Then
'        Stop
'    End If
    coCurFilterOpt.OptionID = Format(oLI.index, "000")
    
    
    Set GetObjectFromListItem = coCurFilterOpt
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function SetControlsFromCurrentObject(Optional oFltrOption As clsFilterOption) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oLV As Object
Dim oLI As ListItem
Dim iCnt As Integer
Dim lNewId As Long


    strProcName = ClassName & ".cmdAddCriteria_Click"
    
    If Not oFltrOption Is Nothing Then
        Stop
        Set coCurFilterOpt = oFltrOption
    End If
    If IsMissing(oFltrOption) = False Then
        Stop
        Set coCurFilterOpt = oFltrOption

    End If
    
    ' If the current one is set then we should be able to use it..
    If coCurFilterOpt Is Nothing Then
'Stop
MakeNew:
        Set coCurFilterOpt = coManFilter.NewFilterOption(CLng(Me.cmbFilterByField))
    Else
        If coCurFilterOpt.UnderlyingOptionId <> CLng(Me.cmbFilterByField) Then
            Stop
            GoTo MakeNew
        End If
    End If



    
    Set oLI = Me.lvwBuiltFilter.SelectedItem
    coCurFilterOpt.LoadFromId (CLng(oLI.Tag))
    
    '' Need to load our combo boxes..
    Call SelectComboBoxByVal(Me.cmbAndOrNot, coCurFilterOpt.Inclusion)

'    Call SelectComboBoxByVal(Me.cmbFilterByField, oLI.SubItems(2))
    Call SelectComboBoxByVal(Me.cmbFilterByField, coCurFilterOpt.FilterBy)

    'Me.cmbOperator = oLI.SubItems(3)
    Me.cmbOperator = coCurFilterOpt.Operator
    
    Call SetRequiredValEntryMode
    
    
    If Me.txtRequiredVal.visible = True Then
'        Me.txtRequiredVal = oLI.SubItems(5)
        Me.txtRequiredVal = coCurFilterOpt.RequiredValue
        Me.cmbRequiredVal = ""
    Else
        Me.txtRequiredVal = ""
        
        Call SelectComboBoxByVal(Me.cmbRequiredVal, coCurFilterOpt.RequiredValue)
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function AddOptionToListView(oFltrOption As clsFilterOption) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oLV As Object
Dim iCnt As Integer

    strProcName = ClassName & ".AddOptionToListView"
    
    Set oLV = Me.lvwBuiltFilter
    iCnt = oLV.ListItems.Count
    iCnt = iCnt + 1

    Set oLI = oLV.ListItems.Add(, , Format(oFltrOption.OptionID, "0##"))
    

'    Set oLI = oLV.ListItems.Add(, , Format(iCnt, "0##"))
    oLI.Tag = CStr(oFltrOption.Id)
    
    oFltrOption.Save = True ' it's in the list so we will save it
    
    If Me.cmbAndOrNot.visible = True Then
        oLI.SubItems(1) = oFltrOption.Inclusion
    End If
    
    oLI.SubItems(2) = oFltrOption.FilterBy
    
    oLI.SubItems(3) = oFltrOption.Operator
    
    If Me.cmbRequiredVal.visible = True Then
        oLI.SubItems(4) = oFltrOption.RequiredValueIdx
        oLI.SubItems(5) = oFltrOption.RequiredValue
        
    Else
        oLI.SubItems(4) = oFltrOption.RequiredValueIdx
        oLI.SubItems(5) = oFltrOption.RequiredValue
    End If

'    oLI.SubItems(6) = Nz(Me.cmbRequiredVal.Column(6, Me.cmbRequiredVal.ListIndex), "")
    'oLI.SubItems(6) = Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "")
    oLI.SubItems(6) = oFltrOption.SampleSource

    Call EnableDisableAndHideControls
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub Form_Timer()

    Me.TimerInterval = 0    ' don't need this anymore..
    Call cmdEditSelected_Click

    
End Sub

Private Sub lvwBuiltFilter_DblClick()
    Call cmdEditSelected_Click
End Sub
