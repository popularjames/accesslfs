Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'Private coManFilter As clsFilter
'Private coCurFilterOpt As clsFilterOption

Private coRule As clsBOLD_LetterRule



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
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_LETTER_Automation_Rules"
        .Parameters.Refresh
        
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    If PopulateListViewFromRs(Me.lvwFilters, oRs) = False Then
        ' may not be any yet..
'        Stop
    End If
    
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


Private Sub cmdAddNew_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sName As String
Dim oAdo As clsADO
Dim lNewId As Long
Dim oDb As DAO.Database
Dim oFrm As Form_frm_BOLD_Letter_Rule_Main

    strProcName = ClassName & ".cmdAddNew_Click"
    
    sName = InputBox("Name of new rule:", "Name filter")
    If Trim(sName) = "" Then
        ' canceled
        GoTo Block_Exit
    End If
    
        
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_LETTER_Automation_CreateNewRule"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pRuleName") = sName
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        lNewId = .Parameters("@pNewId").Value
    End With
    
    Set coRule = New clsBOLD_LetterRule
    If coRule.LoadFromId(lNewId) = False Then
        Stop
    End If
'Stop
    
    Set oFrm = New Form_frm_BOLD_Letter_Rule_Main
    oFrm.RuleObject = coRule
    
        ' we don't want to wait for it to close..
        ' as we want to look at a value from it
        ' so instead of closing it when a user clicks save or cancel
        ' it just makes itself visible = false
        
    Call KDShowFormAndWait(oFrm)
    
    ' if it's still loaded:
    
    If oFrm.Canceled = True Then
   
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_BOLD_LETTER_Automation_RemoveNewRule"
            .Parameters.Refresh
'            .Parameters("@pAccountId") = gintAccountID
            .Parameters("@pRuleId") = lNewId
            .Execute
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                Stop
            End If
        End With

        DoCmd.Close acForm, oFrm.Name, acSaveNo
        
        Set oDb = CurrentDb
        oDb.Execute "DELETE FROM " & cs_TEMP_RULE_TABLE_NAME
        
    End If
    
    Call RefreshData
    
    
    
Block_Exit:
    Set oDb = Nothing
    Set oFrm = Nothing
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdApplyNow_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdEditSelected_Click"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_SetManualOverride"
        .Parameters.Refresh
        .Parameters("@pAccount") = 0
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Problem running sproc", .Parameters("@pErrMsg").Value, True
            GoTo Block_Exit
        End If
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdEditSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim lFilterId As Long
Dim oFrm As Form_frm_BOLD_Letter_Rule_Main

    strProcName = ClassName & ".cmdEditSelected_Click"

    Set oLI = Me.lvwFilters.SelectedItem
    lFilterId = CLng(oLI.Text)
    
    If lFilterId = 0 Then
        Stop
        GoTo Block_Exit
    End If
    
    
    Set oFrm = New Form_frm_BOLD_Letter_Rule_Main
    
    Set coRule = New clsBOLD_LetterRule
    If coRule.LoadFromId(lFilterId) = False Then
        Stop
    End If
  
'    oFrm.OpenArgs = lFilterId
    oFrm.RuleObject = coRule
    Call oFrm.Initialize
    
    
    
    Call KDShowFormAndWait(oFrm)

    
    Call RefreshData


Block_Exit:
    Set oLI = Nothing
    
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
Dim oFilter As clsFilter
Dim lFilterId As Long


    strProcName = ClassName & ".cmdRemoveSelected_Click"
    '   Have to do 2 things here:
        ' 1) Remove from the ListView
        ' 2) Remove from the database
    
    
    Set oLI = Me.lvwFilters.SelectedItem
    ' Should build our object from this LI:
    lFilterId = CLng(oLI.Text)
    
    If lFilterId = 0 Then
        Stop
    End If
    
    
    
    Set oFilter = New clsFilter
    If oFilter.LoadFromId(lFilterId) = False Then
        Stop
    End If
    
    ' Delete it from the database
    ' but first lets prompt to make sure:
    If MsgBox("Are you sure you wish to delete the '" & oFilter.FilterName & "' filter?", vbYesNo, "Delete?") = vbNo Then
        GoTo Block_Exit
    End If
    
    If oFilter.Delete = False Then
        Stop
        GoTo Block_Exit
    End If
        
    
    ' Now we can remove it from the list
    lvwFilters.ListItems.Remove (oLI.index)
'    Call RenumberItems
'    Call cmdViewCollectiveSample_Click
    Call RefreshData
    

Block_Exit:
    Set oLI = Nothing
    Set oFilter = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


'''
'''Private Function LoadCurrentFilterOption(Optional bIsNew As Boolean = False) As Boolean
'''Dim oLV As Object
'''Dim iCnt As Integer
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim lNewId As Long
'''
'''    strProcName = ClassName & ".LoadCurrentFilterOption"
'''Stop
'''
'''    Set coCurFilterOpt = New clsFilterOption
''''    coCurFilterOpt.ManFilterID = coManFilter.ManFilterID
'''
'''    coCurFilterOpt.ManFilterID = 1
'''
'''    coCurFilterOpt.UnderlyingOptionId = CLng(Me.cmbFilterByField)
'''
'''    If bIsNew = True Then
'''        lNewId = coCurFilterOpt.IsNew
'''    End If
'''
'''    Set oLV = Me.lvwBuiltFilter
'''    iCnt = oLV.ListItems.Count + 1
'''
'''    coCurFilterOpt.OptionId = Format(iCnt, "000")
'''
'''
'''    If Me.cmbAndOrNot.visible = True Then
'''        coCurFilterOpt.Inclusion = Nz(Me.cmbAndOrNot, "AND")
'''    End If
'''
'''    coCurFilterOpt.FilterBy = Nz(Me.cmbFilterByField.Column(1, Me.cmbFilterByField.ListIndex), "")
'''
'''    coCurFilterOpt.Operator = Nz(Me.cmbOperator, "=")
'''
'''    If Me.cmbRequiredVal.visible = True Then
'''        coCurFilterOpt.RequiredValueIdx = Nz(Me.cmbRequiredVal.Column(0, Me.cmbRequiredVal.ListIndex), "")
'''        coCurFilterOpt.RequiredValue = Nz(Me.cmbRequiredVal.Column(1, Me.cmbRequiredVal.ListIndex), "")
'''
'''    Else
'''        coCurFilterOpt.RequiredValueIdx = ""    '   Nz(Me.cmbFilterByField.Column(3, Me.cmbFilterByField.ListIndex), "")
'''        coCurFilterOpt.RequiredValue = Nz(Me.txtRequiredVal, "")
'''    End If
'''
'''
''''    oLI.SubItems(6) = Nz(Me.cmbRequiredVal.Column(6, Me.cmbRequiredVal.ListIndex), "")
''''    oLI.SubItems(6) = Nz(Me.cmbFilterByField.Column(5, Me.cmbFilterByField.ListIndex), "")
'''
''''    Me.cmbAndOrNot.visible = True
'''
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    GoTo Block_Exit
'''End Function




Private Sub Form_Load()
    If Me.OpenArgs <> "" Then
Stop
        Set coRule = New clsBOLD_LetterRule
        coRule.LoadFromId (CLng(Me.OpenArgs))
        ' then we want the LAST one to load into the controls
        
    End If

    RefreshData
End Sub


Private Sub Form_Resize()
'Dim dGridSpacer As Double
'
'    ' Very basic stuff here..
'    ' only going to do the main grid and the provider addresses grid
'    If Me.InsideWidth < 12840 Then
'        Exit Sub
'    End If
'
'    dGridSpacer = (Me.oMainGrid.left * 2)
'
'    Me.oMainGrid.Width = Me.InsideWidth - dGridSpacer
'    Me.oProvAddresses.Width = Me.InsideWidth - dGridSpacer
'
'    ' Alright, screw it, it looks un finished so we'll do the top stuff too
'
'    Me.fraRightButtons.left = Me.InsideWidth - Me.fraRightButtons.Width - dGridSpacer
'    Me.lvwBuiltFilter.Width = Me.fraRightButtons.left - (dGridSpacer * 2)
'
'    Me.cmdViewCollectiveSample.left = Me.fraRightButtons.left + dGridSpacer
'    Me.cmdEditSelected.left = Me.cmdViewCollectiveSample.left
'    Me.cmdRemoveSelected.left = Me.cmdViewCollectiveSample.left
'    Me.cmdSave.left = Me.cmdViewCollectiveSample.left
'    Me.CmdCancel.left = Me.cmdViewCollectiveSample.left
'
'    Me.fraUpperRightButtons.Width = Me.InsideWidth - (Me.fraUpperLeftButtons.Width + Me.fraUpperLeftButtons.left) - dGridSpacer
'
    
End Sub

Public Function GetFilterFromListItem(oLI As ListItem) As clsFilter
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetFilterFromListItem"
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub lvwFilters_DblClick()
    Call cmdEditSelected_Click
End Sub
