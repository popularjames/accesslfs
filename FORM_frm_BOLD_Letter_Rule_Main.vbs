Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'' Last Modified: 04/24/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/13/2013  - Created
''
'' AUTHOR
''  =====================================
'' Kevin Dearing
''
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################


Private WithEvents coGroupBy As Form_frm_GENERAL_Select_List_ADO_N_Criteria
Attribute coGroupBy.VB_VarHelpID = -1
Private WithEvents coApplyTo As Form_frm_GENERAL_Select_List_ADO_N_Criteria
Attribute coApplyTo.VB_VarHelpID = -1

Private WithEvents coLetterRule As clsBOLD_LetterRule
Attribute coLetterRule.VB_VarHelpID = -1

Private csGroupByText As String
Private csApplyToText As String
Private cbEditing As Boolean
Private cbCanceled As Boolean

Private clRuleId As Long


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get Canceled() As Boolean
    Canceled = cbCanceled
End Property
Public Property Let Canceled(bCanceled As Boolean)
    cbCanceled = bCanceled
End Property

Public Property Get RuleId() As Long
    RuleId = clRuleId
End Property
Public Property Let RuleId(lRuleId As Long)
    clRuleId = lRuleId
End Property

Public Property Get SelectedId() As Long
    SelectedId = RuleId
End Property


Public Property Get RuleObject() As clsBOLD_LetterRule
    Set RuleObject = coLetterRule
End Property
Public Property Let RuleObject(oRule As clsBOLD_LetterRule)
    Set coLetterRule = oRule
    Call PopulateFromRuleObject(oRule)
End Property

Private Function PopulateFromRuleObject(oRule As clsBOLD_LetterRule) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".PopulateFromRuleObject"
    
    If Not coGroupBy Is Nothing Then
         coGroupBy.ItemDetails = oRule.GetRuleItemDetails()
    End If
    If Not coApplyTo Is Nothing Then
        coApplyTo.ItemDetails = oRule.GetRuleItemDetails()
    End If
    
    Me.txtRuleName = oRule.RuleName
    Me.txtQty = oRule.Qty
    
    Me.cmbWhatToLimit = oRule.ObjectIdToLimit
    
    Me.cmbFinalFormat = oRule.FinalFormatId
    
    
Block_Exit:
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

    strProcName = ClassName & "RefreshData"

    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT GroupById, GroupByName, FieldName FROM BOLD_LETTER_Automation_Req_Xref_GroupBy WHERE Active <> 0  ORDER BY GroupByName"
        Set oRs = .ExecuteRS
    End With
    
    Call coGroupBy.InitData(oRs, "GroupBy", "GroupById", "GroupByName")
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT ApplyToId, ApplyToName, FieldName FROM BOLD_LETTER_Automation_Req_Xref_ApplyTo WHERE Active <> 0  ORDER BY ApplyToName"
        Set oRs = .ExecuteRS
    End With
    
    Call coApplyTo.InitData(oRs, "ApplyTo", "ApplyToId", "ApplyToName")
    
    '' Now need to get the combo box:
    Call RefreshComboBoxADO("SELECT ObjectToLimitId, ObjectName, FieldName FROM BOLD_LETTER_Automation_Xref_ObjectsToLimit WHERE Active <> 0 ORDER BY ObjectName", Me.cmbWhatToLimit, , , "v_Data_Database")
    
    Call RefreshComboBoxADO("SELECT FormatId, Name FROM BOLD_Letter_Automation_XREF_Formats ORDER BY FormatId", Me.cmbFinalFormat, 1, "FormatId", "v_Data_Database")
    
        
    ' so, now, if we loaded an existing rule, then we need to remove the ListItem and populate it in the object
    If Me.OpenArgs <> "" Then
            Stop
    End If
'    Stop
    
Block_Exit:
    Set oAdo = Nothing
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub CmdCancel_Click()
    Me.Canceled = True
    Me.visible = False
End Sub

Private Sub cmdSaveRule_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oLetterDetails As clsBOLD_LetterRuleItemDetails

    strProcName = ClassName & ".cmdSaveRule"
    
    ' we need a RuleId already.. so when we open this we'll get that
    If coLetterRule Is Nothing Then
        Stop
        Set coLetterRule = New clsBOLD_LetterRule
    End If
    
    If coLetterRule.Id = 0 Then
        Stop
        If coLetterRule.CreateNew(Me.txtRuleName, CLng("0" & Nz(Me.txtQty, "")), Me.cmbWhatToLimit, Me.cmbFinalFormat) < 1 Then
            Stop
        End If
        
Stop
        
        Set oLetterDetails = New clsBOLD_LetterRuleItemDetails
'        oLetterDetails.LoadFromItemType ("GroupBy")
        oLetterDetails.LoadFromItemType ("")    ' we want both (or all) types..
        
        coLetterRule.Details = oLetterDetails
        
        coLetterRule.SaveNow
    Else
        ' editing the Rule?
        Stop
        
    End If

    Me.Canceled = False
    Me.visible = False

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub coApplyTo_RuleChanged(sEnglish As String, sSql As String)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".coGroupByRuleChanged"
    csApplyToText = sEnglish
    
    DisplayEnglishText
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub coGroupBy_RuleChanged(sEnglish As String, sSql As String)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".coGroupByRuleChanged"
    csGroupByText = sEnglish
    
    DisplayEnglishText
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub DisplayEnglishText()
On Error GoTo Block_Err
Dim strProcName As String
Dim sFullText As String

    strProcName = ClassName & ".DisplayEnglishText"
    If CInt(Nz(Me.txtQty, 0)) > 1 Then
        sFullText = "Limit to " & CStr(Nz(Me.txtQty, "")) & " " & Me.cmbWhatToLimit.Column(1) & "s" & vbCrLf & "PER: " & csGroupByText
    Else
        sFullText = "Limit to " & CStr(Nz(Me.txtQty, "")) & " " & Me.cmbWhatToLimit.Column(1) & vbCrLf & "PER: " & csGroupByText
    End If
    
    
    If csApplyToText <> "" Then
        sFullText = sFullText & vbCrLf & "Apply to: " & csApplyToText
    End If

    If Me.cmbFinalFormat.Column(1) <> "" Then
        sFullText = sFullText & vbCrLf & "And send as a " & Me.cmbFinalFormat.Column(1)
    End If

    Me.lblEnglishVersion.Caption = sFullText

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
        
    Call RefreshData
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Sub



Public Function Initialize() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDtl As clsBOLD_LetterRuleItemDetail
Dim oDtls As clsBOLD_LetterRuleItemDetails

    strProcName = ClassName & ".Initialize"
    
    cbEditing = True

    If coLetterRule Is Nothing Then
        GoTo Block_Exit
    End If
    If coLetterRule.Id = 0 Then
        Stop
        GoTo Block_Exit
    End If

    Set oDtls = coLetterRule.Details
    
    For Each oDtl In oDtls.Items
        If oDtl.ItemType = "GroupBy" Then
            Call coGroupBy.MoveItemDetailToSelected(oDtl)
            
        Else
            Call coApplyTo.MoveItemDetailToSelected(oDtl)
        
        End If
    Next
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


' This is going to save it to our temp table..
Public Function Save_Backup() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database

    strProcName = ClassName & ".Save_Backup"
    
Stop
    
    Set oDb = CurrentDb
    oDb.Execute "DELETE FROM " & cs_TEMP_RULE_TABLE_NAME
    
    cbEditing = True
    ' bottom line, populate everything from the global coManFilter
    If coLetterRule Is Nothing Then
        GoTo Block_Exit
    End If
    If coLetterRule.Id = 0 Then
        Stop
        GoTo Block_Exit
    End If
    

Block_Exit:
    Set oDb = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub Form_Open(Cancel As Integer)
    Set coLetterRule = New clsBOLD_LetterRule
    
    ' so, we need to populate the subforms
    Set coGroupBy = Me.osfrmPer.Form
    Set coApplyTo = Me.ofrmApplyTo.Form
    
    
    coGroupBy.MainCaption = "Per"
    coApplyTo.MainCaption = "Apply to:"
    
    If Me.OpenArgs <> "" Then
        Stop
        coLetterRule.LoadFromId (CLng(Me.OpenArgs))
    Else
'         secure a new id
    End If
    
End Sub
