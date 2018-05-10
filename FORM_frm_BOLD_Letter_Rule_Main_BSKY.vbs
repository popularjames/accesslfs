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

Private clRuleId As Long


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get RuleId() As Long
    RuleId = clRuleId
End Property
Public Property Let RuleId(lRuleId As Long)
    clRuleId = lRuleId
End Property


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & "RefreshData"
    
    ' so, we need to populate the subforms
    Set coGroupBy = Me.osfrmPer.Form
    Set coApplyTo = Me.ofrmApplyTo.Form
    
    
    coGroupBy.MainCaption = "Per"
    coApplyTo.MainCaption = "Apply to:"
    
    Set oAdo = New clsADO
    With oAdo
        ''.ConnectionString = CurrentProject.Connection
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT GroupById, GroupByName, FieldName FROM BOLD_LETTER_Automation_Req_Xref_GroupBy WHERE Active <> 0  ORDER BY GroupByName"
'        .sqlString = "SELECT GroupById, GroupByName, FieldName FROM BOLD_LETTER_Automation_Req_Xref_GroupBy WHERE Active <> 0  ORDER BY GroupById"
        Set oRs = .ExecuteRS
    End With
    
    Call coGroupBy.InitData(oRs, "GroupBy", "GroupById", "GroupByName")
    
    
    Set oAdo = New clsADO
    With oAdo
        ''.ConnectionString = CurrentProject.Connection
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT ApplyToId, ApplyToName, FieldName FROM BOLD_LETTER_Automation_Req_Xref_ApplyTo WHERE Active <> 0  ORDER BY ApplyToName"
        Set oRs = .ExecuteRS
    End With
    
    Call coApplyTo.InitData(oRs, "ApplyTo", "ApplyToId", "ApplyToName")
    
    '' Now need to get the combo box:
    Call RefreshComboBoxADO("SELECT ObjectToLimitId, ObjectName, FieldName FROM BOLD_LETTER_Automation_Xref_ObjectsToLimit WHERE Active <> 0 ORDER BY ObjectName", Me.cmbWhatToLimit, , , "v_Data_Database")
    
    Call RefreshComboBoxADO("SELECT FormatId, Name FROM BOLD_Letter_Automation_XREF_Formats ORDER BY FormatId", Me.cmbFinalFormat, 1, "FormatId", "v_Data_Database")
    
'
'    Me.cmbWhatToLimit.RowSource = "SELECT ObjectToLimitId, ObjectName FROM BOLD_LETTER_Automation_Xref_ObjectsToLimit WHERE Active <> 0 ORDER BY ObjectName"
        
    ' so, now, if we loaded an existing rule, then we need to remove the ListItem and populate it in the object
    If Me.OpenArgs <> "" Then
            Stop
    End If
    
Block_Exit:
    '   CurrentProject.CloseConnection ???
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub cmdSaveRule_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdSaveRule"
    
    ' we need a RuleId already.. so when we open this we'll get that
    

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
Dim oDb As DAO.Database

    strProcName = ClassName & ".Form_Load"
    Set coLetterRule = New clsBOLD_LetterRule
    
    Set oDb = CurrentDb()
    oDb.Execute "DELETE FROM tmp_BOLD_LETTER_Rules"
    
'    If Me.OpenArgs <> "" Then
        coLetterRule.LoadFromId 1 ' (CLng(Me.OpenArgs))
'    Else
        ' secure a new id
        
'    End If
        
    Call RefreshData
Block_Exit:
    Set oDb = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Sub
