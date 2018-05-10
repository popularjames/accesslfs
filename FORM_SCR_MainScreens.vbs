Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------
'Author :
'Description:
'Create Date:
'Last Modified:
' PD,Oct 20 2011 - CR # 2564 fix - Clear phantom filters during refresh and toggle of 'Filter/Unfiltered' widget
' PD,Oct 22 2011 - CR # 2574 fix - Filters with blank criteria should not be allowed
' SA 8/6/2012 - Added app name to telemetry calls
' SA 11/7/2012 - Moved some methods into SCR_ClsMainScreens to reduce code in form and memory used
'---------------------------------------------------------------------------------------------------------------------------------
Private ClsSCR As New SCR_ClsMainScreens

#If ccDT = 1 Then
Private WithEvents criteriaForm As Form_DT_CustomCriteriaSelection
Attribute criteriaForm.VB_VarHelpID = -1
#End If

Private Const DecipherRestoreCollapseMenu As String = "DecipherRestoreCollapseMenu"
Private mvConfig As CnlyScreenCfg
Private mvSql As CnlyScreenSQL
Private WithEvents MvGridMain As Form_CT_SubGenericDataSheet
Attribute MvGridMain.VB_VarHelpID = -1
Private MvSplitY As Single
Private WithEvents filterForm As Form_SCR_ScreensFilters
Attribute filterForm.VB_VarHelpID = -1
Private WithEvents stateForm As Form_SCR_State
Attribute stateForm.VB_VarHelpID = -1
Private genUtils As New CT_ClsGeneralUtilities
Private holdForm As Object

Private WithEvents restoreButton As CommandBarButton
Attribute restoreButton.VB_VarHelpID = -1
Private WithEvents collapseAllButton As CommandBarButton
Attribute collapseAllButton.VB_VarHelpID = -1
Private WithEvents collapseRecordSetButton As CommandBarButton
Attribute collapseRecordSetButton.VB_VarHelpID = -1
Private restoreMenu As CommandBar
'Private TogValue As Boolean
' hc 3/15/2011 -- added to keep track of the orginal sql filter used on the refresh
Private selectionFilter As String

' DS Mar 12 2010 save filter during restores/splitter resize
Private MvFilter As String

'SA 1/19/2012 - CR1967 Added to fix problem of filters being lost on tab changes
Private TabLoaded() As Boolean

'SA 05/21/2012 - CR2132 Added row change event and variable to track it
Private HasRowChangeEvent As Boolean

Public Event isVisible(data As Boolean)

Public Function GetCustomCriteriaSource() As String
    GetCustomCriteriaSource = ClsSCR.GetCustomCriteriaSource(mvConfig)
End Function

Public Sub ApplyLayout()
    cmdLayoutApply_Click
End Sub

Public Function BuildWhere() As String
    BuildWhere = ClsSCR.BuildWhere(mvSql)
End Function

Public Property Get GridForm() As Form_CT_SubGenericDataSheet
    If Not MvGridMain Is Nothing Then
        Set MvGridMain = Me.SubForm.Form
    End If
    Set GridForm = MvGridMain
End Property
Public Property Get SQL() As CnlyScreenSQL
    SQL = mvSql
End Property
Public Property Let SQL(data As CnlyScreenSQL)
    mvSql = data
End Property
Public Property Let Config(data As CnlyScreenCfg)
    mvConfig = data
End Property
Public Property Get Config() As CnlyScreenCfg
   Config = mvConfig
End Property

'DLC 11/13/2012 - Added to support Audit/Platform selection
Public Property Get Platform() As String
   Platform = mvConfig.Platform
End Property
Public Property Let Platform(Value As String)
   mvConfig.Platform = Value
End Property

Public Function BuildWherePrimary() As String
    BuildWherePrimary = ClsSCR.BuildWherePrimary(mvConfig)
End Function

Public Function BuildWhereSecondary() As String
    BuildWhereSecondary = ClsSCR.BuildWhereSecondary(mvConfig)
End Function

Public Function BuildWhereTertiary() As String
    BuildWhereTertiary = ClsSCR.BuildWhereTertiary(mvConfig)
End Function

Public Function BuildMultiItemSQL(ctlList As MSForms.listBox, txtQual As String) As String
    BuildMultiItemSQL = ClsSCR.BuildMultiItemSQL(ctlList, txtQual)
End Function

Private Sub SetListBy()
Dim X As Integer
For X = 0 To Me.CmboListPrimaryBy.ListCount - 1
    If Me.CmboListPrimaryBy.Column(1, X) Then
        Me.CmboListPrimaryBy = Me.CmboListPrimaryBy.Column(0, X)
        Exit For
    End If
Next X
CmboListPrimaryBy_AfterUpdate


For X = 0 To Me.CmboListSecondaryBy.ListCount - 1
    If Me.CmboListSecondaryBy.Column(1, X) Then
        Me.CmboListSecondaryBy = Me.CmboListSecondaryBy.Column(0, X)
        Exit For
    End If
Next X

'  ** Added Tertiary
For X = 0 To Me.CmboListTertiaryBy.ListCount - 1
    If Me.CmboListTertiaryBy.Column(1, X) Then
        Me.CmboListTertiaryBy = Me.CmboListTertiaryBy.Column(0, X)
        Exit For
    End If
Next X
    If Me.CmboListTertiaryBy <> vbNullString Then
        CmboListTertiaryBy_AfterUpdate
    End If
End Sub
Property Get PrimaryCriteria()
    PrimaryCriteria = mvConfig.PrimaryRecordSource
End Property

Property Let ScreenName(Criteria As String)
    mvConfig.ScreenName = Criteria
End Property
Property Let ScreenID(Criteria As Long)
    mvConfig.ScreenID = Criteria
End Property
Property Get ScreenID() As Long
    ScreenID = mvConfig.ScreenID
End Property
Property Let FormID(Criteria As Long)
    mvConfig.FormID = Criteria
End Property
Property Get FormID() As Long
    FormID = mvConfig.FormID
End Property
Property Get ScreenName() As String
    ScreenName = mvConfig.ScreenName
End Property
Private Sub SortListMoveItem(MoveCount As Integer)
On Error GoTo ErrorHandler
    'SA 03/22/2012 - Changed SortList from ActiveX to Access Listbox
    Dim FieldName As String
    Dim SortDir As String
    Dim CurIndex As Integer
    Dim newIndex As Integer
    
    CurIndex = Me.SortList.ListIndex
    If CurIndex > -1 Then
        FieldName = Me.SortList.Column(1, Me.SortList.ListIndex)
        SortDir = Me.SortList.Column(0, Me.SortList.ListIndex)
        
        'Figure out where the item should go
        If CurIndex + MoveCount < 0 Then
            newIndex = 0
        ElseIf CurIndex + MoveCount >= Me.SortList.ListCount - 1 Then
            newIndex = Me.SortList.ListCount - 1
        Else
            newIndex = CurIndex + MoveCount
        End If
    
        If newIndex <> CurIndex Then
            Me.SortList.RemoveItem CurIndex
            Me.SortList.AddItem SortDir & ";" & FieldName, newIndex
            Me.SortList.Selected(newIndex) = True
        End If
    End If
    
ExitSort:
    ClsSCR.UpdateSortTip
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Sorting"
    Resume ExitSort
    Resume
End Sub
Public Sub BuildDetail()
    selectionFilter = ClsSCR.BuildDetail(mvSql, mvConfig)
End Sub

Private Sub PopulateLists()
'On Error Resume Next

txtFocus.SetFocus

LblPrimary.Caption = mvConfig.PrimaryListBoxCaption
If mvConfig.PrimaryListBoxMulti = True Then
    CmboPrimaryMulti.visible = True
    lblPrimaryMulti.visible = True
    CmboPrimaryMulti.Enabled = True
    CmdBuildMulti1.visible = True
    CmdBuildMulti1.Enabled = True
    CmboPrimary.Enabled = False
    CmboPrimary.visible = False
    LblPrimaryAlternate.visible = False
    CmboListPrimaryBy.Enabled = True
    CmdZoomMulti1.visible = True
Else
    CmboPrimaryMulti.Enabled = False
    CmboPrimaryMulti.visible = False
    lblPrimaryMulti.visible = False
    CmdBuildMulti1.Enabled = False
    CmdBuildMulti1.visible = False
    CmboPrimary.visible = True
    CmboPrimary.Enabled = True
    LblPrimaryAlternate.visible = True
    Me.LblPrimary.Caption = mvConfig.PrimaryListBoxCaption
    LblPrimaryAlternate.Caption = vbNullString
    CmdZoomMulti1.visible = False
End If

If mvConfig.SecondaryListBoxUse = False Then
    Me.CmboSecondary.Enabled = False
    Me.CmboListSecondaryBy.Enabled = False
    Me.LblSecondaryAlternate.visible = False
    Me.CmboSecondaryMulti.visible = False
    Me.lblSecondaryMulti.visible = False
    Me.CmdBuildMulti2.visible = False
    Me.CmdZoomMulti2.visible = False
Else
    If mvConfig.SecondaryListBoxMulti = True Then
        CmboSecondaryMulti.Enabled = True
        CmboSecondaryMulti.visible = True
        lblSecondaryMulti.visible = True
        CmdBuildMulti2.visible = True
        CmdBuildMulti2.Enabled = True
        CmboSecondary.Enabled = False
        CmboSecondary.visible = False
        LblSecondaryAlternate.visible = False
        CmboListSecondaryBy.Enabled = True
        Me.CmdZoomMulti2.visible = True
    Else
        'MULTI SELECT STUFF
        CmboSecondaryMulti.Enabled = False
        CmboSecondaryMulti.visible = False
        lblSecondaryMulti.visible = False
        CmdBuildMulti2.Enabled = False
        CmdBuildMulti2.visible = False
        CmboSecondary.visible = True
        CmboSecondary.Enabled = True
        LblSecondaryAlternate.visible = True
        LblSecondaryAlternate.Caption = vbNullString
        CmdZoomMulti2.visible = False
    End If
    LblSecondary.visible = True
    LblSecondary.Caption = mvConfig.SecondaryListBoxCaption
End If

' ** Added Tertiary **
If mvConfig.TertiaryListBoxUse = False Then
    Me.CmboTertiary.Enabled = False
    Me.CmboListTertiaryBy.Enabled = False
    Me.LblTertiaryAlternate.visible = False
    Me.CmboTertiaryMulti.visible = False
    Me.lblTertiaryMulti.visible = False
    Me.CmdBuildMulti3.visible = False
    Me.CmdZoomMulti3.visible = False
Else
    If mvConfig.TertiaryListBoxMulti = True Then
        CmboTertiaryMulti.Enabled = True
        CmboTertiaryMulti.visible = True
        lblTertiaryMulti.visible = True
        CmdBuildMulti3.visible = True
        CmdBuildMulti3.Enabled = True
        CmboTertiary.Enabled = False
        CmboTertiary.visible = False
        LblTertiaryAlternate.visible = False
        CmboListTertiaryBy.Enabled = True
        Me.CmdZoomMulti3.visible = True
    Else
        'MULTI SELECT STUFF
        CmboTertiaryMulti.Enabled = False
        CmboTertiaryMulti.visible = False
        lblTertiaryMulti.visible = False
        CmdBuildMulti3.Enabled = False
        CmdBuildMulti3.visible = False
        CmboTertiary.visible = True
        CmboTertiary.Enabled = True
        CmboListTertiaryBy.Enabled = True
        LblTertiaryAlternate.visible = True
        LblTertiaryAlternate.Caption = vbNullString
        Me.CmdZoomMulti3.visible = False
    End If
    LblTertiary.visible = True
    LblTertiary.Caption = mvConfig.TertiaryListBoxCaption
End If

If mvConfig.DateUse Then
    PopulateListDateFilters Me.CmboFilterDte, mvConfig.ScreenID
    Me.CmboFilterDte.Enabled = True
    Me.StartDte.Enabled = True
    Me.EndDte.Enabled = True
Else
    Me.CmboFilterDte.Enabled = False
    Me.StartDte.Enabled = False
    Me.EndDte.Enabled = False
End If

' populate the label
Me.lblBanner.Caption = "Screen: " & mvConfig.ScreenName

'Populate Sort Field Lists
Me.CmboSortFieldList.RowSource = mvConfig.PrimaryRecordSource

'POPULATE THE LISTS
PopulateListFunctions Me.CmboFunction, mvConfig.ScreenID
PopulateListTotals Me.CmboTotals, mvConfig.ScreenID
PopulateListReports Me.CmboReports, mvConfig.ScreenID
ClsSCR.PopulateScreenFilters Me.CmboFilters, mvConfig.ScreenID
PopulateListSelects Me.CmboListPrimaryBy, mvConfig.ScreenID, 1
PopulateListSelects Me.CmboListSecondaryBy, mvConfig.ScreenID, 2
PopulateListSelects Me.CmboListTertiaryBy, mvConfig.ScreenID, 3
PopulateListLayouts Me.CmboLayouts, mvConfig.ScreenID

'IF A DEFAULT SORT EXISTS THEN LOAD IT
ClsSCR.SetDefaultSort mvConfig
End Sub

Public Sub RunReport(ListName As String, ReportID As Long)
'SA 03/22/2012 - CR2000 Changed sub to public
'SA 05/12/2012 - CR2131 Added pre and post events
On Error GoTo RunReportError
    Dim RptConfig As New CT_ClsRpt
    Dim db As DAO.Database
    Set db = CurrentDb

    RunEvent "Report Pre-Run", Me.ScreenID, Me.FormID
    
    Call GetReportCfg(ReportID, RptConfig)
    
    'If the global variables in scope have been reset - get them again
    ' delete the current report selections from ReportCriteria
    db.Execute "DELETE FROM SCR_ScreensReportCriteria " & _
            " WHERE Auditor=" & Chr(34) & Identity.Auditor & Chr(34) & " AND RptId=" & ReportID
            
    
    'Primary Criteria
    ' ADDED BY LINO TO FIX LONG FILTER ERROR
    '       The max length of the where clause passed to an Access Report is 32,768.
    '       If this limit would be exceeded by the existing WhereSecondary, build a temp
    '       table to hold the "in" values.
    'Moved to a routine - HC 11/2/2008
    If RptConfig.EnablePrimary And LenB(mvSql.WherePrimary) > 0 Then
        If mvConfig.PrimaryListBoxMulti Then
            If ClsSCR.GetCurrentSQLStringLength(mvSql) > 32767 Then
                ClsSCR.SetReportCriteria RptConfig, ReportID, CmboPrimaryMulti, CmboListPrimaryBy, 1
            Else
                RptConfig.Criteria = mvSql.WherePrimary
            End If
        Else
            RptConfig.Criteria = mvSql.WherePrimary
        End If
    End If
    
    'Secondary Criteria
    If RptConfig.EnableSecondary And mvConfig.SecondaryListBoxUse Then
        If LenB(mvSql.WhereSecondary) > 0 Then
            If mvConfig.SecondaryListBoxMulti Then
                If ClsSCR.GetCurrentSQLStringLength(mvSql) > 32767 Then
                    ClsSCR.SetReportCriteria RptConfig, ReportID, CmboSecondaryMulti, CmboListSecondaryBy, 2
                Else
                    If vbNullString & RptConfig.Criteria <> vbNullString Then
                        RptConfig.Criteria = RptConfig.Criteria & " AND "
                    End If
                    RptConfig.Criteria = RptConfig.Criteria & mvSql.WhereSecondary
                End If
            Else
                If vbNullString & RptConfig.Criteria <> vbNullString Then
                    RptConfig.Criteria = RptConfig.Criteria & " AND "
                End If
                RptConfig.Criteria = RptConfig.Criteria & mvSql.WhereSecondary
            End If
        End If
    End If
    
    If RptConfig.EnableTertiary And mvConfig.TertiaryListBoxUse Then
        If LenB(mvSql.WhereTertiary) > 0 Then
            If mvConfig.TertiaryListBoxMulti Then
                If ClsSCR.GetCurrentSQLStringLength(mvSql) > 32767 Then
                    ClsSCR.SetReportCriteria RptConfig, ReportID, CmboTertiaryMulti, CmboListTertiaryBy, 3
                Else
                    If vbNullString & RptConfig.Criteria <> vbNullString Then
                        RptConfig.Criteria = RptConfig.Criteria & " and "
                    End If
                    RptConfig.Criteria = RptConfig.Criteria & mvSql.WhereTertiary
                End If
            Else
                If vbNullString & RptConfig.Criteria <> vbNullString Then
                    RptConfig.Criteria = RptConfig.Criteria & " and "
                End If
                RptConfig.Criteria = RptConfig.Criteria & mvSql.WhereTertiary
            End If
        End If
    End If
    
    
    'Build Criteria If applicable
    If RptConfig.EnableFilter Then
        If vbNullString & mvSql.WhereDates <> vbNullString Then
            If vbNullString & RptConfig.Criteria <> vbNullString Then
                RptConfig.Criteria = RptConfig.Criteria & " and "
            End If
            RptConfig.Criteria = RptConfig.Criteria & mvSql.WhereDates
        End If
    End If
    
    'Build Sort Order
    ' DS Mar 11 2010 Dave B fix to change request ID#1430- Bug in RunReports - Access is passing the form name along with the field name with the “OrderBy” parameter
    If RptConfig.EnableSort Then
        'rptConfig.SortString = IIf(Me.GridForm.OrderByOn = True, Me.GridForm.ORDERBY, MvSql.ORDERBY)
        RptConfig.SortString = IIf(Me.GridForm.OrderByOn = True, Replace(Me.GridForm.OrderBy, "[" & Me.GridForm.Name & "].", vbNullString), mvSql.OrderBy)
    End If
    
    'Extra SQL
    If vbNullString & RptConfig.ExtraSQL <> vbNullString Then
            If vbNullString & RptConfig.Criteria <> vbNullString Then
                RptConfig.Criteria = RptConfig.Criteria & " and "
            End If
            RptConfig.Criteria = RptConfig.Criteria & RptConfig.ExtraSQL
    End If
    RptConfig.Criteria = BuildCriteriaFromSubform(Me.SubForm, ReportID, RptConfig.Criteria)
    
    'Now add the filter
    If RptConfig.EnableFilter And vbNullString & mvSql.filter <> vbNullString Then
        If vbNullString & RptConfig.Criteria <> vbNullString Then
            RptConfig.Criteria = RptConfig.Criteria & " and "
        End If
        RptConfig.Criteria = RptConfig.Criteria & "(" & mvSql.filter & ") "
    End If
    
    If RptConfig.Criteria <> "ERROR SETTING CRITERIA" Then
        DoCmd.Close acReport, RptConfig.ReportName, acSaveNo
    
    
        'FIRE THE REPORT RUN EVENT
        RunEvent "Report Run", Me.ScreenID, Me.FormID, RptConfig
        
        'If the OpenArgs were not populated by a "Report Run" event handler, set it to the FormID
        If RptConfig.OpenArgs = vbNullString Then
            RptConfig.OpenArgs = mvConfig.FormID
        End If
        
        ' Changed the open report form to place the criteria in the filter rather than in the where clause.
        'DoCmd.OpenReport rptConfig.ReportName, acViewPreview, rptConfig.Criteria, , acWindowNormal, rptConfig.OpenArgs
        DoCmd.OpenReport RptConfig.ReportName, acViewPreview, , RptConfig.Criteria, acWindowNormal, RptConfig.OpenArgs
        If LenB(RptConfig.SortString) > 0 And RptConfig.EnableSort Then
           Reports(RptConfig.ReportName).OrderBy = RptConfig.SortString
           Reports(RptConfig.ReportName).OrderByOn = True
        End If
    End If

RunReportExit:
    On Error Resume Next
    'SA 05/12/2012 - CR2131 Added event
    RunEvent "Report Post-Run", Me.ScreenID, Me.FormID, RptConfig
    Set db = Nothing
Exit Sub
RunReportError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error running report!", vbCritical, "Run Report Error"
    Resume RunReportExit
    Resume
End Sub

Private Sub CmboFilterDte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
CmboFilterDte.Dropdown
End Sub

Private Sub CmboFilters_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboFilters.Dropdown
End Sub

Private Sub CmboFunction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
    CmboFunction.Dropdown
End Sub

Private Sub cmboLayouts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboLayouts.Dropdown
End Sub

Private Sub CmboListPrimaryBy_AfterUpdate()
'On Error Resume Next
    ClsSCR.PopulateCriteriaLists Me.CmboListPrimaryBy, CmboPrimary, 1, mvConfig
    CmboPrimary_AfterUpdate
End Sub

Private Sub CmboListPrimaryBy_Enter()
'On Error Resume Next
Me.CmboListPrimaryBy.Dropdown
End Sub

Private Sub CmboListPrimaryBy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboListPrimaryBy.Dropdown
End Sub

Private Sub CmboListTertiaryBy_AfterUpdate()
'On Error Resume Next
    ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
    CmboTertiary_AfterUpdate
End Sub

Private Sub CmboListTertiaryBy_Enter()
'    On Error Resume Next
    Me.CmboListTertiaryBy.Dropdown
End Sub

Private Sub CmboListTertiaryBy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
    Me.CmboListTertiaryBy.Dropdown
End Sub
Private Sub CmboListSecondaryBy_AfterUpdate()
'On Error Resume Next
    ClsSCR.PopulateCriteriaLists Me.CmboListSecondaryBy, Me.CmboSecondary, 2, mvConfig
    CmboSecondary_AfterUpdate
End Sub
Private Sub CmboListSecondaryBy_Enter()
'On Error Resume Next
Me.CmboListSecondaryBy.Dropdown
End Sub

Private Sub CmboListSecondaryBy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboListSecondaryBy.Dropdown
End Sub

Public Sub CmboPrimary_AfterUpdate()

If CmboPrimary.ListIndex <> -1 Then
    LblPrimaryAlternate.Caption = IIf(mvConfig.PrimaryAlternatePos = 1, Me.CmboPrimary.Column(Me.CmboPrimary.BoundColumn - 1, Me.CmboPrimary.ListIndex + 1), CmboPrimary.Column(mvConfig.PrimaryAlternatePos - 1, Me.CmboPrimary.ListIndex + 1))
    If mvConfig.SecondaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListSecondaryBy, Me.CmboSecondary, 2, mvConfig
        'If using Multi Select Clear the List
        If mvConfig.SecondaryListBoxMulti Then CmboSecondaryMulti.Clear
    End If
    If mvConfig.TertiaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
        If mvConfig.TertiaryListBoxMulti Then CmboTertiaryMulti.Clear
    End If
End If

If CmboPrimaryMulti.ListCount > 0 Then
    If mvConfig.SecondaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListSecondaryBy, Me.CmboSecondary, 2, mvConfig
        'If using Multi Select Clear the List
        If mvConfig.SecondaryListBoxMulti Then CmboSecondaryMulti.Clear
    End If
    If mvConfig.TertiaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
        If mvConfig.TertiaryListBoxMulti Then CmboTertiaryMulti.Clear
    End If
End If
End Sub
Private Sub CmboPrimary_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboPrimary.Dropdown
End Sub

Private Sub CmboReports_BeforeUpdate(Cancel As Integer)
'On Error Resume Next
Me.CmboReports.Dropdown
End Sub

Private Sub CmboReports_Change()
    'SA 05/21/2012 - CR2131 Added event
    RunEvent "Report Selected", Me.ScreenID, Me.FormID
End Sub

Private Sub CmboReports_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
    Me.CmboReports.Dropdown
End Sub

Private Sub CmboSecondary_AfterUpdate()
'On Error Resume Next
If CmboSecondary.ListIndex <> -1 Then
    LblSecondaryAlternate.Caption = vbNullString & IIf(mvConfig.SecondaryAlternatePos = 1, vbNullString & Me.CmboSecondary.Column(Me.CmboSecondary.BoundColumn - 1, Me.CmboSecondary.ListIndex + 1), vbNullString & CmboSecondary.Column(mvConfig.SecondaryAlternatePos - 1, Me.CmboSecondary.ListIndex + 1))
    If mvConfig.TertiaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
        If mvConfig.TertiaryListBoxMulti Then CmboTertiaryMulti.Clear
    End If
End If

If CmboSecondaryMulti.ListCount > 0 Then
    If mvConfig.TertiaryListBoxUse Then
        ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
        If mvConfig.TertiaryListBoxMulti Then CmboTertiaryMulti.Clear
    End If
End If
End Sub

Private Sub CmboSecondary_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboSecondary.Dropdown
End Sub

Private Sub CmboTertiary_AfterUpdate()
'    On Error Resume Next
    LblTertiaryAlternate.Caption = vbNullString & IIf(mvConfig.TertiaryAlternatePos = 1, vbNullString & Me.CmboTertiary.Column(Me.CmboTertiary.BoundColumn - 1, Me.CmboTertiary.ListIndex + 1), vbNullString & Me.CmboTertiary.Column(mvConfig.TertiaryAlternatePos - 1, Me.CmboTertiary.ListIndex + 1))
    ClsSCR.SetLabelFragment CmboTertiaryMulti, lblTertiaryMulti, 3
End Sub

Private Sub CmboTertiary_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboTertiary.Dropdown
End Sub

Private Sub CmboSortFieldList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboSortFieldList.Dropdown
End Sub

Private Sub CmboTotals_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Me.CmboTotals.Dropdown
End Sub

Private Sub CmdApplyTotals_Click()
On Error GoTo ErrorHappened
 
Dim pgCt As Integer
Dim Found As Boolean

If CmboTotals.ListIndex = -1 Then
    Me.Tabs.Pages(glTabTotalsCustom).visible = False
Else
    If Not MvGridMain Is Nothing Then
        If Not MvGridMain.RecordSet Is Nothing Then
            If Me.Tabs.Pages(glTabTotalsCustom).visible = False Then
               Me.Tabs.Pages(glTabTotalsCustom).visible = True
            End If
            If Me.FormFooter.visible = False Then
                Me.FormFooter.visible = True
                Form_Resize
            End If
            DoEvents
            Me.Tabs.Value = glTabTotalsCustom ' Make it the active Tab
        End If
    End If
End If

BuildTotalsCustom True

' find out if there are any tabs to show
Found = False
For pgCt = 0 To Me.Tabs.Pages.Count - 1
    If vbNullString & Tabs.Pages(pgCt).Controls(0).Tag <> vbNullString Then
        Found = True
        pgCt = Me.Tabs.Pages.Count
    End If
Next

If Not Found Then
    Me.FormFooter.visible = False
    Form_Resize
End If


ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Applying Custom Totals : CmdApplyTotals_Click"
    Resume ExitNow
    Resume
End Sub

Private Sub CmdBuildMulti1_Click()
PopupMultiSelect 1
End Sub
Private Sub CmdBuildMulti2_Click()

Dim Prompt As Boolean
Prompt = False

If Not mvConfig.SecondaryListBoxDependency And IsNull(Me.CmboSecondary) Then
    ClsSCR.PopulateCriteriaLists Me.CmboListSecondaryBy, Me.CmboSecondary, 2, mvConfig
End If

If mvConfig.SecondaryListBoxDependency Then
    If mvConfig.PrimaryListBoxMulti Then
        If Nz(ClsSCR.BuildMultiItemSQL(CmboPrimaryMulti.Object, mvConfig.PrimaryQualifier), vbNullString) = vbNullString Then
            Prompt = True
        End If
    ElseIf IsNull(Me.CmboPrimary) Then
            Prompt = True
    End If
End If

If Prompt Then
    MsgBox "Please select '" & Me.LblPrimary.Caption & "' value.", vbInformation + vbOKOnly, LblSecondary.Caption & " Select"
    GoTo Multi2Exit
End If

PopupMultiSelect 2

Multi2Exit:
    Exit Sub
End Sub
Private Sub CmdBuildMulti3_Click()
Dim Prompt As Boolean
Dim prompt1 As Boolean

Prompt = False
prompt1 = False

If Not mvConfig.TertiaryListBoxDependency And IsNull(Me.CmboTertiary) Then
    ClsSCR.PopulateCriteriaLists Me.CmboListTertiaryBy, Me.CmboTertiary, 3, mvConfig
End If

If mvConfig.TertiaryListBoxPrimaryDependency Then
    If mvConfig.PrimaryListBoxMulti Then
        If Nz(ClsSCR.BuildMultiItemSQL(CmboPrimaryMulti.Object, mvConfig.PrimaryQualifier), vbNullString) = vbNullString Then
            prompt1 = True
        End If
    ElseIf IsNull(Me.CmboPrimary) Then
            prompt1 = True
    End If
End If

If mvConfig.TertiaryListBoxDependency Then
    If mvConfig.SecondaryListBoxMulti Then
        If Nz(ClsSCR.BuildMultiItemSQL(CmboSecondaryMulti.Object, mvConfig.SecondaryQualifier), vbNullString) = vbNullString Then
            Prompt = True
        End If
    ElseIf IsNull(Me.CmboSecondary) Then
            Prompt = True
    End If
End If

If prompt1 Then
    MsgBox "Please select '" & Me.LblPrimary.Caption & "' value.", vbInformation + vbOKOnly, LblTertiary.Caption & " Select"
    GoTo Multi3Exit
End If

If Prompt Then
    MsgBox "Please select '" & Me.LblSecondary.Caption & "' value.", vbInformation + vbOKOnly, LblTertiary.Caption & " Select"
    GoTo Multi3Exit
End If

PopupMultiSelect 3

Multi3Exit:
    Exit Sub

End Sub

Private Sub CmdExecuteFunction_Click()
On Error GoTo CmdExecuteFunctionError

    If CmboFunction.ListIndex <> -1 Then 'Not Blank
        'SA 1/17/2012 - Added Telemetry
        Telemetry.RecordOpen "Function", CmboFunction, "Decipher Screens"
    
        Select Case UCase(CmboFunction)
            Case "CLEAR"
                ClearValues Me
            Case "ITEM GRAPH"
                If mvConfig.DateUse Then
                    LaunchItemGraph mvConfig, CmboFunction.Column(1), mvConfig.PrimaryRecordSource, BuildDateCriteria(Me)
                Else
                    LaunchItemGraph mvConfig, CmboFunction.Column(1), mvConfig.PrimaryRecordSource, vbNullString
                End If
            Case "VENDOR NOTES"
                LaunchVendorNotes mvConfig.FormID, CmboFunction.Column(1)
            Case "EXCEL"
                Dim locExcel As New CT_ClsExcel
                With locExcel
                     .FormID = mvConfig.FormID
                    .Run
                End With
                Set locExcel = Nothing
            Case "DISC GRAPH"
                If mvConfig.DateUse Then
                    LaunchDiscGraph mvConfig, CmboFunction.Column(1), mvConfig.PrimaryRecordSource, BuildDateCriteria(Me)
                Else
                    LaunchDiscGraph mvConfig, CmboFunction.Column(1), mvConfig.PrimaryRecordSource, vbNullString
                End If
            Case "EXIT"
                DoCmd.Close acForm, Me.Name, acSaveNo
            Case "CUSTOM DUP CRITERIA SELECTION"
                #If ccDT = 1 Then
                Set criteriaForm = New Form_DT_CustomCriteriaSelection
                With criteriaForm
                    .SetParent Me.Form
                    .visible = True
                    .Initialize
                End With
                #End If
            Case Else
                DoCmd.OpenForm CmboFunction.Column(2), acNormal, , , , , Me.FormID
        End Select
    End If
    
CmdExecuteFunctionExit:
On Error Resume Next
    
Exit Sub
CmdExecuteFunctionError:
    If subFunction.visible Then
        SubForm.visible = False
    End If
    MsgBox Err.Description & vbCrLf & vbCrLf & " Error Intializing Function Selection!", vbCritical, "BAD COMPUTER"
    Resume CmdExecuteFunctionExit
    Resume
End Sub

Private Sub CmdFilterEdit_Click()
'On Error Resume Next
    Set filterForm = New Form_SCR_ScreensFilters
    With filterForm
        .SetParent Me.Form
        .visible = True
        .Initialize
    End With
    
    'SA 1/17/2012 - Added Telemetry
    Telemetry.RecordOpen "Form", filterForm.Name, mvConfig.ScreenName, "Decipher Screens"
End Sub

Private Sub cmdFiltersClear_Click()
    Me.cmboFiltersSelected.RowSource = vbNullString
End Sub

Private Sub cmdFiltersSave_Click()
On Error GoTo Err_CmdSaveFilter_Click
    ClsSCR.SaveScreenFilter Me.CmboFilters
Exit_CmdSaveFilter_Click:
    Me.CmboFilters.Requery
    Exit Sub

Err_CmdSaveFilter_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveFilter_Click
End Sub

Private Sub cmdLayoutApply_Click()
On Error GoTo ErrorHappened
    Dim stName As String
    Dim lngId As Long
    Dim db As DAO.Database
    Dim rst As DAO.RecordSet
    Dim SQL As String
    Dim X As Long
    Dim Y As Long
    Dim TxtFld As Access.TextBox
    Dim Msg As String

    DoCmd.Hourglass True
    Application.Echo False
    
    'Make sure there is a layout selected
    If CmboLayouts.ListIndex = -1 Then
        If MsgBox("Your Current Layout Will Be Cleared!", vbInformation + vbOKCancel, "Clear Layouts") = vbOK Then
            Me.SubformCalcs.Form.ApplyClear
            If Not MvGridMain Is Nothing Then
                MvGridMain.CalcFieldsClear 'CLEAR THE CALCS
                Me.SubformCondFormats.Form.ApplyFormatClear 'Clear FORMAT  List
                MvGridMain.FormatsClear
                MvGridMain.LayoutClear
            End If
        End If
        GoTo ExitNow
    End If
    
    lngId = Me.CmboLayouts
    stName = Me.CmboLayouts.Column(1, CmboLayouts.ListIndex)
    Msg = "Applying Layout " & Chr(34) & stName & Chr(34) & vbCrLf
    
    Set db = CurrentDb
    
    'REMOVE ALL OF THE EXISTING CALCS
    If Not MvGridMain Is Nothing Then
        MvGridMain.CalcFieldsClear
    End If

    'REMOVE CALCS FROM APPLY LIST
    Me.SubformCalcs.Form.ApplyClear
    'GET THE CALCULATED FIELDS
    SQL = "SELECT LC.*, C.CalcName " & _
          "FROM SCR_ScreensLayoutsCalculations AS LC " & _
          "INNER JOIN SCR_ScreensCalculations AS C ON LC.CalcID = C.CalcID " & _
          "WHERE LC.LayoutID =" & CStr(lngId)
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    With rst
        If .EOF And .BOF Then
            Msg = Msg & "Calculations: 0" & vbCrLf
        Else
            Do Until .EOF
                Me.SubformCalcs.Form.ApplyAdd .Fields("CalcID"), .Fields("CalcName")
                'SubformCalcs.Form
                .MoveNext
            Loop
            .Close
            Msg = Msg & "Calculations: " & Me.SubformCalcs.Form.ActiveCalcCount & vbCrLf
            Me.SubformCalcs.Form.ApplyCalcs
        End If
    End With
    Set rst = Nothing

    'REMOVE ALL OF THE EXISTING FORMATS
    Me.SubformCondFormats.Form.ApplyFormatClear 'Clear The List
    Me.SubformCondFormats.Form.ApplyFormats 'Clear the data by applying
    
    'GET THE CONDITIONAL FORMATS
    SQL = "SELECT LF.*, F.FormatName " & _
          "FROM SCR_ScreensLayoutsFormats AS LF " & _
          "INNER JOIN SCR_ScreensCondFormats AS F on " & _
          "LF.CondFormatID = F.CondFormatID " & _
          "WHERE LF.LayoutID =" & CStr(lngId)
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    With rst
        If .EOF And .BOF Then
            Msg = Msg & "Formats: 0" & vbCrLf
        Else
            Do Until .EOF
                Me.SubformCondFormats.Form.ApplyFormatAdd .Fields("CondFormatID"), .Fields("FormatName")
                .MoveNext
            Loop
            .Close
            Msg = Msg & "Formats: " & Me.SubformCondFormats.Form.ActiveFormatCount & vbCrLf
            Me.SubformCondFormats.Form.ApplyFormats
        End If
    End With
    Set rst = Nothing

    'CLEAR THE FIELD LAYOUT
    MvGridMain.LayoutClear
    'GET THE FIELD LAYOUTS FROM THE DATABASE
    SQL = "SELECT FieldName, Ordinal, ColWidth, CalcFld " & _
          "FROM SCR_ScreensLayoutsFields AS LF " & _
          "WHERE LF.LayoutID =" & CStr(lngId) & " AND LF.Identifier='MainGrid' " & _
          "ORDER BY Ordinal"
    X = 0
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    With rst
        If .EOF And .BOF Then
            Msg = Msg & "Layouts: 0" & vbCrLf
        Else
            Do Until .EOF
                X = X + 1
                MvGridMain.LayoutField .Fields("FieldName"), .Fields("Ordinal"), .Fields("ColWidth"), .Fields("CalcFld")
                .MoveNext
            Loop
            Msg = Msg & "Layouts: " & .recordCount & vbCrLf
            .Close
        End If
    End With
    Set rst = Nothing
    
    DoEvents
    
    ' HC 9/25/2008 - load the tab layouts
    Set db = CurrentDb
    Dim Found As Boolean
    For X = 0 To mvConfig.TabsCT - 1
        'GET THE FIELD LAYOUTS FROM THE DATABASE
        Y = 0
        SQL = "SELECT Identifier, FieldName, Ordinal, ColWidth, CalcFld " & _
              "FROM SCR_ScreensLayoutsFields AS LF " & _
              "WHERE LF.LayoutID=" & CStr(lngId) & " AND LF.Identifier='" & mvConfig.Tabs(X).TabID & "' " & _
              "ORDER BY Ordinal"

        Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

        With rst
            If rst.recordCount > 0 Then
                Found = False
                For Y = 0 To Tabs.Pages.Count - 1
                    If Tabs.Pages(Y).Tag = rst.Fields("Identifier") Then
                        Found = True
                        Set holdForm = Nothing
                        Set holdForm = Tabs.Pages(Y).Controls(0).Form
                        
                        'SA 11/15/2012 - Only apply layouts for tabs with datasheet
                        If left(holdForm.Name, 22) = "CT_SubGenericDataSheet" Then
                            holdForm.LayoutClear
                        Else
                            Found = False
                        End If
                        Exit For
                    End If
                Next Y
                If Found Then
                    Do Until .EOF
                        Y = Y + 1
                        holdForm.LayoutField .Fields("FieldName"), .Fields("Ordinal"), .Fields("ColWidth"), .Fields("CalcFld")
                        .MoveNext
                    Loop
                End If
                .Close
            End If
        End With
        Set rst = Nothing
    Next X

    DoEvents
    
ExitNow:
On Error Resume Next
    Set db = Nothing
    Set rst = Nothing
    Set TxtFld = Nothing
    
    'CRAZY FORMAT CODE TO KEEP SCROLL BAR
    MvGridMain.Form.InsideWidth = Me.InsideWidth
    Form_Resize
    
    DoCmd.Hourglass False
    Application.Echo True
Exit Sub
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & Msg, vbCritical, "Error Saving Layout"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdLayoutDelete_Click()
Dim lngId As Long
Dim db As DAO.Database
Dim stName As String
Set db = CurrentDb

'Take current layout as default name
If CmboLayouts.ListIndex <> -1 Then
    stName = CmboLayouts.Column(1, CmboLayouts.ListIndex)
    lngId = CmboLayouts
End If
'Check if it exists
lngId = Nz(DLookup("LayoutID", "SCR_ScreensLayouts", "ScreenID = " & mvConfig.ScreenID & " and LayoutName = " & Chr(34) & stName & Chr(34)), -1)
If lngId <> -1 Then 'It exists - Ask then delete IT
    If MsgBox("Delete Layout '" & stName & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete") = vbYes Then
        db.Execute "Delete * From SCR_ScreensLayouts Where LayoutID = " & CStr(lngId), dbFailOnError
    End If
End If

Me.CmboLayouts.Requery

End Sub

Private Sub cmdLayoutSave_Click()
On Error GoTo ErrorHappened
    Dim stName As String
    Dim lngId As Long
    Dim db As DAO.Database
    Dim SQL As String
    Dim X As Long
    Dim TxtFld As Access.TextBox
    Dim Msg As String
    Dim Y As Long
    
    
    DoCmd.Hourglass True
    'Take current layout as default name
    If CmboLayouts.ListIndex <> -1 Then
        stName = CmboLayouts.Column(1, CmboLayouts.ListIndex)
        lngId = CmboLayouts
    End If
    
GetName:
    'Confirm the name
    stName = InputBox("Please enter the name of the new existing layout.", "Layout Name", stName)
    If vbNullString & stName = vbNullString Then
        GoTo ExitNow
    End If
    
    Set db = CurrentDb
    'Check if it exists
    lngId = Nz(DLookup("LayoutID", "SCR_ScreensLayouts", "ScreenID = " & mvConfig.ScreenID & " and LayoutName = " & Chr(34) & stName & Chr(34)), -1)
    If lngId <> -1 Then 'It exists - Ask then delete IT
        If MsgBox("The Layout '" & stName & "' already exists." & vbCrLf & vbCrLf & "Would you like to replace it?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Replace") = vbYes Then
            db.Execute "DELETE FROM SCR_ScreensLayouts WHERE LayoutID=" & CStr(lngId), dbFailOnError
        Else
            GoTo GetName 'Try Getting a new name
        End If
    End If
    
    'Create The Layout Record
    SQL = "INSERT INTO SCR_ScreensLayouts(ScreenID, LayoutName,Computer,UserName)VALUES(" & _
        mvConfig.ScreenID & ", " & _
        Chr(34) & stName & Chr(34) & ", " & _
        Chr(34) & Identity.Computer & Chr(34) & ", " & _
        Chr(34) & Identity.UserName & Chr(34) & ")"
    db.Execute SQL, dbFailOnError
    lngId = Nz(DLookup("LayoutID", "SCR_ScreensLayouts", "ScreenID = " & mvConfig.ScreenID & " and LayoutName = " & Chr(34) & stName & Chr(34)), -1)
    Msg = "Layout " & Chr(34) & stName & Chr(34) & " Saved!" & vbCrLf
    
    'SAVE THE FORMATS
    With SubformCondFormats.Form
        If .ActiveFormatCount > 0 Then
            'Create The FormatRecords
            SQL = "INSERT INTO SCR_ScreensLayoutsFormats(LayoutID,CondFormatID) " & _
                "SELECT " & CStr(lngId) & ", CondFormatID " & _
                "FROM SCR_ScreensCondFormats " & _
                "WHERE CondFormatID IN (" & .SQLList & ")"
            db.Execute SQL, dbFailOnError + dbSeeChanges
            Msg = Msg & "Formats: " & .ActiveFormatCount & vbCrLf
        Else
            Msg = Msg & "Formats: 0" & vbCrLf
        End If
    End With
    
    'SAVE THE CALCULATIONS
    With Me.SubformCalcs.Form
        If .ActiveCalcCount > 0 Then
            'Create The FormatRecords
            SQL = "INSERT INTO SCR_ScreensLayoutsCalculations(LayoutID,CalcID) " & _
                "SELECT " & CStr(lngId) & ", CalcID " & _
                "FROM SCR_ScreensCalculations " & _
                "WHERE CalcID in (" & .SQLList & ")"
            db.Execute SQL, dbFailOnError
            Msg = Msg & "Calculations: " & .ActiveCalcCount & vbCrLf
        Else
            Msg = Msg & "Calculations: 0" & vbCrLf
        End If
    End With
    
    ' HC 9/25/2008 changed mvgridmain to me.subform.form; added new identifier to screenlayoutfields
    With Me.SubForm.Form    'MvGridMain
        For X = 1 To .FldCT
            Set TxtFld = .Controls("Field" & CStr(X))
            SQL = "INSERT INTO SCR_ScreensLayoutsFields(LayoutID,Identifier,FieldName,CalcFld,ColWidth,Ordinal)VALUES(" & _
                CStr(lngId) & ", 'MainGrid' , "
            If vbNullString & TxtFld.Tag <> vbNullString Then 'Calculated field
                SQL = SQL & Chr(34) & TxtFld.Tag & Chr(34) & ", " & "True, "
            Else
                '021406 David.Brady added "EscapeQuotes" to accomodate control sources with quotes in them.
                SQL = SQL & Chr(34) & EscapeQuotes(TxtFld.ControlSource) & Chr(34) & ", " & "False, "
            End If
            SQL = SQL & IIf(TxtFld.ColumnHidden = True, 0, TxtFld.ColumnWidth) & ", " & _
                TxtFld.ColumnOrder & vbNullString & ")"
            db.Execute SQL, dbFailOnError
        Next X
        Msg = Msg & "Column Layouts: " & .FldCT & vbCrLf
    End With
    
    ' HC 9/24/2008 - build an xml string of the tab layout information
    
    For X = 0 To mvConfig.TabsCT - 1
        With Tabs.Pages(X + 1).Controls(0).Form
            'SA 11/15/2012 - Only save layouts for datasheets
            If left(.Name, 22) = "CT_SubGenericDataSheet" Then
                For Y = 1 To .FldCT
                    SQL = "INSERT INTO SCR_ScreensLayoutsFields(LayoutID,Identifier,FieldName,CalcFld,ColWidth,Ordinal)VALUES(" & _
                        CStr(lngId) & ", " & Chr(34) & mvConfig.Tabs(X).TabID & Chr(34) & ", "
                    Set TxtFld = .Controls("Field" & CStr(Y))
                    If vbNullString & TxtFld.Tag <> vbNullString Then 'Calculated field
                        SQL = SQL & Chr(34) & TxtFld.Tag & Chr(34) & ", " & "True, "
                    Else
                        '021406 David.Brady added "EscapeQuotes" to accomodate control sources with quotes in them.
                        SQL = SQL & Chr(34) & EscapeQuotes(TxtFld.ControlSource) & Chr(34) & ", " & "False, "
                    End If
                    SQL = SQL & IIf(TxtFld.ColumnHidden = True, 0, TxtFld.ColumnWidth) & ", " & _
                        TxtFld.ColumnOrder & vbNullString & ") "
                    db.Execute SQL, dbFailOnError
                Next Y
            End If
        End With ' with
    Next X   ' next x
           
    Msg = Msg & "Tab Layouts Layouts: " & mvConfig.TabsCT & vbCrLf
       
MsgBox Msg, vbInformation, "Layout Saved"
ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Me.CmboLayouts.Requery
    Set db = Nothing
    Set TxtFld = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Saving Layout"
    Resume ExitNow
    Resume
End Sub

Private Sub CmdPowerBar_Click()
On Error GoTo ErrorHappened
Dim SQL As String
Dim frm As New Form_CT_PopupSelect

SQL = "SELECT PwrBarID, ListName, Function "
SQL = SQL & "From SCR_ScreensPowerBars "
SQL = SQL & "WHERE ScreenID = " & mvConfig.ScreenID
SQL = SQL & " ORDER BY ListName"

With frm
    With .Lst
        .RowSource = SQL
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "0;3;0"
        .Requery
        '.MultiSelect = False
    End With
    .Title = "PowerBar Chooser"
    .ListTitle = "Select PowerBar:"
    .StartupWidth = -1   'AUTO SIZE THE FORM TO LIST WIDTH
    .visible = True
    
    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        If .Lst.ListIndex <> -1 Then
            'SA 03/22/2012 - Added for Work Files support
            Me.SubformPowerBar.Tag = .Lst.Column(0, .Lst.ListIndex)
            
            Me.SubformPowerBar.SourceObject = vbNullString & .Lst.Column(2, .Lst.ListIndex)
            Me.TxtCur.ControlSource = "= " & Chr(34) & "Current PowerBar --> " & Chr(34) & " & SubformPowerBar.Form.Caption"
            
            'SA 1/17/2012 - Added Telemetry
            Telemetry.RecordOpen "Powerbar", Me.SubformPowerBar.SourceObject, "Decipher Screens"
        Else
            Me.SubformPowerBar.SourceObject = vbNullString
        End If
    End If
End With

ExitNow:
    On Error Resume Next
    DoCmd.Hourglass False
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, "CmdPowerBar_Click --> Select PowerBar"
    Resume ExitNow
    Resume
End Sub
Private Sub CmdRefreshSmall_Click()
    RefreshForm
End Sub
Public Sub cmdRefresh_Click()
    RefreshForm
End Sub

'/* CR# 2564 fix - Clear phantom filters during refresh.
'                  If there is a grid filter,clear it and set
'                  the grid to a 'No Filter' state on refresh. */

Private Sub RefreshForm()
On Error GoTo ErrorHappened
    DoCmd.Hourglass True
    genUtils.SuspendLayout Me
    
    'SA 05/21/2012 - CR2132 Added event
    RunEvent "Screen Refresh Start", Me.ScreenID, Me.FormID
    
    'SA 1/19/2012 - CR 1967 Reset TabLoaded variable when the form is refreshed
    ClsSCR.ResetTabs
    ReDim TabLoaded(Tabs.Pages.Count)
    
    ' CR# 2564 fix
    If MvGridMain.filter <> vbNullString Then
        With MvGridMain
            .FilterOn = False
            .filter = vbNullString
        End With
    End If
    
    selectionFilter = ClsSCR.BuildDetail(mvSql, mvConfig)
    
    'SA 1/17/2012 - Now calls TabsLoad instead of Tabs_Change
    TabsLoad        'Tabs_Change
    BuildTotalsCustom False
    
    RunEvent "Screen Refresh", Me.ScreenID, Me.FormID
    
    'SA 1/17/2012 - Added telemetry
    'SA 10/1/2012 - CR3128 XML encode screen name to prevent telemetry errors
    Telemetry.RecordAction "Screen Refresh", "<SN>" & genUtils.XMLEncode(Me.ScreenName) & "</SN><P>" & genUtils.XMLEncode(mvSql.SqlAll) & "</P>", "Decipher Screens"
    
ExitNow:
    On Error Resume Next
    genUtils.ResumeLayout Me
    DoCmd.Hourglass False
    
    Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub cmdRunReport_Click()
On Error GoTo CmdRunReportError
    CmdRunReport.Enabled = False
    If Me!CmboReports.ListIndex <> -1 Then
        RunReport Me!CmboReports.Column(0), Me!CmboReports.Column(1)
        
        'SA 1/17/2012 - Added Telemetry
        Telemetry.RecordOpen "Report", Me!CmboReports.Column(0), "Decipher Screens"
    End If
CmdRunReportExit:
    On Error Resume Next
    CmdRunReport.Enabled = True
    Exit Sub
CmdRunReportError:
    MsgBox Err.Description & "Error Intializing Report Selection!", vbCritical, "BAD COMPUTER"
    Resume CmdRunReportExit
End Sub

Private Sub CmdSortAdd_Click()
On Error GoTo CmdSortAddError
    'SA 03/22/2012 - CR1782 Changed SortList from ActiveX to Access Listbox
    Dim i As Integer
    Dim SortField As String
    Dim SortDir As String
    Dim FieldExists As Boolean

    If Me.CmboSortFieldList.ListIndex > -1 Then
        SortField = Me.CmboSortFieldList
        
        'Check to see if field was selected already
        FieldExists = False
        For i = 0 To Me.SortList.ListCount - 1
            If SortField = Me.SortList.Column(1, i) Then
                FieldExists = True
                Exit For
            End If
        Next
        
        'Sort direction
        If Me.ToggleSort Then
            SortDir = "D"
        Else
            SortDir = "A"
        End If
        
        'Add to sort list
        If Not FieldExists Then
            Me.SortList.AddItem SortDir & ";" & SortField
            Me.SortList.Selected(Me.SortList.ListCount - 1) = True
            ClsSCR.UpdateSortTip
        Else
            MsgBox "The field '" & SortField & "' is already in the sort list." & vbCrLf & vbCrLf & _
                    "If you want to sort in a different direction, remove the field from the list first.", vbInformation, "Sorting"
        End If
        Me.CmboSortFieldList = vbNullString
    Else
        MsgBox "Please select a field you want to sort on.", vbInformation, "Sorting"
    End If

CmdSortAddExit:
    On Error Resume Next
    Exit Sub
CmdSortAddError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error adding field to sort list!"
    Resume CmdSortAddExit
    Resume
End Sub

Private Sub CmdSortDelete_Click()
On Error GoTo ErrorHandler
    'SA 03/22/2012 - CR1782 Changed SortList from ActiveX to Access Listbox
    If Me.SortList.ListCount = 1 Then
        Me.SortList.RemoveItem 0
    Else
        If Me.SortList.ListIndex < 0 Then
            MsgBox "Please select a sort item to remove.", vbInformation
        Else
            Me.SortList.RemoveItem (SortList.ListIndex)
        End If
    End If
ExitDelete:
On Error Resume Next
    ClsSCR.UpdateSortTip
Exit Sub
ErrorHandler:
    Resume ExitDelete
    Resume
End Sub

Private Sub CmdSortDeleteAll_Click()
    'SA 03/22/2012 - CR1782 Changed SortList from ActiveX to Access Listbox
    Me.SortList.RowSource = vbNullString
    ClsSCR.UpdateSortTip
End Sub

Private Sub CmdSortMoveDown_Click()
    SortListMoveItem 1
End Sub

Private Sub CmdSortMoveUp_Click()
    SortListMoveItem -1
End Sub

Private Sub CmdSortOpen_Click()
On Error GoTo OpenSortsError
    Dim SortForm As Form
    Dim MySortList As listBox
    
    DoCmd.OpenForm CCASorts
    
    Set SortForm = Forms(CCASorts)
    Set MySortList = Me.SortList
    
    With SortForm
        .ScreenID = mvConfig.ScreenID
        .BoundSortList = MySortList
        .BoundScreenName = mvConfig.ScreenName
        .InitData
    End With
OpenSortsExit:
On Error Resume Next
    Set SortForm = Nothing
    Set MySortList = Nothing
    Exit Sub
OpenSortsError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error opening sorts form!", vbCritical, mvConfig.ScreenName
    Resume OpenSortsExit
End Sub

Private Sub CmdSortSave_Click()
On Error GoTo SaveSortError
    'SA 03/22/2012 - CR1782 Changed SortList from ActiveX to Access Listbox
    Dim SortName As String
    Dim X As Integer
    Dim SQL As String

    'Get sort name
    SortName = Nz(InputBox("Enter a name for this sort." & String(2, vbCrLf) & _
            "Default - Load automatically with screen.", "Save Sort", vbNullString), vbNullString)

    If LenB(SortName) > 0 Then
        'Check if name exists
        If DLookup("SortName", "SCR_ScreensSorts", "ScreenID=" & mvConfig.ScreenID & " and SortName=" & Chr(34) & SortName & Chr(34)) = SortName Then
            If MsgBox("The sort '" & SortName & "' already exists!" & String(2, vbCrLf) & "Would you like to replace it?", vbQuestion + vbYesNo + vbDefaultButton2, "Replace Sort?") = vbYes Then
                CurrentDb.Execute "DELETE FROM SCR_ScreensSorts WHERE ScreenID=" & mvConfig.ScreenID & " and SortName=" & Chr(34) & SortName & Chr(34)
            End If
        End If
        
        'Save
        For X = 0 To Me.SortList.ListCount - 1
            SQL = "INSERT INTO SCR_ScreensSorts(ScreenID, SortName, SortIndex, FieldName, SortOrder)" & _
                    "VALUES(" & mvConfig.ScreenID & ", " & Chr(34) & SortName & Chr(34) & ", " & _
                    X & ", " & Chr(34) & Me.SortList.Column(1, X) & Chr(34) & ", " & Chr(34) & Me.SortList.Column(0, X) & Chr(34) & ")"
            CurrentDb.Execute SQL
        Next X
    Else
        MsgBox "The current sort items were not saved.", vbInformation, "Sorting"
    End If

SaveSortExit:
On Error Resume Next
    Exit Sub
SaveSortError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Saving Current Sort!", vbCritical, "Save Sort"
    Resume SaveSortExit
    Resume
End Sub

Private Sub CmdRowCount_Click()
'SA 1/18/2012 - Button was added to test a query to see how many rows will be returned
'Compile SQL and run a test query to see how many rows will be returned
On Error GoTo ErrorHappened
    DoCmd.Hourglass True
    genUtils.SuspendLayout

    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    
    Dim strSQL As String
    Dim strWherePrimary As String
    Dim StrDateRange As String
    Dim StrFilters As String
    
    strWherePrimary = ClsSCR.BuildWherePrimary(mvConfig)
    
    If Nz(strWherePrimary, vbNullString) <> vbNullString Then
        strSQL = "SELECT COUNT(1) AS RwCnt FROM " & mvConfig.PrimaryRecordSource & " WHERE " & ClsSCR.BuildWherePrimary(mvConfig) & " "
        If mvConfig.SecondaryListBoxUse And (Me.CmboSecondary.ListIndex <> -1 Or Me.CmboSecondaryMulti.ListCount <> 0) Then
            strSQL = strSQL & " AND " & ClsSCR.BuildWhereSecondary(mvConfig) & " "
        End If
        If mvConfig.TertiaryListBoxUse And (Me.CmboTertiary.ListIndex <> -1 Or Me.CmboTertiaryMulti.ListCount <> 0) Then
            strSQL = strSQL & " AND " & ClsSCR.BuildWhereTertiary(mvConfig) & " "
        End If
        If mvConfig.DateUse Then
            StrDateRange = BuildDateCriteria(Me)
            If StrDateRange <> vbNullString Then
                strSQL = strSQL & " AND " & StrDateRange
            End If
        End If
        StrFilters = ClsSCR.ApplyListedFilters
        If LenB(StrFilters) > 0 Then
            strSQL = strSQL & " AND " & StrFilters
        End If

        Set db = CurrentDb
        Set rs = db.OpenRecordSet(strSQL, dbOpenSnapshot, dbForwardOnly)

        If rs.recordCount = 1 Then
            genUtils.ResumeLayout
            DoCmd.Hourglass False
            
            If rs!RwCnt > 0 Then
                If MsgBox("These settings will return " & Format(rs!RwCnt, "#,###") & _
                    " rows of data." & vbCrLf & vbCrLf & _
                    "Would you like to run the query now?", vbYesNo + vbInformation, "Test results") = vbYes Then
                        
                    RefreshForm
                End If
            Else
                MsgBox "These settings will return 0 rows of data.", vbExclamation, "Test results"
            End If
        End If
        
        'SA 03/22/2012 - Added telemetry event
        'SA 10/1/2012 - CR3128 XML encode screen name to prevent telemetry errors
        Telemetry.RecordAction "Query row count", "<SN>" & genUtils.XMLEncode(Me.ScreenName) & "</SN><P>" & genUtils.XMLEncode(strSQL) & "</P>", "Decipher Screens"
    Else
        genUtils.ResumeLayout
        DoCmd.Hourglass False
        
        MsgBox "Please select primary criteria.", vbExclamation
    End If

ExitNow:
    On Error Resume Next
    genUtils.ResumeLayout
    DoCmd.Hourglass False
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    Exit Sub
ErrorHappened:
    MsgBox "There is a problem with this query. Please check you settings and try again.", vbExclamation, "Error"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdToggle_Click()
    ClsSCR.SetTogValue = Not ClsSCR.GetTogValue
    ClsSCR.PopulateScreenFilters Me.CmboFilters, mvConfig.ScreenID
End Sub

Private Sub CmdFiltersApply_Click()
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2667 Added apply filter functionality
    Dim ExistingFilter As String
    Dim newFilter As String
    Dim itemFilter As String
    Dim i As Integer
    
    ExistingFilter = MvGridMain.filter

    'Check for right click filters
    If InStr(1, ExistingFilter, "[" & MvGridMain.Name & "]", vbTextCompare) = 0 Then
        newFilter = vbNullString    'Rebuild filter
    Else
        newFilter = ExistingFilter  'Append existing
    End If

    'Load filters from list and apply if not in exiting filter
    For i = 0 To cmboFiltersSelected.ListCount - 1
        itemFilter = DLookup("FilterSQL", "SCR_ScreensFilters", "FilterID=" & Me.cmboFiltersSelected.Column(0, i))
        
        If InStr(1, newFilter, itemFilter, vbTextCompare) = 0 Then
            If LenB(newFilter) > 0 Then
                newFilter = newFilter & " AND (" & itemFilter & ")"
            Else
                newFilter = "(" & itemFilter & ")"
            End If
        End If
    Next

    'Apply filter if changed
    If newFilter <> ExistingFilter Then
        MvGridMain.filter = newFilter
        MvGridMain.FilterOn = True
        
        TabsLoad
        BuildTotalsCustom False
    End If
ExitNow:

    Exit Sub
ErrorHappened:
    MsgBox "Error applying filter", vbCritical, "Error"
    Resume ExitNow
    Resume
End Sub


Private Sub CmdTotalsDelete_Click()
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim SQL As String

    If Me.CmboTotals.ListIndex = -1 Then
        GoTo ExitNow
    Else
        If MsgBox("Are you sure you want to delete the following Custom Totals:" & vbCrLf & vbCrLf & CmboTotals.Column(1, CmboTotals.ListIndex), vbQuestion + vbYesNo + vbDefaultButton2, "Delete Totals") = vbYes Then
            SQL = "Delete SCR_ScreensTotals.TotalID From SCR_ScreensTotals Where TotalID = " & Me.CmboTotals
            Set db = CurrentDb
            CurrentDb.Execute SQL, dbFailOnError
        End If
    End If

    CmboTotals.Requery

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Loading Total Config Form"
    Resume ExitNow
    Resume
End Sub

Private Sub CmdTotalsEdit_Click()
On Error GoTo ErrorHappened
Dim frmTotals As New Form_SCR_CfgCustomTotals

Dim TotalID As Long
    If Me.CmboTotals.ListIndex = -1 Then
        If MsgBox("You are about to create a new custom totals definition" & vbCrLf & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Create New Totals") = vbNo Then
            GoTo ExitNow
        End If
        TotalID = 0
    Else
        TotalID = Me.CmboTotals
    End If

With frmTotals
    .CurrentScreenID = Me.Config.ScreenID
    .CurrentFormID = Me.Config.FormID
    .CurrentTotalID = TotalID
    .Modal = True
    If TotalID = 0 Then
        .show True
    Else
        .show
    End If
    Do While .Results = vbApplicationModal
        DoEvents
    Loop
End With
ExitNow:
    On Error Resume Next
    Me.CmboTotals.Requery
    Set frmTotals = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error Loading Total Config Form"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdFiltersRemove_Click()
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2667 Changed how filters are removed from the list
    If Me.cmboFiltersSelected.ListCount = 1 Then
        Me.cmboFiltersSelected.RemoveItem 0
    Else
        If Me.cmboFiltersSelected.ItemsSelected.Count = 0 Then
            MsgBox "Please select a filter to remove.", vbInformation
        Else
            Me.cmboFiltersSelected.RemoveItem Me.cmboFiltersSelected.ListIndex
        End If
    End If
    
    Me.CmboFilters = vbNullString
    
ExitNow:

    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error removing filter"
    Resume ExitNow
    Resume
End Sub
Private Sub CmboFilters_Click()
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2667 Add filter to list when selected
    Dim FilterName As String
    Dim FilterID As String
    Dim FilterFound As Boolean
    Dim i As Integer
    
    FilterFound = False
    
    If LenB(Nz(Me.CmboFilters.Value, vbNullString)) > 0 Then
        'Get filter name
        FilterID = Me.CmboFilters.Value
        For i = 0 To CmboFilters.ListCount - 1
            If CmboFilters.Column(0, i) = FilterID Then
                FilterName = Me.CmboFilters.Column(1, i)
                Exit For
            End If
        Next
        
        'Check to see if filter is already in the list
        For i = 0 To Me.cmboFiltersSelected.ListCount - 1
            If FilterID = Me.cmboFiltersSelected.Column(0, i) Then
                FilterFound = True
                Exit For
            End If
        Next
        
        'Add to list if not already there
        If Not FilterFound Then
            If LenB(cmboFiltersSelected.RowSource) = 0 Then
                cmboFiltersSelected.RowSource = FilterID & ";'" & Replace(FilterName, "'", "''") & "'"
            Else
                cmboFiltersSelected.RowSource = cmboFiltersSelected.RowSource & ";" & FilterID & ";'" & Replace(FilterName, "'", "''") & "'"
            End If
        End If
    End If
ExitNow:
    CmboFilters = vbNullString
    Exit Sub
ErrorHappened:
    MsgBox "Error adding filter to list", vbCritical, "Error"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdFiltersAdd_Click()
    Dim strName As String
    Dim FilterID As String
    Dim i As Integer
    Dim bFound As Boolean
    
    If Nz(Me.CmboFilters.Value, vbNullString) = vbNullString Then
        MsgBox "Please select a filter to add."
        Me.CmboFilters.SetFocus
        Me.CmboFilters.Dropdown
        Exit Sub
    Else
        FilterID = Me.CmboFilters.Value
        For i = 0 To CmboFilters.ListCount - 1
            If CmboFilters.Column(0, i) = FilterID Then
                strName = Me.CmboFilters.Column(1, i)
                Exit For
            End If
        Next
    End If
    
    bFound = False
    
    For i = 0 To Me.cmboFiltersSelected.ListCount - 1
        If FilterID = Me.cmboFiltersSelected.Column(0, i) Then
            bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then
        If cmboFiltersSelected.RowSource = vbNullString Then
            cmboFiltersSelected.RowSource = FilterID & ";'" & Replace(strName, "'", "''") & "'"
        Else
            cmboFiltersSelected.RowSource = cmboFiltersSelected.RowSource & ";" & FilterID & ";'" & Replace(strName, "'", "''") & "'"
        End If
    End If
    CmboFilters = vbNullString
End Sub
Private Sub filterForm_Cancelled()
    Me.CmdRefresh.SetFocus
    Me.subFunction.visible = False
    Set filterForm = Nothing
    'SA 1/23/2012 - CR1967 Remove call to RefreshForm to stop screen refresh on filter cancel
    'RefreshForm
    CmboFilters.Requery
End Sub
Private Sub criteriaForm_Cancelled()
    Me.CmdRefresh.SetFocus
    Me.subFunction.visible = False
    
    #If ccDT = 1 Then
    Set criteriaForm = Nothing
    #End If
    
    CmboFilters.Requery
    RefreshForm
End Sub
Private Sub filterForm_FinishedFilter(ByVal FilterID As Long)
    Dim i As Integer

    Me.CmdRefresh.SetFocus
    Me.subFunction.visible = False
    Set filterForm = Nothing
    
    
    ClsSCR.PopulateScreenFilters Me.CmboFilters, mvConfig.ScreenID
    For i = 0 To Me.CmboFilters.ListCount - 1
        If Me.CmboFilters.Column(0, i) = FilterID Then
            Me.CmboFilters.Value = FilterID
            cmdFiltersAdd_Click
            Exit For
        End If
    Next
    
    CmboFilters.Requery
    'SA 03/22/2012 - CR2667 Apply filter or refresh form depending on filter state
    If LenB(MvGridMain.filter) > 0 Then
        CmdFiltersApply_Click
    Else
        RefreshForm
    End If
            
    Me.SubformCondFormats.Form.ApplyFormats
    
End Sub
Private Sub criteriaForm_Finished()
    Dim i As Integer
    Dim SQL As String
    Dim db As Database
    Dim rs As RecordSet

    Me.CmdRefresh.SetFocus
    Me.subFunction.visible = False

    #If ccDT = 1 Then
    Set criteriaForm = Nothing
    #End If
    
    For i = 0 To Me.CmboFilters.ListCount - 1
        ' HC 6/2010 - Changed the filter to look for either one of these as the filter to apply for the Custom Dup Tool Filters.
        If Me.CmboFilters.Column(1, i) = "1 - Custom Criteria Results" Or Me.CmboFilters.Column(1, i) = "1 - All Custom Criteria Results" Then
            Me.CmboFilters.Value = Me.CmboFilters.Column(0, i)
            cmdFiltersAdd_Click
            Exit For
        End If
    Next
    
    RefreshForm
    
    On Error GoTo eTrap
    Me.SubformCondFormats!LstApply.RowSource = vbNullString
    SQL = " SELECT SCR_ScreensCondFormats.CondFormatID, " & _
        " SCR_ScreensCondFormats.FormatName  " & _
        " FROM SCR_ScreensCondFormats " & _
        " WHERE ScreenId = " & mvConfig.ScreenID & "  AND SCR_ScreensCondFormats.Expression1 = 'CustomID Mod 2 = 0'"
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbReadOnly)
    If Not rs.EOF Then
        Me.SubformCondFormats!LstApply.RowSource = rs(0) & ";" & rs(1)
    End If
        
eTrap:
    rs.Close
    db.Close
    Me.SubformCondFormats.Form.ApplyFormats
    
End Sub
Private Sub filterForm_LoadError()
    Me.subFunction.visible = False
    Set filterForm = Nothing
End Sub

Private Sub criteriaForm_LoadError()
    Me.subFunction.visible = False
    
    #If ccDT = 1 Then
    Set criteriaForm = Nothing
    #End If
End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    RaiseEvent isVisible(True)
    
    'SA 05/21/2012 - CR2132 Added event
    RunEvent "Screen Activate", Me.ScreenID, Me.FormID
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
Me.SubForm.Form.Form_ApplyFilter Cancel, ApplyType
End Sub

Public Sub InitData()
On Error GoTo InitDataError
    DoCmd.Hourglass True
    genUtils.SuspendLayout Me
    
    ClsSCR.SetMyForm = Me
    
    Me.Caption = mvConfig.ScreenName
    Me!SubForm.Form.InitData mvConfig.PrimaryRecordSource, mvConfig.PrimaryRecordSourceType
    SetTabs Me, mvConfig

    ' Test if there are any tabs specified
    If mvConfig.TabsCT = 0 Then
        Me.FormFooter.visible = False
    End If
    
    'Are There any Powerbars???
    Me.PgPowerBars.visible = mvConfig.PowerBars
    
    'User tabs
    ClsSCR.LoadUserTab PgTabsHeadUser1, mvConfig.TabsHeadUser1
    ClsSCR.LoadUserTab PgTabsHeadUser2, mvConfig.TabsHeadUser2
    ClsSCR.LoadUserTab PgTabsHeadUser3, mvConfig.TabsHeadUser3
    
    'Set Up Fields
    PopulateLists
    If mvConfig.DateUse Then
        Me.StartDte.Value = mvConfig.StartDate
        Me.EndDte.Value = mvConfig.EndDate
    End If
    
    'Sets The Default List By
    SetListBy
    Form_Resize
    DoCmd.Hourglass False

    Dim FrmTmp As Form_SCR_CondFormats
    Set FrmTmp = SubformCondFormats.Form
    FrmTmp.InitData mvConfig.ScreenID, mvConfig.FormID
    Set FrmTmp = Nothing
    
    Me.SubformCalcs.Form.InitData mvConfig.ScreenID, mvConfig.FormID
    genUtils.ResumeLayout Me
    
InitDataExit:
On Error Resume Next
    Exit Sub
InitDataError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error in Sub InitData!", vbCritical, mvConfig.ScreenName
    Resume InitDataExit
End Sub

Private Sub CmdFilterSave_Click()
On Error GoTo Err_CmdSaveFilter_Click
    SaveFilter Me.CmboFilters
Exit_CmdSaveFilter_Click:
    Exit Sub

Err_CmdSaveFilter_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveFilter_Click
    
End Sub

Private Sub Form_Close()
SubformPowerBar.SourceObject = vbNullString
End Sub

Private Sub Form_Deactivate()
'On Error Resume Next
    screen.MousePointer = 0
    RaiseEvent isVisible(False)
    
    'SA 05/21/2012 - CR2132 Added event
    RunEvent "Screen Deactivate", Me.ScreenID, Me.FormID
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Beep
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim bHotKeySequence As Boolean
    
    bHotKeySequence = False
    
    If Shift = 3 Then bHotKeySequence = True
    
    If bHotKeySequence Then
        If KeyCode = Asc("S") Then
            ScreenRecordSourceOnly
        End If
        If KeyCode = Asc("P") Then
            ScreenCollapse
        End If
        If KeyCode = Asc("T") Then
            ScreenRestore
        End If
    End If
End Sub

Private Sub Form_Load()
    ' do not add suspend and resume layout to the load, causes a white screen on a Terminal Server
    
    DoCmd.Maximize
    ClsSCR.SetTogValue = True
    HasRowChangeEvent = True
    Set stateForm = Me.Child164.Form
    genUtils.ToggleAccessMenus (False)
    DoCmd.ShowToolbar "Filter/Sort", acToolbarYes
    
    SetRestoreCollapseMenu
    Me.CmdRefresh.SetFocus
    
    ' HC - do not move this set
    Set MvGridMain = Me.SubForm.Form
    
    CmdRefresh.SetFocus
    
    'SA 1/19/2012 - CR1967 Set array count to number of tabs
    ReDim TabLoaded(Tabs.Pages.Count)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    On Error GoTo Form_ResizeError:
    Dim GridHeight As Single
    Dim GridWidth As Single
    Dim GridPadding As Single
    Dim TabHeight As Single
    Dim i As Integer

    If Me.ScreenID <> 0 Then
    'suppress the suspend layout when creating a new screen for faster loading
        genUtils.SuspendLayout
    End If

    ' changed by SC Since the form footer property cangrow = true the following will prevent the footer from
    ' growing larger than the main grid which will cause a link break down between the main grid and the tabs
    ' HC 4/24/2008 - only change size if form footer is visible
    If FormFooter.visible And (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) > 0 And _
         FormFooter.Height >= (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) Then
            FormFooter.Height = Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))
    End If

    Me.lblBanner.top = 0
    Me.SubForm.top = lblBanner.Height
    GridPadding = Me.SubForm.left * 2
    If Me.FormFooter.visible Then
        GridHeight = Me.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height + lblBanner.Height)
    Else
        GridHeight = Me.InsideHeight - (Me.FormHeader.Height + lblBanner.Height)
    End If

    If GridHeight > 0 And (Me.WindowHeight > GridHeight) Then
        Me.Detail.Height = GridHeight + lblBanner.Height
        Me.SubForm.Height = GridHeight
    End If

    GridWidth = Me.InsideWidth
    If GridWidth > 0 Then
        Me.SubForm.Width = GridWidth
        lblBanner.Width = GridWidth
        Me.TabsHead.Width = GridWidth
        Me.Splitter.top = 0
        Me.Splitter.Width = GridWidth
        Me.Tabs.Width = GridWidth
    End If

    If Me.Tabs.Value > -1 Then
        txtFocus.SetFocus
        Me.Tabs.visible = True
        Me.Splitter.visible = True
        With Me.Tabs
            .left = 0
            .visible = False
            TabHeight = Me.FormFooter.Height - Me.Splitter.Height
            For i = 0 To mvConfig.TabsCT + glTabsUsed - 1
                With Tabs.Pages(i)
                    .left = Me.Tabs.left
                    If GridWidth > 0 Then
                        .Width = GridWidth - (GridPadding * 3)
                    End If
                    With .Controls(0)
                        .top = Tabs.Pages(i).top
                        If GridWidth > 0 Then
                            .Width = GridWidth - (GridPadding * 4)
                        End If
                        .Height = Tabs.Pages(i).Height - GridPadding
                        .left = Tabs.Pages(i).left
                    End With
                End With
            Next i
            If GridWidth > 0 Then
                .Width = GridWidth - (GridPadding * 1)
            End If
            If screen.MousePointer <> 7 Then
                .visible = True
            End If
        End With

    Else
        txtFocus.SetFocus
        Me.Tabs.visible = False
        Me.Splitter.visible = False
        Me.FormFooter.Height = 0
    End If
    
    'SA 03/22/2012 - CR2713 Resize power par container width to screen width
    Me.SubformPowerBar.Width = Me.WindowWidth - 20
    Me.sfTabsHeadUser1.Width = Me.SubformPowerBar.Width
    Me.sfTabsHeadUser2.Width = Me.SubformPowerBar.Width
    Me.sfTabsHeadUser3.Width = Me.SubformPowerBar.Width
    
    'SA 8/2/2012 - Reposition restore/collapse buttons to right side of screen
    Child164.left = Me.WindowWidth - Child164.Width - 50

Form_ResizeExit:
    On Error Resume Next
    ' DS Mar 8, 2010 Change Request #1812 - Links no longer work when application window is resized
    Set MvGridMain = Me.SubForm.Form
    ' DS Mar 12 2010 save and restore filter that is lost during resize
    If Me.SubForm.Form.filter <> MvFilter Then
        Me.SubForm.Form.filter = MvFilter
    End If
    
    If Me.ScreenID <> 0 Then
    'suppress the suspend layout when creating a new screen for faster loading
         genUtils.ResumeLayout
    End If
   
Exit Sub

Form_ResizeError:
    ' Apr 5 2010 Adam Miles fix from discussion boards
    Select Case Err
        Case 0: 'Exit Sub ' Case 0 fires all the time, so I always exit sub on that one. The err in  question is 2759.
        Case 2759: 'Exit Sub 'occurs when a screen is open and maximized; ok to exit
        Case Else: MsgBox "Error " & Err.Number & " ( " & Err.Description & ")"
    End Select
 
    Resume Form_ResizeExit
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo UnloadError
RunEvent "Screen Unload", Me.ScreenID, Me.FormID
Set Scr(mvConfig.FormID) = Nothing
 screen.MousePointer = 0

Exit Sub

UnloadError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Clearing Screens Form Variable!", vbCritical, mvConfig.ScreenName
    Exit Sub
End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub imgFilters_Click()
    ScreenRecordSourceOnly
End Sub

Private Sub imgFilters_Collapsed_Click()
    ScreenRestore
End Sub

Private Sub imgSelection_Click()
    ScreenCollapse
End Sub

Private Sub MvGridMain_ApplyFilter(filter As String)

If mvSql.filter <> filter Then
    
    ' HC 3/15/2011 make the mvsql.filter the combination of the selection criteria filter and the right-click filter
    If selectionFilter <> vbNullString Then
       'DLC 11/15/11 added iif to prevent "and" being added where Filter is blank
        mvSql.filter = selectionFilter & IIf(LenB(filter) > 0, " and ", vbNullString) & filter
    Else
        mvSql.filter = filter
    End If
    ' hc 3/15/2011 - commented out.  the right click filter was overwriting the filter from the screen refresh
    'MvSql.Filter = Filter
    
    'SA 8/6/2012 - Added tab reset to make sure tabs get refilled when applying filters
    ReDim TabLoaded(Tabs.Pages.Count)
    
    'SA 1/17/2012 - Now calls TabsLoad instead of Tabs_Change
    TabsLoad        'Tabs_Change
    
    BuildTotalsCustom False
End If
End Sub

'/*DS Mar 12 2010 save and restore filter that is lost during resize
'  CR# 2564 fix - Clear phantom filters during refresh and toggle of 'Filter/Unfiltered' widget.
'               - When the 'Filtered/Unfiltered' widget is toggled on the grid , the true or
'                 false status of the grid filter flag can only be captured in the current event of the grid.
'               - If the grid filter flag is false(i.e., Unfiltered) but the grid filter string holds a filter condition,it will need
'                 the tabs to be requeried as the tabs could be reflecting (phantom)filtered record counts(if a sort was done).
'               - If custom filters are present when the grid is in the 'Unfiltered' state,those will need to be retained and
'                 applied;only grid filters need to be cleared.
'*/
Private Sub mvGridMain_Current()
    'SA 1/19/2012 - CR1967 Reset TabLoaded array on main record change
    ReDim TabLoaded(Tabs.Pages.Count)

    If Me.Tabs.Value >= glTabsUsed Then
        If Me.chkResetTabFilter Then
            Tabs.Pages(Tabs.Value).Controls(0).Form.filter = vbNullString
        End If
        
        If Nz(Tabs.Pages(Me.Tabs.Value).Controls(0).LinkChildFields, vbNullString) <> vbNullString Then
            Tabs.Pages(Tabs.Value).Controls(0).Form.Requery
        End If
    End If
    
    'CR# 2564 fix:Begin
    If MvGridMain.FilterOn Then
        On Error Resume Next
        MvFilter = Me.SubForm.Form.filter ' DS Mar 12 2010 save and restore filter that is lost during resize
        On Error GoTo 0
    ElseIf mvSql.filter <> vbNullString Then
        mvSql.filter = selectionFilter 'CR# 2564 - Retain custom filters
        MvFilter = mvSql.filter
        TabsLoad
    End If
    'CR# 2564 fix:End
    
    
    ' HC highlight the current row
    If Me.chkHighlight Then
        SendKeys "+ ", True
    End If
    
    'SA 05/21/2012 - CR2132 Added event
    If HasRowChangeEvent Then
        HasRowChangeEvent = RunEvent("Screen Row Change", Me.ScreenID, Me.FormID)
    End If
End Sub

Private Sub MvGridMain_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub MvGridMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub Page_2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub Page1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub Page2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub SortList_DblClick(Cancel As Integer)
On Error GoTo ErrorHandler
    'SA 03/22/2012 - CR1782 Remove selected item when double clicked
    Dim i As Integer
    i = SortList.ListIndex
    If i > -1 Then
        SortList.RemoveItem (i)
    End If
ExitRemove:
    ClsSCR.UpdateSortTip
Exit Sub
ErrorHandler:
    Resume ExitRemove
    Resume
End Sub

Private Sub Tabs_Change()
    'SA 1/17/2012 - Moved code in Tabs_Change to TabsLoad
    TabsLoad
    
    'SA 1/17/2012 - Added telemetry
    Telemetry.RecordAction "Tab Change Main", "<FRM>" & Me.Name & "</FRM><TABNUM>" & Me.Tabs.Value & "</TABNUM><TABNAME>" & genUtils.XMLEncode(Me.Tabs.Pages(Me.Tabs.Value).Caption) & "</TABNAME>", "Decipher Screens"
End Sub

Public Sub TabsLoad()
'SA 1/17/2012 - Moved code in Tabs_Change to this sub
'   Changed calls in MvGridMain_ApplyFilter, MvGridMain_Current, RefreshForm
'   to point to this sub
'SA 8/27/2012 - Made this method public
On Error GoTo ErrorHappened
    Dim CurTab As Byte
    Dim curPage As Byte
    
    If Me!Tabs.Value = 0 And Tabs.Pages(0).visible Then
        BuildTotalsCustom False
    End If
    
    If Me!Tabs.Value <= 0 Then Exit Sub
    
    curPage = Me!Tabs.Value
    
    'SA 1/19/2012 - CR1967 Added check to see if the current tab is already loaded
    If Not TabLoaded(curPage - 1) Then
        For CurTab = 1 To Me.Tabs.Pages.Count - 1
            ' If CurTab <> CurPage Then
            ' DS 17 Feb 2010 skip any tabs where the grid was removed
            If Tabs.Pages(CurTab).Controls(0).SourceObject <> vbNullString And CurTab <> curPage Then
                If Not TabLoaded(CurTab - 1) Then
                    Tabs.Pages(CurTab).Controls(0).Form.RecordSource = vbNullString
                End If
            End If
        Next CurTab
    
        With Tabs.Pages(curPage).Controls(0)
            If Nz(Tabs.Pages(curPage).Controls(0).LinkChildFields, vbNullString) = vbNullString Then
                'SetTabRecordSource (CurPage)
                ClsSCR.SetTabRecordSource curPage, mvConfig, mvSql
            Else
                If vbNullString & .Tag <> vbNullString Then
                    .Form.RecordSource = .Tag
                End If
            End If
        End With
        
        'SA 1/19/2012 - CR1967 Set tab status to loaded
        TabLoaded(curPage - 1) = True
    End If
    
ExitNow:
    On Error Resume Next
    Exit Sub
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error Changing tabs!", vbCritical
    Resume ExitNow
    Resume
End Sub

Private Sub TabsHead_Change()
    'SA 1/17/2012 Added this event for telemetry
    If Me.TabsHead.Value > 0 Then
        Telemetry.RecordAction "Tab Change Head", "<FRM>" & Me.Name & "</FRM><TABNUM>" & Me.TabsHead.Value & "</TABNUM><TABNAME>" & genUtils.XMLEncode(Me.TabsHead.Pages(Me.TabsHead.Value).Caption) & "</TABNAME>", "Decipher Screens"
    End If
End Sub

Private Sub ToggleSort_AfterUpdate()
    TogleAscDesc Me.ToggleSort
End Sub

Private Sub ToggleSort_Click()
    TogleAscDesc Me.ToggleSort
End Sub

Private Sub PrepareSplitterResize()
    Dim pgCt As Integer
    Dim pgIDx As Integer
    Dim pg As Page
    Dim pgCtrlCt As Integer
    Dim pgCtrlIdx As Integer
    Dim pgCtrl As Control
    
    genUtils.SuspendLayout Me
    
    screen.MousePointer = 7
    Me.Splitter.BorderColor = 0
    Me.Decoy.SetFocus
    Me.Tabs.Height = 1
    Me.Tabs.visible = False
    Me.SubForm.visible = False

    pgCt = Me.Tabs.Pages.Count - 1
    
    For pgIDx = 0 To pgCt
        Set pg = Me.Tabs.Pages.Item(pgIDx)
        
        pgCtrlCt = pg.Controls.Count - 1
        
        For pgCtrlIdx = 0 To pgCtrlCt
            Set pgCtrl = pg.Controls(pgCtrlIdx)
            With pgCtrl
                .Height = 1
                .Width = 1
                .visible = False
            End With
        Next
        
        pg.Height = 1
    Next
    
    Me.Tabs.Height = 1
    
    genUtils.ResumeLayout Me
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    MvSplitY = Y
    PrepareSplitterResize
End If
End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        screen.MousePointer = 7
    End If
    
    If Button = 1 And Y <> 0 Then
        If Me.FormFooter.Height + (MvSplitY + (Y * -1)) < (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) Then
                Me.FormFooter.Height = Me.FormFooter.Height + (MvSplitY + (Y * -1))
        End If
    End If
End Sub
Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        screen.MousePointer = 0
        Me.Splitter.BorderColor = Splitter.BackColor
        ClsSCR.ReleaseFooterForResize
    End If
End Sub

Private Sub CmdZoomMulti1_Click()
    ClsSCR.ZoomToControl CmboPrimaryMulti, LblPrimary.Caption
    ClsSCR.SetLabelFragment CmboPrimaryMulti, lblPrimaryMulti, 3
End Sub

Private Sub CmdZoomMulti2_Click()
    ClsSCR.ZoomToControl CmboSecondaryMulti, LblSecondary.Caption
    ClsSCR.SetLabelFragment CmboSecondaryMulti, lblSecondaryMulti, 2
End Sub

Private Sub CmdZoomMulti3_Click()
    ClsSCR.ZoomToControl CmboTertiaryMulti, LblTertiary.Caption
    ClsSCR.SetLabelFragment CmboTertiaryMulti, lblTertiaryMulti, 3
End Sub

Private Sub PopupMultiSelect(lvl As Byte)
On Error GoTo ErrorHappened
Dim frm As Form_SCR_CfgMultiItemSelect

' check the level, if greater than the available number on the screen there is an error, exit.
If lvl > 3 Then GoTo ExitNow

Set frm = New Form_SCR_CfgMultiItemSelect

With frm
    .InitData mvConfig.FormID, lvl
    .visible = True
End With

Do While frm.visible = True
    DoEvents
Loop

If frm.Results <> vbCancel Then
    Select Case lvl
        Case 1
            CmboPrimary_AfterUpdate
        Case 2
            CmboSecondary_AfterUpdate
        Case 3
            CmboTertiary_AfterUpdate
    End Select
End If


ExitNow:
    On Error Resume Next
    Set frm = Nothing
    Exit Sub
       
ErrorHappened:
    MsgBox Err.Description, vbCritical, "PopupMultiSelect"
    Resume ExitNow
    Resume
End Sub

Public Sub Resize()
    Form_Resize
End Sub

Public Sub BuildTotalsCustom(ByVal forceRebuild As Boolean)
    ClsSCR.BuildTotalsCustom forceRebuild
End Sub

Private Sub cmdScreenSaveSmall_Click()
    CmdScreenSave_Click
End Sub

Public Sub CmdScreenSave_Click() ' Routine to Save Screen Configuration by User
    'SA 03/22/2012 - CR2708 Moved code to boolean function SaveOpenScreen
    If ClsSCR.SaveOpenScreen(mvConfig) Then
        MsgBox "Current Layout and Position Saved!", vbInformation, "Save Confirmation"
    End If
End Sub

Public Function SaveOpenScreen() As Boolean
    SaveOpenScreen = ClsSCR.SaveOpenScreen(mvConfig)
End Function

Private Sub CmdScreenLoadSmall_Click()
    CmdScreenLoad_Click
End Sub

Public Sub CmdScreenLoad_Click() 'Load Users last save configuration
On Error GoTo ErrorHappened
'SA 03/22/2012 - CR2708 Made sub public
Dim db As DAO.Database
Dim rs As DAO.RecordSet
Dim FormatID As Long
Dim CalcID As Long
Dim StArray() As String
Dim i As Integer

DoCmd.Hourglass True
Set db = CurrentDb              'Retrieve values from table
Set rs = db.OpenRecordSet("SELECT * FROM SCR_SaveScreens WHERE ScreenID=" & mvConfig.ScreenID & " AND UserName=" & Chr(34) & Identity.UserName & Chr(34), dbOpenSnapshot)
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    
    Me.CmboListPrimaryBy = Nz(rs!PrimaryListBy, vbNullString)        'PrimaryListBy
    If Me.CmboListPrimaryBy <> vbNullString Then
        CmboListPrimaryBy_AfterUpdate
    End If
    
    If mvConfig.PrimaryListBoxMulti Then            'PrimaryCriteria
        If CmboPrimaryMulti.ListCount > 0 Then
            For i = Me.CmboPrimaryMulti.ListCount - 1 To 0 Step -1
                Me.CmboPrimaryMulti.RemoveItem i
            Next i
        End If
        StArray = Split(Nz(rs!PrimaryCriteria, vbNullString), ",")
        For i = 0 To UBound(StArray)
            Me.CmboPrimaryMulti.AddItem StArray(i)
        Next i
        ClsSCR.SetLabelFragment CmboPrimaryMulti, lblPrimaryMulti, 1
    Else
        Me.CmboPrimary = Nz(rs!PrimaryCriteria, vbNullString)
        Me.CmboPrimary_AfterUpdate
    End If
    
    Me.CmboListSecondaryBy = Nz(rs!SecondaryListBy, vbNullString)      'SecondaryListBy
    If Me.CmboListSecondaryBy <> vbNullString Then
        CmboListSecondaryBy_AfterUpdate
    End If
    
    If Me.CmboSecondaryMulti.ListCount > 0 Then     'SecondaryCriteria
        For i = Me.CmboSecondaryMulti.ListCount - 1 To 0 Step -1
            Me.CmboSecondaryMulti.RemoveItem (i)
        Next i
    End If
    If mvConfig.SecondaryListBoxMulti Then
        StArray = Split(Nz(rs!SecondaryCriteria, vbNullString), ",")
        For i = 0 To UBound(StArray)
            Me.CmboSecondaryMulti.AddItem StArray(i)
        Next i
        ClsSCR.SetLabelFragment CmboSecondaryMulti, lblSecondaryMulti, 2
    Else
        Me.CmboSecondary = Nz(rs!SecondaryCriteria, vbNullString)
    End If
    
    Me.CmboListTertiaryBy = Nz(rs!TertiaryListBy, vbNullString)
    If Me.CmboListTertiaryBy <> vbNullString Then
        CmboListTertiaryBy_AfterUpdate
    End If
    
    If Me.CmboTertiaryMulti.ListCount > 0 Then     'TertiaryCriteria
        For i = Me.CmboTertiaryMulti.ListCount - 1 To 0 Step -1
            Me.CmboTertiaryMulti.RemoveItem (i)
        Next i
    End If
    If mvConfig.TertiaryListBoxMulti Then
        StArray = Split(Nz(rs!TertiaryCriteria, vbNullString), ",")
        For i = 0 To UBound(StArray)
            Me.CmboTertiaryMulti.AddItem StArray(i)
        Next i
        ClsSCR.SetLabelFragment CmboTertiaryMulti, lblTertiaryMulti, 3
    Else
        Me.CmboTertiary = Nz(rs!TertiaryCriteria, vbNullString)
    End If
    
    Me.CmboFilterDte = Nz(rs!DateFilter, vbNullString)                 'DateFilterField
    Me.StartDte = Nz(rs!StartDte, vbNullString)                        'StartDate
    Me.EndDte = Nz(rs!EndDte, vbNullString)                            'EndDate
    
    'SA 03/22/2012 - CR1882 Clear and re-fill SortList
    Me.SortList.RowSource = vbNullString
    StArray = Split(Nz(rs!Sort, vbNullString), ",")
    For i = 0 To UBound(StArray) Step 2
        Me.SortList.AddItem StArray(i + 1) & ";" & StArray(i)
    Next i
    ClsSCR.UpdateSortTip
    
    Me.CmboFunction = Nz(rs!Function, vbNullString)                       'Function
    Me.CmboReports = Nz(rs!Report, 0)                           'Report
    'SA 03/22/2012 - Added for Work Files compatibility
    Me.SubformPowerBar.Tag = vbNullString & DLookup("PwrBarID", "SCR_ScreensPowerBars", "ScreenID = " & mvConfig.ScreenID & " AND Function = '" & Nz(rs!PowerBar, vbNullString) & "'")
    Me.SubformPowerBar.SourceObject = Nz(rs!PowerBar, vbNullString)       'PowerBar
                                    
    Dim StFilter As String
    Dim stFilterArray() As String
    Dim X As Integer
    StFilter = vbNullString
    If Nz(rs!filter, vbNullString) <> vbNullString Then
        ' add the split chars to the end so it will split correctly
        stFilterArray = Split(Nz(rs!filter, vbNullString) & ";", ";")
        For X = 0 To UBound(stFilterArray) - 1 Step 2
            StFilter = StFilter & stFilterArray(X) & ";'" & Replace(stFilterArray(X + 1), "'", "''") & "';"
        Next X
    End If
    If StFilter <> vbNullString Then
        StFilter = Mid(StFilter, 1, Len(StFilter) - 1)
    End If
    
    cmboFiltersSelected.RowSource = StFilter

        
    If Nz(rs!MainGridAdditionalFilter, vbNullString) <> vbNullString Then
        If mvSql.filter = vbNullString Then
            mvSql.filter = Nz(rs!MainGridAdditionalFilter, vbNullString)
        Else
            If Nz(rs!MainGridAdditionalFilter, vbNullString) <> vbNullString Then
                mvSql.filter = mvSql.filter & " AND " & rs!MainGridAdditionalFilter
            End If
        End If
    End If
    
    cmdRefresh_Click                                            'Refreshing Screen
       
    ' DS Apr 30 2010 fix lack of restore of filters
    Me.SubForm.Form.filter = Nz(rs!MainGridAdditionalFilter, vbNullString)
    Me.SubForm.Form.FilterOn = True

    Me.SubForm.Form.OrderBy = Nz(rs!MainGridAdditionalSort, vbNullString)
    Me.SubForm.Form.OrderByOn = True
    
    If Nz(Me.CmboLayouts.Value, 0) <> Nz(rs!layout, -1) Then
        Me.CmboLayouts.Value = rs!layout                'Applying Layout
        Me.ApplyLayout
    End If
    Me.chkHighlight = rs!highlightcurrentrow            ' highlight current row
    'SA 1/30/2012 - Added new checkbox
    Me.chkResetTabFilter = rs!ResetTabFilter            ' Reset tab filter
    
    If Nz(Me.CmboTotals.Value, 0) <> Nz(rs!Totals, -1) Then
        Me.CmboTotals.Value = rs!Totals                'Applying Totals
        CmdApplyTotals_Click
    End If

    If rs!ConditionalFormats <> vbNullString Then             'Applying Conditional Formatting
        StArray = Split(Nz(rs!ConditionalFormats, vbNullString), ",")
        For i = 0 To UBound(StArray)
            FormatID = Nz(DLookup("CondFormatID", "SCR_ScreensCondFormats", "ScreenID = " & mvConfig.ScreenID & " and FormatName = " & Chr(34) & StArray(i) & Chr(34)), 0)
            Me.SubformCondFormats.Form.ApplyFormatAdd FormatID, StArray(i)
        Next i
        Me.SubformCondFormats.Form.ApplyFormats
    End If

    If rs!CustomCalculations <> vbNullString Then             'Applying Custom Calculations
        StArray = Split(Nz(rs!CustomCalculations, vbNullString), ",")
        For i = 0 To UBound(StArray)
            CalcID = Nz(DLookup("CalcID", "SCR_ScreensCalculations", "ScreenID = " & mvConfig.ScreenID & " and CalcName = " & Chr(34) & StArray(i) & Chr(34)), 0)
            Me.SubformCalcs.Form.ApplyAdd CalcID, StArray(i)
        Next i
        Me.SubformCalcs.Form.ApplyCalcs
    End If
    
    ' DS 9 Mar 2010 fix to bug in the Decipher Save/Restore Screen feature
    ' The value returned by "Me.SubForm.Form.Recordset.RecordCount" is inaccurate without being preceded by a movefist/movelast on the recordset
    ' for performance reasons, we do not want to check to see if saved record is still in the grid since MoveLast would be slow
    ' If Nz(Me.SubForm.Form.Recordset.RecordCount, 0) >= Nz(rs!maingridrecpos, 0) - 1 Then                            End If
        ' Me.SubForm.Form.Recordset.AbsolutePosition = Nz(rs!maingridrecpos, 0) - 1 'Moving to last current record
    If Nz(rs!maingridrecpos, 0) - 1 > 0 Then
        'Moving to last current record
        Me.SubForm.Form.RecordSet.Move Nz(rs!maingridrecpos, 0) - 1
    End If
    
    Select Case Nz(rs!Tab, 0)       'Setting Current Tab Record
        Case -1
        Case 0
            If Nz(Me.Subform_2.Form.RecordSource, vbNullString) <> vbNullString Then
                Me.Page_2.SetFocus
                Me.Subform_2.Form.Requery
                If Me.Subform_2.Form.RecordSet.recordCount > 0 And Me.Subform_2.Form.RecordSet.recordCount >= Nz(rs!TabGridRecPos, 0) - 1 Then
                    Me.Subform_2.Form.RecordSet.AbsolutePosition = Nz(rs!TabGridRecPos, 0) - 1
                End If
            End If
        Case Else
            Me.Controls.Item("Page" & CStr(Nz(rs!Tab + 1, 0))).SetFocus
            Me.Controls.Item("Subform" & CStr(rs!Tab + 1)).Form.Requery
            If Me.Controls.Item("Subform" & CStr(rs!Tab + 1)).Form.RecordSet.recordCount > 0 And Me.Controls.Item("Subform" & CStr(rs!Tab + 1)).Form.RecordSet.recordCount >= Nz(rs!TabGridRecPos, 0) - 1 Then
                Me.Controls.Item("Subform" & CStr(rs!Tab + 1)).Form.RecordSet.AbsolutePosition = Nz(rs!TabGridRecPos, 0) - 1
            End If
    End Select
    
    Me.SubForm.SetFocus     'Setting Focus to main grid

End If

AllDone:
On Error Resume Next
rs.Close
Set rs = Nothing
db.Close
Set db = Nothing
DoCmd.Hourglass False
Exit Sub

ErrorHappened:
MsgBox "Error Loading Saved Screen" & vbCrLf & vbCrLf & Err.Description
Resume AllDone

End Sub

Private Sub ScreenCollapse()
On Error GoTo ErrorHappened
    Dim labelTop As Integer
    DoCmd.Hourglass True
    
    stateForm.SetState_Collapsed False
    DoEvents
    
    genUtils.SuspendLayout Me
    
    Me.imgSelection_Collapsed.visible = True
   
    Me.cmdScreenSaveSmall.visible = True
    Me.CmdScreenLoadSmall.visible = True
    Me.CmdRefreshSmall.visible = True
    
    labelTop = lblSelectionLabel.top + lblSelectionLabel.Height + 18
    lblFiltersLabel.top = labelTop
    lblDateRangeLabel.top = labelTop
    lblSortingLabel.top = labelTop
    lblTotalsLabel.top = labelTop
    lblLayoutLabel.top = labelTop
    lblLayout2Label.top = labelTop - 30
    With imgFilters_Collapsed
        .visible = True
        .top = labelTop - 20
    End With
    
    lblLayoutLabel.Width = cmdScreenSaveSmall.left - lblLayoutLabel.left - 40
    lblLayout2Label.Width = cmdScreenSaveSmall.left - lblLayoutLabel.left - 40
        
    lblTools2Label.Width = cmdScreenSaveSmall.left - lblToolsLabel.left - 40
    lblToolsLabel.Width = cmdScreenSaveSmall.left - lblToolsLabel.left - 40
       
    genUtils.LoadScreenProfile "MainScreensCollapseAll", Me
    
    Me.SubForm.SetFocus
    Me.TabsHead.Height = 0
    Me.FormHeader.Height = Me.TabsHead.Height
    
    Form_Resize
    
    'SA 03/22/2012 - Added telemetry
    'SA 10/1/2012 - CR3128 XML encode screen name to prevent telemetry errors
    Telemetry.RecordAction "Screen Collapse", "<SN>" & genUtils.XMLEncode(Me.ScreenName) & "</SN>", "Decipher Screens"
    
AllDone:
    genUtils.ResumeLayout Me
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical
    Resume AllDone
    Resume
End Sub

Private Sub ScreenRestore()
On Error GoTo ErrorHappened
    DoCmd.Hourglass True
    
    stateForm.SetState_Restored False
    DoEvents
    
    genUtils.SuspendLayout Me
    
    Me.FormHeader.Height = 5000 '2985
    Me.TabsHead.Height = 5000 '2985
    
    Me.SubForm.SetFocus
    
    genUtils.LoadScreenProfile "MainScreensRestoreAll", Me
   
    txtFocus.SetFocus
    imgSelection_Collapsed.visible = False
   
    cmdScreenSaveSmall.visible = False
    CmdScreenLoadSmall.visible = False
    CmdRefreshSmall.visible = False
    
    ClsSCR.RestoreTitles
    
    With imgFilters_Collapsed
        .visible = False
        .top = lineSelection.top + lineSelection.Height + 75
    End With
    
    
    Me.TabsHead.Height = 2985 'pnlFilters.Top + pnlFilters.Height + 50 '2985
    Me.FormHeader.Height = 2985 'pnlFilters.Top + pnlFilters.Height + 50 '2985

    Form_Resize

    'SA 03/22/2012 - Added telemetry
    'SA 10/1/2012 - CR3128 XML encode screen name to prevent telemetry errors
    Telemetry.RecordAction "Screen Restore", "<SN>" & genUtils.XMLEncode(Me.ScreenName) & "</SN>", "Decipher Screens"
    
AllDone:
    genUtils.ResumeLayout Me
    DoCmd.Hourglass False
    DoEvents
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical
    Resume AllDone
    Resume
End Sub

Private Sub ScreenRecordSourceOnly()
On Error GoTo ErrorHappened
    DoCmd.Hourglass True

    stateForm.SetState_SelectionOnly False
    DoEvents
    
    genUtils.SuspendLayout Me
    
    Me.TabsHead.Height = 1752 'lineSelection.Top + lineSelection.Height + 75 + imgFilters_Collapsed.Height + 50 ' 1752
    Me.FormHeader.Height = 1752 'lineSelection.Top + lineSelection.Height + 75 + imgFilters_Collapsed.Height + 50 '1752
    
    genUtils.LoadScreenProfile "MainScreensPrimaryOnly", Me
    
    txtFocus.SetFocus
    imgSelection_Collapsed.visible = False
    cmdScreenSaveSmall.visible = False
    CmdScreenLoadSmall.visible = False
    CmdRefreshSmall.visible = False
     
    ClsSCR.RestoreTitles
      
    Me.TabsHead.Height = 1752 'pnlSelection.Top + pnlSelection.Height + 75 + imgFilters_Collapsed.Height + 50 '1752
    Me.FormHeader.Height = 1752 'pnlSelection.Top + pnlSelection.Height + 75 + imgFilters_Collapsed.Height + 50 '1752
   
    Form_Resize

    'SA 03/22/2012 - Added telemetry
    'SA 10/1/2012 - CR3128 XML encode screen name to prevent telemetry errors
    Telemetry.RecordAction "Screen Select Only", "<SN>" & genUtils.XMLEncode(Me.ScreenName) & "</SN>", "Decipher Screens"
    
AllDone:
    genUtils.ResumeLayout Me
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical
    Resume AllDone
    Resume
End Sub

Private Sub stateForm_CollapseAll()
    ScreenCollapse
End Sub

Private Sub stateForm_RecordSourceOnly()
    ScreenRecordSourceOnly
End Sub

Private Sub stateForm_RestoreAll()
    ScreenRestore
End Sub

Private Sub collapseRecordSetButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ScreenRecordSourceOnly
End Sub

Private Sub collapseAllButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ScreenCollapse
End Sub

Private Sub restoreButton_Click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    ScreenRestore
End Sub

Private Sub imgSelection_Collapsed_Click()
    ScreenRecordSourceOnly
End Sub

Private Sub lblFiltersLabel_Click()
    If CmdFilterEdit.Height > 1 Then
        ScreenRecordSourceOnly
   Else
        ScreenRestore
   End If
End Sub

Private Sub lblSelectionLabel_Click()
    If Not cmdScreenSaveSmall.visible Then
        ScreenCollapse
    Else
        ScreenRecordSourceOnly
    End If
End Sub


Private Sub SetRestoreCollapseMenu()
On Error GoTo ErrorHappened
    Dim objCommandBar As CommandBar
    Dim objCommandBarButton As CommandBarButton
  
    Set collapseAllButton = Nothing
    Set restoreButton = Nothing
    Set collapseRecordSetButton = Nothing
    Set restoreMenu = Nothing
            
    ' clear the existing DecipherRestoreCollapseMenu
    genUtils.ClearMenu (DecipherRestoreCollapseMenu)
    
    Set objCommandBar = CommandBars.Add(Name:=DecipherRestoreCollapseMenu, position:=msoBarPopup, Temporary:=False, MenuBar:=False)
    Set restoreMenu = objCommandBar
    With restoreMenu
        Set objCommandBarButton = .Controls.Add(msoControlButton, , , , False)
        Set restoreButton = objCommandBarButton
        With restoreButton
            .Caption = "Restore All"
            .FaceId = 298
            .Tag = "DecipherRestore"
            .style = msoButtonIconAndCaption
            .ToolTipText = "Restore all controls"
        End With
        Set objCommandBarButton = .Controls.Add(msoControlButton, , , , False)
        Set collapseRecordSetButton = objCommandBarButton
        With collapseRecordSetButton
            .BeginGroup = True
            .Caption = "Collapse To Selection"
            .FaceId = 1023
            .Tag = "DecipherCollapseRecordSet"
            .style = msoButtonIconAndCaption
            .ToolTipText = "Collapse To Selection Only"
        End With
        Set objCommandBarButton = .Controls.Add(msoControlButton, , , , False)
        Set collapseAllButton = objCommandBarButton
        With collapseAllButton
            .Caption = "Collapse All"
            .Tag = "DecipherCollapseAll"
            .FaceId = 3838
            .style = msoButtonIconAndCaption
            .ToolTipText = "Collapse To Tabs"
        End With
    End With

    ' Associate the control shortcutmenu bar with DecipherRestoreCollapseMenu
    ' example
        Me.Form.ShortcutMenu = True
        Me.Form.ShortcutMenuBar = DecipherRestoreCollapseMenu
    Exit Sub

ErrorHappened:
    
    Set collapseAllButton = Nothing
    Set restoreButton = Nothing
    Set collapseRecordSetButton = Nothing
    Set restoreMenu = Nothing
End Sub

Public Function RequeryGrid() As Boolean
'SA 9/4/2012 - Added method to requery main recordset from powerbars etc.
On Error GoTo ErrorHappened
    Dim Result As Boolean

    Me.GridForm.RecordSet.Requery

    Result = True
ExitNow:
On Error Resume Next
    RequeryGrid = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function
