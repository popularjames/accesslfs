Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : SCR_ClsMainScreens
' Author    : SA
' Date      : 11/7/2012
' Purpose   : Moved methods from SCR_MainScreens to save them from being loaded with each form
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private myForm As Form_SCR_MainScreens
Private TogValue As Boolean


Public Property Let SetMyForm(ByRef frm As Form_SCR_MainScreens)
    Set myForm = frm
End Property

Public Property Let SetTogValue(ByVal Value As Boolean)
    TogValue = Value
End Property

Public Property Get GetTogValue() As Boolean
    GetTogValue = TogValue
End Property

Private Sub Class_Initialize()
    TogValue = True
End Sub

Public Function GetName(ByVal Request As String, ByVal originalName As String) As String
'Get filter name from input box
On Error GoTo ErrorHappened
    Dim FilterName As String
    
    FilterName = InputBox(Request, "Save Filter", Replace(originalName, "'", vbNullString))
    FilterName = Replace(FilterName, "'", "''")
    FilterName = EscapeQuotes(FilterName)
ExitNow:
On Error Resume Next
    GetName = FilterName
Exit Function
ErrorHappened:
    FilterName = vbNullString
    MsgBox Err.Description, vbCritical, "SCR_ClsMainScreens:GetName"
    Resume ExitNow
    Resume
End Function

Public Function GetFilterData(ByVal FilterID As Long) As String
'Get filter sql from id
On Error GoTo ErrorHappened
    Dim filterString As String
    filterString = DLookup("FilterSql", "SCR_ScreensFilters", "FilterId = " & FilterID)

ExitNow:
On Error Resume Next
    GetFilterData = filterString
Exit Function
ErrorHappened:
    filterString = vbNullString
    Resume ExitNow
    Resume
End Function

Public Sub SetReportCriteria(ByVal RptConfig As CT_ClsRpt, ByVal ReportID As Long, ByVal multiList As Control, ByVal listBox As Control, ByVal controlValue As Integer)
'Set Report Criteria
On Error GoTo ErrorHappened
    Dim X As String
    Dim N As Integer
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    
    Set db = CurrentDb

    Set rs = db.OpenRecordSet("SCR_ScreensReportCriteria")
    With multiList
        For N = 0 To .ListCount - 1
            rs.AddNew
            rs!Auditor = Identity.Auditor
            rs!RptId = ReportID
            rs!CriteriaType = controlValue
            rs!Criteria = .list(N)
            rs.Update
        Next N
    End With
    rs.Close
     
    'added by GI 09/10/07 To bound primary criteria control to default bound column - start
     N = InStr(1, listBox.RowSource, " order by ")
     X = left(listBox.RowSource, N) & " and Bound = True"
     Set rs = db.OpenRecordSet(X)
     X = rs!FieldName
     'added by GI 09/10/07 To bound secondary criteria control to default bound column - en
         
    rs.Close
    db.Close
     
    If vbNullString & RptConfig.Criteria <> vbNullString Then
        RptConfig.Criteria = RptConfig.Criteria & " AND "
    End If
     
    RptConfig.Criteria = RptConfig.Criteria & X & _
    " IN (Select Criteria From SCR_ScreensReportCriteria " & _
    " WHERE Auditor = " & Chr(34) & Identity.Auditor & Chr(34) & _
    " and RptId = " & ReportID & " AND CriteriaType = " & controlValue & " AND Criteria <> = '')"
 
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "SCR_ClsMainScreens:SetReportCriteria"
    Resume ExitNow
    Resume
End Sub

Public Sub SetLabelFragment(ByRef list As Control, ByRef Lbl As Label, Optional ByVal itemCount As Integer = 0)
'SA 11/26/2012 - Made itemCount optional since it is unused. Kept for backward compatibility
On Error GoTo ErrorHappened
    
    Const MaxCaptionLength As Integer = 255
    'SA 1/18/2012 - CR1042 Variables used to set label
    Dim X As Long
    Dim strCaption As String
    Dim BlnTruncated As Boolean
    
    DoCmd.Hourglass True
    
    For X = 0 To list.ListCount - 1
        If Len(strCaption) > MaxCaptionLength Then
            BlnTruncated = True
            Exit For
        End If
        strCaption = strCaption & list.list(X) & ", "
    Next X
    
    'Remove the trailing comma
    If list.ListCount > 0 Then
        strCaption = left(strCaption, Len(strCaption) - 2)
    End If
    
    If BlnTruncated Or Len(strCaption) > MaxCaptionLength Then
        strCaption = left(strCaption, MaxCaptionLength - 3) & "..."
    End If

    Lbl.ControlTipText = strCaption
        
    If Len(strCaption) > 21 Then
        strCaption = "Items selected: " & list.ListCount
    End If

    Lbl.Caption = vbNullString
    If list.ListCount > 0 Then
        Lbl.Caption = strCaption
    ElseIf list.ListCount = 0 Then
        Lbl.Caption = "(No Items Selected)"
        Lbl.ControlTipText = vbNullString
    End If
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    'Display error in label
    Lbl.ControlTipText = vbNullString
    Lbl.Caption = "(Error)"
    
    Resume ExitNow
    Resume
End Sub

Public Sub ZoomToControl(ByRef ctrl As Control, Optional ByVal Title As String = vbNullString)
'Open zoom dialog
On Error GoTo ErrorHappened
    Dim FrmZoom As New Form_SCR_Zoom
    
    FrmZoom.Move ctrl.left
    With FrmZoom
        .Bind ctrl
        .Title = Title
        .visible = True
    End With
    
    Do While FrmZoom.visible
        DoEvents
    Loop
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "SCR_ClsMainScreens:ZoomToControl"
    Resume ExitNow
    Resume
End Sub

Public Sub LoadUserTab(ByRef tabPage As Page, ByRef tabConfig As CnlyScreenTabsHead)
On Error GoTo ErrorHappened
    With tabPage
        .visible = tabConfig.ShowTab
        .Caption = tabConfig.Caption
        .ControlTipText = tabConfig.ControlTip
        .StatusBarText = tabConfig.StatusBar
        .Controls.Item(0).SourceObject = tabConfig.SubForm
        If LenB(tabConfig.Image) > 0 Then
            .Picture = tabConfig.Image
        End If
    End With
ExitNow:

Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Public Function ApplyListedFilters() As String
On Error GoTo ApplyListedFiltersError
    Dim i As Integer
    Dim strFilter As String
    Dim strValue As String

    strFilter = vbNullString
    
    For i = 0 To myForm.cmboFiltersSelected.ListCount - 1
        If Nz(myForm.cmboFiltersSelected.ItemData(i), vbNullString) <> vbNullString Then
            strValue = GetFilterData(myForm.cmboFiltersSelected.ItemData(i))
            If Nz(strValue, vbNullString) <> vbNullString Then
                strFilter = strFilter & "(" & strValue & ") AND "
            End If
        End If
    Next

    If Len(strFilter) > 4 Then
        strFilter = left(strFilter, Len(strFilter) - 4)
    End If
        
    RunEvent "Apply Filter", myForm.ScreenID, myForm.FormID

ApplyListedFiltersExit:
On Error Resume Next
    ApplyListedFilters = strFilter
Exit Function
ApplyListedFiltersError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error applying saved filter", vbCritical, "Filter"
    Resume ApplyListedFiltersExit
    Resume
End Function

Public Sub ResetTabs()
'SA 1/19/2012 - CR 1967 Reset the RecordSource to null for all tabs
On Error GoTo ErrorHandler
    Dim CurTab As Byte

    For CurTab = 1 To myForm.Tabs.Pages.Count - 1
        If LenB(myForm.Tabs.Pages(CurTab).Controls(0).SourceObject) > 0 Then
            myForm.Tabs.Pages(CurTab).Controls(0).Form.RecordSource = vbNullString
        End If
    Next CurTab
Exit Sub
ErrorHandler:

End Sub

Public Sub SetTabRecordSource(ByVal curPage As Integer, ByRef mvConfig As CnlyScreenCfg, ByRef mvSql As CnlyScreenSQL)
On Error GoTo BuildError
    Dim SQL As String
    Dim pos As Long
    Dim tmpStr As String
    Dim filter As String
    
    myForm.Tabs.Pages(curPage).Controls(0).visible = True
    If CByte(mvConfig.Tabs(curPage - 1).SourceType) = 1 Then 'query
        SQL = CurrentDb.QueryDefs(myForm.Tabs.Pages(curPage).Controls(0).Tag).SQL
    Else
        SQL = "SELECT * FROM " & myForm.Tabs.Pages(curPage).Controls(0).Tag
    End If
           
    'Chop Off Ending ;
    pos = InStr(1, SQL, ";")
    If pos > 0 Then SQL = left(SQL, pos - 1)

    'BUILD THE WHERE CLAUSE
    filter = " WHERE " & myForm.BuildWhere()
    pos = InStr(1, UCase(SQL), "GROUP BY")
    If pos > 0 Then
        tmpStr = Mid(SQL, 1, pos - 1)
        If vbNullString & filter <> vbNullString Then
            tmpStr = tmpStr & filter & " "
        End If
        SQL = tmpStr & Mid(SQL, pos)
    Else
        SQL = SQL & " " & filter
    End If

    mvSql.SqlTotals = SQL

BuildExit:
On Error Resume Next
    myForm.Tabs.Pages(curPage).Controls(0).Form.RecordSource = SQL
    myForm.Tabs.Pages(curPage).Controls(0).Form.Requery
Exit Sub
BuildError:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error populating Totals!", vbCritical, "SQL ERROR"
    Resume BuildExit
    Resume
End Sub

Public Function BuildWhere(ByRef mvSql As CnlyScreenSQL) As String
    Dim filter As String
    Dim HoldFilter As String
    
    With mvSql
        filter = vbNullString & .WherePrimary
        If vbNullString & .WhereSecondary <> vbNullString Then
            filter = filter & " and " & .WhereSecondary
        End If
        If vbNullString & .WhereTertiary <> vbNullString Then
            filter = filter & " and " & .WhereTertiary
        End If
        If vbNullString & .WhereDates <> vbNullString Then
            filter = filter & " and " & .WhereDates
        End If
        If vbNullString & .filter <> vbNullString Then
            filter = filter & " and " & .filter
        End If
        HoldFilter = ApplyListedFilters
        If vbNullString & HoldFilter <> vbNullString And HoldFilter <> .filter Then
            filter = filter & " AND " & HoldFilter
        End If
        
    End With

    BuildWhere = filter
End Function

Public Function BuildWherePrimary(ByRef mvConfig As CnlyScreenCfg) As String
    Dim SQL As String
    Dim returnString As String
    
    SQL = " " & mvConfig.PrimaryField
    If mvConfig.PrimaryListBoxMulti Then
        returnString = Trim(BuildMultiItemSQL(myForm.CmboPrimaryMulti.Object, mvConfig.PrimaryQualifier))
        If returnString <> vbNullString Then
            SQL = SQL & " in " & returnString
        Else
            SQL = vbNullString
        End If
    Else
        returnString = mvConfig.PrimaryQualifier & myForm.CmboPrimary & mvConfig.PrimaryQualifier
        If returnString <> vbNullString Then
            SQL = SQL & " = " & returnString & " "
        Else
            SQL = vbNullString
        End If
    End If
    BuildWherePrimary = SQL
End Function

Public Function BuildWhereSecondary(ByRef mvConfig As CnlyScreenCfg) As String
    Dim SQL As String
    Dim returnString As String
    
    SQL = " " & mvConfig.SecondaryField
    If mvConfig.SecondaryListBoxMulti Then
        returnString = BuildMultiItemSQL(myForm.CmboSecondaryMulti.Object, mvConfig.SecondaryQualifier)
        If returnString <> vbNullString Then
            SQL = SQL & " in " & returnString
        Else
            SQL = vbNullString
        End If
    Else
        returnString = mvConfig.SecondaryQualifier & myForm.CmboSecondary & mvConfig.SecondaryQualifier
        If returnString <> vbNullString Then
            SQL = SQL & " = " & returnString & " "
        Else
            SQL = vbNullString
        End If
    End If
    BuildWhereSecondary = SQL
End Function
Public Function BuildWhereTertiary(ByRef mvConfig As CnlyScreenCfg) As String
    Dim SQL As String
    Dim returnString As String

    SQL = " " & mvConfig.TertiaryField & " "
    If mvConfig.TertiaryListBoxMulti Then
        returnString = BuildMultiItemSQL(myForm.CmboTertiaryMulti.Object, mvConfig.TertiaryQualifier)
        If returnString <> vbNullString Then
            SQL = SQL & " in " & returnString
        Else
            SQL = vbNullString
        End If
    Else
        returnString = mvConfig.TertiaryQualifier & myForm.CmboTertiary & mvConfig.TertiaryQualifier
        If returnString <> vbNullString Then
            SQL = SQL & " = " & returnString & " "
        Else
            returnString = vbNullString
        End If
    End If
    BuildWhereTertiary = SQL
End Function
Public Function BuildMultiItemSQL(ByRef ctlList As MSForms.listBox, ByVal txtQual As String)
    Dim SQL As String
    Dim X As Integer

    With ctlList
        'Make Sure the are items in the list
        If .ListCount = 0 Then
            BuildMultiItemSQL = vbNullString
            GoTo ExitBuild
        End If
        
        SQL = "("
        For X = 0 To ctlList.ListCount - 1
            If X <> 0 Then SQL = SQL & ","
            SQL = SQL & txtQual & .list(X) & txtQual
        Next X
        SQL = SQL & ")"
        
    End With
    
    BuildMultiItemSQL = SQL
ExitBuild:

Exit Function
ErrorBuild:
    MsgBox Err.Description & String(2, vbCrLf) & "Error building SQL for Multi Item Select!", vbCritical, "Function BuildMultiItemSQL"
    BuildMultiItemSQL = vbNullString
    Resume ExitBuild
End Function

Public Sub PopulateCriteriaLists(ByRef listBy As Control, ByRef criteriaList As Control, _
        ByVal itemNumber As Integer, ByRef mvConfig As CnlyScreenCfg)
On Error GoTo PopulateCriteriaListsError
    Dim SQL As String
    Dim listByField As String
    Dim pastListBy As Boolean
    Dim colWidths As String
    Dim listWidth As Single
    Dim OrderBy As String
    Dim X As Integer
    
    DoCmd.Hourglass True
    listByField = vbNullString & listBy
    OrderBy = IIf(listBy.Column(4, listBy.ListIndex) = dbCurrency, " Desc", vbNullString)
    SQL = "SELECT " & listByField & " "
    colWidths = CStr(listBy.Column(3, listBy.ListIndex)) & " in "
    For X = 0 To listBy.ListCount - 1
        If listBy.Column(0, X) <> listByField Then
            SQL = SQL & ", " & listBy.Column(0, X) & " "
            colWidths = colWidths & ";" & CStr(listBy.Column(3, X)) & " in "
        Else
            pastListBy = True
        End If
        If listBy.Column(1, X) Then  'BoundField
            If listBy.Column(0, X) <> listByField Then
                criteriaList.BoundColumn = X + IIf(pastListBy, 1, 2)
            Else
                criteriaList.BoundColumn = 1
            End If
        End If
        If listBy.Column(2, X) Then
            If listBy.Column(0, X) <> listByField Then
                If itemNumber = 1 Then
                    mvConfig.PrimaryAlternatePos = X + IIf(pastListBy, 1, 2)
                ElseIf itemNumber = 2 Then
                    mvConfig.SecondaryAlternatePos = X + IIf(pastListBy, 1, 2)
                ElseIf itemNumber = 3 Then
                    mvConfig.TertiaryAlternatePos = X + IIf(pastListBy, 1, 2)
                End If
            Else
                If itemNumber = 1 Then
                    mvConfig.PrimaryAlternatePos = 1
                ElseIf itemNumber = 2 Then
                    mvConfig.SecondaryAlternatePos = 1
                ElseIf itemNumber = 3 Then
                    mvConfig.TertiaryAlternatePos = 1
                End If
            End If
        End If
        listWidth = listWidth + listBy.Column(3, X)
    Next X
    
    'Buid the sql command
    
    If itemNumber = 1 Then
        SQL = SQL & "FROM " & mvConfig.PrimaryListBoxRecordSource & " "
    ElseIf itemNumber = 2 Then
        SQL = SQL & "FROM " & mvConfig.SecondaryListBoxRecordSource & " "
        ' if the secondary list box is dependent on primary, add where clause for primary
        If mvConfig.SecondaryListBoxDependency Then
            SQL = SQL & "WHERE " & BuildWherePrimary(mvConfig)
        End If
    ElseIf itemNumber = 3 Then
        SQL = SQL & "From " & mvConfig.TertiaryListBoxRecordSource & " "
        ' if tertiary list box dependent on secondary, add where clause for secondary
        If mvConfig.TertiaryListBoxPrimaryDependency Then
            SQL = SQL & " WHERE " & BuildWherePrimary(mvConfig)
            If mvConfig.TertiaryListBoxDependency Then
                SQL = SQL & " AND " & BuildWhereSecondary(mvConfig)
            End If
        ElseIf mvConfig.TertiaryListBoxDependency Then
            SQL = SQL & " WHERE " & BuildWhereSecondary(mvConfig)
        End If
    End If
    
    SQL = SQL & " Order By " & listByField & OrderBy
    With criteriaList
        .listWidth = (listWidth + 0.2) * 1440
        .ColumnCount = listBy.ListCount
        .ColumnWidths = colWidths
        .RowSource = SQL
    End With
PopulateCriteriaListsExit:
On Error Resume Next
    DoCmd.Hourglass False
Exit Sub
PopulateCriteriaListsError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Setting Criteria Lists!", vbCritical, mvConfig.ScreenName
    Resume PopulateCriteriaListsExit
End Sub

Public Sub SetDefaultSort(ByRef mvConfig As CnlyScreenCfg)
On Error GoTo SetDefaultSortError
    'SA 03/22/2012 - CR1782 Changed SortList from ActiveX to Access object
    Dim SortRst As DAO.RecordSet
    Dim SortDb As DAO.Database
    Dim SQL As String
    
    SQL = "SELECT SortOrder,FieldName FROM SCR_ScreensSorts WHERE ScreenID=" & _
            mvConfig.ScreenID & " AND SortName='Default' ORDER BY SortIndex"
    
    Set SortDb = CurrentDb
    Set SortRst = SortDb.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    
    With SortRst
        Do Until .EOF
            myForm.SortList.AddItem !SortOrder & ";" & !FieldName
            .MoveNext
        Loop
    End With
    
SetDefaultSortExit:
On Error Resume Next
    UpdateSortTip
    SortRst.Close
    Set SortRst = Nothing
    Set SortDb = Nothing
    Exit Sub
SetDefaultSortError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error Loading Default Sort!", vbCritical, mvConfig.ScreenName
    Resume SetDefaultSortExit
    Resume
End Sub

Public Sub UpdateSortTip()
On Error GoTo ErrorHandler
    'SA 03/22/2012 - CR1782 Display sort fields in control text
    Dim i As Integer
    Dim sortTip As String
    
    'Build control tip for SortList that displays all sort fields
    If myForm.SortList.ListCount > 0 Then
        sortTip = "ORDER BY "
        For i = 0 To myForm.SortList.ListCount - 1
            If myForm.SortList.Column(0, i) = "A" Then
                sortTip = sortTip & myForm.SortList.Column(1, i) & ", "
            Else
                sortTip = sortTip & myForm.SortList.Column(1, i) & " DESC, "
            End If
        Next
        
        sortTip = left(sortTip, Len(sortTip) - 2)
        myForm.SortList.ControlTipText = left(sortTip, 255)
    Else
        myForm.SortList.ControlTipText = vbNullString
    End If
    
ExitNow:
    
Exit Sub
ErrorHandler:
    myForm.SortList.ControlTipText = vbNullString
    Resume ExitNow
    Resume
End Sub

Public Function BuildDetail(ByRef mvSql As CnlyScreenSQL, ByRef mvConfig As CnlyScreenCfg) As String
On Error GoTo BuildError
    Dim Result As String
    
    With mvSql
        .Select = "SELECT " & mvConfig.PrimaryRecordSource & ".* "
        .From = " FROM " & mvConfig.PrimaryRecordSource & " "
        
        '** START THE PRIMARY LIST BOX ***
        .WherePrimary = BuildWherePrimary(mvConfig)
        ' make sure the primary selection has been made, if not exit sub
        If Nz(.WherePrimary, vbNullString) = vbNullString Then
            Exit Function
        End If
            
        '*** START THE SECONDARY LIST BOX ***
        If mvConfig.SecondaryListBoxUse And (myForm.CmboSecondary.ListIndex <> -1 Or myForm.CmboSecondaryMulti.ListCount <> 0) Then
            .WhereSecondary = " " & BuildWhereSecondary(mvConfig)
        Else
            .WhereSecondary = vbNullString
        End If
        
        ' ** START THE Tertiary LIST BOX **
        If mvConfig.TertiaryListBoxUse And (myForm.CmboTertiary.ListIndex <> -1 Or myForm.CmboTertiaryMulti.ListCount <> 0) Then
            .WhereTertiary = " " & BuildWhereTertiary(mvConfig)
        Else
            .WhereTertiary = vbNullString
        End If
        
        If mvConfig.DateUse Then .WhereDates = BuildDateCriteria(myForm)
        ' added to change the filtering to be on sql
        .filter = ApplyListedFilters
        .OrderBy = BuildDetailSort(myForm, True)
        .SqlAll = .Select & .From & "Where " & .WherePrimary
        If vbNullString & .WhereSecondary <> vbNullString Then
            .SqlAll = .SqlAll & " and " & .WhereSecondary
        End If
        If vbNullString & .WhereTertiary <> vbNullString Then
            .SqlAll = .SqlAll & " and " & .WhereTertiary
        End If
        '.Sql hold values Select, From, WherePrimary, WhereSecondary, and WhereTertiary, only
        .SQL = .SqlAll
        
        If vbNullString & .WhereDates <> vbNullString Then
            .SqlAll = .SqlAll & " and " & .WhereDates
        End If
        
        'Retain the filter used on the refresh
        Result = .filter
        
        If vbNullString & .filter <> vbNullString Then
            .SqlAll = .SqlAll & " and " & .filter
        End If
        If vbNullString & .OrderBy <> vbNullString Then
            .SqlAll = .SqlAll & " ORDER BY " & .OrderBy
        End If
        myForm.SubForm.Form.RecordSource = .SqlAll
    End With

BuildExit:
On Error Resume Next
    BuildDetail = Result
Exit Function
BuildError:
    Result = vbNullString
    If Err.Number = 3075 Then
        MsgBox "There is a syntax error in the SQL Statement used to build the report." & vbCrLf & _
            "The most likely cause of the error is a syntax error in the Filter." & vbCrLf & vbCrLf & _
            "Error populating Detail!", vbCritical, "SQL Error"
    Else
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error populating Detail!", vbCritical, "SQL ERROR"
    End If
    Resume BuildExit
    Resume
End Function

Public Sub PopulateScreenFilters(ByRef MyControl As Control, ByVal ScreenID As Long)
On Error GoTo PopulateListError
    Dim SQL As String
    SQL = "SELECT FilterId, FilterName, UserName " & _
          "FROM SCR_ScreensFilters " & _
          "WHERE ScreenID = " & ScreenID & " "
        
    If TogValue Then
        myForm.cmdToggle.Caption = "Mine"
        ' HC 6/2010 - changed filter for mine to display those items matching the user name and those w/o user name
        ' DS 11/15/11 - changed from UserName Is NULL to Nz(UserName, '') = '' to display imports screen filters with user name as empty string
        SQL = SQL & " AND (UserName ='" & Identity.UserName & "' or Nz(UserName, '') = '')"
    Else
        myForm.cmdToggle.Caption = "All"
    End If
    
    'PD,Oct 22 2011 - CR # 2574 fix - Added clause to prevent filters with blank criteria from showing up on the main screen
    'SQL = SQL & " AND FilterSQL <> '' ORDER BY FilterName;"
    'SA 10/1/2012 - Changed to IS NOT NULL for SQL Server
    SQL = SQL & " AND FilterSQL IS NOT NULL ORDER BY FilterName;"
    
    MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Error Loading Filter"
    Resume PopulateListExit
End Sub

Public Function GetCurrentSQLStringLength(ByRef mvSql As CnlyScreenSQL) As Integer
    GetCurrentSQLStringLength = Len(mvSql.Select) + Len(mvSql.From) + Len(mvSql.WherePrimary) + Len(mvSql.WhereSecondary) + Len(mvSql.WhereTertiary) + Len(mvSql.WhereDates) + Len(mvSql.OrderBy)
End Function

Public Sub SaveScreenFilter(ByRef MyControl As Control)
On Error GoTo SaveFilterError
    Dim sFilterName As String
    Dim tmpFilter As String
    Dim FilterID As Long
    Dim filterString As String
    Dim frm As Form_SCR_MainScreens
    Dim SQL As String
    Dim continue As Boolean
    Dim Response As VbMsgBoxResult
    
    Set frm = MyControl.Parent.Parent.Parent
    tmpFilter = frm.SubForm.Form.filter
    
    ' make sure there really is a filter to save
    If tmpFilter = vbNullString Then
        Exit Sub
    End If
        
    ' HC 11/8/2010 - 2010 changed to handler new format on For Name
    ' modified the replacement string for the form name to include the brackets; this seems to be a change in the way 2010 references object names.
    tmpFilter = Replace(tmpFilter, "[" & frm.SubForm.Form.Name & "].", vbNullString)
    filterString = Replace(tmpFilter, "'", "''")
    filterString = EscapeQuotes(filterString)
    filterString = Replace(filterString, Chr(34) & Chr(34), "'")
    
    sFilterName = vbNullString
    sFilterName = GetName("Please enter the filter name." & vbCrLf & vbCrLf, tmpFilter)
    If sFilterName = vbNullString Then
        Exit Sub
    End If
    
    continue = True
    While continue
        FilterID = Nz(DLookup("FilterId", "SCR_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
            "ScreenId = " & myForm.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34)), -1)
        If FilterID <> -1 Then
            Response = MsgBox("Filter: " & sFilterName & "  already exists!" & vbCrLf & vbCrLf & _
                "Do you want to overwrite it?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Replace Filter")
            If Response = vbCancel Then
                continue = False
            ElseIf Response = vbYes Then
                continue = False
                SQL = " UPDATE SCR_ScreensFilters SET filterName = " & Chr(34) & sFilterName & Chr(34) & _
                    ", filterSQL = " & Chr(34) & filterString & Chr(34) & _
                    " WHERE filterId = " & FilterID
                RunDAO SQL
                SQL = "DELETE FROM SCR_ScreensFiltersDetails WHERE FilterId = " & FilterID
                RunDAO SQL
                SQL = " INSERT INTO SCR_ScreensFiltersDetails(FilterId,Operator,SqlString) VALUES(" & FilterID & "," & _
                    Chr(34) & "CUSTOM" & Chr(34) & "," & Chr(34) & filterString & Chr(34) & ")"
                RunDAO SQL
            Else
                sFilterName = GetName("Please enter a different name for the filter." & vbCrLf & vbCrLf & _
                    tmpFilter & vbCrLf & vbCrLf & sFilterName & " is already taken!", tmpFilter)
                If sFilterName = vbNullString Then
                    continue = False
                End If
            End If
        Else
            continue = False
            SQL = " INSERT INTO SCR_ScreensFilters(ScreenID,FilterName,FilterSQL,UserName) " & _
                    " VALUES (" & myForm.ScreenID & "," & Chr(34) & sFilterName & Chr(34) & "," & _
                     Chr(34) & filterString & Chr(34) & "," & Chr(34) & Identity.UserName & Chr(34) & ")"
            RunDAO SQL
            FilterID = DLookup("FilterId", "SCR_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
                "ScreenId = " & myForm.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34))
            ' first remove all the items in the table for this filter
            SQL = "DELETE FROM SCR_ScreensFiltersDetails WHERE FilterId = " & FilterID
            RunDAO SQL
            SQL = " INSERT INTO SCR_ScreensFiltersDetails(FilterId,Operator,SqlString) VALUES(" & FilterID & "," & _
                Chr(34) & "CUSTOM" & Chr(34) & "," & Chr(34) & filterString & Chr(34) & ")"
            RunDAO SQL
        End If
    Wend

SaveFilterExit:
On Error Resume Next
    Set frm = Nothing
Exit Sub
SaveFilterError:
    If Err.Number = 3022 Then 'Duplicate Query String
        MsgBox "The same filter already exists under a different name!" & String(2, vbCrLf) & "Error saveing filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    Else
        MsgBox Err.Description & String(2, vbCrLf) & "Error saveing filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    End If
    Resume SaveFilterExit
End Sub

Public Function GetCustomCriteriaSource(ByRef mvConfig As CnlyScreenCfg) As String
    Select Case mvConfig.CustomCriteriaListBoxRecordSource
        Case mvConfig.PrimaryListBoxRecordSource
            GetCustomCriteriaSource = myForm.CmboPrimary
        Case mvConfig.SecondaryListBoxRecordSource
            GetCustomCriteriaSource = myForm.CmboSecondary
        Case mvConfig.TertiaryListBoxRecordSource
            GetCustomCriteriaSource = myForm.CmboTertiary
        Case Else
            GetCustomCriteriaSource = vbNullString
    End Select
End Function

Public Sub RestoreTitles()
    Dim labelTop As Integer
    
    With myForm
        labelTop = .lineSelection.top + .lineSelection.Height + 85
        If labelTop <> .lblFiltersLabel.top Then
            .lblFiltersLabel.top = labelTop
            .lblDateRangeLabel.top = labelTop
            .lblSortingLabel.top = labelTop
            .lblTotalsLabel.top = labelTop
            .lblLayoutLabel.top = labelTop
            .lblLayout2Label.top = labelTop - 30
            
            .lblLayoutLabel.Width = .CmdScreenSave.left - .lblLayoutLabel.left - 51
            .lblLayout2Label.Width = .lblLayoutLabel.Width
            
            .lblTools2Label.Width = 3262
            .lblToolsLabel.Width = 3262
    
            .imgFilters_Collapsed.top = labelTop
           
        End If
        .imgFilters_Collapsed.visible = True
    End With
End Sub

Public Function SaveOpenScreen(ByRef mvConfig As CnlyScreenCfg) As Boolean
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2708 Changed to boolean function so message isn't displayed each time a screen is saved
    ' ** Added Tertiary
    Dim db As DAO.Database
    Dim SQL As String
    Dim StPrimary As String
    Dim StSecondary As String
    Dim stTertiary As String
    Dim StSort As String
    Dim StFormat As String
    Dim StCalc As String
    Dim StFilter As String
    Dim stFilterArray() As String
    Dim X As Integer
    Dim genUtils As New CT_ClsGeneralUtilities
    
    DoCmd.Hourglass True
    DoCmd.SetWarnings False
    Set db = CurrentDb
    
    If mvConfig.PrimaryListBoxMulti Then                'PrimaryCriteria
        StPrimary = BuildMultiItemSQL(myForm.CmboPrimaryMulti.Object, mvConfig.PrimaryQualifier)
    Else
        StPrimary = mvConfig.PrimaryQualifier & myForm.CmboPrimary & mvConfig.PrimaryQualifier
    End If
    
    If Nz(Replace(StPrimary, Chr(34), vbNullString), vbNullString) <> vbNullString Then
        ' Delete Saved Configuration
        'SA 10/1/2012 - Added CreakePK for SQL Server compatibility
        genUtils.CreatePK "SCR_SaveScreens", "ScreenID,UserName"
        SQL = "DELETE FROM SCR_SaveScreens WHERE ScreenID=" & mvConfig.ScreenID & _
              " AND UserName = '" & Identity.UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        ' Getting all Screen values and Inserting into Save Table
        'SA 1/30/12 - Added field ResetTabFilter at the end
        SQL = "INSERT INTO SCR_SaveScreens(ScreenID, UserName, CreatedDte, PrimaryListBy, PrimaryCriteria, " & _
            "SecondaryListBy, SecondaryCriteria, TertiaryListBy,TertiaryCriteria, DateFilter, StartDte, EndDte, " & _
             "Sort, Function, Report, Filter, Layout, Totals, Tab, ConditionalFormats, CustomCalculations, " & _
             "PowerBar, MainGridRecPos, TabGridRecPos, MainGridAdditionalFilter,MainGridAdditionalSort,HighlightCurrentRow,ResetTabFilter)"
        SQL = SQL & " VALUES(" & mvConfig.ScreenID & ","           'ScreenID
        SQL = SQL & Chr(34) & Identity.UserName & Chr(34) & ","    'UserName
        SQL = SQL & "#" & Now() & "#,"                             'CurrentDate
        SQL = SQL & Chr(34) & myForm.CmboListPrimaryBy & Chr(34) & ","             'PrimaryListBy
        
        If mvConfig.PrimaryListBoxMulti Then                'PrimaryCriteria
            StPrimary = BuildMultiItemSQL(myForm.CmboPrimaryMulti.Object, mvConfig.PrimaryQualifier)
        Else
            StPrimary = mvConfig.PrimaryQualifier & myForm.CmboPrimary & mvConfig.PrimaryQualifier
        End If
        StPrimary = Replace(Replace(Replace(StPrimary, Chr(34), vbNullString), ")", vbNullString), "(", vbNullString)
        SQL = SQL & Chr(34) & StPrimary & Chr(34) & ","
        
        SQL = SQL & Chr(34) & myForm.CmboListSecondaryBy & Chr(34) & ","  'SecondaryListBy
        
        If mvConfig.SecondaryListBoxMulti Then              'SecondaryCriteria
            StSecondary = BuildMultiItemSQL(myForm.CmboSecondaryMulti.Object, mvConfig.SecondaryQualifier)
        Else
            StSecondary = mvConfig.SecondaryQualifier & myForm.CmboSecondary & mvConfig.SecondaryQualifier
        End If
        StSecondary = Replace(Replace(Replace(StSecondary, Chr(34), vbNullString), ")", vbNullString), "(", vbNullString)
        SQL = SQL & Chr(34) & StSecondary & Chr(34) & ", "
        
        SQL = SQL & Chr(34) & myForm.CmboListTertiaryBy & Chr(34) & ","   'TertiaryListBy
        If mvConfig.TertiaryListBoxMulti Then              'TertiaryCriteria
            stTertiary = BuildMultiItemSQL(myForm.CmboTertiaryMulti.Object, mvConfig.TertiaryQualifier)
        Else
            stTertiary = mvConfig.TertiaryQualifier & myForm.CmboTertiary & mvConfig.TertiaryQualifier
        End If
        stTertiary = Replace(Replace(Replace(stTertiary, Chr(34), vbNullString), ")", vbNullString), "(", vbNullString)
        SQL = SQL & Chr(34) & stTertiary & Chr(34) & ","
        
        SQL = SQL & Chr(34) & myForm.CmboFilterDte.Value & Chr(34) & ","   'DateFilterField
        SQL = SQL & Chr(34) & myForm.StartDte.Value & Chr(34) & "," 'StartDate
        SQL = SQL & Chr(34) & myForm.EndDte.Value & Chr(34) & ","    'EndDate
    
        If myForm.SortList.ListCount > 0 Then                   'SortList
            For X = 0 To myForm.SortList.ListCount - 1
                If X > 0 Then StSort = StSort & ","
                StSort = StSort & myForm.SortList.Column(1, X) & "," & myForm.SortList.Column(0, X)
            Next X
            SQL = SQL & Chr(34) & StSort & Chr(34) & ","
        Else
            SQL = SQL & "'',"
        End If
        
        SQL = SQL & Chr(34) & myForm.CmboFunction.Value & Chr(34) & ","     'Function
        SQL = SQL & Nz(myForm.CmboReports.Value, 0) & ","      'Report
        StFilter = vbNullString
        If Nz(myForm.cmboFiltersSelected.RowSource, vbNullString) <> vbNullString Then
            ' add the split sequence to the end so it will split correctly
            stFilterArray = Split(myForm.cmboFiltersSelected.RowSource & ";", ";")
            For X = 0 To UBound(stFilterArray) - 1 Step 2
                StFilter = StFilter & stFilterArray(X) & ";" & Mid$(stFilterArray(X + 1), 2, Len(stFilterArray(X + 1)) - 2) & ";"
            Next X
        End If
        If StFilter <> vbNullString Then
            StFilter = Mid(StFilter, 1, Len(StFilter) - 1)
        End If
        
        SQL = SQL & Chr(34) & StFilter & Chr(34) & ","
        SQL = SQL & Nz(myForm.CmboLayouts.Value, 0) & ","      'Layout
        SQL = SQL & Nz(myForm.CmboTotals.Value, 0) & ","       'Totals
        SQL = SQL & Nz(myForm.Tabs.Value, 0) & ","             'Tab
        
        If myForm.SubformCondFormats.Form!LstApply.ListCount > 0 Then   'ConditionalFormat
            For X = 0 To myForm.SubformCondFormats.Form!LstApply.ListCount - 1
                If X > 0 Then StFormat = StFormat & ","
                StFormat = StFormat & myForm.SubformCondFormats.Form!LstApply.Column(1, X)
                Next X
            SQL = SQL & Chr(34) & StFormat & Chr(34) & ","
        Else
            SQL = SQL & "'',"
        End If
        
        If myForm.SubformCalcs.Form!LstApply.ListCount > 0 Then         'CustomCalcuations
            For X = 0 To myForm.SubformCalcs.Form!LstApply.ListCount - 1
                If X > 0 Then StCalc = StCalc & ","
                StCalc = StCalc & myForm.SubformCalcs.Form!LstApply.Column(1, X)
                Next X
            SQL = SQL & Chr(34) & StCalc & Chr(34) & ","
        Else
            SQL = SQL & "'',"
        End If
        
        SQL = SQL & Chr(34) & myForm.SubformPowerBar.SourceObject & Chr(34) & ","    'PowerBar
        
        If myForm.SubForm.Form.RecordSource <> vbNullString Then
            SQL = SQL & Nz(myForm.SubForm.Form.CurrentRecord, 1) & ","          'Current Grid Record
        Else
            SQL = SQL & "1, "
        End If
        
        Select Case myForm.Tabs.Value   'Current Tab Record
            Case Is < 0
                SQL = SQL & "0 ,"
            Case 0
                If Nz(myForm.Subform_2.Form.RecordSource, vbNullString) <> vbNullString Then
                    SQL = SQL & myForm.Subform_2.Form.CurrentRecord & ","
                Else
                    SQL = SQL & "0,"
                End If
            Case Else
                If Nz(myForm.Controls.Item("Subform" & CStr(myForm.Tabs.Value + 1)).Form.RecordSource, vbNullString) <> vbNullString Then
                    SQL = SQL & Nz(myForm.Controls.Item("Subform" & CStr(myForm.Tabs.Value + 1)).Form.CurrentRecord, 1) & ","
                Else
                    SQL = SQL & "0,"
                End If
        End Select
        
        'SA 4/20/2012 - Hotfix: Changed to use single quotes so that filters with double quotes are saved
        SQL = SQL & "'" & Replace(Nz(myForm.SubForm.Form.filter, vbNullString), "'", "''") & "',"          'MainGridAdditionalFilter

        SQL = SQL & Chr(34) & Nz(myForm.SubForm.Form.OrderBy, vbNullString) & Chr(34) & ","                 'MainGridAdditionalSort
        'SA 1/30/2012 - Added chkResetTabFilter
        SQL = SQL & myForm.chkHighlight & "," & myForm.chkResetTabFilter & ")"

        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        DoCmd.SetWarnings True
        DoCmd.Hourglass False
    End If

    SaveOpenScreen = True
AllDone:
On Error Resume Next
    Set db = Nothing
    Set genUtils = Nothing
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
Exit Function
ErrorHappened:
    SaveOpenScreen = False
    MsgBox Err.Description, vbCritical
    Resume AllDone
    Resume
End Function

Public Sub BuildTotalsCustom(ByVal forceRebuild As Boolean)
On Error GoTo ErrorHappened
    'SA 2012-05-31 - Modified to load Tab0 with source object only when used
    Dim strTotalsSQL As CnlyScreenSQLTotalsCustom
    Dim SQL As String
    Dim frm As Form_CT_SubGenericDataSheet100
    
    If myForm.FormFooter.visible = False Then GoTo ExitNow
      
    If myForm.CmboTotals.ListIndex <> -1 Then
        'Set source object on first run and load into frm
        With myForm.Tabs.Pages(glTabTotalsCustom).Controls(0)
            If .SourceObject = vbNullString Then
                .SourceObject = "CT_SubGenericDataSheet100"
            End If
            Set frm = .Form
        End With
        
        strTotalsSQL = SCR_TotalsCustom.BuildAll(myForm.CmboTotals, myForm)
        SQL = SCR_TotalsCustom.ToSQL(strTotalsSQL)
        
        If (vbNullString & frm.Tag <> myForm.CmboTotals) Or (forceRebuild) Then  ' THE ACTIVE TOTALID IS STORED ON THE FORM TAG
            frm.IsCustomTotal = True
            frm.Tag = myForm.CmboTotals '- Must come before init so that formats can be looked up
            frm.InitData SQL, 3 '- Recordset (SLOW) - may need to be replace in the future
            frm.visible = True
        End If
        
        myForm.Tabs.Pages(glTabTotalsCustom).Controls(0).Tag = SQL
        
        If myForm.Tabs.Value = glTabTotalsCustom Then
            frm.RecordSource = SQL
        Else
            frm.RecordSource = vbNullString
            myForm.Tabs.Pages(glTabTotalsCustom).Controls(0).Tag = SQL 'THIS IS WHERE SQL IS STORED FOR INNACTIVE TABS
        End If
    
        If Not myForm.Tabs.Pages(glTabTotalsCustom).visible Then
           myForm.Tabs.Pages(glTabTotalsCustom).visible = True
        End If
    Else
        If myForm.Tabs.Pages(glTabTotalsCustom).visible Then 'IF IT IS VISIBLE THEN CLEAR - INFO AND HIDE
            myForm.Tabs.Pages(glTabTotalsCustom).visible = False
            myForm.Tabs.Pages(glTabTotalsCustom).Controls(0).Tag = vbNullString
            
            'Clear frm settings if used
            With myForm.Tabs.Pages(glTabTotalsCustom).Controls(0)
                If .SourceObject <> vbNullString Then
                    Set frm = .Form
                    
                    frm.Tag = vbNullString
                    frm.RecordSource = vbNullString
                End If
            End With
        End If
    End If

ExitNow:
On Error Resume Next
    Set frm = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, "Error Applying Custom Totals : BuildTotalsCustom"
    Resume ExitNow
    Resume
End Sub

Public Sub ReleaseFooterForResize()
On Error GoTo ReleaseError
    Dim pgCt As Integer
    Dim pgIDx As Integer
    Dim pg As Page
    Dim pgCtrlCt As Integer
    Dim pgCtrlIdx As Integer
    Dim pgCtrl As Control
    Dim TabHeight As Integer
    
    myForm.Splitter.top = 0
    myForm.Tabs.top = myForm.Splitter.top + myForm.Splitter.Height
    myForm.Tabs.Height = myForm.FormFooter.Height - myForm.Tabs.top
    
    TabHeight = myForm.Tabs.Height - 500
    pgCt = myForm.Tabs.Pages.Count - 1

    Application.Echo False
    For pgIDx = 0 To pgCt
        Set pg = myForm.Tabs.Pages.Item(pgIDx)
        pg.Height = TabHeight
        pgCtrlCt = pg.Controls.Count - 1
        For pgCtrlIdx = 0 To pgCtrlCt
            Set pgCtrl = pg.Controls(pgCtrlIdx)
            
            With pgCtrl
                If pg.Height - 500 > 0 Then
                    .Height = pg.Height - 500
                End If
                .Width = pg.Width
                .visible = True
            End With
        Next
    Next
    
    myForm.Tabs.visible = True
    myForm.SubForm.visible = True

    myForm.Resize
ReleaseExit:
Exit Sub
    Application.Echo True
ReleaseError:
    MsgBox (Err.Number & ": " & Err.Description)
    Resume ReleaseExit
End Sub

Public Function RunDAO(ByVal SQL As String) As Boolean
'Execute SQL
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim db As DAO.Database
    
    Set db = CurrentDb
    db.Execute SQL, dbFailOnError
    
    Result = True
ExitNow:
On Error Resume Next
    Set db = Nothing
    RunDAO = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function