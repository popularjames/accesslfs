Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrUserProfile As String
Private mbRecordChanged As Boolean
Private miAppPermission As Integer
Private mbLocked As Boolean
Private mReturnDate As Date
Private ColReSize As clsAutoSizeColumns
Private mstrStoredProcName As String

Private mstrReportNumber As Integer
Private mstrReportName As String

Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Const CstrFrmAppID As String = "QueueProductivity"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Let StoredProcName(data As String)
    mstrStoredProcName = data
End Property

Public Property Get StoredProcName() As String
    StoredProcName = mstrStoredProcName
End Property



Public Property Let ReportName(data As String)
    mstrReportName = data
End Property

Public Property Get ReportName() As String
    ReportName = mstrReportName
End Property



Public Property Let ReportNumber(data As Integer)
    mstrReportNumber = data
End Property

Public Property Get ReportNumber() As Integer
    ReportNumber = mstrReportNumber
End Property

Private Sub cmdClearFilter_Click()
Me.txtFilter = ""

End Sub

Private Sub cmdExportToExcel_Click()
    Dim bExport As Boolean
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    If lstDetail.ListCount = 1 Then Exit Sub     'only row headers, nothing to do
    
    bExport = ExportDetails(Me.lstDetail.RecordSet, "Queue_Productivity.xls")
    
    'If bExport = False Then
    '     MsgBox "An error was encountered while attempting to export Detail data to Excel.", vbCritical
    'End If
End Sub


Private Function ExportDetails(rst As ADODB.RecordSet, strFilePath As String) As Boolean
'Function to export recordset to excel file

    Dim dlg As clsDialogs
    Dim cie As clsImportExport

    Set cie = New clsImportExport
    Set dlg = New clsDialogs

    
    If rst Is Nothing Then Exit Function     ' nothing to save
    
    With dlg
    
        strFilePath = .SavePath(Identity.CurrentFolder, xlsf, strFilePath)
        strFilePath = .CleanFileName(strFilePath, CleanPath)
     
        If strFilePath <> "" Then
        
            If .FileExists(strFilePath) = True Then
            
                If MsgBox("Overwrite existing file?", vbYesNo) = vbYes Then
                    .DeleteFile strFilePath
                Else
                    GoTo exitHere
                End If
            
            End If
            
            If rst.recordCount > 65535 Then
                MsgBox "Warning: Your recordset contains more than 65535 rows, the maximum number of rows allowed in Excel.  " & _
                Trim(str(rst.recordCount - 65535)) & " rows will not be displayed.", vbCritical
            End If
                        
        Else
            GoTo exitHere
        End If
     
        With cie
            .ExportExcelRecordset rst, strFilePath, True
        End With
         
    End With
    
    ExportDetails = True

exitHere:
    Set cie = Nothing
    Set dlg = Nothing
    Exit Function
    
HandleError:
    ExportDetails = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function



Public Sub RefreshData()
    On Error GoTo ErrHandler
    
    'thieu code begin
    Dim cmd As ADODB.Command
    'thieu code end

    'Dim rst As ADODB.Recordset
    Dim iResult As Integer
    Dim strRPTTableName As String
    Dim strSQL As String
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strRPTTableName = Replace(Identity.UserName(), ".", "_") & "_Adhoc_Report"
    myCode_ADO.SQLTextType = sqltext
    'myCode_ADO.sqlString = "exec " & Me.txtStoredProc & " '" & Identity.UserName() & "', '" & Format(Me.txtSQLFromDt, "mm-dd-yyyy") & "','" & Format(txtSQLThruDt, "mm-dd-yyyy") & "'"
    'Changing the reports to be triggered by a common stored proc run as DBO
    myCode_ADO.sqlString = "exec [dbo].[usp_Run_UniversalReport] '" & Me.txtStoredProc & "', '" & Identity.UserName() & "', '" & Format(Me.txtSQLFromDt, "mm-dd-yyyy") & "','" & Format(txtSQLThruDt, "mm-dd-yyyy") & "'"

    'MsgBox myCode_ADO.SQLstring

    'iResult = myCode_ADO.Execute    'thieu commented out
    
    'thieu code begin
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandTimeout = 0
    cmd.commandType = adCmdText
    cmd.CommandText = myCode_ADO.sqlString
    cmd.Execute
    'thieu code end
    
    
    
  'Alex commented this out on 11/18/09.  This is what populates the old list box
    'strSQL = "select * from " & strRPTTableName
    'Set rst = myCode_ADO.OpenRecordSet(strSQL)
    
    ''Setting the list record set equal to the specic ADO-class record set
    'lstDetail.ColumnCount = rst.Fields.Count
    'Set lstDetail.Recordset = rst

    'Set ColReSize = New clsAutoSizeColumns
    'ColReSize.SetControl Me.lstDetail
    ''don't resize if lstclaims is null
    'On Error Resume Next
    'If Me.lstDetail.ListCount - 1 > 0 Then
    '    ColReSize.AutoSize
    'End If
    'Set ColReSize = Nothing

   'DoCmd.OpenReport Me.ReportName, acViewPreview

   
   
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Sub

Private Sub cmdRunReport_Click()
On Error GoTo ErrorOnRunReport
    
    DoCmd.Hourglass True
    
    'Dim db As Database
    'Set db = CurrentDb
    
    'Dim qry1 As QueryDef
    'Dim sSQL1 As String
    
    'Get rid of Excel option
    'If Me.optOutput = 3 Then
    '    MsgBox "Working on Excel output.  In the meantime, run it by table, then do File->Export and save as type Microsoft Excel 97-2002."
    '    Me.optOutput = 2
    'End If
    
    If Trim(Me.txtAccessReportName) & "" = "" And Trim(Me.txtOutputTable) & "" = "" Then
        MsgBox "Report in development."
        DoCmd.Hourglass False
        Exit Sub
    End If
    
    If Me.optOutput = 1 And Trim(Me.txtAccessReportName) & "" = "" Then
        MsgBox "No access report for this one."
        DoCmd.Hourglass False
        Me.optOutput = 2
    End If
    
    If Me.optOutput = 2 And Trim(Me.txtOutputTable) & "" = "" Then
        MsgBox "No table for this one."
        DoCmd.Hourglass False
        Me.optOutput = 1
    End If
    
    If Me.optOutput = 3 And Trim(Me.txtAccessFormName) & "" = "" Then
        MsgBox "No table for this one."
        DoCmd.Hourglass False
        Me.optOutput = 2
    End If
    
    'log user and report into CONCEPT_RunLog when using "last run"
    If Me.optData = 1 Then 'if just opening table
    
        Dim MyAdo As clsADO
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("REPORT_RunLog")
        MyAdo.sqlString = "INSERT INTO CMS_AUDITORS_CLAIMS.dbo.REPORT_RunLog (ReportID, LastRunDt, LastRunBy, RecordCount, LastRunDuration, LastRunDurationText) VALUES ('LR-" & Me.txtReportNumber & "', '" & Now & "', '" & Me.txtUserId & "', 0, 0, '0 Seconds')"
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute
    
    End If
    
    If Me.optData = 1 And Me.optOutput = 1 Then
       LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
       DoCmd.OpenReport Me.txtAccessReportName, acViewPreview, , Nz(Me.txtFilter, "")
       'DoCmd.Maximize
        SeudoMaximize
    End If
    
    'This is going to be the new standard:
    'what used to be open table (with or without query) now will be opening a form using ADO for the recordsource
    
    'open table
    If Me.optData = 1 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" Then
        '' 20121211 KD : Back to DAO for now due to 2010's issue with filter by form when form is bound to ADO RS
        LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable

        Call ReportingAccessForm("frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, ""))
        'If ReportingAccessFormADO("frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, "")) Then
            SeudoMaximize
        'End If
        'DoCmd.Maximize
    End If
    
    'refreshdata and open table
    If Me.optData = 2 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" Then
        If txtSQLFromDt & "" = "" Then
            MsgBox "Please enter a from date"
        ElseIf txtSQLThruDt & "" = "" Then
            MsgBox "Please enter a through date"
        Else
            RefreshData
        End If
        
        'Then repopulate parameters
        'ClearReportParameters
        'PopulateReportParameters
        Me.optData = 1
        Me.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
        Me.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), "") & ")"
        '' 20120214 JS : Since it was moved back to DAO we need to relink the table when refreshing the report
        LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
        '' 20121211 KD : Back to DAO for now due to 2010's issue with filter by form when form is bound to ADO RS
        Call ReportingAccessForm("frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, ""))
'        If ReportingAccessFormADO("frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, "")) Then
            SeudoMaximize
'        End If
        'DoCmd.Maximize
    End If
    
    
    'Use last run and show table if QueryOrderBy field is empty
    'If Me.optData = 1 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" And Nz(Me.txtQueryOrderBy, "") = "" Then
    '    LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
    '
    '    DoCmd.OpenTable Me.txtOutputTable
    '
    '    DoCmd.Maximize
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    'End If
    
    'refreshdata and open table if QueryOrderBy is empty
    'If Me.optData = 2 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" And Nz(Me.txtQueryOrderBy, "") = "" Then
    '   'Run the SQL, prep the parameters, flip button to old data, change last run dt, launch report no filter
    '    If txtSQLFromDt & "" = "" Then
    '        MsgBox "Please enter a from date"
    '    ElseIf txtSQLThruDt & "" = "" Then
    '        MsgBox "Please enter a through date"
    '    Else
    '        RefreshData
    '    End If
    '
    '    'Then repopulate parameters
    '    'ClearReportParameters
    '    'PopulateReportParameters
    '    Me.optData = 1
    '    Me.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
    '    Me.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), "") & ")"
    '    LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
    '    DoCmd.OpenTable Me.txtOutputTable
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    'End If
    '
    
    ''use last run and create new query that can be sorted using QueryOrderBy
    'If Me.optData = 1 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" And Not Nz(Me.txtQueryOrderBy, "") = "" Then
    '
    '
    '    On Error Resume Next
    '    db.QueryDefs.Delete "qry_TMP_" & Me.txtOutputTable
    '    On Error GoTo 0
    '
    '    sSQL1 = "SELECT * from " & Me.txtOutputTable & " order by " & txtQueryOrderBy
    '
    '    Set qry1 = db.CreateQueryDef("qry_TMP_" & Me.txtOutputTable, sSQL1)
    '
    '    DoCmd.OpenQuery qry1.Name
    '
    '    DoCmd.Maximize
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    'End If
    '
    ''refreshdata and create a new query that can be sorted with QueryOrderBy
    'If Me.optData = 2 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "r" And Not Nz(Me.txtQueryOrderBy, "") = "" Then
    '   'Run the SQL, prep the parameters, flip button to old data, change last run dt, launch report no filter
    '    If txtSQLFromDt & "" = "" Then
    '        MsgBox "Please enter a from date"
    '    ElseIf txtSQLThruDt & "" = "" Then
    '        MsgBox "Please enter a through date"
    '    Else
    '        RefreshData
    '    End If
    '
    '    'Then repopulate parameters
    '    'ClearReportParameters
    '    'PopulateReportParameters
    '    Me.optData = 1
    '    Me.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
    '    Me.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), "") & ")"
    '    LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
    '
    '
    '    On Error Resume Next
    '    db.QueryDefs.Delete "qry_TMP_" & Me.txtOutputTable
    '    On Error GoTo 0
    '
    '    sSQL1 = "SELECT * from " & Me.txtOutputTable & " order by " & txtQueryOrderBy
    '
    '    Set qry1 = db.CreateQueryDef("qry_TMP_" & Me.txtOutputTable, sSQL1)
    '
    '    DoCmd.OpenQuery qry1.Name
    '
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    'End If
    
    
    
    'use last run and show query
    If Me.optData = 1 And Me.optOutput = 2 And left(Me.txtOutputTable, 1) = "q" Then
        DoCmd.OpenQuery Me.txtOutputTable
        'DoCmd.Maximize
        SeudoMaximize
        If Trim(Me.txtFilter) <> "" Then
            DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
        End If
    End If
    
    
    'use last run and open form. this is for reports with over 10k records
    If Me.optData = 1 And Me.optOutput = 3 Then
        
        LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
        If Me.txtAccessFormName = "frm_RPT_AccessForm" Then
            ReportingAccessForm "frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, "")
        Else
            DoCmd.OpenForm Me.txtAccessFormName, acFormDS, , Nz(Me.txtFilter, "")
        End If
        'DoCmd.Maximize
        SeudoMaximize
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    End If
    
    'refresh data and open form
    If Me.optData = 2 And Me.optOutput = 3 Then
       'Run the SQL, prep the parameters, flip button to old data, change last run dt, launch report no filter
        If txtSQLFromDt & "" = "" Then
            MsgBox "Please enter a from date"
        ElseIf txtSQLThruDt & "" = "" Then
            MsgBox "Please enter a through date"
        Else
            RefreshData
        End If
        
        'need to link table since now (with ADO) this is not done anywhere else, this is for reports with over 100,000 records (usually details)  JS 09/10/2012

         LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable 'VS 3/3/2015  for Dev Version
        'Then repopulate parameters
        'ClearReportParameters
        'PopulateReportParameters
        Me.optData = 1
        Me.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
        Me.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), "") & ")"

        LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable 'VS 3/3/2015  for Dev Version
        If Me.txtAccessFormName = "frm_RPT_AccessForm" Then
            ReportingAccessForm "frm_RPT_AccessForm", Me.txtOutputTable, Nz(Me.txtFilter, "1=1"), Nz(txtQueryOrderBy, "")
        Else
            DoCmd.OpenForm Me.txtAccessFormName, acFormDS, , Nz(Me.txtFilter, "")
        End If
        'DoCmd.Maximize
        SeudoMaximize
    '    If Trim(Me.txtFilter) <> "" Then
    '        DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
    '    End If
    End If
    
    'Get rid of Excel option
    'If Me.optData = 1 And Me.optOutput = 3 Then
    '
    '    'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, Me.txtOutputTable, "ORAC2.xls", True
    '
    '
    '    Dim myADO As clsADO
    '    Dim rs As ADODB.Recordset
    '    Set myADO = New clsADO
    '    myADO.ConnectionString = GetConnectString(Me.txtOutputTable)
    '    myADO.SQLstring = "select * from " & Me.txtOutputTable
    '
    '    Set rs = myADO.OpenRecordSet
    '
    '    Dim bExport As Boolean
    '
    '    MsgBox rs.RecordCount
    '
    '    If rs Is Nothing Then Exit Sub     'nothing to do
    '    If rs.RecordCount = 1 Then Exit Sub     'only row headers, nothing to do
        
    '    bExport = ExportDetails(rs, "orca.xls")
    '
    '    Set rs = Nothing
    '    Set myADO = Nothing
        
    'End If
    
    'refresh data and open report
    If Me.optData = 2 And Me.optOutput = 1 Then
       'Run the SQL, prep the parameters, flip button to old data, change last run dt, launch report no filter
        If txtSQLFromDt & "" = "" Then
            MsgBox "Please enter a from date"
        ElseIf txtSQLThruDt & "" = "" Then
            MsgBox "Please enter a through date"
        Else
            RefreshData
        End If
        
        'Then repopulate parameters
        'ClearReportParameters
        'PopulateReportParameters
        Me.optData = 1
        Me.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
        Me.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & Me.txtReportNumber & Chr(34)), "") & ")"
        'LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable
        DoCmd.OpenReport Me.txtAccessReportName, acViewPreview, , Nz(Me.txtFilter, "")
    End If
    
'*******************************************************************
' THIS IS THE NEW PART FOR THE EXCEL EXPORT
'*******************************************************************

'MsgBox Me.optData
'MsgBox Me.optOutput
    
'use last run and open form. this is for reports with over 10k records
If Me.optData = 1 And Me.optOutput = 4 Then

    LinkTable "SQL", CurrentCMSServer(), "CMS_AUDITORS_Reports", Me.txtOutputTable 'VS 3/3/2015  for Dev Version
    'Get drop location of Excel file
    Dim intChoice As Integer
    Dim strPath As String
    Application.FileDialog(msoFileDialogSaveAs).Title = "Save your spreadsheet with the extension .xlsx"
    intChoice = Application.FileDialog(msoFileDialogSaveAs).show
    
    
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogSaveAs).SelectedItems(1)
        'MsgBox "This is the path the Excel spreadsheet will be placed: " & strPath
    ElseIf intChoice = 0 Then
        DoCmd.Hourglass False
        Exit Sub
    End If

    'Delete temp query def if it already exists
    Dim qdf As DAO.QueryDef
        For Each qdf In CurrentDb.QueryDefs
            If qdf.Name = "xxxExcelExportxxx" Then
                CurrentDb.QueryDefs.Delete "xxxExcelExportxxx"
                Exit For
            End If
        Next
        
    'Put the records you want to export into a temporary query
    Set qdf = CurrentDb.CreateQueryDef("xxxExcelExportxxx")
    'MsgBox "This is the where clause: " & qdf.SQL
    'MsgBox "this is the filter: [" & Me.txtFilter & "]"
    Dim FormFilter As String
    FormFilter = Trim(Me.txtFilter)
    If FormFilter = "" Then
        FormFilter = " 1 = 1"
    End If
    
    'MsgBox "this is the clean filter: [" & FormFilter & "]"
    
    qdf.SQL = "SELECT * FROM " & Me.txtOutputTable & " WHERE " & FormFilter
    'MsgBox "This is the query that is going to get exported: " & qdf.SQL

    
    'Transfer records in the temporary query to Excel
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "xxxExcelExportxxx", strPath, True
    CurrentDb.QueryDefs.Delete "xxxExcelExportxxx"
    MsgBox "Export complete!" & Chr(10) & Chr(10) & "PATH: " & strPath & Chr(10) & Chr(10) & "SQL: " & qdf.SQL
    
    Set qdf = Nothing
    'Me.optOutput = 2
End If
    
If Me.optData = 2 And Me.optOutput = 4 Then
    MsgBox "Please refresh data by selecting 'Table' first then select 'Use Last Run' to export to excel."
End If
    
'*******************************************************************
' END OF NEW EXCEL EXPORT
'*******************************************************************
    
    
    DoCmd.Hourglass False

Exit Sub

ErrorOnRunReport:
    DoCmd.Hourglass False
    MsgBox "Error while running the report.  Please take a screen shot of this error and send to Alex." & Chr(10) & Chr(10) & "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Exit Sub

End Sub

Private Sub cmdRunSQL_Click()
'    If txtSQLFromDt & "" = "" Then
'        MsgBox "Please enter a from date"
'    ElseIf txtSQLThruDt & "" = "" Then
'        MsgBox "Please enter a through date"
'    Else
'        RefreshData
'    End If
    
'    Me.cmdViewReportNoFilter.SetFocus
'    Me.cmdRunSQL.Enabled = False
End Sub



Private Sub cmdViewData_Click()
DoCmd.OpenTable Me.txtOutputTable
If Trim(Me.txtFilter) <> "" Then
    DoCmd.ApplyFilter , Nz(Me.txtFilter, "1=1")
End If

End Sub

Private Sub cmdViewReport_Click()
DoCmd.OpenReport Me.txtAccessReportName, acViewPreview, , Me.txtFilter
End Sub

Public Sub ClearReportParameters()

Me.txtFilter = ""

Me.T1L.Caption = ""
Me.T2L.Caption = ""
Me.T3L.Caption = ""
Me.T4L.Caption = ""
Me.T5L.Caption = ""

Me.T1 = ""
Me.T2 = ""
Me.T3 = ""
Me.T4 = ""
Me.T5 = ""

Me.T1F = ""
Me.T2F = ""
Me.T3F = ""
Me.T4F = ""
Me.T5F = ""

Me.T1R = ""
Me.T2R = ""
Me.T3R = ""
Me.T4R = ""
Me.T5R = ""

Me.T1O = "="
Me.T2O = "="
Me.T3O = "="
Me.T4O = "="
Me.T5O = "="

Me.L1L.Caption = ""
Me.L2L.Caption = ""
Me.L3L.Caption = ""
Me.L4L.Caption = ""
Me.L5L.Caption = ""

Me.L1F = ""
Me.L2F = ""
Me.L3F = ""
Me.L4F = ""
Me.L5F = ""

Me.L1.RowSource = ""
Me.L2.RowSource = ""
Me.L3.RowSource = ""
Me.L4.RowSource = ""
Me.L5.RowSource = ""

Me.L1K = True
Me.L2K = True
Me.L3K = True
Me.L4K = True
Me.L5K = True

End Sub
Public Sub PopulateReportParameters(Optional pControlID As String = "")

Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Dim ControlToSet As String
Set MyAdo = New clsADO

MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
MyAdo.sqlString = "select * from cms_auditors_claims.dbo.REPORT_Parameter where ReportID = '" & Me.txtReportNumber & "'"
Set rs = MyAdo.OpenRecordSet

If Not rs Is Nothing Then
If Not (rs.BOF Or rs.EOF) Then
    rs.MoveFirst
End If
While Not rs.EOF
    If pControlID = "" Then ControlToSet = rs("ControlID") Else ControlToSet = pControlID
    Select Case ControlToSet
        Case "T1"
            Me.T1L.Caption = Nz(rs("PromptText"), "")
            Me.T1F = Nz(rs("FieldName"), "")
            Me.T1 = Nz(rs("DefaultValue"), "")
            Me.T1R = Nz(rs("DefaultValueRange"), "")
        Case "L1"
            Me.L1L.Caption = Nz(rs("PromptText"), "")
            Me.L1F = Nz(rs("FieldName"), "")
            MyAdo.sqlString = Nz(rs("RecordSet"), "SELECT ListItem, ListItemDesc FROM cms_auditors_claims.dbo.REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L1""")
            Set Me.L1.RecordSet = MyAdo.OpenRecordSet
        Case "T2"
            Me.T2L.Caption = Nz(rs("PromptText"), "")
            Me.T2F = Nz(rs("FieldName"), "")
            Me.T2 = Nz(rs("DefaultValue"), "")
            Me.T2R = Nz(rs("DefaultValueRange"), "")
        Case "L2"
            Me.L2L.Caption = Nz(rs("PromptText"), "")
            Me.L2F = Nz(rs("FieldName"), "")
            MyAdo.sqlString = Nz(rs("Recordset"), "SELECT ListItem, ListItemDesc FROM cms_auditors_claims.dbo.REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L2""")
            Set Me.L2.RecordSet = MyAdo.OpenRecordSet
        Case "T3"
            Me.T3L.Caption = Nz(rs("PromptText"), "")
            Me.T3F = Nz(rs("FieldName"), "")
            Me.T3 = Nz(rs("DefaultValue"), "")
            Me.T3R = Nz(rs("DefaultValueRange"), "")
        Case "L3"
            Me.L3L.Caption = Nz(rs("PromptText"), "")
            Me.L3F = Nz(rs("FieldName"), "")
            MyAdo.sqlString = Nz(rs("RecordSet"), "SELECT ListItem, ListItemDesc FROM cms_auditors_claims.dbo.REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L3""")
            Set Me.L3.RecordSet = MyAdo.OpenRecordSet
    End Select
    If pControlID <> "" Then rs.MoveLast
    rs.MoveNext
Wend
End If

rs.Close
Set rs = Nothing

AddFieldstoTComboBoxesADO Me.txtOutputTable

'Me.T1L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'Me.T2L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'Me.T3L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'Me.T4L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T4"""))
'Me.T5L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T5"""))
'
'Me.T1F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'Me.T2F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'Me.T3F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'Me.T4F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T4"""))
'Me.T5F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T5"""))
'
'
'Me.L1L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L1"""))
'Me.L2L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L2"""))
'Me.L3L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L3"""))
'Me.L4L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L4"""))
'Me.L5L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L5"""))
'
'Me.L1F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L1"""))
'Me.L2F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L2"""))
'Me.L3F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L3"""))
'Me.L4F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L4"""))
'Me.L5F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L5"""))
'
'
'Me.L1.RowSource = Nz(DLookup("[RowSource]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L1"""))
'Me.L2.RowSource = Nz(DLookup("[RowSource]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L2"""))
'Me.L3.RowSource = Nz(DLookup("[RowSource]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L3"""))
'Me.L4.RowSource = Nz(DLookup("[RowSource]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L4"""))
'Me.L5.RowSource = Nz(DLookup("[RowSource]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""L5"""))
'
'If Me.L1.RowSource = "" Then
'  Me.L1.RowSource = "SELECT ListItem, ListItemDesc FROM REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L1"""
'End If
'
'If Me.L2.RowSource = "" Then
'  Me.L2.RowSource = "SELECT ListItem, ListItemDesc FROM REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L2"""
'End If
'
'If Me.L3.RowSource = "" Then
'  Me.L3.RowSource = "SELECT ListItem, ListItemDesc FROM REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L3"""
'End If
'
'If Me.L4.RowSource = "" Then
'  Me.L4.RowSource = "SELECT ListItem, ListItemDesc FROM REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L4"""
'End If
'
'If Me.L5.RowSource = "" Then
'  Me.L5.RowSource = "SELECT ListItem, ListItemDesc FROM REPORT_ListBox WHERE ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " and ControlID = ""L5"""
'End If
'
'Me.T1 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'Me.T2 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'Me.T3 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'Me.T4 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T4"""))
'Me.T5 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T5"""))
'
'Me.T1R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'Me.T2R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'Me.T3R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'Me.T4R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T4"""))
'Me.T5R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T5"""))

If Trim(Nz(Me.T1R, "")) + "" <> "" Then
  Me.T1O = "BETWEEN"
End If

If Trim(Nz(Me.T2R, "")) + "" <> "" Then
  Me.T2O = "BETWEEN"
End If

If Trim(Nz(Me.T3R, "")) + "" <> "" Then
  Me.T3O = "BETWEEN"
End If

'If Trim(Nz(Me.T4R, "")) + "" <> "" Then
'  Me.T4O = "BETWEEN"
'End If
'
'If Trim(Nz(Me.T5R, "")) + "" <> "" Then
'  Me.T5O = "BETWEEN"
'End If

Me.Refresh

End Sub

Private Sub cmdViewReportNoFilter_Click()
DoCmd.OpenReport Me.txtAccessReportName, acViewPreview
End Sub

Private Sub cmdPrepParameters_Click()
ClearReportParameters
PopulateReportParameters

End Sub

Private Sub cmdBuildFilter_Click()
    BuildReportFilter

End Sub

Private Sub BuildReportFilter()

Dim SQLFilter As String

Dim T1DataType As String
Dim T2DataType As String
Dim T3DataType As String
Dim T4DataType As String
Dim T5DataType As String
Dim L1DataType As String
Dim L2DataType As String
Dim L3DataType As String
Dim L4DataType As String
Dim L5DataType As String


Dim T1Range As Boolean
Dim T2Range As Boolean
Dim T3Range As Boolean
Dim T4Range As Boolean
Dim T5Range As Boolean

Dim T1Contain As String
Dim T2Contain As String
Dim T3Contain As String
Dim T4Contain As String
Dim T5Contain As String
Dim L1Contain As String
Dim L2Contain As String
Dim L3Contain As String
Dim L4Contain As String
Dim L5Contain As String

'T1DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'T2DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'T3DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))

T1DataType = Nz(T1F.Column(1, T1F.ListIndex), "")
T2DataType = Nz(T2F.Column(1, T2F.ListIndex), "")
T3DataType = Nz(T3F.Column(1, T3F.ListIndex), "")


'T4DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""T4"""))
'T5DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""T5"""))


Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Set MyAdo = New clsADO

MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
MyAdo.sqlString = "select * from cms_auditors_claims.dbo.REPORT_Parameter where ReportID = '" & Me.txtReportNumber & "'"
Set rs = MyAdo.OpenRecordSet

If Not rs Is Nothing Then
If Not (rs.BOF Or rs.EOF) Then rs.MoveFirst
While Not rs.EOF
    Select Case rs("ControlID")
        Case "L1"
            L1DataType = Nz(rs("ParamDataType"), "")
        Case "L2"
            L2DataType = Nz(rs("ParamDataType"), "")
        Case "L3"
            L3DataType = Nz(rs("ParamDataType"), "")
    End Select
    rs.MoveNext
Wend
End If
rs.Close
Set rs = Nothing

'L1DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""L1"""))
'L2DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""L2"""))
'L3DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""L3"""))


'L4DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""L4"""))
'L5DataType = Nz(DLookup("[ParamDataType]", "REPORT_Parameter", "ReportID = " & Chr(39) & Me.txtReportNumber & Chr(39) & " AND ControlID = ""L5"""))

'If T1DataType = "" And T1F.ListCount > 0 And Not IsNull(T2F) Then
'    Select Case CurrentDb.TableDefs(txtOutputTable).Fields(T1F).Type
'        Case 10
'            T1DataType = "T"
'        Case 2 To 7
'            T1DataType = "N"
'        Case 8
'            T1DataType = "D"
'        Case Else
'            T1DataType = ""
'    End Select
'End If

'If T2DataType = "" And T2F.ListCount > 0 And Not IsNull(T2F) Then
'    Select Case CurrentDb.TableDefs(txtOutputTable).Fields(T2F).Type
'        Case 10
'            T2DataType = "T"
'        Case 2 To 7
'            T2DataType = "N"
'        Case 8
'            T2DataType = "D"
'        Case Else
'            T2DataType = ""
'    End Select
'End If

'If T3DataType = "" And T3F.ListCount > 0 And Not IsNull(T3F) Then
'    Select Case CurrentDb.TableDefs(txtOutputTable).Fields(T3F).Type
'        Case 10
'            T3DataType = "T"
'        Case 2 To 7
'            T3DataType = "N"
'        Case 8
'            T3DataType = "D"
'        Case Else
'            T3DataType = ""
'    End Select
'End If

Select Case T1DataType
    Case Is = "D"
        T1Contain = "#"
    Case Is = "N"
        T1Contain = ""
    Case Else
        T1Contain = Chr(39)
End Select

Select Case T2DataType
    Case Is = "D"
        T2Contain = "#"
    Case Is = "N"
        T2Contain = ""
    Case Else
        T2Contain = Chr(39)
End Select

Select Case T3DataType
    Case Is = "D"
        T3Contain = "#"
    Case Is = "N"
        T3Contain = ""
    Case Else
        T3Contain = Chr(39)
End Select

Select Case T4DataType
    Case Is = "D"
        T4Contain = "#"
    Case Is = "N"
        T4Contain = ""
    Case Else
        T4Contain = Chr(39)
End Select

Select Case T5DataType
    Case Is = "D"
        T5Contain = "#"
    Case Is = "N"
        T5Contain = ""
    Case Else
        T5Contain = Chr(39)
End Select

Select Case L1DataType
    Case Is = "D"
        L1Contain = "#"
    Case Is = "N"
        L1Contain = ""
    Case Else
        L1Contain = Chr(39)
End Select

Select Case L2DataType
    Case Is = "D"
        L2Contain = "#"
    Case Is = "N"
        L2Contain = ""
    Case Else
        L2Contain = Chr(39)
End Select

Select Case L3DataType
    Case Is = "D"
        L3Contain = "#"
    Case Is = "N"
        L3Contain = ""
    Case Else
        L3Contain = Chr(39)
End Select

Select Case L4DataType
    Case Is = "D"
        L4Contain = "#"
    Case Is = "N"
        L4Contain = ""
    Case Else
        L4Contain = Chr(39)
End Select

Select Case L5DataType
    Case Is = "D"
        L5Contain = "#"
    Case Is = "N"
        L5Contain = ""
    Case Else
        L5Contain = Chr(39)
End Select


T1Range = False
T2Range = False
T3Range = False
T4Range = False
T5Range = False


If Trim(Nz(Me.T1R, "")) + "" <> "" Then
  T1Range = True
  Me.T1O = "BETWEEN"
End If

If Trim(Nz(Me.T2R, "")) + "" <> "" Then
  T2Range = True
  Me.T2O = "BETWEEN"
End If

If Trim(Nz(Me.T3R, "")) + "" <> "" Then
  T3Range = True
  Me.T3O = "BETWEEN"
End If

If Trim(Nz(Me.T4R, "")) + "" <> "" Then
  T4Range = True
  Me.T4O = "BETWEEN"
End If

If Trim(Nz(Me.T5R, "")) + "" <> "" Then
  T5Range = True
  Me.T5O = "BETWEEN"
End If


'If T1Range = True And T1DataType = "T" Then
'  MsgBox "Unfortunately, you can't use a range on a text field."
'End If

'If T2Range = True And T2DataType = "T" Then
'  MsgBox "Unfortunately, you can't use a range on a text field."
'End If

'If T3Range = True And T3DataType = "T" Then
'  MsgBox "Unfortunately, you can't use a range on a text field."
'End If

'If T4Range = True And T4DataType = "T" Then
'  MsgBox "Unfortunately, you can't use a range on a text field."
'End If

'If T5Range = True And T5DataType = "T" Then
'  MsgBox "Unfortunately, you can't use a range on a text field."
'End If



SQLFilter = "<start>"

Dim NewT1 As String
Dim j As Integer

If Trim(Nz(Me.T1, "")) + "" <> "" Then
    If T1Range = True Then
      SQLFilter = SQLFilter & " AND [" & Me.T1F & "] BETWEEN " & T1Contain & Me.T1 & T1Contain & " and " & T1Contain & Me.T1R & T1Contain
    ElseIf Me.T1O = "IN" Then
      If T1Contain = Chr(39) Then
        For j = 1 To Len(Me.T1)
            Select Case j
                Case 1
                    NewT1 = Chr(39) & Mid(Me.T1, j, 1)
                Case Len(Me.T1)
                    NewT1 = NewT1 & Mid(Me.T1, j, 1) & Chr(39)
                Case Else
                    Select Case Mid(Me.T1, j, 1)
                        Case ","
                            NewT1 = NewT1 & Chr(39) & ", " & Chr(39)
                        Case " "
                            'do nothing
                        Case Else
                            NewT1 = NewT1 & Mid(Me.T1, j, 1)
                    End Select
            End Select
        Next
        SQLFilter = SQLFilter & " AND [" & Me.T1F & "] IN (" & NewT1 & ")"
     Else
        SQLFilter = SQLFilter & " AND [" & Me.T1F & "] IN (" & Me.T1 & ")"
     End If
    ElseIf Me.T1O = "LIKE" Then
      SQLFilter = SQLFilter & " AND [" & Me.T1F & "] LIKE " & Chr(39) & "*" & Me.T1 & "*" & Chr(39)
    Else
      SQLFilter = SQLFilter & " AND [" & Me.T1F & "] " & Me.T1O & " " & T1Contain & Me.T1 & T1Contain
    End If
End If

If Trim(Nz(Me.T2, "")) + "" <> "" Then
    If T2Range = True Then
      SQLFilter = SQLFilter & " AND [" & Me.T2F & "] BETWEEN " & T2Contain & Me.T2 & T2Contain & " and " & T2Contain & Me.T2R & T2Contain
    ElseIf Me.T2O = "IN" Then
      SQLFilter = SQLFilter & " AND [" & Me.T2F & "] IN (" & Me.T2 & ")"
    ElseIf Me.T2O = "LIKE" Then
      SQLFilter = SQLFilter & " AND [" & Me.T2F & "] LIKE " & Chr(39) & "*" & Me.T2 & "*" & Chr(39)
    Else
      SQLFilter = SQLFilter & " AND [" & Me.T2F & "] " & Me.T2O & " " & T2Contain & Me.T2 & T2Contain
    End If
End If

If Trim(Nz(Me.T3, "")) + "" <> "" Then
    If T3Range = True Then
      SQLFilter = SQLFilter & " AND [" & Me.T3F & "] BETWEEN " & T3Contain & Me.T3 & T3Contain & " and " & T3Contain & Me.T3R & T3Contain
    ElseIf Me.T3O = "IN" Then
      SQLFilter = SQLFilter & " AND [" & Me.T3F & "] IN (" & Me.T3 & ")"
    ElseIf Me.T3O = "LIKE" Then
      SQLFilter = SQLFilter & " AND [" & Me.T3F & "] LIKE " & Chr(39) & "*" & Me.T3 & "*" & Chr(39)
    Else
      SQLFilter = SQLFilter & " AND [" & Me.T3F & "] " & Me.T3O & " " & T3Contain & Me.T3 & T3Contain
    End If
End If
    
If Trim(Nz(Me.T4, "")) + "" <> "" Then
    If T4Range = True Then
      SQLFilter = SQLFilter & " AND [" & Me.T4F & "] BETWEEN " & T4Contain & Me.T4 & T4Contain & " and " & T4Contain & Me.T4R & T4Contain
    ElseIf Me.T4O = "IN" Then
      SQLFilter = SQLFilter & " AND [" & Me.T4F & "] IN (" & Me.T4 & ")"
    ElseIf Me.T4O = "LIKE" Then
      SQLFilter = SQLFilter & " AND [" & Me.T4F & "] LIKE " & Chr(39) & "*" & Me.T4 & "*" & Chr(39)
    Else
      SQLFilter = SQLFilter & " AND [" & Me.T4F & "] " & Me.T4O & " " & T4Contain & Me.T4 & T4Contain
    End If
End If
    
If Trim(Nz(Me.T5, "")) + "" <> "" Then
    If T5Range = True Then
      SQLFilter = SQLFilter & " AND [" & Me.T5F & "] BETWEEN " & T5Contain & Me.T5 & T5Contain & " and " & T5Contain & Me.T5R & T5Contain
    ElseIf Me.T5O = "IN" Then
      SQLFilter = SQLFilter & " AND [" & Me.T5F & "] IN (" & Me.T5 & ")"
    ElseIf Me.T5O = "LIKE" Then
      SQLFilter = SQLFilter & " AND [" & Me.T5F & "] LIKE " & Chr(39) & "*" & Me.T5 & "*" & Chr(39)
    Else
      SQLFilter = SQLFilter & " AND [" & Me.T5F & "] " & Me.T5O & " " & T5Contain & Me.T5 & T5Contain
    End If
End If


'Now get the list boxes
  Dim ParamListBox As Control
  Dim ListItem As Variant
  Dim BuildParameterList As String
  Dim Comma As String
  
  
  Set ParamListBox = Me.L1
  BuildParameterList = "("
  Comma = ""
  If ParamListBox.ItemsSelected.Count > 0 Then
    For Each ListItem In ParamListBox.ItemsSelected
      BuildParameterList = BuildParameterList & Comma & L1Contain & ParamListBox.ItemData(ListItem) & L1Contain
      Comma = ","
    Next ListItem
    BuildParameterList = BuildParameterList & ")"
    SQLFilter = SQLFilter & " AND [" & Me.L1F & "] IN " & BuildParameterList
  End If

  Set ParamListBox = Me.L2
  BuildParameterList = "("
  Comma = ""
  If ParamListBox.ItemsSelected.Count > 0 Then
    For Each ListItem In ParamListBox.ItemsSelected
      BuildParameterList = BuildParameterList & Comma & L2Contain & ParamListBox.ItemData(ListItem) & L2Contain
      Comma = ","
    Next ListItem
    BuildParameterList = BuildParameterList & ")"
    SQLFilter = SQLFilter & " AND [" & Me.L2F & "] IN " & BuildParameterList
  End If

  Set ParamListBox = Me.L3
  BuildParameterList = "("
  Comma = ""
  If ParamListBox.ItemsSelected.Count > 0 Then
    For Each ListItem In ParamListBox.ItemsSelected
      BuildParameterList = BuildParameterList & Comma & L3Contain & ParamListBox.ItemData(ListItem) & L3Contain
      Comma = ","
    Next ListItem
    BuildParameterList = BuildParameterList & ")"
    SQLFilter = SQLFilter & " AND [" & Me.L3F & "] IN " & BuildParameterList
  End If
  
  Set ParamListBox = Me.L4
  BuildParameterList = "("
  Comma = ""
  If ParamListBox.ItemsSelected.Count > 0 Then
    For Each ListItem In ParamListBox.ItemsSelected
      BuildParameterList = BuildParameterList & Comma & L4Contain & ParamListBox.ItemData(ListItem) & L4Contain
      Comma = ","
    Next ListItem
    BuildParameterList = BuildParameterList & ")"
    SQLFilter = SQLFilter & " AND [" & Me.L4F & "] IN " & BuildParameterList
  End If

  Set ParamListBox = Me.L5
  BuildParameterList = "("
  Comma = ""
  If ParamListBox.ItemsSelected.Count > 0 Then
    For Each ListItem In ParamListBox.ItemsSelected
      BuildParameterList = BuildParameterList & Comma & L5Contain & ParamListBox.ItemData(ListItem) & L5Contain
      Comma = ","
    Next ListItem
    BuildParameterList = BuildParameterList & ")"
    SQLFilter = SQLFilter & " AND [" & Me.L5F & "] IN " & BuildParameterList
  End If
  
  
SQLFilter = Replace(SQLFilter, "<start> AND", "")
SQLFilter = Replace(SQLFilter, "<start> ", "")
SQLFilter = Replace(SQLFilter, "<start>", "")

Me.txtFilter = SQLFilter



End Sub







Private Sub Command90_Click()
Call ClearReportParameters

End Sub

Private Sub FromDt_Exit(Cancel As Integer)
    If txtSQLFromDt & "" <> "" Then
        If Not IsDate(txtSQLFromDt) Then
            MsgBox "Please enter a valid date"
            Cancel = True
        End If
    End If
End Sub

Private Sub L1_AfterUpdate()
BuildReportFilter
End Sub

Private Sub L1_Click()
L1K = False
BuildReportFilter
End Sub



Private Sub L1_KeyDown(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub



Private Sub L1_KeyUp(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub

Private Sub L1K_AfterUpdate()
  Dim intcurrentrow As Integer
  Dim ParamListBox As Control
  Set ParamListBox = Me.L1
  
If L1K = True Then
    For intcurrentrow = 0 To ParamListBox.ListCount - 1
    ParamListBox.Selected(intcurrentrow) = False
    Next intcurrentrow
End If

BuildReportFilter

End Sub



Private Sub L2_AfterUpdate()
BuildReportFilter
End Sub

Private Sub L2_Click()
L2K = False
BuildReportFilter
End Sub

Private Sub L2_KeyDown(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub

Private Sub L2_KeyUp(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub

Private Sub L2K_AfterUpdate()
  Dim intcurrentrow As Integer
  Dim ParamListBox As Control
  Set ParamListBox = Me.L2
  
If L2K = True Then
    For intcurrentrow = 0 To ParamListBox.ListCount - 1
    ParamListBox.Selected(intcurrentrow) = False
    Next intcurrentrow

End If

BuildReportFilter
End Sub

Private Sub L3_AfterUpdate()
BuildReportFilter
End Sub

Private Sub L3_Click()
L3K = False
BuildReportFilter
End Sub

Private Sub L3_KeyDown(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub

Private Sub L3_KeyUp(KeyCode As Integer, Shift As Integer)
BuildReportFilter
End Sub

Private Sub L3K_AfterUpdate()
  Dim intcurrentrow As Integer
  Dim ParamListBox As Control
  Set ParamListBox = Me.L3
  
If L3K = True Then
    For intcurrentrow = 0 To ParamListBox.ListCount - 1
    ParamListBox.Selected(intcurrentrow) = False
    Next intcurrentrow

End If

BuildReportFilter


End Sub

Private Sub L4_Click()
L4K = False
BuildReportFilter
End Sub

Private Sub L4K_AfterUpdate()
  Dim intcurrentrow As Integer
  Dim ParamListBox As Control
  Set ParamListBox = Me.L4
  
If L4K = True Then
    For intcurrentrow = 0 To ParamListBox.ListCount - 1
    ParamListBox.Selected(intcurrentrow) = False
    Next intcurrentrow

End If

BuildReportFilter

End Sub

Private Sub L5_Click()
L5K = False
BuildReportFilter
End Sub

Private Sub L5K_AfterUpdate()
  Dim intcurrentrow As Integer
  Dim ParamListBox As Control
  Set ParamListBox = Me.L5
  
If L5K = True Then
    For intcurrentrow = 0 To ParamListBox.ListCount - 1
    ParamListBox.Selected(intcurrentrow) = False
    Next intcurrentrow

End If

BuildReportFilter
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & vbCrLf & ErrSource
End Sub

Private Sub ThruDt_Exit(Cancel As Integer)
    If txtSQLThruDt & "" <> "" Then
        If Not IsDate(txtSQLThruDt) Then
            MsgBox "Please enter a valid date"
            Cancel = True
        End If
    End If
End Sub

Private Sub T1_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T1F_Change()
BuildReportFilter
End Sub

Private Sub T1O_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T1R_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T2_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T2F_Change()
BuildReportFilter
End Sub

Private Sub T2O_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T2R_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T3_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T3F_Change()
BuildReportFilter
End Sub

Private Sub T3O_AfterUpdate()
BuildReportFilter
End Sub

Private Sub T3R_AfterUpdate()
BuildReportFilter
End Sub
Private Sub cmdRiskForm_Click()
On Error GoTo Err_cmdRiskForm_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Concept_RiskFactor"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdRiskForm_Click:
    Exit Sub

Err_cmdRiskForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdRiskForm_Click
    
End Sub

Private Sub AddFieldstoTComboBoxes(TableName As String)

    Dim db As DAO.Database
    Dim TB As Object
    Dim FieldsSet As DAO.Fields
    
    Dim i As Integer
    Dim FieldDataType As String
    
    Set db = CurrentDb
    If LCase(left(TableName, 3)) = "qry" Then
        Set TB = db.QueryDefs(TableName)
    Else
        Set TB = db.TableDefs(TableName)
    End If
    Set FieldsSet = TB.Fields
    
    For i = 0 To FieldsSet.Count - 1
        Select Case FieldsSet(i).Type
            Case 10
                FieldDataType = "T"
            Case 2 To 7
                FieldDataType = "N"
            Case 8
                FieldDataType = "D"
            Case Else
                FieldDataType = ""
        End Select
        Me.T1F.AddItem Item:=FieldsSet(i).Name & "; " & FieldDataType
        Me.T2F.AddItem Item:=FieldsSet(i).Name & "; " & FieldDataType
        Me.T3F.AddItem Item:=FieldsSet(i).Name & "; " & FieldDataType
    Next
End Sub

Private Sub AddFieldstoTComboBoxesADO(TableName As String)


    Dim MyAdo As clsADO
    Dim ADOtbl As ADODB.RecordSet
    Dim ADOfld As ADODB.Field
    Dim FieldDataType As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "Select top 1 * from cms_auditors_reports.dbo." & TableName
    Set ADOtbl = MyAdo.OpenRecordSet
    
    If ADOtbl Is Nothing Then GoTo GetOut
    
    For Each ADOfld In ADOtbl.Fields
        Select Case ADOfld.Type
            Case 10
                FieldDataType = "T"
            Case 2 To 7
                FieldDataType = "N"
            Case 8
                FieldDataType = "D"
            Case Else
                FieldDataType = ""
        End Select
        Me.T1F.AddItem Item:=ADOfld.Name & "; " & FieldDataType
        Me.T2F.AddItem Item:=ADOfld.Name & "; " & FieldDataType
        Me.T3F.AddItem Item:=ADOfld.Name & "; " & FieldDataType
    Next
    ADOtbl.Close
    
GetOut:
    
    Set ADOtbl = Nothing
    
End Sub

Private Sub ClearT1_Click()
'    Me.T1L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'    Me.T1F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'    Me.T1 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'    Me.T1R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T1"""))
'    If Trim(Nz(Me.T1R, "")) + "" <> "" Then
'      Me.T1O = "BETWEEN"
'    Else
'      Me.T1O = "="
'    End If
    PopulateReportParameters "T1"
    PopulateReportParameters "T1R"
    BuildReportFilter
End Sub

Private Sub ClearT2_Click()
'    Me.T2L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'    Me.T2F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'    Me.T2 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'    Me.T2R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T2"""))
'    If Trim(Nz(Me.T2R, "")) + "" <> "" Then
'      Me.T2O = "BETWEEN"
'    Else
'      Me.T2O = "="
'    End If
    PopulateReportParameters "T2"
    PopulateReportParameters "T2R"
    BuildReportFilter
End Sub

Private Sub ClearT3_Click()
'    Me.T3L.Caption = Nz(DLookup("[PromptText]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'    Me.T3F = Nz(DLookup("[FieldName]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'    Me.T3 = Nz(DLookup("[DefaultValue]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'    Me.T3R = Nz(DLookup("[DefaultValueRange]", "REPORT_Parameter", "ReportID = " & Chr(34) & Me.txtReportNumber & Chr(34) & " AND ControlID = ""T3"""))
'    If Trim(Nz(Me.T3R, "")) + "" <> "" Then
'      Me.T3O = "BETWEEN"
'    Else
'      Me.T3O = "="
'    End If
    
    PopulateReportParameters "T3"
    PopulateReportParameters "T3R"
    BuildReportFilter
End Sub

Function ReportingAccessFormADO(strSearchType As String, ReportTable As String, ReportFilter As String, ReportOrderBy As String) As Boolean

    ReportingAccessFormADO = True

    Dim frmNew As New Form_frm_RPT_AccessForm
    Set frmNew = New Form_frm_RPT_AccessForm
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.SearchType = strSearchType
    'frmNew.RecordSource = ReportTable
    
    
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select top 100000 * from cms_auditors_reports.dbo." & ReportTable & IIf(ReportFilter <> "", " where " & ReportFilter, "") & IIf(ReportOrderBy <> "", " order by " & ReportOrderBy, "")

    Set rs = MyAdo.OpenRecordSet
    
    If rs Is Nothing Then
        MsgBox "There was an error obtaining the data from the report table, please contact Alex or James, Thanks.", vbExclamation, "Error in ReportingAccessFormADO"
        GoTo OuttaHere
    End If
    Set frmNew.RecordSet = rs
    
    
    If rs.recordCount = 100000 Then
        MsgBox "This report has reached the limit of 100,000 records!" & vbNewLine & _
            "Please be aware you are not seeing all the results!" & vbNewLine & _
            "Use the filters on the report screen or contact IT" & vbNewLine & _
            vbNewLine & _
            "IT: Enter frm_RPT_AccessForm to the AccessFormName field in Report_Hdr for this report", vbExclamation, "Results Limit Reached"
    End If
    
'    Dim db As DAo.Database
'    Dim tdfld As DAo.TableDef
    Dim fld As ADODB.Field
    
    Dim FieldCounter As Integer
 
'    Set db = CurrentDb()
'    Set tdfld = db.TableDefs(ReportTable)
    
    frmNew.Caption = "frm_" & ReportTable & " (AccessForm ADO)"
    
    FieldCounter = 1
    
    For Each fld In rs.Fields   'loop through all the fields of the tables
        frmNew("Text" & FieldCounter).Properties("ColumnHidden") = False
        frmNew("Text" & FieldCounter).ControlSource = fld.Name
        frmNew("Label" & FieldCounter).Caption = fld.Name
        FieldCounter = FieldCounter + 1
    Next
    
    Dim i As Integer
    
    For i = FieldCounter To 270 'rest of columns are not used so hide them
        frmNew("Text" & i).Properties("ColumnHidden") = True
    Next

'    frmNew.FilterOn = True
'    frmNew.Filter = ReportFilter
    frmNew.visible = True
    rs.Close
    Set rs = Nothing
    Exit Function
    
    
OuttaHere:
    
    ReportingAccessFormADO = False
    Set rs = Nothing
    Exit Function

   'DoCmd.Close acForm, Me.Name
    
End Function


Sub SeudoMaximize()
    DoCmd.MoveSize 0, 0, (MonitorWidth * 15) - 300, (MonitorHeight * 15) - 1450
End Sub


Sub ReportingAccessForm(strSearchType As String, ReportTable As String, ReportFilter As String, ReportOrderBy As String)

    Dim frmNew As New Form_frm_RPT_AccessForm
    Set frmNew = New Form_frm_RPT_AccessForm
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.SearchType = strSearchType
    frmNew.RecordSource = "Select * from " & ReportTable & IIf(ReportOrderBy <> "", " order by " & ReportOrderBy, "")
    
    Dim db As DAO.Database
    Dim tdfld As DAO.TableDef
    Dim fld As DAO.Field
    
    Dim FieldCounter As Integer
 
    Set db = CurrentDb()
    Set tdfld = db.TableDefs(ReportTable)
    
    frmNew.Caption = "frm_" & ReportTable & " (AccessForm DAO)"
    
    FieldCounter = 1
    
    For Each fld In tdfld.Fields    'loop through all the fields of the tables
        frmNew("Text" & FieldCounter).Properties("ColumnHidden") = False
        frmNew("Text" & FieldCounter).ControlSource = fld.Name
        frmNew("Label" & FieldCounter).Caption = fld.Name
        FieldCounter = FieldCounter + 1
    Next
    
    Dim i As Integer
    
    For i = FieldCounter To 270 'rest of columns not used so hide
        frmNew("Text" & i).Properties("ColumnHidden") = True
    Next

    frmNew.filter = ReportFilter
    frmNew.FilterOn = True
    frmNew.visible = True

    'DoCmd.Close acForm, Me.Name
    
End Sub
