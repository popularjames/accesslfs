Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_KeyPress(KeyAscii As Integer)

'capturing keystroke

If KeyAscii = 99 Or KeyAscii = 67 Then '"C" or "c" for Copy
    
    'selection width should be the total number of fields and the selection height should be one
    'so this will work only if the record selector was clicked for only one record
    
    If Me.SelWidth = Me.RecordSet.Fields.Count And Me.SelHeight = 1 Then
    
     Dim ErrorMessage As String
    
     Dim F1 As Form
     
     Dim rsRHdr As RecordSet

     'expanding sub form and sub sub form
     SendKeys "+^{down}"
     SendKeys "+^{down}"
   
     'continue only after expanding
     DoEvents

     Set F1 = Me
     
     Set rsRHdr = F1.RecordsetClone
    
     NewReportName = InputBox("You are making a copy of report " & Me("ReportID") & vbNewLine & vbNewLine & _
                            "Enter New Report ID: ", "Copying Report...")
                            
     ExistingReport = DLookup("[ReportID]", "Report_Hdr", "[ReportID]='" & NewReportName & "'")
     
     ErrorMessage = ""
     
     If NewReportName = "" Then
        ErrorMessage = "Error: Empty New Report ID!"
     End If
     
     
     If Not Nz(ExistingReport, "") = "" Then
        ErrorMessage = "Error: Report ID " & ExistingReport & " already exists!"
     End If
     
     If ErrorMessage = "" Then
     
        Dim MyAdo As clsADO
    
   
        'making up a "random" number for new report name
'        currentTime = Format(Now, "mmss")
'        NewReportName = "NewReport" & currentTime
    
        'Adding Record into Reports_Hdr
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("REPORT_Hdr")
        strsql1 = " INSERT INTO CMS_AUDITORS_CLAIMS.dbo.REPORT_Hdr (ReportID, ReportingGroup, OutputTable, StoredProc, DisableFrontEnd, RunSchedule, SQLStartDt, SQLEndDt, FilterType, ReportName, ActiveFlag) "
        strsql2 = " VALUES ('" & NewReportName & "', '" & Me.ReportingGroup & "', 'RPT_" & NewReportName & "', 'usp_RPT_" & NewReportName & "', '" & Me.DisableFrontEnd & "', '" & Me.RunSchedule & "', '" & Me.SQLStartDt & "', '" & Me.SQLEndDt & "', '" & Me.FilterType & "', '" & Me.ReportName & " COPY', '" & 1 & "')"
        MyAdo.sqlString = strsql1 & strsql2
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute

        'Adding Record into General_Tabs
        With Me.ctl_frm_ADMIN_Report_sub2.Form
            MyAdo.ConnectionString = GetConnectString("GENERAL_Tabs")
            strsql1 = " INSERT INTO CMS_AUDITORS_CLAIMS.dbo.GENERAL_Tabs (TabName, FormName, AccessForm, FormValue) "
            strsql2 = " VALUES ('" & NewReportName & "', '" & .Controls("FormName") & "', '" & .Controls("AccessForm") & "', '" & NewReportName & "')"
            MyAdo.sqlString = strsql1 & strsql2
            MyAdo.SQLTextType = sqltext
            MyAdo.Execute
        End With
        
        'Finding Original RowId to copy permissions from
        OriginalRowID = DLookup("[RowID]", "General_Tabs", "[Tabname]='" & Me.ctl_frm_ADMIN_Report_sub2.Form.Controls("TabName") & "'")
        
        'Finding New RowID to copy permissions to
        DestinationRowID = DLookup("[RowID]", "General_Tabs", "[Tabname]='" & NewReportName & "'")
        
        'Adding copied permission records into GENERAL_Tabs_Linked_ProfileIDs
        MyAdo.ConnectionString = GetConnectString("GENERAL_Tabs_Linked_ProfileIDs")
        strsql1 = " INSERT INTO CMS_AUDITORS_CLAIMS.dbo.GENERAL_Tabs_Linked_ProfileIDs (RowID, ProfileID) "
        strsql2 = " SELECT '" & DestinationRowID & "', ProfileID FROM CMS_AUDITORS_CLAIMS.dbo.GENERAL_Tabs_Linked_ProfileIDs"
        strsql3 = " WHERE RowID = '" & OriginalRowID & "'"
        MyAdo.sqlString = strsql1 & strsql2 & strsql3
        MyAdo.SQLTextType = sqltext
        MyAdo.Execute
        
        MsgBox "New Report (ID: " & NewReportName & ") has been added under (" & Me.ReportingGroup & ")"
        
        Me.Requery
        
        
     Else
        
        MsgBox ErrorMessage & vbNewLine & vbNewLine & "Copy operation was cancelled.", vbOKOnly, "Copying Report... "
        
     End If
     
     Me.SubdatasheetExpanded = False
     
     End If
     
End If

End Sub
