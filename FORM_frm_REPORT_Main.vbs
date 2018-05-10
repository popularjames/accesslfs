Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrUserProfile As String

'=============================================
' ID:          frm_REPORT_Main
' Author:
' Create Date:
' Description:
'      Prompt the user to select a report.
'
' Modification History:
'   2010-03-15 by BJD to Find the q_ReportTree2 record cooresponding to the currently selected
' report in tvwTasks_NodeClick and add the option to use the old generic report form.
'
' =============================================

Const CstrFrmAppID As String = "ReportMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub Form_Close()
    Me.sub_form.SourceObject = Me.sub_form.Tag
End Sub


Private Sub Form_Load()

    Dim iAppPermission As Integer
    
    Me.Caption = "Report Maintenance"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
        
    mstrUserProfile = GetUserProfile()
    
    Me.sub_form.SourceObject = ""
    'Me.sub_form.visible = False
    'Me.lblSubAppTitle.visible = False
    
    RefreshReportTree
    
End Sub


Private Sub cmdExpandList_Click()
  RefreshReportTree
  Dim ReportNode As node
  For Each ReportNode In Me.tvwTasks.Nodes
    ReportNode.Expanded = True
  Next
End Sub


Private Sub cmdCollapseList_Click()
  RefreshReportTree
  Dim ReportNode As node
  For Each ReportNode In Me.tvwTasks.Nodes
    ReportNode.Expanded = False
  Next
End Sub

Private Sub ResetNodes()
  Dim ReportNode As node
  For Each ReportNode In Me.tvwTasks.Nodes
    ReportNode.ForeColor = vbBlack
    ReportNode.Bold = False
  Next
End Sub

Private Sub RefreshReportTree()

'On Error GoTo RefreshReportTreeFail
    
mstrUserProfile = GetUserProfile()

'First, clear out and rebuild queries that control the tree
'Dim StrSQL As String
'Dim qdf As DAo.QueryDef
'
''If they already exist, delete all three queries
'  For Each qdf In CurrentDb.QueryDefs
'    If qdf.Name = "q_ReportTree1" Then
'      CurrentDb.QueryDefs.Delete "q_ReportTree1"
'
'    End If
'
'    If qdf.Name = "q_ReportTree2" Then
'      CurrentDb.QueryDefs.Delete "q_ReportTree2"
'
'    End If
'
'  Next
'
'
'   'Build fresh queries
'    StrSQL = ""
'    StrSQL = StrSQL & " SELECT DISTINCT ReportingGroup as Level1 "
'    StrSQL = StrSQL & " FROM Report_Hdr "
'    StrSQL = StrSQL & " WHERE ActiveFlag = 1 "
'    StrSQL = StrSQL & " ORDER BY 1 "
'    Set qdf = CurrentDb.CreateQueryDef("q_ReportTree1", StrSQL)
'
'    StrSQL = ""
'    StrSQL = StrSQL & " SELECT DISTINCT ReportingGroup as Level1, hd.SubGroupSort, hd.ReportingSubGroup as Level2, hd.ReportSort, hd.ReportName as Level3, hd.ReportId, pr.ProfileID, gt.FormName, gt.AccessForm "
'    StrSQL = StrSQL & " FROM (Report_Hdr hd LEFT JOIN GENERAL_Tabs gt ON hd.ReportID = gt.FormValue) "
'    StrSQL = StrSQL & " LEFT JOIN GENERAL_Tabs_Linked_ProfileIDs pr ON gt.RowID = pr.RowID "
'    StrSQL = StrSQL & " WHERE pr.ProfileID = " & Chr(34) & mstrUserProfile & Chr(34)
'    StrSQL = StrSQL & " AND gt.AccessForm = " & Chr(34) & "frm_REPORT_Main" & Chr(34)
'    StrSQL = StrSQL & " AND ActiveFlag = 1"
'    StrSQL = StrSQL & " ORDER BY 1, 2, 3, 4, 5"
'
'    Set qdf = CurrentDb.CreateQueryDef("q_ReportTree2", StrSQL)
'
'    'strSQL = ""
'    'strSQL = strSQL & " SELECT DISTINCT Level1, Level2, Level3 "
'    'strSQL = strSQL & " FROM TestTable "
'    'strSQL = strSQL & " ORDER BY Level1, Level2, Level3"
'    'Set qdf = CurrentDb.CreateQueryDef("q_ReportTree3", strSQL)
'
'
'  'Populate the nodes
'   Dim strMessage As String
'   Dim dbs As DAo.Database
'   Dim rst As DAo.Recordset
'

    Dim MyAdo As clsADO
    Dim rs1 As ADODB.RecordSet
    
    Dim strSQL As String
    Dim iResult As Integer
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT ReportingGroup as Level1 "
    strSQL = strSQL & " FROM CMS_AUDITORS_CLAIMS.dbo.Report_Hdr "
    strSQL = strSQL & " WHERE ActiveFlag = 1 "
    strSQL = strSQL & " ORDER BY 1 "

    MyAdo.sqlString = strSQL
    Set rs1 = MyAdo.OpenRecordSet
    
    Dim rs2 As ADODB.RecordSet
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT ReportingGroup as Level1, hd.SubGroupSort, hd.ReportingSubGroup as Level2, hd.ReportSort, hd.ReportName as Level3, hd.ReportId, pr.ProfileID, gt.FormName, gt.AccessForm "
    strSQL = strSQL & " FROM cms_auditors_claims.dbo.Report_Hdr hd LEFT JOIN cms_auditors_claims.dbo.GENERAL_Tabs gt ON hd.ReportID = gt.FormValue "
    strSQL = strSQL & " LEFT JOIN cms_auditors_claims.dbo.GENERAL_Tabs_Linked_ProfileIDs pr ON gt.RowID = pr.RowID "
    strSQL = strSQL & " WHERE pr.ProfileID = '" & mstrUserProfile & "' "
    strSQL = strSQL & " AND gt.AccessForm = 'frm_REPORT_Main' "
    strSQL = strSQL & " AND ActiveFlag = 1"
    strSQL = strSQL & " ORDER BY 1, 2, 3, 4, 5"

    MyAdo.sqlString = strSQL
    Set rs2 = MyAdo.OpenRecordSet


'   Dim strQuery1 As String
'   Dim strQuery2 As String
'   Dim strQuery3 As String
   
   Dim strNode1Text As String
   Dim strNode2Text As String
   Dim strNode3Text As String
   
   Dim strNodeReportId As String
      
   Dim nod As Object
  
   Dim strVisibleText As String
   
   Dim LastSubGroup As String
   
'   Set dbs = CurrentDb()
'   strQuery1 = "q_ReportTree1"
'   strQuery2 = "q_ReportTree2"
   'strQuery3 = "q_ReportTree3"
     
   Me![tvwTasks].Nodes.Clear
   With Me![tvwTasks]

      'Fill Level 1
      'Set rst = dbs.OpenRecordSet(strQuery1, dbOpenForwardOnly)
      
      Do Until rs1.EOF
         strNode1Text = "L1" & rs1![Level1]
         Set nod = .Nodes.Add(Key:=strNode1Text, Text:=rs1![Level1])
         nod.Expanded = False
         rs1.MoveNext
      Loop
      rs1.Close
      
      
      'Fill Level 2
      'Will fill subgroups if exist. If don't it will fill Level 3 (reports)
      'Set rst = dbs.OpenRecordSet(strQuery2, dbOpenForwardOnly) '

      Do Until rs2.EOF
         strNode1Text = "L1" & rs2![Level1]
         strNode2Text = "L2" & rs2![Level2]
         strNode3Text = "L3" & rs2![Level3]
         strNodeReportId = rs2![ReportID]

        If Not (IsNull(rs2![Level2]) Or rs2![Level2] = "N/A") Then
            If rs2![Level2] <> LastSubGroup Then
                strVisibleText = Nz(rs2![Level2], "")
                .Nodes.Add relative:=strNode1Text, relationship:=tvwChild, Key:=strNode1Text & strNode2Text, Text:=rs2![Level2]
            End If
        Else
            strVisibleText = Nz(rs2![Level3], "")
            .Nodes.Add relative:=strNode1Text, relationship:=tvwChild, Key:=strNode1Text & strNode3Text, Text:=rs2![Level3]
        End If
         LastSubGroup = Nz(rs2![Level2], "")
         rs2.MoveNext
      Loop
      'rst.Close
      
     
      'Fill Level 3
      'Only for reports nested under subgroups
      'Set rst = dbs.OpenRecordSet(strQuery2, dbOpenForwardOnly) '

      rs2.MoveFirst
      Do Until rs2.EOF
         strNode1Text = "L1" & rs2![Level1]
         strNode2Text = "L2" & rs2![Level2]
         strNode3Text = "L3" & rs2![Level3]
      
          If Not (IsNull(rs2![Level2]) Or rs2![Level2] = "N/A") Then
            strVisibleText = Nz(rs2![Level3], "")
            .Nodes.Add relative:=strNode1Text & strNode2Text, relationship:=tvwChild, Key:=strNode2Text & strNode3Text, Text:=rs2![Level3]
          End If
   
         rs2.MoveNext
      Loop

      
   End With
   
   'dbs.Close
 
 
   rs2.Close

   Exit Sub
    

    
RefreshReportTreeFail:
    MsgBox "RefreshReportTree procedure has failed.  Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    Exit Sub

End Sub



Private Sub tvwTasks_NodeClick(ByVal node As Object)

'If node.FullPath <> node.Text Then 'i.e. only perform this procedure if this is NOT a root node
If node.Child Is Nothing Then

ResetNodes
node.ForeColor = vbBlue
node.Bold = True


'Dim rs As DAo.Recordset
Dim ReportNumber As String
Dim DisableFrontEnd As String
Dim ReportParent As String
Dim AccessForm As String

    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim LookUp As ADODB.RecordSet
    Dim vLastRun As ADODB.RecordSet
    Dim strSQL As String
    Dim iResult As Integer
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT ReportingGroup as Level1, hd.SubGroupSort, hd.ReportingSubGroup as Level2, hd.ReportSort, hd.ReportName as Level3, hd.ReportId, pr.ProfileID, gt.FormName, gt.AccessForm "
    strSQL = strSQL & " FROM cms_auditors_claims.dbo.Report_Hdr hd LEFT JOIN cms_auditors_claims.dbo.GENERAL_Tabs gt ON hd.ReportID = gt.FormValue "
    strSQL = strSQL & " LEFT JOIN cms_auditors_claims.dbo.GENERAL_Tabs_Linked_ProfileIDs pr ON gt.RowID = pr.RowID "
    strSQL = strSQL & " WHERE pr.ProfileID = '" & mstrUserProfile & "' "
    strSQL = strSQL & " AND gt.AccessForm = 'frm_REPORT_Main' "
    strSQL = strSQL & " AND ActiveFlag = 1"
    strSQL = strSQL & " ORDER BY 1, 2, 3, 4, 5"
    
    MyAdo.sqlString = strSQL
    Set rs = MyAdo.OpenRecordSet


'Set rs = CurrentDb.OpenRecordSet("q_ReportTree2", dbOpenSnapshot, dbSeeChanges)
If Not (rs.BOF And rs.EOF) Then

    If node.Parent.Parent Is Nothing Then ReportParent = node.Parent Else ReportParent = node.Parent.Parent

    'Find the record cooresponding to the currently selected report.
    'ReportNumber = Nz(DLookup("[ReportId]", "Report_Hdr", "[ReportingGroup] = " & Chr(34) & ReportParent & Chr(34) & " And [ReportName] = " & Chr(34) & node.Text & Chr(34)), "N")
    
    MyAdo.sqlString = "Select *  from cms_auditors_claims.dbo.Report_Hdr where ReportingGroup = '" & ReportParent & "' And ReportName = '" & node.Text & "'"
    Set LookUp = MyAdo.OpenRecordSet
    
    
    
    ReportNumber = Nz(LookUp("ReportID"), "N")

    rs.MoveFirst
    Do While Not rs.EOF
        If rs("ReportId") = ReportNumber Then
            Exit Do
        End If
        rs.MoveNext
    Loop
    'Error if record not found.
    If rs.EOF Then
        MsgBox "Report has not been defined.  Contact application support.  "
        LookUp.MoveFirst
        Exit Sub
    End If

    'Me.sub_form.SourceObject = ""
    'Me.sub_form.visible = True
    Me.sub_form.SourceObject = rs("FormName")
    
    Select Case Nz(rs("FormName"), "")
        Case "frm_REPORT_Generic"
            
            'ReportNumber = Nz(DLookup("[ReportId]", "Report_Hdr", "[ReportingGroup] = " & Chr(34) & ReportParent & Chr(34) & " And [ReportName] = " & Chr(34) & node.Text & Chr(34)), "N")

            'working on adding all the fields on one shot here!!!!!!!!!!!!!!!!!!!!
            
         
            'DisableFrontEnd = Nz(DLookup("[DisableFrontEnd]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "N")
            If Nz(LookUp("DisableFrontEnd"), "N") = "Y" Then
              Me.sub_form.Form.opt2Refresh.Enabled = False
            End If
            
            If Nz(LookUp("AccessReportName"), "") = "" Then
              Me.sub_form.Form.opt1Report.Enabled = False
              Me.sub_form.Form.optOutput = 2
            End If
            
            'AccessForm = Nz(DLookup("[AccessFormName]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            Select Case Nz(LookUp("AccessFormName"), "")
                Case ""
                    Me.sub_form.Form.opt3Form.Enabled = False
                    Me.sub_form.Form.optOutput = 2
                Case "frm_RPT_AccessForm"
                    Me.sub_form.Form.opt3Form.Enabled = True
                    Me.sub_form.Form.optOutput = 3
                Case Else
                    Me.sub_form.Form.opt3Form.Enabled = True
                    Me.sub_form.Form.optOutput = 2
            End Select
            
           
            Me.sub_form.Form.txtUserId = Identity.UserName()
            Me.sub_form.Form.txtReportNumber = ReportNumber
            
            Me.sub_form.Form.txtReportName = Nz(LookUp("ReportingGroup"), "") & " / " & Nz(LookUp("ReportingSubGroup"), "") & " / " & Nz(LookUp("ReportName"), "")
            'Nz(DLookup("[ReportingGroup]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "") & _
            '    Nz(" / " + DLookup("[ReportingSubGroup]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "") & _
            '    Nz(" : " + DLookup("[ReportName]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")

            Me.sub_form.Form.txtStoredProc = Nz(LookUp("StoredProc"), "")
            'Nz(DLookup("[StoredProc]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            Me.sub_form.Form.txtOutputTable = Nz(LookUp("OutputTable"), "")
            'Nz(DLookup("[OutputTable]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            Me.sub_form.Form.txtAccessReportName = Nz(LookUp("AccessReportName"), "")
            'Nz(DLookup("[AccessReportName]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            
            Me.sub_form.Form.txtAccessFormName = Nz(LookUp("AccessFormName"), "")
            'Nz(DLookup("[AccessFormName]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            MyAdo.sqlString = "Select LastRunDt, LastRunDurationText from cms_auditors_code.dbo.v_REPORT_LastRun where ReportID = '" & ReportNumber & "'"
            Set vLastRun = MyAdo.OpenRecordSet
            If Not (vLastRun.EOF Or vLastRun.BOF) Then
            
                Me.sub_form.Form.lblOptionLastRun.Caption = "Use last run  (" & Format(Nz(vLastRun("LastRunDt"), ""), "m/d/yy hh:mm AMPM") & ")"
                '"Use last run  (" & Format(Nz(DLookup("[LastRunDt]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), ""), "m/d/yy hh:mm AMPM") & ")"
                
                Me.sub_form.Form.lblOptionRefreshData.Caption = "Refresh  (approximately " & Nz(vLastRun("LastRunDurationText"), "") & ")"
                '"Refresh  (approximately " & Nz(DLookup("[LastRunDurationText]", "v_REPORT_LastRun", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "") & ")"
            End If
            vLastRun.Close
            Set vLastRun = Nothing

            
            Me.sub_form.Form.txtReportDesc = Nz(LookUp("ReportDesc"), "")
            'Nz(DLookup("[ReportDesc]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            Me.sub_form.Form.txtQueryOrderBy = Nz(LookUp("QueryOrderBy"), "")
            'Nz(DLookup("[QueryOrderBy]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
            
            

            If Me.sub_form.Form.txtStoredProc <> "" Then
                Me.sub_form.Form.txtSQLFromDt = #1/1/2000#
                Me.sub_form.Form.txtSQLThruDt = #12/31/2020#
                
            Else
                Me.sub_form.Form.txtSQLFromDt = ""
                Me.sub_form.Form.txtSQLThruDt = ""
                
            End If
            
            If left(Me.sub_form.Form.txtOutputTable, 1) = "r" And Me.sub_form.Form.txtAccessFormName <> "" Then
                LinkTable "SQL", "DS-FLD-009", "CMS_AUDITORS_Reports", Me.sub_form.Form.txtOutputTable '--added 3/30/2011
            End If
            
            Call Me.Controls("sub_form").Form.ClearReportParameters
            Call Me.Controls("sub_form").Form.PopulateReportParameters
            
            Me.sub_form.Form.CmdRunReport.SetFocus
            
        'Call the older simple generic report form under the new Report Model.
        Case "frm_RPT_Generic_Report_Simple"
            Me.sub_form.Form.StoredProcName = Nz(LookUp("StoredProc"), "")
            'Nz(DLookup("[StoredProc]", "Report_Hdr", "[ReportID] = " & Chr(34) & ReportNumber & Chr(34)), "")
           
            
            
    End Select
    
    rs.Close
    Set rs = Nothing
    
    LookUp.Close
    Set LookUp = Nothing
    
    
Else
    MsgBox "Application form has not been defined"
End If
    
End If


End Sub
