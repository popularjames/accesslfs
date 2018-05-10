Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This form is used to import objects and data from legacy Decipher
' databases in to Decipher 3+
'
' SA 7/17/2012 - Created class
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private WithEvents clsImport As MUT_ClsImportObjects
Attribute clsImport.VB_VarHelpID = -1
Private Const StatusImport As String = "IMPORT"
Private Const StatusExclude As String = "EXCLUDE"
Private Const MinSourceVersion As String = "2.5.1000"
Private ImportFileVersion As String

Private Sub CmdCodeUpgradeManager_Click()
DoCmd.OpenForm "MUT_CodeUpgradeManager", acNormal
End Sub

Private Sub Form_Load()
    'Defaults
    lblImportStatus.Caption = vbNullString
    lblImportFileVersion.Caption = vbNullString
    AddImportHeader
    
    'Hide tabs for apps that are not installed
    If Not IsProductInstalled("Decipher Screens") And Not IsProductInstalled("App Screens") Then
        pgImportScreens.visible = False
    End If
    If Not TableExists("CFG_CfgLink", CurrentDb) Then
        pgLinks.visible = False
    End If
    If Not IsProductInstalled("ClaimsPlus Framework") Then
        pgClaimsPlus.visible = False
    End If
    If Not IsProductInstalled("Projects") Then
        pgProjects.visible = False
    End If
    If Not IsProductInstalled("WorkFiles") Then
        pgWorkFiles.visible = False
    End If
    If Not IsProductInstalled("Vendor Management") Then
        pgVendorManagement.visible = False
    End If
    If Not IsProductInstalled("Project Management") Then
        pgProjectManagement.visible = False
    End If
    If Not IsProductInstalled("Dup Tool") Then
        PgDupTool.visible = False
    End If
End Sub

Private Sub AddImportHeader()
    lstAccessObjects.RowSource = vbNullString
    lstAccessObjects.AddItem "Type;Name;Status", 0
End Sub

Private Sub cmdBrowse_Click()
'Browse for app file
On Error GoTo ErrorHappened
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Decipher file (version 2.6)"
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add "Access 2010", "*.accdb"
        If .show Then
            txtAccessFile = .SelectedItems(1)
            GetImportFileVersion
        End If
    End With
ExitNow:
On Error Resume Next
    Set fd = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error loading Decipher file"
    Resume ExitNow
End Sub

Private Sub GetImportFileVersion()
'Get version of legacy Decipher
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(txtAccessFile)
    ImportFileVersion = db.TableDefs("CnlyScreensVersions").Properties("Description")
    
    'Set label color
    If Not IsVersionOK Then
        lblImportFileVersion.ForeColor = vbRed
    Else
        lblImportFileVersion.ForeColor = vbBlack
    End If
    
ExitNow:
On Error Resume Next
    lblImportFileVersion.Caption = "Decipher Version: " & ImportFileVersion
    db.Close
    Set db = Nothing
Exit Sub
ErrorHappened:
    ImportFileVersion = "Error"
    Resume ExitNow
End Sub

Private Function IsVersionOK() As Boolean
    'Check that the version is not older than the minimum
    Dim Result As Boolean
    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If SortableVersion(ImportFileVersion) < SortableVersion(MinSourceVersion) Then
            Result = False
        Else
            Result = True
        End If
    Else
        Result = False
    End If
    
    IsVersionOK = Result
End Function

Private Sub cmdImportObjects_Click()
    'Import select objects in list
    If lstAccessObjects.ListCount > 1 Then
        DoCmd.Hourglass True
        If IsVersionOK Then
            RebuildObjectLists
            With clsImport
                .ImportAllObjects
                
                lblImportStatus.Caption = "Imported " & .GetTotalObjectCount & _
                    " objects with " & .GetImportErrorCount & " errors."
            End With
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                    MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
        DoCmd.Hourglass False
    Else
        MsgBox "There aren't any objects in the list to import", vbCritical, "No objects to import"
    End If
    
End Sub

Private Sub RebuildObjectLists()
'Build new lists based on list box selections
On Error GoTo ErrorHappened
    Dim i As Integer
    Dim TableList As New Collection
    Dim QueryList As New Collection
    Dim FormList As New Collection
    Dim ModuleList As New Collection
    Dim MacroList As New Collection
    Dim ReportList As New Collection
    Dim ObjType As String
    Dim ObjName As String

    If lstAccessObjects.ListCount > 1 Then
        For i = 0 To lstAccessObjects.ListCount - 1
            ObjType = lstAccessObjects.Column(0, i)
            ObjName = lstAccessObjects.Column(1, i)
            
            If lstAccessObjects.Column(2, i) = StatusImport Then
                Select Case ObjType
                    Case "TABLE"
                        TableList.Add ObjName
                    Case "QUERY"
                        QueryList.Add ObjName
                    Case "FORM"
                        FormList.Add ObjName
                    Case "MODULE"
                        ModuleList.Add ObjName
                    Case "MACRO"
                        MacroList.Add ObjName
                    Case "REPORT"
                        ReportList.Add ObjName
                End Select
            End If
        Next
        
        'Update import lists
        With clsImport
            .SetTableList = TableList
            .SetQueryList = QueryList
            .SetFormList = FormList
            .SetModuleList = ModuleList
            .SetMacroList = MacroList
            .SetReportList = ReportList
        End With
    End If
ExitNow:
On Error Resume Next
    
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error updating object list"
    Resume ExitNow
End Sub

Private Sub cmdLoadObjectList_Click()
'Load objects into listbox
On Error GoTo ErrorHappened

    If LenB(Nz(txtAccessFile, vbNullString)) > 0 Then
        If Not chkImportForms And Not chkImportMacros And Not chkImportModules And _
            Not chkImportQueries And Not chkImportReports And Not chkImportTables Then
            
            MsgBox "You must select at least 1 object type", vbInformation, "Select object type"
        Else
            lblImportStatus.Caption = vbNullString
            LoadObjectList
        End If
    Else
        MsgBox "Please select a file to import", vbInformation, "Select Decipher File"
    End If
ExitNow:
On Error Resume Next
    
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error loading loading list"
    Resume ExitNow
End Sub

Private Sub LoadObjectList()
'Load objects into listbox for review before import
On Error GoTo ErrorHappened
    Dim i As Integer
    
    DoCmd.Hourglass True
    
    Set clsImport = New MUT_ClsImportObjects
    With clsImport
        .ExcludeObjectsTable "MUT_ExcludeObjects", "ObjectName", "ObjectType"
        .SourceDatabase = txtAccessFile
        
        .CopyForms = chkImportForms
        .CopyModules = chkImportModules
        .CopyTables = chkImportTables
        .CopyQueries = chkImportQueries
        .CopyReports = chkImportReports
        .CopyMacros = chkImportMacros
        
        .LoadObjectList
        
        lstAccessObjects.RowSource = vbNullString
        AddImportHeader
        
        'Add forms to list
        For i = 1 To .GetFormList.Count
            lstAccessObjects.AddItem "FORM;" & .GetFormList.Item(i) & ";" & StatusImport
        Next
        
        'Add modules to list
        For i = 1 To .GetModuleList.Count
            lstAccessObjects.AddItem "MODULE;" & .GetModuleList.Item(i) & ";" & StatusImport
        Next
        
        'Add tables to list
        For i = 1 To .GetTableList.Count
            lstAccessObjects.AddItem "TABLE;" & .GetTableList.Item(i) & ";" & StatusImport
        Next
    
        'Add queries to list
        For i = 1 To .GetQueryList.Count
            lstAccessObjects.AddItem "QUERY;" & .GetQueryList.Item(i) & ";" & StatusImport
        Next
        
        'Add reports to list
        For i = 1 To .GetReportList.Count
            lstAccessObjects.AddItem "REPORT;" & .GetReportList.Item(i) & ";" & StatusImport
        Next
        
        'Add macros to list
        For i = 1 To .GetMacroList.Count
            lstAccessObjects.AddItem "MACRO;" & .GetMacroList.Item(i) & ";" & StatusImport
        Next
        
        If lstAccessObjects.ListCount = 1 Then
            DoCmd.Hourglass False
            MsgBox "No new objects were found that match your settings." & vbCrLf & vbCrLf & _
                "If you alread imported your objects, you won't see them again." & vbCrLf & vbCrLf & _
                "If you think there is an error, you can adjust the settings in the table 'MUT_ExcludeObjects' to make sure your objects are imported.", vbInformation, "No objects to import"
        End If
    End With
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error loading loading list"
    Resume ExitNow
End Sub

Private Sub lstAccessObjects_DblClick(Cancel As Integer)
    UpdateListItemStatus lstAccessObjects.ListIndex + 1
End Sub

Private Sub clsImport_ImportStatus(ByVal ObjectType As String, ByVal ObjectName As String, ByVal Status As String)
'Locate item in listbox and update status
     
     Dim i As Integer
     For i = 0 To lstAccessObjects.ListCount - 1
        If ObjectType = lstAccessObjects.Column(0, i) And ObjectName = lstAccessObjects.Column(1, i) Then
            Exit For
        End If
     Next
     
     UpdateListItemStatus i, Status

End Sub

Private Sub clsImport_StatusMessage(ByVal Message As String)
    lblImportStatus.Caption = Message
End Sub

Private Sub UpdateListItemStatus(ByVal index As Integer, Optional ByVal Status As String = vbNullString)

    Dim ObjType As String
    Dim ObjName As String
    Dim ObjStatus As String
    
    If index >= 0 Then
        'Get existing values
        ObjType = lstAccessObjects.Column(0, index)
        ObjName = lstAccessObjects.Column(1, index)
        ObjStatus = lstAccessObjects.Column(2, index)
        
        'Toggle status
        If LenB(Status) = 0 Then
            If ObjStatus = StatusImport Then
                Status = StatusExclude
            ElseIf ObjStatus = StatusExclude Then
                Status = StatusImport
            Else
                Status = ObjStatus
            End If
        End If
        
        'Update row
        lstAccessObjects.RemoveItem index
        lstAccessObjects.AddItem ObjType & ";" & ObjName & ";" & Status, index
        lstAccessObjects.Selected(index) = True
    End If
End Sub

Private Sub cmdImportProjects_Click()
'Copy Projects from specified database into current database
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim ErrorMsg As String
    Dim ErrorCount As Integer
    Dim addInManager As New CT_ClsCnlyAddinSupport

    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All Projects in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "Confirm Projects import") = vbYes Then
            
                SQL = "SELECT VPS.SQL, VPS.Notes " & _
                    "FROM MUT_ProjectsVersionsPaths AS VP " & _
                    "INNER JOIN MUT_ProjectsVersionsPathsSQL AS VPS ON " & _
                    "VP.PathID = VPS.PathID " & _
                    "WHERE '" & SortableVersion(ImportFileVersion) & "' Between SortableVersion(VP.MinSrcVer) AND SortableVersion(MaxSrcVer) " & _
                    "AND '" & SortableVersion(VersionTemplate) & "' BETWEEN SortableVersion(MinDestVer) AND SortableVersion(MaxDestVer) " & _
                    "ORDER BY VPS.Sort"
                
                ErrorCount = 0
                lstProjectImport.RowSource = vbNullString
                
                DoCmd.Hourglass True
                Application.Echo False
                
                Set db = CurrentDb
                Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
    
                Do Until rs.EOF
                    SQL = rs!SQL
                    SQL = Replace(SQL, "$DbName", txtAccessFile)
    
                    If ExecuteScript(db, SQL, ErrorMsg) Then
                        lstProjectImport.AddItem Nz(rs!Notes, "Executed SQL Statement")
                    Else
                        ErrorCount = ErrorCount + 1
                        lstProjectImport.AddItem "*** ERROR!: " & rs!Notes & " - " & ErrorMsg & " ***"
                    End If
        
                    rs.MoveNext
                Loop
    
                MsgBox "Projects import completed with " & ErrorCount & " errors.", vbInformation, "Projects import complete"
                
                'Move to end of list
                If lstProjectImport.ListCount > 0 Then
                    lstProjectImport.Selected(lstProjectImport.ListCount - 1) = True
                End If
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
    
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Application.Echo True
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, "Projects import error"
    Resume ExitNow

End Sub

Private Sub cmdImportScreens_Click()
'Copy screens from specified database into current database
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim ErrorMsg As String
    Dim ErrorCount As Integer
    Dim addInManager As New CT_ClsCnlyAddinSupport

#If ccSCR = 1 Then
    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All screens in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "Confirm screens import") = vbYes Then
            
                SQL = "SELECT VPS.SQL, VPS.Notes " & _
                    "FROM MUT_ScreensVersionsPaths AS VP " & _
                    "INNER JOIN MUT_ScreensVersionsPathsSQL AS VPS ON " & _
                    "VP.PathID = VPS.PathID " & _
                    "WHERE '" & SortableVersion(ImportFileVersion) & "' Between SortableVersion(VP.MinSrcVer) AND SortableVersion(MaxSrcVer) " & _
                    "AND '" & SortableVersion(VersionTemplate) & "' BETWEEN SortableVersion(MinDestVer) AND SortableVersion(MaxDestVer) " & _
                    "ORDER BY VPS.Sort"
                
                ErrorCount = 0
                lstScreenImport.RowSource = vbNullString
                
                DoCmd.Hourglass True
                Application.Echo False
                
                Set db = CurrentDb
                Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
    
                Do Until rs.EOF
                    SQL = rs!SQL
                    SQL = Replace(SQL, "$DbName", txtAccessFile)
    
                    If ExecuteScript(db, SQL, ErrorMsg) Then
                        lstScreenImport.AddItem Nz(rs!Notes, "Executed SQL Statement")
                    Else
                        ErrorCount = ErrorCount + 1
                        lstScreenImport.AddItem "*** ERROR!: " & rs!Notes & " - " & ErrorMsg & " ***"
                    End If
        
                    rs.MoveNext
                Loop
    
                'Restore screens that were installed as an app and rebuild if required
                If RestoreInstalledScreens Then
                    addInManager.BuildRibbonBar "BuildSilent"
                    MsgBox "Screens import completed with " & ErrorCount & " errors." & vbCrLf & vbCrLf & _
                        "You will need to restart Decipher to see changes in the Ribbon Bar.", vbInformation, "Screens import complete"
                Else
                    MsgBox "Screens import completed with " & ErrorCount & " errors.", vbInformation, "Screens import complete"
                End If
                
                'Move to end of list
                If lstScreenImport.ListCount > 0 Then
                    lstScreenImport.Selected(lstScreenImport.ListCount - 1) = True
                End If
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
#ElseIf ccSCR = 2 Then
    RunImportConfigurations SCR_AppID, "App Screens", Me.lstScreenImport
    If RestoreInstalledScreens Then
        addInManager.BuildRibbonBar "BuildSilent"
        MsgBox "You will need to restart Decipher to see changes in the Ribbon Bar.", vbInformation, "Screens import complete"
    End If
#End If
    
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Application.Echo True
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, "Screens import error"
    Resume ExitNow
End Sub

Private Function RestoreInstalledScreens() As Boolean
'Restore screens that were installed as an app
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    Dim ScreenName As String
    Dim Result As Boolean
    
    Set db = CurrentDb
    #If ccSCR = 1 Then
        SQL = "SELECT ScreenName FROM SCR_ScreensXML WHERE ProductPrefix<>'' Group By ScreenName"
        Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
        If rs.recordCount > 0 Then
            Do Until rs.EOF
                ScreenName = rs!ScreenName
        
                'Backup existing screen
                SQL = "UPDATE SCR_Screens SET Included=0, ScreenName='" & Replace(ScreenName, "'", "''") & " - Migrated copy (Do not use)" & Now & _
                    "' WHERE ScreenName='" & Replace(ScreenName, "'", "''") & "'"
                db.Execute SQL
                
                'Install screen
                Run "RestoreScreenFromXML", "", "", ScreenName, True
                
                lstScreenImport.AddItem "Restored screen " & ScreenName & " from XML"
                
                DoEvents
                
                rs.MoveNext
            Loop
            Result = True
        End If
    #ElseIf ccSCR = 2 Then
        SQL = "SELECT ConfigName FROM APPF_ManageToXML WHERE AppID = " & SCR_AppID & " AND ProductPrefix<>'' Group By ConfigName"
        Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
        If rs.recordCount > 0 Then
            Do Until rs.EOF
                ScreenName = rs!ConfigName
        
                'Backup existing screen
                SQL = "UPDATE SCR_Screens SET Included=0, ScreenName='" & Replace(ScreenName, "'", "''") & " - Migrated copy (Do not use)" & Now & _
                    "' WHERE ScreenName='" & Replace(ScreenName, "'", "''") & "'"
                db.Execute SQL
                
                'Install screen
                Run "RestoreAppsFromXML", "", "", ScreenName, SCR_AppID
                
                lstScreenImport.AddItem "Restored screen " & ScreenName & " from XML"
                
                DoEvents
                
                rs.MoveNext
            Loop
            Result = True
        End If
    #End If
    
ExitNow:
On Error Resume Next
    RestoreInstalledScreens = Result
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Function
ErrorHappened:
    Result = False
    lstScreenImport.AddItem "*** ERROR!: Restoring screens from XML ***"
    Resume ExitNow
End Function

Private Sub cmdImportLinks_Click()
'Copy ConfigLinks data
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim ErrorMsg As String

    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All link settings in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "Confirm links import") = vbYes Then
                
                DoCmd.Hourglass True
                
                Set db = CurrentDb
                
                'Clear table
                SQL = "DELETE FROM CFG_CfgLink"
                db.Execute SQL, dbFailOnError
                
                'Copy
                SQL = "INSERT INTO CFG_CfgLink" & _
                    "([Location],[LinkType],[Prefix],[Server],[Database],[Suffix],[LastMessage],[Schema]) " & _
                    "SELECT [Location],[LinkType],[Prefix],[Server],[Database],[Suffix],[LastMessage],[Schema] " & _
                    "FROM [" & txtAccessFile & "].[CcaCfgLink]"
    
                If ExecuteScript(db, SQL, ErrorMsg) Then
                    MsgBox "Your link settings have been imported.", vbInformation, "Link import complete"
                Else
                    MsgBox "There was an error importing your link settings." & vbCrLf & ErrorMsg, vbCritical, "Link import error"
                End If
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    Set db = Nothing
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, "Links import error"
    Resume ExitNow
End Sub

Private Sub cmdImportDupTool_Click()
'Copy DupTool data
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim ErrorMsg As String
    
    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All DupTool settings in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "DupTool import") = vbYes Then
                
                DoCmd.Hourglass True
                
                Set db = CurrentDb
                
                'Clear table
                SQL = "DELETE FROM DT_DupCriteriaSavedConfigs"
                db.Execute SQL, dbFailOnError
                
                SQL = "DELETE FROM DT_DupCriteriaSavedConfigsSub1"
                db.Execute SQL, dbFailOnError
                
                'Copy
                SQL = "INSERT INTO DT_DupCriteriaSavedConfigs" & _
                    "([Rpt],[ConfigName],[ConfigWhereClause],[ConfigHavingClause]) " & _
                    "SELECT [Rpt],[ConfigName],[ConfigWhereClause],[ConfigHavingClause] " & _
                    "FROM [" & txtAccessFile & "].[CnlyDtDupCriteriaSavedConfigs] " & _
                    "WHERE ConfigName <> 'test' AND ConfigName not like 'David*s'"
                If ExecuteScript(db, SQL, ErrorMsg) Then
                    SQL = "INSERT INTO DT_DupCriteriaSavedConfigsSub1" & _
                        "([Rpt],[ConfigName],[CriteriaType],[FieldName]) " & _
                        "SELECT [Rpt],[ConfigName],[CriteriaType],[FieldName] " & _
                        "FROM [" & txtAccessFile & "].[CnlyDtDupCriteriaSavedConfigsSub1] " & _
                        "WHERE ConfigName <> 'test' AND ConfigName not like 'David*s'"
                    If ExecuteScript(db, SQL, ErrorMsg) Then
                        MsgBox "Your DupTool settings have been imported.", vbInformation, "DupTool import complete"
                    Else
                        MsgBox "There was an error importing your DupTool settings. " & vbCrLf & ErrorMsg, vbCritical, "DupTool import error"
                    End If
                Else
                    MsgBox "There was an error importing your DupTool settings." & vbCrLf & ErrorMsg, vbCritical, "DupTool import error"
                End If
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    Set db = Nothing
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, "DupTool import error"
    Resume ExitNow
End Sub

Private Function ExecuteScript(ByRef db As DAO.Database, ByVal SQL As String, ByRef ErrMsg As String) As Boolean
'Execute specified script on specified db
On Error GoTo ErrorHappened
    Dim Result As Boolean

    db.Execute SQL, dbFailOnError
    
    Result = True
ExitNow:
On Error Resume Next
    ExecuteScript = Result
Exit Function
ErrorHappened:
    If Err.Number = 3201 Then
        db.Execute SQL
        Result = True
    Else
        Result = False
        ErrMsg = Err.Description
    End If
    Resume ExitNow
End Function

Private Sub cmdImportClaimsPlus_Click()
'Copy ClaimsPlus settings
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim ErrorCount As Integer
    Dim TableName As String
    Dim ErrorMsg As String

    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All ClaimsPlus settings in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "Confirm ClaimsPlus import") = vbYes Then
            
                SQL = "SELECT VPS.SQL, VPS.Notes, VPS.TableName " & _
                    "FROM MUT_CpVersionsPaths AS VP " & _
                    "INNER JOIN MUT_CpVersionsPathsSQL AS VPS ON " & _
                    "VP.PathID = VPS.PathID " & _
                    "WHERE '" & SortableVersion(ImportFileVersion) & "' Between SortableVersion(VP.MinSrcVer) AND SortableVersion(MaxSrcVer) " & _
                    "AND '" & SortableVersion(VersionTemplate) & "' BETWEEN SortableVersion(MinDestVer) AND SortableVersion(MaxDestVer) " & _
                    "ORDER BY VPS.Sort"
                
                ErrorCount = 0
                lstClaimsPlusImport.RowSource = vbNullString
                
                DoCmd.Hourglass True
                
                Set db = CurrentDb
                Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
    
                Do Until rs.EOF
                    TableName = Nz(rs!TableName, vbNullString)
                    
                    'Execute if table name is not specified or specified table exists
                    If LenB(TableName) = 0 Or TableExists(TableName, CurrentDb) Then
                        SQL = rs!SQL
                        SQL = Replace(SQL, "$DbName", txtAccessFile)
        
                        If ExecuteScript(db, SQL, ErrorMsg) Then
                            lstClaimsPlusImport.AddItem Nz(rs!Notes, "Executed SQL Statement")
                        Else
                            ErrorCount = ErrorCount + 1
                            lstClaimsPlusImport.AddItem "*** ERROR!: " & rs!Notes & " - " & ErrorMsg & " ***"
                        End If
                    End If
                    rs.MoveNext
                Loop
                
                'Move to end of list
                If lstClaimsPlusImport.ListCount > 0 Then
                    lstClaimsPlusImport.Selected(lstScreenImport.ListCount - 1) = True
                End If
                
                MsgBox "ClaimsPlus settings import completed with " & ErrorCount & " errors.", vbInformation, "ClaimsPlus import complete"
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, "ClaimsPlus import error"
    Resume ExitNow
End Sub

Private Sub cmdObjectsExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - Imported objects.txt", lstAccessObjects
End Sub

Private Sub cmdScreenExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - Screens Import Log.txt", lstScreenImport
End Sub

Private Sub cmdCpExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - ClaimsPlus Import Log.txt", lstClaimsPlusImport
End Sub

Private Sub cmdDTExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - DupTool Import Log.txt", lstClaimsPlusImport
End Sub

Private Sub cmdProjectExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - Projects Import Log.txt", lstProjectImport
End Sub

Private Sub ExportLogFile(ByVal FileName As String, ByRef Lst As listBox)
'Write screen import log to file
On Error GoTo ErrorHappened
    Dim fso As Object
    Dim fts As Object
    Dim r As Integer
    Dim c As Integer
    Dim temp As String

    If (Not Lst.ColumnHeads And Lst.ListCount > 0) Or (Lst.ColumnHeads And Lst.ListCount > 1) Then
        Application.Echo False
        Set fso = CreateObject("Scripting.Filesystemobject")
        '2=Overwrite, 8=Append
        Set fts = fso.OpenTextFile(FileName, 2, True)
        
        'Header
        fts.WriteLine String(Len(CurrentProject.FullName) + 2, "*")
        fts.WriteLine "* " & Now
        fts.WriteLine "* " & Identity.UserName
        fts.WriteLine "* " & CurrentProject.FullName
        fts.WriteLine String(Len(CurrentProject.FullName) + 2, "*")
        fts.WriteLine
        
        'Write listbox to file
        For r = 0 To Lst.ListCount - 1
            temp = vbNullString
            
            'Build string based on column count
            For c = 0 To Lst.ColumnCount - 1
                temp = temp & Lst.Column(c, r) & vbTab
            Next c
            
            fts.WriteLine temp
        Next r
        
        MsgBox "The log was exported to:" & vbCrLf & FileName, vbInformation, "Log export complete"
    Else
        MsgBox "There isn't any information in the list to export.", vbInformation, "Nothing to export"
    End If

ExitNow:
On Error Resume Next
    fts.Close
    Set fts = Nothing
    Set fso = Nothing
    Application.Echo True
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error exporting log file"
    Resume ExitNow
End Sub

Private Sub RunImportConfigurations(AppID As Long, AppName As String, ByRef Lst As Access.listBox)
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim ErrorMsg As String
    Dim ErrorCount As Integer

    If LenB(ImportFileVersion) > 0 And ImportFileVersion <> "Error" Then
        If IsVersionOK Then
            If MsgBox("All " & AppName & " configurations in this copy of Decipher will be deleted and imported from" & vbCrLf & _
                txtAccessFile & vbCrLf & vbCrLf & "Are you sure you want to proceed?", _
                vbQuestion + vbYesNo, "Confirm " & AppName & " Import") = vbYes Then
            
                SQL = "SELECT VPS.SQL, VPS.Notes " & _
                    "FROM MUT_AppsVersionsPaths AS VP " & _
                    "INNER JOIN MUT_AppsVersionsPathsSQL AS VPS ON " & _
                    "VP.PathID = VPS.PathID " & _
                    "WHERE '" & SortableVersion(ImportFileVersion) & "' Between SortableVersion(VP.MinSrcVer) AND SortableVersion(MaxSrcVer) " & _
                    "AND '" & SortableVersion(VersionTemplate) & "' BETWEEN SortableVersion(MinDestVer) AND SortableVersion(MaxDestVer) " & _
                    "AND VP.AppID = " & AppID & _
                    " ORDER BY VPS.Sort"
                
                ErrorCount = 0
                Lst.RowSource = vbNullString
                
                DoCmd.Hourglass True
                Application.Echo False
                
                Set db = CurrentDb
                Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)
    
                Do Until rs.EOF
                    SQL = rs!SQL
                    SQL = Replace(SQL, "$DbName", txtAccessFile)
    
                    If ExecuteScript(db, SQL, ErrorMsg) Then
                        Lst.AddItem Nz(rs!Notes, "Executed SQL Statement")
                    Else
                        Debug.Print SQL & vbCrLf & vbCrLf
                        ErrorCount = ErrorCount + 1
                        Lst.AddItem "*** ERROR!: " & rs!Notes & " - " & ErrorMsg & " ***"
                    End If
        
                    rs.MoveNext
                Loop
    
                MsgBox AppName & " import completed with " & ErrorCount & " errors.", vbInformation, AppName & " import complete"
                
                'Move to end of list
                If Lst.ListCount > 0 Then
                    Lst.Selected(Lst.ListCount - 1) = True
                End If
            End If
        Else
            MsgBox "The version you selected is too old to be imported." & vbCrLf & "Please update to version " & _
                MinSourceVersion & " or higher before you import.", vbInformation, "Old source version"
        End If
    Else
        MsgBox "Please specify a valid Decipher file to import", vbInformation, "Missing import file"
    End If
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Application.Echo True
Exit Sub
ErrorHappened:
    DoCmd.Hourglass False
    MsgBox Err.Description, vbCritical, AppName & " import error"
    Resume ExitNow
End Sub

Private Sub cmdImportProjectManagement_Click()
#If ccPMT = 1 Then
    RunImportConfigurations PMT_AppID, "Project Management", Me.lstPMTImport
#End If
End Sub

Private Sub cmdImportVendorManagement_Click()
#If ccVMT = 1 Then
    RunImportConfigurations VMT_AppID, "Vendor Management", Me.lstVMTImport
#End If
End Sub

Private Sub cmdImportWorkFiles_Click()
#If ccWFT = 1 Then
    RunImportConfigurations WFT_AppID, "WorkFiles", Me.lstWorkFilesImport
#End If
End Sub

Private Sub cmdPMTExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - Project Management Import Log.txt", lstPMTImport
End Sub

Private Sub cmdVMTExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - Vendor Management Import Log.txt", lstVMTImport
End Sub

Private Sub cmdWorkFilesExportToFile_Click()
    ExportLogFile CurrentProject.Path & "\Migration - WorkFiles Import Log.txt", lstWorkFilesImport
End Sub
