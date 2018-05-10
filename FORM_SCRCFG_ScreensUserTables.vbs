Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_SCRCFG_ScreensUserTables
' Author    : Scott Akam
' Date      : 10/8/2012
' Purpose   : Manage location of screens user tables
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

#If ccSCR = 1 Then      'Decipher Screens
    Private Const AppName As String = "Decipher Screens"
    Private Const DefaultDatabaseNameAccess As String = "<ShortName>DecipherSettings<Maj><Min>.accdb"
    Private Const DefaultDatabaseNameSQL As String = "[<ShortName>AuditorsDecipherSettings]"
#ElseIf ccSCR = 2 Then  'App Screens
    Private Const AppName As String = "App Screens"
    Private Const DefaultDatabaseNameAccess As String = "<ShortName>AppScreensSettings<Maj><Min>.accdb"
    Private Const DefaultDatabaseNameSQL As String = "[<ShortName>AuditorsAppScreensSettings]"
#End If

Private CurrentTableLocation As UserTableLocation
Private CurrentConnectString As String
Private BackupDbName As String
Private SqlScriptOpened As Boolean

Private Enum UserTableLocation
    Unknown = 0
    AccessLocal = 1
    AccessLinked = 2
    SQLLinked = 3
End Enum

Private Sub Form_Load()
'Form load
On Error GoTo ErrorHappened
    FillUserTableCombo
    SetCurrentLocationLabel
    
    cmdMoveTables.Enabled = False
    
    txtAccessFile.top = cboSQLServer.top
    cmdBrowseFile.top = cboSQLServer.top
    lblAccessFile.top = lblSQLServer.top
    
    lblSQLScriptLink.HyperlinkAddress = ScriptCreateDbFile
    
    If IsProductInstalled("Config Links") Then
        cboSQLServer.RowSource = "SELECT Server FROM CFG_CfgLink GROUP BY Server ORDER BY Server"
    End If
        
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:Form_Load"
    Resume ExitNow
    Resume
End Sub

Private Sub FillUserTableCombo()
'Fill user table combo and select current location
On Error GoTo ErrorHappened
    Dim TableLocation As UserTableLocation
    
    With cboUserTablesLocation
        .RowSource = vbNullString
        .AddItem UserTableLocation.AccessLocal & ";Local Access Tables"
        .AddItem UserTableLocation.AccessLinked & ";Linked Access Tables"
        .AddItem UserTableLocation.SQLLinked & ";Linked SQL Server Tables"
    End With
    
    TableLocation = GetCurrentTableLocation
    CurrentTableLocation = TableLocation
    If TableLocation > UserTableLocation.Unknown Then
        cboUserTablesLocation.Value = TableLocation
    End If
    
    Select Case CurrentTableLocation
        Case UserTableLocation.AccessLinked
            txtAccessFile = Right(CurrentConnectString, Len(CurrentConnectString) - InStr(CurrentConnectString, "="))
        Case UserTableLocation.SQLLinked
            cboSQLServer = GetServerNameFromConnectionString(CurrentConnectString)
            txtSQLDatabase = GetDatabaseNameFromConnectionString(CurrentConnectString)
    End Select
    
    ToggleMode
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:FillUserTableCombo"
    Resume ExitNow
    Resume
End Sub

Private Function GetCurrentTableLocation() As UserTableLocation
'Figure out where the user tables are
On Error GoTo ErrorHappened
    Const TableName As String = "SCR_SaveScreens"
    Dim db As DAO.Database
    Dim Result As UserTableLocation
    
    Set db = CurrentDb
    
    If TableExists(TableName, db) Then
        CurrentConnectString = db.TableDefs(TableName).Connect

        If LenB(CurrentConnectString) = 0 Then
            Result = UserTableLocation.AccessLocal
        ElseIf Right(CurrentConnectString, 6) = ".Accdb" Then
            Result = UserTableLocation.AccessLinked
        ElseIf InStr(1, CurrentConnectString, "DRIVER=SQL Server;") > 0 Then
            Result = UserTableLocation.SQLLinked
        Else
            Result = UserTableLocation.Unknown
        End If
        
    Else
        Result = UserTableLocation.Unknown
    End If
ExitNow:
On Error Resume Next
    Set db = Nothing
    GetCurrentTableLocation = Result
Exit Function
ErrorHappened:
    Result = UserTableLocation.Unknown
    Resume ExitNow
    Resume
End Function

Private Sub cboUserTablesLocation_Change()
    ToggleMode
End Sub

Private Sub ToggleMode()
'Toggle items on form based on selection
On Error GoTo ErrorHappened

    'Defaults
    lblSQLServer.visible = False
    cboSQLServer.visible = False
    lblSQLDatabase.visible = False
    txtSQLDatabase.visible = False
    lblAccessFile.visible = False
    txtAccessFile.visible = False
    cmdBrowseFile.visible = False
    lblSQLScriptLink.visible = False
    
    Select Case cboUserTablesLocation.Value
        Case UserTableLocation.AccessLinked
            lblAccessFile.visible = True
            txtAccessFile.visible = True
            cmdBrowseFile.visible = True
        Case UserTableLocation.SQLLinked
            lblSQLServer.visible = True
            cboSQLServer.visible = True
            lblSQLDatabase.visible = True
            txtSQLDatabase.visible = True
            lblSQLScriptLink.visible = True
            If LenB(Nz(txtSQLDatabase, vbNullString)) = 0 Then
                txtSQLDatabase = GetShortNameFromFilePath & "AuditorsDecipherSettings"
            End If
    End Select
    
    If CurrentTableLocation = UserTableLocation.Unknown Then
        cboUserTablesLocation.Enabled = False
    Else
        cboUserTablesLocation.Enabled = True
    End If
    
    If CurrentTableLocation = cboUserTablesLocation.Value Then
        cmdMoveTables.Enabled = False
        txtAccessFile.Enabled = False
    Else
        If CurrentTableLocation = UserTableLocation.SQLLinked Then
            cmdMoveTables.Enabled = False
            txtAccessFile.Enabled = False
        Else
            cmdMoveTables.Enabled = True
            txtAccessFile.Enabled = True
        End If
    End If
    
    If Not UserTablesVersionOK And CurrentTableLocation > UserTableLocation.Unknown Then
        cmdUpdateTableVersion.visible = True
        lblVersionOutOfSync.visible = True
    Else
        cmdUpdateTableVersion.visible = False
        lblVersionOutOfSync.visible = False
    End If
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:ToggleMode"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdBrowseFile_Click()
'Browse for file
On Error GoTo ErrorHandler
    Dim DefaultFileName As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Specify file you want your user tables stored in"
        .AllowMultiSelect = False
        
        DefaultFileName = Replace(DefaultDatabaseNameAccess, "<ShortName>", GetShortNameFromFilePath)
        DefaultFileName = Replace(DefaultFileName, "<Maj>", CT_GetAppVersionMaj(AppName))
        DefaultFileName = Replace(DefaultFileName, "<Min>", CT_GetAppVersionMin(AppName))
        
        .InitialFileName = CurrentProject.Path & "\" & DefaultFileName
        
        If .show Then
            txtAccessFile = .SelectedItems(1)
            If Right(txtAccessFile, 6) <> ".accdb" Then
                txtAccessFile = txtAccessFile & ".accdb"
            End If
        End If
    End With
ExitNow:
On Error Resume Next
    Set fd = Nothing
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error loading app file"
    Resume ExitNow
End Sub

Private Sub SetCurrentLocationLabel()
'Fill out current location label
On Error GoTo ErrorHappened
    Dim LableText As String
    
    Select Case CurrentTableLocation
        Case UserTableLocation.Unknown
            LableText = "Unknown - Make sure your database is linked."
        Case UserTableLocation.AccessLocal
            LableText = cboUserTablesLocation.Column(1, 0)
        Case UserTableLocation.AccessLinked
            LableText = cboUserTablesLocation.Column(1, 1)
            LableText = LableText & ", File: " & Right(CurrentConnectString, Len(CurrentConnectString) - InStr(CurrentConnectString, "="))
        Case UserTableLocation.SQLLinked
            LableText = cboUserTablesLocation.Column(1, 2)
            LableText = LableText & ", Server: " & GetServerNameFromConnectionString(CurrentConnectString)
            LableText = LableText & ", Database: " & GetDatabaseNameFromConnectionString(CurrentConnectString)
    End Select
    
    lblCurrentLocation.Caption = LableText
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:SetCurrentLocationLabel"
    Resume ExitNow
    Resume
End Sub

Private Sub cmdMoveTables_Click()
'Move tables based on selection
On Error GoTo ErrorHappened
    BackupDbName = CT_BackupDatabase
    
    DoCmd.Hourglass True
    
    If LenB(BackupDbName) > 0 Then
        Select Case cboUserTablesLocation.Value
            Case UserTableLocation.AccessLocal
                ImportTablesAccessToLocal
            Case UserTableLocation.AccessLinked
                ExportTablesToAccess
            Case UserTableLocation.SQLLinked
                MoveTablesToSQLServer
        End Select
        
        Application.RefreshDatabaseWindow
        
        Telemetry.RecordAction "Moved user table location", "<TBLLOC>" & cboUserTablesLocation & "</TBLLOC>", AppName
    Else
        DoCmd.Hourglass False
        MsgBox "Failed to make a backup of the database file.", vbCritical, "Aborting!"
    End If
    
ExitNow:
On Error Resume Next
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:cmdMoveTables_Click"
    Resume ExitNow
    Resume
End Sub

Private Function UserTablesList() As Collection
'List of user tables for import/export
    #If ccSCR = 1 Then
        Set UserTablesList = SCR_ScreensUserTablesList
    #ElseIf ccSCR = 2 Then
        Set UserTablesList = SCR_AppScreensUserTablesList
    #End If
End Function

Private Sub ImportTablesAccessToLocal()
'Pull tables in from linked Access DB
'Note: delete, import and copy relationships needs to be seperate loops
On Error GoTo ErrorHappened

    Dim i As Integer
    Dim TableCol As New Collection
    Dim AccessFile As String
    
    AccessFile = Right(CurrentConnectString, Len(CurrentConnectString) - InStr(CurrentConnectString, "="))

    If LenB(AccessFile) > 0 Then
        
        Set TableCol = UserTablesList

        'Delete
        RemoveLocalUserTables
        
        'Import tables
        For i = 1 To TableCol.Count
            DoCmd.TransferDatabase acImport, "Microsoft Access", AccessFile, acTable, TableCol.Item(i), TableCol.Item(i)
        Next
        
        'Copy relationships
        For i = 1 To TableCol.Count
            CopyTableRelationships TableCol.Item(i), AccessFile, CurrentDb.Name
        Next
        
        FillUserTableCombo
        SetCurrentLocationLabel
        
        MsgBox "User tables were successfully imported from:" & vbCrLf & AccessFile & vbCrLf & vbCrLf & _
            "A backup of your original file was created:" & vbCrLf & BackupDbName, vbInformation, "Success"
    Else
        MsgBox "Please make sure your user tables are linked.", vbExclamation, "Missing linked tables"
    End If
ExitNow:
On Error Resume Next
    Set TableCol = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:ImportTablesToLocal"
    Resume ExitNow
    Resume
End Sub

Private Sub ExportTablesToAccess()
'Create blank database and move tables
'Note: Export, copy relationships and delete needs to be seperate loops
On Error GoTo ErrorHappened
    Dim App As New Access.Application
    Dim i As Integer
    Dim TableCol As New Collection
    
    If LenB(txtAccessFile) > 0 Then
        'Check for and remove target
        If LenB(Dir(txtAccessFile, vbNormal)) > 0 Then
            Kill txtAccessFile
        End If

        'Create file
        App.DBEngine.CreateDatabase txtAccessFile, dbLangGeneral

        Set TableCol = UserTablesList

        'Export tables
        For i = 1 To TableCol.Count
            DoCmd.TransferDatabase acExport, "Microsoft Access", txtAccessFile, acTable, TableCol.Item(i), TableCol.Item(i)
        Next
        
        'Copy relationships
        For i = 1 To TableCol.Count
            CopyTableRelationships TableCol.Item(i), CurrentDb.Name, txtAccessFile
        Next

        'Delete
        RemoveLocalUserTables
        
        'Link to new database
        LinkTables vbNullString, txtAccessFile
        
        'Add to config links table
        UpdateCfgLinkData "ACCESS", vbNullString, txtAccessFile
        
        FillUserTableCombo
        SetCurrentLocationLabel
        
        MsgBox "User tables were successfully exported to:" & vbCrLf & txtAccessFile & vbCrLf & vbCrLf & _
            "A backup of your original file was created:" & vbCrLf & BackupDbName, vbInformation, "Success"
    Else
        MsgBox "Please select a target ACCDB file.", vbExclamation, "Missing ACCDB file name"
    End If
ExitNow:
On Error Resume Next
    Set App = Nothing
    Set TableCol = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:ExportTablesToAccess"
    Resume ExitNow
    Resume
End Sub

Private Sub MoveTablesToSQLServer()
'Create blank database and move tables
'Note: Export, copy relationships and delete needs to be seperate loops
On Error GoTo ErrorHappened
    Dim genUtils As New CT_ClsGeneralUtilities
    
    If LenB(cboSQLServer) > 0 And LenB(txtSQLDatabase) > 0 Then
        If SqlScriptOpened Then
            'Delete local tables
            RemoveLocalUserTables
            
            'Link to new database
            LinkTables cboSQLServer, txtSQLDatabase
            
            If TableExists(UserTablesList.Item(1), CurrentDb) Then
                genUtils.CreatePK "SCR_SaveScreens", "ScreenID,UserName"
        
                'Add to config links table
                UpdateCfgLinkData "SQL", cboSQLServer, txtSQLDatabase
                
                FillUserTableCombo
                SetCurrentLocationLabel
         
                If IsProductInstalled("Decipher Screens Sync") Then
                    'Copy data from backup to SQL
                    If MsgBox("You are now linked into the SQL database:" & vbCrLf & _
                        "Server: " & cboSQLServer & ", Database: " & txtSQLDatabase & vbCrLf & vbCrLf & _
                        "A backup of your original file was created:" & vbCrLf & BackupDbName & vbCrLf & vbCrLf & _
                        "Would you like to export your user settings to SQL Server?", vbQuestion + vbYesNo, "Export settings") = vbYes Then
                        
                        DoCmd.OpenForm "SCRSYNC_Sync", acNormal, , , , , BackupDbName
                    End If
                Else
                    MsgBox "You are now linked into the SQL database:" & vbCrLf & _
                        "Server: " & cboSQLServer & ", Database: " & txtSQLDatabase & vbCrLf & vbCrLf & _
                        "A backup of your original file was created:" & vbCrLf & BackupDbName & vbCrLf & vbCrLf & _
                        "If you want to copy you user settings over, install Decipher Screens Sync and run the utility.", vbInformation, "Success"
                End If
            Else
                MsgBox "There was a problem linking to your database. Please confirm your server and database settings and try again.", vbCritical, "Link error"
            End If
        Else
            If MsgBox("Have you created the User Table database on your server?" & vbCrLf & vbCrLf & _
                "If not select NO and click the link below to open the database creation script.", _
                vbQuestion + vbYesNo, "Database created?") = vbYes Then
                
                SqlScriptOpened = True
                MoveTablesToSQLServer
            End If
        End If
    Else
        MsgBox "Please select a server and database.", vbExclamation, "Missing information"
    End If
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:ExportTablesToAccess"
    Resume ExitNow
    Resume
End Sub

Private Sub CopyTableRelationships(ByVal TableName As String, ByVal SourceDB As String, ByVal DestDB As String)
'Copy table relationships
On Error GoTo ErrorHappened
    Dim DbDest As DAO.Database
    Dim DbSource As DAO.Database
    Dim RelSource As Relation
    Dim RelTarget As Relation
    Dim i As Integer
    
    Set DbSource = DBEngine.OpenDatabase(SourceDB)
    Set DbDest = DBEngine.OpenDatabase(DestDB)
    
    DbSource.Relations.Refresh

    For Each RelSource In DbSource.Relations
        If RelSource.ForeignTable = TableName Then
            Set RelTarget = DbDest.CreateRelation(RelSource.Name, RelSource.Table, TableName, RelSource.Attributes)
            For i = 0 To RelSource.Fields.Count - 1
                RelTarget.Fields.Append RelTarget.CreateField(RelSource.Fields(i).Name)
                RelTarget.Fields(i).ForeignName = RelSource.Fields(i).ForeignName
            Next i
            DbDest.Relations.Append RelTarget
        End If
    Next RelSource

ExitNow:
On Error Resume Next
    DbDest.Close
    DbSource.Close
    Set DbDest = Nothing
    Set DbSource = Nothing
    Set RelSource = Nothing
    Set RelTarget = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
End Sub

Private Sub LinkTables(ByVal Server As String, ByVal Database As String)
'Call config links "API" to link tables to Access or SQL
On Error GoTo ErrorHappened
    Dim cfg As New Form_CFG_CfgLink
    With cfg
        .visible = False
        If LenB(Server) > 0 And LenB(Database) > 0 Then
           .LinkThisDatabase Server, Database, vbNullString, vbNullString, vbNullString
        Else
            .LinkAccessDatabase Database, vbNullString, vbNullString, True
        End If
        
    End With
ExitNow:
On Error Resume Next
    Set cfg = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub UpdateCfgLinkData(ByVal LinkType As String, ByVal Server As String, ByVal Database As String)
'Add a record into config links for each location
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim LastLinkLocation As String
    Dim ServerSave As String
    Dim SQL As String
    
    Set db = CurrentDb
    LastLinkLocation = db.Properties("LastLinkLocation")
    
    SQL = "SELECT [Location] FROM CFG_CfgLink GROUP BY [Location]"
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
    
    Do Until rs.EOF
        If LinkType = "ACCESS" Then
            SQL = "Location='" & rs!Location & "' AND LinkType='" & LinkType & "' AND Database='" & Replace(Database, "'", "''") & "'"
        Else
            'Try to find server for given location.
            'Note: This may be wrong if they have more than 1 server for a location.
            If LastLinkLocation <> rs!Location Then
                ServerSave = Nz(DLookup("Server", "CFG_CfgLink", "LinkType='SQL' AND Location='" & rs!Location & "'"), Server)
            Else
                ServerSave = Server
            End If
            
            SQL = "Location='" & rs!Location & "' AND LinkType='" & LinkType & "' AND Server='" & Replace(ServerSave, "'", "''") & "' AND Database='" & Replace(Database, "'", "''") & "'"
        End If
        
        'Insert row if not found
        If DCount("Location", "CFG_CfgLink", SQL) = 0 Then
            SQL = "INSERT INTO CFG_CfgLink ([Location],[LinkType],[Server],[Database])VALUES" & _
                  "('" & Replace(rs!Location, "'", "''") & "','" & LinkType & "','" & _
                  Replace(ServerSave, "'", "''") & "','" & Replace(Database, "'", "''") & "')"

            db.Execute SQL, dbFailOnError
        End If

        rs.MoveNext
    Loop

ExitNow:
On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:UpdateCfgLinkData"
    Resume ExitNow
    Resume
End Sub

Private Sub RemoveLocalUserTables()
'Remove local user tables
On Error GoTo ErrorHappened
    Dim TableCol As Collection
    Dim i As Integer
    
    Set TableCol = UserTablesList

    For i = 1 To TableCol.Count
        DeleteTable TableCol.Item(i)
    Next
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub DeleteTable(ByVal TableName As String)
'Delete table and fail silently if it doesn't exist
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Set db = CurrentDb
    
    db.TableDefs.Delete TableName
ExitNow:
On Error Resume Next
    Set db = Nothing
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub lblSQLScriptLink_Click()
    cboUserTablesLocation.SetFocus
    lblSQLScriptLink.HyperlinkAddress = CopySQLScript(ScriptCreateDbFile)
    SqlScriptOpened = True
End Sub

Private Function UserTablesVersionOK() As Boolean
'Check to see if use table and config table versions are in sync
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim TableVersionConfig As Integer
    Dim TableVersionUser As Integer
    
    If Not CurrentTableLocation = UserTableLocation.Unknown Then
        TableVersionConfig = Nz(DLookup("VersionNum", "SCR_TablesVersionConfig"), 0)
        TableVersionUser = Nz(DLookup("VersionNum", "SCR_TablesVersionUser"), 0)
        
        If TableVersionConfig = TableVersionUser Then
            Result = True
        End If
    End If
ExitNow:
On Error Resume Next
    UserTablesVersionOK = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function

Private Sub cmdUpdateTableVersion_Click()
'Update user tables based on table type
On Error GoTo ErrorHappened
     
    Select Case CurrentTableLocation
        Case UserTableLocation.AccessLocal
        
        Case UserTableLocation.AccessLinked
        
        Case UserTableLocation.SQLLinked
            Application.FollowHyperlink CopySQLScript(ScriptUpdateDbFile)
    End Select

ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_SCRCFG_ScreensUserTables:cmdUpdateTableVersion_Click"
    Resume ExitNow
    Resume
End Sub

Private Function ScriptCreateDbFile() As String
'Get path to SQL server create DB script
On Error GoTo ErrorHappened
    Dim Result As String
    Dim FileName As String
    
    FileName = DLookup("Setting", "SCR_Settings", "Tag='UserTablesSqlCreateScript'")
    FileName = Replace(FileName, "{VERSION}", DLookup("VersionNum", "SCR_TablesVersionConfig"))
    
    Result = FileName
ExitNow:
On Error Resume Next
    ScriptCreateDbFile = Result
Exit Function
ErrorHappened:
    Result = vbNullString
    Resume ExitNow
End Function

Private Function ScriptUpdateDbFile() As String
'Get path to SQL server update DB script
On Error GoTo ErrorHappened
    Dim Result As String
    Dim FileName As String
    
    FileName = DLookup("Setting", "SCR_Settings", "Tag='UserTablesSqlUpdateScript'")
    FileName = Replace(FileName, "{VERSION}", DLookup("VersionNum", "SCR_TablesVersionConfig"))
    
    Result = FileName
ExitNow:
On Error Resume Next
    ScriptUpdateDbFile = Result
Exit Function
ErrorHappened:
    Result = vbNullString
    Resume ExitNow
End Function

Private Function GetShortNameFromFilePath() As String
'Get the shortname from the current db path
On Error GoTo ErrorHappened
    Dim Result As String
    Dim Default As String
    Dim SearchString As String
    Dim AuditPos As Integer
    
    Default = "<ShortName>"
    Result = CurrentDb.Name
    SearchString = "\Audits\"
    AuditPos = InStr(Result, SearchString)
    
    If AuditPos > 0 Then
        Result = Right(Result, Len(Result) - AuditPos - Len(SearchString) + 1)
        Result = left(Result, InStr(Result, "\") - 1)
    Else
        Result = Default
    End If
ExitNow:
On Error Resume Next
    GetShortNameFromFilePath = Result
Exit Function
ErrorHappened:
    Result = Default
    Resume ExitNow
End Function

Private Function CopySQLScript(ByVal FileNameMaster As String) As String
'Copy script to local directory and replace the database name
On Error GoTo ErrorHappened
    Const ForReading = 1
    Const ForWriting = 2
    
    Dim fso As Object 'FileSystemObject
    Dim file As Object
    Dim FileText As String
    Dim FileNameLocal As String
    Dim DatabaseName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Copy file in from Source folder
    FileNameLocal = CurrentProject.Path & "\" & fso.GetFileName(FileNameMaster)
    If fso.FileExists(FileNameLocal) Then
        fso.DeleteFile FileNameLocal, True
    End If
    fso.CopyFile FileNameMaster, FileNameLocal, True
    
    'Read file contents
    Set file = fso.OpenTextFile(FileNameLocal, ForReading)
    FileText = file.ReadAll
    file.Close
    
    'Replace database name in file and write to new sql file
    If LenB(Nz(txtSQLDatabase, vbNullString)) > 0 Then
        DatabaseName = "[" & txtSQLDatabase & "]"
    Else
        DatabaseName = "[" & GetShortNameFromFilePath & "AuditorsDecipherSettings]"
    End If
    
    FileText = Replace(FileText, DefaultDatabaseNameSQL, DatabaseName)
    
    'Create new file to prevent problems with read only files.
    If fso.FileExists(FileNameLocal) Then
        fso.DeleteFile FileNameLocal, True
    End If
    fso.CreateTextFile FileNameLocal, True
    
    Set file = fso.OpenTextFile(FileNameLocal, ForWriting)
    file.WriteLine FileText
    
ExitNow:
On Error Resume Next
    file.Close
    Set file = Nothing
    Set fso = Nothing
    CopySQLScript = FileNameLocal
Exit Function
ErrorHappened:
    CopySQLScript = vbNullString
    Resume ExitNow
    Resume
End Function

Private Sub cmdCleanUserTables_Click()
    Dim Msg As String
    
    Msg = "This will delete any user data that is not associated with a Screen." & vbCrLf & vbCrLf & "Would you like to continue?"
    
    If MsgBox(Msg, vbQuestion + vbYesNo, "Clean user data") = vbYes Then
        SCR_CleanUserTables
    End If
End Sub
