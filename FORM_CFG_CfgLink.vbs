Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 6/26/2012 - Replaced table name with constant and updated import procedure

Private Const CfgTableName As String = "CFG_CfgLink"

Public Enum LinkType
    LNK_ACCESS = 0
    LNK_SQL = 1
End Enum

Private Type LinkCnf
    LinkType As LinkType
    Server As String
    Database As String
    Schema As String
    Connect As String
    Prefix As String
    Suffix As String
    UseViews As Boolean
    TableName As String
    DropIfExists As Boolean
End Type

Private MvarAutoGetServers As Boolean
Private MvarAutoDcServers As Boolean
Private MvarAutoFldServers As Boolean

Private MvMasterLocations As Collection
Private mvCancel As Boolean
Private mvLink As LinkCnf
Private Const QI As String = """"

' HC 5/2010 -- added constants for the property values saved and retrieved
Private Const PROPERTYLINK As String = "LastLinkLocation"
Private Const PROPERTYAUTOSERVERS As String = "AutoGetServers"
Private Const PROPERTYDCSERVERS As String = "AutoDcServers"
Private Const PROPERTYFLDSERVERS As String = "AutoFldServers"

Private Const DCSERVERCLASS As String = "DC"
Private Const FLDSERVERCLASS As String = "Field"

Private Const LINK_SRC_ACCESS = "Provider=Microsoft.ACE.OLEDB.12.0;"
Private Const LINK_SRC_SQL = "Provider=SQLOLEDB.1;Integrated Security='SSPI';"
Private Const ConfigUrl As String = "http://util.svc.ccaintranet.com/smo/default.asmx"
Private Const adOpenStatic = 3
Private Const adUseClient = 3
Private Const adLockBatchOptimistic = 4

Public Event TableLinked(LinkType As String, RemoteName As String, LocalName As String)
Public Event LinkLocationChanged()

'WINDOWS API STUFF
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Properties
Public Property Get MasterLocations() As Collection
    'Get a list (collection) of available link locations.
    'for use with LinkAll

    Dim i As Integer
    Me.MasterLocation.Requery
    
    'Clear the collection
    For i = 1 To MvMasterLocations.Count
        MvMasterLocations.Remove (1)
    Next i
    
    For i = 0 To MasterLocation.ListCount - 1
        MvMasterLocations.Add MasterLocation.ItemData(i)
    Next i
    
    Set MasterLocations = MvMasterLocations
End Property
Public Property Let AutoGetServer(data As Boolean)
    ' HC 5/2010 changed the Auto Get Servers feature to use the database properites rather than the registry
    MvarAutoGetServers = data
    SetProperty PROPERTYAUTOSERVERS, CStr(MvarAutoGetServers)
End Property
Public Property Get AutoGetServer() As Boolean
    AutoGetServer = MvarAutoGetServers
End Property
Public Property Let AutoDcServers(data As Boolean)
' HC 5/2010 -- added the saving of the auto dc servers setting to the database properties
    MvarAutoDcServers = data
    SetProperty PROPERTYDCSERVERS, CStr(MvarAutoDcServers)
End Property
Public Property Get AutoDcServer() As Boolean
    AutoDcServer = MvarAutoDcServers
End Property
Public Property Let AutoFldServers(data As Boolean)
' HC 5/2010 -- added the saving of the auto fld servers setting to the database properties
    MvarAutoFldServers = data
    SetProperty PROPERTYFLDSERVERS, CStr(MvarAutoFldServers)
End Property
Public Property Get AutoFldServer() As Boolean
    AutoFldServer = MvarAutoFldServers
End Property

Public Property Let Cancel(Value As Boolean)
    mvCancel = Value
End Property
Public Property Get Cancel() As Boolean
    Cancel = mvCancel
End Property

Private Sub Form_Load()
    ' HC 3/2010 - Updated function during upgrade to support schemas.
    Dim tableVersion As String
    Dim formVersion As String
    Dim Msg As String
    Dim Result As String
    
    ' HC 3/2010 - Added routine to update the title so it could be updated when it changes
    CreateHeaderTitle
    Me.visible = True
    ClearMessages
    DoEvents
    
    'DLC 05/19/10 Set the default Hight and Width for the form
    Me.Form.InsideHeight = Me.FormFooter.Height + Me.FormHeader.Height + (Me.Detail.Height * 12)
    Me.Form.InsideWidth = Me.CmdLink.left + Me.CmdLink.Width + 400
    
    'HC - 3/2010 - Check the form version against the table version make sure they match before continuing
    tableVersion = GetTableVersion
    formVersion = GetFormVersion
    If tableVersion <> formVersion Then
        Msg = "Incorrect version of the Config Links Form." & vbCrLf & vbCrLf
        Msg = Msg & "Correct version is: " & tableVersion & vbCrLf
        Msg = Msg & "Your form version is: " & formVersion & vbCrLf & vbCrLf
        Msg = Msg & "Check the table description property for version info."
        MsgBox Msg, vbCritical, Me.Name & " - " & CodeContextObject.Name
        DoCmd.Close acForm, Me.Name, acSavePrompt
        Exit Sub
    End If
        
    ' HC 5/2010 -- changed over the AutoGetServers to use the db properties rather than the registry
    MvarAutoGetServers = GetBoolProperty(PROPERTYAUTOSERVERS)
    Me.ChkGetServers = MvarAutoGetServers
    
    ' HC 5/2010 -- added settings for DC Servers and Fld servers
    MvarAutoDcServers = GetBoolProperty(PROPERTYDCSERVERS)
    MvarAutoFldServers = GetBoolProperty(PROPERTYFLDSERVERS)
    
    ' if both are off then set the dc servers on
    If MvarAutoDcServers = False And MvarAutoFldServers = False Then
        AutoDcServers = True
        MvarAutoDcServers = True
    End If
    
    Me.chkDc = MvarAutoDcServers
    Me.chkFld = MvarAutoFldServers
    
    'populate the server list
    If MvarAutoGetServers = True Then
        Call SetSQLCombo(Me.CmboServer)
    End If
    
   
    'SELECT THE LAST LINKED LOCATION AS CURRENT LINK LOCATION. HC 5/2010 - modified to use the generic get property
    Result = GetProperty(PROPERTYLINK)
    If MasterLocation.ListCount > 0 Then
        If Nz(Result, "") <> "" Then
            Me.MasterLocation = Result
        Else 'No last linked location
            Me.MasterLocation = Me.MasterLocation.ItemData(Me.MasterLocation.ListCount - 1)
        End If
        MasterLocation_AfterUpdate
    Else
        If Nz(Result, "") = "" Then
            SavedLocationSet "Not Linked"
        End If
        Me.MasterLocation = "Not Linked"
    End If
    
    ' HC 5/2010 -- fixed the initial display of the items linked when the last location is Not Linked.
    If MasterLocation.ListIndex <= 0 Then
        CmboLocation.DefaultValue = Me.MasterLocation
        Me.filter = "Location = " & QI & Me.MasterLocation & QI
        Me.FilterOn = True
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MvMasterLocations = Nothing
End Sub
Private Sub Form_AfterDelConfirm(Status As Integer)
    If Status = acDeleteOK Then
        Form_AfterUpdate
    End If
    Me.MasterLocation.Requery
    Me.CmboLocation.Requery
End Sub
Private Sub Form_AfterUpdate()
    Me.MasterLocation.Requery
    Me.CmboLocation.Requery
End Sub
Private Sub Form_Current()
    If Me.NewRecord = True Then
        Me.CmboLocation = Nz(Me.MasterLocation, "Development")
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
    Case 3314 'MISSING FIELD
        On Error GoTo Done
        DoCmd.RunCommand acCmdUndo
        Response = 0
    Case 2169
        On Error GoTo Done
        DoCmd.RunCommand acCmdUndo
        Response = 0
    Case 3022
        'Attempt to add duplicate, ignore and allow the message to be displayed
    Case Else
        MsgBox DataErr
    End Select
Done:
End Sub

Private Sub MasterLocation_AfterUpdate()
On Error GoTo ErrorStuff
    
    If Me.NewRecord = True Then
        'DLC 05/19/10 Use ListIndex check if a combo item was selected
        If Me.CmboLocation.ListIndex < 0 Or Me.CmboDatabase.ListIndex < 0 Then
            If Me.Dirty Then DoCmd.RunCommand (acCmdUndo)
        Else
            Me.Dirty = False
        End If
    End If
    
    If Me.MasterLocation.ListIndex >= 0 Then
        CmboLocation.DefaultValue = Me.MasterLocation
        Me.filter = "Location = " & QI & Me.MasterLocation & QI
        Me.FilterOn = True
    End If
    
    ' HC 3/2010 - update the title when the default locations change
    CreateHeaderTitle
    Exit Sub
    
ErrorStuff:
    MsgBox Err.Description & " -- Master Location After Update"
End Sub
Private Sub MasterLocation_Click()
    Me.MasterLocation.Requery
    Me.CmboLocation.Requery
    UpdateStatus ("")
End Sub
Private Sub ChkGetServers_AfterUpdate()
    Me.AutoGetServer = Me.ChkGetServers
End Sub

Private Sub chkDc_AfterUpdate()
    AutoDcServers = chkDc
End Sub

Private Sub chkFld_AfterUpdate()
    AutoFldServers = chkFld
End Sub
Private Sub cmdBrowse_Click()
On Error GoTo ErrorHandler
    Dim FileName As String

    UpdateStatus ("")
    Select Case Nz(Me.TypeLink, "")
        Case "SQL"
            Dim StServer As String
            StServer = "" & Me.CmboServer
            If Me.TypeLink = "SQL" And StServer <> "" Then
                If Me.CmboDatabase.Tag <> StServer Then
                    screen.MousePointer = 11
                    Call GetSQLDatabases(Me.CmboServer, Me.CmboDatabase)
                    screen.MousePointer = 0
                    Me.CmboDatabase.Tag = StServer
                End If
            Else
                Me.CmboDatabase.Tag = ""
                Me.CmboDatabase.RowSource = ""
            End If
        Case "ACCESS"
            FileName = ShowOpenFile(, , "" & Me.CmboDatabase)
            If FileName <> "" Then
                Me.CmboDatabase = FileName
            End If
        Case Else
            MsgBox "Please select one of the supported link types!"
    End Select
Exit_ErrorHandler:
    On Error Resume Next
    screen.MousePointer = 0
    Exit Sub
ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Exit_ErrorHandler
End Sub
Private Sub CmdSchemaBrowse_Click()
On Error GoTo ErrorHandler
    ' fill the schema list
    Dim StServer As String
    Dim StDatabase As String
    
    UpdateStatus ("")
    
    If Nz(TypeLink, "") = "SQL" Then
        StServer = Nz(Me.CmboServer, "")
        StDatabase = Nz(Me.CmboDatabase, "")
        If StServer <> "" And StDatabase <> "" Then
            If Me.CmboSchema.Tag <> StDatabase Then
                screen.MousePointer = 11
                Call GetSQLDatabaseSchemas(Me.CmboServer, Me.CmboDatabase, Me.CmboSchema)
                screen.MousePointer = 0
                Me.CmboSchema.Tag = StDatabase
            End If
        Else
            Me.CmboSchema.Tag = ""
            Me.CmboSchema.RowSource = ""
        End If
    Else
        MsgBox "Schema is supported only for link types of SQL"
    End If
Exit_ErrorHandler:
    On Error Resume Next
    screen.MousePointer = 0
    Exit Sub
ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Exit_ErrorHandler
End Sub
Private Sub CmdCopyLoc_Click()
On Error GoTo ErrorHappened
    Dim StLoc As String
    Dim SQL As String
    Dim StLocNew As String
    
    UpdateStatus ("")

    'GET THE CURRENT LOCATION
    StLoc = Nz(Me.MasterLocation, "")
    If StLoc = "" Then 'IF NOT SPECIFIED THEN ERROR OUT
        MsgBox "You must have a location specified to create a copy.", vbCritical, "Error Copying Location"
    Else
        'GET THE NAME OF THE NEW LOCATION
        StLocNew = InputBox("Please enter the name of the new location.", "Copy Location", "New Location")
        If Nz(StLocNew, "") <> "" Then
            SQL = "INSERT INTO " & CfgTableName & " (Location, LinkType, Prefix, Server, [Database], Suffix, LastMessage, Schema ) "
            SQL = SQL & "SELECT '" & StLocNew & "' AS Location, LinkType, Prefix, Server, Database, Suffix, LastMessage, Schema "
            SQL = SQL & "FROM " & CfgTableName & " "
            SQL = SQL & "WHERE Location ='" & StLoc & "'"

            CurrentDb.Execute SQL

            Me.MasterLocation.Requery
            Me.MasterLocation = StLocNew
            MasterLocation_AfterUpdate
        End If
    End If

Done:
    On Error Resume Next
    Exit Sub

ErrorHappened:
    MsgBox "From Location: " & StLoc & vbCrLf & "To Location: " & StLocNew & vbCrLf & Err.Description, vbCritical, "Error Copying Location Records"
    Resume Done
End Sub

Private Sub CmdUnlink_Click()
On Error Resume Next
    RunningMode True
    UpdateStatus ("")
    
    screen.MousePointer = 11
    FormUnlinkTables
    CreateHeaderTitle
    CurrentDb.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    
    ' HC 5/2010 update the display to show not linked
    Me.MasterLocation = GetProperty(PROPERTYLINK)
    CmboLocation.DefaultValue = Me.MasterLocation
    Me.filter = "Location = " & QI & Me.MasterLocation & QI
    Me.FilterOn = True
    
    DoEvents
    
    screen.MousePointer = 0
    RunningMode False
End Sub
Private Sub CmdGetServer_Click()
    UpdateStatus ("")
    Call SetSQLCombo(Me.CmboServer)
End Sub
Private Sub cmdImport_Click()
    UpdateStatus ("")
    If Not GetLinksToImport Then
        MsgBox "Import Failed", vbCritical, "Importing Links"
    End If
End Sub
Public Sub CmdLink_Click()
'SA 9/20/12 - Made method public for people that use the form like an API.
    UpdateStatus ("")
    ' verify there is a master location
    If Nz(Me.MasterLocation, "") <> "" Then
        If Me.Dirty = True Then
            If Me.NewRecord = True Then
                'DLC 05/19/10 Use ListIndex check if a combo item was selected (HC 5/2010 -- applied same fix here)
                If Me.CmboLocation.ListIndex < 0 Or Me.CmboDatabase.ListIndex < 0 Then
                    If Me.Dirty Then DoCmd.RunCommand (acCmdUndo)
                Else
                    Me.Dirty = False
                End If
            Else
                Me.Dirty = False 'Save the Record
            End If
        End If
        ' link all the items in the list
        LinkAll Nz(Me.MasterLocation, "")
        CmdLink.SetFocus
    Else    ' HC 5/2010 - added a message when unable to link due to missing location
        MsgBox "Please select the location to link.", vbOKOnly, "Link"
    End If
End Sub
Private Sub CmdLinkItem_Click()
' link a single item
On Error GoTo ErrorHandler
    
    'HC - 5/2010 -- added a message when unable to link -- Only attempt to link when a location is specified
    If CmboLocation.ListIndex >= 0 Then
        UpdateStatus ("")
        RunningMode True
        screen.MousePointer = 11
        
        lblType.Caption = "Type:"
        lblServer.Caption = "Server:"
        lblDatabase.Caption = "Database:"
        lblTable.Caption = "Table:"
        
        ' create the linkcnf
        Select Case UCase(Me.TypeLink)
            Case "ACCESS"
                lblPrefix.Caption = ""
                lblSuffix.Caption = ""
                LinkAccessDatabase Nz(CmboDatabase, ""), Nz(Prefix, ""), Nz(Suffix, ""), True
            Case "SQL"
                lblPrefix.Caption = "Prefix:"
                lblSuffix.Caption = "Suffix:"
                LinkThisDatabase Nz(CmboServer, ""), Nz(CmboDatabase, ""), Nz(CmboSchema, ""), Nz(Prefix, ""), Nz(Suffix, "")
        End Select
    Else
        MsgBox "Cannot link without a location", vbOKOnly, "Link"
        Exit Sub
    End If

Done:
On Error Resume Next
    CurrentDb.TableDefs.Refresh
    DoEvents
    Application.RefreshDatabaseWindow
    DoEvents
    screen.MousePointer = 0
    ClearMessages
    If GetProperty(PROPERTYLINK) <> Me.MasterLocation Then
        SetProperty PROPERTYLINK, Me.MasterLocation
    End If
    CreateHeaderTitle
    DoEvents
    RunningMode False
    CmdLinkItem.SetFocus
    Exit Sub

ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Done
End Sub
Private Sub CmdRefreshViews_Click()
    ' for all linked sql databases, refresh all the views.  Allow to fail 3 times before giving up
    ' HC - 3/2010 - Updated function during upgrade to support SQL Schemas
On Error GoTo ErrorHandler
    Dim db 'As DAO.Database
    Dim rst 'As DAO.Recordset
    Dim TryCount As Integer
    Dim RfshSuccess As Boolean
    Dim sqlString As String
    Dim dName As String
    Dim sName As String
    
    UpdateStatus ("")
    
    RunningMode True
    
    Set rst = Nothing
    If Nz(Me.MasterLocation, "") <> "" Then
        Set db = CurrentDb
        sqlString = "SELECT Location, LinkType, Prefix, Server, Database, Suffix, LastMessage,Schema " & _
            " FROM " & CfgTableName & " " & _
            " WHERE Location = " & QI & Me.MasterLocation & QI & _
            " AND LinkType = " & QI & "SQL" & QI
        Set rst = db.OpenRecordSet(sqlString, dbReadOnly)
        
        If rst.BOF And rst.EOF Then
            MsgBox "Specified location does not exist.", vbCritical, "Linking Location: " & Me.MasterLocation
        End If
        
        screen.MousePointer = 11
        UpdateStatus ("Refreshing Views..." & vbCrLf)
        
        Do Until rst.EOF
            dName = Nz(rst!Database, "")
            sName = Nz(rst!Server, "")
            If left$(sName, 2) = DCSERVERCLASS Or sName = "DEV-SQL-002" Then
                RfshSuccess = False
                TryCount = 0
                Do Until RfshSuccess = True Or TryCount = 3
                   RfshSuccess = RefreshThisDatabaseViews(sName, dName, Nz(rst!Schema, ""), Nz(rst!Prefix, ""), Nz(rst!Suffix, ""))
                     If Not RfshSuccess Then
                         TryCount = TryCount + 1
                         RfshSuccess = False
                     End If
                     If Cancel Then
                        UpdateStatusAppend (Me.txtStatus.Text & vbCrLf & "Canceled")
                        GoTo Done
                    End If
                    DoEvents
                 Loop '--Try again loop
                
                 'show result of refesh attempt.
                    If RfshSuccess = True Then
                       If Len(Me.txtStatus.Text) + 100 < 2000 Then
                            UpdateStatusAppend (Me.txtStatus.Text & vbCrLf & "Refresh Succeeded On: " & sName & "." & dName)
                        Else
                            UpdateStatusAppend ("Refresh Succeeded On: " & sName & "." & sName)
                        End If
                    Else
                      UpdateStatusAppend ("Refresh Failed On: " & sName & "." & dName & vbCrLf & ">>>" & CompleteDBExecuteError & vbCrLf)
                    End If
            Else
                UpdateStatusAppend (Me.txtStatus.Text & vbCrLf & "Refresh Skipped On: " & sName & "." & dName & vbCrLf)
            End If 'is this a DC server?
            rst.MoveNext
        Loop '--Next DB to link loop
    End If 'MasterLocation <> ""
    UpdateStatusAppend (Me.txtStatus.Text & vbCrLf & "Done")
   
Done:
    On Error Resume Next
    If Not rst Is Nothing Then
        rst.Close
    End If
    Set rst = Nothing
    Set db = Nothing
    RunningMode False
    screen.MousePointer = 0
    Me.CmdRefreshViews.SetFocus
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Done
End Sub
Private Function LinkAll(strLocation As String) As Boolean
On Error GoTo ErrorHandler
    If Me.CmdLink.Caption = "&Cancel" Then
        Cancel = True
        DoEvents
        Exit Function
    Else
        Cancel = False
    End If

    'DLC 05/20/2010 Do not link if there are no rows
    If strLocation <> "" And Me.RecordSet.recordCount > 0 Then
        If Me.Dirty = True Then
            If Me.NewRecord = True Then
                'DLC 05/19/10 Use ListIndex check if a combo item was selected (HC 5/2010 -- applied same fix here)
                If Me.CmboLocation.ListIndex < 0 Or Me.CmboDatabase.ListIndex < 0 Then
                    If Me.Dirty Then DoCmd.RunCommand (acCmdUndo)
                Else
                    Me.Dirty = False
                End If
            Else
                Me.Dirty = False 'Save the Record
            End If
        End If
            
        FormUnlinkTables
        
        RunningMode True
        screen.MousePointer = 11
        LinkLocation strLocation
        SavedLocationSet strLocation
        
        DoEvents
        On Error Resume Next
        CurrentDb.TableDefs.Refresh
        Application.RefreshDatabaseWindow
        screen.MousePointer = 0
        
        ClearMessages
        
        'The linked location appears in the title.  Refresh it.
        CreateHeaderTitle
        DoEvents
        
        RunningMode False
        LinkAll = True
        
    Else
        MsgBox "You must select a location"
        LinkAll = False
    End If
Exit_ErrorHandler:
    On Error Resume Next
    screen.MousePointer = 0
    Exit Function
ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Exit_ErrorHandler
End Function
Private Function GetSQLDatabaseSchemas(ByVal StServer As String, ByVal StDatabase As String, Optional CmboBox As Variant = Null) As String()
    ' Get the list of available schemas
    On Error GoTo ErrorHappened
    Dim rsDB As Object
    Dim sqlStr As String
    Dim recordCount As Integer
    recordCount = 0
    Dim StTmp() As String
    
    If IsObject(CmboBox) Then
        Select Case CmboBox.ControlType
            Case 119 'MSForms.ComboBox
                CmboBox.Clear
            Case 111 'ComboBox
                CmboBox.RowSourceType = "Value List"
                CmboBox.RowSource = ""
            Case Else
                MsgBox "An unsupport control type has been passed in." & vbCrLf & vbCrLf & "Error In GetSQLDatabaseSchemas", _
                    vbCritical, "CCACfgLinks.GetSQLDatabaseSchemas"
                GoTo AllDone
        End Select
    End If
    
    UpdateStatus ("Retrieving schemas... ")
    ' Retrieve only the schemas with tables or views, only get those schemas owned by dbo
    sqlStr = " SELECT DISTINCT S.Name" & _
                " FROM sys.schemas  AS S" & _
                " INNER JOIN sys.objects AS O" & _
                " ON S.schema_id = O.schema_id" & _
                " WHERE S.principal_id = 1" & _
                " AND S.schema_id < 16000 AND O.type in ('V','U')" & _
                " ORDER BY Name"
    If GetSQLServerInfo(sqlStr, StServer, StDatabase, rsDB) Then
        If rsDB.EOF And rsDB.BOF Then
            GoTo AllDone
        End If
        ReDim StTmp(rsDB.recordCount)
        Do Until rsDB.EOF
            recordCount = recordCount + 1
            StTmp(recordCount) = rsDB!Name
            If IsObject(CmboBox) Then
                Select Case CmboBox.ControlType
                    Case 119 'MSForms.ComboBox
                        CmboBox.AddItem (StTmp(recordCount))
                    Case 111 'ComboBox
                        If recordCount = 1 Then
                            CmboBox.RowSource = ""
                        End If
                        CmboBox.RowSource = CmboBox.RowSource & ";" & StTmp(recordCount)
                End Select
            End If
            DoEvents
            rsDB.MoveNext
        Loop
        
        If Not rsDB Is Nothing Then
            'need to dissconnect the recordset before closing the connection.
            Set rsDB.ActiveConnection = Nothing
        End If
    End If
        
    GetSQLDatabaseSchemas = StTmp
    
AllDone:
    UpdateStatusAppend (vbCrLf & "Done")
    Me.CmboSchema.SetFocus
        
    On Error Resume Next
    If Not rsDB Is Nothing Then
        rsDB.Close
    End If
    Set rsDB = Nothing
    Exit Function
        
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error In GetSQLDatabaseSchemas", vbCritical, "CCACfgLinks.GetSQLDatabaseSchemas"
    Resume AllDone
    Resume
End Function
Private Function GetSQLDatabases(ByVal StServer As String, Optional CmboBox As Variant = Null) As String()
    ' Get the list of available databases on the server
    ' HC 2010 upgrade, modified to use common routine GetSQLServerInfo
    On Error GoTo ErrorHappened
    Dim rsDB As Object
    Dim sqlStr As String
    Dim recordCount As Integer
    recordCount = 0
    Dim StTmp() As String
    If IsObject(CmboBox) Then
        Select Case CmboBox.ControlType
            Case 119 'MSForms.ComboBox
                CmboBox.Clear
            Case 111 'ComboBox
                CmboBox.RowSourceType = "Value List"
                CmboBox.RowSource = ""
            Case Else
                MsgBox "An unsupport control type has been passed in." & vbCrLf & vbCrLf & "Error In GetSQLDatabases", vbCritical, "CCACfgLinks.GetSQLDatabases"
                GoTo AllDone
        End Select
    End If
    
    UpdateStatus ("Retrieving databases... ")
        
    sqlStr = "SELECT Name FROM sys.databases WHERE database_id  > 4 ORDER BY Name"
    If GetSQLServerInfo(sqlStr, StServer, "master", rsDB) Then
        If rsDB.EOF And rsDB.BOF Then
            GoTo AllDone
        End If
        ReDim StTmp(rsDB.recordCount)
        Do Until rsDB.EOF
            recordCount = recordCount + 1
            StTmp(recordCount) = rsDB!Name
            If IsObject(CmboBox) Then
                Select Case CmboBox.ControlType
                    Case 119 'MSForms.ComboBox
                        CmboBox.AddItem (StTmp(recordCount))
                    Case 111 'ComboBox
                        If recordCount = 1 Then
                            CmboBox.RowSource = ""
                        End If
                        CmboBox.RowSource = CmboBox.RowSource & ";" & StTmp(recordCount)
                End Select
            End If
            DoEvents
            rsDB.MoveNext
        Loop
        
        If Not rsDB Is Nothing Then
            'need to dissconnect the recordset before closing the connection.
            Set rsDB.ActiveConnection = Nothing
        End If
    End If
        
    GetSQLDatabases = StTmp
    
AllDone:
    UpdateStatusAppend (vbCrLf & "Done")
        
    On Error Resume Next
    If Not rsDB Is Nothing Then
        rsDB.Close
    End If
    Set rsDB = Nothing
    Exit Function
        
ErrorHappened:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error In GetSQLDatabases", vbCritical, "CCACfgLinks.GetSQLDatabases"
    Resume AllDone
    Resume
End Function
Private Sub LinkLocation(ByVal strLocation As String)
On Error GoTo ErrorHandler
    Dim db 'As DAO.Database
    Dim rst 'As DAO.Recordset
    Dim sqlString As String
    
    Set db = CurrentDb
    sqlString = "SELECT Location, LinkType,Prefix, Server, Database, Suffix, LastMessage, Schema " & _
        " FROM " & CfgTableName & " " & _
        " WHERE Location = " & QI & strLocation & QI
    Set rst = db.OpenRecordSet(sqlString, dbReadOnly)
    
    If rst.BOF And rst.EOF Then
        MsgBox "Specified location does not exist.", vbCritical, "Linking Location: " & strLocation
    End If
    
    lblType.Caption = "Type:"
    lblServer.Caption = "Server:"
    lblDatabase.Caption = "Database:"
    lblTable.Caption = "Table:"
    
    Do Until rst.EOF
        Select Case Nz(rst!LinkType, "")
            Case "ACCESS"
                lblPrefix.Caption = ""
                lblSuffix.Caption = ""
                LinkAccessDatabase Nz(rst!Database, ""), Nz(rst!Prefix, ""), Nz(rst!Suffix, ""), True
            Case "SQL"
                lblPrefix.Caption = "Prefix:"
                lblSuffix.Caption = "Suffix:"
                LinkThisDatabase Nz(rst!Server, ""), Nz(rst!Database, ""), Nz(rst!Schema, ""), Nz(rst!Prefix, ""), Nz(rst!Suffix, "")
            Case Else
                GoTo NEXTCat
        End Select
NEXTCat:
        If Cancel = True Then
            Exit Do
        End If
        rst.MoveNext
    Loop

Done:
On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Done
End Sub
Private Sub FormUnlinkTables()
On Error GoTo ErrorHandler
    
    lblType.Caption = ""
    lblServer.Caption = ""
    lblDatabase.Caption = ""
    lblTable.Caption = ""
    lblPrefix.Caption = ""
    lblSuffix.Caption = ""
        
    screen.MousePointer = 11
    UnLinkTables (CurrentDb.Name)
    
    'Clear the saved location property now that we are unlinked.
    SavedLocationSet "Not Linked"
    MasterLocation_AfterUpdate
    

Done:
On Error Resume Next
    screen.MousePointer = 0
    Exit Sub

ErrorHandler:
    MsgBox CompleteDBExecuteError, vbCritical, Me.Name & " - " & CodeContextObject.Name
    Resume Done
End Sub
Private Sub UpdateLinkStatus()
    Dim thisLink As LinkCnf
    thisLink = mvLink
    If thisLink.LinkType = LNK_ACCESS Then
        typeMessage.Caption = "Access"
    Else
        typeMessage.Caption = "SQL"
    End If
    serverMessage.Caption = thisLink.Server
    databaseMessage.Caption = thisLink.Database
    tableMessage.Caption = thisLink.TableName
    prefixMessage.Caption = thisLink.Prefix
    suffixMessage.Caption = thisLink.Suffix

    RaiseEvent TableLinked(typeMessage.Caption, thisLink.Prefix & thisLink.TableName & thisLink.Suffix, thisLink.TableName)

    DoEvents

End Sub
Private Sub UpdateStatus(ByVal Status As String)
    Me.txtStatus.Text = Status
    DoEvents
End Sub
Private Sub UpdateStatusAppend(ByVal Status As String)
    If Len(Me.txtStatus.Text) + Len(Status) > 2000 Then
        UpdateStatus (Status & vbCrLf)
    Else
        UpdateStatus (Me.txtStatus.Text & Status & vbCrLf)
    End If
End Sub
Sub RunningMode(Running As Boolean)
    If Running = False Then
        Me.CmdLink.Caption = "&Link"
        Me.CmdGetServer.Enabled = True
        Me.CmdUnlink.Enabled = True
    Else
        Me.CmdLink.Caption = "&Cancel"
        Me.CmdGetServer.Enabled = False
        Me.CmdUnlink.Enabled = False
    End If
    Cancel = False
    DoEvents
    Me.Refresh
End Sub
Private Sub SavedLocationSet(StLoc As String)
    ' HC 5/2010 modified to use generic property setting function
    If SetProperty(PROPERTYLINK, StLoc) = False Then
        MsgBox CompleteDBExecuteError, vbInformation, "Error : Saved LocationSet Set"
    End If
    
End Sub
Private Function ShowOpenFile(Optional sFilter As String = "", Optional sPath As String = "", Optional SFileName As String = "") As String
On Error GoTo CmdBrowseError
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim tmpStr As String
    Dim pos As String
    
    If sFilter = "" Then
' HC 5/2010 -- changed extension to accdb
        sFilter = "Access Files (*.accdb)" & Chr(0) & "*.Accdb" & Chr(0) & _
                  "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    End If
    
    If SFileName <> "" Then
        pos = InStrRev(SFileName, "\")
        If sPath = "" Then
            sPath = Mid(SFileName, 1, pos)
        End If
        If pos <> 0 Then
            SFileName = Mid(SFileName, pos + 1)
        End If
    End If
    
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hWndOwner = Me.hwnd
        .lpstrFilter = sFilter
        .lpstrFile = SFileName
        .nFilterIndex = 1
        .lpstrFile = SFileName & String(257 - Len(SFileName), 0)
        .nMaxFile = Len(OpenFile.lpstrFile) - 1
        .lpstrFileTitle = OpenFile.lpstrFile
        .nMaxFileTitle = OpenFile.nMaxFile
        .lpstrInitialDir = sPath
        .lpstrTitle = "Select a database"
        .flags = 0
    End With
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        Exit Function
    Else
        sFilter = Trim(OpenFile.lpstrFile)
        tmpStr = ""
        For lReturn = 1 To OpenFile.nMaxFile
            If Asc(Mid(sFilter, lReturn, 1)) = 0 Then Exit For
            tmpStr = tmpStr & Mid(sFilter, lReturn, 1)
        Next lReturn
        
        ShowOpenFile = tmpStr
    End If

CmdBrowsExit:
    On Error Resume Next
    Exit Function
CmdBrowseError:
    Select Case Err.Number
    Case 68, 3043, 71, 76, 68 'Disc/Path not ready
        Me.CmboDatabase = ""
        tmpStr = ""
        Resume
    Case Else
        MsgBox Err.Description & String(3, vbCrLf) & "Error getting data location!", vbCritical, "CCA Link Config"
        Resume CmdBrowsExit
        Resume
    End Select
End Function
Private Sub ClearMessages()
    lblType.Caption = ""
    typeMessage.Caption = ""
    lblServer.Caption = ""
    serverMessage.Caption = ""
    lblDatabase.Caption = ""
    databaseMessage.Caption = ""
    lblTable.Caption = ""
    tableMessage.Caption = ""
    lblPrefix.Caption = ""
    prefixMessage.Caption = ""
    lblSuffix.Caption = ""
    suffixMessage.Caption = ""
End Sub
Private Sub MasterLocation_NotInList(NewData As String, Response As Integer)
    Me.MasterLocation = Me.MasterLocation.ItemData(Me.MasterLocation.ListCount - 1)
End Sub
Private Sub SetSQLCombo(ByVal ctl As ComboBox)
    ' HC 5/2010 -- changed to add the ability to select both the dc and the fld servers
    Dim DCList() As String
    Dim FldList() As String
    Dim Result As String
    Result = ""
    
    ' see if both are set off, if so set to default of dc on
    If MvarAutoDcServers = False And MvarAutoFldServers = False Then
        chkDc = True
        chkDc_AfterUpdate
    End If
    
    ' see what we are retrieving
    If MvarAutoDcServers Then
        DCList = GetSqlServers(DCSERVERCLASS)
        Result = ProcessList(DCList, DCSERVERCLASS)
    End If
    
    If MvarAutoFldServers Then
        FldList = GetSqlServers("Field")
        Result = Result & ProcessList(FldList, FLDSERVERCLASS)
    End If
    
    ' clean up the value list
    If Len(Result) > 0 Then
        Result = Mid(Result, 1, Len(Result) - 1)
    End If
    
    ctl.RowSourceType = "Value List"
    ctl.RowSource = Result
    
End Sub
Private Function ProcessList(ByRef list() As String, ByVal serverClass As String) As String
On Error GoTo ErrorRoutine
    Dim Z As Integer
    Dim Result As String
    Result = ""
    
    For Z = 0 To UBound(list)
        Result = Result & list(Z) & ";"
    Next Z

    ProcessList = Result
    Exit Function
    
ErrorRoutine:
    If Err.Number = 9 Then
        MsgBox "SMO request did not return any '" + serverClass + "' servers.", vbOKOnly, "Get Sql Servers"
    Else
        MsgBox "Error retrieving list of '" + serverClass + "' servers" + vbCrLf + Err.Description, vbOKOnly, "Get Sql Servers"
    End If
End Function

' Support Routines
Private Sub CreateHeaderTitle()
    ' Set the Form caption title
    Dim tableVersion As String
    Dim formVersion As String
    Dim Msg As String
    tableVersion = GetTableVersion
    formVersion = GetFormVersion
    
    Msg = "Connolly Config Links (Form  " & formVersion & " / Table  " & tableVersion & ")"
    
    tableVersion = GetProperty(PROPERTYLINK)
    If Nz(tableVersion, "") <> "" Then
        Msg = Msg & " - " & tableVersion
    End If
    Me.Caption = Msg
    DoEvents
    
    'DLC 05/19/2010 Attempt to redraw the application title
    On Error Resume Next
    Application.Run "SetApplicationTitle"

End Sub
Private Function GetTableVersion() As String
' Get the version of the config links table, stored as a property of the table
On Error Resume Next
    Dim Ver As String

    Ver = "0"
    Ver = CurrentDb.TableDefs(CfgTableName).Properties("Description")
    GetTableVersion = Ver
End Function
Private Function GetModuleVersion(ByVal moduleName As String) As String
' Get the version of the config links modules
On Error Resume Next
    Dim Ver As String
    Ver = "0"
    Ver = CurrentDb.Containers("Modules").Documents(moduleName).Properties("Description")
    GetModuleVersion = Ver
End Function

Private Function GetFormVersion() As String
' HC 3/2010 - created a routine
On Error Resume Next
    GetFormVersion = Me.Tag
End Function

Private Function GetLinksToImport(Optional StrDatatbase As String)
'Import Links From external Decipher database.
'StrDatabase: Optional path to external Decipher.  If missing, will prompt user.

On Error GoTo ErrorHappened
    Dim FileName As String
    Dim bReturn As Boolean
    
    If StrDatatbase = "" Then
        'Get the file name to import
        FileName = ShowOpenFile()
        If FileName = "" Then
            GoTo Done
        End If
    Else
        FileName = StrDatatbase
    End If
    
    bReturn = ImportLinks(FileName)
    MasterLocation.Requery

    If Not bReturn Then
        MsgBox CompleteDBExecuteError, vbCritical, "Error Importing Links"
    End If
    
Done:
    On Error Resume Next
    GetLinksToImport = bReturn
    Exit Function
    
ErrorHappened:
    bReturn = False
    MsgBox Err.Description, vbCritical, "Error Importing Links"
    Resume Done
End Function

Public Function LinkAccessDatabase(Database As String, Prefix As String, Suffix As String, IncludeHidden As Boolean) As Boolean
On Error GoTo ErrorHandler
    Dim Link As LinkCnf
    Dim bReturn As Boolean
    
    Dim adCatCurrentDb 'As ADOX.Catalog
    Dim AdCatSourceDb 'As ADOX.Catalog

    bReturn = False
    
    With Link
        .LinkType = LNK_ACCESS
        .Server = ""
        .Database = "" & Database
        .Connect = ""
        .Schema = ""
        .Prefix = "" & Prefix
        .Suffix = "" & Suffix
        .DropIfExists = True
    End With
    ' create the catelog for holding the items we are linking
    Set adCatCurrentDb = CreateObject("ADOX.Catalog")
    ' HC 11/16/2010 commented out and replaced with currentproject.active connection.  Seems to work better at avoiding some corruption in 2010
    adCatCurrentDb.ActiveConnection = CurrentProject.Connection
    
    ' create the source catalog
    Set AdCatSourceDb = CreateObject("ADOX.Catalog")
    ' HC 11/16/2010 - commented out and replaced with currentproject.active connection.  Seems to work better at avoiding some corruption in 2010
    'AdCatSource.activeconnection = LINK_SRC_ACCESS & "Data Source=" & Database & ";"
    'AdCatSource.activeconnection = CurrentProject.Connection
    
    'JL 02/14/2010 Set the connection for the database that we wanto to link.
    AdCatSourceDb.ActiveConnection = LINK_SRC_ACCESS & "Data Source=" & Database & ";"
    bReturn = LinkAccessTables(AdCatSourceDb, Link, adCatCurrentDb, IncludeHidden)
    
    DBEngine.Idle 9 'dbForceOSFlush + dbRefreshCache
    CurrentDb.TableDefs.Refresh

Done:
On Error Resume Next
    Set adCatCurrentDb.ActiveConnection = Nothing
    Set adCatCurrentDb = Nothing
    
    LinkAccessDatabase = bReturn
    Exit Function

ErrorHandler:
    bReturn = False
    Resume Done
End Function


Public Function LinkThisDatabase(ByVal Server As String, ByVal Database As String, ByVal Schema As String, ByVal Prefix As String, ByVal Suffix As String) As Boolean
    Dim AdCat 'As ADOX.Catalog
    Dim Link As LinkCnf
   
    ' create the catalog we are linking to
    Set AdCat = CreateObject("ADOX.Catalog")
    ' HC 11/16/2010 - commented out and replaced with currentproject.active connection.  Seems to work better at avoiding some corruption in 2010
    'adCat.activeconnection = LINK_SRC_ACCESS & "Data Source=" & CurrentDb.Name & ";"
    AdCat.ActiveConnection = CurrentProject.Connection

    Link.LinkType = LNK_SQL
    Link.Server = "" & Server
    Link.Database = "" & Database
    Link.Connect = "ODBC;DRIVER=SQL Server;SERVER=" & Link.Server & ";APP=" & CurrentProject.Name & ";Trusted_Connection=Yes;DATABASE=" & Link.Database & ";"
    Link.Schema = "" & Schema
    Link.Prefix = "" & Prefix
    Link.Suffix = "" & Suffix
    Link.DropIfExists = True
    LinkThisDatabase = GetSQLTables(Link, AdCat)
    
    Set AdCat.ActiveConnection = Nothing
    Set AdCat = Nothing
End Function
Public Function LinkByLocation(ByVal strLocation As String) As Boolean
On Error GoTo ErrorHandler
    Dim db 'As DAO.Database
    Dim rst 'As DAO.Recordset
    Dim sqlString As String
    Dim bReturn As Boolean
    
    bReturn = False
    Set db = CurrentDb
    sqlString = "SELECT Location, LinkType,Prefix, Server, Database, Suffix, LastMessage, Schema " & _
        " FROM " & CfgTableName & " " & _
        " WHERE Location = " & QI & strLocation & QI
    Set rst = db.OpenRecordSet(sqlString, dbReadOnly)
    
    If rst.BOF And rst.EOF Then
        bReturn = False
    Else
        rst.MoveFirst
        Do Until rst.EOF
            Select Case Nz(rst!LinkType, "")
                Case "ACCESS"
                    LinkAccessDatabase Nz(rst!Database, ""), Nz(rst!Prefix, ""), Nz(rst!Suffix, ""), True
                Case "SQL"
                    LinkThisDatabase Nz(rst!Server, ""), Nz(rst!Database, ""), Nz(rst!Schema, ""), Nz(rst!Prefix, ""), Nz(rst!Suffix, "")
                Case Else
                    GoTo NEXTCat
            End Select
NEXTCat:
            If mvCancel = True Then
                Exit Do
            End If
            rst.MoveNext
        Loop
        bReturn = True
    End If
    
Done:
On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    LinkByLocation = bReturn
    Exit Function

ErrorHandler:
    bReturn = False
    MsgBox "Link By Location Error" + vbCrLf + Err.Description, vbOKOnly, "Link By Location"
    Resume Done
End Function

Public Sub UnLinkTables(Optional ByVal Database As String = vbNullString, Optional ByVal Prefix As String = vbNullString)
'Delete linked tables (all or by prefix)
'SA 9/18/2012 - Swithced to use DAO instead of ADOX
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String

    If LenB(Database) = 0 Then
        Set db = CurrentDb()
    Else
        Set db = DBEngine.OpenDatabase(Database)
    End If
    
    '4 - Linked Access Table
    '6 - Linked SQL table
    If LenB(Prefix) = 0 Then
        SQL = "SELECT [Name] FROM MSysobjects WHERE [Type]=4 OR [Type]=6"
    Else
        SQL = "SELECT [Name] FROM MSysobjects WHERE [Name] LIKE '" & Prefix & "*' AND ([Type]=4 OR [Type]=6)"
    End If
    
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    
    Do Until rs.EOF
        If Not DeleteLinkedTable(db, rs![Name]) Then
            UpdateStatus "Unable to unlink " & rs![Name]
        End If
        
        If mvCancel = True Then
            Exit Do
        End If

        rs.MoveNext
    Loop
    
    Application.RefreshDatabaseWindow
ExitNow:
On Error Resume Next
    Set rs = Nothing
    db.Close
    Set db = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Form_CFG_CfgLink:UnlinkTables"
    Resume ExitNow
End Sub

Public Function GetSQLServerInfo(ByVal sqlString As String, ByVal StServer As String, ByVal StDatabase As String, ByRef rsDB As Object) As Boolean
On Error GoTo ErrorHappened
    Dim StConnect As String
    Dim LocConn As Object 'As ADODB.Connection
    Dim retValue As Boolean
    retValue = False
    
    StConnect = LINK_SRC_SQL & "Persist Security Info=False;Data Source=" & StServer & ";" & _
        "Initial Catalog=" & StDatabase & ";"
    
    Set LocConn = CreateObject("ADODB.Connection")
    LocConn.ConnectionTimeout = 3 'secs
    LocConn.Open StConnect
        
    Set rsDB = CreateObject("ADODB.Recordset")
    rsDB.CursorLocation = adUseClient  'Is neccessary for creating a disconnected recordset
    rsDB.CursorType = adOpenStatic
    rsDB.LockType = adLockBatchOptimistic
    rsDB.Open sqlString, LocConn, adOpenStatic, adLockBatchOptimistic
    
    If Not rsDB Is Nothing Then
        'need to dissconnect the recordset before closing the connection.
        Set rsDB.ActiveConnection = Nothing
        retValue = True
    End If
    

Done:
    On Error Resume Next
    ' Close the connection.
    LocConn.Close
    Set LocConn = Nothing
    GetSQLServerInfo = retValue
    Exit Function
    
ErrorHappened:
    retValue = False
    Resume Done
End Function
Public Function RefreshThisDatabaseViews(ByVal Server As String, ByVal Database As String, ByVal Schema As String, ByVal Prefix As String, ByVal Suffix As String) As Boolean
    Dim stProvider As String
    Dim Link As LinkCnf
    Dim bReturn As Boolean
    
    bReturn = True
    
    Link.LinkType = LNK_SQL
    Link.Server = "" & Server
    Link.Database = "" & Database
    Link.Connect = "ODBC;DRIVER=SQL Server;SERVER=" & Link.Server & ";APP=" & CurrentProject.Name & ";Trusted_Connection=Yes;DATABASE=" & Link.Database & ";"
    Link.Schema = "" & Schema
    Link.Prefix = "" & Prefix
    Link.Suffix = "" & Suffix
    Link.DropIfExists = True
    stProvider = LINK_SRC_SQL & "Server=" & Link.Server & ";" & "Database=" & Link.Database & ";"
    bReturn = RefreshSqlViews(stProvider, Link)

    RefreshThisDatabaseViews = bReturn
    Exit Function
End Function
Private Function RefreshSqlViews(ByVal stProvider As String, ByRef Link As LinkCnf) As Boolean
On Error GoTo ErrorHandler
    Dim cn 'As ADODB.Connection
    Dim cmdExec 'As ADODB.Command
    Dim ErrCt As Integer
    Dim sqlTablePrefix As String
    Dim sqlString As String
    Dim rsDB As Object
    Dim LinkTable As Boolean
    Dim Status As String
    Dim bReturn As Boolean
    
    Status = ""
    bReturn = False
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = stProvider
    cn.Open
    
    ' Create the sql String needed to get the sql views
    sqlString = "SELECT O.Name as TableName, S.Name as [SchemaName]" & _
        " FROM sys.views O" & _
        " INNER JOIN sys.schemas S" & _
        " ON O.schema_id = S.schema_id"
    ' Link the views
    If GetSQLServerInfo(sqlString, Link.Server, Link.Database, rsDB) Then
        Set cmdExec = CreateObject("ADODB.Command")
        cmdExec.ActiveConnection = cn
        On Error GoTo LogError
        If Not rsDB.EOF Then
            rsDB.MoveFirst
            Do Until rsDB.EOF
                LinkTable = False
                ' if the prefix is "", link all the tables otherwise only those matching the schema
                If Nz(Link.Schema, "") = "" Then
                    LinkTable = True
                ElseIf Nz(Link.Schema, "") = Nz(rsDB!schemaname, "") Then
                    LinkTable = True
                End If
    
                If LinkTable Then
                    ' set the correct prefix name
                    If Nz(rsDB!schemaname, "dbo") = "dbo" Then
                        sqlTablePrefix = ""
                    Else
                        sqlTablePrefix = rsDB!schemaname & "."
                    End If
                    On Error GoTo LogError
                    cmdExec.CommandText = "Exec sp_refreshview N'" & sqlTablePrefix & rsDB!TableName & "'"
                    cmdExec.Execute
                End If
                rsDB.MoveNext
            Loop
        End If
    End If

Done:
    On Error Resume Next
    If ErrCt >= 1 Then
        Status = CStr(ErrCt) & " Additional Failures."
        UpdateStatusAppend (Status)
    End If
    
    cn.Close
    Set cmdExec = Nothing
    Set cn = Nothing
    RefreshSqlViews = bReturn
    Exit Function

ErrorHandler:
    bReturn = False
    Resume Done

LogError:
    If ErrCt >= 1 Then
        ErrCt = ErrCt + 1
    ElseIf Len(Status) >= 2000 Then
        If Len(Status) > 2000 Then
            Status = left(Status, 2000)
            UpdateStatusAppend (Status)
        End If
        ErrCt = ErrCt + 1
    Else
        Status = rsDB!TableName & " Not Refreshed."
        UpdateStatusAppend (Status)
    End If
    Resume Next
    
End Function
Public Function ImportLinks(ByVal FileName As String) As Boolean
'Import Links From external Decipher database.
'StrDatabase: Optional path to external Decipher.  If missing, will prompt user.
On Error GoTo ErrorHandler
    Dim SQL As String
    Dim bReturn As Boolean
    
    If Nz(FileName, "") = "" Then
        FileName = CurrentDb.Name
    End If
    
    ' HC 5/2010 --purposefully left as select * so old table format items will insert
    SQL = "INSERT INTO " & CfgTableName & " "
    SQL = SQL & "SELECT * "
    SQL = SQL & "FROM " & ExportTableName(FileName) & " IN '" & FileName & "'"
    
    CurrentDb.Execute SQL ' , 128 'dbFailOnError
    DBEngine.Idle 9 'dbRefreshCache + dbForceOSFlush
    bReturn = True

Done:
    On Error Resume Next
    'MsgBox err.Description, vbCritical, "Import error"
    ImportLinks = bReturn
    Exit Function
    
ErrorHandler:
    bReturn = False
    Resume Done
End Function

Private Function ExportTableName(ByVal FileName As String) As String
'Return which table name to import from
'SA 6/26/2012 - Added this function
'New = CFG_CfgLink
'Old = CcaCfgLink
On Error GoTo ErrorHandler
    Dim TableName As String
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(FileName, , True)
    
    TableName = db.TableDefs(CfgTableName).Name
ExitNow:
On Error Resume Next
    ExportTableName = TableName
    Set db = Nothing
Exit Function
ErrorHandler:
    TableName = "CcaCfgLink"
    Resume ExitNow
End Function

Public Function GetSqlServers(ByVal serverClass As String) As String()
' get the list of servers
On Error GoTo ErrorHandler
    Dim xReq 'As New MSXML2.xmlHttp
    Dim xNode
    Dim xDoc 'As DOMDocument
    Dim xDocSub 'As DOMDocument
    Dim SoapEnvelope As String
    Dim Results() As String
    Dim Z As Integer
    SoapEnvelope = GetConfigSoapEnvelope(serverClass)
        
    Set xReq = CreateObject("MSXML2.xmlHttp")
    xReq.Open "POST", ConfigUrl, False, Nothing, Nothing

    xReq.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
    xReq.setRequestHeader "Content-Length", Len(SoapEnvelope)

    xReq.Send (SoapEnvelope)
    
    Set xDoc = CreateObject("MSXML2.DOMDocument")
    Set xDocSub = CreateObject("MSXML2.DOMDocument")
    
    xDoc.loadXML (xReq.ResponseText)
    
    ReDim Results(xDoc.selectNodes("//SERVER").Length - 1)
    Z = 0
    For Each xNode In xDoc.selectNodes("//SERVER")
        Results(Z) = xNode.Attributes.getNamedItem("NAME").Text
        Z = Z + 1
    Next
    
Done:
    Set xReq = Nothing
    Set xDoc = Nothing
    Set xDocSub = Nothing
    GetSqlServers = Results
    Exit Function
    
ErrorHandler:
    Resume Done
End Function

'Private copy of CompleteDBExecuteError from ClsGeneralUtilities (to enable ConfigLinks to be standalone)
Private Function CompleteDBExecuteError(Optional ByVal errDescription As String = "", _
    Optional ByVal bStripCRLF As Boolean = True) As String
    Dim strErrMsg As String
    Dim strErrLoopMsg As String
    Dim errLoop As Error
    If DBEngine.Errors.Count > 0 Then
        'If the DBEngine Error Matches the current error, get the deatils
        If DBEngine.Errors(DBEngine.Errors.Count - 1).Number = Err.Number Then
            'Notify user of any errors that result from executing the query.
            For Each errLoop In DBEngine.Errors
                strErrLoopMsg = Replace(errLoop.Description, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "")
                If bStripCRLF Then
                    strErrLoopMsg = Replace(strErrLoopMsg, vbCrLf, " ", 1, 1, vbTextCompare)
                End If
                strErrMsg = strErrMsg & " " & strErrLoopMsg
            Next errLoop
        End If
    End If
    'If there was a non DBEngine error, use specified or standard error desc/#
    If strErrMsg = "" And Err.Number <> 0 Then
        strErrMsg = IIf(errDescription = "", Err.Description & " (" & Err.Number & ")", errDescription)
    End If
    CompleteDBExecuteError = Trim(strErrMsg)
End Function

' -- Private class functions
Private Function LinkAccessTables(AdCatSrc, ByRef Link As LinkCnf, AdCat, Optional IncludeHidden As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    Dim AdSrc 'As ADOX.Table
    Dim bReturn As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb

    If AdCatSrc Is Nothing Then
        GoTo ExitNow
    End If

    For Each AdSrc In AdCatSrc.Tables
        If mvCancel Then
            Exit For
        End If
        
        Link.TableName = AdSrc.Name
        If UCase(left(Link.TableName, 4)) <> "MSYS" Then
            If ((Not AdSrc.Properties("Jet OLEDB:Table Hidden In Access") Is Nothing And AdSrc.Properties("Jet OLEDB:Table Hidden In Access").Value = False) Or _
                IncludeHidden = True) And (Not AdSrc.Properties("Jet OLEDB:Create Link") Is Nothing And AdSrc.Properties("Jet OLEDB:Create Link").Value = False) Then
                
                Set tdf = db.CreateTableDef(Link.Prefix & Link.TableName & Link.Suffix)
                tdf.SourceTableName = Link.TableName
            
                tdf.Connect = ";DATABASE=" & Link.Database
                
                If Link.DropIfExists Then
                    DeleteLinkedTable db, Link.Prefix & Link.TableName & Link.Suffix
                End If
                If Not AddLinkedTable(db, tdf) Then
                    UpdateStatus "Unable to create link to " & Link.TableName & vbTab & "(LinkType=" & Link.Connect & ")"
                End If
                
                mvLink = Link
                UpdateLinkStatus
            End If
        End If
    Next AdSrc

    bReturn = True
    
ExitNow:
On Error Resume Next
    Set AdSrc = Nothing
    Set db = Nothing
    
    LinkAccessTables = bReturn
Exit Function
ErrorHandler:
    bReturn = False
    Resume ExitNow
End Function
Private Function GetSQLTables(ByRef Link As LinkCnf, AdCat) As Boolean
    
    Dim sqlString As String
    Dim rsDB As Object
    Dim matchPrefix As String
    Dim bReturn As Boolean
    Dim tablesDone As Boolean
    Dim viewsDone As Boolean
    
    ' save the current link prefix
    matchPrefix = Nz(Link.Schema, "")
    bReturn = True
    tablesDone = False
    viewsDone = False
    
    
    ' Create the sql String needed to get the sql table information
    sqlString = "SELECT O.Name as TableName, S.Name as [SchemaName]" & _
        " FROM sys.objects O" & _
        " INNER JOIN sys.schemas S" & _
        " ON O.schema_id = S.schema_id" & _
        " WHERE type in (N'U')"
    ' Get the table info
    If GetSQLServerInfo(sqlString, Link.Server, Link.Database, rsDB) Then
        If Not LinkSQLTables(Link, matchPrefix, rsDB, "Table links") Then
            UpdateStatusAppend (CompleteDBExecuteError)
            bReturn = False
        End If
        rsDB.Close
    Else
        UpdateStatusAppend "No tables found for Server: " & Link.Server & " Database: " & Link.Database
    End If
    
    Set rsDB = Nothing
       
    ' Create the sql String needed to get the sql views
    sqlString = "SELECT O.Name as TableName, S.Name as [SchemaName]" & _
        " FROM sys.views O" & _
        " INNER JOIN sys.schemas S" & _
        " ON O.schema_id = S.schema_id"
    ' Link the views
    If GetSQLServerInfo(sqlString, Link.Server, Link.Database, rsDB) Then
        ' link the Views
        If Not LinkSQLTables(Link, matchPrefix, rsDB, "View links") Then
            UpdateStatusAppend (CompleteDBExecuteError)
            bReturn = False
        End If
        rsDB.Close
    Else
        UpdateStatusAppend "No views found for Server: " & Link.Server & " Database: " & Link.Database
    End If
    
    Set rsDB = Nothing
    GetSQLTables = bReturn
End Function

Private Function LinkSQLTables(ByRef Link As LinkCnf, ByVal schemaToMatch As String, ByRef rsDB As Object, Optional ByVal ObjectType = "Table links") As Boolean
'SA 9/17/2012 - Rewrote to use DAO instead of ADO. This fixes the problem of the ADOX catalog getting corrupted
'               which caused config links to stop working.
On Error GoTo ErrorHandler
    Dim accessNamePrefix As String
    Dim accessNameSuffix As String
    Dim sqlTablePrefix As String
    Dim LinkTable As Boolean
    Dim bReturn As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb
    
    If Nz(Link.Suffix, "") <> "" Then
        accessNameSuffix = "_" & Link.Suffix
    End If
    
    If Not rsDB.EOF And Not rsDB.BOF Then
        rsDB.MoveFirst
        Do Until rsDB.EOF
            If mvCancel Then
                Exit Do
            End If
            
            LinkTable = False
            ' if the prefix is "", link all the tables otherwise only those matching the schema
            If schemaToMatch = vbNullString Then
                LinkTable = True
            ElseIf schemaToMatch = Nz(rsDB!schemaname, vbNullString) Then
                LinkTable = True
            End If

            ' skip the sys diagrams
            If rsDB!TableName = "sysdiagrams" Then
                LinkTable = False
            End If

            If LinkTable Then
                sqlTablePrefix = vbNullString
                accessNamePrefix = vbNullString
                If Nz(Link.Prefix, vbNullString) <> vbNullString Then
                    accessNamePrefix = Link.Prefix & "_"
                End If
                ' set the correct prefix name
                If Nz(rsDB!schemaname, "dbo") = "dbo" Then
                    sqlTablePrefix = vbNullString
                Else
                    sqlTablePrefix = rsDB!schemaname & "."
                End If

                Link.TableName = rsDB!TableName
                
                Set tdf = db.CreateTableDef(accessNamePrefix & Link.TableName & accessNameSuffix)
                tdf.SourceTableName = sqlTablePrefix & Link.TableName
            
                tdf.Connect = Link.Connect
                
                If Link.DropIfExists Then
                    DeleteLinkedTable db, accessNamePrefix & Link.TableName & accessNameSuffix
                End If
                If Not AddLinkedTable(db, tdf) Then
                    UpdateStatus "Unable to create link to " & Link.TableName & vbTab & "(LinkType=" & Link.Connect & ")"
                End If
                
                mvLink = Link
                UpdateLinkStatus
            End If

            rsDB.MoveNext
        Loop
        
        Application.RefreshDatabaseWindow
        
        bReturn = True
        UpdateStatusAppend ObjectType + " complete for Server: " & Link.Server & " Database: " & Link.Database
    Else
        UpdateStatusAppend "No " + ObjectType + " for Server: " & Link.Server & " Database: " & Link.Database
        bReturn = True
    End If

ExitNow:
On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    LinkSQLTables = bReturn
Exit Function
ErrorHandler:
    UpdateStatus Link.TableName & vbTab & Err.Description
    Resume ExitNow
End Function

Private Function AddLinkedTable(ByRef db As DAO.Database, ByRef tdf As DAO.TableDef) As Boolean
'Append table definition to database
On Error GoTo ErrorHappened
    Dim Result As Boolean
    db.TableDefs.Append tdf
    Result = True
ExitNow:
On Error Resume Next
    AddLinkedTable = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function DeleteLinkedTable(ByRef db As DAO.Database, ByVal TableName As String) As Boolean
'Delete specified linked table. Fail silently if the table doesn't exist.
On Error GoTo ErrorHappened
    Dim Result As Boolean
    db.TableDefs.Delete TableName
    Result = True
ExitNow:
On Error Resume Next
    DeleteLinkedTable = Result
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
End Function

Private Function GetConfigSoapEnvelope(ByVal serverType As String) As String
    ' create the soap envelope to for the smo web service
    Dim SoapReq As String
    Dim serverTypeRequest As String
    
    ' HC 5/2010 - modified to accept a server type and insert the request for that server type to soap
    serverTypeRequest = "       <ServerClass>" & serverType & "\SqlServers</ServerClass>"
    SoapReq = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
                "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" & _
                "  <soap12:Body>" & _
                "    <Test_ListServers xmlns=""http://services.ccaintranet.com/SMO/"">" & _
                serverTypeRequest & _
                "       <ServerNamePattern />" & _
                "    </Test_ListServers>" & _
                "  </soap12:Body>" & _
                "</soap12:Envelope>"

    
    GetConfigSoapEnvelope = SoapReq
End Function

Private Function SetProperty(ByVal propertyName As String, ByVal propertyValue As String) As Boolean
' HC 5/2010 -- added generic function to set db properties
On Error GoTo ErrorHappened
    Dim Prop 'As DAO.Property
    Dim Exists As Boolean
    Dim db 'As DAO.Database
    Dim bReturn As Boolean

    bReturn = True
    propertyName = UCase(propertyName)
    
    Set db = CurrentDb
    For Each Prop In db.Properties
        If UCase(Prop.Name) = propertyName Then
            Exists = True
            Exit For
        End If
    Next Prop

    If Exists = True Then
        db.Properties(propertyName) = propertyValue
    Else
        Set Prop = db.CreateProperty(propertyName, 10, propertyValue) 'dbText = 10
        db.Properties.Append Prop
    End If
Done:
    SetProperty = bReturn
    Set db = Nothing
    Set Prop = Nothing

    Exit Function
    
ErrorHappened:
    On Error Resume Next
    bReturn = False
    Resume Done
End Function
Private Function GetProperty(ByVal propertyName As String) As String
' HC 5/2010 - generic routine to get a property setting value
On Error Resume Next
    Dim Result  As String
    Result = ""
    propertyName = UCase(propertyName)
    Result = CurrentDb.Properties(propertyName)
    GetProperty = Result
End Function

Private Function GetBoolProperty(ByVal propertyName As String) As Boolean
' HC 5/2010 - generic routine to get a property setting value
On Error Resume Next
    Dim Result  As String
    Result = ""
    propertyName = UCase(propertyName)
    Result = CurrentDb.Properties(propertyName)
    If Nz(Result, "") = "" Then
        GetBoolProperty = False
    Else
        GetBoolProperty = CBool(Result)
    End If
End Function
