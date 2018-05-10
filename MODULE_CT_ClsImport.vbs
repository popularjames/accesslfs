Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Variables For Properties
Private MvVersion As Double
Private MvVersionRemote As Double
Private MvRemoteDB As String
Private MvLocalDB As String
Private MvImpForms As Boolean
Private MvImpReports As Boolean
Private MvImpModules As Boolean
Private MvImpQueries As Boolean
Private MvImpScreens As Boolean
Private MvImpLinks As Boolean
Private MvImpWorkfiles As Boolean
' HC 9/22/2008
Private MvImpClaimsPlus As Boolean

Private StRowSource As String

Private genUtils As New CT_ClsGeneralUtilities

#If ccCFG = 1 Then
Private mvCfgLinks As Form_CFG_CfgLink
Private WithEvents MvFrmCfg As Form_CFG_CfgLink
Attribute MvFrmCfg.VB_VarHelpID = -1
#End If

Public Event StatusMessage(Src As String, Msg As String, lvl As Integer)

Public Property Let ImportQueries(data As Boolean)
    MvImpQueries = data
End Property
Public Property Get ImportQueries() As Boolean
    ImportQueries = MvImpQueries
End Property
Public Property Let ImportModules(data As Boolean)
    MvImpModules = data
End Property
Public Property Get ImportModules() As Boolean
    ImportModules = MvImpModules
End Property

Public Property Let ImportReports(data As Boolean)
    MvImpReports = data
End Property
Public Property Get ImportReports() As Boolean
    ImportReports = MvImpReports
End Property
Public Property Let ImportForms(data As Boolean)
    MvImpForms = data
End Property
Public Property Get ImportForms() As Boolean
    ImportForms = MvImpForms
End Property
Public Property Let ImportScreens(data As Boolean)
    MvImpScreens = data
End Property
Public Property Get ImportScreens() As Boolean
    ImportScreens = MvImpScreens
End Property
Public Property Let ImportLinks(data As Boolean)
    MvImpLinks = data
End Property
Public Property Get ImportLinks() As Boolean
    ImportLinks = MvImpLinks
End Property
Public Property Let ImportWorkfiles(data As Boolean)
    MvImpWorkfiles = data
End Property
Public Property Get ImportWorkfiles() As Boolean
    ImportWorkfiles = MvImpWorkfiles
End Property
Public Property Let ImportClaimsPlus(data As Boolean)
    MvImpClaimsPlus = data
End Property
Public Property Get ImportClaimsPlus() As Boolean
    ImportClaimsPlus = MvImpClaimsPlus
End Property
Public Property Let DatabaseRemote(data As String)
    Dim fso ' Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(data) = False Then
        RaiseEvent StatusMessage("DatabaseRemote", "The Database you specified does not exits!", 10)
    Else
        MvRemoteDB = data
        MvVersionRemote = genUtils.getVersion(MvRemoteDB)
        RaiseEvent StatusMessage("DatabaseRemote", "Remote Database set.", 0)
        RaiseEvent StatusMessage("DatabaseRemote", MvRemoteDB, 0)
    End If
    Set fso = Nothing
End Property
' HC 5/2010 -- moved properties and class initialization to the top
Public Property Get DatabaseRemote() As String
   DatabaseRemote = MvRemoteDB
End Property

Public Property Get DatabaseLocal() As String
   DatabaseLocal = MvLocalDB
End Property

Public Property Get Version() As Double
    Version = MvVersion
End Property

Public Property Get VersionRemote() As Double
If "" & MvRemoteDB <> "" Then
   VersionRemote = MvVersionRemote
Else
    RaiseEvent StatusMessage("Initialize", "Error Getting Remote Version: " & "Remote Database Not Specified", 10)
    VersionRemote = -1
End If
End Property
Private Sub Class_Initialize()
    MvLocalDB = CurrentDb.Name
    RaiseEvent StatusMessage("Initialize", "Local DB:" & MvLocalDB, 0)
    MvVersion = genUtils.getVersion(CurrentDb.Name)
    RaiseEvent StatusMessage("Initialize", "Local Ver:" & MvVersion, 0)
    MvImpForms = True
    MvImpReports = True
    MvImpModules = True
    MvImpQueries = True
    MvImpScreens = True
    MvImpLinks = True
    MvImpWorkfiles = True
    MvImpClaimsPlus = True

    #If ccCFG = 1 Then
    Set mvCfgLinks = Nothing
    #End If
End Sub

Public Function RunUtility(ByVal UtilityID As Long, Optional RemoteDB As String) As Boolean
'It Synchronizes selected tables from an external database with tables in this database.
    On Error GoTo ErrorHappened
    Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
    
    RaiseEvent StatusMessage("RunUtility", "UTILITY (" & UtilityID & ") IMPORT AT " & Now, 0)
    
    If "" & RemoteDB <> "" Then
        Me.DatabaseRemote = RemoteDB
    End If
    RaiseEvent StatusMessage("RunUtility", "  -" & MvRemoteDB, 0)
    'Get the SQL for the Utilities to run by UtilityID
    SQL = "Select SQL, Notes "
    SQL = SQL & "FRom SCR_ScreensVersionsUtilitiesSQL "
    SQL = SQL & "Where " & MvVersionRemote & " Between MinVer and MaxVer "
    SQL = SQL & "   and UtilityID = " & UtilityID & " "
    SQL = SQL & "Order By SEQ"
    
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    
    With rst
        If .EOF And .BOF Then
            RaiseEvent StatusMessage("RunUtility", "  -No SQL Statements Available for this Utility", 2)
            RunUtility = True
            GoTo ExitNow
        End If
        Do Until .EOF
            If "" & .Fields("Notes") <> "" Then
                RaiseEvent StatusMessage("RunUtility", "  -" & .Fields("Notes"), 0)
            End If
            SQL = "" & .Fields("SQL")
            'Replace all instances where '$Dbname' is located in the query for MvRemoteDB
            SQL = Replace(SQL, "$Dbname", MvRemoteDB)
            db.Execute SQL, dbFailOnError
            RaiseEvent StatusMessage("RunUtility", "  -Records Affected - " & db.RecordsAffected, 0)
             
            .MoveNext
        Loop
    
    End With
    RunUtility = True

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set rst = Nothing
    RaiseEvent StatusMessage("RunUtility", "FINISHED UTILITY (" & UtilityID & "} AT " & Now, 0)
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("RunUtility", "  -" & Err.Description, 10)
    MsgBox "Error in Utility:" & vbCrLf & MvRemoteDB & vbCrLf & vbCrLf & Err.Description, vbCritical, CodeContextObject.Name & ".RunUtility()"
    RunUtility = False
    Resume ExitNow
    Resume

End Function

Public Function RunUtilityEx(ByVal UtilityID As Long, RemoteDB As String, Pairs() As ReplacePairs) As Boolean
On Error GoTo ErrorHappened
Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
Dim X As Integer
RaiseEvent StatusMessage("RunUtilityEx", "UTILITY (" & UtilityID & ") IMPORT AT " & Now, 0)

If "" & RemoteDB <> "" Then
    Me.DatabaseRemote = RemoteDB
End If
Debug.Print Me.DatabaseRemote

RaiseEvent StatusMessage("RunUtilityEx", "  -" & MvRemoteDB, 0)

SQL = "Select SQL, Notes "
SQL = SQL & "FRom SCR_ScreensVersionsUtilitiesSQL "
SQL = SQL & "Where " & MvVersionRemote & " Between MinVer and MaxVer "
SQL = SQL & "   and UtilityID = " & UtilityID & " "
SQL = SQL & "Order By SEQ"
Debug.Print SQL
Set db = CurrentDb
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)

With rst
    If .EOF And .BOF Then
        RaiseEvent StatusMessage("RunUtilityEx", "  -No SQL Statements Available for this Utility", 2)
        RunUtilityEx = True
        GoTo ExitNow
    End If
    Do Until .EOF
        If "" & .Fields("Notes") <> "" Then
            RaiseEvent StatusMessage("RunUtilityEx", "  -" & .Fields("Notes"), 0)
        End If
        SQL = "" & .Fields("SQL")
        SQL = Replace(SQL, "$Dbname", MvRemoteDB)
        
            'REPLACE WITH ARRAY
            For X = 0 To UBound(Pairs())
                SQL = Replace(SQL, Pairs(X).From, Pairs(X).To)
            Next X
        
        db.Execute SQL, dbFailOnError
        RaiseEvent StatusMessage("RunUtilityEx", "  -Records Affected - " & db.RecordsAffected, 0)
         
        .MoveNext
    Loop

End With
RunUtilityEx = True

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set rst = Nothing
    RaiseEvent StatusMessage("RunUtilityEx", "FINISHED UTILITY (" & UtilityID & "} AT " & Now, 0)
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("RunUtilityEx", "  -" & Err.Description, 10)
    MsgBox "Error in Utility:" & vbCrLf & MvRemoteDB & vbCrLf & vbCrLf & Err.Description, vbCritical, CodeContextObject.Name & ".RunUtility()"
    RunUtilityEx = False
    Resume ExitNow
    Resume

End Function

Public Sub Run(Optional RemoteDB As String)
'TODO Need to rework import to work with new CFG.net if ever available
#If ccCFG = 1 Then
' HC 5/2010 modified to use the config links form
On Error GoTo ErrorHappened
    RaiseEvent StatusMessage("Run", "STARTED IMPORT AT " & Now, 0)
    StRowSource = ""
    If "" & RemoteDB <> "" Then
        Me.DatabaseRemote = RemoteDB
    End If

    RaiseEvent StatusMessage("Run", "Linking Remote database.", 0)
    RaiseEvent StatusMessage("Run", MvRemoteDB, 0)


    If mvCfgLinks Is Nothing Then
        Set mvCfgLinks = New Form_CFG_CfgLink
        mvCfgLinks.visible = False
    End If

    If mvCfgLinks.LinkAccessDatabase(MvRemoteDB, "Import_", "", True) = False Then
        RaiseEvent StatusMessage("CFG_CfgLink", "Error linking database ", 0)
        GoTo ExitNow
    End If

    RaiseEvent StatusMessage("Run", "Finished Linking.", 0)
    DBEngine.Idle dbForceOSFlush + dbRefreshCache
    CurrentDb.TableDefs.Refresh
    RefreshDatabaseWindow
    DoEvents


    If MvImpScreens = True Then
        RaiseEvent StatusMessage("Run", "Importing Screens", 0)
        If SQLExecute = False Then
            GoTo ExitNow
        End If
    End If

    If MvImpWorkfiles = True Then
        RaiseEvent StatusMessage("Run", "Importing Workfiles", 0)
        If SQLExecuteWorkfiles = False Then
            GoTo ExitNow
        End If
    End If

    ' HC 9/22/2008
    If MvImpClaimsPlus = True Then
        RaiseEvent StatusMessage("Run", "Importing Claims Plus Tables", 0)
        If SQLExecuteClaimsPlus = False Then
            GoTo ExitNow
        End If
    End If

    If MvImpReports = True Then
        RaiseEvent StatusMessage("Run", "Importing Reports", 0)
        If ImportAll(MvRemoteDB, ObjType.objReport) = False Then
            GoTo ExitNow
        End If
    End If
    
    If MvImpModules = True Then
        RaiseEvent StatusMessage("Run", "Importing Modules", 0)
        If ImportAll(MvRemoteDB, ObjType.objModule) = False Then
            GoTo ExitNow
        End If
    End If
    
    If MvImpQueries = True Then
        RaiseEvent StatusMessage("Run", "Importing Tables/Queries", 0)
        If ImportAll(MvRemoteDB, ObjType.objQuery) = False Then
            GoTo ExitNow
        End If
    End If

    If MvImpForms = True Then
        RaiseEvent StatusMessage("Run", "Importing Forms", 0)
        If ImportAll(MvRemoteDB, ObjType.objForm) = False Then
            GoTo ExitNow
        End If
    End If

    If MvImpLinks = True Then
        RaiseEvent StatusMessage("Run", "Importing Links", 0)
        If mvCfgLinks Is Nothing Then
            Set mvCfgLinks = New Form_CFG_CfgLink
            mvCfgLinks.visible = False
        End If
        If mvCfgLinks.ImportLinks(MvRemoteDB) = False Then
            GoTo ExitNow
        End If
    End If

    'Unlink remote db
    UnLinkImportTables
    
    If Nz(StRowSource, "") <> "" Then
        StRowSource = left(StRowSource, Len(StRowSource) - 1)
        OpenForm StRowSource, MvRemoteDB
    End If

ExitNow:
    On Error Resume Next
    Set mvCfgLinks = Nothing
    RaiseEvent StatusMessage("Run", "FINISHED IMPORT AT " & Now, -1)
    Exit Sub

ErrorHappened:
    RaiseEvent StatusMessage("Run", Err.Description, 10)
    MsgBox "Error Importing Database:" & vbCrLf & MvRemoteDB & vbCrLf & vbCrLf & Err.Description, vbCritical, CodeContextObject.Name & ".Run()"
    Resume ExitNow
    Resume
#Else
    MsgBox "Config links must be installed", vbCritical, "Missing App"
#End If
End Sub

Private Sub UnLinkImportTables()
'TODO Need to rework import to work with new CFG.net if ever available
#If ccCFG = 1 Then
' HC 2010 upgrade, config links
    If mvCfgLinks Is Nothing Then
        Set mvCfgLinks = New Form_CFG_CfgLink
        mvCfgLinks.visible = False
    End If
    
    screen.MousePointer = 11
    mvCfgLinks.UnLinkTables CurrentDb.Name, "Import_"
    
    CurrentDb.TableDefs.Refresh
    RefreshDatabaseWindow
    DoEvents
    
    screen.MousePointer = 0
    Set mvCfgLinks = Nothing
#Else
    MsgBox "Config links must be installed", vbCritical, "Missing App"
#End If
End Sub

Private Sub Import_SavedLocationSet(StLoc As String)
On Error GoTo ErrorHappened
Dim Prop 'As DAO.Property
Dim Exists As Boolean
Dim db 'As DAO.Database

Set db = CurrentDb
For Each Prop In db.Properties
    If UCase(Prop.Name) = UCase("LastLinkLocation") Then
        Exists = True
        Exit For
    End If
Next Prop

If Exists = True Then
    db.Properties("LastLinkLocation") = StLoc
Else
    Set Prop = db.CreateProperty("LastLinkLocation", 10, StLoc) 'dbText = 10
    db.Properties.Append Prop
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set Prop = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbInformation, "Error : Saved LocationSet Set"
    Resume ExitNow
    

End Sub

Private Function SQLExecute() As Boolean
On Error GoTo ErrorHappened
Dim SQL As String, X As Integer
Dim db As DAO.Database, rst As DAO.RecordSet
SQL = "SELECT VPS.SQL, VPS.NOTES "
SQL = SQL & "FROM SCR_ScreensVersionsPaths AS VP "
SQL = SQL & "   INNER JOIN SCR_ScreensVersionsPathsSQL AS VPS ON "
SQL = SQL & "           VP.PathID = VPS.PathID "
SQL = SQL & "WHERE " & MvVersionRemote & " Between VP.MinSrcVer and MaxSrcVer "
SQL = SQL & " And " & MvVersion & " Between MinDestVer and MaxDestVer "
SQL = SQL & " And " & MvVersionRemote & " >= VPS.MinSrcVer "
SQL = SQL & "ORDER BY VPS.Sort;"

Set db = CurrentDb
db.TableDefs.Refresh
DoEvents
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)

With rst
    If .EOF And .BOF Then
        RaiseEvent StatusMessage("SQLExecute", "No SQL Defined for this upgrade path", 10)
        SQLExecute = False
    End If
    On Error GoTo DataError
    Do Until rst.EOF
        X = X + 1
        RaiseEvent StatusMessage("SQLExecute", "Executing SQL: " & .Fields("Notes"), 1)
        SQL = .Fields("SQL")
        SQL = Replace(SQL, "$DbName", MvRemoteDB)
        SQL = Replace(SQL, "$Identity.UserName", Chr(34) & Identity.UserName & Chr(34))
        db.Execute SQL, dbFailOnError
        .MoveNext
    Loop
End With

SQLExecute = True
ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("SQLExecute", "Error Running Upgrade SQL:" & vbCrLf & vbCrLf & Err.Description, 10)
    SQLExecute = False
    Resume ExitNow
    Resume
DataError:
    RaiseEvent StatusMessage("SQLExecute", Err.Description, 10)
    RaiseEvent StatusMessage("SQLExecute", "Failed Executing a SQL Statement", 10)
    SQLExecute = False
    Resume ExitNow
    Resume
End Function

Private Function SQLExecuteWorkfiles() As Boolean
On Error GoTo ErrorHappened
Dim SQL As String, X As Integer
Dim db As DAO.Database, rst As DAO.RecordSet
SQL = "SELECT VPS.SQL, VPS.NOTES "
SQL = SQL & "FROM PRJ_ProjectsVersionsPaths AS VP "
SQL = SQL & "   INNER JOIN PRJ_ProjectsVersionsPathsSQL AS VPS ON "
SQL = SQL & "           VP.PathID = VPS.PathID "
SQL = SQL & "WHERE " & MvVersionRemote & " Between VP.MinSrcVer and MaxSrcVer "
SQL = SQL & " And " & MvVersion & " Between MinDestVer and MaxDestVer "
SQL = SQL & " And " & MvVersionRemote & " >= VPS.MinSrcVer "
SQL = SQL & "ORDER BY VPS.Sort;"

Set db = CurrentDb
db.TableDefs.Refresh
DoEvents
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)

With rst
    If .EOF And .BOF Then
        RaiseEvent StatusMessage("SQLExecuteWorkfiles", "No SQL Defined for this upgrade path", 10)
        SQLExecuteWorkfiles = False
    End If
    On Error GoTo DataError
    Do Until rst.EOF
        X = X + 1
        'RaiseEvent StatusMessage("SQLExecute", "Executing SQL: " & Right("000" & x, 3), 0)
        RaiseEvent StatusMessage("SQLExecuteWorkfiles", "Executing SQL: " & .Fields("Notes"), 1)
        SQL = .Fields("SQL")
        SQL = Replace(SQL, "$DbName", MvRemoteDB)
        db.Execute SQL, dbFailOnError
        .MoveNext
    Loop
End With

SQLExecuteWorkfiles = True
ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("SQLExecuteWorkfiles", "Error Running Upgrade SQL:" & vbCrLf & vbCrLf & Err.Description, 10)
    SQLExecuteWorkfiles = False
    Resume ExitNow
    Resume
DataError:
    RaiseEvent StatusMessage("SQLExecuteWorkfiles", Err.Description, 10)
    RaiseEvent StatusMessage("SQLExecuteWorkfiles", "Failed Executing a SQL Statement", 10)
    SQLExecuteWorkfiles = False
    Resume ExitNow
    Resume
End Function
' HC 9/22/2008
Private Function SQLExecuteClaimsPlus() As Boolean
On Error GoTo ErrorHappened
Dim SQL As String, X As Integer
Dim db As DAO.Database, rst As DAO.RecordSet
SQL = "SELECT VPS.SQL, VPS.NOTES "
SQL = SQL & "FROM CP_ClaimsPlusVersionsPaths AS VP "
SQL = SQL & "   INNER JOIN CP_ClaimsPlusVersionsPathsSql AS VPS ON "
SQL = SQL & "           VP.PathID = VPS.PathID "
SQL = SQL & "WHERE " & MvVersionRemote & " Between VP.MinSrcVer and MaxSrcVer "
SQL = SQL & " And " & MvVersion & " Between MinDestVer and MaxDestVer "
SQL = SQL & " And " & MvVersionRemote & " >= VPS.MinSrcVer "
SQL = SQL & "ORDER BY VPS.Sort;"

Set db = CurrentDb
db.TableDefs.Refresh
DoEvents
Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)

With rst
    If .EOF And .BOF Then
        RaiseEvent StatusMessage("Update Claims Plus tables", "No SQL Defined for this upgrade path", 10)
        SQLExecuteClaimsPlus = False
    End If
    On Error GoTo DataError
    Do Until rst.EOF
        X = X + 1
        'RaiseEvent StatusMessage("SQLExecute", "Executing SQL: " & Right("000" & x, 3), 0)
        RaiseEvent StatusMessage("Update Claims Plus tables", "Executing SQL: " & .Fields("Notes"), 1)
        SQL = .Fields("SQL")
        SQL = Replace(SQL, "$DbName", MvRemoteDB)
        db.Execute SQL, dbFailOnError
        .MoveNext
    Loop
End With

SQLExecuteClaimsPlus = True
ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("Update Claims Plus tables", "Error Running Upgrade SQL:" & vbCrLf & vbCrLf & Err.Description, 10)
    SQLExecuteClaimsPlus = False
    Resume ExitNow
    Resume
DataError:
    RaiseEvent StatusMessage("Update Claims Plus tables", Err.Description, 10)
    RaiseEvent StatusMessage("SQLExecuUpdate Claims Plus tablesteWorkfiles", "Failed Executing a SQL Statement", 10)
    SQLExecuteClaimsPlus = False
    Resume ExitNow
    Resume
End Function

Public Function ObjectExists(stName As String, Tp As ObjType, dbName As String) As Boolean
On Error GoTo ErrorHappened
    Dim db As DAO.Database, rst As DAO.RecordSet
    Dim SQL As String
    
    SQL = "SELECT Replace([O1].[Name],chr(34),chr(39)) as Name From MSysObjects as O1 " & _
           " INNER JOIN MSysObjects as O2 on O1.ParentID = O2.[ID] "
    stName = Replace(stName, Chr(34), Chr(39))
    Set db = DBEngine.OpenDatabase(dbName)
    Select Case Tp
    Case ObjType.objTable, ObjType.objQuery 'Does Not matter if it is a table or query
        SQL = SQL & "Where O2.Name = 'Tables' and Replace([O1].[Name],chr(34),chr(39))  = " & Chr(34) & stName & Chr(34)
    Case ObjType.objForm
        SQL = SQL & "Where O2.Name = 'Forms' and Replace([O1].[Name],chr(34),chr(39))  = " & Chr(34) & stName & Chr(34)
    Case ObjType.objReport
        SQL = SQL & "Where O2.Name = 'Reports' and Replace([O1].[Name],chr(34),chr(39))  = " & Chr(34) & stName & Chr(34)
    Case ObjType.objModule
        SQL = SQL & "Where O2.Name = 'Modules' and Replace([O1].[Name],chr(34),chr(39))  = " & Chr(34) & stName & Chr(34)
        'ACCOUNT FOR RETIRED NAMES
        SQL = SQL & " or " & Chr(39) & stName & Chr(39) & " in ('CCA Screens CommandBars ') "
    Case Else
        RaiseEvent StatusMessage("ObjectExists", "Unhandled object type", 8)
        ObjectExists = False
        GoTo ExitNow
    End Select
    
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    If rst.EOF And rst.BOF Then
        ObjectExists = False
    Else
        ObjectExists = True
    End If
    rst.Close
    db.Close

ExitNow:
    On Error Resume Next
    Set rst = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("ObjectExists", Err.Description, 8)
    RaiseEvent StatusMessage("ObjectExists", "Unable to check existance of " & stName, 8)
    ObjectExists = False
    Resume ExitNow
End Function


Public Function ImportAll(SrcDb As String, Tp As ObjType) As Boolean
On Error GoTo ErrorHappened
Dim DbSrc As DAO.Database, oVar As Variant
Dim ctr As DAO.Container, Doc As DAO.Document
Dim CtrName As String

'RaiseEvent StatusMessage("ImportAll", "Begin", 0)

Select Case Tp
Case ObjType.objTable, ObjType.objQuery
    CtrName = "TABLES"
Case ObjType.objForm
    CtrName = "Forms"
Case ObjType.objReport
    CtrName = "Reports"
Case ObjType.objModule
    CtrName = "Modules"
Case Else
    RaiseEvent StatusMessage("ImportAll", "Undocument object type specified", 8)
    GoTo ExitNow
End Select

Set DbSrc = DBEngine.OpenDatabase(SrcDb)
With DbSrc
    For Each ctr In .Containers
        If UCase(ctr.Name) = UCase(CtrName) Then
            For Each Doc In ctr.Documents
                'Only import the object if an object by the same name/type does not exist in the local db.
                If ObjectExists(Doc.Name, Tp, MvLocalDB) = False _
                    And Nz(DLookup("ID", "CT_ExcludeObjects", "ObjectType = '" & CtrName & _
                            "' AND ObjectName = '" & Replace(Doc.Name, "'", "''") & "'"), 0) = 0 Then
                    Select Case Tp
                    Case ObjType.objQuery
                        On Error GoTo BruteForceTrap
                            Set oVar = DbSrc.QueryDefs(Doc.Name)
                        On Error GoTo ErrorHappened
                        If Not oVar Is Nothing Then ' THE RIGHT TYPE
                            If left(Doc.Name, 1) <> "~" Then 'TEMP QUERIES
                                RaiseEvent StatusMessage("ImportAll", "Import:  Queries." & Doc.Name, 1)
                                DoEvents
                                DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, Tp, Doc.Name, Doc.Name
                            End If
                        Else
                            On Error GoTo BruteForceTrap
                                Set oVar = DbSrc.TableDefs(Doc.Name)
                            On Error GoTo ErrorHappened
                            If Not oVar Is Nothing Then ' THE RIGHT TYPE
                                If oVar.Connect = "" Then   ' not a linked table, import it.  Added by DBrady 10.16.2008
                                    RaiseEvent StatusMessage("ImportAll", "Import:  Tables." & Doc.Name, 1)
                                    DoEvents
                                    
                                    On Error Resume Next 'error trapping and messaging added by Dbrady 10.16.2008
                                    DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, acTable, Doc.Name, Doc.Name
                                    If Err.Number > 0 Then
                                        RaiseEvent StatusMessage("ImportAll", "Import:  Tables." & Doc.Name & " - FAILED", 8)
                                        RaiseEvent StatusMessage("ImportAll", "    " & Err.Description, 8)
                                    End If
                                    On Error GoTo ErrorHappened
                                End If
                            End If
                        End If
                    Case ObjType.objForm
                        Select Case UCase(left(Doc.Name, 3))
                        Case "CCA", "CNL", "SCR", "CDT" 'Assume These are system and not user forms
                            RaiseEvent StatusMessage("ImportAll", "Skip:  " & CtrName & "." & Doc.Name, 2)
                            StRowSource = StRowSource & Doc.Name & ",Form;" & Tp & ","
                        Case Else
                            RaiseEvent StatusMessage("ImportAll", "Import:  " & CtrName & "." & Doc.Name, 1)
                            DoEvents
                            DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, Tp, Doc.Name, Doc.Name
                        End Select
                    Case ObjType.objModule
                        Select Case UCase(left(Doc.Name, 3))
                        Case "CCA", "CNL", "CDT" 'Assume These are system and not user modules
                            RaiseEvent StatusMessage("ImportAll", "Skip:  " & CtrName & "." & Doc.Name, 2)
                            StRowSource = StRowSource & Doc.Name & ",Module;" & Tp & ","
                        Case Else
                            RaiseEvent StatusMessage("ImportAll", "Import:  " & CtrName & "." & Doc.Name, 1)
                            DoEvents
                            DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, Tp, Doc.Name, Doc.Name
                        End Select
                    Case Else
                        RaiseEvent StatusMessage("ImportAll", "Import:  " & CtrName & "." & Doc.Name, 1)
                        DoEvents
                        DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, Tp, Doc.Name, Doc.Name
                    End Select
                Else
                    RaiseEvent StatusMessage("ImportAll", "Exists:  " & CtrName & "." & Doc.Name, 2)
                End If
            Next Doc
        End If
    Next ctr
End With


ImportAll = True

ExitNow:
    On Error Resume Next
    Set DbSrc = Nothing
    Set ctr = Nothing
    Set Doc = Nothing
    Set oVar = Nothing
    Exit Function

ErrorHappened:
    RaiseEvent StatusMessage("ImportAll", Err.Description, 8)
    If Err.Number = 3011 Then
        Resume Next
    Else
        ImportAll = False
        Resume ExitNow
        Resume
    End If
BruteForceTrap:
    Set oVar = Nothing
    Resume Next
End Function

Private Sub OpenForm(StRowSource As String, SrcDb As String)
On Error GoTo ErrorHappened
Dim frm As New Form_CT_PopupSelect
Dim i As Variant

With frm
    With .Lst
        .RowSource = StRowSource
        .RowSourceType = "Value List"
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "2.9" & Chr(34) & ";0.4" & Chr(34) & ";0" & Chr(34)
        .Requery
    End With
    .Title = "Objects Not Imported due to Naming Conventions"
    .ListTitle = "Select the Objects that you want to import"
    .StartupWidth = -1  'AUTO
    .visible = True
    
    Do While .Results = vbApplicationModal
        DoEvents
    Loop

    If .Results = vbOK Then
        RaiseEvent StatusMessage("Run", "Importing Missed Objects", 0)
        For Each i In frm.Lst.ItemsSelected
            If frm.Lst.Selected(i) Then
                RaiseEvent StatusMessage("OpenForm", "Import:  " & frm.Lst.Column(1, i) & "." & frm.Lst.Column(0, i), 1)
                DoCmd.TransferDatabase acImport, "Microsoft Access", SrcDb, frm.Lst.Column(2, i), frm.Lst.Column(0, i), frm.Lst.Column(0, i)
            End If
        Next i

    End If

End With
ExitNow:
    On Error Resume Next
   ' Me.requery
    DoCmd.Hourglass False
    Set frm = Nothing
    Exit Sub

ErrorHappened:
    Resume
    
End Sub