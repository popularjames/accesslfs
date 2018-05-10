Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' DLC 11/16/11 - Log startup information to the ILM Telemetry database
' DLC 11/29/11 - Fix Bug 8911 - ClientName not logged in Telemetry Startup
' DLC 01/24/12 - Updated to store all telemetry events as an action.
' DLC 04/07/12 - Added support for Decipher App Name
'              - If non XML parameter is passed to RecordAction, wrap it in <P> tags.
'              - Fix Telemetry Typo
' DLC 10/22/12 - Included missing Decipher App Name on some telemetry events.
'
' Sample Usage Code:
'
'    Telemetry.RecordOpen "Report", "CP_CpRptAging"
'    Telemetry.RecordOpen "Report", "CP_CpRptAging", "Audit=-1"
'
'    Telemetry.RecordPerformance "Telemetry Startup", CDbl(dEnd - dStart)
'    Telemetry.RecordPerformance "Telemetry Startup", CDbl(dEnd - dStart), "Processing Time" '
'
'    Telemetry.RecordAction "Refresh Grid", "<SQL>SELECT F1 FROM T1</SQL><ROWS>1092</ROWS><TIME>2.33</TIME>"
'    Telemetry.RecordAction "Change Audit", "<AUDIT>2938</AUDIT>"
'

Private Const LogoutWatcherForm As String = "CT_LogoutWatcher"
Private MvTelemetryDB As String
Private MvInsertSproc As String
Private MvTransactionVersion As Integer
Private MvTransactionXml As String
Private MvTransactionAction As String
Private MvUseTelemetry As Boolean
Private MvDecipherVersion As String
Private MvSessionID As String
Private MvTelemetryServer As String

Private genUtils As New CT_ClsGeneralUtilities

Public Property Get SessionID() As String
On Error GoTo ErrorHappened
    'When inserting or removing objects the sessionID can be cleared. Storing it on the
    'logout watcher form allows the SessionID to be recovered.
    If LenB(MvSessionID) = 0 Then
        With Forms(LogoutWatcherForm)
            If LenB(Nz(.txtSessionID, vbNullString)) = 0 Then
                MvSessionID = Mid$(genUtils.NewID(), 2, 36)
                .txtSessionID = MvSessionID
            Else
                MvSessionID = .txtSessionID
                ReadConfig
                UseTelemetry = True
            End If
        End With
    End If
ExitNow:
    On Error Resume Next
    SessionID = MvSessionID
    Exit Property
ErrorHappened:
    'If there is an error getting the sessionID, fail silently
    Resume ExitNow
End Property

Public Property Get AppTitle() As String
On Error Resume Next
AppTitle = CurrentDb.Properties("AppTitle")
End Property

Public Property Get UseTelemetry() As Boolean
On Error GoTo ErrorHappened
    'If the sessionID is lost but the backup sessionID exists, use Telemetry
    If LenB(MvSessionID) = 0 And LenB(Forms(LogoutWatcherForm).txtSessionID) > 0 Then
        UseTelemetry = True
    Else
        UseTelemetry = MvUseTelemetry
    End If
ExitNow:
    On Error Resume Next
    Exit Property
ErrorHappened:
    'Fail silently - this will return False
    Resume ExitNow
End Property

Public Property Let UseTelemetry(Value As Boolean)
    MvUseTelemetry = Value
End Property

Public Property Get DecipherVersion()
    If MvDecipherVersion = vbNullString Then
        On Error Resume Next
        MvDecipherVersion = VersionTemplate
        On Error GoTo 0
    End If
    DecipherVersion = MvDecipherVersion
End Property

'Record a Telemetry action on application startup passing in a session ID that will be used to link to all subsequent actions.
Public Sub Startup()
    Dim ActionDetails As String
    ReadConfig
    'If the Insert Sproc is set log the startup. If it is not, no configuration was selected in cnlyTelemetry so do not record any telemetry during this session.
    If LenB(MvInsertSproc) > 0 Then
        ActionDetails = "<TA ID=""" & SessionID & """>" & _
                           SharedElements("START") & SessionDetails & _
                        "</TA>"
        'If this call fails, there is an unrecoverable error with calling the Telemetry stored procedure so no more telemetry will be recorded during this session.
        UseTelemetry = RecordTelemetry("Exec " & MvInsertSproc & " " & MvTransactionVersion & ", '" & GetTransactionXml(MvTransactionAction) & "', '" & ActionDetails & "'")
    Else
        UseTelemetry = False
    End If
End Sub

'Record a Telemetry action on application shutdown passing in the session ID to link back to the Startup Action and UTC time of the shutdown.
'The data passed to the Telemetry stored procedure is XML encoded to the relevant specification in the Sample payload in the Telemetry.Actions table.
Public Sub Shutdown()
    Dim ActionDetails As String
On Error GoTo ErrorHappened
    If UseTelemetry Then
        ActionDetails = "<TA ID=""" & SessionID & """>" & _
                           SharedElements("SHUTDOWN") & _
                           "<AP ID=""" & SessionID & """>" & SharedElements("SHUTDOWN") & "</AP>" & _
                        "</TA>"
        'If this call fails, there is an unrecoverable error with calling the Telemetry stored procedure so no more telemetry will be recorded during this session.
        UseTelemetry = RecordTelemetry("Exec " & MvInsertSproc & " " & MvTransactionVersion & ", '" & GetTransactionXml(MvTransactionAction) & "', '" & ActionDetails & "'")
    End If
    'Remove the backup copy of the SessionID on the logout watcher form
    Forms(LogoutWatcherForm).txtSessionID = vbNullString
ExitNow:
    On Error Resume Next
    MvSessionID = vbNullString
    UseTelemetry = False
    Exit Sub
ErrorHappened:
    'Fail Silently
    Resume ExitNow
End Sub

'Record an open action to record a standard OPEN event (eg. opening a form or report) based on the name and type of object being opened.
'The data passed to the Telemetry stored procedure is XML encoded to the relevant specification in the Sample payload in the Telemetry.Actions table.
Public Sub RecordOpen(ByVal ObjectType As String, ByVal ObjectName As String, Optional ByVal ObjectParameter As String = vbNullString, Optional ByVal DecipherAppName As String = vbNullString)
    Dim ActionDetails As String
    If UseTelemetry Then
        ActionDetails = "<TA ID=""" & SessionID & """>" & _
                           SharedElements("OPEN", DecipherAppName) & _
                           "<AP ID=""" & SessionID & """>" & SharedElements("OPEN", DecipherAppName) & _
                              "<OT>" & genUtils.XMLEncode(ObjectType) & "</OT>" & _
                              "<ON>" & genUtils.XMLEncode(ObjectName) & "</ON>" & _
                              IIf(LenB(ObjectParameter) > 0, "<P>" & genUtils.XMLEncode(ObjectParameter) & "</P>", vbNullString) & _
                           "</AP>" & _
                        "</TA>"
        'If this call fails, there is an unrecoverable error with calling the Telemetry stored procedure so no more telemetry will be recorded during this session.
        UseTelemetry = RecordTelemetry("Exec " & MvInsertSproc & " " & MvTransactionVersion & ", '" & GetTransactionXml(MvTransactionAction) & "', '" & ActionDetails & "'")
    End If
End Sub

'Record a duration of an event to monitor (eg. the time taken to refresh data grid) based on the task and length of time taken to complete.
'The data passed to the Telemetry stored procedure is XML encoded to the relevant specification in the Sample payload in the Telemetry.Actions table.
Public Sub RecordPerformance(ByVal Task As String, ByVal Duration As Double, Optional ByVal ObjectParameter As String = vbNullString, Optional ByVal DecipherAppName As String = vbNullString)
    Dim ActionDetails As String
    If UseTelemetry Then
        ActionDetails = "<TA ID=""" & SessionID & """>" & _
                           SharedElements("PERFORMANCE", DecipherAppName) & _
                           "<AP ID=""" & SessionID & """>" & SharedElements("PERFORMANCE", DecipherAppName) & _
                              "<T>" & genUtils.XMLEncode(Task) & "</T>" & _
                              "<D>" & Format(Duration, "0.000") & "</D>" & _
                              IIf(LenB(ObjectParameter) > 0, "<P>" & genUtils.XMLEncode(ObjectParameter) & "</P>", vbNullString) & _
                           "</AP>" & _
                        "</TA>"
        'If this call fails, there is an unrecoverable error with calling the Telemetry stored procedure so no more telemetry will be recorded during this session.
        UseTelemetry = RecordTelemetry("Exec " & MvInsertSproc & " " & MvTransactionVersion & ", '" & GetTransactionXml(MvTransactionAction) & "', '" & ActionDetails & "'")
    End If
End Sub

'Record a generic action that is not covered by the standard actions in the Telemetry.Actions table.
'The ActionName allows actions to be grouped once telemetry data is made available in a data warehouse.
'The XML parameter will be stored in addition to the standard elements captured (SessionID, Application Name, UTC Time, Action Name).
'The format of the XML supplied must be a valid list of elements with no root e.g.  <ROWS>123</ROWS><SQL>SELECT 1 FROM MyTable</SQL>
'The following elements will be recorded in the ActionGeneral table:
'<P>  - Parameter
'<SN> - ScreenName
Public Sub RecordAction(ByVal ActionName As String, ByVal strActionDetailsXml As String, Optional ByVal DecipherAppName As String = vbNullString)
    Dim ActionDetails As String
    'If a non XML string was specified, wrap it in <P> tags
    If LenB(strActionDetailsXml) > 0 And left(LTrim(strActionDetailsXml), 1) <> "<" Then
        strActionDetailsXml = "<P>" & strActionDetailsXml & "</P>"
    End If
    If UseTelemetry Or ActionName = "TELEMETRY ERROR" Then
        ActionDetails = "<TA ID=""" & SessionID & """>" & _
                           SharedElements(ActionName, DecipherAppName) & _
                           "<AP ID=""" & SessionID & """>" & _
                            SharedElements(ActionName, DecipherAppName) & _
                            strActionDetailsXml & _
                           "</AP>" & _
                        "</TA>"
        UseTelemetry = RecordTelemetry("Exec " & MvInsertSproc & " " & MvTransactionVersion & ", '" & GetTransactionXml(MvTransactionAction) & "', '" & ActionDetails & "'")
        'If the stored procedure failed, there are 2 likely causes. The first is that bad XML was supplied - if this is the case it is recorded as a "TELEMETRY ERROR" action type.
        If Not UseTelemetry And ActionName <> "TELEMETRY ERROR" Then
            'Attempt to record the error with the original XML wrapped in CDATA tags, if this fails then this is the second type of error (an unrecoverable error
            'with the Telemetry stored procedure). In this case, UseTelemetry will be set to False and no more telemetry will be recorded during this session.
            RecordAction "TELEMETRY ERROR", "<ACTION>" & genUtils.XMLEncode(ActionName) & "</ACTION><DTL><![CDATA[" & strActionDetailsXml & "]]></DTL>"
        End If
    End If
End Sub

'Returns all of the current session infomation as XML to be logged by Telemetry
Private Function SessionDetails() As String
    'SA 03/22/2012 - Added XMLEncode to ClientName. Added monitor dimensions.
    With Identity
        SessionDetails = "<AP ID=""" & SessionID & """>" & SharedElements("START") & "<V>" & genUtils.XMLEncode(DecipherVersion) & "</V>" & _
                        "<U>" & genUtils.XMLEncode(.UserName) & "</U>" & _
                        "<C>" & genUtils.XMLEncode(.Computer) & "</C><TS>" & Format(NowAsUTC, "yyyy-mm-dd hh:mm:ss") & "</TS>" & _
                        "<F>" & genUtils.XMLEncode(.CurrentFolder) & "</F><N>" & genUtils.XMLEncode(CurrentDb.Name) & "</N>" & _
                        "<L>" & genUtils.XMLEncode(.CurrentLocation) & "</L><DA>" & .AuditNum & "</DA>" & _
                        "<DAP>" & .AuditPass & "</DAP><DT>" & .UseDup & "</DT><CP>" & .UseCP & "</CP>" & _
                        "<CN>" & genUtils.XMLEncode(.ClientName) & "</CN><ABD>" & .ShowAuditByDiv & "</ABD><OA>" & .UseOvApp & "</OA>" & _
                        "<MONX>" & ScreenWidth & "</MONX><MONY>" & ScreenHeight & "</MONY></AP>"
    End With
End Function

Private Function SharedElements(ByVal Action As String, Optional ByVal DecipherAppName As String = vbNullString) As String
    SharedElements = "<AN>Decipher</AN><AUDID>" & Identity.AuditNum & "</AUDID><AC>" & genUtils.XMLEncode(Action) & "</AC><TS>" & Format(NowAsUTC, "yyyy-mm-dd hh:mm:ss") & "</TS>" & _
                      IIf(LenB(DecipherAppName) > 0, "<DAN>" & genUtils.XMLEncode(DecipherAppName) & "</DAN>", vbNullString)
End Function

Private Function RecordTelemetry(ByVal sqlStr As String) As Boolean
On Error GoTo ErrorHappened
    Dim StConnect As String
    Dim LocConn 'As ADODB.Connection
    
    StConnect = LINK_SRC_SQL & "Persist Security Info=False;" & _
                 "Data Source=" & MvTelemetryServer & ";" & _
                 "Initial Catalog=" & MvTelemetryDB & ";Connect Timeout = 3"
    
    Set LocConn = CreateObject("ADODB.Connection")
    LocConn.ConnectionString = StConnect
    
    'Open the connection Asynchronous to allow Decipher to continue to load if there is a delay locating the server
    LocConn.Open , , , 16 'adAsyncConnect
    Do While LocConn.State = 2 'adStateConnecting
        DoEvents
    Loop
     
    LocConn.Execute sqlStr, , &H90  'H84 is adAsyncExecute (x10) Or'd with adExecuteNoRecords (x80)
    Do Until LocConn.State = 1  'adStateExecuting
        DoEvents
    Loop
    RecordTelemetry = True
ExitNow:
    If Not LocConn Is Nothing Then
        If LocConn.State > 0 Then
            LocConn.Close
        End If
        Do Until LocConn.State = 0  'adStateClosed
            DoEvents
        Loop
        Set LocConn = Nothing
    End If
    Exit Function
ErrorHappened:
    'Ignore any errors recording Telemetry Data
    RecordTelemetry = False
    Resume ExitNow
    Resume
End Function

Private Function GetTransactionXml(TransactionType)
    GetTransactionXml = Replace(MvTransactionXml, "{TransactionType}", TransactionType)
End Function

Private Sub ReadConfig()
On Error GoTo ErrorHappened
    Dim rs As DAO.RecordSet
    Dim db As DAO.Database
    Set db = CurrentDb
    Set rs = db.OpenRecordSet("SELECT Server, Database, InsertSproc, TransactionVersion, TransactionXML, ActionTransaction from CT_Telemetry WHERE Active")
    If rs.recordCount > 0 Then
        MvTelemetryServer = Nz(rs!Server, vbNullString)
        MvTelemetryDB = Nz(rs!Database, vbNullString)
        MvInsertSproc = Nz(rs!InsertSproc, vbNullString)
        MvTransactionVersion = Nz(rs!TransactionVersion, vbNullString)
        MvTransactionXml = Nz(rs!TransactionXml, vbNullString)
        MvTransactionAction = Nz(rs!ActionTransaction, vbNullString)
    End If
ExitNow:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub