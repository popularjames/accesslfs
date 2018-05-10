Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Darren Collard  11/15/11 - Added additional properties from CP General tab
'                          - Updated CurrentFolder to set FileSystemObject to Nothing

'SA 11/28/2012 - Removed options table checks for CP, Dup Tool and Overrides

Private MvAuditor As String
Private MvDataSheetStyle As CnlyDataSheetStyle
Private MvDataSheetStyleDefault As CnlyDataSheetStyle
Private MvFolderCurrent As String
Private MvFolderOutput As String
Private MvAuditNum As Long
Private MvAuditPass As Integer
Private MvClientName As String
Private MvCCAHelp As String
Private MvShowAuditByDiv As Integer
Private MvLocation As String

Public Property Get UseDup() As Integer
'SA 11/28/2012 - Changed to check if app is installed
    If IsProductInstalled("Dup Tool") Then
        UseDup = 1
    Else
        UseDup = 0
    End If
End Property

Public Property Get UseCP() As Integer
'SA 11/28/2012 - Changed to check if app is installed
    If IsProductInstalled("ClaimsPlus Framework") Then
        UseCP = 1
    Else
        UseCP = 0
    End If
End Property

Public Property Get UseOvApp() As Integer
'SA 11/28/2012 - Kept for backward compatibility
    UseOvApp = 0
End Property

Public Property Get CurrentLocation() As String
    Dim rtn As String
    On Error Resume Next  'If Prop does not exits then just return blank
    rtn = CurrentDb.Properties("LastLinkLocation")
    CurrentLocation = rtn
End Property

Public Property Let CCAHelp(Fname As String)
    MvCCAHelp = Fname
End Property

Public Property Get CCAHelp() As String
    CCAHelp = MvCCAHelp
End Property

Public Property Get AuditNum() As Long
    AuditNum = MvAuditNum
End Property
'MN -- sets default Audit ID -- added as an requirement for Contract Compliance
Public Property Let AuditNum(intAuditNum As Long)
    MvAuditNum = intAuditNum
End Property

Public Property Get ShowAuditByDiv() As Integer
    ShowAuditByDiv = MvShowAuditByDiv
End Property

Public Property Get ClientName() As String
    ClientName = MvClientName
End Property

Public Property Get AuditPass() As Integer
    AuditPass = MvAuditPass
End Property

Public Property Get FolderOutput() As String
    If "" & MvFolderOutput = "" Then
        MvFolderOutput = Interaction.GetSetting("AuditProbe", "General", "FolderOutput", "")
    End If
    'If Still Blank Get Current Path
    If "" & MvFolderOutput = "" Then
        MvFolderOutput = Me.CurrentFolder
    End If
    FolderOutput = MvFolderOutput
End Property

Public Property Let FolderOutput(data As String)
    MvFolderOutput = data
    Call SaveSetting("AuditProbe", "General", "FolderOutput", MvFolderOutput)
End Property

Public Property Get CurrentFolder() As String
On Error GoTo ErrorHappened
    Dim fso
    If Nz(MvFolderCurrent, vbNullString) = vbNullString Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        MvFolderCurrent = fso.GetFile(CurrentDb.Name).ParentFolder.Path
    End If
ExitNow:
    On Error Resume Next
    Set fso = Nothing
    CurrentFolder = MvFolderCurrent
    Exit Function
ErrorHappened:
    Resume ExitNow
End Property

Public Property Get Auditor() As String
On Error GoTo ErrorHappened

If "" & MvAuditor = "" Then
    MvAuditor = Interaction.GetSetting("AuditProbe", "General", "Auditor", "")
End If

ExitNow:
    On Error Resume Next

    Auditor = MvAuditor
    Exit Property
ErrorHappened:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Property
Public Property Let Auditor(data As String)
    If UCase(data) <> MvAuditor Then
        Call SaveSetting("AuditProbe", "General", "Auditor", UCase(data)) 'save new value in registry
        MvAuditor = UCase(data) ' set new value
    End If
End Property

Public Property Get DataSheetStyle() As CnlyDataSheetStyle
On Error GoTo ErrorHappened
DataSheetStyle = MvDataSheetStyle
ExitNow:
    On Error Resume Next
    Exit Property
ErrorHappened:
    MsgBox Err.Description, vbInformation, "ClsIDentity.DataSheetStyle"
    Resume ExitNow
End Property

Public Property Let DataSheetStyle(data As CnlyDataSheetStyle)
MvDataSheetStyle = data
With MvDataSheetStyle
    Call SaveSetting("AuditProbe", "DataSheetStyle", "BackGroundColor", .BackGroundColor)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "BorderLineStyle", .BorderLineStyle)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "CellsEffect", .CellsEffect)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontFamily", .FontFamily)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontItalic", .FontItalic)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontSize", .fontsize)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontUnderline", .FontUnderline)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontWeight", .FontWeight)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "FontSize", .fontsize)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "ForeColor", .ForeColor)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "GridlinesBehavior", .GridlinesBehavior)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "GridlinesColor", .GridlinesColor)
    Call SaveSetting("AuditProbe", "DataSheetStyle", "HeaderUnderlineStyle", .HeaderUnderlineStyle)
End With
End Property

Public Sub ReloadOptions()
On Error GoTo ErrorHandler
    Dim rs As DAO.RecordSet
    Dim db As DAO.Database
    Set db = CurrentDb
    Set rs = db.OpenRecordSet("CT_Options", dbOpenSnapshot)
    Do While Not rs.EOF
        Select Case Nz(rs!OptionName, vbNullString)
        Case "AuditNum"
            MvAuditNum = CLng(Nz(rs!Value, 0))
        Case "AuditPass"
            MvAuditPass = CInt(Nz(rs!Value, 0))
        Case "ClientName"
            MvClientName = Nz(rs!Value, vbNullString)
        Case "Show Audits by Div"
            MvShowAuditByDiv = CInt(Nz(rs!Value, 0))
        End Select
        rs.MoveNext
    Loop
exitHere:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Resume exitHere
End Sub

Public Sub DataSheetStylesDefaults()
    Me.DataSheetStyle = MvDataSheetStyleDefault
End Sub

Public Function UserName() As String
On Error GoTo ErrorHappened
Dim WshNetwork, StReturn As String
'
'StReturn = "Donald.Krupens"
'GoTo ExitNow

    Set WshNetwork = CreateObject("WScript.Network")
    StReturn = WshNetwork.UserName
ExitNow:
    On Error Resume Next
    Set WshNetwork = Nothing
    UserName = StReturn
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Function

Public Function Computer() As String
On Error GoTo ErrorHappened
Dim WshNetwork, StReturn As String
    Set WshNetwork = CreateObject("WScript.Network")
    StReturn = WshNetwork.ComputerName
ExitNow:
    On Error Resume Next
    Set WshNetwork = Nothing
    Computer = StReturn
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Function



Public Function UserSupervisorId() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim rs As ADODB.RecordSet
Dim strUserName As String
Dim oAdo As clsADO

    
    strProcName = "mod_Identity.UserSupervisorId"
    
    
    strUserName = GetUserName
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_ADMIN_Get_User_SupervisorId"
        
        .Parameters.Refresh
        .Parameters("@pUserID") = strUserName
        Set rs = .ExecuteRS
        If .GotData = False Then
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                LogMessage strProcName, "ERROR", "Could not get the Supervisor Id for user: '" & strUserName & "'"
                GoTo Block_Exit
            End If
        End If
    End With
    
    If rs.EOF = True And rs.BOF = True Then
        UserSupervisorId = ""
    Else
        UserSupervisorId = Trim("" & rs("SupervisorId").Value)
    End If
    
Block_Exit:
    Set oAdo = Nothing
    Set rs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function Domain() As String
On Error GoTo ErrorHappened
Dim WshNetwork, StReturn As String
    Set WshNetwork = CreateObject("WScript.Network")
    StReturn = WshNetwork.UserDomain
ExitNow:
    On Error Resume Next
    Set WshNetwork = Nothing
    Domain = StReturn
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbInformation, Application.CodeContextObject.Name
    Resume ExitNow
End Function

Private Sub Class_Initialize()
Me.ReloadOptions
With MvDataSheetStyleDefault
    .BackGroundColor = RGB(255, 255, 255) 'White
    .BorderLineStyle = 1 'Solid
    .CellsEffect = acEffectNormal 'Flat
    .FontFamily = "Arial"
    .FontItalic = False
    .fontsize = 9
    .FontUnderline = False
    .FontWeight = 400 'Normal
    .ForeColor = 0 'Black
    .GridlinesBehavior = acGridlinesBoth 'Both Vertical and Horizontal
    .GridlinesColor = 12632256 'Some Light Color
    .HeaderUnderlineStyle = 1 'Solid
End With
With MvDataSheetStyle
    .BackGroundColor = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "BackGroundColor", MvDataSheetStyleDefault.BackGroundColor)
    .BorderLineStyle = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "BorderLineStyle", MvDataSheetStyleDefault.BorderLineStyle)
    .CellsEffect = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "CellsEffect", MvDataSheetStyleDefault.CellsEffect)
    .FontFamily = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "FontFamily", MvDataSheetStyleDefault.FontFamily)
    .FontItalic = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "FontItalic", MvDataSheetStyleDefault.FontItalic)
    .fontsize = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "FontSize", MvDataSheetStyleDefault.fontsize)
    .FontUnderline = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "FontUnderline", MvDataSheetStyleDefault.FontUnderline)
    .FontWeight = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "FontWeight", MvDataSheetStyleDefault.FontWeight)
    .ForeColor = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "ForeColor", MvDataSheetStyleDefault.ForeColor)
    .GridlinesBehavior = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "GridlinesBehavior", MvDataSheetStyleDefault.GridlinesBehavior)
    .GridlinesColor = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "GridlinesColor", MvDataSheetStyleDefault.GridlinesColor)
    .HeaderUnderlineStyle = Interaction.GetSetting("AuditProbe", "DataSheetStyle", "HeaderUnderlineStyle", MvDataSheetStyleDefault.HeaderUnderlineStyle)
End With
End Sub