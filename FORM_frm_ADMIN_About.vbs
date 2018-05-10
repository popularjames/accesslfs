Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mbFirstRun As Boolean

Private Sub Form_Close()
    Dim MyAdo As clsADO
   ' Dim Identity As New ClsIdentity
    Dim strSQL As String
    Dim db As Database
    
    
    On Error Resume Next
    
    ' delete entry in the user log.
    Set db = CurrentDb
    strSQL = "delete from ADMIN_LoggedIn_Users where UserID = '" & Identity.UserName & "'"
    db.Execute (strSQL)
    
    LogMessage TypeName(Me) & ".Form_Close", "LOGOUT", Identity.UserName & " logging out", "Version: " & CStr(GetLocalVersionNum) & " on " & Identity.Computer
    
  '  Set Identity = Nothing
    Set db = Nothing
    
End Sub



Private Sub Form_GotFocus()
    Me.TimerInterval = 2500
End Sub



Private Sub Form_Load()
  '  Dim Identity As New ClsIdentity
    Dim strSQL As String
    Dim db As Database
    
    Me.Caption = ""
    
    
    On Error Resume Next

    ' 3.0.1101 CA: 01.16

    Me.lblVersion.Caption = VersionTemplate
    

    ' add entry in the user log.
    Set db = CurrentDb
    db.Execute ("delete from ADMIN_LoggedIn_Users where UserID = '" & Identity.UserName & "'")
    
    strSQL = "insert into ADMIN_LoggedIn_Users(UserID, LoggedIn,Computer,AppPath)" & _
                        " values('" & Identity.UserName & "','" & Now() & "','" & Identity.Computer & "','" & CurrentDb.Name & "')"
    db.Execute (strSQL)

    ' 20130328 KD: Lets log who is logging in on the server so we can see the version and stuff!
    LogMessage TypeName(Me) & ".Form_Load", "LOGIN", Identity.UserName() & " is logging into the 2010 version of claim admin from " & Identity.Computer, "Version " & CStr(mod_ClaimAdmin_Tools.GetSQLServerVersionNum)

    WriteIcon
    SetStartupProperty "AppIcon", dbText, GetClaimAdminIconName
    
    SetApplicationTitle
    Application.RefreshTitleBar
    
    RunScheduledJobs
    
    Me.TimerInterval = 1000
    mbFirstRun = True


  '  Set Identity = Nothing
    Set db = Nothing
    
    
    
End Sub

Private Function GetClaimAdminIconName() As String
On Error GoTo ErrorHappened

Dim oShell ' WScript.Shell
Dim oFso ' as Scripting.FileSystemObject
Dim FileName As String
    
    Set oShell = CreateObject("WScript.Shell")
    
    
    FileName = oShell.SpecialFolders("AppData") & "\Connolly"
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    If oFso.FolderExists(FileName) = False Then
        Call oFso.CreateFolder(FileName)
    End If
    
    FileName = FileName & "\" & gcClaimAdminName & ".ico"
    


    GetClaimAdminIconName = FileName
ExitNow:
    On Error Resume Next
    Set oShell = Nothing
    Set oFso = Nothing
    Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical, TypeName(Me) & ".GetClaimAdminIconName"
    Resume ExitNow
    Resume
    
End Function


Private Sub WriteIcon()
On Error GoTo ErrorHappened
Dim ClIcon As New CT_ClsIcon


    With ClIcon
        .FileName = GetClaimAdminIconName
        .SaveFile gcClaimAdminName
    End With

ExitNow:
    On Error Resume Next
    Set ClIcon = Nothing
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical, "Save Icon File"
    Resume ExitNow
    Resume

End Sub

Private Sub Form_Timer()
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim rsDAO As DAO.RecordSet
    
    Dim strErrMsg As String
    Dim strMessage As String
    Dim db As Database
    Dim strSQL As String
    
    Me.visible = False
    
    
    

    If mbFirstRun Then
    
    
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
        MyAdo.sqlString = "select * from ADMIN_User_Account where UserID = '" & Identity.UserName() & "'"
    
        Set rs = MyAdo.OpenRecordSet
    
        If rs.recordCount = 0 Then
            strErrMsg = "Error: You are not setup on any account." & vbCrLf & vbCrLf & "Please notify your account administrator."
            MsgBox strErrMsg, vbCritical + vbInformation
            gintAccountID = -1
            gstrAcctAbbrev = ""
            gstrAcctDesc = ""
        ElseIf rs.recordCount = 1 Then
            
            gintAccountID = rs("AccountID")
            MyAdo.sqlString = "select * from ADMIN_Client_Account where AccountID = " & gintAccountID
            Set rs = MyAdo.OpenRecordSet
            gstrAcctAbbrev = rs("AcctAbbrev")
            gstrAcctDesc = rs("AcctDesc")
            If gstrAcctDesc = "CMS" Then
                'DPR - ATO REQUIREMENT
                
                ' If this is being opened by another tool do not show the msgbox..
                If Nz(Command, "") = "" Then
                    
                    strMessage = "***WARNING***" & vbCrLf
                    strMessage = strMessage & "(a) This application accesses a U.S. Government information" & vbCrLf
                    strMessage = strMessage & "(b) Users must adhere to U.S. Government Information Security Policies, Standards, and Procedures;" & vbCrLf
                    strMessage = strMessage & "(c) Usage may be monitored, recorded, and audited;" & vbCrLf
                    strMessage = strMessage & "(d) Unauthorized use is prohibited and subject to criminal and civil penalties; and" & vbCrLf
                    strMessage = strMessage & "(e) The use of the information system establishes your consent to any and all monitoring and recording of your activities." & vbCrLf
                    MsgBox strMessage, vbCritical, "WARNING"
                End If
           End If
        Else
            ' user is associated with more than one account
            DoCmd.OpenForm "frm_ADMIN_Account_Selection", acNormal
        End If
        
        If AppAccess_Check_Passive("DashManagement") <> 0 Then
            DoCmd.OpenForm "frm_Dash_Main"
        End If
            
        
        mbFirstRun = False
        
        'JS 08/21/2012 Changing it to 10 minutes instead of 1 due to CA performance issues
        
        Me.TimerInterval = 600000
        
    Else
        Set db = CurrentDb
        
        strSQL = "select * from ADMIN_LoggedIn_Users where UserID = '" & Identity.UserName & "'"
        
        Set rsDAO = db.OpenRecordSet(strSQL)
        If rsDAO("ForcedLogOff") = "Y" Then
            DoCmd.SetWarnings False
            Application.Quit
        ElseIf rsDAO("BroadcastMsg") & "" <> "" Then
            MsgBox rsDAO("BroadCastMsg"), vbCritical, "SYSTEM ALERT"
            Me.TimerInterval = 15000
        End If
    End If
    
    '' 20130328: KD: Added this to check the version and notify the user that they have an old version..
    Call CheckVersion
    
    Set rs = Nothing
    Set rsDAO = Nothing
    Set MyAdo = Nothing
    Set db = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Form_Close
End Sub
