Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'DLC 05/19/10 Replaced all code related to monitoring logged in users with the launch of scrLogoutWatcher in the Form_Load

Private Sub Form_Open(Cancel As Integer)
    ' 20130418 KD: Added this to make sure that when we implement the Version control
    ' our startup forms don't do anything..
    If Application.UserControl = True Then
        lstInstalledApps.RowSource = "SELECT ProductName,LocalVersion,MaturityDesc FROM CT_InstalledApps ORDER BY ProductName"
    Else
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    'DLC 05/24/2010 Changed the version color to make it visible in TS environment
    LblVer2.ForeColor = 12615680

    Me.LblVer2.Caption = VersionTemplate
    Me.visible = True
    On Error Resume Next
    If CInt(Nz(Me.OpenArgs, 0)) <> 1 Then
         Me.TimerInterval = 1500
        'Loading the logout watcher with broken references will prevent users from closing the database
        If HasBrokenReference Then
            RemoveBrokenReferences
            If Not HasBrokenReference Then
                'Resolved issue, carry on as normal
                DoCmd.OpenForm "CT_LogoutWatcher", , , , , acHidden
            Else
                MsgBox "Please resolve the reference issue before using this application"
            End If
        Else
            DoCmd.OpenForm "CT_LogoutWatcher", , , , , acHidden
        End If
    End If
    
    
        
    'DPR CMS - check if we are on CMS
    If InStr(Environ$("computername"), "CMS") > 0 Then
        DoCmd.OpenForm "frm_ADMIN_About"
    End If


End Sub

Private Sub Form_Timer()
    CloseAbout
End Sub

Private Sub UserClickedMe()
    If Me.TimerInterval > 0 Then
        'Keep about screen open until clicked again
        Me.TimerInterval = 0
    Else
        'Close now
        CloseAbout
    End If
End Sub

Private Sub CloseAbout()
    'Stop timer and close form
    Me.visible = True
    Me.TimerInterval = 0
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


''''''' Form click events below here ''''''''

Private Sub LblCopyright_Click()
    UserClickedMe
End Sub

Private Sub LblCopyright_DblClick(Cancel As Integer)
    CloseAbout
End Sub

Private Sub Logo_Click()
    UserClickedMe
End Sub

Private Sub Logo_DblClick(Cancel As Integer)
    CloseAbout
End Sub

Private Sub Detail_Click()
    UserClickedMe
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
    CloseAbout
End Sub
