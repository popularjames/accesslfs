Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------------------------------------------------------------------------------
' DLC 05/20/2010
' Empty form to record login and logout information and store it in CT_CurrentlyLoggedIn table
'
' DLC 11/16/2011 - Record Startup Telemetry information
' DLC 10/22/2012 - Store a backup of the Telemetry sessionID on the LogoutWatcher Form
' SA 11/1/2012 - Added memory meter using form timer
'-----------------------------------------------------------------------------------------------
Private genUtils As New CT_ClsGeneralUtilities
Private TelemetryStarted As Boolean
Private MaxMemory As Integer

Private Sub Form_Load()
On Error GoTo ErrorHandler
    UpdateUserLog True
    Call SetAppStartup
    
    'Wait for the about screen to close or become invisible
'    Do While FormIsLoadedAndVisible("CT_About")
    Do While IsLoaded("CT_About")   ' 20121211 KD:  I changed this to use the IsLoaded function because it doesn't rely on errors
        DoEvents
    Loop
    
    MaxMemory = Nz(DLookup("Value", "CT_Options", "OptionName='MemoryMeterMax'"), 0)

    Me.TimerInterval = 1
    
    'Open CPCC audit picker
    If IsProductInstalled("ClaimsPlus for Contract Compliance") Then
        DoCmd.OpenForm "CPCC_SelectAudit"
    End If
    
ExitLoad:

Exit Sub
ErrorHandler:
    Resume ExitLoad
End Sub

Private Sub Form_Timer()
'Start telemetry and memory meter
On Error GoTo ErrorHappened
    Dim Mem As Integer

    If Not TelemetryStarted Then
        Telemetry.Startup
        TelemetryStarted = True
    End If
    
    If MaxMemory > 0 Then
        If Not Application.VBE.MainWindow.visible Then
            Mem = GetPagefileUsage
            If Mem > MaxMemory Then
                Mem = MaxMemory
            End If
            SysCmd acSysCmdInitMeter, "Decipher Memory: " & Mem & " mb", MaxMemory
            SysCmd acSysCmdUpdateMeter, Mem
            
            Me.visible = False  'SA - There is something that occasionally makes this form visible.
            Me.TimerInterval = Nz(DLookup("Value", "CT_Options", "OptionName='MemoryMeterTime'"), 2500)
        Else
            'Hide progress bar and stop timer when code window is opened
            SysCmd acSysCmdRemoveMeter
            Me.TimerInterval = 0
        End If
    Else
        Me.TimerInterval = 0
    End If
    
ExitNow:
On Error Resume Next

Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UpdateUserLog
    
    'Record Telemetry Information
    If Not Telemetry Is Nothing Then
        Telemetry.Shutdown
    End If
    
    ' JL - 3/2011 -- Unload all loaded add-ins prior to closing Decipher.  This will release all the allocated resources.
    Dim addInManager As New CT_ClsCnlyAddinSupport
    addInManager.UnloadAddins
    
    ' LG - 6/2012 -- Added to remove reference to the App Manager
On Error Resume Next
    DoCmd.Save acForm, Me.Name
    DoEvents
    RemoveAppManagerRef
End Sub

Private Sub UpdateUserLog(Optional ByVal blnAddUser As Boolean = False)
'This will remove the current user from the CnlyCurrentLoggedIn table and re-add if blnAdd is True
On Error GoTo ErrorHandler
    Dim db As DAO.Database
    'Only attempt to update the table if the database is not readonly
    If Not genUtils.ApplicationIsReadOnly Then
        Set db = CurrentDb
        db.Execute "DELETE FROM CT_CurrentlyLoggedIn WHERE empCurrentlyLoggedIn='" & Replace(GetNetworkUserName(), "'", "''") & "' AND Computer='" & Replace(Identity.Computer, "'", "''") & "'"
        If blnAddUser Then
            db.Execute "INSERT INTO CT_CurrentlyLoggedIn(empCurrentlyLoggedIn, UDate, Computer)VALUES(" & _
                       "'" & Replace(Identity.UserName, "'", "''") & "', Now(), '" & Replace(Identity.Computer, "'", "''") & "')"
        End If
    End If
Exit_ErrorHandler:
On Error Resume Next
    db.Close
    Set db = Nothing
Exit Sub
ErrorHandler:
    'Ignore any errors
    Resume Exit_ErrorHandler
End Sub

Private Function FormIsLoadedAndVisible(ByVal strForm As String) As Boolean
'DLC 11/10/2012 - Added function to see if form is loaded
On Error GoTo ErrorHappened
    Dim blnReturn As Boolean
    If LenB(Forms(strForm).Caption) > -1 Then
        blnReturn = Forms(strForm).visible
    End If
ExitNow:
    On Error Resume Next
    FormIsLoadedAndVisible = blnReturn
    Exit Function
ErrorHappened:
    blnReturn = False
    Resume ExitNow
End Function
