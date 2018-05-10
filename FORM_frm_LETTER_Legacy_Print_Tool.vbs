Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'Private Const cs_LETTER_TYPE_COMBO_SQL As String = "SELECT DISTINCT LetterType, LetterDesc FROM Letter_Type WHERE AccountId = 1 AND Uses2DBarcodes = 1 "
Private Const cs_LETTER_TYPE_COMBO_SQL As String = "SELECT DISTINCT LT.LetterType, LT.LetterDesc FROM Letter_Type LT WITH (NOLOCK) INNER JOIN  LETTER_Work_Queue WQ WITH (NOLOCK) ON LT.LetterType = WQ.LetterType " & _
        " WHERE LT.AccountId = 1  AND WQ.ProcessedDt > DATEADD(week, 1, WQ.RowCreateDt)"


Private csSelectedLetterType As String
Private clCurrentLetter As Long
Private cbStopNow As Boolean
Private cbPaused As Boolean
Private cdtStarted As Date

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Let Status(sStatus As String)
    Me.txtStatus = sStatus
    Me.txtStatus.visible = True
End Property

Public Property Get StopNow() As Boolean
    StopNow = cbStopNow
End Property
Public Property Let StopNow(bStop As Boolean)
    cbStopNow = bStop
End Property


Public Property Get Paused() As Boolean
    Select Case UCase(Me.cmdPause.Caption)
    Case "PAUSE"
        cbPaused = False
    Case "CONTINUE"
        cbPaused = True
    Case Else
        Stop
    End Select
    Paused = cbPaused
End Property
Public Property Let Paused(bPause As Boolean)
    cbPaused = bPause
End Property


Public Property Get MaxToProcess() As Long
    MaxToProcess = Me.prgbStatus.max
End Property
Public Property Let MaxToProcess(lMaxToProcess As Long)
    Me.prgbStatus.max = lMaxToProcess
End Property


Public Property Get CurrentLetterNum() As Long
    CurrentLetterNum = clCurrentLetter
End Property
Public Property Let CurrentLetterNum(lCurrentLetter As Long)
    clCurrentLetter = lCurrentLetter
    Me.prgbStatus.Value = lCurrentLetter
End Property



Public Property Get SelectedLetterType() As String
    If Nz(Me.cmbLetterType, "") = "" Then
        csSelectedLetterType = GetSetting("LRT_DEFAULT_LETTER_TYPE")
    End If
    SelectedLetterType = csSelectedLetterType
End Property
Public Property Let SelectedLetterType(sLetterType As String)
    csSelectedLetterType = sLetterType
    Me.cmbLetterType.Text = sLetterType
        '' SelectComboBoxItemFromText
End Property


Private Sub cmbLetterType_AfterUpdate()
On Error GoTo Block_Err
Dim strProcName As String
Dim sDefFolder As String

    strProcName = ClassName & ".cmbLetterType_AfterUpdate"

    sDefFolder = QualifyFldrPath(GetSetting("MAILROOM_LETTER_PATH")) & Format(Now(), "yyyy-mm-dd") & "\" & Me.SelectedLetterType

    Me.txtFolderToProcess = sDefFolder
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdBrowse_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sSelFolder As String


    strProcName = ClassName & ".cmdBrowse_Click"
    Me.txtProvNums = ""
    
    Me.txtProvderResult_Match.visible = False
    
    sSelFolder = Me.txtFolderToProcess
    
    sSelFolder = BrowseForFolderMSOFfice("Select the folder containing the letters to reconcile", sSelFolder, msoFileDialogViewDetails)
    
    If sSelFolder = "" Then
        ' canceled:
        GoTo Block_Exit
    End If
    
    Me.txtFolderToProcess = sSelFolder
    
    Call txtFolderToProcess_AfterUpdate
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sDefaultLetterType As String


    strProcName = ClassName & ".RefreshData"
    sDefaultLetterType = Me.SelectedLetterType
    
    ' load the letter combo box
    Call RefreshComboBoxADO(cs_LETTER_TYPE_COMBO_SQL, Me.cmbLetterType, sDefaultLetterType, "LetterType", "v_Data_Database")
    
    ' refresh the 'Results stuff...
    
    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub cmdPause_Click()
    If Me.TimerInterval > 0 Then
        Select Case UCase(Me.cmdPause.Caption)
        Case "PAUSE"
            Me.cmdPause.Caption = "Continue"
        Case "CONTINUE"
            Me.cmdPause.Caption = "Pause"
        Case Else
            Stop
        End Select
    Else
        Me.cmdPause.visible = False
    End If
End Sub

Private Sub cmdStart_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim lProvCount As Long
Dim sLetterType As String
Dim sLtrReqDt As String
Dim oFso As Scripting.FileSystemObject
Dim oRegEx As RegExp

    strProcName = ClassName & ".cmdStart_Click"
    Set oFso = New Scripting.FileSystemObject
    
    If Me.cmdStart.Caption = "Cancel" Then
        Me.StopNow = True
        Me.cmdPause.visible = False
        Me.prgbStatus.visible = False
        Me.cmdStart.Caption = "Start"
        GoTo Block_Exit
    End If

    
    If FolderExists(Me.txtFolderToProcess) = False Then
        GoTo Block_Exit
    End If
    
    
    sLetterType = Me.txtFolderToProcess

    sLtrReqDt = oFso.GetParentFolderName(sLetterType)

    
    Set oRegEx = New RegExp
    With oRegEx
        .IgnoreCase = True
        .Global = False
        .Pattern = "^(.+?[\\\/])([^\\\/]+?)$"
    End With
    
    sLetterType = oRegEx.Replace(sLetterType, "$2")
    sLtrReqDt = oRegEx.Replace(sLtrReqDt, "$2")
    
    Me.cmdPause.Caption = "Pause"
    Me.cmdPause.visible = True
    
    ' clear up stuff here first:
    Me.StopNow = False
    Me.cmdStart.Caption = "Cancel"
    DoCmd.Hourglass True
    Me.TimerInterval = 500
    Me.txtProvNums = 0
    Me.prgbStatus.max = CLng(Me.txtLettersFound)
    
    Me.txtProvderResult_Match.visible = False
    
    Me.prgbStatus.visible = True
    Me.prgbStatus.Value = 1
    Me.txtStatus.visible = True
    
    cdtStarted = Now()
    
    If PrintFolder(Me.txtFolderToProcess, sLetterType, sLtrReqDt, Me) = True Then
        lProvCount = Me.txtLettersFound
    End If
    Me.TimerInterval = 0
    
    Call SleepEvents(2)
    
    
'    lProvCount = DLookup("CountOfProvNum", "qry_LETTER_Reconciliation_Tool_Counts")
    
    Me.txtProvNums = lProvCount
    
    If Trim(Me.txtProvNums) = Trim(Me.txtLettersFound) Then
        Me.txtProvderResult_Match.BackColor = RGB(0, 255, 0)
    Else
        Me.txtProvderResult_Match.BackColor = RGB(255, 0, 0)
        Me.Status = "There is at least 1 duplicate in that folder! Please contact Biz Ops to investigate and fix!"
    End If
    Me.txtProvderResult_Match.visible = True
    
    Me.prgbStatus.visible = False
    Me.cmdStart.Caption = "Start"

    
Block_Exit:
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String
Dim sDefFolder As String


    strProcName = ClassName & ".Form_Load"
    
    Me.txtProvderResult_Match.visible = False
    Me.prgbStatus.visible = False
    Me.prgbStatus.Value = 1
    Me.txtProvNums = ""

    Call RefreshData
    
    ' get default selections
    
    If Nz(Me.txtFolderToProcess, "") = "" Then
        sDefFolder = QualifyFldrPath(GetSetting("MAILROOM_LETTER_PATH")) & Format(Now(), "yyyy-mm-dd") & "\" & Me.SelectedLetterType
        
        Me.txtFolderToProcess = sDefFolder
    
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Timer()
Dim sElapsedTime As String
Dim sPctDone As String

    DoEvents
    DoEvents
    
    sElapsedTime = ProcessTookHowLong(cdtStarted, Now()) & " elapsed"
    
    sPctDone = Format(CDbl(Me.prgbStatus.Value / Me.prgbStatus.max), "00.00 %")  ' & "%"
    Me.Status = " Processing " & CStr(Me.prgbStatus.Value) & " of " & CStr(Me.prgbStatus.max) & " (" & sPctDone & ") " & sElapsedTime
    
    DoEvents
    DoEvents
    
End Sub

Public Function UpdateStatus(Optional sMsg As String) As Boolean
Dim sElapsedTime As String
Dim sPctDone As String

    DoEvents
    DoEvents
    
    If sMsg <> "" Then
        Me.Status = sMsg
    Else
        sElapsedTime = ProcessTookHowLong(cdtStarted, Now()) & " elapsed"
        
        sPctDone = Format(CDbl(Me.prgbStatus.Value / Me.prgbStatus.max), "00.00 %")  ' & "%"
        Me.Status = " Processing " & CStr(Me.prgbStatus.Value) & " of " & CStr(Me.prgbStatus.max) & " (" & sPctDone & ") " & sElapsedTime
    End If
    
    
End Function

Private Sub txtFolderToProcess_AfterUpdate()
    If FolderExists(Me.txtFolderToProcess) = True Then
        Me.txtLettersFound = CStr(GetFileCountFromDir(Me.txtFolderToProcess, False))
        MaxToProcess = CLng(Me.txtLettersFound)
    End If
    
End Sub
