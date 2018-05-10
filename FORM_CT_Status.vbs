Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' DLC 11/16/11 - Set focus to an invisible textbox before attempting make any control invisible to prevent error

'THIS IS TO PASS EVENTS TO THE CALLING CLASS
Public Enum StatusConditions
    Idle = 0
    Running = 1
    Canceled = 2
    CanceledAll = 4
End Enum


'***** START API FOR DRAGGING FORM WITHOUT BORDER
Private Type RECT
    x1 As Long
    y1 As Long
    X2 As Long
    y2 As Long
End Type
                                               
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_SYSCOMMAND = &H112
Private Const MOUSE_MOVE = &HF012
'***** END API FOR DRAGGING FORM WITHOUT BORDER

'***** PRIVATE VARIABLES FOR PROPERTIES
Private MvShowCancel As Boolean
Private MvShowCancelAll As Boolean
Private MvShowMessage As Boolean
Private MvShowTime As Boolean
Private MvStatus As Long
Private MvStartTime As Date

Private MvAllMax As Double
Private MvAllVal As Double
Private Mv1Max As Double
Private Mv1Val As Double

Public Sub StatusMessage(data As String)
    Message.Caption = data
End Sub

Public Property Let ProgVal(data As Double)
    Mv1Val = data
End Property

Public Property Get ProgVal() As Double
    ProgVal = Mv1Val
End Property

Public Property Let ProgMax(data As Double)
    Mv1Max = data
End Property

Public Property Get ProgMax() As Double
    ProgMax = Mv1Max
End Property

Public Property Let ProgAllVal(data As Double)
    MvAllVal = data
End Property

Public Property Get ProgAllVal() As Double
    ProgAllVal = MvAllVal
End Property

Public Property Let ProgAllMax(data As Double)
    MvAllMax = data
End Property

Public Property Get ProgAllMax() As Double
    ProgAllMax = MvAllMax
End Property

Public Property Let ShowTime(data As Boolean)
    MvShowTime = data
    If Not data Then txtFocus.SetFocus
    TimeElapsed.visible = MvShowTime
    TimeLeft.visible = MvShowTime
End Property
Public Property Get ShowTime() As Boolean
    ShowTime = MvShowTime
End Property

Public Property Let ShowMessage(data As Boolean)
    MvShowMessage = data
    If Not data Then txtFocus.SetFocus
    Message.visible = MvShowMessage
End Property
Public Property Get ShowMessage() As Boolean
    ShowMessage = MvShowMessage
End Property

Public Property Let ShowCancelAll(data As Boolean)
    MvShowCancelAll = data
    If Not data Then txtFocus.SetFocus
    StatusCaptionTotal.visible = MvShowCancelAll
    CmdCancelAll.visible = MvShowCancelAll
    Pb2.visible = MvShowCancelAll
End Property
Public Property Get ShowCancelAll() As Boolean
    ShowCancelAll = MvShowCancelAll
End Property

Public Property Let ShowCancel(data As Boolean)
    If Not data Then txtFocus.SetFocus
    MvShowCancel = data
    CmdCancel.visible = MvShowCancel
    Pb1.visible = MvShowCancel
End Property
Public Property Get ShowCancel() As Boolean
    ShowCancel = MvShowCancel
End Property
Public Sub Cancel()
    MvStatus = StatusConditions.Canceled
End Sub
Public Sub CancelAll()
    MvStatus = StatusConditions.Canceled Or StatusConditions.CanceledAll
End Sub
Public Sub show()
    With Me
        If MvShowMessage = True Then
            .BOX.Height = .Message.top + .Message.Height + (0.3 * .StatusCaption.top)
            .InsideHeight = .BOX.Height + .BOX.top + 15
            .Pb2.Width = 0
        ElseIf MvShowCancelAll = True Then
            .BOX.Height = .CmdCancelAll.top + .CmdCancelAll.Height + (0.3 * .StatusCaption.top)
            .InsideHeight = .BOX.Height + .BOX.top + 15
            .Pb2.Width = 0
        Else
            .BOX.Height = .CmdCancel.top + .CmdCancel.Height + (0.3 * .StatusCaption.top)
            .InsideHeight = .BOX.Height + .BOX.top + 15
        End If
        .Pb1.Width = 0
        MvStartTime = Now
        .visible = True
    End With
End Sub
Private Sub Box_DblClick(Cancel As Integer)
On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub BOX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        On Error Resume Next
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub
' HC 5/2010 - centralized the mouse capture relase
Private Sub ReleaseMouse()
    Call CaptureRelease
    Call SendWindowMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Private Sub CmdCancel_Click()
    Cancel
End Sub

Private Sub CmdCancelAll_Click()
    CancelAll
End Sub

Public Function EvalStatus(StatusIN As StatusConditions) As Boolean
    Select Case MvStatus
    Case 0
        EvalStatus = (StatusIN = Idle)
    Case Else
        EvalStatus = (StatusIN And MvStatus)
    End Select

End Function

Private Sub cmdHide_Click()
    Me.visible = False
End Sub

Private Sub Form_Open(Cancel As Integer)
    'SET THE DEFAULTS
    Me.ShowCancel = True
    Me.ShowCancelAll = False
    Me.ShowMessage = False
    Me.ShowTime = False
End Sub

Private Sub Form_Timer()
    Dim Dbl As Double
    Dim Pct As Double
    Dim PctAll As Double
    
    If Mv1Max > 0 And Mv1Val > 0 Then
        Pct = (Mv1Val / Mv1Max)
        If Pct > 1 Then
            Pct = 1
        End If
    End If
    
    If MvAllMax > 0 And MvAllVal > 0 Then
        PctAll = (MvAllVal / MvAllMax)
        If PctAll > 1 Then
            PctAll = 1
        End If
    End If
    If MvShowTime = True Then
        Dbl = DateDiff("s", MvStartTime, Now)
        TimeElapsed.Caption = "Elapsed: " & FormatTimeInterval(Dbl)
        If Pct > 0 Then
            Dbl = Dbl * (1 - Pct)
            TimeLeft.Caption = "Remaining: " & FormatTimeInterval(Dbl)
        Else
            TimeLeft.Caption = "Remaining: Div/0"
        End If
    End If
    
    If MvShowCancel = True Then
        If Pct > 0 Then
            Pb1.Width = Me.StatusCaption.Width * Pct
        End If
        StatusCaption.Caption = Format(Pct, "#,##0.00%")
    End If
    
    If MvShowCancelAll = True Then
        If PctAll > 0 Then
            Pb2.Width = Me.StatusCaptionTotal.Width * PctAll
        End If
        StatusCaption.Caption = Format(PctAll, "#,##0.00%")
    End If
    

End Sub
Private Function FormatTimeInterval(DblSecondsPassed As Double) As String
On Error Resume Next
    Dim Dbl As Double
    Dim LngHrs As Double
    Dim LngMinutes As Double
    Dim lngSeconds As Double
    Dbl = DblSecondsPassed
    lngSeconds = Dbl Mod 60
    Dbl = Dbl - lngSeconds
    Dbl = Dbl / 60
    LngMinutes = Dbl Mod 60
    Dbl = Dbl - LngMinutes
    Dbl = Dbl / 60
    LngHrs = Dbl
    FormatTimeInterval = Format(LngHrs, "#,#00") & ":" & Format(LngMinutes, "00") & ":" & Format(lngSeconds, "00")
End Function
Private Sub Message_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        On Error Resume Next
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub

Private Sub StatusCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        On Error Resume Next
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub

Private Sub TimeElapsed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        On Error Resume Next
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub

Private Sub TimeLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        On Error Resume Next
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub

Private Sub TotalStatusCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' HC 5/2010 Call common Routine
    If Button = 1 Then
        ReleaseMouse
    End If
End Sub
