Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'
' DLC 06/03/2010 - Replacement Progress Bar
' ------------------------------------------------------------------
' Usage: Add CnlyProgress SubForm (in this example called progDone)
'
' Private Progress As Form_CT_Progress
' Set Progress = progDone.Form
' Progress.PercentComplete = n
'
' to change the color, use:
'    Progress.BarColor = RGB(255, 0, 255)
'    Progress.TextColor = RGB(255, 255, 255)
'
Private Const BORDER As Integer = 50
Private mvComplete As Integer
Private mvCurrentValue As Integer
Private mvBarColor As Long
Private mvTextColor As Long

Public Property Let PercentComplete(Value As Integer)
    If Value >= 0 And Value <= 100 Then
        mvComplete = Value
        'Don't bother redrawing if the percentage did not change.
        'Ideally, this would be done by the calling form to improve performance.
        If mvComplete <> mvCurrentValue Then
            RedrawProgressBar
            mvCurrentValue = mvComplete
        End If
    Else
        Err.Raise 7102, "CnlyProgress", "Complete must be the percentage complete as an integer in the range 0 to 100"
    End If
End Property

Public Property Let BarColor(Value As Long)
    mvBarColor = Value
    lblBack.BackColor = mvBarColor
    lblProgressComplete.BackColor = mvBarColor
End Property

Public Property Let TextColor(Value As Long)
    mvTextColor = Value
    lblProgressComplete.ForeColor = mvTextColor
End Property

Public Property Get PercentComplete() As Integer
    PercentComplete = mvComplete
End Property

Public Property Get BarColor() As Long
    BarColor = mvBarColor
End Property

Public Property Get TextColor() As Long
    TextColor = mvTextColor
End Property

Private Sub Form_Load()
    mvCurrentValue = 0
    mvComplete = 0
    mvBarColor = lblProgressIncomplete.BackColor
    lblBack.left = BORDER
    lblProgressComplete.left = BORDER
    lblProgressIncomplete.left = BORDER
    lblProgressComplete.LeftMargin = BORDER
    lblProgressIncomplete.LeftMargin = BORDER
    lblBack.top = BORDER
    Form_Resize
End Sub

Private Sub Form_Resize()
    With Me.Form
        lblProgressComplete.top = (.InsideHeight - lblProgressComplete.Height) \ 2     'integer divide
        lblProgressIncomplete.top = (.InsideHeight - lblProgressIncomplete.Height) \ 2 'integer divide
        lblBack.Height = .InsideHeight - (BORDER * 2)
        lblProgressIncomplete.Width = .InsideWidth - (BORDER * 2)
        RedrawProgressBar
    End With
End Sub

Private Sub RedrawProgressBar()
    'Set the captions on both labels
    lblProgressComplete.Caption = mvComplete & "% Complete"
    lblProgressIncomplete.Caption = lblProgressComplete.Caption
    'This requires 2 labels to prevent the caption word wrapping when lblProgressComplete is not wide enough to display the whole message
    lblProgressComplete.Width = (Me.InsideWidth - (BORDER * 2)) / 100 * mvComplete
    lblBack.Width = lblProgressComplete.Width
    'This hides the background which is still visible when width = 0
    lblBack.visible = (mvComplete > 0)
    lblProgressComplete.visible = (mvComplete > 0)
End Sub
