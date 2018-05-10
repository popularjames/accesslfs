Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mstrSelHeight As Integer
Dim mstrSelWidth As Integer

Public Function WholeRowSelected() As Boolean
    If Me.RecordSource = "" Then
        WholeRowSelected = False
        Exit Function
    End If
    If mstrSelWidth = 11 And mstrSelHeight = 1 Then
        WholeRowSelected = True
    Else
        WholeRowSelected = False
    End If
End Function

Private Sub Form_Click()
    mstrSelHeight = Me.SelHeight
    mstrSelWidth = Me.SelWidth
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    
    'cmdOpen_Click

End Sub

Private Sub Text58_DblClick(Cancel As Integer)
    OpenScreenByField Me.ActiveControl
End Sub

Sub OpenScreenByField(CallingTextBox As TextBox)

'since double click selects the whole field I will keep the same functionality
'which wont work if the ctrl shift C combination is pressed
If Me.ActiveControl = CallingTextBox Then
    CallingTextBox.SelStart = 0
    CallingTextBox.SelLength = Len(CallingTextBox.Value)
End If

'a different action depending of the field associated with the double clicked field
Select Case CallingTextBox.ControlSource
    Case "CnlyClaimNum" 'Open Claim form
        Navigate "frm_RPT_AccessForm", "AUDITCLM", "DblClick", CallingTextBox.Value
    Case "ICN" 'since this is also a claim number I will pass the value from that field. If cnlyclaimnum does not exist it will triger an error
        Navigate "frm_RPT_AccessForm", "AUDITCLM", "DblClick", Me("CnlyClaimNum")
End Select


End Sub
