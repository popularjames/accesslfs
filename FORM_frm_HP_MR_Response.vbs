Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const ColorRed = 8421631
Const ColorNormal = -2147483643

Private Sub Form_Current()
    Me.AllowEdits = False
'    If Me.AckCode.Value <> "1" Then
'        Me.AckCode.BackColor = ColorRed
'        Me.AckDesc.BackColor = ColorRed
'    Else
'        Me.AckCode.BackColor = ColorNormal
'        Me.AckDesc.BackColor = ColorNormal
'    End If
    
    'Me.ClaimStatusDisplay = Me.ClmStatus & " - " & Me.ClmStatusDesc
End Sub
