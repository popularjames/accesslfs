Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    ' JL 6/13/2011 change recordsource grid to not include Decipher config values in CnlyScreenOptions added 12 - 6
    Me.RecordSource = "SELECT * FROM CT_Options WHERE OptionID Not In (4,13,12,6)"
    
    
End Sub

Private Sub Value_AfterUpdate()
'SA 11/27/2012 - Added 'AuditPass' and 'Show Audits by Div' to select case
    Select Case Me.OptionName
        Case "ClientName", "AuditNum", "AuditPass", "Show Audits by Div"
            Me.Dirty = False
            Identity.ReloadOptions
    End Select
End Sub

Private Sub Value_BeforeUpdate(Cancel As Integer)
Select Case Me.DataType
Case 0 'Text
Case 1, 4 ' INT, Decimal
    If "" & Me.Value = "" Then Exit Sub 'BLANKS ARE OKAY
    If IsNumeric(Me.Value) = False Then
        MsgBox "Numeric Value Expected for Option (" & Me.OptionName & ")", vbInformation, "Option Validation Error"
        Cancel = True
        Exit Sub
    Else
        If Me.DataType = 1 And CDbl(CLng(Me.Value)) <> CDbl(Me.Value) Then
            MsgBox "Numeric Value Expected for Option (" & Me.OptionName & ") - INT ONLY", vbInformation, "Option Validation Error"
            Cancel = True
            Exit Sub
        End If
    End If
Case 2 'Date
End Select
End Sub
