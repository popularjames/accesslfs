Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    Select Case Me.OpenArgs
        Case "Provider"
            'NewProvider "", ""
            NewQuickLook "PROVHDR", "PROVIDERS"
        Case "Claim"
            'NewMain "", "Claim Administration"
            NewQuickLook "AUDITCLM", "CLAIMS"
        Case "Ledger"
            NewQuickLook "COLLMANUAL", "LEDGER Connolly Adjustments"
    
    
    End Select

    DoCmd.Close acForm, Me.Name

End Sub
