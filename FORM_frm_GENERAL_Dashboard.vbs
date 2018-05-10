Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub cmdAppMaintenance_Click()
    DoCmd.OpenForm "frm_ADMIN_Main", acNormal
End Sub
Private Sub cmdDecipher_Click()

    Dim strAppPath As String
    
    strAppPath = DLookup("AppPath", "GENERAL_DecipherPath")
    
    Shell "explorer.exe " & strAppPath, vbMinimizedFocus

End Sub

Private Sub cmdLaunchPool_Click()
    DoCmd.OpenForm "frm_POOL_Main", acNormal
End Sub

Private Sub cmdLetterMaintenance_Click()
    DoCmd.OpenForm "frm_LETTER_MAIN", acNormal
End Sub

Private Sub cmdNewSearch_Click()
    NewMainSearch "", "", ""
End Sub

Private Sub cmdOpenClaim_Click()
    NewMain "", "Claim Administration"
End Sub

Private Sub cmdQueueManagement_Click()
    DoCmd.OpenForm "frm_QUEUE_Main", acNormal
End Sub
Private Sub cmdUserMaintenance_Click()
    DoCmd.OpenForm "frm_ADMIN_User_Main", acNormal
End Sub

Private Sub cmdImageMaintenance_Click()
    DoCmd.OpenForm "frm_SCANNING_Main", acNormal
End Sub

Private Sub cmdProviderMaint_Click()
    NewProvider "", ""
End Sub
