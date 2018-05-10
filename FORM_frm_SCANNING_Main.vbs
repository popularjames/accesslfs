Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const CstrFrmAppID As String = "ImageMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Private Sub cmdCDAutoLoad_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_CDLoad_Main", , , , , , Me.Name
End Sub

Private Sub cmdClaimLookUp_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Claim_Lookup_ByName", , , , , , Me.Name
End Sub

Private Sub cmdCountPDF_Click()
    Me.visible = False
    
    
    If gintAccountID = 1 Then
        ' TK 5/4/2011 This only apply to CMS
        
    Else
        ' TK 5/4/2011  This applies to all MCR accounts
        DoCmd.OpenForm "frm_SCANNING_PDF_PageCount", , , , , , Me.Name
    End If
End Sub

Private Sub cmdDataEntry_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Image_Log (PHILLY)", , , , , , Me.Name
End Sub

Private Sub cmdErrorRpt_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Error", , , , , , Me.Name

End Sub

Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub cmdFastScanFixNoMatch_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_FastScan_IssuesMain", , , , , , Me.Name
End Sub

Private Sub cmdFastScanMatch_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_FastScan_Main", , , , , , Me.Name
End Sub

Private Sub cmdFastScanPrint_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_FastScan_PrintCoverSheets", , , , , , Me.Name
End Sub

Private Sub cmdFastScanScan_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_FastScan_ScanCoverSheets", , , , , , Me.Name
End Sub

Private Sub cmdLetterLookUp_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Claim_Lookup_By_InstanceID", , , , , , Me.Name
End Sub

Private Sub cmdQuickImageLog_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Quick_Image_Log", , , , , , Me.Name

End Sub

Private Sub cmdQuickImageLogUpdate_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Quick_Image_Log_Update", , , , , , Me.Name
End Sub

Private Sub cmdSpotCheck_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_Item_Validation", , , , , , Me.Name
End Sub

Private Sub cmdTrashValidation_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_TrashReport", , , , , , Me.Name
End Sub

Private Sub cmdUpdatePageCount_Click()
    Me.visible = False
    DoCmd.OpenForm "frm_SCANNING_UpdatePageCount", , , , , , Me.Name
End Sub

Private Sub cmdValidation_Click()
    Me.visible = False
    
    If gintAccountID = 1 Then
        ' TK 5/4/2011 This only apply to CMS
        DoCmd.OpenForm "frm_SCANNING_Image_Validation", , , , , , Me.Name
    Else
        ' TK 5/4/2011  This applies to all MCR accounts
        DoCmd.OpenForm "frm_SCANNING_Image_Validation (MCR)", , , , , , Me.Name
    End If
End Sub

Private Sub cmdValidationRpt_Click()
    Me.visible = False
    On Error GoTo Show_Main_Form
    
    DoCmd.OpenReport "rpt_Validated_MRs", acViewPreview, , , , Me.Name
    Exit Sub

Show_Main_Form:
    Me.visible = True
    
End Sub







Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Me.Caption = "Scanning"
    screen.MousePointer = 0
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    ' 5/10/2011 TK: Disable this feature for CMS. This functionality is integrated with the "Process data entry"
    If gintAccountID = 1 Then
        Me.cmdCountPDF.visible = False
    End If
    
End Sub
    
