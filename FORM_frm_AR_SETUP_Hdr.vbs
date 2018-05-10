Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AR_SETUP_Hdr
' Author:      Barbara Dyroff
' Create Date: 2010-05-17
' Description:
'   Display the Accounts Receivable Information.  Sort the data.
'
' Modification History:
'   2010-07-13 by BJD TO Add fields for CR6928 FISS, CR6943 VMS and CR6554 MCS.
'   2014-08-19 by BJD to add the txtClmStatus display.
'   2014-12-17 by BJD to add the option to Update Orphan.
'
' =============================================

Const CstrFrmAppID As String = "ARSetupM"

Private frmUpdateOrphanPopup As Form_frm_AR_SETUP_Orphan_Popup


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Private Sub cmdLauchClaim_Click()
On Error GoTo ErrHandler

    Dim strCnlyClaimNum As String
    strCnlyClaimNum = Me.CnlyClaimNum
    NewMain Me.CnlyClaimNum, "Main Claim"

Exit Sub

ErrHandler:
    MsgBox "Error Launching Claim - " & strCnlyClaimNum, vbOKOnly + vbExclamation, "Launch Claim"

End Sub

Private Sub cmdUpdateOrphan_Click()

    On Error GoTo Err_UpdateOrphan_Click
    
    'Call the form to update the AR orphan and research info.
    Set frmUpdateOrphanPopup = New Form_frm_AR_SETUP_Orphan_Popup
    frmUpdateOrphanPopup.CurrentCnlyClaimARID = Me.CnlyClaimARID
    frmUpdateOrphanPopup.RefreshData
    ShowFormAndWait frmUpdateOrphanPopup
  
    Me.Refresh

Exit_UpdateOrphan_Click:
    Set frmUpdateOrphanPopup = Nothing
    Exit Sub

Err_UpdateOrphan_Click:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".cmdUpdateOrphan_Click"
    Resume Exit_UpdateOrphan_Click

End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    
    If Not (IsSubForm(Me)) Then
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
    
    Me.OrderBy = "BatchPrcsDt DESC"
    Me.OrderByOn = True
End Sub
