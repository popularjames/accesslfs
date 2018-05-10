Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AR_SETUP
' Author:      Barbara Dyroff
' Create Date: 2010-05-17
' Description:
'      Display Accounts Receivable (AR) Setup data and maintain AR Setup errors.  The first tab
' displays (display only) the AR Setup records.  The second tab is used for maintaining the
' errors.  Each error records displays the associated error notes.  New notes can be added.  After
' the error is corrected and succesfully processed (by the AR Setup load stored procedure), the
' AR Setup information is available in the first tab -- AR Setup display.
'
'
' Modification History:
'   2010-05-22 by BJD to remove the filter button, add Close button and refresh button.
'   2010-07-13 by BJD TO Add fields for CR6928 FISS, CR6943 VMS and CR6554 MCS.
'   2012-08-10 by BJD to update the Form Caption.
'   2014-07-30 by BJD to add the option to select a set of records.
'
' =============================================

Private Const strAR_SETUP_APP_ID As String = "ARSetupM"  'Used for Note ID
Const CstrFrmAppID As String = "ARSetupM"  'Used for form security

Private frmNoteDetailPopup As Form_frm_Note_Detail_Insert_Popup
Attribute frmNoteDetailPopup.VB_VarHelpID = -1
Private frmARSetupSelectPopup As Form_frm_AR_SETUP_Hdr_Select_Popup
Private frmARSetupErrorFilterPopup As Form_frm_AR_SETUP_Error_Filter_Popup 'bjd change name

Private bolNoteInd As Boolean

Private strARSetupSelect As String 'User selection SQL for the AR Setup Hdr page.
Private strARSetupErrorFilter As String 'User Filter for the AR Setup Error page.

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Let NoteInd(data As Boolean)
    bolNoteInd = data
End Property

Public Property Get NoteInd() As Boolean
    NoteInd = bolNoteInd
End Property

Public Property Get AppID() As String
    AppID = strAR_SETUP_APP_ID
End Property

Public Property Let NoteSeqNo(data As String)
    Me.txtSeqNo = data
End Property

Public Property Let ARSetupSelect(data As String)
    strARSetupSelect = data
End Property

Public Property Get ARSetupSelect() As String
    ARSetupSelect = strARSetupSelect
End Property

Public Property Let ARSetupErrorFilter(data As String)
    strARSetupErrorFilter = data
End Property

Public Property Get ARSetupErrorFilter() As String
    ARSetupErrorFilter = strARSetupErrorFilter
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Private Sub cboProcessFlag_AfterUpdate()
    'Display the ProcessFlag description.
    Me.txtProcessFlagTxt = Me.cboProcessFlag.Column(1)
End Sub

Private Sub cmdLauchClaim_Click()
On Error GoTo ErrHandler

    Dim strCnlyClaimNum As String
    strCnlyClaimNum = Me.CnlyClaimNum
    NewMain Me.CnlyClaimNum, "Main Claim"

Exit Sub

ErrHandler:
    MsgBox "Error Launching Claim - " & strCnlyClaimNum, vbOKOnly + vbExclamation, "Launch Claim"

End Sub

'Add a Note.  Call the popup (modal) form to prompt the user to add a new new.
'The popup form will execute the stored procedure to add the note to Note_Detail.
Private Sub cmdNoteInsert_Click()
    Dim intResult As Integer
    On Error GoTo ErrHandler
    
    'Save any changes before adding a note because the SeqNo on the record needs to be updated.
    If Me.Dirty Then
        intResult = MsgBox("Changes must be saved before adding a Note.  Do you want to save your change?", vbYesNo + vbQuestion)
        If intResult = vbYes Then
            Me.NoteInd = True
            cmdSaveRecord_Click
        Else
            Exit Sub
        End If
    End If
    
    'Call the form to add a note.
    Set frmNoteDetailPopup = New Form_frm_Note_Detail_Insert_Popup
    frmNoteDetailPopup.txtNoteID = Me.NoteID
    frmNoteDetailPopup.txtAppID = Me.AppID
    frmNoteDetailPopup.txtSeqNo = Me.SeqNo + 1
    frmNoteDetailPopup.Refresh
    ShowFormAndWait frmNoteDetailPopup
  
    'Update the SeqNo for the note.
    Me.SeqNo = DLookup("MAX([SeqNo])", "[NOTE_Detail]", "([NoteID] = [txtNoteID])")
    Me.NoteInd = True
    cmdSaveRecord_Click
    Me.Refresh

ErrHandler_Exit:
    Set frmNoteDetailPopup = Nothing
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".cmdNoteInsert_Click"
    Resume ErrHandler_Exit
    
    
End Sub


Private Sub cmdSearchErr_Click()
On Error GoTo Err_cmdSearchErr_Click


    screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_cmdSearchErr_Click:
    Exit Sub

Err_cmdSearchErr_Click:
    MsgBox Err.Description
    Resume Exit_cmdSearchErr_Click
    
End Sub


Private Sub cmdSelect_Click()

    Dim rst As DAO.RecordSet
    Dim strCurrentRecordSource As String
    
    On Error GoTo ErrHandler
    
    If Me.TabCtl0.Value = 0 Then

        ARSetupSelect = ""  ' Init for new selection.
        
        'Hold the last selection.
        If Me.subfrm_AR_SETUP_Hdr.Form.RecordSource = "AR_SETUP_Hdr" Then
            strCurrentRecordSource = "SELECT * FROM AR_SETUP_Hdr ORDER BY BatchPrcsDt DESC"
        Else
            strCurrentRecordSource = Me.subfrm_AR_SETUP_Hdr.Form.RecordSource
        End If
        
        'Call the form
        Set frmARSetupSelectPopup = New Form_frm_AR_SETUP_Hdr_Select_Popup
        Set frmARSetupSelectPopup.FormARSetup = Me.Form

        ShowFormAndWait frmARSetupSelectPopup
        
        'Testing
        'MsgBox ARSetupSelect, vbOKOnly + vbExclamation, "Select Testing"
        
        'Update the Record Source selection.
        If ARSetupSelect <> "" Then
            Me.subfrm_AR_SETUP_Hdr.Form.RecordSource = ARSetupSelect()
            Me.subfrm_AR_SETUP_Hdr.Requery
            
            Set rst = Me.subfrm_AR_SETUP_Hdr.Form.RecordSet
            If rst.recordCount = 0 Then
                MsgBox "No records were selected.  Please make a new selection.  ", vbOKOnly + vbExclamation, "Select No Records Found"
                Me.subfrm_AR_SETUP_Hdr.Form.RecordSource = strCurrentRecordSource
                Me.subfrm_AR_SETUP_Hdr.Requery
            End If

        End If
        
    ElseIf Me.TabCtl0.Value = 1 Then
        
        ARSetupErrorFilter = ""  ' Init for new filter.
                
        'Call the form
        Set frmARSetupErrorFilterPopup = New Form_frm_AR_SETUP_Error_Filter_Popup
        Set frmARSetupErrorFilterPopup.FormARSetup = Me.Form

        ShowFormAndWait frmARSetupErrorFilterPopup
        
        'Testing
'        MsgBox ARSetupErrorFilter, vbOKOnly + vbExclamation, "Filter Testing"
        
        'Update the Record Source selection.
        If ARSetupErrorFilter <> "" Then
            Me.Form.filter = ARSetupErrorFilter()
            Me.Form.FilterOn = True
            Me.Form.Refresh
        End If
    Else
        ' This should not happen.  The Tab Index should be set.
        MsgBox "Please click on your tab page and try again. ", vbOKOnly + vbExclamation, "Select Error"
    End If

ErrHandler_Exit:
    Set frmARSetupSelectPopup = Nothing
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".cmdSelect_Click"
    Resume ErrHandler_Exit


End Sub

Private Sub cmdUnSelect_Click()
    On Error GoTo Err_UnSelect_Click
        
    'Restore the Record Source selection or Filter to include all records in the table.
    If Me.TabCtl0.Value = 0 Then
        Me.subfrm_AR_SETUP_Hdr.Form.RecordSource = "SELECT * FROM AR_SETUP_Hdr ORDER BY BatchPrcsDt DESC"
        Me.subfrm_AR_SETUP_Hdr.Requery
    ElseIf Me.TabCtl0.Value = 1 Then
        Me.Form.filter = ""
        Me.Form.FilterOn = False
        Me.Form.Refresh
    Else
        ' This should not happen.  The Tab Index should be set.
        MsgBox "Please click on your tab page and try again. ", vbOKOnly + vbExclamation, "Unselect Error"
    End If
    
Exit_UnSelect_Click:
    Exit Sub

Err_UnSelect_Click:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".cmdUnSelect_Click"
    Resume Err_UnSelect_Click
    
End Sub

Private Sub Form_Current()
    'Initial New Note Indicator to false.
    Me.NoteInd = False

    'Display the ProcessFlag description.
    Me.txtProcessFlagTxt = Me.cboProcessFlag.Column(1)
End Sub

' Check permissions and set the sort order.
Private Sub Form_Load()
    Dim iAppPermission As Integer
        
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    If iAppPermission <> 0 Then
        Me.OrderBy = "BatchPrcsDt DESC"
        Me.OrderByOn = True
    End If
    
End Sub


Private Sub cmdSaveRecord_Click()
On Error GoTo Err_cmdSaveRecord_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_cmdSaveRecord_Click:
    Exit Sub

Err_cmdSaveRecord_Click:
    MsgBox Err.Description
    Resume Exit_cmdSaveRecord_Click
    
End Sub


Private Sub cmdUndo_Click()
On Error GoTo Err_cmdUndo_Click

    If Me.Dirty Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    
        'Reset the diplay of the ProcessFlag text.
        Me.txtProcessFlagTxt = Me.cboProcessFlag.Column(1)
    End If

Exit_cmdUndo_Click:
    Exit Sub

Err_cmdUndo_Click:
    MsgBox Err.Description
    Resume Exit_cmdUndo_Click
    
End Sub

' Prompt the user to confirm changes to be saved.
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Me.NoteInd = False Then 'Do not confirm after New Note popup.
        If MsgBox("Save changes to the record? ", vbYesNo + vbQuestion, "Confirm Change") = vbNo Then
            cmdUndo_Click
        End If
    Else
        'Reset New Note Indicator
        Me.NoteInd = False
    End If
End Sub


Private Sub cmdCloseForm_Click()
On Error GoTo Err_cmdCloseForm_Click


    DoCmd.Close

Exit_cmdCloseForm_Click:
    Exit Sub

Err_cmdCloseForm_Click:
    MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click
    
End Sub


Private Sub cmdRefresh_Click()
On Error GoTo Err_cmdRefresh_Click

    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70

Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click
    
End Sub
