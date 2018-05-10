Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AUDITCLM_Main
' Description:
'   Main Audit Claim maintenance form.
'
' Modification History:
'   20130723 KD: Added some properties to deal with Therapy (Congress) concepts where an auditor needs to
'       select the appropriate error code from a drop down.  Also changed the AUDITCLM_ReviewChart form which
'       is where the auditor sets the error code..
'   2011-11-17 by Barbara Dyroff to display the Adjusted To ICN from the last AR received for a demand.  Added a hidden field to
'       the form to still get the Adj_To from the from the Audit Claim record for the cross reference. (See upper right side of window for the field.)
'   2009-12-18 by Barbara Dyroff to truncate the ConceptDesc in the ComboBox to solve a
'       a property too long error.
'
'
' =============================================


Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1
Private WithEvents myAuditClaim As clsAUDITCLM
Attribute myAuditClaim.VB_VarHelpID = -1
Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private frmIncompleteMRRequest As Form_frm_Incomplete_MR_Request

'TKL 3/2/2011: auditor credit override
Private WithEvents frmAuditorCreditOverride As Form_frm_AUDITCLM_Auditor_Credit_Override
Attribute frmAuditorCreditOverride.VB_VarHelpID = -1

Private frmAUDITTracking As Form_frm_AUDIT_TRACKING_Main
Private mrsConcept As ADODB.RecordSet

Private mstrCnlyClaimNum As String
Private mstrUserProfile As String
Private mstrUserName As String
Private mstrConceptCategory As String
Private miAppPermission As Integer

Private mbAllowChange As Boolean
Private mbAllowDelete As Boolean

'Private variable to keep track of the claim status
Private strClaimStatus As String
Private mReturnDate As Date
Private mbRecordLocked As Boolean
Public mbRecordChanged As Boolean
Private mdtMaxThreshold As Date

'Private variables to keep track of adj values
Private mstrAdj_ReimbAmt
Private mstrAdj_ProjectedSavings
Private mstrAdj_DRG

'Private variables to keep track of original ConceptID
Private mstrConceptID

Private mstrLastUpDt As Date

'TKL 3/2/2011: Auditor credit override
Private mstrOverrideAuditor As String
Private mstrCreditOverrideReason As String
Private mCreditAssignment As CreditAssignment
Private mstrPrevClaimAuditor As String


Private miListItmSelected As Integer

Enum CreditAssignment
    NotAssigned = 0
    AutoAssignment = 1
    Override = 2
End Enum

' 7/16/2013 : KD: Added ErrorCode string for Prepay therapy stuff which is set by the frm_AUDITCLM_ReviewChart form
'   This is so we don't have to check the database again for it when we go to validate whether the claim can be saved or not
Private cstrErrorCode As String
Private cblnIsTherapyClaim As Boolean

Const CstrFrmAppID As String = "AuditClm"

'Added per Mona's design changes
Const DME As Integer = 1
Const DME_1454 = 2
Const DME_0851 = 3
Const STANDARD = 4
Const HH = 5
Const OP_CARR = 6
Const SNF = 7

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get ErrorCodePrpty() As String
    ErrorCodePrpty = cstrErrorCode
End Property
Public Property Let ErrorCodePrpty(strErrorCode As String)
    cstrErrorCode = strErrorCode
    'VS 11/17/15 Allow user to clear error code by clearing ErrorCode drop down menu - part of RVC effort
    If strErrorCode <> "" Then
            RecordChanged = True
        Else
            If Not myAuditClaim.rsAuditClmHdrAdditionalInfo Is Nothing Then
                If strErrorCode = "" And Nz(myAuditClaim.rsAuditClmHdrAdditionalInfo.Fields("ErrorCode"), "") <> "" Then
                    RecordChanged = True
                End If
            End If
    End If
End Property


Public Property Get IsTherapyConcept() As Boolean
    IsTherapyConcept = cblnIsTherapyClaim
End Property
Public Property Let IsTherapyConcept(blnIsTherapyConcept As Boolean)
    cblnIsTherapyClaim = blnIsTherapyConcept
End Property


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
    Me.txtAppID = CstrFrmAppID
End Property

Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property

Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property

Property Let CnlyClaimNum(data As String)
    mstrCnlyClaimNum = data
    Me.txtCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = mstrCnlyClaimNum
End Property

Property Let RecordLocked(data As Boolean)
    mbRecordLocked = data
    If mbRecordLocked Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Property

'TKL 3/2/2011: auditor credit override
Private Sub Adj_Auditor_AfterUpdate()
    If Me.Adj_Auditor & "" = myAuditClaim.ClaimAuditor Then
        mCreditAssignment = NotAssigned
        mstrCreditOverrideReason = ""
    ElseIf mCreditAssignment <> AutoAssignment Then
        If Me.Adj_Auditor & "" <> mstrPrevClaimAuditor Then
            ' credit override
            Set frmAuditorCreditOverride = New Form_frm_AUDITCLM_Auditor_Credit_Override
            ShowFormAndWait frmAuditorCreditOverride
            Set frmAuditorCreditOverride = Nothing
        End If
    End If
    mstrPrevClaimAuditor = Me.Adj_Auditor & ""
End Sub

Private Sub Adj_Bic_AfterUpdate()
    ' TK 10/4/2013 validate Adj_BIC must be less than or equal to 2 characters
    If Len(Me.Adj_Bic.Value) > 2 Then
        MsgBox "Error: Adj_BIC must be 9 or less characters. Please re-enter the correct Adj_BIC ", vbOKOnly + vbCritical
        Me.Adj_Bic.Value = ""
    End If
End Sub

Private Sub Adj_Can_AfterUpdate()
    ' TK 10/4/2013 validate Adj_CAN must be less than or equal to 9 characters
    If Len(Me.Adj_Can.Value) > 9 Then
        MsgBox "Error: Adj_CAN must be 9 or less characters. Please re-enter the correct Adj_CAN ", vbOKOnly + vbCritical
        Me.Adj_Can.Value = ""
    End If
        
End Sub

Private Sub Adj_ConceptID_BeforeUpdate(Cancel As Integer)
    
    myAuditClaim.rsAuditClmDtl.MoveFirst
    myAuditClaim.rsAuditClmDtl.Find "Adj_ConceptId <> ''", , adSearchForward
    
    If Not (myAuditClaim.rsAuditClmDtl.EOF) Then
    
        If MsgBox("Changing the Concept ID at the Claim level will overwrite any codes associated with Detail lines.  Proceede?", vbYesNo) = vbYes Then
           
           myAuditClaim.SyncConceptCodes ClmDetail
            
            If Nz(Me.Adj_ProjectedSavings, 0) <> 0 Then
            
                If MsgBox("Clear Projected Savings?", vbQuestion + vbYesNo, "Clear Projected Savings") = vbYes Then
                   Me.Adj_ProjectedSavings = Null
                End If
            End If

        Else
            Me.Adj_ConceptID.Undo
            Cancel = True
        End If
    
    End If

End Sub

Private Sub Adj_ProjectedSavings_AfterUpdate()
    Me.Adj_ReimbAmt = Me.ReimbAmt - Me.Adj_ProjectedSavings
End Sub

Private Sub Adj_ReimbAmt_AfterUpdate()
    Me.Adj_ProjectedSavings = Me.ReimbAmt - Me.Adj_ReimbAmt
End Sub


Private Sub cmdAction_Click()
    'Launches another action based on what is selected
    'Damon 05/08
    On Error GoTo ErrHandler
    Dim strError As String
    Dim aParameters
    Dim rst As DAO.RecordSet

    
    Dim strSQL As String
    Dim strFunction As String
    Dim strFunctionResult As String
    Dim strParameterName As String
    Dim lngi As Long
    
    If Nz(Me.cboAction, "") = "" Then
        Exit Sub
    End If
    
    
    'Get the Access function for the selected action in the AUDITCLM_Action table
    strSQL = " select * from GENERAL_Action where FormName = '" & Me.Name & "' and AutoID = " & Me.cboAction & " "
    
    Set rst = CurrentDb.OpenRecordSet(strSQL)
    
    If Not rst.EOF Then
        strFunction = rst!Function
        strParameterName = ""
        
        'Multiple parameters can be passed to a function
        'In the table, they are a comma delimited list.  Here, we split them into an array
        'and build an arguement
        'ACTION function variables must be passed as strings.  Do your conversion once they are in
        aParameters = Split(rst!ParameterName, ",")
        For lngi = 0 To UBound(aParameters)
            If lngi = 0 Then
                strParameterName = strParameterName & "'" & Me.Controls(aParameters(lngi)) & "'"
            Else
                strParameterName = strParameterName & ",'" & Me.Controls(aParameters(lngi)) & "'"
            End If
        Next lngi
               
        'The EVAL() function calls the Action.  The "StrFunction" variable has to point to a public module somewhere in Decipher
        'These exist in mod_AUDITCLM_Action
        strFunctionResult = Eval(strFunction & "(" & strParameterName & ")")
        RefreshMain
                
    End If
    
Block_Exit:

    Set rst = Nothing
Exit Sub
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : cmdAction_Click"
    GoTo Block_Exit
End Sub

Private Sub cmdAdj_ChartReviewDt_Click()
    On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.Adj_ChartReviewDt, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.Adj_ChartReviewDt = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub

Private Sub cmdAdj_From_Click()
    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.Adj_From, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.Adj_From = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click


End Sub

Private Sub cmdAdj_To_Click()
    On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.Adj_To, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.Adj_To = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click


End Sub

Private Sub cmdddNote_Click()
Dim noClaimNotes As String
    
    On Error GoTo Err_cmdddNote_Click
    
    'Alex C 3/8/2012 - Added restriction on adding notes for customer service people, per Kathy Bingnear
    noClaimNotes = Nz(DLookup("NoClaimNotes", "CUST_Security_User", "UserID = '" & mstrUserName & "'"), "")
    If noClaimNotes = "y" Then
        MsgBox mstrUserName & ", please use the Customer Service Screen to add notes to a claim."
        Exit Sub
    End If
    
     Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
    
     frmGeneralNotes.frmAppID = Me.frmAppID
     Set frmGeneralNotes.NoteRecordSource = myAuditClaim.rsNotes
     frmGeneralNotes.RefreshData
     ShowFormAndWait frmGeneralNotes
     'lstTabs_Click
     Set frmGeneralNotes = Nothing

Exit_cmdddNote_Click:
    Exit Sub

Err_cmdddNote_Click:
    MsgBox Err.Description
    Resume Exit_cmdddNote_Click
End Sub

Private Sub cmdExit_Click()

    On Error GoTo HandleError
   
    DoCmd.Close acForm, Me.Name
    
exitHere:
    Exit Sub
HandleError:
'* Error 2501 will be caused by canceling the form close.

    If Err.Number = 2501 Then
        Resume Next
    Else
        MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
        GoTo exitHere
    End If
End Sub

Private Sub cmdLaunchTab_Click()
    'Launches a new window displaying the data selected in the tab list box
    'Damon 05/08
    On Error GoTo ErrHandler
    Dim strError As String
    Dim lngTabID As Long
    Dim strSQL As String
    Dim strTabName As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    
    'Get the ID of the currently selected tab
    lngTabID = Me.cboTabs
    strSQL = GetNavigateTabSQL(lngTabID, Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
    NewMainTab strSQL, mstrCnlyClaimNum, strTabName

Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : cmdLaunchTab_Click"
End Sub
Private Sub cmdOpen_Click()
    Dim strCnlyClaimNum As String
    If myAuditClaim.ClaimExists Then
        If Me.Dirty Or Me.NewRecord Or mbRecordChanged Then
            If MsgBox("Record has changed.  Do you want to save it first?", vbYesNo) = vbYes Then
                SaveData
            End If
        End If
        
        If myAuditClaim.LockedForEdit Then
            myAuditClaim.UnlockClaim
        Else
            myAuditClaim.UnlockClaim (True)
        End If
    End If
    
    strCnlyClaimNum = InputBox("Enter Connolly Claim ID.")
    
    If StrPtr(strCnlyClaimNum) <> 0 Then 'If StrPtr function returns 0, then the user pressed cancel
        If strCnlyClaimNum <> "" Then
            'Me.Dirty = False
            mbRecordChanged = False
            
            Me.CnlyClaimNum = strCnlyClaimNum
            Me.LoadData
        Else
            MsgBox "You entered an invalid Connolly Claim ID."
        End If
    End If
    
End Sub


Private Sub cmdRollbackClaimStatus_Click()
Dim strTempCnlyClaimNum As String

strTempCnlyClaimNum = Me.CnlyClaimNum

 ' check and assign operation mode (user vs manager)

'Check if user is a manager
'    If Me.LockUserID <> Identity.Username Then
'JS 03/28/2013 There is already a limitation in the stored procedure to not allow rollbacks on statuses from other users and the 24h cap so I am removing this mballowdelete limitation to allow users to be able to rollback their own mistakes
    'If mbAllowDelete = True Then
        
        If MsgBox("Are you sure you wish to rollback the status?", vbYesNo) = vbYes Then
            
            If myAuditClaim.RollbackStatus() = True Then
                MsgBox "Status was rolled back!", vbInformation, "Rollback successful!"
                mbRecordChanged = False
                Me.CnlyClaimNum = strTempCnlyClaimNum
                Me.LoadData
                Me.lblAppTitle.Caption = "Claim Administration"
            Else
                MsgBox "The status was not able to be rolled back at this time.", vbOKOnly, "Unsuccessful rollback!"
            End If
        Else
            MsgBox "Rollback Cancelled!", vbCritical
        End If
'    Else
'        MsgBox "You do not appear to have permission to rollback the status.", vbInformation
'        Exit Sub
'    End If
    
End Sub

'4/23/2012 - DPR - using this setting to allow a user to manuyally unlock a claim
Private Sub cmdUnlock_Click()
Dim strTempCnlyClaimNum As String

strTempCnlyClaimNum = Me.CnlyClaimNum

'Check if user is a manager
    If myAuditClaim.LockedForEdit = False And Me.LockUserID <> Identity.UserName Then
        
        If MsgBox("Are you sure you wish to release the lock on this claim?", vbYesNo) = vbYes Then
            
            If myAuditClaim.UnLockClaimForce(True) Then
                MsgBox "Claim Unlocked!", vbInformation
                mbRecordChanged = False
                Me.CnlyClaimNum = strTempCnlyClaimNum
                Me.LoadData
                Me.lblAppTitle.Caption = "Claim Administration"
            End If
        Else
            MsgBox "Unlock Cancelled!", vbCritical
        End If
    Else
        MsgBox "Claim is not currently locked", vbInformation
        Exit Sub
    End If
    
End Sub
'4/23/2012 - DPR - using this setting to allow a user to manuyally unlock a claim

'Alex C 2/12/2012 - added for launching Customer Service for this claim
Private Sub Command305_Click()
    lngEventID = 0
    LaunchNewCustClaimEvent (CnlyClaimNum)
End Sub

Private Sub Command309_Click()
Stop
End Sub

Private Sub Form_Close()
    If myAuditClaim.ClaimExists Then
        If myAuditClaim.LockedForEdit = True Then
            myAuditClaim.UnlockClaim
        Else
            myAuditClaim.UnlockClaim (True)
        End If
    End If
    RemoveObjectInstance Me
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub

Private Sub Form_Load()
Dim strSQL As String
    
    Me.Caption = "Claim Processing"
    Me.RecordSource = ""
    
    Me.txtAppID = CstrFrmAppID
    
    IsTherapyConcept = False
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
    Set myAuditClaim = New clsAUDITCLM
            
    'Setting the APPID hardcoded for now to test other functionality
    mstrUserName = Identity.UserName
    mstrUserProfile = GetUserProfile()

    mbAllowChange = (miAppPermission And gcAllowChange)
    
    '4/23/2012 - DPR - using this setting to allow a user to manuyally unlock a claim
    mbAllowDelete = (miAppPermission And gcAllowDelete)
           
    If mbAllowDelete Then
        Me.cmdUnlock.Enabled = True
        Me.cmdRollbackClaimStatus.Enabled = True
    Else
        Me.cmdUnlock.Enabled = False
        '' Since they aren't either admin or mgr then we need to see if they were the last one to update the thing
        '' (and maybe if it's within 24 hours
        If AllowNonMgrToRollback = True Then
            Me.cmdRollbackClaimStatus.Enabled = True
        Else
            Me.cmdRollbackClaimStatus.Enabled = False
        End If
    End If
    '4/23/2012 - DPR - using this setting to allow a user to manuyally unlock a claim
           
    If mbAllowChange = False Then
        Me.cmdddNote.Enabled = False
        Me.cmdSave.Enabled = False
    End If
    
    'initial setup of boxes
    
    ' JS 08/21/20102 Original pre changes to switch it to ADO where possible due to CA performance issues
'    RefreshComboBox "SELECT AutoID, ActionName FROM GENERAL_Action WHERE FormName = '" & Me.Name & "'", Me.cboAction
'    RefreshListBox "SELECT RowID, TabName FROM GENERAL_Tabs WHERE AccessForm = '" & Me.Name & "'", Me.lstTabs
'    RefreshComboBox "SELECT RowID, TabName FROM GENERAL_Tabs WHERE Launch is null AND AccessForm = '" & Me.Name & "'", Me.cboTabs
'    RefreshComboBox "SELECT reviewtype, reviewtypedesc FROM XREF_ReviewType", Me.Adj_ReviewType, "", "ReviewType"
'    RefreshComboBox "SELECT vulnID, VulnDesc FROM XREF_Vulnerability", Me.Adj_VulnerabilityCd, "", "vulnID"
'
'    'RefreshComboBox "SELECT ConceptID, ConceptID + ' - ' + Left(ConceptDesc, 30) FROM Concept_Hdr where AccountID = " & gintAccountID, Me.Adj_ConceptID, "", "ConceptID"
'    StrSQl = "SELECT ConceptID, ConceptID + ' - ' + Left(ConceptDesc, 30), ConceptGroup, ConceptStatus FROM Concept_Hdr where AccountID = " & gintAccountID
'    Me.Adj_ConceptID.RowSource = StrSQl
'    Me.Adj_ConceptID.Requery
'
'    'DPR 7/25/2012 - changed as this is set afterwards and we do not want to incur the cost
'    'RefreshComboBox "SELECT S.ClmStatus, S.ClmStatusDesc FROM XREF_ClaimStatus S ", Me.ClmStatus, "", "ClmStatus"
'
'    ' TKL 3/1/2011: auditor credit override
'    mCreditAssignment = NotAssigned
'    mstrCreditOverrideReason = ""
'    mstrPrevClaimAuditor = myAuditClaim.ClaimAuditor
'    Me.Adj_Auditor.RowSource = "select UserID from ADMIN_User_Account where AccountID = " & gintAccountID
    
    
    Dim rst As ADODB.RecordSet
    Dim myCode_ADO As clsADO
    
    'Dim StrSQl As String
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_DATA_Database")
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strSQL = "SELECT RowID, TabName FROM cms_auditors_claims.dbo.GENERAL_Tabs WHERE AccessForm = '" & Me.Name & "'"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    RefreshListBoxFromRecordset rst, Me.lstTabs
    
    RefreshComboBox "SELECT AutoID, ActionName FROM GENERAL_Action WHERE FormName = '" & Me.Name & "'", Me.cboAction
    
    strSQL = "SELECT RowID, TabName FROM cms_auditors_claims.dbo.GENERAL_Tabs WHERE Launch is null AND AccessForm = '" & Me.Name & "'"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    RefreshComboBoxFromRecordset rst, Me.cboTabs
    
    strSQL = "SELECT reviewtype, reviewtypedesc FROM cms_auditors_claims.dbo.XREF_ReviewType"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    RefreshComboBoxFromRecordset rst, Me.Adj_ReviewType, "", "ReviewType"
    
    strSQL = "SELECT vulnID, VulnDesc FROM cms_auditors_claims.dbo.XREF_Vulnerability"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    RefreshComboBoxFromRecordset rst, Me.Adj_VulnerabilityCd, "", "vulnID"
    
    'JS 20121106 took the load of the concept from here to the LoadData sub
    
    'Set rst = mycode_ado.OpenRecordSet(StrSQl)
    'RefreshComboBoxFromRecordset rst, Me.Adj_ConceptID, "", "ConceptID"
    
    ' TKL 3/1/2011: auditor credit override
    mCreditAssignment = NotAssigned
    mstrCreditOverrideReason = ""
    mstrPrevClaimAuditor = myAuditClaim.ClaimAuditor
    
    strSQL = "select UserID from ADMIN_User_Account where AccountID = " & gintAccountID
    Set Me.Adj_Auditor.RecordSet = myCode_ADO.OpenRecordSet(strSQL)
    
    
    
    'RefreshMain
   
    Me.Detail.visible = False
    cmdSave.Enabled = False
    
    If Me.CnlyClaimNum <> "" Then LoadData

End Sub

Public Sub RefreshMain()
    'Refresh the main form
    'This form has control names that match the column names in the AuditCLM_Hdr table
    'Only textboxes and comboboxes with a TAG property of "R" are filled
    'To add a field to this form
    '   Add the column to the data table
    '   Create a textbox (or combobox) on the form wit the same name as the column name in the table
    '   Set the control's tag property to "R"
    On Error GoTo ErrHandler
    
    Dim strError As String
    Dim rst As ADODB.RecordSet
    Dim strSQL As String
    Dim Field As Field
    Dim ctl As Control
    
    IsTherapyConcept = False
    
    'Keep load data seperate from form refresh
        
    Me.Caption = gstrAcctDesc & ": ClaimNum : " & mstrCnlyClaimNum
    
    '2014:03:12:JS Prepay claims for Concept Cm_C2019 are being imported without ReimbAmt, need to allow editing this field only for this concept
    Dim bolAllowEditReimbAmt As Boolean
    bolAllowEditReimbAmt = False
    If Not (myAuditClaim.rsAuditClmDtl.EOF And myAuditClaim.rsAuditClmDtl.BOF) Then
        myAuditClaim.rsAuditClmDtl.MoveFirst
        myAuditClaim.rsAuditClmDtl.Find "Adj_ConceptId = 'CM_C2019'", , adSearchForward
    End If
    If Not (myAuditClaim.rsAuditClmDtl.EOF) Or myAuditClaim.rsAuditClmHdr("Adj_ConceptID").Value = "CM_C2019" Then
        bolAllowEditReimbAmt = True
    End If
    
    If myAuditClaim.ClaimExists Then
        Me.Detail.visible = True
        Set Me.Form.RecordSet = myAuditClaim.rsAuditClmHdr
        'Loop through the controls setting their control source to the recordset
        'JS 20121107 Added the possibility to include buttons to the controls marked with R, they will be enabled or disabled (i.e. Rollback button)
        For Each ctl In Me.Controls
            If ctl.Tag = "R" Then
                If Controls(ctl.Name).ControlType = acCommandButton Then
                    If myAuditClaim.LockedForEdit = False And mbAllowChange Then
                        Me.Controls(ctl.Name).Enabled = False
                    ElseIf myAuditClaim.LockedForEdit = True And mbAllowChange Then
                        Me.Controls(ctl.Name).Enabled = True
                    End If
                Else
                    Me.Controls(ctl.Name).ControlSource = myAuditClaim.rsAuditClmHdr.Fields(ctl.Name).Name
                    If myAuditClaim.LockedForEdit = False And mbAllowChange Then
                        Me.Controls(ctl.Name).Locked = True
                    ElseIf myAuditClaim.LockedForEdit = True And mbAllowChange Then
                        If Me.Controls(ctl.Name).Name = "ReimbAmt" And Not bolAllowEditReimbAmt Then
                            Me.Controls(ctl.Name).Locked = True
                        Else
                            Me.Controls(ctl.Name).Locked = False
                        End If
                    End If
                End If
            End If
        Next
        
        
    '------------------------------------------------------------------------------------------------
    '---- JS 20121102
    '---- getting the available claims statuses was a DAO call,
    '---- now it is ADO and uses the function "cms_auditors_code.dbo.udf_clmstatus_available"
    '------------------------------------------------------------------------------------------------
        
    Dim myCode_ADO As clsADO
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_DATA_Database")
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strSQL = "select ClmStatus, ClmStatusDesc from cms_auditors_code.dbo.udf_clmstatus_available('" & myAuditClaim.CnlyClaimNum & "', " & gintAccountID & ")"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    
    'update the combobox
    RefreshComboBoxFromRecordset rst, Me.ClmStatus
    Me.ClmStatusCode = Nz(Me.ClmStatus, "")
    
    '------------------------------------------------------------------------------------------------
        
        
'        'get allowable statuses for the claim status combo box
'        strSQL = " SELECT X.ClmStatus, X.ClmStatusDesc " & _
'                    " FROM XREF_ClaimStatus as X INNER JOIN AUDITCLM_Process_Logics as P " & _
'                    "   ON X.ClmStatus = P.NextStatus " & _
'                    " WHERE P.CurrStatus = '" & myAuditClaim.rsAuditClmHdr("ClmStatus") & "'" & _
'                    " AND P.DataType = '" & myAuditClaim.rsAuditClmHdr("DataType") & "'" & _
'                    " AND P.AccountID = " & gintAccountID & _
'                    " AND P.ProcessModule = 'AUDITCLM'" & _
'                    " AND P.ProcessType = 'Manual'" & _
'                    " AND P.ReviewType = '" & myAuditClaim.rsAuditClmHdr("Adj_ReviewType") & "'" & _
'                    " AND (X.isSubonly = 0 or x.clmStatus = '" & Me.ClmStatus & "') " 'DPR 4/23/2012
'
'
'
'        RefreshComboBox strSQL, Me.ClmStatus, "", "ClmStatus"
        
 

        If lstTabs.Column(1, 0) = "_Alerts" And FormExists("frm_AUDITCLM_Alerts") Then lstTabs.Selected(0) = True
        
        
        If Me.lstTabs.ListIndex > -1 Then
            lstTabs_Click
        End If
           
        mstrAdj_ProjectedSavings = Me.Adj_ProjectedSavings
        mstrAdj_ReimbAmt = Me.Adj_ReimbAmt
        mstrAdj_DRG = Me.Adj_DRG
        mstrLastUpDt = Me.LastUpDt
        mstrConceptID = Me.Adj_ConceptID
           
           
        'Save the current claim status, so we can validate the claim before saving
        'strCurrentStatus = Me.ClmStatus (thieu)
    Else
        Me.Detail.visible = False
        Set Me.Form.RecordSet = Nothing
        Exit Sub
    End If
    
    
    
    cmdSave.Enabled = False
    cmdddNote.Enabled = False
    If myAuditClaim.LockedForEdit = False And mbAllowChange Then
        Me.lblAppTitle.Caption = "Claim Administration" & " - Locked by " & myAuditClaim.LockedUser
    ElseIf myAuditClaim.LockedForEdit = True And mbAllowChange Then
        cmdSave.Enabled = True
        cmdddNote.Enabled = True
    End If
    
Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : RefreshMain"
End Sub
Private Sub CheckProviderStatus()
    Dim strCnlyProvID As String
    Dim strProvStatus As String
    Dim strHPProviderStatus As String
    Dim strStatusDesc As String
    
    Dim strCustServiceUser As String
    
    
    
    If myAuditClaim.rsAuditClmHdr.EOF <> True And myAuditClaim.rsAuditClmHdr.BOF <> True Then
        
        
        'This was replaced with the alerts tab JS 06/28/2012
        'Dim strMsg As String
        'strMsg = ""
        'strCnlyProvID = myAuditClaim.rsAuditClmHdr("CnlyProvID")
        'strProvStatus = DLookup("Status", "PROV_Hdr", "CnlyProvID = '" & strCnlyProvID & "'") & ""
        'If strProvStatus <> "01" Then
        '    strStatusDesc = DLookup("Description", "PROV_Xref_Status_Code", "ProvStatus = '" & strProvStatus & "'") & ""
        '    If strStatusDesc <> "" Then
        '        strMsg = "The provider is '" & strStatusDesc & "'"
        '    End If
        'End If
        'If strMsg <> "" Then
        '    strMsg = "ALERT: " & vbCrLf & strMsg
        '    MsgBox strMsg, vbCritical
        'End If
        
        
        'Alex C 02/19/2012 - For Customer Service only, see if this provider is active with HealthPort - if so, make the label visible
        'that indicates that the medical record is likely coming electronically form HealthPort, in case of delays caused by HP
        strCustServiceUser = Nz(DLookup("UserID", "CUST_Security_User", "UserID = '" & mstrUserName & "'"), "none")
        If strCustServiceUser <> "none" Then
            strHPProviderStatus = DLookup("CurStatus", "HP_Providers", "CnlyProvID = '" & strCnlyProvID & "'") & ""
       
            If UCase(strHPProviderStatus) = "ACTIVE" Then
                lblHPeMR.visible = True
            Else
                lblHPeMR.visible = False
            End If
        End If
    End If
    

End Sub
Private Sub CheckClaimOwner()
    
    Dim strLOB As String
        
    If myAuditClaim.rsAuditClmHdr.EOF <> True And myAuditClaim.rsAuditClmHdr.BOF <> True Then
        strLOB = Nz(myAuditClaim.rsAuditClmHdr("LOB"), "")
        If strLOB <> "" Then
            MsgBox "ALERT: the claim belongs to '" & strLOB & "'", vbCritical
        End If
    End If

End Sub

Private Sub Form_Resize()
    If Me.lstTabs.ListIndex <> -1 Then
        miListItmSelected = Me.lstTabs.ListIndex
    End If
    ResizeControls Me.Form
End Sub


Private Sub Form_Timer()
    Me.TimerInterval = 0
        '' 20130205 KD: Very annoying fix to very annoying issue Microsoft introduced with MS Access 2007
        '' subforms bound to an ADO recordset fire Form_Unload then Form_Close
        '' when the main form is minimized (but after the main form's Resize event)
        '' also, the list box looses the selected item
    If miListItmSelected > -1 Then
        Me.lstTabs.Selected(miListItmSelected) = True
        Call lstTabs_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)


    If myAuditClaim Is Nothing Then
        GoTo exitHere
    End If

    If myAuditClaim.ClaimExists = False Then
        GoTo exitHere
    End If
    
    If Me.RecordSource = "" Then
        GoTo exitHere
    End If
    
    If mbAllowChange = False Or (Me.Dirty = False And mbRecordChanged = False) Then
        GoTo exitHere
    End If
    
    '* JC The SaveData function contains the save confirmation.  I don't like it there, but don't want to restructure the proc now.
    'SaveData
    
    If mbRecordChanged Or Me.Dirty Then
        If MsgBox("Record has changed.   Would you like to save it?", vbYesNo) = vbYes Then
            SaveData
      
        End If
    End If
        
exitHere:
    Exit Sub
End Sub

'TKL 3/2/2011: Auditor credit override
Private Sub frmAuditorCreditOverride_OverrideReason(OverrideReason As String, Cancel As Boolean)
    If Cancel Then
        Me.Adj_Auditor = mstrPrevClaimAuditor
    Else
        mstrCreditOverrideReason = OverrideReason
        mCreditAssignment = Override
    End If
End Sub


Private Sub lstTabs_Click()
'Tues 2/5/2013 by KCF - Add code to handle new module for DME Documentation Review \ Rationale template
    Dim strSQL As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    Dim strNoteIDs As String
    Dim MyAdo As clsADO
'    Dim rs As DAo.Recordset         '' KD 20120910 - Change to ADO
    Dim rs As ADODB.RecordSet
    Dim oAdo As clsADO        '' KD 20120910 - Change to ADO
    
    
    If Not myAuditClaim.rsAuditClmHdr Is Nothing And lstTabs.ListIndex >= 0 Then
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("v_Data_Database")
            .SQLTextType = sqltext
            .sqlString = GetListBoxRowSQL(lstTabs.Column(0), Me.Name)
            Set rs = .ExecuteRS
        End With
        
        '' /  KD 20120910 - Change to ADO
    
    
'        Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstTabs.Column(0), Me.Name), dbOpenSnapshot, dbSeeChanges)        '' KD 20120830 - Change to ADO
        If Not (rs.BOF And rs.EOF) Then
        
            'check if form exists first
            If Not FormExists(rs("FormName")) Then
                MsgBox "The form: '" & rs("FormName") & "' " & vbNewLine & _
                        "does not exist in your copy of Claim Admin." & vbNewLine & vbNewLine & _
                        "You cannot click on '" & lstTabs.Column(1) & "' until you contact IT to have it updated.", vbExclamation, "LstTabs_Click Error"
                Exit Sub
            End If
            
            Me.subFrmMain.visible = True
            Select Case rs("FormName")
            Case "frm_CONCEPT_Hdr"
                If Nz(Me.Adj_ConceptID, "") <> "" Then
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
                    MyAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.Adj_ConceptID & "'"
                    Set mrsConcept = MyAdo.OpenRecordSet()
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Me.subFrmMain.Form.FormConceptID = Me.Adj_ConceptID
                    
                    Set Me.subFrmMain.Form.RecordSet = mrsConcept
                    Set Me.subFrmMain.Form.ConceptRecordSource = mrsConcept
                    
                    Me.subFrmMain.Form.RefreshData
                    Me.subFrmMain.Form.Controls("ConceptLevel").SetFocus
                    Me.subFrmMain.Form.Controls("cmdSave").Enabled = False
                    Set MyAdo = Nothing
                  End If
            Case "frm_GENERAL_Notes_Display"
                If rs("TabName") = "Queue Notes" Then
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    
'                    StrSQl = "select n.* from NOTE_Detail n join QUEUE_Dtl q on n.NoteID = q.NoteID " & _
'                             " where q.CnlyClaimNum = '" & Me.CnlyClaimNum & "' " & _
'                             " order by n.NoteID asc, n.SeqNo"
                             
                    strSQL = "select rb.* ,nd.* " & _
                    "From CMS_AUDITORS_code.dbo.v_queue_notes_w_rollback rb left join CMS_AUDITORS_Claims.dbo.NOTE_Detail nd on rb.noteid = nd.NoteID " & _
                    "where rb.CnlyClaimNum = '" & Me.CnlyClaimNum & "' and nd.noteid is not null order by rb.lastupdate, rb.seqno, nd.SeqNo"
                    
                    MyAdo.sqlString = strSQL
                    Set Me.subFrmMain.Form.NoteRecordSource = MyAdo.OpenRecordSet
                    Set MyAdo = Nothing
                ElseIf rs("TabName") = "Provider Notes" Then
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    
                    strSQL = "select n.* from NOTE_Detail n join PROV_Hdr p on n.NoteID = p.NoteID " & _
                             " where p.CnlyProvID = '" & Me.cnlyProvID & "' " & _
                             " order by n.SeqNo asc"
                    MyAdo.sqlString = strSQL
                    Set Me.subFrmMain.Form.NoteRecordSource = MyAdo.OpenRecordSet
                    Set MyAdo = Nothing
                'BEGIN: QA Notes tab option added KCF Tues 10/2/2012 to show the results of Claim QA
                ElseIf rs("TabName") = "QA Summary" Then
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    
                    strSQL = "Select * from cms_auditors_code..v_QA_NotesToDisplay QA " & _
                             "where QA.CnlyClaimNum = '" & Me.CnlyClaimNum & "' " & _
                             "Order by NoteDate desc"
                             
                    MyAdo.sqlString = strSQL
                    Set Me.subFrmMain.Form.NoteRecordSource = MyAdo.OpenRecordSet
                    Set MyAdo = Nothing
                'END: QA Notes tab option added KCF Tues 10/2/2012 to show the results of Claim QA
                Else
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Set Me.subFrmMain.Form.NoteRecordSource = myAuditClaim.rsNotes
                End If
                Me.subFrmMain.Form.RefreshData
            Case "frm_AUDITCLM_ReviewChart"
                Me.subFrmMain.SourceObject = rs("FormName")
                Set Me.subFrmMain.Form.DiagCodeRecordsource = myAuditClaim.rsAuditClmDiag
                Set Me.subFrmMain.Form.ProcCodeRecordsource = myAuditClaim.rsAuditClmProc
                Set Me.subFrmMain.Form.DiagCodeRevRecordsource = myAuditClaim.rsAuditClmDiagRev
                Set Me.subFrmMain.Form.ProcCodeRevRecordsource = myAuditClaim.rsAuditClmProcRev
                Set Me.subFrmMain.Form.HdrRecordsource = myAuditClaim.rsAuditClmHdr
                Set Me.subFrmMain.Form.HdrAddInfoRecordsource = myAuditClaim.rsAuditClmHdrAdditionalInfo '5/13/2013 TK getting additional_info data
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.DRG = Me.DRG 'to be DRG Andrew Assigns Surg
                Me.subFrmMain.Form.RefreshData
            
            
            Case "frm_AUDITCLM_RELATED"
                Me.subFrmMain.SourceObject = rs("FormName")
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.RefreshData
            
            
            
            Case "frm_AUDITCLM_Main_Dtl"
                Me.subFrmMain.SourceObject = rs("FormName")
                'Me.subFrmMain.SetFocus
                Set Me.subFrmMain.Form.DtlRecordSource = myAuditClaim.rsAuditClmDtl
                Me.subFrmMain.Form.RefreshData
            Case "frm_NOTE_Detail_GridView"
                Me.subFrmMain.SourceObject = rs("FormName")
                Set Me.subFrmMain.Form.NoteRecordSource = myAuditClaim.rsNotes
            
            Case "frm_QUEUE_Exception_Info_Grid_View"
                Me.subFrmMain.SourceObject = rs("FormName")
                Me.subFrmMain.Form.FormFilter = "cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
                Me.subFrmMain.Form.RefreshData
            Case "frm_AUDITCLM_ClaimsPlus"
                Me.subFrmMain.SourceObject = rs("FormName")
                Set Me.subFrmMain.Form.ClaimsPlusRecordsource = myAuditClaim.rsAuditClmClaimsPlus
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.RefreshData
            
            Case "frm_AUDIT_TRACKING_Main"
                Set frmAUDITTracking = New Form_frm_AUDIT_TRACKING_Main
                    
                frmAUDITTracking.AppTitle = "Change history for claim : " & CnlyClaimNum
                frmAUDITTracking.AuditTableName = "AUDITCLM_Hdr_Audit_Hist"
                frmAUDITTracking.AuditKey = "CnlyClaimNum = '" & CnlyClaimNum & "'"
                frmAUDITTracking.RefreshData
                ColObjectInstances.Add frmAUDITTracking, frmAUDITTracking.hwnd & ""
                frmAUDITTracking.visible = True
            Case "frm_AUDITCLM_Rationale"
                Me.subFrmMain.SourceObject = rs("FormName")
                Set Me.subFrmMain.Form.HdrRecordsource = myAuditClaim.rsAuditClmHdr
                Me.subFrmMain.Form.RefreshData
                
                
                
            Case "frm_AUDITCLM_Hdr_Additional_Info"
                Select Case gintAccountID
                    Case 1
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Me.subFrmMain.Form.filter = "cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
                        Me.subFrmMain.Form.FilterOn = True
                    Case 2
                    Case 3
                    Case 4
                        Me.subFrmMain.SourceObject = "frm_AMERIGROUP_AUDITCLM_Hdr_AdditionalInfo"
                        Set Me.subFrmMain.Form.AdditionalHdrInfoRecordSource = myAuditClaim.rsAuditClmHdrAdditionalInfo
                        Me.subFrmMain.Form.RefreshData
                End Select
            Case "frm_APPEAL_TIMELINE" 'andrew
                Me.subFrmMain.SourceObject = rs("FormName")
                If rs("FormName") = "frm_AUDITCLM_References_Grid_View" Then
                    Me.subFrmMain.Form.FieldReference = "cnlyClaimNum"
                    Me.subFrmMain.Form.FieldValue = Me.CnlyClaimNum
                End If
                Dim frm_Appeal_Time As Form_frm_APPEAL_TIMELINE
                Set frm_Appeal_Time = Me.subFrmMain.Form
                frm_Appeal_Time.CnlyClaimNum = mstrCnlyClaimNum
                ''
                      
                strSQL = GetNavigateTabSQL(lstTabs.Column(0), Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
                If strSQL <> "" Then
                    Me.subFrmMain.Form.CnlyRowSource = strSQL
                End If
                Me.subFrmMain.Form.RefreshData
            Case "frm_AUDITCLM_Alerts"
                Me.subFrmMain.SourceObject = "frm_AUDITCLM_Alerts"
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.RefreshData
            'BEGIN: 2/5/2013 Frm_AuditClm_DocRevie_DME by KCF
            Case "frm_AuditClm_DocReview_DME"
                Me.subFrmMain.SourceObject = "frm_AuditClm_DocReview_DME"
                Me.subFrmMain.Form.txtCnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.DocReview_DME_ReviewSetup
                Me.subFrmMain.Form.Rationale_Populate
                Set Me.subFrmMain.Form.HdrRecordsource = myAuditClaim.rsAuditClmHdr
            'END: 2/5/2013 Frm_AuditClm_DocRevie_DME by KCF
            
            'BEGIN: 3/6/2014 Frm_AuditClm_RulesEngine by KCF
            Case "frm_AuditClm_RulesEngine"
                Me.subFrmMain.SourceObject = "frm_AuditClm_RulesEngine"
                Me.subFrmMain.Form.txtCnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.CnlyClaimNum = Me.CnlyClaimNum
                Me.subFrmMain.Form.txtDataType = Me.DataType
                'Me.subFrmMain.Form.DataType = Me.DataType
                Me.subFrmMain.Form.AuditClm_RulesEngine_Eligibility
                Set Me.subFrmMain.Form.HdrRecordsource = myAuditClaim.rsAuditClmHdr
            'END: 3/6/2014 Frm_AuditClm_RulesEngine by KCF
            
            'MG add screen scraping audit trail for R'Lay
            Case "frm_SS_Output_FISS_Audit_Trail"
                Me.subFrmMain.SourceObject = "frm_SS_Output_FISS_Audit_Trail"
                
                'MG filters are executed in the subform as it will only display the claim data filtered on
                'Executing code here will work, but users will see all claims before it gets filtered unless if I add additional VBA in subform, but why bother?
                
                'Dim sqlString As String
                'sqlString = " SS_CnlyClaimNum = " & Chr(34) & Me.CnlyClaimNum & Chr(34)
                'Me.subFrmMain.Form.filter = sqlString
                'Me.subFrmMain.Form.FilterOn = True
                'Me.subFrmMain.Form.Requery
                'Me.subFrmMain.Form.Refresh
            '12/10/2013 MG add provider mr extension info for auditors to see
            Case "frm_PROV_MR_Extension_Info"
                Me.subFrmMain.SourceObject = "frm_PROV_MR_Extension_Info"
            '01/08/2014 MG add exception history
            Case "frm_QUEUE_Exception_Hist"
                Me.subFrmMain.SourceObject = "frm_QUEUE_Exception_Hist"
            Case Else
                Me.subFrmMain.SourceObject = rs("FormName")
                If rs("FormName") = "frm_AUDITCLM_References_Grid_View" Then
                    Me.subFrmMain.Form.FieldReference = "cnlyClaimNum"
                    Me.subFrmMain.Form.FieldValue = Me.CnlyClaimNum
                End If
                
                
                strSQL = GetNavigateTabSQL(lstTabs.Column(0), Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
                If strSQL <> "" Then
                    Me.subFrmMain.Form.CnlyRowSource = strSQL
                End If
                Me.subFrmMain.Form.RefreshData
            End Select
            
        'JS 20121107 not allowing changes in the subforms on a locked claim
        If myAuditClaim.LockedForEdit = False And mbAllowChange Then
            Me.subFrmMain.Locked = True
        ElseIf myAuditClaim.LockedForEdit = True And mbAllowChange Then
            Me.subFrmMain.Locked = False
        End If
            
        Else
            MsgBox "Application form has not been defined"
        End If
    End If
    
Block_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close '' KD 20120830 - Change to ADO
    End If
    Set rs = Nothing
    Set oAdo = Nothing
End Sub
Private Sub cmdSave_Click()
    Sleep 1000
    SaveData
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Err_cmdSearch_Click
    NewMainSearch "AUDITCLM", "v_AUDITCLM_Hdr", "Claims"
    
Exit_cmdSearch_Click:
    Exit Sub
Err_cmdSearch_Click:
    MsgBox Err.Description
    Resume Exit_cmdSearch_Click
End Sub
Private Sub frmGeneralNotes_NoteAdded()
    mbRecordChanged = True
End Sub
Public Sub LoadData()
    Dim bLoaded As Boolean
    bLoaded = myAuditClaim.LoadClaim(Me.CnlyClaimNum, mbAllowChange)
    
    cmdSave.Enabled = False
    cmdddNote.Enabled = False
    
    If myAuditClaim.ClaimExists Then
        If mbAllowChange Then
            If myAuditClaim.LockedForEdit = False Then
                'this has been replaced by the alerts tab JS 06/28/2012
                'MsgBox "Record is being locked by " & myAuditClaim.LockedUser & " at " & myAuditClaim.LockedDate
            ElseIf mbAllowChange Then
                cmdSave.Enabled = True
                cmdddNote.Enabled = True
            End If
        End If
    Else
        MsgBox "Claim '" & CnlyClaimNum & "' does not exist for this account"
    End If
    
    'JS 20121106 had to move the fill of the Concept combo box here because I need the cnlyclaimnum
    Dim myCode_ADO As clsADO
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_DATA_Database")
    myCode_ADO.sqlString = "SELECT ConceptID, ConceptShow, ConceptGroup, ConceptStatus FROM cms_auditors_code.dbo.udf_GetAvailableConceptsnOriginal('" & CnlyClaimNum & "','" & gintAccountID & "')"
    Set Me.Adj_ConceptID.RecordSet = myCode_ADO.OpenRecordSet
    
    RefreshMain
    
    CheckProviderStatus
    
    
    'This has been replaced with the alerts tab JS 06/28/2012
    'CheckClaimOwner
    
End Sub
Private Sub SaveData()
Dim strProcName As String
'    Dim rs As DAo.Recordset        '' KD 20120910 - Change to ADO
    Dim rs As ADODB.RecordSet
    Dim rs2 As ADODB.RecordSet
    Dim oAdo As clsADO
    
    Dim strClaimStatusGroups As String
    Dim strSupervisorIDGroups As String
    Dim strError As String
    Dim bSaved As Boolean
    Dim bLoaded As Boolean
    
    strProcName = ClassName & ".SaveData"
    
    On Error GoTo Err_SaveData

    If (mbRecordChanged Or Me.Dirty) = False Then
'    If mbRecordChanged = False Or Me.Dirty = False Then
        MsgBox "There are no changes to save."
        Exit Sub
    End If
    
     'Bring up the form to request more information for Incomplete Records
    If Me.ClmStatus = "313" Then
        Call GetIncompleteInfo
    End If
    
    
    ' TK 11/21/2013 per Gautam request: please make a change in CA that doesn't let auditors move a claim to status 354 if no rationale was entered.
    If Me.ClmStatus = "354" And Trim(Nz(myAuditClaim.rsAuditClmHdr.Fields("adj_rationale").Value, "")) = "" Then
        MsgBox "Claim NOT saved. Please type up a valid rationale for this claim."
        Exit Sub
    End If
    
    strError = ""
    
    '
    ' JS 20121107 Changes to check if claim was changed simultaneously by someone else. If so the claim will not be saved
    '
    Dim myCode_ADO As clsADO
    Dim rst As ADODB.RecordSet
    Dim strSQL As String
    Set myCode_ADO = New clsADO
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_DATA_Database")
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strSQL = "select ModifAfterLoadFlag, ModifAfterLoadUserID, ModifAfterLoadUpDt from cms_auditors_code.dbo.udf_AUDITCLM_ClaimModifAfterLoad('" & myAuditClaim.CnlyClaimNum & "', '" & Format(mstrLastUpDt, "yyyy-mm-dd hh:mm:ss") & "')"
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    'If the claim was modified after it was loaded
    If rst("ModifAfterLoadFlag") Then
        MsgBox "Claim was modified simultaneously by " & rst("ModifAfterLoadUserID") & vbNewLine & _
            " on " & rst("ModifAfterLoadUpDt") & "." & vbNewLine & vbNewLine & _
                    "Cannot Overwrite, Changes will be Lost!", vbExclamation, "Claim changed simultaneously"
        MsgBox "Record not saved." & vbCrLf & strError, vbCritical
        Exit Sub
    End If
    
    'Check the claim status
    If CheckClaimStatus(strError) = False Then
        MsgBox "Record not saved." & vbCrLf & strError, vbCritical
        Exit Sub
    Else
        If Not myAuditClaim.rsAuditClmHdr Is Nothing Then
            
            ' TKL 3/2/2011: auditor credit override begin
            If Me.Adj_Auditor & "" = "" Or mCreditAssignment = AutoAssignment Then
                ' auto assignment
                strClaimStatusGroups = "RC,NR,TM,RJ"
        
                
                Set oAdo = New clsADO        '' KD 20120910 - Change to ADO
                With oAdo
                    .ConnectionString = GetConnectString("v_Data_Database")
                    .SQLTextType = sqltext
                    .sqlString = "select * from XREF_ClaimStatus where ClmStatus = '" & Me.ClmStatus & "'"
                    Set rs = .ExecuteRS
                End With
                
                Set oAdo = New clsADO
                With oAdo
                    .ConnectionString = GetConnectString("v_Data_Database")
                    .SQLTextType = sqltext
                    .sqlString = "select  * from admin_user where UserID = '" & mstrUserName & "'"
                    Set rs2 = .ExecuteRS
                End With
                
                '' /  KD 20120910 - Change to ADO
        
                'Set rs = CurrentDb.OpenRecordSet("select * from XREF_ClaimStatus where ClmStatus = '" & Me.ClmStatus & "'")
                
                'JS 2013/11/08
                'adj_Auditor will not be updated when user belongs to the following supervisorIDs
                strSupervisorIDGroups = "DATA CENTER,ADMIN"
                
                If InStr(1, strClaimStatusGroups, UCase(rs("ClmStatusGroup"))) > 0 And Not InStr(1, strSupervisorIDGroups, UCase(rs2("SupervisorID"))) > 0 Then
                    Me.Adj_Auditor = mstrUserName
                    mCreditAssignment = AutoAssignment
                    myAuditClaim.OverrideAuditor = ""
                    myAuditClaim.CreditOverrideReason = ""
                End If
            ElseIf Me.Adj_Auditor & "" <> myAuditClaim.ClaimAuditor And Me.Adj_Auditor & "" <> "" Then
                ' credit override
                myAuditClaim.OverrideAuditor = Me.Adj_Auditor
                myAuditClaim.CreditOverrideReason = mstrCreditOverrideReason
            End If
            ' TKL 3/2/2011: auditor credit override end
            
            '' 7/16/2013 KD: Prepay Therapy concepts need to have an ERror code selected
            '' otherwise we should not save.
            '' Ok, but also, it should ONLY require it for Recovery clm status..
            '' 7/23/2013 KD: Ok, it's NOT just Prepay, there are also post pay ones..
            If IsThisClaimTherapyCongress(Me.CnlyClaimNum) = True Then
'            If Me.IsTherapyConcept = True Then
                '            If Me.Adj_ReviewType = "PRP" Then
                '' what about only specific concepts?
                If Me.ClmStatus = "320.T" Then   ' Prepay Recovery Therapy
                    If Me.ErrorCodePrpty = "" Then
                        LogMessage TypeName(Me) & ".SaveData", "ERROR", "Validation error! Pre-pay claims need an error code set in the Chart Review - IP tab", Me.CnlyClaimNum & " : " & Identity.UserName, True
                        GoTo Exit_SaveData
                    End If
                End If
            End If
            
            bSaved = myAuditClaim.SaveClaim
            If bSaved Then
                ' 20130723 KD: Ok, we haven't saved the Error Code yet. so do that here
                'If Me.IsTherapyConcept = True Then
                If ClaimOfRightReviewTypeDoesNotHaveLineLevelReason(Me.CnlyClaimNum) = True Then
                    'VS 11/19/2015 RVC Project Save Header Level Error Code
                        If Me.ErrorCodePrpty <> "" Then
                                Call SaveTherapyErrorCode
                        ElseIf Not myAuditClaim.rsAuditClmHdrAdditionalInfo Is Nothing Then
                            If myAuditClaim.rsAuditClmHdrAdditionalInfo.EOF = False And myAuditClaim.rsAuditClmHdrAdditionalInfo.BOF = False Then
                                If Me.ErrorCodePrpty = "" And Nz(myAuditClaim.rsAuditClmHdrAdditionalInfo.Fields("ErrorCode"), "") <> "" Then
                                    Call SaveTherapyErrorCode
                                End If
                            End If
                        Else
                            LogMessage strProcName, "ERROR", "Somehow the user didn't set the error code for this claim", , , Me.Adj_ConceptID, Me.CnlyClaimNum
                        End If
                End If
            
                MsgBox "Record saved.", vbOKOnly + vbInformation
                
                bLoaded = myAuditClaim.LoadClaim(Me.CnlyClaimNum, mbAllowChange)
                
                RefreshMain
                mbRecordChanged = False
                
                ' TKL 3/2/2011: auditor credit override
                mstrCreditOverrideReason = ""
                mstrPrevClaimAuditor = myAuditClaim.ClaimAuditor
                mCreditAssignment = NotAssigned
            Else
                MsgBox "Record not saved." & vbCrLf & strError, vbCritical
            End If
        End If
    End If
    
Exit_SaveData:

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close '' KD 20120910 - Change to ADO
    End If
    Set rs = Nothing
    Set oAdo = Nothing
    Exit Sub

Err_SaveData:
   strError = Err.Description
   MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
   Resume Exit_SaveData
    
End Sub

Private Sub GetIncompleteInfo()
      
        Dim myCode_ADO As New clsADO
        Dim rs As ADODB.RecordSet
        Dim rs_ED As ADODB.RecordSet
        Dim strSQL As String

        On Error GoTo Err_handler

        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        myCode_ADO.SQLTextType = StoredProc
        myCode_ADO.sqlString = "usp_Incomplete_MR_Data"
        myCode_ADO.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
        Set rs = myCode_ADO.ExecuteRS

        If rs.EOF = True Then
            ' Were not able to retrieve data for Incomplete Records
            Exit Sub
        End If
        
         Set frmIncompleteMRRequest = New Form_frm_Incomplete_MR_Request
         
          If Me.DataType = "DME" Then
            If Me.Adj_ConceptID = "CM_C1454" Or Me.Adj_ConceptID = "CM_C1637" Or Me.Adj_ConceptID = "CM_C1444" Then
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_DME_CM_C1454"
                 frmIncompleteMRRequest.FormType = DME_1454
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.Height = frmIncompleteMRRequest.frm_MR_Needed_Subform.Height - 500
                 frmIncompleteMRRequest.InsideHeight = frmIncompleteMRRequest.InsideHeight - 500
                 frmIncompleteMRRequest.lblAuditorNotes.top = frmIncompleteMRRequest.lblAuditorNotes.top - 500
                 frmIncompleteMRRequest.Box35.top = frmIncompleteMRRequest.Box35.top - 500
                 frmIncompleteMRRequest.txtNotes.top = frmIncompleteMRRequest.txtNotes.top - 500
                 
            ElseIf Me.Adj_ConceptID = "CM_C0851" Then
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_DME_CM_C0851"
                 frmIncompleteMRRequest.FormType = DME_0851
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.Height = frmIncompleteMRRequest.frm_MR_Needed_Subform.Height - 1000
                 frmIncompleteMRRequest.InsideHeight = frmIncompleteMRRequest.InsideHeight - 1000
                 frmIncompleteMRRequest.lblAuditorNotes.top = frmIncompleteMRRequest.lblAuditorNotes.top - 1000
                 frmIncompleteMRRequest.Box35.top = frmIncompleteMRRequest.Box35.top - 1000
                 frmIncompleteMRRequest.txtNotes.top = frmIncompleteMRRequest.txtNotes.top - 1000
                 
            Else
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_DME"
                 frmIncompleteMRRequest.FormType = DME
                 frmIncompleteMRRequest.frm_MR_Needed_Subform.Height = frmIncompleteMRRequest.frm_MR_Needed_Subform.Height - 950
                 frmIncompleteMRRequest.InsideHeight = frmIncompleteMRRequest.InsideHeight - 950
                 frmIncompleteMRRequest.lblAuditorNotes.top = frmIncompleteMRRequest.lblAuditorNotes.top - 950
                 frmIncompleteMRRequest.Box35.top = frmIncompleteMRRequest.Box35.top - 950
                 frmIncompleteMRRequest.txtNotes.top = frmIncompleteMRRequest.txtNotes.top - 950
            
            End If
            
            
         ElseIf Me.DataType = "OP" Or Me.DataType = "CARR" Then
                frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_Pharmacy"
                frmIncompleteMRRequest.FormType = OP_CARR
                
                frmIncompleteMRRequest.frm_MR_Needed_Subform.Height = frmIncompleteMRRequest.frm_MR_Needed_Subform.Height + 350
                frmIncompleteMRRequest.InsideHeight = frmIncompleteMRRequest.InsideHeight + 350
                frmIncompleteMRRequest.lblAuditorNotes.top = frmIncompleteMRRequest.lblAuditorNotes.top + 350
                frmIncompleteMRRequest.Box35.top = frmIncompleteMRRequest.Box35.top + 350
                frmIncompleteMRRequest.txtNotes.top = frmIncompleteMRRequest.txtNotes.top + 350
                
         ElseIf Me.DataType = "HH" Then
                frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_HH"
                frmIncompleteMRRequest.FormType = HH
                
                frmIncompleteMRRequest.frm_MR_Needed_Subform.Height = frmIncompleteMRRequest.frm_MR_Needed_Subform.Height - 500
                frmIncompleteMRRequest.InsideHeight = frmIncompleteMRRequest.InsideHeight - 500
                frmIncompleteMRRequest.lblAuditorNotes.top = frmIncompleteMRRequest.lblAuditorNotes.top - 500
                frmIncompleteMRRequest.Box35.top = frmIncompleteMRRequest.Box35.top - 500
                frmIncompleteMRRequest.txtNotes.top = frmIncompleteMRRequest.txtNotes.top - 500
                             
         
         ElseIf Me.DataType = "SNF" Then
            frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform_SNF"
            frmIncompleteMRRequest.FormType = SNF
         
         ElseIf Me.DataType = "HSP" Or Me.DataType = "IP" Then
            frmIncompleteMRRequest.frm_MR_Needed_Subform.SourceObject = "frm_MR_Needed_Subform"
            frmIncompleteMRRequest.FormType = STANDARD
         End If
 
         Set frmIncompleteMRRequest.Form.RecordSet = rs
         Set frmIncompleteMRRequest.NoteRecordSource = myAuditClaim.rsNotes
         
         myCode_ADO.sqlString = "udf_EDAdmission"
         myCode_ADO.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
         Set rs_ED = myCode_ADO.ExecuteRS

         strSQL = "Select dbo.udf_EDAdmission('" & Me.CnlyClaimNum & "')"
         rs_ED.Open strSQL, myCode_ADO.CurrentConnection

         If Not (rs_ED.EOF Or rs_ED.BOF) Then
             frmIncompleteMRRequest.txtEd = rs_ED(0).Value
             
             If frmIncompleteMRRequest.txtEd <> "Yes" And (Me.DataType = "IP" Or Me.DataType = "OP" Or Me.DataType = "CARR" Or (Me.DataType = "HH" And Me.Adj_ConceptID <> "CM_C0834")) Then
             frmIncompleteMRRequest.frm_MR_Needed_Subform.Form!ch8.Enabled = False
             frmIncompleteMRRequest.frm_MR_Needed_Subform.Form!ch32.Enabled = False
             End If
             
         End If
             
         
         frmIncompleteMRRequest.SelectedBoxes = 0
         
         If (rs("attributes").Value <> "") Then
         SetSelectedCheckBoxes (rs("attributes").Value)
         frmIncompleteMRRequest.SelectedBoxes = rs("attributes").Value
         End If
         
         If (rs("AuditorNotes").Value <> "") Then
         frmIncompleteMRRequest.Notes = rs("AuditorNotes").Value
         End If
         
         If (rs("Other").Value <> "") Then
         frmIncompleteMRRequest.other = rs("Other").Value
         End If
         
         frmIncompleteMRRequest.txtNotes = rs("AuditorNotes").Value
         frmIncompleteMRRequest.frm_MR_Needed_Subform.Form!txtOther = rs("Other").Value
         ShowFormAndWait frmIncompleteMRRequest

Exit_Sub:

        Set myCode_ADO = Nothing
        Set rs = Nothing
        Set rs_ED = Nothing
        Exit Sub

Err_handler:
        MsgBox "Error populating Incomplete MR Request Form: " & Err.Description
        Resume Exit_Sub
End Sub

Public Function SetSelectedCheckBoxes(allSelected As Long) As Boolean
 
    Dim i As Long

    i = 1

    While (i < allSelected)
           i = i * 2
    Wend

    If i > 1 And i <> allSelected Then
    i = i / 2
    End If

        While (allSelected > 1)
        
            While (allSelected < i)
            i = i / 2
            Wend
            
            frmIncompleteMRRequest.frm_MR_Needed_Subform.Controls.Item(chBoxName & i).Value = -1
            allSelected = allSelected - i
            i = i / 2
        
        Wend

    If (allSelected = 1) Then
     frmIncompleteMRRequest.frm_MR_Needed_Subform.Controls.Item(chBoxName & allSelected).Value = -1
    End If
    
    SetSelectedCheckBoxes = True
    End Function

Private Sub myAuditClaim_AuditClmError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub


Private Function CheckClaimStatus(ByRef strValidationError As String) As Boolean
'    Dim rs As DAo.Recordset        '' KD 20120830 - Change to ADO
    Dim rs As ADODB.RecordSet
    Dim oAdo As clsADO
    Dim bValid As Boolean
    Dim strNewStatus As String
    Dim strOldStatus As String
    Dim MaxThresholdDt As String
    Dim bDateOK As Boolean
    Dim strFilter As String
    
    strNewStatus = Me.ClmStatus
    strOldStatus = myAuditClaim.LoadClaimStatus
    
    bDateOK = False
    bValid = False
    
    'Check to see if this is a legal status change
'    strFilter = " ProcessModule = 'AuditClm' and ProcessType = 'Manual' and CurrStatus = '" & _
'                strOldStatus & "' and NextStatus = '" & strNewStatus & "' AND AccountID = " & _
'                gintAccountID & " and ReviewType = '" & Me.Adj_ReviewType & "'"
'
'                    '' HMMM DLookup.. DAO
'    If Nz(DLookup("CurrStatus", "AUDITCLM_Process_Logics", strFilter), "") = "" Then
'        bValid = False
'        strValidationError = "Error: Cannot move the claim to this status"
'        CheckClaimStatus = False
'        Exit Function
'    End If

    
    'JS 20121106 if the status of the concept is not in (250, 360, 380) and claim status is not "Not Recovery" then it's a no go!
    '               i added the original claim concept to the concept list regardless of its current status
    '               because if the claims becomes no recovery then they need to select it again in the header and it wasn't possible for concepts that were 250 and moved to let's say 300
    'JS 20120508 Excluding claims in status 4xx from this validation as AR people usually work on this claims and they don't need to be hit by this restriction
    'JS 20130709 Adding condition: "Concept was changed" to this validation. so auditor doesnt get bothered by this if they didnt change the conceptid
    If mstrConceptID <> Me.Adj_ConceptID And InStr(1, "380,360,250", Me.Adj_ConceptID.Column(3, Me.Adj_ConceptID.ListIndex)) = 0 Then
       If Me.ClmStatus <> "321" And Not (Me.ClmStatus Like "4??") Then
            MsgBox "You can only select a Non Active Concept if the claim status is No Recovery!", vbExclamation, "Non Active Concept"
            bValid = False
            CheckClaimStatus = bValid
            Exit Function
       End If
    End If



    Set oAdo = New clsADO
    oAdo.ConnectionString = GetConnectString("v_Data_Database")
    
    With oAdo
        .SQLTextType = sqltext
        .sqlString = "select * from cms_auditors_claims.dbo.AUDITCLM_Process_Logics where ProcessModule = 'AuditClm' and ProcessType = 'Manual' and CurrStatus = '" & _
                strOldStatus & "' and NextStatus = '" & strNewStatus & "' AND AccountID = " & _
                gintAccountID & " and ReviewType = '" & Me.Adj_ReviewType & "' and datatype = '" & Me.DataType & "'"
        Set rs = .ExecuteRS
    End With
    
    If rs.EOF And rs.BOF Then
        bValid = False
        strValidationError = "Error: Cannot move the claim to this status"
        CheckClaimStatus = False
        Exit Function
    End If
    rs.Close


        
    'OK We have a legal change validate the specifics
'    Set rs = CurrentDb.OpenRecordSet("select * from XREF_ClaimStatus where ClmStatus = '" & strNewStatus & "'")
    
    '' KD 20120910 - Change to ADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select * from XREF_ClaimStatus where ClmStatus = '" & strNewStatus & "'"
        Set rs = .ExecuteRS
    End With

    Select Case strNewStatus
        Case "305", "307", "309"
            'Placed First Call to Provider for non receipt of medical records
            ' pop up box for next threshold date
            While bDateOK = False
                MaxThresholdDt = InputBox("Please enter next review date: ")
                If IsDate(MaxThresholdDt) Then
                    If CDate(MaxThresholdDt) <= Date Then
                        MsgBox "Error: next review date can not be prior to today's date", vbInformation
                    Else
                        bDateOK = True
                        mdtMaxThreshold = MaxThresholdDt
                    End If
                Else
                    MsgBox "Please enter a valid date", vbInformation + vbCritical
                End If
            Wend
            bValid = True
        Case "501" 'Claims Plus
           bValid = True
        Case "705"
            'Demand Letter Queue
            bValid = Check705(strValidationError)
        Case Else
            'Set rs = CurrentDb.OpenRecordSet("select * from XREF_ClaimStatus where ClmStatus = '" & strNewStatus & "'")
            bValid = True
                
            If rs.BOF = True And rs.EOF = True Then
                strValidationError = "Status is not on master status list"
                bValid = False
            Else
                                                                        'JS 2013/11/08 and not change in status (like when a customer service person just wants to add a note)
                If InStr(1, "RC", UCase(rs("ClmStatusGroup"))) > 0 And Not (rs("ClmStatus") = strOldStatus) Then
                    ' recovery claim check
                    bValid = CheckRecovery(strValidationError)
                End If
                    
                If InStr(1, "NR,RJ,TM", UCase(rs("ClmStatusGroup"))) > 0 Then
                    ' no recovery claim check
                    bValid = CheckNoRecovery(strValidationError)
                End If
                
                If UCase(rs("ValidationInd")) = "Y" Then
                    'Check to see if this status need to be validated
                    strValidationError = strValidationError & "Status requires validation but there is no validation routine for it.  Please alert IT"
                    bValid = False
                End If
            End If
    End Select
    

    'set the function to the output of the check
    CheckClaimStatus = bValid

Block_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close '' KD 20120910 - Change to ADO
    End If
    Set rs = Nothing
    Set oAdo = Nothing
End Function


Function Check705(ByRef strValidationError As String) As Boolean
    
    Check705 = CheckRecovery(strValidationError)
    
End Function

Function Check501(ByRef strValidationError As String) As Boolean
    
    'This is to move this claim to claims plus
    'Check that all required fields are filled in in AUDITCLM_ClaimsPlus before moving
    
    
    
    strValidationError = ""
    
    If (myAuditClaim.rsAuditClmClaimsPlus.BOF And myAuditClaim.rsAuditClmClaimsPlus.EOF) Then
        strValidationError = strValidationError & vbCrLf & "Claims Plus information is required in order to move to Claims Plus."
        Check501 = False
        Exit Function
    Else
        myAuditClaim.rsAuditClmClaimsPlus.MoveFirst
    End If
    
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("GrossAmt")), "") = "" And Not IsNumeric(Trim(myAuditClaim.rsAuditClmClaimsPlus("GrossAmt"))) Then
        strValidationError = strValidationError & vbCrLf & "Gross Amount is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("RootCause")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Root Cause is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("ClaimSrc")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Claim Source is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("ClaimSrcRootTxt")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Claim Source is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
        
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("ClaimSrcTxt")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Claim Source is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
    
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("ClaimCode")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Connolly Reason is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
    
    
    If Nz(Trim(myAuditClaim.rsAuditClmClaimsPlus("ClaimCodeText")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Connolly Reason is required to move to Claims Plus."
        Check501 = False
        Exit Function
    End If
        
    Check501 = True
End Function


Private Function CheckNoRecovery(strValidationError As String) As Boolean
    If Nz(Me.Adj_ProjectedSavings, "") <> "" Then
        strValidationError = strValidationError & vbCrLf & "There should not be a Projected Savings for non-recovery claim."
        CheckNoRecovery = False
        Exit Function
    End If
    
    If Nz(Me.Adj_ReimbAmt, "") <> "" Then
        strValidationError = strValidationError & vbCrLf & "There should not be an Adjusted Reimbursement for non-recovery claim."
        CheckNoRecovery = False
        Exit Function
    End If
        
    If Nz(Me.Adj_DRG, "") <> "" Then
        strValidationError = strValidationError & vbCrLf & "There should not be an Adjusted DRG for non-recovery claim."
        CheckNoRecovery = False
        Exit Function
    End If
    
    CheckNoRecovery = True

End Function

Private Function CheckRecovery(strValidationError As String) As Boolean
    ' Adj_ProjectedSavings
    ' Adj_ReviewType 10/23/2012 KCF - Need to allow Prepay Review Type Claims to bypass Project Savings Check
    If (Nz(Trim(Me.Adj_ProjectedSavings), "") = "" And Me.Adj_ReviewType <> "PRP") Then
        strValidationError = strValidationError & vbCrLf & "Projected Savings can not be blank for recovery claim."
        CheckRecovery = False
        Exit Function
    End If
    
    'Adj_ReimbAmt
    ' Adj_ReviewType 10/23/2012 KCF - Need to allow Prepay Review Type Claims to bypass Reimbursement Amount Check
    If (Nz(Trim(Me.Adj_ReimbAmt), "") = "" And Me.Adj_ReviewType <> "PRP") Then
        strValidationError = strValidationError & vbCrLf & "Adjusted Reimbursement can not be blank for recovery claim."
        CheckRecovery = False
        Exit Function
    End If
        
    'Adj_Drg
    If (Nz(Trim(Me.Adj_DRG), "") = "" And Nz(Trim(Me.DataType), "") = "IP") Then
        strValidationError = strValidationError & vbCrLf & "Adjusted DRG cannot be blank for recovery claim"
        CheckRecovery = False
        Exit Function
    End If
    
    'Adj_Rationale
    If Nz(Trim(myAuditClaim.rsAuditClmHdr("Adj_Rationale")), "") = "" Then
        strValidationError = strValidationError & vbCrLf & "Rationale should not be blank for recovery claim"
        CheckRecovery = False
        Exit Function
    End If
        
    'Make sure there is a chart review date
    'If (Nz(Trim(Me.Adj_ChartReviewDt), "") = "" And Nz(Trim(Me.DataType), "") = "IP") Then
    '    strValidationError = strValidationError & vbCrLf & "No date specified for chart review."
    '    CheckRecovery = False
    '    Exit Function
    'End If
    
    CheckRecovery = True

End Function

Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mReturnDate = ReturnDate
End Sub

Private Function AllowNonMgrToRollback() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet


    strProcName = TypeName(Me) & "AllowNonMgrToRollback"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("V_DATA_DATABASE")
        .SQLTextType = sqltext
        .sqlString = "SELECT LastUpdate, UpdateUser FROM QUEUE_Hdr WHERE CnlyClaimNum = '" & Me.CnlyClaimNum & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            AllowNonMgrToRollback = False
            GoTo Block_Exit
        End If
        
        If LCase(Identity.UserName) = LCase("" & oRs("UpdateUser").Value) And DateDiff("hh", Now(), oRs("LastUpdate").Value) < 24 Then
            AllowNonMgrToRollback = True
        End If
        
        '' Now we need to know if they've already rolled back or made more than 1 change to this in the last 24 hours
        '' I guess I should do something like select count(*) from auditclm_hdr_audit_hist where cnlyclaimnum = ... and DATEDIFF(HH,GetDate(),histcreatedt ) < 24
        '' but for now, I'm not going to..
        
    End With
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    AllowNonMgrToRollback = False
    GoTo Block_Exit
End Function


Private Sub ClmStatus_Change()
    
    mstrConceptCategory = ""
    
    If Nz(Me.Adj_ConceptID, "") <> "" And Nz(Me.Adj_ReviewType, "") <> "" Then
        mstrConceptCategory = GetClaimConceptCategory(Me.Adj_ConceptID, Me.Adj_ReviewType)
    End If
    
    'warning in case recovery underpayment is selected JS 10/17/2012
    If Me.ClmStatus = "322" And mstrConceptCategory = "MN" Then
        MsgBox "You selected an UNDERPAYMENT Recovery status." & vbNewLine & vbNewLine & "Click OK to continue", vbInformation + vbOKOnly, "Underpyament Recovery"
    End If
    
    'if recovery (OP or UP) is selected
    If (Me.ClmStatus = "320" Or Me.ClmStatus = "322" Or Me.ClmStatus = "353" Or Me.ClmStatus = "354") And _
        Me.Adj_ConceptID.Column(2, Me.Adj_ConceptID.ListIndex) = "Medical Necessity" Then
    
        If Nz(Me.Adj_ReimbAmt, "") = "" Then Adj_ReimbAmt = 0
        If Nz(Me.Adj_DRG, "") = "" Then Adj_DRG = "000"
        If Nz(Me.Adj_ProjectedSavings, "") = "" Then Me.Adj_ProjectedSavings = mstrAdj_ProjectedSavings
        
    'If non recovery
    ElseIf (Me.ClmStatus = "321" Or Me.ClmStatus = "355") Then
        
        Adj_ProjectedSavings = Null
        If Nz(Me.Adj_ReimbAmt, "") = 0 Then Adj_ReimbAmt = mstrAdj_ReimbAmt
        If Nz(Me.Adj_DRG, "") = "000" Then Adj_DRG = mstrAdj_DRG
        
    Else 'undo changes if user changes from OP, UP or NR
    
        If Nz(Me.Adj_ReimbAmt, "") = 0 Then Adj_ReimbAmt = mstrAdj_ReimbAmt
        If Nz(Me.Adj_DRG, "") = "000" Then Adj_DRG = mstrAdj_DRG
        If Nz(Me.Adj_ProjectedSavings, "") = "" Then Me.Adj_ProjectedSavings = mstrAdj_ProjectedSavings
        
    End If
    
    Me.ClmStatusCode = Nz(Me.ClmStatus, "")
    
End Sub


Function GetClaimConceptCategory(ConceptID As String, ReviewType As String) As String

    Dim rst As ADODB.RecordSet
    Dim myCode_ADO As clsADO
    Dim strSQL As String
    
    'Dim StrSQl As String
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_DATA_Database")
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strSQL = "SELECT     ConceptCategory = " & _
                " CASE " & _
                " WHEN '" & ReviewType & "' = 'S' THEN 'SEMI' " & _
                " WHEN '" & ReviewType & "' = 'C' AND ConceptGroup = 'Medical Necessity' THEN 'MN' " & _
                " WHEN '" & ReviewType & "' = 'C' AND (ConceptGroup = 'Pharmacy' or LOB = 'CNLY Pharmacy') THEN 'PHARM' " & _
                " WHEN '" & ReviewType & "' = 'C' THEN 'DRG' " & _
                " Else 'AUTO' " & _
                " End from (select distinct conceptid, conceptgroup, lob from cms_auditors_code.dbo.v_concept_hdr_payer) a where conceptid = '" & ConceptID & "'"
    
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    
    GetClaimConceptCategory = ""
    
    If Not (rst Is Nothing) Then
        If Not (rst.EOF Or rst.BOF) Then
            rst.MoveFirst
            GetClaimConceptCategory = Nz(rst("ConceptCategory"), "")
        End If
    End If
    
    
End Function

Private Function SaveTherapyErrorCode() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim strSQL As String
Dim strErrorCodeNew As String

    strProcName = ClassName & ".SaveTherapyErrorCode"
    
    strErrorCodeNew = Me.ErrorCodePrpty
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
            '    strSql = " EXEC cms_auditors_code.dbo.usp_ErrorCodeUpdate '" & Me.CnlyClaimNum & "', '" & strErrorCodeNew & "'"
    strSQL = "usp_ErrorCodeUpdate"
    Debug.Print strSQL
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = strSQL
        .Parameters.Refresh
        .Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
        .Parameters("@pErrorCode") = strErrorCodeNew
        
        Set oRs = .ExecuteRS
    End With
    
    SaveTherapyErrorCode = True
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
