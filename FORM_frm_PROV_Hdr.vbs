Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_PROV_Hdr
' Description:
'   Main Provider maintenance form.
'
' Modification History:
'   2010-02-03 by Barbara Dyroff to add the Provider MR Scanning Invoices tab list option.
'   2012-01-09 by Andrew Lauer to fix the bug when creating a new payer
' =============================================


Public WithEvents myPROV As clsPROV
Attribute myPROV.VB_VarHelpID = -1
Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1
Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents frmProvAddr As Form_frm_PROV_Addr
Attribute frmProvAddr.VB_VarHelpID = -1
Private WithEvents frmGetNewCnlyProvID As Form_frm_PROV_Create_New_ID
Attribute frmGetNewCnlyProvID.VB_VarHelpID = -1

Private frmAUDITTracking As Form_frm_AUDIT_TRACKING_Main
Private frmScanMRInvMain As Form_frm_SCANNING_MR_Invoice_Main

Private mrsPROVHdr As ADODB.RecordSet
Private mrsPROVAddrDeleted As ADODB.RecordSet

Private mstrCnlyProvID As String
Private mbInsert As Boolean
Private mstrUserProfile As String
Private mbRecordChanged As Boolean
Private mbRecordLocked As Boolean
Private mReturnDate As Date

Private miAppPermission As Integer
Private mbAllowView As Boolean
Private mbAllowChange As Boolean
Private mbAllowDelete As Boolean
Private mbAllowAdd As Boolean
Private mbLocked As Boolean

Private mstrRtnCnlyProvID As String
Private mstrRtnProvID As String
Private mstrRtnPayerID As String

Private mbolProviderSaved As Boolean

Const CstrFrmAppID As String = "ProvHdr"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let cnlyProvID(data As String)
    mstrCnlyProvID = data
End Property

Property Get cnlyProvID() As String
    cnlyProvID = mstrCnlyProvID
End Property

Property Let Insert(data As Boolean)
    mbInsert = data
End Property

Property Get Insert() As Boolean
    Insert = mbInsert
End Property

Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property

Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property

Property Let RecordLocked(data As Boolean)
    mbRecordLocked = data
    If mbRecordLocked Then
        cmdSave.Enabled = False
        cmdNewProviderAddress.Enabled = False
        cmdNewProviderNote.Enabled = False
        Me.AllowEdits = False
    Else
        cmdSave.Enabled = True
        cmdNewProviderAddress.Enabled = True
        cmdNewProviderNote.Enabled = True
        Me.AllowEdits = True
    End If
End Property

Property Get RecordLocked() As Boolean
    RecordLocked = mbRecordLocked
End Property


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdNewProviderNote_Click()
    On Error GoTo Err_cmdNewProviderNote
    
     Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
    
     frmGeneralNotes.frmAppID = Me.frmAppID
     Set frmGeneralNotes.NoteRecordSource = myPROV.PROVNotes
     frmGeneralNotes.RefreshData
     ShowFormAndWait frmGeneralNotes
     lstTabs_Click
     Set frmGeneralNotes = Nothing

Exit_cmdNewProviderNote:
    Exit Sub

Err_cmdNewProviderNote:
    MsgBox Err.Description
    Resume Exit_cmdNewProviderNote

End Sub

Private Sub cmdNewProviderAddress_Click()
    If myPROV.PROVAddr.recordCount = 0 Then
        ' this is needed to reset the recordset bookmark
        Set myPROV.PROVAddr = myPROV.PROVAddr.Clone
    End If
    
    Set frmProvAddr = New Form_frm_PROV_Addr
    Set frmProvAddr.AddrRecordSource = myPROV.PROVAddr
    'Set frmProvAddr.PortalAddrRecordSource = myPROV.PROVAddrPortal
    frmProvAddr.ProvID = mstrCnlyProvID
    frmProvAddr.AddNewAddress
    frmProvAddr.DisableMousewheel = True
    ShowFormAndWait frmProvAddr
    Set frmProvAddr = Nothing
    
    If myPROV.PROVAddr.recordCount = 0 Then
        ' this is needed to reset the recordset bookmark
        Set myPROV.PROVAddr = myPROV.PROVAddr.Clone
    End If
    
    lstTabs_Click
    
Exit_NewProviderAddress:
    Exit Sub

Err_NewProviderAddress:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "from_Pro_Hdr : NewProviderAddress"
    Resume Exit_NewProviderAddress
    
End Sub

Private Sub cmdNewProvider_Click()
    On Error GoTo Err_NewProvider
    
    If myPROV.ProviderExists Then
        If Me.Dirty Then
            If MsgBox("Record has been changed.  Do you want to save the changes?", vbYesNo) = vbYes Then
                SaveData
            End If
        End If
        
        If myPROV.LockedForEdit Then
            If myPROV.UnLockProv = False Then
                MsgBox "Error unlocking the provider.  Please notify IT"
                Exit Sub
            End If
        End If
    End If
        
    If MsgBox("Would you like to create a new provider?", vbYesNo + vbQuestion) = vbYes Then
        mbInsert = True
    
        Set frmGetNewCnlyProvID = New Form_frm_PROV_Create_New_ID
        ShowFormAndWait frmGetNewCnlyProvID
        Set frmGetNewCnlyProvID = Nothing
        
        If mstrRtnCnlyProvID = "" Then Exit Sub
        
        Me.cnlyProvID = mstrRtnCnlyProvID
        
        LoadProvider
        If myPROV.ProviderExists Then
            MsgBox "Provider " & Me.cnlyProvID & " already exists in dataabase"
            mbInsert = False
        Else
            'DoCmd.GoToRecord , , acNewRec
            mrsPROVHdr.AddNew
            Me.txtCnlyProvID = Me.cnlyProvID
            Me.ProvNum = mstrRtnProvID
            Me.PayerNum = mstrRtnPayerID
            Me.AccountID = gintAccountID
            ' thieu added new defaults
            Me.Status = "01"
            Me.StatusEffDt = Date
            Me.StatusTermDt = "12/31/9999"
            Me.ProvRisk = 0
            Me.MaxChartReq = 0
        End If
    End If

Exit_NewProvider:
    Exit Sub

Err_NewProvider:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "from_Pro_Hdr : NewProvider"
    Resume Exit_NewProvider
    
End Sub

Private Sub cmdOpen_Click()
    Dim strProvID As String
  
    If myPROV.ProviderExists Then
        If Me.Dirty Or Me.NewRecord Or mbRecordChanged Then
            SaveData
        End If
        
        If myPROV.LockedForEdit Then
            If myPROV.UnLockProv = False Then
                MsgBox "Error unlocking the provider."
                Exit Sub
            End If
        End If
    End If
    
    strProvID = InputBox("Enter Connolly Provider ID.")
    
    If StrPtr(strProvID) <> 0 Then 'If StrPtr function returns 0, then the user pressed cancel
        If strProvID <> "" Then
            Me.Insert = False
            Me.cnlyProvID = Trim(strProvID)
            LoadProvider
        Else
            MsgBox "You entered an invalid Connolly Provider ID."
        End If
    End If
End Sub

Public Sub LoadProvider()
    Dim bResult As Boolean

    If Me.cnlyProvID = "" Then Exit Sub

    bResult = myPROV.LoadProv(Me.cnlyProvID, mbAllowChange)
    
    cmdSave.Enabled = False
    cmdNewProviderAddress.Enabled = False
    cmdNewProviderNote.Enabled = False
    
    If myPROV.ProviderExists Then
        mbInsert = False
        
        If mbAllowChange Then
            If myPROV.LockedForEdit = False Then
                MsgBox "Record is being locked by " & myPROV.LockedUser & " at " & myPROV.LockedDate
            Else
                cmdSave.Enabled = True
                cmdNewProviderAddress.Enabled = True
                cmdNewProviderNote.Enabled = True
            End If
        End If
        
    ElseIf mbInsert Then
        
        If mbAllowAdd Then
            cmdSave.Enabled = True
            cmdNewProviderAddress.Enabled = True
            cmdNewProviderNote.Enabled = True
        End If
    
    Else
        MsgBox "Provider '" & cnlyProvID & "' does not exist"
    End If
    
    RefreshMain
    
End Sub

Private Sub Command65_Click()
'Alex C 3/5/2012 - added for launching Customer Service for this provider
    LaunchNewCustProviderEvent (cnlyProvID)
End Sub

Private Sub Form_Close()
    If myPROV.ProviderExists And myPROV.LockedForEdit Then
        myPROV.UnLockProv
    End If
    RemoveObjectInstance Me
End Sub

Private Sub RefreshMain()
    'Refresh the main form
    
    On Error GoTo ErrHandler
    
    ' set default display
    Me.Caption = "Provider: " & Me.cnlyProvID
    Me.lblAppTitle.Caption = "Provider:"
    
    Set mrsPROVHdr = myPROV.PROVHdr
    
    If myPROV.ProviderExists Then
        Me.lblAppTitle.Caption = "Provider: " & Me.cnlyProvID
    ElseIf mbInsert Then
        Me.lblAppTitle.Caption = "Provider: " & " - NEW PROVIDER"
    Else
        Set Me.RecordSet = Nothing
        screen.PreviousControl.SetFocus
        Me.Detail.visible = False
        Exit Sub
    End If
    
    
    If myPROV.LockedForEdit = False And mbInsert = False And mbAllowChange Then
        Me.lblAppTitle.Caption = Me.lblAppTitle.Caption & " - Locked by " & myPROV.LockedUser
        Me.RecordLocked = True
    End If
    
    
    Set Me.RecordSet = mrsPROVHdr
    
    'thieu
    If myPROV.ProviderExists Then
        If Me.PayerNum & "" = "" Then Me.PayerNum = "00000"  ' set default payer
    End If
    
    Me.Detail.visible = True

    lstTabs_Click
    
    'contractid. TK 2015-04-13 added for new contract
    ProviderContractID
                  
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : RefreshMain"
End Sub

Private Sub ProviderContractID()
Dim strSQL As String
Dim strCnlyProvID As String
Dim strContractID As String

    strCnlyProvID = myPROV.cnlyProvID
    strSQL = "select ContractID from CMS_AUDITORS_CODE.dbo.ProviderContractID WHERE CnlyProvID = '" & strCnlyProvID & "'"

Dim oAdo As clsADO
Dim rs As ADODB.RecordSet

Set oAdo = New clsADO
With oAdo
    .ConnectionString = GetConnectString("v_Data_Database")
    .SQLTextType = sqltext
    .sqlString = strSQL
    Set rs = .ExecuteRS
End With
    
    strContractID = rs.Fields(0)
    txtContractID.Value = strContractID
    
End Sub

Private Sub Form_Current()
    'Check if there's a pending address change for this provider
    txtAddrChange.visible = myPROV.mbProvAddrChanged
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub

Private Sub Form_Load()
    Me.Caption = "Provider Maintenance"
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
    Set myPROV = New clsPROV
    mstrUserProfile = GetUserProfile()
    CheckPermission


    If mbAllowAdd = False Then
        Me.cmdNewProvider.Enabled = False
    End If

    If mbAllowChange = False Then
        Me.cmdSave.Enabled = False
        Me.cmdNewProviderAddress.Enabled = False
        Me.cmdNewProviderNote.Enabled = False
    End If
    
    lstTabs.RowSource = GetListBoxSQL(Me.Name, mstrUserProfile)
    lstTabs.Requery
    RefreshComboBox "SELECT ProvStatus, Description  FROM PROV_XREF_Status_Code ", Me.Status, "", ""
    RefreshComboBox "SELECT PayerNum, PayerName FROM XREF_Payer WHERE AccountID =  " & gintAccountID, Me.PayerNum, "", ""

    lblAppTitle.Caption = gstrAcctDesc & " - Provider: "
    Me.Caption = lblAppTitle.Caption
    
    Me.Detail.visible = False
    cmdNewProviderNote.Enabled = False
    cmdNewProviderAddress.Enabled = False
    cmdSave.Enabled = False
    If Me.cnlyProvID <> "" Then LoadProvider
    
End Sub



Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If myPROV.ProviderExists = False Then
        GoTo exitHere
    End If
    
    If Me.RecordSource = "" Then
        GoTo exitHere
    End If
    
    If mbAllowChange = False Or (Me.Dirty = False And mbRecordChanged = False) Then
        GoTo exitHere
    End If
    
    SaveData

exitHere:
    Exit Sub
    
End Sub

Private Sub lstTabs_Click()
    Dim strSQL As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    
    Dim rs As DAO.RecordSet
    
    If Not mrsPROVHdr Is Nothing And lstTabs.ListIndex >= 0 Then
        Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstTabs.Column(1), Me.Name), dbOpenSnapshot, dbSeeChanges)
        If Not (rs.BOF And rs.EOF) Then
            Select Case rs("FormName")
                Case "frm_GENERAL_Notes_Display"
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Set Me.subFrmMain.Form.NoteRecordSource = myPROV.PROVNotes
                    Me.subFrmMain.Form.RefreshData
                
                Case "frm_PROV_Address_Grid_View"
                    Me.subFrmMain.SourceObject = rs("FormName")
                    If myPROV.PROVAddr.recordCount = 0 Then
                        ' this is to reset record bookmark
                        Set myPROV.PROVAddr = myPROV.PROVAddr.Clone
                    End If
                    
                    Set Me.subFrmMain.Form.AddrRecordSource = myPROV.PROVAddr
                    Set Me.subFrmMain.Form.PortalAddrRecordSource = myPROV.PROVAddrPortal
                    Set Me.subFrmMain.Form.DeletedAddrRecord = myPROV.PROVAddrDeleted
                    
                    Me.subFrmMain.Form.RecordLocked = Me.RecordLocked
                    Me.subFrmMain.Form.RefreshData
                    Me.subFrmMain.Form.Requery
                
                Case "frm_QUEUE_Exception_Info_Grid_View"
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Me.subFrmMain.Form.FormFilter = "CnlyProvID = '" & mstrCnlyProvID & "'"
                    Me.subFrmMain.Form.RefreshData
                
                Case "frm_AUDIT_TRACKING_Main"
                
                    If Me.subFrmMain.SourceObject <> "frm_PROV_Address_Grid_View" And lstTabs.Value = "Provider Address Change History" Then
                        MsgBox "You need to select an address first!", vbExclamation, "Address History needs an address selected"
                        Exit Sub
                    End If
                    
                    Set frmAUDITTracking = New Form_frm_AUDIT_TRACKING_Main
                    
                    Select Case rs("RowSource")
                        Case "PROV_HDR"
                            frmAUDITTracking.AuditTableName = "PROV_Hdr_Audit_Hist"
                            frmAUDITTracking.AppTitle = "Audit History for Provider : " & cnlyProvID
                            frmAUDITTracking.AuditKey = "CnlyProvID = '" & cnlyProvID & "'"
                        Case "PROV_ADDRESS"
                            frmAUDITTracking.AuditTableName = "PROV_Address_Audit_Hist"
                            frmAUDITTracking.AppTitle = "Audit History of provider addresses for Provider : " & cnlyProvID
                            frmAUDITTracking.AuditKey = "CnlyProvID = '" & cnlyProvID & "' and AddrType = '" & Me.subFrmMain.Form.AddrType & "'"
                    End Select
                    
                    frmAUDITTracking.RefreshData
                    ColObjectInstances.Add frmAUDITTracking, frmAUDITTracking.hwnd & ""
                    frmAUDITTracking.visible = True
                
                'Display Scanning Medical Record Invoices for the provider.
                Case "frm_SCANNING_MR_Invoice_Main"
                    Set frmScanMRInvMain = New Form_frm_SCANNING_MR_Invoice_Main
                    
                    frmScanMRInvMain.PropTableName = rs("RowSource")
                    frmScanMRInvMain.AppTitle = "Medical Record Scanning Invoices"
            
                    frmScanMRInvMain.PropKey = "CnlyProvID = '" & cnlyProvID & "'"
                    frmScanMRInvMain.cnlyProvID = cnlyProvID
                    
                    frmScanMRInvMain.RefreshData
                    ColObjectInstances.Add frmScanMRInvMain, frmScanMRInvMain.hwnd & ""
                    
                    frmScanMRInvMain.visible = True

                Case "frm_PROV_References_Grid_View"
                    Me.subFrmMain.SourceObject = rs("FormName")
                    Me.subFrmMain.Form.FieldReference = "CnlyProvID"
                    Me.subFrmMain.Form.FieldValue = Me.cnlyProvID
                    strSQL = GetNavigateTabSQL(lstTabs.Column(1), Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
                  
                    If strSQL <> "" Then
                        Me.subFrmMain.Form.CnlyRowSource = strSQL
                    End If
                
                   Me.subFrmMain.Form.RefreshData
                
                'MG 9/10/2013 faxing multiple documents per provider
                Case "frm_PROV_Fax_Documents_Grid_View"
                
                    'MsgBox "frm_PROV_Fax_Documents_Grid_View Test !!!!"
                    Me.subFrmMain.SourceObject = rs("FormName")
                    
                Case Else
                    Me.subFrmMain.SourceObject = rs("FormName")
                    
                    strSQL = GetNavigateTabSQL(lstTabs.Column(1), Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
                    If strSQL <> "" Then
                        Me.subFrmMain.Form.CnlyRowSource = strSQL
                    End If
                    Me.subFrmMain.Form.RefreshData
            End Select
        Else
            MsgBox "Application form has not been defined"
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    SaveData
End Sub

Private Function SelectListBoxItemFromText(strTextToFind As String) As Integer
On Error GoTo Block_Err
Dim iIndex As Integer

    For iIndex = 0 To Me.lstTabs.ListCount - 1

        
        If UCase(Me.lstTabs.ItemData(iIndex)) = UCase(strTextToFind) Then
            SelectListBoxItemFromText = iIndex
            lstTabs.Selected(iIndex) = True
            GoTo Block_Exit
        End If
    Next
    ' if we get here then we didn't find it
    iIndex = -1

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, "SelectListBoxItemFromText", "ERROR!"
    GoTo Block_Exit
End Function





Private Sub SaveData()
    Dim bSaved As Boolean
    Dim strError As String
    
    On Error GoTo Err_SaveData
    
    mbolProviderSaved = False
    
    SelectListBoxItemFromText ("Provider Notes")
    Call lstTabs_Click
    
    If mbRecordChanged = False And Me.Dirty = False Then
        MsgBox "There are no changes to save."
        Exit Sub
    End If
    
    If Nz(Me.Status, "") = "" Then
        MsgBox "'Status' field cannot be blank.", vbOKOnly + vbInformation
        Me.Status.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.StatusEffDt, "") = "" Then
        MsgBox "'Status Eff Dt' field cannot be blank.", vbOKOnly + vbInformation
        Me.StatusEffDt.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.ProvName, "") = "" Then
        MsgBox "'Prov Name' field cannot be blank.", vbOKOnly + vbInformation
        Me.ProvName.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.ProvNum, "") = "" Then
        MsgBox "'Prov Num' field cannot be blank.", vbOKOnly + vbInformation
       Me.ProvNum.SetFocus
       Exit Sub
       
    ElseIf Nz(Me.PayerNum, "") = "" Then
        MsgBox "'Payer Num' field cannot be blank.", vbOKOnly + vbInformation
        Me.PayerNum.SetFocus
        Exit Sub
    ElseIf Nz(Me.ProvRisk, "") = "" Then
        MsgBox "'ProvRisk' field cannot be blank.", vbOKOnly + vbInformation
        Me.ProvRisk.SetFocus
        Exit Sub
    ElseIf Nz(Me.MaxChartReq, "") = "" Then
        MsgBox "'MaxChartReq' field cannot be blank.", vbOKOnly + vbInformation
        Me.MaxChartReq.SetFocus
        Exit Sub
'    ElseIf Nz(Me.NPI, "") <> "" And (Len(Trim(Me.NPI)) <> 10 Or Not IsNumeric(Me.NPI)) Then
'        MsgBox "If NPI is entered it must be 10 digits and only numeric.", vbOKOnly + vbInformation
'        Me.MaxChartReq.SetFocus
'        Exit Sub
'    ElseIf Nz(Me.TIN, "") <> "" And (Len(Trim(Me.TIN)) <> 9 Or Not IsNumeric(Me.TIN)) Then
'        MsgBox "If TIN is entered it must be 9 digits and only numeric.", vbOKOnly + vbInformation
'        Me.MaxChartReq.SetFocus
'        Exit Sub
    End If
    
    If mbInsert Then
        'Me.CnlyProvID = Me.PayerNum & Me.ProvNum  thieu
        mbInsert = False
    End If
    
    If MsgBox("Record has changed.  Would you like to save changes to Provider - " & Me.cnlyProvID & "?", vbYesNo + vbQuestion) = vbYes Then
        bSaved = myPROV.SaveProv
                
        If bSaved Then
            mbolProviderSaved = True
            MsgBox "Record saved.", vbOKOnly + vbInformation
            RefreshMain
        Else
            MsgBox "Record not saved.", vbOKOnly + vbCritical
        End If
    End If

    mbRecordChanged = False
    
Exit_SaveData:
    Exit Sub

Err_SaveData:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume Exit_SaveData
    
End Sub

Private Sub StatusEffDt_Exit(Cancel As Integer)
    If Not (IsDate(StatusEffDt.Value) Or IsNull(StatusEffDt.Value)) Then
        MsgBox "Please enter a valid 'Status Eff Dt'."
        StatusEffDt.SetFocus
    End If

End Sub

Private Sub StatusTermDt_Exit(Cancel As Integer)
    If Not (IsDate(StatusTermDt.Value) Or IsNull(StatusTermDt.Value)) Then
        MsgBox "Please enter a valid 'Status Term Dt'."
        StatusTermDt.SetFocus
    End If

End Sub

Private Sub cmdStatusEffDt_Click()
    On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.StatusEffDt, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.StatusEffDt = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click

End Sub

Private Sub cmdStatusTermDt_Click()
    On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.StatusTermDt, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    Me.StatusTermDt = mReturnDate

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click

End Sub

Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mReturnDate = ReturnDate
End Sub

Private Sub myPROV_PROVError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Error: " & ErrNum & " (" & ErrMsg & ")" & vbCr & vbCr & "Source: " & ErrSource, vbOKOnly + vbCritical, "frm_AuditClm_Main : ADO Error"
End Sub


Private Sub frmProvAddr_RecordChanged()
    mbRecordChanged = True
End Sub

Private Sub frmGeneralNotes_NoteAdded()
    mbRecordChanged = True
End Sub


Private Sub CheckPermission()
    If miAppPermission = gcLocked Then mbLocked = True Else mbLocked = False
    
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowDelete = (miAppPermission And gcAllowDelete)
    mbAllowChange = (miAppPermission And gcAllowChange) Or mbAllowAdd Or mbAllowDelete
    mbAllowView = (miAppPermission And gcAllowView) Or mbAllowChange
       
End Sub


Private Sub frmGetNewCnlyProvID_ReturnIDs(cnlyProvID As String, ProvID As String, PayerID As String)
    mstrRtnProvID = ProvID
    mstrRtnPayerID = PayerID
    mstrRtnCnlyProvID = ProvID  ' thieu
End Sub
