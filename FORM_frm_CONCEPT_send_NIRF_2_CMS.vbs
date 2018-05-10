Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 12/27/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 12/27/2012 - fixed bug for when this form is loaded as a subform in concept
'''     mgmt and a concept doesn't have any records.. Synch'd this better..
'''  - 12/03/2012 - Created
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################


Private csConceptId As String



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get FormConceptID() As String
    FormConceptID = IdValue
End Property
Public Property Let FormConceptID(sConceptId As String)
    IdValue = sConceptId
End Property

' frmAppID
Public Property Get frmAppID() As String
    frmAppID = 1
End Property

Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    
    Call Me.RefreshData

Block_Exit:
    Exit Property
End Property


Public Sub RefreshData()
Dim strProcName As String
On Error GoTo Block_Err
Dim oRs As ADODB.RecordSet
'Dim oFrmGeneric As Form_frm_GENERAL_Datasheet_ADO


    strProcName = ClassName & ".RefreshData"
    
    ' Refresh only the page selected..
'    Select Case Me.tabConceptStatusRpts
'    Case 0

'        Set oFrmGeneric = Me.sfrm_Generic.Form
        
        Set oRs = GetGenericReport()

        If Not oRs Is Nothing Then
            
            If oRs.recordCount > 0 Then
                Set Me.RecordSet = oRs
            Else
'                Set Me.Recordset = Nothing
                Set Me.RecordSet = oRs
            End If
        Else
            Set Me.RecordSet = Nothing
        End If
        
'    Case 1
'
'        Set oFrmGeneric = Me.sfrm_StatusOutOfLine.Form
'
'        Set oRs = GetOutOfLineReport()
'
'        If Not oRs Is Nothing Then
'            oFrmGeneric.InitDataADO oRs, "v_Data_Database"
'            If oRs.RecordCount > 0 Then
'                Set oFrmGeneric.Recordset = oRs
'            End If
'        End If
'
'
'    Case 2
'        Stop
'    Case Else
'        Stop
'    End Select
    
    

Block_Exit:
'    Set oFrmGeneric = Nothing
    Set oRs = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub





Private Sub cmbAuditor_AfterUpdate()
    Call RefreshData
End Sub




Private Sub cmbConcept_AfterUpdate()
    Call RefreshData
End Sub


Private Sub cmbPayer_AfterUpdate()
    Call RefreshData
End Sub




Private Sub cmbStatus_AfterUpdate()
    Call RefreshData
End Sub



Private Sub cmdClear_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdClear_Click"
    
    
    Me.cmbConcept = ""
'    Me.cmbStatus = ""
    Me.cmbAuditor = ""
    Me.cmbPayer = ""
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdRefresh_Click"
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String
Dim sUsername As String
Dim sUserProfile As String


    strProcName = ClassName & ".Form_Load"
    Me.Form.RecordSource = ""
    Me.ckShowSent = 0
    
    '' Only Ken should be able to do this:
    sUsername = Identity.UserName()
    sUserProfile = GetUserProfile()
    
'    Select Case UCase(sUsername)
'    Case "KENNETH.TURTURRO" ' , "KEVIN.DEARING"
'        Me.cmdCreateEmail.Enabled = True
'    Case Else
'
'        Me.cmdCreateEmail.Enabled = False
'    End Select
    Select Case UCase(sUserProfile)
    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
        Me.cmdCreateEmail.Enabled = True
    Case Else
        Me.cmdCreateEmail.Enabled = False
    End Select
    
    'Me.sfrm_Generic.Form.AllowFilters = True
    Me.AllowFilters = True
    
'    Me.sfrm_StatusOutOfLine.Form.AllowFilters = True
    
    
    ' refresh all of the combo boxes..
    Call RefreshComboBoxes
    
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit


End Sub


Private Sub RefreshComboBoxes()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshComboBoxes"
    
    ''-- Concept:
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct ConceptId, ConceptDesc from CONCEPT_Hdr"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbConcept.ColumnCount = 2
            Me.cmbConcept.ColumnWidths = "1000;2880;"
            Set Me.cmbConcept.RecordSet = oRs
        End If
    
    
'        .SqlString = "SELECT ConceptStatus, StatusDescription FROM CONCEPT_XREF_Status"
'        Set oRs = .ExecuteRS
'        If .GotData Then
''            Me.cmbStatus.ColumnCount = 2
''            Me.cmbStatus.ColumnWidths = "1000;2880;"
'            Set Me.cmbStatus.Recordset = oRs
'        End If
    
    
        .sqlString = "SELECT DISTINCT Auditor FROM CONCEPT_Hdr WHERE Auditor IS NOT NULL ORDER BY Auditor"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbAuditor.ColumnCount = 1
            Me.cmbAuditor.ColumnWidths = "2880;"
            Set Me.cmbAuditor.RecordSet = oRs
        End If
    
    
    
        .sqlString = "SELECT PayerNameId, PayerName FROM XREF_Payernames WHERE PayerNameId > 999 ORDER BY PayerName"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbPayer.ColumnCount = 2
            Me.cmbPayer.ColumnWidths = "1000;2880;"
            Set Me.cmbPayer.RecordSet = oRs
        End If
    
    
    
    
    
    End With
    
    
    
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Function GetGenericReport() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetGenericReport"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Send_Nirf_to_CMS_Rpt"
        .Parameters.Refresh
        
        If IsSubForm(Me) Then
'            Me.cmbConcept = Nz(Me.Parent.Form.txtConceptID, "")
        End If
        
        If Nz(Me.cmbConcept, "") <> "" Then
            .Parameters("@pConceptId") = Me.cmbConcept
        End If
        
        If Nz(Me.cmbPayer, "") <> "" Then
            .Parameters("@pPayerNameID") = Me.cmbPayer
        End If
        
        If Nz(Me.ckShowSent, 0) <> 0 Then
            .Parameters("@pShowSent") = 1
        End If
        
'        If Nz(Me.cmbStatus, "") <> "" Then
'            .Parameters("@pConceptStatus") = Me.cmbStatus
'        End If
        
        If Nz(Me.cmbAuditor, "") <> "" Then
            .Parameters("@pAuditor") = Me.cmbAuditor
        End If
     
        Set oRs = .ExecuteRS
        If .GotData = False Then
'            GoTo Block_Exit
        End If

    End With
    Debug.Print oRs.recordCount
    
    Set GetGenericReport = oRs
Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function






'Private Sub Form_Resize()
'Dim lWidth As Long
'Dim lHeight As Long
'
'    '' Resize stuff..
'    ' just in the details section
'    lWidth = Me.InsideWidth - 200
'    lHeight = Me.InsideHeight - 700
'
'
'    Me.sfrm_Generic.width = lWidth - 100
'    Me.sfrm_Generic.Height = lHeight - 300
''    Me.sfrm_StatusOutOfLine.width = Me.tabConceptStatusRpts.Pages(1).width - 100
''    Me.sfrm_StatusOutOfLine.Height = Me.tabConceptStatusRpts.Pages(1).Height - 300
'
'
'End Sub

Private Sub tabConceptStatusRpts_Change()
    Call RefreshData
End Sub



Private Sub cmdCreateEmail_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sConceptId As String
Dim lPayerNameId As Long
Dim oConcept As clsConcept
Dim oRs As ADODB.RecordSet
Dim sUsername As String
Dim sUserProfile As String

    strProcName = ClassName & ".cmdCreateEmail_Click"
    sUserProfile = GetUserProfile()
    
    '' Only Ken should be able to do this:
'    sUsername = GetUserName()
'    Select Case UCase(sUsername)
'    Case "KENNETH.TURTURRO" ' , "KEVIN.DEARING"
'        Stop
'    Case Else
'        MsgBox "You do not have adequate permissions to do this!"
'        GoTo Block_Exit
'    End Select
    
    Select Case UCase(sUserProfile)
    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
                
    Case Else
        MsgBox "You do not have adequate permissions to do this!"
        GoTo Block_Exit
    End Select
    
    
        ''' So, we need the concept id
        ''  payernameid
        '' and that should be about it..
    
    Set oRs = Me.RecordSet
    
    
    sConceptId = Nz(oRs("ConceptID").Value, "")
    lPayerNameId = Nz(oRs("PayerNameId").Value, 1000)
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "There was a problem loading the concept object!", , True, sConceptId
        GoTo Block_Exit
    End If
    
    
    Call PrepConceptSubmitEmail(oConcept, lPayerNameId)
    
    
    Call mod_Concept_Specific.MarkConceptAsSentToCms(sConceptId, lPayerNameId)
    
    ' So, now let's just assume that the email is sent and we'll mark the database as sent..
    Call Me.RefreshData

Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub




Private Sub SortForm(strFieldToSortOn As String)
On Error GoTo Block_Err
Dim strProcName As String
Static bAscending As Boolean
Dim sFilter As String
Dim oAdoRs As ADODB.RecordSet
Dim oDaoRs As DAO.RecordSet


    strProcName = ClassName & ".SortForm"
        ' flip it
    bAscending = Not bAscending
    
    sFilter = strFieldToSortOn & IIf(bAscending, " ASC", " DESC")

    
    If TypeOf Me.RecordSet Is ADODB.RecordSet Then
        Set oAdoRs = Me.RecordSet
        oAdoRs.Sort = sFilter
        Set Me.RecordSet = Nothing
        Set Me.RecordSet = oAdoRs
    Else
        Set oDaoRs = Me.RecordSet
        oDaoRs.Sort = sFilter
        Set Me.RecordSet = Nothing
        Set Me.RecordSet = oDaoRs
    End If
    
    


Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub






Private Sub lblAuditor_Click()
    SortForm ("Auditor")
End Sub

Private Sub lblConceptID_Click()
    SortForm ("ConceptId")
End Sub

Private Sub lblConceptLevel_Click()
    SortForm ("ConceptLevel")
End Sub

Private Sub lblConceptStatus_Click()
    SortForm ("CStatus")
End Sub

Private Sub lblDataType_Click()
    SortForm ("DataType")
End Sub

Private Sub lblDateFinalized_Click()
    SortForm ("DateFinalized")
End Sub

Private Sub lblDateSentToPayer_Click()
    SortForm ("DateSentToPayer")
End Sub

Private Sub lblDtSubmittedonNIRF_Click()
    SortForm ("DtSubmittedonNIRF")
End Sub

Private Sub lblLastUpDt_Click()
    SortForm ("LastUpDt")
End Sub

Private Sub lblLastUpUser_Click()
    SortForm ("LastUpUser")
End Sub

Private Sub lblLOB_Click()
    SortForm ("LOB")
End Sub

Private Sub lblNirfSentToCMSDt_Click()
    SortForm ("NirfSentToCMSDt")
End Sub

Private Sub lblNirfSentToCmsUser_Click()
    SortForm ("NirfSentToCmsUser")
End Sub

Private Sub lblPackageCreatedDt_Click()
    SortForm ("PackageCreatedDt")
End Sub

Private Sub lblPackageCreateUser_Click()
    SortForm ("PackageCreateUser")
End Sub

Private Sub lblPayerName_Click()
    SortForm ("PayerName")
End Sub

Private Sub lblPayerNameID_Click()
    SortForm ("PayerNameId")
End Sub

Private Sub lblQAUser_Click()
    SortForm ("QAUser")
End Sub

Private Sub lblReviewType_Click()
    SortForm ("ReviewType")
End Sub

Private Sub lblStatusDescription_Click()
    SortForm ("StatusDescr")
End Sub

Private Sub lblUserWhoSentToPayer_Click()
    SortForm ("UserWhoSentToPayer")
End Sub
