Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 10/08/2012
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
'''  - 10/08/2012 - Created
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
Dim oFrmGeneric As Form_frm_GENERAL_Datasheet_ADO


    strProcName = ClassName & ".RefreshData"
    
    ' Refresh only the page selected..
    Select Case Me.tabConceptStatusRpts
    Case 0

        Set oFrmGeneric = Me.sfrm_Generic.Form
        
        Set oRs = GetGenericReport()

        If Not oRs Is Nothing Then
            oFrmGeneric.InitDataADO oRs, "v_Data_Database"
            If oRs.recordCount > 0 Then
                Set oFrmGeneric.RecordSet = oRs
            End If
        End If
        
    Case 1

        Set oFrmGeneric = Me.sfrm_StatusOutOfLine.Form
        
        Set oRs = GetOutOfLineReport()

        If Not oRs Is Nothing Then
            oFrmGeneric.InitDataADO oRs, "v_Data_Database"
            If oRs.recordCount > 0 Then
                Set oFrmGeneric.RecordSet = oRs
            End If
        End If


    Case 2
        Stop
    Case Else
        Stop
    End Select
    
    

Block_Exit:
    Set oFrmGeneric = Nothing
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
    Me.cmbStatus = ""
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

    strProcName = ClassName & ".Form_Load"
    
    
    Me.sfrm_Generic.Form.AllowFilters = True
    Me.sfrm_StatusOutOfLine.Form.AllowFilters = True
    
    
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
    
    
        .sqlString = "SELECT ConceptStatus, StatusDescription FROM CONCEPT_XREF_Status"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbStatus.ColumnCount = 2
            Me.cmbStatus.ColumnWidths = "1000;2880;"
            Set Me.cmbStatus.RecordSet = oRs
        End If
    
    
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
        .sqlString = "usp_CONCEPT_Status_Report_Generic"
        .Parameters.Refresh
        
        If Nz(Me.cmbConcept, "") <> "" Then
            .Parameters("@pConceptId") = Me.cmbConcept
        End If
        
        If Nz(Me.cmbPayer, "") <> "" Then
            .Parameters("@pPayerNameID") = Me.cmbPayer
        End If
        
        If Nz(Me.cmbStatus, "") <> "" Then
            .Parameters("@pConceptStatus") = Me.cmbStatus
        End If
        
        If Nz(Me.cmbAuditor, "") <> "" Then
            .Parameters("@pAuditor") = Me.cmbAuditor
        End If
     
'        Set oRs = .OpenRecordSet()
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
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




Private Function GetOutOfLineReport() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetOutOfLineReport"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Status_Report_OutOfLine"
        .Parameters.Refresh
        
        If Nz(Me.cmbConcept, "") <> "" Then
            .Parameters("@pConceptId") = Me.cmbConcept
        End If
        
        If Nz(Me.cmbPayer, "") <> "" Then
            .Parameters("@pPayerNameID") = Me.cmbPayer
        End If
        
        If Nz(Me.cmbStatus, "") <> "" Then
            .Parameters("@pConceptStatus") = Me.cmbStatus
        End If
        
        If Nz(Me.cmbAuditor, "") <> "" Then
            .Parameters("@pAuditor") = Me.cmbAuditor
        End If
     
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If

    End With
    Debug.Print oRs.recordCount
    
    Set GetOutOfLineReport = oRs
Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub Form_Resize()

    '' Resize stuff..
    ' just in the details section
    Me.tabConceptStatusRpts.Width = Me.InsideWidth - 200
    Me.tabConceptStatusRpts.Height = Me.InsideHeight - 700
    
    
    Me.sfrm_Generic.Width = Me.tabConceptStatusRpts.Pages(0).Width - 100
    Me.sfrm_Generic.Height = Me.tabConceptStatusRpts.Pages(0).Height - 300
    Me.sfrm_StatusOutOfLine.Width = Me.tabConceptStatusRpts.Pages(1).Width - 100
    Me.sfrm_StatusOutOfLine.Height = Me.tabConceptStatusRpts.Pages(1).Height - 300
    

End Sub

Private Sub tabConceptStatusRpts_Change()
    Call RefreshData
End Sub
