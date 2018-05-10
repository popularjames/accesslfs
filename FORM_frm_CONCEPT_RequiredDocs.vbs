Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 10/17/2012
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
'''  - 10/17/2012 - Created
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
    Me.txtSelectedId = sConceptId
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
    Me.txtSelectedId = sValue
    Me.cmbConcept = csConceptId
    
    Call Me.RefreshData

Block_Exit:
    Exit Property
End Property


Public Sub RefreshData()
Dim strProcName As String
On Error GoTo Block_Err
Dim oRs As ADODB.RecordSet



    strProcName = ClassName & ".RefreshData"
    
    ' Refresh only the page selected..
    Select Case Me.tabConceptStatusRpts
    Case 0

        
        
        Set oRs = GetDocumentRS()

        If oRs Is Nothing Then
            Me.lvwDocuments.ListItems.Clear
        Else
            Call FillListBox(oRs)
        End If
        

        
''            oFrmGeneric.InitDataADO oRs, "v_Data_Database"
'            If oRs.RecordCount > 0 Then
''                Set oFrmGeneric.Recordset = oRs
'            End If
        
        
    Case 1

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


    Case 2
        Stop
    Case Else
        Stop
    End Select
    
    

Block_Exit:
'    Set oFrmGeneric = Nothing
    Set oRs = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub FillListBox(oRs As ADODB.RecordSet)
On Error GoTo Block_Err
Dim strProcName As String

Dim oLView As Object
Dim oLItem As Object

'Dim oLView As ListView
'Dim oLItem As ListItem

    strProcName = ClassName & ".FillListBox"
    
    If oRs Is Nothing Then GoTo Block_Exit
    
    Set oLView = Me.lvwDocuments
    oLView.ListItems.Clear
    
    While Not oRs.EOF
        Set oLItem = oLView.ListItems.Add(, , Nz(oRs("ReviewTypeName").Value, ""))
        oLItem.SubItems(1) = Nz(oRs("DataTypeCode").Value, "")
        oLItem.SubItems(2) = Nz(oRs("DocTypeId").Value, "")
        oLItem.SubItems(3) = Nz(oRs("DocName").Value, "")
        oLItem.SubItems(4) = Nz(oRs("Description").Value, "")
        oLItem.SubItems(5) = Nz(oRs("PerPayer").Value, 0)
        oLItem.SubItems(6) = Nz(oRs("CnlyAttachType").Value, "")
        
        oRs.MoveNext
    Wend
    
    
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Err
End Sub






Private Sub cmbDataType_AfterUpdate()
    Call Me.RefreshData
End Sub

Private Sub cmbReviewType_AfterUpdate()
    Call Me.RefreshData
End Sub

Private Sub cmdClear_Click()
    Me.cmbConcept = ""
    Me.cmbDataType = ""
    Me.cmbReviewType = ""
    Call Me.RefreshData
End Sub

Private Sub cmdRefresh_Click()
    Call Me.RefreshData
End Sub

Private Sub Command54_Click()
Stop    ' get your column widths..

'Dim oLView As ListView
'Dim oLItem As ListItem
'Dim oLCol As ColumnHeader
Dim i As Integer

Dim oLView As Object
Dim oLItem As Object
Dim oLCol As Object

    Set oLView = Me.lvwDocuments
    For Each oLCol In oLView.ColumnHeaders
        i = i + 1
        'Debug.Print oLCol.Name & " " & CStr(oLView.ColumnHeaders(i).width)
        ' oLCol
        Debug.Print oLCol.Key & " " & CStr(oLCol.Width)
    Next


End Sub

Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
    
    
'    Me.sfrm_Generic.Form.AllowFilters = True
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
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "select distinct CnlyReviewTypeCode, ReviewTypeName FROM XREFReviewType"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbReviewType.ColumnCount = 2
            Me.cmbReviewType.ColumnWidths = "1000;2880;"
            Set Me.cmbReviewType.RecordSet = oRs
        End If
    
        .ConnectionString = GetConnectString("v_Data_Database")
        .sqlString = "SELECT DataType, DataTypeDesc FROM XREF_DataType"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbDataType.ColumnCount = 2
            Me.cmbDataType.ColumnWidths = "1000;2880;"
            Set Me.cmbDataType.RecordSet = oRs
        End If
    

        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct ConceptId, ConceptDesc from CONCEPT_Hdr"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbConcept.ColumnCount = 2
            Me.cmbConcept.ColumnWidths = "1000;2880;"
            Set Me.cmbConcept.RecordSet = oRs
        End If
        

'        .SQLstring = "SELECT DISTINCT Auditor FROM CONCEPT_Hdr WHERE Auditor IS NOT NULL ORDER BY Auditor"
'        Set oRs = .ExecuteRS
'        If .GotData Then
'            Me.cmbAuditor.ColumnCount = 1
'            Me.cmbAuditor.ColumnWidths = "2880;"
'            Set Me.cmbAuditor.Recordset = oRs
'        End If
'
'
'
'        .SQLstring = "SELECT PayerNameId, PayerName FROM XREF_Payernames WHERE PayerNameId > 999 ORDER BY PayerName"
'        Set oRs = .ExecuteRS
'        If .GotData Then
'            Me.cmbPayer.ColumnCount = 2
'            Me.cmbPayer.ColumnWidths = "1000;2880;"
'            Set Me.cmbPayer.Recordset = oRs
'        End If
'
'
'
'
'
    End With
    
    
    
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Function GetDocumentRS(Optional sConceptId As String) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetDocumentRS"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_GetRequiredDocs"
        .Parameters.Refresh
        
        If Nz(Me.cmbConcept, "") <> "" Then
        
        
'        If sConceptId <> "" Then
            '.Parameters("@pConceptId") = sConceptId
            .Parameters("@pConceptId") = Me.cmbConcept
        Else
            If Nz(Me.cmbReviewType, "") <> "" Then
                .Parameters("@pReviewTypeId") = Me.cmbReviewType
            End If
            
            If Nz(Me.cmbDataType, "") <> "" Then
                .Parameters("@pDataType") = Me.cmbDataType
            End If
        
        End If
        
     
'        Set oRs = .OpenRecordSet()
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If

    End With
    Debug.Print oRs.recordCount
    
    Set GetDocumentRS = oRs
Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




'Private Function GetOutOfLineReport() As ADODB.Recordset
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim sSql As String
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".GetOutOfLineReport"
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("v_CODE_Database")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_CONCEPT_Status_Report_OutOfLine"
'        .Parameters.Refresh
'
'        If Nz(Me.cmbConcept, "") <> "" Then
'            .Parameters("@pConceptId") = Me.cmbConcept
'        End If
'
'        If Nz(Me.cmbPayer, "") <> "" Then
'            .Parameters("@pPayerNameID") = Me.cmbPayer
'        End If
'
'        If Nz(Me.cmbStatus, "") <> "" Then
'            .Parameters("@pConceptStatus") = Me.cmbStatus
'        End If
'
'        If Nz(Me.cmbAuditor, "") <> "" Then
'            .Parameters("@pAuditor") = Me.cmbAuditor
'        End If
'
'        Set oRs = .ExecuteRS
'        If .GotData = False Then
'            GoTo Block_Exit
'        End If
'
'    End With
'    Debug.Print oRs.RecordCount
'
'    Set GetOutOfLineReport = oRs
'Block_Exit:
'    Exit Function
'
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function


Private Sub Form_Resize()

    '' Resize stuff..
    ' just in the details section
    Me.tabConceptStatusRpts.Width = Me.InsideWidth - 200
    Me.tabConceptStatusRpts.Height = Me.InsideHeight - 700
    
    
    Me.lvwDocuments.Width = Me.tabConceptStatusRpts.Width
    Me.lvwDocuments.Height = Me.tabConceptStatusRpts.Height
    
    
'    Me.sfrm_Generic.width = Me.tabConceptStatusRpts.Pages(0).width - 100
'    Me.sfrm_Generic.Height = Me.tabConceptStatusRpts.Pages(0).Height - 300
'    Me.sfrm_StatusOutOfLine.width = Me.tabConceptStatusRpts.Pages(1).width - 100
'    Me.sfrm_StatusOutOfLine.Height = Me.tabConceptStatusRpts.Pages(1).Height - 300
    

End Sub

Private Sub tabConceptStatusRpts_Change()
    Call RefreshData
End Sub
