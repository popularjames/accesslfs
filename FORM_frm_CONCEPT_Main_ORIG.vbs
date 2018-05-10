Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private WithEvents oMainGrid As Form_frm_GENERAL_Datasheet_ADO
Private WithEvents oMainGrid As Form_frm_GENERAL_Datasheet
Attribute oMainGrid.VB_VarHelpID = -1
Private WithEvents frmConceptHdr As Form_frm_CONCEPT_Hdr
Attribute frmConceptHdr.VB_VarHelpID = -1
Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1

Private cdctSFrmRefreshTimes As Scripting.Dictionary

Private WithEvents ofrmNewConcept As Form_frm_CONCEPT_New_Concept
Attribute ofrmNewConcept.VB_VarHelpID = -1
Private cstrNewConceptId As String

Private mNoteID As Long
Private mrsConcept As ADODB.RecordSet
Private mrsNotes As ADODB.RecordSet

Private mstrUserProfile As String
Private mstrUserName As String
Private miAppPermission As Integer
Private strCurrentStatus As String

Private mbAllowChange As Boolean
Private mbAllowView As Boolean
Private mbAllowAdd As Boolean

Private clSelectedPayerNID As Long


Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Const CstrFrmAppID As String = "ConceptHdr"
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get SelectedPayerNameId() As Long
    If clSelectedPayerNID = 0 Then
        ' Reach into the subform and get it..
        clSelectedPayerNID = Me.subFrmMain.Form.Controls("cmbPayer").Value
    End If
    SelectedPayerNameId = clSelectedPayerNID
End Property
Public Property Let SelectedPayerNameId(lPayerNameId As Long)
    clSelectedPayerNID = lPayerNameId
    ' reach into subform and select it in the dropdown
'    If Me.subFrmMain.Form.Controls("cmbPayer").Value <> lPayerNameId Then
'        Me.subFrmMain.Form.Controls("cmbPayer").Value = clSelectedPayerNID
'    End If
End Property


Public Sub RefreshData()
Dim strError As String
Dim sDefaultWhere As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim strProcName As String
Dim ctl As Control

    On Error GoTo ErrHandler
    
    strProcName = TypeName(Me) & ".RefreshData"
    
    DoCmd.Hourglass True
    DoCmd.Echo True, "Searching ..."

    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_ConMgmt_Search"
        .Parameters.Refresh
        .Parameters("@pKeyword") = Nz(Me.txtSearchBox, "")
        .Parameters("@pSearchAllFields") = IIf(Me.ckExpandSearch, 1, 0)
        .Parameters("@pSearchCodes") = IIf(Me.ckIncludeCodes, 1, 0)
        .Parameters("@pAllFieldWhereClause") = ""
        Set oRs = .ExecuteRS
        If .Parameters("@pErrMsg").Value <> "" Then
            LogMessage strProcName, "ERROR", "Problem searching for a concept", "Keyword: " & Nz(Me.txtSearchBox, "") & " Expand Search: " & IIf(Me.ckExpandSearch, 1, 0) & " Include Codes: " & IIf(Me.ckIncludeCodes, 1, 0)
        End If
    End With
        
    
    'Refresh the grid based on the rowsource passed into the form
'    Me.frm_GENERAL_Datasheet.Form.InitDataADO oRs, "v_ConceptMgmt_MainGrid_View"
    Me.frm_GENERAL_Datasheet.Form.InitData oRs, "v_ConceptMgmt_MainGrid_View"
    
    Set Me.frm_GENERAL_Datasheet.Form.RecordSet = oRs
    Set oMainGrid = Me.frm_GENERAL_Datasheet.Form
    
    oMainGrid.AllowFilters = True
    DoCmd.Echo True, "Refreshing grids"
    
    'Loop through the controls and size them correctly.
    For Each ctl In Me.frm_GENERAL_Datasheet.Form.Controls
      If ctl.ControlType = acTextBox Then
          ctl.ColumnWidth = -2
      End If
   Next
   oMainGrid_Current
   
   

   
exitHere:
    DoCmd.Echo True, "Ready..."
    DoCmd.Hourglass False

Exit Sub
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshDataADO"
    Resume exitHere
End Sub



Private Sub cmdConceptStatusReports_Click()
    DoCmd.OpenForm "frm_CONCEPT_Status_Reports", acNormal, , , , acWindowNormal
End Sub

'''Public Sub RefreshData() ' LEGACY
'''Dim sTableList As String
'''Dim sWhere As String
'''Dim sFrom As String
'''Dim StrSQL As String
'''Dim strError As String
'''Dim sDefaultWhere As String
'''Dim oAdo As clsADO
'''Dim oRs As ADODB.Recordset
'''Dim sDaoSql As String
'''
'''    On Error GoTo ErrHandler
'''
'''    DoCmd.Hourglass True
'''    DoCmd.Echo True, "Searching ..."
'''
'''    'Build the SQL string and query the data
'''    StrSQL = "     select DISTINCT CONCEPT_Hdr.ConceptID,"
'''    StrSQL = StrSQL & "         ClientIssueNum, "
'''    StrSQL = StrSQL & "         ConceptDesc, "
'''    StrSQL = StrSQL & "         ConceptSource, "
'''    StrSQL = StrSQL & "         DataType,  "
'''    StrSQL = StrSQL & "         CONCEPT_Hdr.ConceptLevel,"
'''    StrSQL = StrSQL & "         CONCEPT_Hdr.ReviewType,"
'''    StrSQL = StrSQL & "         CONCEPT_Hdr.OpportunityType,"
'''    StrSQL = StrSQL & "         CONCEPT_Hdr.ConceptStatus,"
'''    StrSQL = StrSQL & "         StatusDescription as CStatus,"
'''    StrSQL = StrSQL & "         ConceptGroup,"
'''    StrSQL = StrSQL & "         Auditor,"
'''    StrSQL = StrSQL & "         CreateDate,"
'''    StrSQL = StrSQL & "         LastUpDt,"
'''    StrSQL = StrSQL & "         LastUpUser,"
'''    'strSQL = strSQL & "         ConceptRationale,"
'''    StrSQL = StrSQL & "         NoteID,"
'''    StrSQL = StrSQL & "         AccountID,  "
'''    StrSQL = StrSQL & "         ConceptLogic"
'''
'''
'''    sFrom = " FROM ((CONCEPT_Hdr LEFT JOIN CONCEPT_XREF_Level ON CONCEPT_Hdr.ConceptLevel = CONCEPT_XREF_Level.ConceptLevel) LEFT JOIN CONCEPT_XREF_Opportunity ON CONCEPT_Hdr.OpportunityType = CONCEPT_XREF_Opportunity.OpportunityType) LEFT JOIN CONCEPT_XREF_Status ON CONCEPT_Hdr.ConceptStatus = CONCEPT_XREF_Status.ConceptStatus "
'''
'''
'''    sDefaultWhere = sDefaultWhere & " (CONCEPT_Hdr.ConceptDesc like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR CONCEPT_Hdr.ConceptRationale like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR CONCEPT_Hdr.ConceptLogic like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR CONCEPT_Hdr.ConceptID like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR CONCEPT_Hdr.ClientIssueNum like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR CONCEPT_Hdr.Auditor like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34)
'''    sDefaultWhere = sDefaultWhere & " OR EditParameters like " & Chr(34) & "*" & Me.txtSearchBox & "*" & Chr(34) & ")"
'''
'''    If Nz(Me.txtSearchBox, "") & "" <> "" Then
'''        If Me.ckIncludeCodes = True Then
'''            sFrom = Replace(sFrom, "FROM ((", "FROM (((", 1, 1, vbTextCompare)
'''            sFrom = sFrom & ") LEFT JOIN CONCEPT_Dtl_Codes ON CONCEPT_Hdr.ConceptId = CONCEPT_Dtl_Codes.ConceptID "
'''        End If
'''    Else
'''    End If
'''
'''
'''    StrSQL = StrSQL & sFrom '   & " )"
'''
'''    'Check to see if the user had built on criteria
'''    'TL add account id logic
'''
'''    '' 20120323 KD expanded search capability.. NOTE: if this starts taking too long
'''    '' change it to ADO...
'''    If Nz(Me.txtSearchBox, "") & "" <> "" Then
'''
'''        If Me.ckExpandSearch = True Then
'''            sTableList = "CONCEPT_Hdr,CONCEPT_XREF_Level,CONCEPT_XREF_Opportunity,CONCEPT_XREF_Status,"
'''        End If
'''        If Me.ckIncludeCodes = True Then
'''            sTableList = sTableList & "CONCEPT_Dtl_Codes"
'''        End If
'''        If sTableList <> "" Then
'''            sWhere = mod_Concept_Specific.MakeWhereListFromTableList(sTableList, Me.txtSearchBox, False, False)
'''            If Me.ckIncludeCodes = True And Me.ckExpandSearch = False Then
'''                    ' Have to remove the last paran from the one and the beginning paren from the other and
'''                    '' put in an OR in the place:
'''                    ' FROM: like "*EDIC*")([CONCEPT_Dtl_Codes].[ConceptID] LIKE '*EDIC*'
'''                    ' TO: like "*EDIC*" OR [CONCEPT_Dtl_Codes].[ConceptID] LIKE '*EDIC*'
'''                sWhere = left(sWhere, Len(sWhere) - 1)
'''                sDefaultWhere = Right(Trim(sDefaultWhere), Len(Trim(sDefaultWhere)) - 1)
'''
'''                sWhere = sWhere & " OR " & sDefaultWhere
'''            End If
'''
'''        Else
'''           sWhere = sDefaultWhere
'''        End If
'''
'''    End If
'''
'''    StrSQL = StrSQL & " WHERE  AccountID = " & gintAccountID & IIf(sWhere = "", "", " AND " & sWhere)
'''
'''    sDaoSql = StrSQL
'''
'''    StrSQL = Replace(StrSQL, """", "'", 1, -1, vbBinaryCompare)
'''    StrSQL = Replace(StrSQL, "*", "%", 1, -1, vbBinaryCompare)
'''
'''
'''
'''    Set oAdo = New clsADO
'''    With oAdo
'''        .ConnectionString = GetConnectString("CONCEPT_Hdr")
'''        .SQLTextType = sqltext
'''        .SQLstring = StrSQL
'''        Set oRs = .ExecuteRS
'''
'''    End With
'''
'''
'''
'''    'Refresh the grid based on the rowsource passed into the form
'''    Me.frm_GENERAL_Datasheet.Form.InitData sDaoSql, 2
'''    Me.frm_GENERAL_Datasheet.Form.RecordSource = sDaoSql
''''    Set Me.frm_GENERAL_Datasheet.Form.Recordset = oRs
'''    Set oMainGrid = Me.frm_GENERAL_Datasheet.Form
'''
'''    DoCmd.Echo True, "Refreshing grids"
'''
'''
'''    Dim ctl As Control
'''    'Loop through the controls and size them correctly.
'''    For Each ctl In Me.frm_GENERAL_Datasheet.Form.Controls
'''      If ctl.ControlType = acTextBox Then
'''          ctl.ColumnWidth = -2
'''      End If
'''   Next
'''   oMainGrid_Current
'''
'''ExitHere:
'''    DoCmd.Echo True, "Ready..."
'''    DoCmd.Hourglass False
'''
'''Exit Sub
'''ErrHandler:
'''    strError = Err.Description
'''    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshDataADO"
'''    Resume ExitHere
'''End Sub






Private Sub cmdddNote_Click()
    Dim bNotes As Boolean
On Error GoTo Err_cmdddNote_Click
    
     Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
    
     frmGeneralNotes.frmAppID = Me.frmAppID
     Set frmGeneralNotes.NoteRecordSource = mrsNotes
     frmGeneralNotes.RefreshData
     ShowFormAndWait frmGeneralNotes
     lstTabs_Click
     Set frmGeneralNotes = Nothing

Exit_cmdddNote_Click:
    Exit Sub

Err_cmdddNote_Click:
    MsgBox Err.Description
    Resume Exit_cmdddNote_Click
End Sub
Private Sub cmdNew_Click()

    Me.subFrmMain.SourceObject = ""
    Set ofrmNewConcept = New Form_frm_CONCEPT_New_Concept
    ShowFormAndWait ofrmNewConcept
    
    Me.txtSearchBox = cstrNewConceptId
    
    Call cmdSearch_Click
    
End Sub

Private Sub cmdRefresh_Click()
  Me.txtSearchBox = ""
    '  Me.RefreshData
    ' KD: SearchChange, From DAO to ADO
    RefreshData
  
End Sub

Private Sub cmdSearch_Click()
    RefreshData
End Sub



'Private Sub Command18_Click()
'
'On Error GoTo ErrHandler
''
''    Dim StrSQL As String
''    Dim strFormValue As String
''    Dim strSQLCharacter As String
''    Dim strSQLValue As String
''    Dim strFormName As String
''    Dim lngNoteID As Long
''
'
'
'
'    Me.subFrmMain.SourceObject = "frm_CONCEPT_AddPayer"
'    Me.subFrmMain.Form.FormConceptID = Me.txtConceptID
'    Set Me.subFrmMain.Form.Recordset = mrsConcept
''    Set Me.subFrmMain.Form.ConceptRecordSource = mrsConcept
'
'    Me.subFrmMain.Form.RefreshData
'
'    Me.lblTabs.Caption = Me.lstTabs
'
'
'Exit Sub
'ErrHandler:
'    MsgBox Err.Description
'
'End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer


    Me.frm_GENERAL_Datasheet.Form.RecordSource = ""
    Me.subFrmMain.Form.RecordSource = ""

    Me.Caption = "Concept Maintenance"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub
    
'    Select Case UCase(gstrProfileID)
'    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
        Me.cmdConceptStatusReports.visible = True
'    Case Else
'        Me.cmdConceptStatusReports.visible = False
'    End Select
    
    miAppPermission = GetAppPermission(Me.frmAppID)
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowView = (miAppPermission And gcAllowView)
    
    If mbAllowChange Then
        Me.cmdddNote.Enabled = True
    Else
        Me.cmdddNote.Enabled = False
    End If
    
    If mbAllowAdd Then
        Me.CmdNew.Enabled = True
    Else
        Me.CmdNew.Enabled = False
    End If
    
    
    
    If mbAllowView Then
'        lstTabs.RowSource = GetListBoxSQL(Me.Name)
        lstTabs.RowSource = GetListBoxSQL("frm_CONCEPT_Main")
        If lstTabs.ListCount > 1 Then
            Me.lstTabs = Me.lstTabs.ItemData(0)
        End If
        
    Else
        MsgBox "You do not have permission to view this form.  Please contact your system admin", vbInformation
        DoCmd.Close acForm, Me.Name
    End If
    ' KD: SearchChange, From DAO to ADO
    RefreshData
    
End Sub

Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub frmConceptHdr_ConceptSaved()
    'Me.RefreshData
End Sub

Private Sub frmGeneralNotes_NoteAdded()
    If SaveData_Notes Then
        MsgBox "Note added"
        'RefreshData
    End If
End Sub
Private Function SaveData_Notes() As Boolean
    Dim bResult As Boolean
    On Error GoTo ErrHandler
    
    Set MyAdo = New clsADO
    Set myCode_ADO = New clsADO
    
    
    myCode_ADO.ConnectionString = GetConnectString("v_Code_Database")
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    
    If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
        'If the noteID is -1 then we need to create a new ID
        If mNoteID = -1 Then
            'This is a public function that gets a unique ID based on the app being passed to the method
            mNoteID = GetAppKey("NOTE")
        End If
            'Set the recordset of the header to contain the new note ID
            'Apply this new noteID to all of the records in the note recordset
            If Not (mrsNotes.BOF = True And mrsNotes.EOF = True) Then
                mrsNotes.MoveFirst
                While Not mrsNotes.EOF
                    mrsNotes.Update
                    mrsNotes("NoteID") = mNoteID
                    mrsNotes.MoveNext
                Wend
            End If
        'End If
        'Pass the recordset back to SQL synching the results
        bResult = myCode_ADO.Update(mrsNotes, "usp_NOTE_Detail_Apply")
    Else
        bResult = True
    End If
    If bResult Then
        MyAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.txtConceptID & "'"
        Set mrsConcept = MyAdo.OpenRecordSet()
        If Not mrsConcept.EOF Then
            mrsConcept.Fields("NoteID") = mNoteID
        End If
        bResult = myCode_ADO.Update(mrsConcept, "usp_CONCEPT_Hdr_Apply")
    End If
    
    
    SaveData_Notes = bResult
    
Exit_Sub:
    Exit Function
ErrHandler:
    'Rollback anything we did up until this point
    SaveData_Notes = False
    GoTo Exit_Sub
End Function



Private Sub lstTabs_Click()
On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    Dim lngNoteID As Long
    
'    Set MYADO = New clsADO
    

    If Me.lstTabs.ListIndex <> -1 Then
        Dim rs As DAO.RecordSet
            'Get a recordset of tabs for this form
'            Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstTabs.Column(1), Me.Name), dbOpenSnapshot, dbSeeChanges)
            Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstTabs.Column(1), "frm_CONCEPT_Main"), dbOpenSnapshot, dbSeeChanges)
            If Not (rs.BOF And rs.EOF) Then
                Select Case rs("FormName")
                    Case "frm_CONCEPT_RequiredDocs"
                        '' IdValue
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Me.subFrmMain.Form.IdValue = Me.txtConceptID
'                        Set Me.subFrmMain.Form.Recordset = mrsConcept
'                        Set Me.subFrmMain.Form.ConceptRecordSource = mrsConcept
'                        Me.subFrmMain.Form.RefreshData
'                        Call SetSubFormPayerSel
                        Me.subFrmMain.Form.RefreshData
                    'Everytime there is a new tab, we have to add a case statement to make sure the form loads correctly
                    ' (kd: OR, we could just standardize the code interface for each of the subforms!!!)
                    Case "frm_Concept_Hdr"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Me.subFrmMain.Form.FormConceptID = Me.txtConceptID
                        Set Me.subFrmMain.Form.RecordSet = mrsConcept
                        Set Me.subFrmMain.Form.ConceptRecordSource = mrsConcept
                        Me.subFrmMain.Form.RefreshData
                        Call SetSubFormPayerSel
                    Case "frm_GENERAL_Notes_Display"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Set Me.subFrmMain.Form.NoteRecordSource = mrsNotes
                        
                        If RefreshSubform(rs("FormName")) = True Then
                            Me.subFrmMain.Form.RefreshData
                        End If
                        
                        ColObjectInstances.Add Item:=Me.subFrmMain.Form, Key:=Me.subFrmMain.Form.hwnd & " "
                        
                    Case "frm_Concept_Dtl_Codes"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        ColObjectInstances.Add Item:=Me.subFrmMain.Form, Key:=Me.subFrmMain.Form.hwnd & " "
                        Call SetSubFormPayerSel
                    Case "frm_Concept_Dtl_State"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        ColObjectInstances.Add Item:=Me.subFrmMain.Form, Key:=Me.subFrmMain.Form.hwnd & " "
                        Call SetSubFormPayerSel
                    Case "frm_CONCEPT_AddPayer"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Me.subFrmMain.Form.FormConceptID = Me.txtConceptID
                        Me.lblTabs.Caption = Me.lstTabs
                        
                        If RefreshSubform(rs("FormName")) = True Then
                            Me.subFrmMain.Form.RefreshData
                        End If
                    Case Else
                        Me.subFrmMain.SourceObject = rs("FormName")
                            
                        If rs("FormName") = "frm_CONCEPT_References_Grid_View" Then
                            Me.subFrmMain.Form.FieldReference = "ConceptID"
                            Me.subFrmMain.Form.FieldValue = Me.txtConceptID
                            Me.subFrmMain.Form.IdValue = Me.txtConceptID    '' KD 20120416
                            Call SetSubFormPayerSel
                        End If

                        
                        If rs("FormName") = "frm_CONCEPT_Tagged_Claims" Then
                            Me.subFrmMain.Form.IdValue = Me.txtConceptID
                            
                            Me.subFrmMain.Form.RefreshData
                            Call SetSubFormPayerSel
                        End If
                        
                        strSQL = GetNavigateTabSQL(lstTabs.Column(1), Me, strFormValue, strSQLCharacter, strSQLValue, strFormName)
                        If rs("FormName") = "frm_CONCEPT_References_Grid_View" Then
                            If strSQL <> "" Then strSQL = strSQL & " order by ConceptID, RefSequence"
                            Me.subFrmMain.Form.RefreshData
                            Call SetSubFormPayerSel
                        End If
                
                        
                        If strSQL <> "" Then
                            Me.subFrmMain.Form.CnlyRowSource = strSQL
                            
                            'commented this because it was preventing the tagged claims to show correctly when
                            'switching payers quicky (not refreshing for the right payer)
                            'JS 07/26/2012
                            Me.subFrmMain.Form.RefreshData
                            ColObjectInstances.Add Item:=Me.subFrmMain.Form, Key:=Me.subFrmMain.Form.hwnd & " "
                            Call SetSubFormPayerSel
                        End If

'                        If InStr(1, Me.subFrmMain.Form.Name, "frm_GENERAL_Datasheet", vbTextCompare) > 0 Then
                            
                        If InStr(1, rs("FormName"), "frm_GENERAL_Datasheet", vbTextCompare) > 0 Then
                            If RefreshSubform(Me.subFrmMain.Form.Name) = True Then
                                Me.subFrmMain.Form.RefreshData
                            End If
                        End If
                End Select


                
                Me.lblTabs.Caption = Me.lstTabs
            Else
                MsgBox "Application form has not been defined"
            End If
    End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub


Private Sub SetSubFormPayerSel()
                ' set the cmbPayer value
        If Me.subFrmMain.Form.Name <> "frm_GENERAL_Tab" Then
            Me.subFrmMain.Form.Controls("cmbPayer").Value = Me.SelectedPayerNameId
            Call Me.subFrmMain.Form.PayerChange
        End If
End Sub

Public Sub SetSubformRefreshTime(ByVal sSubFormName As String)
    sSubFormName = UCase(sSubFormName)
    
    If cdctSFrmRefreshTimes Is Nothing Then
        Set cdctSFrmRefreshTimes = New Scripting.Dictionary
    End If

    If cdctSFrmRefreshTimes.Exists(sSubFormName) = True Then
        cdctSFrmRefreshTimes.Item(sSubFormName) = Now()
    Else
        cdctSFrmRefreshTimes.Add sSubFormName, Now()
    End If
    
End Sub

Private Function RefreshSubform(ByVal sSubFormName As String) As Boolean
    sSubFormName = UCase(sSubFormName)
    
    If cdctSFrmRefreshTimes Is Nothing Then
        Set cdctSFrmRefreshTimes = New Scripting.Dictionary
    End If
    
    If cdctSFrmRefreshTimes.Exists(sSubFormName) = True Then
        RefreshSubform = IIf(DateDiff("s", cdctSFrmRefreshTimes.Item(sSubFormName), Now()) > 2, True, False)
        cdctSFrmRefreshTimes.Item(sSubFormName) = Now()
    Else
        RefreshSubform = True
        cdctSFrmRefreshTimes.Add sSubFormName, Now()
    End If

End Function

Private Sub ofrmNewConcept_ConceptSaved(strNewConceptId As String)
    cstrNewConceptId = strNewConceptId
End Sub

Private Sub oMainGrid_Current()
    Set MyAdo = New clsADO

    Me.txtConceptID = Nz(oMainGrid.Controls("ConceptID"), "")
    Me.txtNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
    mNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
    'Refresh the tabs to ensure the main form is in sync with the other forms.
    
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    
    MyAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.txtConceptID & "'"
    Set mrsConcept = MyAdo.OpenRecordSet()
        
    MyAdo.sqlString = " SELECT * from Note_Detail WHERE NoteID = '" & Me.txtNoteID & "'"
    Set mrsNotes = MyAdo.OpenRecordSet()


    lstTabs_Click
End Sub

Public Property Get TabSelected() As Integer
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = TypeName(Me) & ".TabSelected"
    TabSelected = Me.lstTabs.ListIndex
        
    
Block_Exit:
    Exit Property
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Property

Public Property Let TabSelected(iItemToSelect As Integer)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = TypeName(Me) & ".TabSelected"
    Me.lstTabs.ListIndex = iItemToSelect
    
Block_Exit:
    Exit Property
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Property




Private Sub txtSearchBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ' don't care if it's shift, alt, or ctrl
        RefreshData
    End If
End Sub
