Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents oMainGrid As Form_frm_GENERAL_Datasheet_DAO
Attribute oMainGrid.VB_VarHelpID = -1
Private WithEvents frmConceptHdr As Form_frm_CONCEPT_Hdr
Attribute frmConceptHdr.VB_VarHelpID = -1
Private WithEvents frmGeneralNotes As Form_frm_GENERAL_Notes_Add
Attribute frmGeneralNotes.VB_VarHelpID = -1

Private WithEvents filterForm As Form_SCR_ScreensFilters
Attribute filterForm.VB_VarHelpID = -1

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

Private mbSearching As Boolean

Private mbAllowChange As Boolean
Private mbAllowView As Boolean
Private mbAllowAdd As Boolean

Private clSelectedPayerNID As Long
''' For form resizing and 'stuff'
Private csgSplitter As Single
Private genUtils As New CT_ClsGeneralUtilities

Private gbTogValue As Boolean

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Const CstrFrmAppID As String = "ConceptHdr"
Private cbLayoutApplied As Boolean

' KD Not sure what is going on - shelly is only getting 1 search result

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ScreenID() As Long
    ScreenID = 1
End Property

Public Property Let SetTogValue(ByVal Value As Boolean)
    gbTogValue = Value
End Property
Public Property Get GetTogValue() As Boolean
    GetTogValue = gbTogValue
End Property


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get Searching() As Boolean
    Searching = mbSearching
End Property
Public Property Let Searching(bSearching As Boolean)
    mbSearching = bSearching
End Property


Public Property Get LayoutApplied() As Boolean
    LayoutApplied = cbLayoutApplied
End Property
Public Property Let LayoutApplied(bLayoutApplied As Boolean)
    cbLayoutApplied = bLayoutApplied
End Property


Public Property Get SelectedPayerNameId() As Long
    If clSelectedPayerNID = 0 Then
        ' Reach into the subform and get it..
        clSelectedPayerNID = Nz(Me.subFrmMain.Form.Controls("cmbPayer").Value, 1000)
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
On Error GoTo Block_Err
Dim strProcName As String
Dim strError As String
Dim sDefaultWhere As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim ctl As Control
Dim lContractId As Long
    
    strProcName = TypeName(Me) & ".RefreshData"
    
    DoCmd.Hourglass True
    DoCmd.Echo True, "Searching ..."

    '' Need to release the lock on the table
    Me.frm_GENERAL_Datasheet.Form.RecordSource = ""
    
    lContractId = Me.cmbContractID
    

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_DATABASE")
        .SQLTextType = StoredProc
        .sqlString = "usp_ConMgmt_Search_2010"   ' _2010
'        .sqlString = "usp_ConMgmt_Search_NC"   ' _2010
        .Parameters.Refresh
        .Parameters("@pKeyword") = Nz(Me.txtSearchBox, "")
        .Parameters("@pSearchAllFields") = IIf(Me.ckExpandSearch, 1, 0)
        .Parameters("@pSearchCodes") = IIf(Me.ckIncludeCodes, 1, 0)
        .Parameters("@pAllFieldWhereClause") = ""
        .Parameters("@pContractId") = lContractId
        Set oRs = .ExecuteRS
        If .Parameters("@pErrMsg").Value <> "" Then
            LogMessage strProcName, "ERROR", "Problem searching for a concept", "Keyword: " & Nz(Me.txtSearchBox, "") & " Expand Search: " & IIf(Me.ckExpandSearch, 1, 0) & " Include Codes: " & IIf(Me.ckIncludeCodes, 1, 0)
        End If
    End With
        
    
    '' 20121214 KD: Going to change the sproc to populate a table for each user..
    '' we'll then try a pass through query to bind to the form and see if the right click filtering / sorting works.
    
    'Refresh the grid based on the rowsource passed into the form
'    Me.frm_GENERAL_Datasheet.Form.InitDataADO oRs, "v_ConceptMgmt_MainGrid_View"

'    Me.frm_GENERAL_Datasheet.Form.InitData "CONCEPT_ConMgmtSearch_NC", 2  '', "v_ConceptMgmt_MainGrid_View"
'    Me.frm_GENERAL_Datasheet.Form.InitData "v_ConceptMgmt_MainGrid_View", 2  '', "v_ConceptMgmt_MainGrid_View"

    If oRs Is Nothing Then
        Stop
    End If
    If oRs.State = adStateOpen Then
        If oRs.EOF And oRs.BOF Then
    '    If oRs.recordCount < 1 Then
            LogMessage strProcName, "USER MSG", "No results found!", , True
            GoTo Block_Exit
        End If
Else
'Stop
'            LogMessage strProcName, "USER MSG", "No results found!", , True
'            GoTo Block_Exit
    End If
    
    Me.frm_GENERAL_Datasheet.Form.InitData "CONCEPT_ConMgmtSearch", 2  '', "v_ConceptMgmt_MainGrid_View"

'    Set Me.frm_GENERAL_Datasheet.Form.Recordset = oRs
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

   If Me.LayoutApplied = True Then
        Call cmdLayoutApply_Click
   End If
    
    If DCount("*", "CONCEPT_ConMgmtSearch", "SearchUserId = '" & Identity.UserName & "'") = 0 Then
        LogMessage strProcName, "USER NOTICE", "No records found!", Nz(Me.txtSearchBox, "") & "::" & IIf(Me.ckExpandSearch, 1, 0) & "::" & IIf(Me.ckIncludeCodes, 1, 0) & "::" & CStr(lContractId), True
    End If
    
   
Block_Exit:
    DoCmd.Echo True, "Ready..."
    DoCmd.Hourglass False

    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Sub RefreshData_ADO_Version()
Dim strError As String
Dim sDefaultWhere As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim strProcName As String
Dim ctl As Control

    On Error GoTo ErrHandler
    
    strProcName = TypeName(Me) & ".RefreshData_ADO_Version"
    
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




Private Sub ckSharedLayouts_AfterUpdate()
   
    Call PopulateListLayouts(Me.CmboLayouts, IIf(Me.ckSharedLayouts.Value = 0, False, True))
    
End Sub

Private Sub CmboFilters_Click()
On Error GoTo Block_Err
Dim strProcName As String
    'SA 03/22/2012 - CR2667 Add filter to list when selected
Dim sFilterName As String
Dim sFilterID As String
Dim bFilterFound As Boolean
Dim i As Integer
    
    strProcName = ClassName & ".CmboFilters_Click"
    
    bFilterFound = False
    
    If LenB(Nz(Me.CmboFilters.Value, vbNullString)) > 0 Then
        'Get filter name
        sFilterID = Me.CmboFilters.Value
        For i = 0 To CmboFilters.ListCount - 1
            If CmboFilters.Column(0, i) = sFilterID Then
                sFilterName = Me.CmboFilters.Column(1, i)
                Exit For
            End If
        Next
        
        'Check to see if filter is already in the list
        For i = 0 To Me.cmboFiltersSelected.ListCount - 1
            If sFilterID = Me.cmboFiltersSelected.Column(0, i) Then
                bFilterFound = True
                Exit For
            End If
        Next
        
        'Add to list if not already there
        If Not bFilterFound Then
            If LenB(cmboFiltersSelected.RowSource) = 0 Then
                cmboFiltersSelected.RowSource = sFilterID & ";'" & Replace(sFilterName, "'", "''") & "'"
            Else
                cmboFiltersSelected.RowSource = cmboFiltersSelected.RowSource & ";" & sFilterID & ";'" & Replace(sFilterName, "'", "''") & "'"
            End If
        End If
    End If
Block_Exit:
    CmboFilters = vbNullString
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub CmboFilters_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.CmboFilters.Dropdown
End Sub

Private Sub cmboLayouts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.CmboLayouts.Dropdown
End Sub

Private Sub cmdConceptStatusReports_Click()
    DoCmd.OpenForm "frm_CONCEPT_Status_Reports", acNormal, , , , acWindowNormal
End Sub


Private Sub cmdddNote_Click()
Dim bNotes As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdddNote_Click"
    
     Set frmGeneralNotes = New Form_frm_GENERAL_Notes_Add
    
     frmGeneralNotes.frmAppID = Me.frmAppID
     Set frmGeneralNotes.NoteRecordSource = mrsNotes
     frmGeneralNotes.RefreshData
     ShowFormAndWait frmGeneralNotes
     lstTabs_Click
     Set frmGeneralNotes = Nothing

Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdFieldFinder_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim sFind As String
Dim bFound As Boolean
Static lLastColTried As Long
Dim lThisCtl As Long
Dim lWrapBack As Long
Dim bWrapped  As Boolean
Dim lEndOfLoop As Long

    strProcName = ClassName & ".cmdFieldFinder_Click"
    
    If Nz(Me.txtFieldFinder, "") = "" Then
        GoTo Block_Exit
    End If
    
    sFind = Me.txtFieldFinder
    
    If lLastColTried >= oMainGrid.Controls.Count Then
        lLastColTried = 0
        lEndOfLoop = oMainGrid.Controls.Count - 1
    Else
        lWrapBack = lLastColTried - 1
        lEndOfLoop = oMainGrid.Controls.Count - 1
    End If
    
WrapBack:
    For lThisCtl = lLastColTried To lEndOfLoop
        Set oCtl = oMainGrid.Controls(lThisCtl)

        If TypeName(oCtl) = "Textbox" Then

            If Me.ckFieldFinderLike = False Then
                If UCase(oCtl.ControlSource) = UCase(sFind) Then
                    If oCtl.visible = True Then
                        oCtl.SetFocus
                        bFound = True
                        Exit For
                    End If
                End If
            Else
                If InStr(1, oCtl.ControlSource, sFind, vbTextCompare) > 0 Then
                    If oCtl.visible = True Then
                        oCtl.SetFocus
                        bFound = True
                        Exit For
                    End If
                End If
            End If
            
        End If
    Next
    
    lLastColTried = lThisCtl + 1
    
    ' If we get here and haven't found it
    ' should we wrap back around or not?
    If bFound = False And lWrapBack > 0 And bWrapped = False Then
        bWrapped = True
        lLastColTried = 0
        lEndOfLoop = lWrapBack
        GoTo WrapBack
    End If
    
    If bFound = False Then
        MsgBox "Could not find the field by that exact name.. "
    End If
    
Block_Exit:
    Set oCtl = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub CmdFilterEdit_Click()
'On Error Resume Next
    Set filterForm = New Form_SCR_ScreensFilters
    With filterForm
        .SetParent Me.Form
        .visible = True
        .Initialize
    End With
    
End Sub

Private Sub cmdFiltersAdd_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim strName As String
Dim FilterID As String
Dim i As Integer
Dim bFound As Boolean
    
    strProcName = ClassName & ".cmdFiltersAdd_Click"
    
    If Nz(Me.CmboFilters.Value, vbNullString) = vbNullString Then
        MsgBox "Please select a filter to add."
        Me.CmboFilters.SetFocus
        Me.CmboFilters.Dropdown
        Exit Sub
    Else
        FilterID = Me.CmboFilters.Value
        For i = 0 To CmboFilters.ListCount - 1
            If CmboFilters.Column(0, i) = FilterID Then
                strName = Me.CmboFilters.Column(1, i)
                Exit For
            End If
        Next
    End If
    
    bFound = False
    
    For i = 0 To Me.cmboFiltersSelected.ListCount - 1
        If FilterID = Me.cmboFiltersSelected.Column(0, i) Then
            bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then
        If cmboFiltersSelected.RowSource = vbNullString Then
            cmboFiltersSelected.RowSource = FilterID & ";'" & Replace(strName, "'", "''") & "'"
        Else
            cmboFiltersSelected.RowSource = cmboFiltersSelected.RowSource & ";" & FilterID & ";'" & Replace(strName, "'", "''") & "'"
        End If
    End If
    CmboFilters = vbNullString
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub CmdFiltersApply_Click()
On Error GoTo Block_Err
    'SA 03/22/2012 - CR2667 Added apply filter functionality
Dim ExistingFilter As String
Dim newFilter As String
Dim itemFilter As String
Dim i As Integer

    ExistingFilter = oMainGrid.filter

    'Check for right click filters
    If InStr(1, ExistingFilter, "[" & oMainGrid.Name & "]", vbTextCompare) = 0 Then
        newFilter = vbNullString    'Rebuild filter
    Else
        newFilter = ExistingFilter  'Append existing
    End If

    'Load filters from list and apply if not in exiting filter
    For i = 0 To cmboFiltersSelected.ListCount - 1
        itemFilter = DLookup("FilterSQL", "CA_ScreensFilters", "FilterID=" & Me.cmboFiltersSelected.Column(0, i))
        
        If InStr(1, newFilter, itemFilter, vbTextCompare) = 0 Then
            If LenB(newFilter) > 0 Then
                newFilter = newFilter & " AND (" & itemFilter & ")"
            Else
                newFilter = "(" & itemFilter & ")"
            End If
        End If
    Next

    'Apply filter if changed
    If newFilter <> ExistingFilter Then
        oMainGrid.filter = newFilter
        oMainGrid.FilterOn = True
        
'        TabsLoad
        
        'SA 1/29/2013 - CR3448 Removed call to BuildTotalsCustom to prevent multiple calls
        'BuildTotalsCustom False
    End If
Block_Exit:

    Exit Sub
Block_Err:
    MsgBox "Error applying filter", vbCritical, "Error"
    GoTo Block_Exit
End Sub

Private Sub cmdFiltersClear_Click()
    Me.cmboFiltersSelected.RowSource = vbNullString
End Sub

Private Sub cmdFiltersRemove_Click()
On Error GoTo Block_Err

    'SA 03/22/2012 - CR2667 Changed how filters are removed from the list
    If Me.cmboFiltersSelected.ListCount = 1 Then
        Me.cmboFiltersSelected.RemoveItem 0
    Else
        If Me.cmboFiltersSelected.ItemsSelected.Count = 0 Then
            MsgBox "Please select a filter to remove.", vbInformation
        Else
            Me.cmboFiltersSelected.RemoveItem Me.cmboFiltersSelected.ListIndex
        End If
    End If
    
    Me.CmboFilters = vbNullString
    
Block_Exit:

    Exit Sub
Block_Err:
    MsgBox Err.Description, vbCritical, "Error removing filter"
    GoTo Block_Exit
End Sub

Private Sub cmdFiltersSave_Click()
On Error GoTo Block_Err
    Call SaveScreenFilter(Me.CmboFilters)

Block_Exit:
    Me.CmboFilters.Requery
    Exit Sub

Block_Err:
    MsgBox Err.Description
    GoTo Block_Exit
End Sub

Private Sub cmdLayoutApply_Click()
On Error GoTo Block_Err
Dim stName As String
Dim lngId As Long
Dim db As DAO.Database
Dim rst As DAO.RecordSet
Dim SQL As String
Dim X As Long
Dim Y As Long
Dim TxtFld As Access.TextBox
Dim Msg As String

    DoCmd.Hourglass True
    Application.Echo False
    
    'Make sure there is a layout selected
    If CmboLayouts.ListIndex = -1 Then
        If MsgBox("Your Current Layout Will Be Cleared!", vbInformation + vbOKCancel, "Clear Layouts") = vbOK Then
'            Me.SubformCalcs.Form.ApplyClear
            If Not oMainGrid Is Nothing Then
'                MvGridMain.CalcFieldsClear 'CLEAR THE CALCS
'                Me.SubformCondFormats.Form.ApplyFormatClear 'Clear FORMAT  List
'                oMainGrid.FormatsClear
                oMainGrid.LayoutClear
            End If
        End If
        GoTo Block_Exit
    End If
    
    lngId = Me.CmboLayouts
    stName = Me.CmboLayouts.Column(1, CmboLayouts.ListIndex)
    Msg = "Applying Layout " & Chr(34) & stName & Chr(34) & vbCrLf
    
    Set db = CurrentDb
    
'    'REMOVE ALL OF THE EXISTING CALCS
'    If Not oMainGrid Is Nothing Then
'        oMainGrid.CalcFieldsClear
'    End If

    'REMOVE CALCS FROM APPLY LIST
'    Me.SubformCalcs.Form.ApplyClear
    'GET THE CALCULATED FIELDS
'    sql = "SELECT LC.*, C.CalcName " & _
'          "FROM CA_ScreenLayOutsCalculations AS LC " & _
'          "INNER JOIN CA_ScreensCalculations AS C ON LC.CalcID = C.CalcID " & _
'          "WHERE LC.LayoutID =" & CStr(LngID)
'    Set rst = db.OpenRecordSet(sql, dbOpenSnapshot)
'    With rst
'        If .EOF And .BOF Then
'            Msg = Msg & "Calculations: 0" & vbCrLf
'        Else
'            Do Until .EOF
'                Me.SubformCalcs.Form.ApplyAdd .Fields("CalcID"), .Fields("CalcName")
'                'SubformCalcs.Form
'                .MoveNext
'            Loop
'            .Close
'            Msg = Msg & "Calculations: " & Me.SubformCalcs.Form.ActiveCalcCount & vbCrLf
'            Me.SubformCalcs.Form.ApplyCalcs
'        End If
'    End With
'    Set rst = Nothing

'    'REMOVE ALL OF THE EXISTING FORMATS
'    Me.SubformCondFormats.Form.ApplyFormatClear 'Clear The List
'    Me.SubformCondFormats.Form.ApplyFormats 'Clear the data by applying
    
'    'GET THE CONDITIONAL FORMATS
'    sql = "SELECT LF.*, F.FormatName " & _
'          "FROM CA_ScreenLayOutsFormats AS LF " & _
'          "INNER JOIN CA_ScreensCondFormats AS F on " & _
'          "LF.CondFormatID = F.CondFormatID " & _
'          "WHERE LF.LayoutID =" & CStr(LngID)
'    Set rst = db.OpenRecordSet(sql, dbOpenSnapshot)
'    With rst
'        If .EOF And .BOF Then
'            Msg = Msg & "Formats: 0" & vbCrLf
'        Else
'            Do Until .EOF
'                Me.SubformCondFormats.Form.ApplyFormatAdd .Fields("CondFormatID"), .Fields("FormatName")
'                .MoveNext
'            Loop
'            .Close
'            Msg = Msg & "Formats: " & Me.SubformCondFormats.Form.ActiveFormatCount & vbCrLf
'            Me.SubformCondFormats.Form.ApplyFormats
'        End If
'    End With
'    Set rst = Nothing

    'CLEAR THE FIELD LAYOUT
    oMainGrid.LayoutClear
    'GET THE FIELD LAYOUTS FROM THE DATABASE
    SQL = "SELECT FieldName, Ordinal, ColWidth, CalcFld " & _
          "FROM CA_ScreenLayOutsFields AS LF " & _
          "WHERE LF.LayoutID =" & CStr(lngId) & " AND LF.Identifier='MainGrid' " & _
          "ORDER BY Ordinal"
    X = 0
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    With rst
        If .EOF And .BOF Then
            Msg = Msg & "Layouts: 0" & vbCrLf
        Else
            Do Until .EOF
                X = X + 1
                oMainGrid.LayoutField .Fields("FieldName"), .Fields("Ordinal"), .Fields("ColWidth"), .Fields("CalcFld")
                .MoveNext
            Loop
            Msg = Msg & "Layouts: " & .recordCount & vbCrLf
            .Close
        End If
    End With
    Set rst = Nothing
    
    Me.LayoutApplied = True
    
    DoEvents
    
'    ' HC 9/25/2008 - load the tab layouts
'    Set db = CurrentDb
'    Dim Found As Boolean
'    For x = 0 To mvConfig.TabsCT - 1
'        'GET THE FIELD LAYOUTS FROM THE DATABASE
'        Y = 0
'        sql = "SELECT Identifier, FieldName, Ordinal, ColWidth, CalcFld " & _
'              "FROM CA_ScreenLayOutsFields AS LF " & _
'              "WHERE LF.LayoutID=" & CStr(LngID) & " AND LF.Identifier='" & mvConfig.Tabs(x).TabID & "' " & _
'              "ORDER BY Ordinal"
'
'        Set rst = db.OpenRecordSet(sql, dbOpenSnapshot, dbReadOnly)
'
'        With rst
'            If rst.recordCount > 0 Then
'                Found = False
'                For Y = 0 To Tabs.Pages.Count - 1
'                    If Tabs.Pages(Y).Tag = rst.Fields("Identifier") Then
'                        Found = True
'                        Set holdForm = Nothing
'                        Set holdForm = Tabs.Pages(Y).Controls(0).Form
'
'                        'SA 11/15/2012 - Only apply layouts for tabs with datasheet
'                        If left(holdForm.Name, 22) = "CT_SubGenericDataSheet" Then
'                            holdForm.LayoutClear
'                        Else
'                            Found = False
'                        End If
'                        Exit For
'                    End If
'                Next Y
'                If Found Then
'                    Do Until .EOF
'                        Y = Y + 1
'                        holdForm.LayoutField .Fields("FieldName"), .Fields("Ordinal"), .Fields("ColWidth"), .Fields("CalcFld")
'                        .MoveNext
'                    Loop
'                End If
'                .Close
'            End If
'        End With
'        Set rst = Nothing
'    Next x

'    DoEvents
    
Block_Exit:

    Set db = Nothing
    Set rst = Nothing
    Set TxtFld = Nothing
    
    'CRAZY FORMAT CODE TO KEEP SCROLL BAR
    oMainGrid.Form.InsideWidth = Me.InsideWidth
    Form_Resize
    
    DoCmd.Hourglass False
    Application.Echo True
Exit Sub
Block_Err:
    MsgBox Err.Description & vbCrLf & vbCrLf & Msg, vbCritical, "Error Saving Layout"
    GoTo Block_Exit
End Sub

Private Sub cmdLayoutDelete_Click()
Dim lngId As Long
Dim db As DAO.Database
Dim stName As String
    Set db = CurrentDb
    
    ' Only allow them to delete their own
    
    'Take current layout as default name
    If CmboLayouts.ListIndex <> -1 Then
        stName = CmboLayouts.Column(1, CmboLayouts.ListIndex)
        lngId = CmboLayouts
    End If
    'Check if it exists
Stop
    lngId = Nz(DLookup("LayoutID", "CA_ScreenLayOuts", "ScreenID = 1 and LayoutName = " & Chr(34) & stName & Chr(34)) & " And UserName = """ & GetUserName & """", -1)
    If lngId <> -1 Then 'It exists - Ask then delete IT
        If MsgBox("Delete Layout '" & stName & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Delete") = vbYes Then
            db.Execute "DELETE FROM CA_ScreenLayOuts WHERE LayoutID=" & CStr(lngId), dbFailOnError + dbSeeChanges
        End If
    End If
    
    Me.CmboLayouts.Requery

End Sub

Private Sub cmdLayoutSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim stName As String
Dim lngId As Long
Dim db As DAO.Database
Dim SQL As String
Dim X As Long
Dim TxtFld As Access.TextBox
Dim Msg As String
Dim Y As Long
    
    strProcName = ClassName & ".cmdLayoutSave_Click"
    
    DoCmd.Hourglass True
    'Take current layout as default name
    If CmboLayouts.ListIndex <> -1 Then
        stName = CmboLayouts.Column(1, CmboLayouts.ListIndex)
        lngId = CmboLayouts
    End If
    
GetName:
    'Confirm the name
    stName = InputBox("Please enter the name of the new existing layout.", "Layout Name", stName)
    If vbNullString & stName = vbNullString Then
        GoTo Block_Exit
    End If
    
    Set db = CurrentDb
    'Check if it exists
    lngId = Nz(DLookup("LayoutID", "CA_ScreenLayOuts", "ScreenID = 1 and LayoutName = " & Chr(34) & stName & Chr(34) & " AND UserName = """ & GetUserName & """"), -1)
    If lngId <> -1 Then 'It exists - Ask then delete IT
        If MsgBox("The Layout '" & stName & "' already exists." & vbCrLf & vbCrLf & "Would you like to replace it?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Replace") = vbYes Then
            db.Execute "DELETE FROM CA_ScreenLayOuts WHERE LayoutID=" & CStr(lngId), dbFailOnError + dbSeeChanges
        Else
            GoTo GetName 'Try Getting a new name
        End If
    End If
    
    'Create The Layout Record
    SQL = "INSERT INTO CA_ScreenLayOuts(ScreenID, LayoutName, Computer, UserName) VALUES (" & _
        "1, " & _
        Chr(34) & stName & Chr(34) & ", " & _
        Chr(34) & Identity.Computer & Chr(34) & ", " & _
        Chr(34) & Identity.UserName & Chr(34) & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    lngId = Nz(DLookup("LayoutID", "CA_ScreenLayOuts", "ScreenID = 1 and LayoutName = " & Chr(34) & stName & Chr(34)), -1)
    Msg = "Layout " & Chr(34) & stName & Chr(34) & " Saved!" & vbCrLf
    
    'SAVE THE FORMATS
'    With SubformCondFormats.Form
'        If .ActiveFormatCount > 0 Then
'            'Create The FormatRecords
'            sql = "INSERT INTO CA_ScreenLayOutsFormats(LayoutID,CondFormatID) " & _
'                "SELECT " & CStr(LngID) & ", CondFormatID " & _
'                "FROM CA_ScreensCondFormats " & _
'                "WHERE CondFormatID IN (" & .SQLList & ")"
'            db.Execute sql, dbFailOnError + dbSeeChanges
'            Msg = Msg & "Formats: " & .ActiveFormatCount & vbCrLf
'        Else
'            Msg = Msg & "Formats: 0" & vbCrLf
'        End If
'    End With
    
    'SAVE THE CALCULATIONS
'    With Me.SubformCalcs.Form
'        If .ActiveCalcCount > 0 Then
'            'Create The FormatRecords
'            sql = "INSERT INTO CA_ScreenLayOutsCalculations(LayoutID,CalcID) " & _
'                "SELECT " & CStr(LngID) & ", CalcID " & _
'                "FROM CA_ScreensCalculations " & _
'                "WHERE CalcID in (" & .SQLList & ")"
'            db.Execute sql, dbFailOnError + dbSeeChanges
'            Msg = Msg & "Calculations: " & .ActiveCalcCount & vbCrLf
'        Else
'            Msg = Msg & "Calculations: 0" & vbCrLf
'        End If
'    End With
    
    ' HC 9/25/2008 changed mvgridmain to me.subform.form; added new identifier to screenlayoutfields
'    With Me.SubForm.Form    'MvGridMain
    With Me.frm_GENERAL_Datasheet.Form
        For X = 1 To .FldCT
            Set TxtFld = .Controls("Field" & CStr(X))
            SQL = "INSERT INTO CA_ScreenLayOutsFields(LayoutID,Identifier,FieldName,CalcFld,ColWidth,Ordinal)VALUES(" & CStr(lngId) & ", 'MainGrid' , "
            If vbNullString & TxtFld.Tag <> vbNullString Then 'Calculated field
                SQL = SQL & Chr(34) & TxtFld.Tag & Chr(34) & ", " & "True, "
            Else
                '021406 David.Brady added "EscapeQuotes" to accomodate control sources with quotes in them.
                SQL = SQL & Chr(34) & EscapeQuotes(TxtFld.ControlSource) & Chr(34) & ", " & "False, "
            End If
            SQL = SQL & IIf(TxtFld.ColumnHidden = True, 0, TxtFld.ColumnWidth) & ", " & TxtFld.ColumnOrder & vbNullString & ")"
            db.Execute SQL, dbFailOnError + dbSeeChanges
        Next X
        Msg = Msg & "Column Layouts: " & .FldCT & vbCrLf
    End With
    
    ' HC 9/24/2008 - build an xml string of the tab layout information
    
'    For x = 0 To mvConfig.TabsCT - 1
'        With Tabs.Pages(x + 1).Controls(0).Form
'            'SA 11/15/2012 - Only save layouts for datasheets
'            If left(.Name, 22) = "CT_SubGenericDataSheet" Then
'                For Y = 1 To .FldCT
'                    sql = "INSERT INTO CA_ScreenLayOutsFields(LayoutID,Identifier,FieldName,CalcFld,ColWidth,Ordinal)VALUES(" & _
'                        CStr(LngID) & ", " & Chr(34) & mvConfig.Tabs(x).TabID & Chr(34) & ", "
'                    Set TxtFld = .Controls("Field" & CStr(Y))
'                    If vbNullString & TxtFld.Tag <> vbNullString Then 'Calculated field
'                        sql = sql & Chr(34) & TxtFld.Tag & Chr(34) & ", " & "True, "
'                    Else
'                        '021406 David.Brady added "EscapeQuotes" to accomodate control sources with quotes in them.
'                        sql = sql & Chr(34) & EscapeQuotes(TxtFld.ControlSource) & Chr(34) & ", " & "False, "
'                    End If
'                    sql = sql & IIf(TxtFld.ColumnHidden = True, 0, TxtFld.ColumnWidth) & ", " & _
'                        TxtFld.ColumnOrder & vbNullString & ") "
'                    db.Execute sql, dbFailOnError + dbSeeChanges
'                Next Y
'            End If
'        End With ' with
'    Next x   ' next x
           
'    Msg = Msg & "Tab Layouts Layouts: " & mvConfig.TabsCT & vbCrLf
       
    MsgBox Msg, vbInformation, "Layout Saved"
Block_Exit:
    On Error Resume Next
    DoCmd.Hourglass False
    Me.CmboLayouts.Requery
    Set db = Nothing
    Set TxtFld = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdNew_Click()

    Me.subFrmMain.SourceObject = ""
    Set ofrmNewConcept = New Form_frm_CONCEPT_New_Concept
    ShowFormAndWait ofrmNewConcept
    
    Me.txtSearchBox = cstrNewConceptId
    
    Call cmdSearch_Click
    
End Sub

Private Sub cmdOldContract_Click()
    DoCmd.OpenForm "frm_CONCEPT_Main", acNormal
End Sub

Private Sub cmdRefresh_Click()
  Me.txtSearchBox = ""
    RefreshData
  
End Sub

Private Sub cmdSearch_Click()
    RefreshData
    
End Sub



Private Sub Command20_Click()

Debug.Print Me.lblTabs.top

Stop
End Sub

Private Sub cmdTstStuff_Click()
Dim oFrm As Form
Dim oCtl As Control

Set oFrm = Me.subFrmMain.Form

oFrm.InsideHeight = 7800
oFrm.InsideWidth = 14000
oFrm.SelHeight = 1000
oFrm.SelWidth = 1000
oFrm.SplitFormDatasheet = acDatasheetReadOnly
'oFrm.SplitFormSize = 500
'oFrm.WindowHeight = 700
'oFrm.WindowWidth = 700
'Me.subFrmMain.SizeToFit

Set oCtl = oFrm.Controls(0)

Stop

End Sub



Private Sub Command23_Click()
Stop
        Me.subFrmMain.SourceObject = "frm_CONCEPT_Validation"
        Me.subFrmMain.Form.IdValue = Me.txtConceptID
    '                        Set Me.subFrmMain.Form.Recordset = mrsConcept
    '                        Set Me.subFrmMain.Form.ConceptRecordSource = mrsConcept
    '                        Me.subFrmMain.Form.RefreshData
    '                        Call SetSubFormPayerSel
        Me.subFrmMain.Form.RefreshData
End Sub



Private Sub cmdToggle_Click()
    SetTogValue = Not GetTogValue
    PopulateScreenFilters Me.CmboFilters, 1
End Sub

Private Sub Command65_Click()
Stop
End Sub



Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    screen.MousePointer = 0
End Sub

Private Sub Form_Current()
Debug.Print Me.Name & ".Form_Current"
        screen.MousePointer = 0
End Sub

Private Sub Form_Deactivate()
    screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim iSetting As Integer
Dim iAppPermission As Integer
    
    gbTogValue = True
    iSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    On Error Resume Next
    giHdrFormSelectedPage = 0

    Me.frm_GENERAL_Datasheet.Form.RecordSource = ""
    Me.subFrmMain.Form.RecordSource = ""

    Me.Caption = "Concept Maintenance"
    Me.Detail.AutoHeight = True
    
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
        lstTabs.RowSource = GetListBoxSQL(Me.Name)
        If lstTabs.ListCount > 1 Then
            Me.lstTabs = Me.lstTabs.ItemData(0)
        End If
        
    Else
        MsgBox "You do not have permission to view this form.  Please contact your system admin", vbInformation
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
  
    Call PopulateListLayouts(Me.CmboLayouts, IIf(Me.ckSharedLayouts.Value = 0, False, True))
    Call PopulateScreenFilters(Me.CmboFilters)
    
    On Error GoTo 0
    Application.SetOption "Error Trapping", iSetting
'    If Me.frm_GENERAL_Datasheet.Form.CompletedLoad = True Then
        RefreshData
'    End If

    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error GoTo Block_Err
Dim strProcName As String

'    ResizeControls Me.Form
Dim dblHeight As Double
Dim sglGridPadding As Single
Dim sglGridHeight As Single
Dim sglGridWidth As Single
Dim sglTabHeight As Single

    strProcName = ClassName & ".Form_Resize"
    

    ''' I'm guessing that Decipher only sets the ScreenId after it's loaded
'    If Me.ScreenID <> 0 Then
    'suppress the suspend layout when creating a new screen for faster loading
    If Me.visible = True Then
        genUtils.SuspendLayout
    End If
'    End If

    ' changed by SC Since the form footer property cangrow = true the following will prevent the footer from
    ' growing larger than the main grid which will cause a link break down between the main grid and the tabs
    ' HC 4/24/2008 - only change size if form footer is visible
    
    '' Ok, TabsHead = height of the form header
    '' lblBanner is in my case the top of the main subform (but it's 270 in height in Decipher so I'm hard codeing it
    '' until I figure out why not just the top of the Detail section...
            '
            '    If FormFooter.visible And (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) > 0 And _
            '         FormFooter.Height >= (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) Then
            '            FormFooter.Height = Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))
            '    End If
            '
    If Me.FormFooter.visible And (Me.InsideHeight - (Me.lblAppTitle.Height + (270 * 2))) > 0 And _
         Me.FormFooter.Height >= (Me.InsideHeight - (Me.lblAppTitle.Height + (270 * 2))) Then
            Me.FormFooter.Height = Me.InsideHeight - (Me.lblAppTitle.Height + (270 * 2))
    End If


    '' KD My own, limit to smallest size..
    If Me.InsideHeight < 11760 Then
        Me.InsideHeight = 11760
    End If
    If Me.InsideWidth < 16620 Then
        Me.InsideWidth = 16620
    End If


        '            Me.lblBanner.top = 0
        '            Me.SubForm.top = lblBanner.Height
    Me.frm_GENERAL_Datasheet.top = 0
    Me.subFrmMain.top = lblTabs.Height
    Me.lstTabs.top = Me.subFrmMain.top
    
        '            GridPadding = Me.SubForm.left * 2
    sglGridPadding = Me.lstTabs.left * 2
    
        '            If Me.FormFooter.visible Then
        '                GridHeight = Me.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height + lblBanner.Height)
        '            Else
        '                GridHeight = Me.InsideHeight - (Me.FormHeader.Height + lblBanner.Height)
        '            End If
    If Me.FormFooter.visible = True Then
'        sglGridHeight = Me.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height + 270)
        sglGridHeight = Me.InsideHeight - (Me.FormHeader.Height + Me.FormFooter.Height)
    Else
'        sglGridHeight = Me.InsideHeight - (Me.FormHeader.Height + 270)
        sglGridHeight = Me.InsideHeight - (Me.FormHeader.Height)
    End If
    
    

                '    If GridHeight > 0 And (Me.WindowHeight > GridHeight) Then
                '        Me.Detail.Height = GridHeight + lblBanner.Height
                '        Me.SubForm.Height = GridHeight
                '    End If
    If sglGridHeight > 0 And (Me.WindowHeight > sglGridHeight) Then
'        Me.Detail.Height = sglGridHeight + 270
        Me.Detail.Height = sglGridHeight
        'Me.subFrmMain.Height = sglGridHeight
        Me.frm_GENERAL_Datasheet.Height = sglGridHeight
    End If
    
                '    GridWidth = Me.InsideWidth
                '    If GridWidth > 0 Then
                '        Me.SubForm.Width = GridWidth
                '        lblBanner.Width = GridWidth
                '        Me.TabsHead.Width = GridWidth
                '        Me.Splitter.top = 0
                '        Me.Splitter.Width = GridWidth
                '        Me.Tabs.Width = GridWidth
                '    End If
    sglGridWidth = Me.InsideWidth
    If sglGridWidth > 0 Then
        Me.frm_GENERAL_Datasheet.Width = sglGridWidth
'        lblBanner.Width = sglGridWidth
        Me.lblAppTitle.Width = sglGridWidth
        Me.lblTabs.top = 0
        Me.lblTabs.Width = sglGridWidth - Me.lblTabs.left
        Me.subFrmMain.Width = sglGridWidth - Me.subFrmMain.left
    End If


                    '''    If Me.Tabs.Value > -1 Then
    If Me.FormFooter.visible = True Then
    
                    '''        txtFocus.SetFocus
        Me.txtDecoy.SetFocus
                    '''        Me.Tabs.visible = True
'        Me.subFrmMain.visible = True
'        Me.lstTabs.visible = True
        
                    '''        Me.Splitter.visible = True
        Me.lblTabs.visible = True
        
                    '''        With Me.Tabs
                    
                    '''            .left = 0
        Me.subFrmMain.visible = False
        Me.lstTabs.visible = False
        
                    '''            .visible = False
                    '''            TabHeight = Me.FormFooter.Height - Me.Splitter.Height
        sglTabHeight = Me.FormFooter.Height - Me.lblTabs.Height
        
                    '''            'SA 1/10/2012 - Changed resize code to go through all visible bottom tabs instead of using mvConfig.TabsCT
                    '''            For i = 0 To Tabs.Pages.Count - 1
                    '''                With Tabs.Pages(i)
                    '''                    If .visible Or i = 0 Then
                    '''                        .left = Me.Tabs.left
                    '''                        If GridWidth > 0 Then
                    '''                            .Width = GridWidth - (GridPadding * 3)
                    '''                        End If
                    '''                        With .Controls(0)
                    '''                            .top = Tabs.Pages(i).top
                    '''                            If GridWidth > 0 Then
                    '''                                .Width = GridWidth - (GridPadding * 4)
                    '''                            End If
                    '''                            .Height = Tabs.Pages(i).Height - GridPadding
                    '''                            .left = Tabs.Pages(i).left
                    '''                        End With
                    '''                    End If
                    '''                End With
                    '''            Next
                    '''            If GridWidth > 0 Then
                    '''                .Width = GridWidth - (GridPadding * 1)
                    '''            End If
        If sglGridWidth > 0 Then
            Me.subFrmMain.Width = sglGridWidth - (sglGridPadding * 1) - Me.subFrmMain.left
            ' I am not resizing the tabs width
'            Me.lstTabs.Width = sglGridWidth - (sglGridPadding * 1) - Me.lstTabs.left
        End If
                    '''            If screen.MousePointer <> 7 Then
                    '''                .visible = True
                    '''            End If
                    '''        End With
        If screen.MousePointer <> 7 Then
            Me.subFrmMain.visible = True
            Me.lblTabs.visible = True
            Me.lstTabs.visible = True
        End If
                    '''    Else
                    '''        txtFocus.SetFocus
                    '''        Me.Tabs.visible = False
                    '''        Me.Splitter.visible = False
                    '''        Me.FormFooter.Height = 0
                    '''    End If
                    '''
    Else
        txtDecoy.SetFocus
        Me.subFrmMain.visible = False
        Me.lstTabs.visible = False
        Me.lblTabs.visible = False
        Me.FormFooter.Height = 0
    End If

    Me.cmdOldContract.left = Me.InsideWidth - sglGridPadding - Me.cmdOldContract.Width

'' Below is KD's original Resize code:
'    Me.frm_GENERAL_Datasheet.Width = Me.InsideWidth - 275
''    Me.frm_GENERAL_Datasheet.Form.Width = Me.Width - 700
'
'    dblHeight = Me.InsideHeight - Me.subFrmMain.Height - Me.lblTabs.Height + 800  '(Me.lblTabs.Height * 2)
'    If dblHeight < 4155 Then
'        dblHeight = 4155
'    End If
''    Me.frm_GENERAL_Datasheet.Height = Me.Detail.Height - Me.subFrmMain.Height - 200 - Me.lblTabs.Height  '(Me.lblTabs.Height * 2)
'    Me.frm_GENERAL_Datasheet.Height = dblHeight  '(Me.lblTabs.Height * 2)
'
''    Me.frm_GENERAL_Datasheet.Form.Refresh
'    Me.frm_GENERAL_Datasheet.Form.Repaint
'
'    ' me.subFrmMain.Height
'    'Me.Label41.top = Me.frm_GENERAL_Datasheet.top + Me.subFrmMain.Height '- 1400
'    Me.Label41.top = Me.frm_GENERAL_Datasheet.top + Me.frm_GENERAL_Datasheet.Height
'
'    Me.lblTabs.top = Me.Label41.top
'
'
'
'    Me.lblAppTitle.Width = Me.InsideWidth
'    Me.lblTabs.Width = Me.InsideWidth
'
'Dim oSection As Section
'Set oSection = Me.subFrmMain.Form.Section(0)
'
'
'    Me.subFrmMain.Width = Me.InsideWidth - 275
'    Me.subFrmMain.Height = Me.InsideHeight - 275
'    Me.subFrmMain.Form.Width = Me.Width - 700
'    Me.subFrmMain.Form.InsideHeight = Me.subFrmMain.Height - 700
'
'    Me.subFrmMain.top = Me.lblTabs.top + Me.lblTabs.Height + 20
'    Me.lstTabs.top = Me.Label41.top + Me.Label41.Height + 40
'
'    Me.subFrmMain.Form.Repaint
    
    If Me.visible = True Then
    'suppress the suspend layout when creating a new screen for faster loading
         genUtils.ResumeLayout
         If Me.subFrmMain.visible = False Then
            Me.subFrmMain.visible = True
         End If
         If Me.lstTabs.visible = False Then
            Me.lstTabs.visible = True
         End If
    End If
   
   
   
Block_Exit:
    ' insure the "links" didn't get lost
    Set oMainGrid = Me.frm_GENERAL_Datasheet.Form
    If TypeName(Me.subFrmMain.Form) = "Form_frm_CONCEPT_Hdr" Then
        Set frmConceptHdr = Me.subFrmMain.Form
    End If
    
    ' this is for when I apply the Decipher saved filter features
'    If Me.SubForm.Form.filter <> MvFilter Then
'        Me.SubForm.Form.filter = MvFilter
'    End If
    
    Me.Repaint
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Unload(Cancel As Integer)
Debug.Print Me.Name & ".Form_Unload"
    screen.MousePointer = 0
End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub frmGeneralNotes_NoteAdded()
    If SaveData_Notes Then
        MsgBox "Note added"
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





Private Sub lblAppTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub lblTabs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        csgSplitter = Y
        PrepareSplitterResize
    End If
End Sub

Private Sub lblTabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        screen.MousePointer = 7
    End If
    
    If Button = 1 And Y <> 0 Then
        screen.MousePointer = 0

        ' TabsHead is used for the Form Header height
        ' lblBanner is the separator between the form Header and the main grid in the form detail
'        If Me.FormFooter.Height + (csgSplitter + (Y * -1)) < (Me.InsideHeight - (Me.TabsHead.Height + (lblBanner.Height * 2))) Then
'                Me.FormFooter.Height = Me.FormFooter.Height + (MvSplitY + (Y * -1))
'        End If
        If Me.FormFooter.Height + (csgSplitter + (Y * -1)) < (Me.InsideHeight - (Me.lblAppTitle.Height + (270 * 2))) Then
                Me.FormFooter.Height = Me.FormFooter.Height + (csgSplitter + (Y * -1))
        End If

    End If
End Sub

Private Sub lblTabs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        screen.MousePointer = 0
'        Me.Splitter.BorderColor = Splitter.BackColor
        Call ReleaseFooterForResize
        screen.MousePointer = 0
    End If
End Sub

Private Sub lstTabs_Click()
'Tues 2/5/2013 by KCF - Add code to handle new module for DME Documentation Review \ Rationale template
On Error GoTo ErrHandler
    
    Dim strSQL As String
    Dim strFormValue As String
    Dim strSQLCharacter As String
    Dim strSQLValue As String
    Dim strFormName As String
    Dim lngNoteID As Long
    
'    Set MYADO = New clsADO
    If Me.Searching = True Then
        GoTo Block_Exit
    End If


    If Me.lstTabs.ListIndex <> -1 Then
        Dim rs As DAO.RecordSet
            'Get a recordset of tabs for this form
            Set rs = CurrentDb.OpenRecordSet(GetListBoxRowSQL(lstTabs.Column(1), Me.Name), dbOpenSnapshot, dbSeeChanges)
            If Not (rs.BOF And rs.EOF) Then
            
                If rs("FormName") <> "frm_Concept_Hdr" Then
                    Set frmConceptHdr = Nothing
                End If
            
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

                        If Not frmConceptHdr Is Nothing Then
                            Me.subFrmMain.Form.CurrentPageSelected = giHdrFormSelectedPage
                        End If
                        If frmConceptHdr Is Nothing Then
                            Set frmConceptHdr = Me.subFrmMain.Form
                        End If
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
                    Case "frm_CONCEPT_AddPayer", "frm_CONCEPT_LCD"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        Me.subFrmMain.Form.FormConceptID = Me.txtConceptID
                        Me.lblTabs.Caption = Me.lstTabs
                        
                        If RefreshSubform(rs("FormName")) = True Then
                            Me.subFrmMain.Form.RefreshData
                        End If
                        
                    Case "frm_CONCEPT_Validation"
                        Me.subFrmMain.SourceObject = rs("FormName")
                        
                        Me.subFrmMain.Form.IdValue = Me.txtConceptID
                        Me.subFrmMain.Form.FormConceptID = Me.txtConceptID
                        'Me.lblTabs.Caption = Me.lstTabs
'
'                        If RefreshSubform(rs("FormName")) = True Then
''                            Me.subFrmMain.Form.RefreshData
'                        End If
                        
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
    
    With Me.subFrmMain.Form
        .AutoResize = True
        .ScrollBars = 3
        .InsideHeight = Me.subFrmMain.Height
        .InsideWidth = Me.subFrmMain.Width
'        .Recordset.MoveLast
'        .Recordset.MoveFirst
        
    End With
    
Block_Exit:
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

Private Sub lstTabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub ofrmNewConcept_ConceptSaved(strNewConceptId As String)
    cstrNewConceptId = strNewConceptId
End Sub

Private Sub oMainGrid_Current()
On Error GoTo Block_Err
Dim strProcName As String
Dim sNewConcept As String


    strProcName = ClassName & ".oMainGrid_Current"

    Set MyAdo = New clsADO

    If oMainGrid Is Nothing Then
        GoTo Block_Exit
    End If
    If oMainGrid.RecordSource = "" Then
'        Stop
        GoTo Block_Exit
    End If

    sNewConcept = Nz(oMainGrid.Controls("ConceptId"), "")

    
    If Me.txtConceptID = sNewConcept Then
        ' same concept - nothing to do
        ' well, let's check the frmConceptHdr
        If Not frmConceptHdr Is Nothing Then
            If Me.subFrmMain.Form.Controls("txtSelectedId") = sNewConcept Then
                ' Ok, NOW, nothing to do
                GoTo Block_Exit
            End If
        End If
        'GoTo Block_Exit
    End If

    Me.txtConceptID = Nz(oMainGrid.Controls("ConceptID"), "")
    Me.txtNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
    mNoteID = Nz(oMainGrid.Controls("NoteID"), -1)
    'Refresh the tabs to ensure the main form is in sync with the other forms.
    
    MyAdo.ConnectionString = GetConnectString("v_Data_Database")
    
    MyAdo.sqlString = " SELECT * from CONCEPT_Hdr WHERE ConceptID = '" & Me.txtConceptID & "'"
    Set mrsConcept = MyAdo.OpenRecordSet()
        
    MyAdo.sqlString = " SELECT * from Note_Detail WHERE NoteID = '" & Me.txtNoteID & "'"
    Set mrsNotes = MyAdo.OpenRecordSet()

    ' if it's the same concept, no need to "click" the tab
    If Not frmConceptHdr Is Nothing Then
        If Me.txtConceptID <> frmConceptHdr.FormConceptID Then
            lstTabs_Click
        Else
'            Stop
            lstTabs_Click
        End If
    Else
        lstTabs_Click
    End If
    
    ' HC highlight the current row
    If Me.chkHighlight Then
        SendKeys "+ ", True
    End If
Block_Exit:
    
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
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



Private Sub oMainGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If screen.MousePointer = 7 Then screen.MousePointer = 0
End Sub

Private Sub txtSearchBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        ' don't care if it's shift, alt, or ctrl
        RefreshData
    End If
End Sub



Private Sub PrepareSplitterResize()
On Error GoTo Block_Err
Dim strProcName As String
Dim pgCt As Integer
Dim pgIDx As Integer
Dim pg As Page
Dim pgCtrlCt As Integer
Dim pgCtrlIdx As Integer
Dim pgCtrl As Control
    
    strProcName = ClassName & ".PrepareSplitterResize"
    
    genUtils.SuspendLayout Me
    
    screen.MousePointer = 7
'Me.lblTabs.BorderColor = 0
    Me.txtDecoy.SetFocus
    ' make stuff in the footer hidden as that's the section that's going to "take the brunt of the resize"
'    Me.Tabs.Height = 1
'    Me.Tabs.visible = False

    ' This is the decipher Main grid so it's our main grid..
'    Me.SubForm.visible = False

    Me.frm_GENERAL_Datasheet.visible = False
    Me.subFrmMain.visible = False
    Me.lstTabs.visible = False

    ' this part makes all of the pages in the footer's single tabs control tiny and invisible
    Call MakeFormFooterSmallNHidden
    
    '    pgCt = Me.Tabs.Pages.Count - 1
    '
    '    For pgIDx = 0 To pgCt
    '        Set pg = Me.Tabs.Pages.Item(pgIDx)
    '
    '        pgCtrlCt = pg.Controls.Count - 1
    '
    '        For pgCtrlIdx = 0 To pgCtrlCt
    '            Set pgCtrl = pg.Controls(pgCtrlIdx)
    '            With pgCtrl
    '                .Height = 1
    '                .Width = 1
    '                .visible = False
    '            End With
    '        Next
    '
    '        pg.Height = 1
    '    Next
    '
    '    Me.Tabs.Height = 1
    
    genUtils.ResumeLayout Me
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub MakeFormFooterSmallNHidden()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".MakeFormFooterSmallNHidden"
    
    Me.lstTabs.Height = 1
    
    Me.lstTabs.visible = False
    
    Me.subFrmMain.visible = False
    Me.subFrmMain.Height = 1
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub ReleaseFooterForResize()
On Error GoTo Block_Err
Dim strProcName As String
Dim pgCt As Integer
Dim pgIDx As Integer
Dim pg As Page
Dim pgCtrlCt As Integer
Dim pgCtrlIdx As Integer
Dim pgCtrl As Control
Dim TabHeight As Integer
    
    strProcName = ClassName & ".ReleaseFooterForResize"
    
'    ''' Decipher:
'
'    myForm.Splitter.top = 0
'    myForm.Tabs.top = myForm.Splitter.top + myForm.Splitter.Height
'    myForm.Tabs.Height = myForm.FormFooter.Height - myForm.Tabs.top
'

    Me.lblTabs.top = 0
    Me.subFrmMain.top = Me.lblTabs.top + Me.lblTabs.Height
    Me.subFrmMain.Height = Me.FormFooter.Height - Me.subFrmMain.top
    
    '' Have to do the list box too:
    Me.lstTabs.top = Me.subFrmMain.top
    Me.lstTabs.Height = Me.FormFooter.Height - Me.lstTabs.top
    

'    TabHeight = myForm.Tabs.Height - 500
'    pgCt = myForm.Tabs.Pages.Count - 1
'
'    Application.Echo False
'    For pgIDx = 0 To pgCt
'        Set pg = myForm.Tabs.Pages.Item(pgIDx)
'        pg.Height = TabHeight
'        pgCtrlCt = pg.Controls.Count - 1
'        For pgCtrlIdx = 0 To pgCtrlCt
'            Set pgCtrl = pg.Controls(pgCtrlIdx)
'
'            With pgCtrl
'                If pg.Height - 500 > 0 Then
'                    .Height = pg.Height - 500
'                End If
'                .Width = pg.Width
'                .visible = True
'            End With
'        Next
'    Next
'
'    myForm.Tabs.visible = True
'    myForm.SubForm.visible = True
    Me.frm_GENERAL_Datasheet.visible = True
    Me.subFrmMain.visible = True
    Me.lstTabs.visible = True

'    myForm.Resize
    Call Form_Resize
    

Block_Exit:
    
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub PopulateListLayouts(oCtl As Control, Optional bAllUsers As Boolean = False)
Dim strProcName As String
On Error GoTo Block_Err
Dim SQL As String
    strProcName = ClassName & ".PopulateListLayouts"
    
'    sql = "Select DISTINCT '' as LayoutID, '' as LayoutName FROM SCR_Screens Union ALL "
    If bAllUsers = True Then
        SQL = SQL & "SELECT LayoutID, LayoutName, UserName "
    Else
        SQL = SQL & "SELECT LayoutID, LayoutName "
    End If
    SQL = SQL & "FROM CA_ScreenLayOuts "
    If bAllUsers = False Then
        SQL = SQL & " WHERE UserName = """ & GetUserName & """ "
        SQL = SQL & " ORDER BY UserName, LayoutName;"
    Else
        SQL = SQL & " ORDER BY LayoutName;"
    End If

    
    oCtl.RowSource = SQL
    oCtl.Requery
    
Block_Exit:
    On Error Resume Next
    Exit Sub
Block_Err:
'    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Public Function GetName(ByVal Request As String, ByVal originalName As String) As String
'Get filter name from input box
On Error GoTo Block_Err
Dim FilterName As String
Dim sSuggested As String
Dim sOrigRequest As String

    sSuggested = originalName
    sOrigRequest = Request
    
Again:
    sSuggested = RemoveMeta(sSuggested)

    
    FilterName = InputBox(Request, "Save Filter", Replace(sSuggested, "'", vbNullString))
    FilterName = Replace(FilterName, "'", "''")
    FilterName = EscapeQuotes(FilterName)
    
    If InStr(1, FilterName, "[") > 0 Then
        sSuggested = FilterName
        Request = "Please try again: " & Request
        GoTo Again
    End If

    If InStr(1, FilterName, "(") > 0 Then
        Request = "Please try again: " & Request
        sSuggested = FilterName
        GoTo Again
    End If
    
Block_Exit:
    GetName = FilterName
    Exit Function
Block_Err:
    FilterName = vbNullString
    MsgBox Err.Description, vbCritical, "SCR_ClsMainScreens:GetName"
    GoTo Block_Exit
End Function


Private Sub SaveScreenFilter(ByRef oCtl As Control)
On Error GoTo Block_Err
Dim strProcName As String
Dim sFilterName As String
Dim tmpFilter As String
Dim FilterID As Long
Dim filterString As String
Dim frm As Form_frm_CONCEPT_Main
Dim SQL As String
Dim continue As Boolean
Dim Response As VbMsgBoxResult
    
    strProcName = ClassName & ".SaveScreenFilter"
    
    Set frm = oCtl.Parent.Form
    tmpFilter = Me.frm_GENERAL_Datasheet.Form.filter
    
    ' make sure there really is a filter to save
    If tmpFilter = vbNullString Then
        GoTo Block_Exit
    End If
        
    ' HC 11/8/2010 - 2010 changed to handler new format on For Name
    ' modified the replacement string for the form name to include the brackets; this seems to be a change in the way 2010 references object names.
    tmpFilter = Replace(tmpFilter, "[" & Me.frm_GENERAL_Datasheet.Form.Name & "].", vbNullString)
    filterString = Replace(tmpFilter, "'", "''")
    filterString = EscapeQuotes(filterString)
    filterString = Replace(filterString, Chr(34) & Chr(34), "'")
    
    sFilterName = vbNullString
    sFilterName = GetName("Please enter the filter name." & vbCrLf & vbCrLf, tmpFilter)
    If sFilterName = vbNullString Then
        GoTo Block_Exit
    End If
    
    continue = True
    While continue
        FilterID = Nz(DLookup("FilterId", "CA_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
            "ScreenId = " & Me.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34)), -1)
        If FilterID <> -1 Then
            Response = MsgBox("Filter: " & sFilterName & "  already exists!" & vbCrLf & vbCrLf & _
                "Do you want to overwrite it?", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Replace Filter")
            If Response = vbCancel Then
                continue = False
            ElseIf Response = vbYes Then
                continue = False
                SQL = " UPDATE CA_ScreensFilters SET filterName = " & Chr(34) & sFilterName & Chr(34) & _
                    ", filterSQL = " & Chr(34) & filterString & Chr(34) & _
                    " WHERE filterId = " & FilterID
                RunDAO SQL
                SQL = "DELETE FROM CA_ScreensFiltersDetails WHERE FilterId = " & FilterID
                RunDAO SQL
                SQL = " INSERT INTO CA_ScreensFiltersDetails(FilterId,Operator,SqlString) VALUES(" & FilterID & "," & _
                    Chr(34) & "CUSTOM" & Chr(34) & "," & Chr(34) & filterString & Chr(34) & ")"
                RunDAO SQL
            Else
                sFilterName = GetName("Please enter a different name for the filter." & vbCrLf & vbCrLf & _
                    tmpFilter & vbCrLf & vbCrLf & sFilterName & " is already taken!", tmpFilter)
                If sFilterName = vbNullString Then
                    continue = False
                End If
            End If
        Else
            continue = False
            SQL = " INSERT INTO CA_ScreensFilters(ScreenID,FilterName,FilterSQL,UserName) " & _
                    " VALUES (" & Me.ScreenID & "," & Chr(34) & sFilterName & Chr(34) & "," & _
                     Chr(34) & filterString & Chr(34) & "," & Chr(34) & Identity.UserName & Chr(34) & ")"
            RunDAO SQL
            FilterID = DLookup("FilterId", "CA_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
                "ScreenId = " & Me.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34))
            ' first remove all the items in the table for this filter
            SQL = "DELETE FROM CA_ScreensFiltersDetails WHERE FilterId = " & FilterID
            RunDAO SQL
            SQL = " INSERT INTO CA_ScreensFiltersDetails(FilterId,Operator,SqlString) VALUES(" & FilterID & "," & _
                Chr(34) & "CUSTOM" & Chr(34) & "," & Chr(34) & filterString & Chr(34) & ")"
            RunDAO SQL
        End If
    Wend

Block_Exit:
    Set frm = Nothing
    Exit Sub
Block_Err:
    If Err.Number = 3022 Then 'Duplicate Query String
        MsgBox "The same filter already exists under a different name!" & String(2, vbCrLf) & "Error saving filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    Else
        MsgBox Err.Description & String(2, vbCrLf) & "Error saving filter:" & String(2, vbCrLf) & tmpFilter, vbInformation, "Control Fucntions"
    End If
    GoTo Block_Exit
    
End Sub

Private Sub PopulateScreenFilters(oCtl As Control, Optional ByVal lScreenId As Long = 1)
On Error GoTo Block_Err
Dim strProcName As String
Dim SQL As String

    
    strProcName = ClassName & ".PopulateScreenFilters"
    
    
    SQL = "SELECT FilterId, FilterName, UserName " & "FROM CA_ScreensFilters " & "WHERE ScreenID = " & ScreenID & " "
        
    If gbTogValue Then
        Me.cmdToggle.Caption = "Mine"
        ' HC 6/2010 - changed filter for mine to display those items matching the user name and those w/o user name
        ' DS 11/15/11 - changed from UserName Is NULL to Nz(UserName, '') = '' to display imports screen filters with user name as empty string
        SQL = SQL & " AND (UserName ='" & Identity.UserName & "' or Nz(UserName, '') = '')"
    Else
        Me.cmdToggle.Caption = "All"
    End If
    
    'PD,Oct 22 2011 - CR # 2574 fix - Added clause to prevent filters with blank criteria from showing up on the main screen
    'SQL = SQL & " AND FilterSQL <> '' ORDER BY FilterName;"
    'SA 10/1/2012 - Changed to IS NOT NULL for SQL Server
    SQL = SQL & " AND FilterSQL IS NOT NULL ORDER BY FilterName;"
    
    oCtl.RowSource = SQL

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Function RunDAO(ByVal SQL As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim Result As Boolean
Dim db As DAO.Database
    
    
    strProcName = ClassName & ".RunDAO"
    
    Set db = CurrentDb
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    Result = True
Block_Exit:

    Set db = Nothing
    RunDAO = Result
    
    Exit Function
Block_Err:
    GoTo Block_Exit
    Resume
End Function
