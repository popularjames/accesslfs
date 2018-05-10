Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AuditClm_DocReview_DME
'
' Description: Provides a template for the DME team to review DME documentation, which can then be utilized to generate the Rationale for the claim.
' a) Pulls template setup from the dbo.AuditClm_DocReview_DME_Template tables (dtl & hdr);
' b) Updates saved records in dbo.AuditClm_DocReview_DME_Saved
'
' Developed:  December 2012 by Kathleen C Flanagan
'
'
' =============================================

Private strCnlyClaimNum As String
Private mstrUserName As String
Const CstrFrmAppID As String = "DocReviewDME"

Private rsAuditClmHdr As ADODB.RecordSet

Property Set HdrRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmHdr = data
End Property

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property

Public Sub DocReview_DME_ReviewSetup()
'===================================================================
'Called when the user selects the 'Doc Review - DME' option in the listbox on CA AuditClm_Main form
'Will call the usp_AuditClm_DocReview_DME_ReviewSetup; will setup the display for the form based upon the returned values
'If there are review guidelines for the concept will call the sub ReviewGuidelines_Popualate and display either the Overview tab or the Review tab depending upon the usp returned values
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================
    
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    mstrUserName = GetUserName()
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_AuditClm_DocReview_DME_ReviewSetup"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = Me.txtCnlyClaimNum
    cmd.Parameters("@pMasterUser") = mstrUserName
    cmd.Execute
    
    'strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_AuditClm_DocReview_DME_ReviewSetup"
        Err.Raise 50001, "usp_AuditClm_DocReview_DME_ReviewSetup", strErrMsg
    End If
    
    Me.txtRevGuide = (cmd.Parameters("@pReviewTextHdr"))
    Me.txtOverview = (cmd.Parameters("@pOverview"))
    Me.txtOverview.Height = (cmd.Parameters("@pOverviewHeight"))
    Me.lblTxtInstructions.top = 780 + cmd.Parameters("@pOverviewHeight") + 100
    Me.txtInstructions.top = Me.lblTxtInstructions.top + 240
    Me.txtInstructions = (cmd.Parameters("@pInstructions"))
    Me.txtInstructions.Height = (cmd.Parameters("@pInstructionsHeight"))
    
    If cmd.Parameters("@pReviewGuideStatus") = 0 Then 'If there is no template for the Adj_ConceptID
            Me.Rationale.SetFocus 'Display only the Rationale input page
            Me.Overview.visible = False
            Me.ReviewGuidelines.visible = False
            Me.Determination.visible = False
    ElseIf cmd.Parameters("@pReviewGuideStatus") = 1 Then 'If there is a template, but the stp had to create records for the AuditClm_DocReview_DME_Saved table for the Claim
        ReviewGuidelines_Populate
            Me.Overview.visible = True
            Me.ReviewGuidelines.visible = True
            Me.Determination.visible = False
            Me.Overview.SetFocus 'Initial display for a new review will be the Overview tab
    ElseIf cmd.Parameters("@pReviewGuideStatus") = 2 Then 'If there are records already saved to the AuditClm_DocReview_DME_Saved table
        ReviewGuidelines_Populate
            Me.Overview.visible = True
            Me.ReviewGuidelines.visible = True
            Me.Determination.visible = False
            Me.ReviewGuidelines.SetFocus 'If already a review setup, go directly to the Review Guidelines page
    End If

Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    
End Sub

Public Sub ReviewGuidelines_Populate()
'===================================================================
'Will populate & format the 'Review Guidelines' page of the form with the table values
'Called as part of the DocRevew_DME_Review_Setup subroutine which is called when the user selects the 'Doc Review - DME' list option the main AuditClm form
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================
Dim strSQL As String
Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Dim intI As Integer
Dim lngPreviousTop As Long

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = " SELECT * from AuditClm_DocReview_DME_saved where cnlyclaimnum = '" & Me.txtCnlyClaimNum & "' Order by ControlID"
    Set rs = MyAdo.OpenRecordSet(strSQL)
     
    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No Template Information for this Claim "
        Exit Sub
    End If
    
    lngPreviousTop = 500
    
'Set the page to expand to fit the data
    Me.Controls("ReviewGuidelines").Height = 20000
    
    For intI = 1 To rs.recordCount
      'For each record, set the value on the page & format the display
      
      Me.Controls("chk" & Trim(str(rs!ControlID))).visible = rs!ChkVisible
      Me.Controls("chk" & Trim(str(rs!ControlID))).Enabled = rs!chkEnabled
      Me.Controls("chk" & Trim(str(rs!ControlID))).left = rs!ChkLeft
      Me.Controls("chk" & Trim(str(rs!ControlID))) = rs!CtlValue  'new line
      Me.Controls("chk" & Trim(str(rs!ControlID))).top = lngPreviousTop
      Me.Controls("chk" & Trim(str(rs!ControlID))).AfterUpdate = "=chkConfirmDataEntryEnable()"
      
      Me.Controls("cmb" & Trim(str(rs!ControlID))).RowSource = rs!cmbRowSource
      Me.Controls("cmb" & Trim(str(rs!ControlID))).visible = rs!cmbVisible
      Me.Controls("cmb" & Trim(str(rs!ControlID))).Enabled = rs!CmbEnabled
      Me.Controls("cmb" & Trim(str(rs!ControlID))).left = rs!CmbLeft
      Me.Controls("cmb" & Trim(str(rs!ControlID))).Height = rs!CmbHeight
      Me.Controls("cmb" & Trim(str(rs!ControlID))).Width = rs!CmbWidth
      Me.Controls("cmb" & Trim(str(rs!ControlID))).top = lngPreviousTop
      Me.Controls("cmb" & Trim(str(rs!ControlID))).AfterUpdate = "=chkConfirmDataEntryEnable()"
      
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).Caption = rs!CtlLblCaption
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).FontWeight = rs!CtlLblFontWeight
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).visible = rs!CtlLblVisible
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).left = rs!CtlLblLeft
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).Height = rs!CtlLblHeight
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).Width = rs!CtlLblWidth
      Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).top = lngPreviousTop
      
      Me.Controls("txt" & Trim(str(rs!ControlID))).visible = rs!txtVisible
      Me.Controls("txt" & Trim(str(rs!ControlID))).Enabled = rs!txtEnabled
      Me.Controls("txt" & Trim(str(rs!ControlID))).left = rs!txtLeft
      Me.Controls("txt" & Trim(str(rs!ControlID))).Height = rs!TxtHeight
      Me.Controls("txt" & Trim(str(rs!ControlID))).Width = rs!txtWidth
      Me.Controls("txt" & Trim(str(rs!ControlID))).top = lngPreviousTop
      Me.Controls("txt" & Trim(str(rs!ControlID))) = rs!txtValue
      Me.Controls("txt" & Trim(str(rs!ControlID))).AfterUpdate = "=chkConfirmDataEntryEnable()"
      
      Me.Controls("txtLbl" & Trim(str(rs!ControlID))).Caption = rs!txtLblCaption
      Me.Controls("TxtLbl" & Trim(str(rs!ControlID))).visible = rs!txtLblVisible
      Me.Controls("txtLbl" & Trim(str(rs!ControlID))).left = rs!txtLblLeft
      Me.Controls("txtLbl" & Trim(str(rs!ControlID))).Height = rs!TxtLblHeight
      Me.Controls("txtLbl" & Trim(str(rs!ControlID))).Width = rs!txtlblWidth
      Me.Controls("txtLbl" & Trim(str(rs!ControlID))).top = lngPreviousTop
      
      lngPreviousTop = Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).top + Me.Controls("CtlLbl" & Trim(str(rs!ControlID))).Height + 50

      rs.MoveNext
    Next intI
    
Debug.Print lngPreviousTop
    
    Me.chkConfirmDataEntry.top = lngPreviousTop + 50
    Me.lbl_chkConfirmDataEntry.top = lngPreviousTop + 50
  
End Sub

Sub ReviewGuidelines_Update()
'===================================================================
'Will update the dbo.AuditClm_DocReview_DME_Saved; will call the usp_AuditClm_DocReview_DME_Saved_Update
'
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================
Dim strSQL As String
Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Dim intI As Integer

Dim myCode_ADO As clsADO
Dim cmd As ADODB.Command
Dim strErrMsg As String

Dim strAdj_ConceptID As String

On Error GoTo Err_handler

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = " SELECT * from AuditClm_DocReview_DME_saved where cnlyclaimnum = '" & strCnlyClaimNum & "' Order by ControlID"

    Set rs = MyAdo.OpenRecordSet(strSQL)

    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No records for this Claim "
        Exit Sub
    End If
    
    strAdj_ConceptID = rs.Fields("Adj_ConceptID")
    
        
    
    For intI = 1 To rs.recordCount
     
'===========================================================
        Set myCode_ADO = New clsADO
        myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = myCode_ADO.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "dbo.usp_AuditClm_DocReview_DME_Saved_Update"
        cmd.Parameters.Refresh
'==========================================================================
        cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
        cmd.Parameters("@pAdj_ConceptID") = strAdj_ConceptID
        cmd.Parameters("@pControlID") = CInt(rs!ControlID)
            If Me.Controls("chk" & Trim(str(rs!ControlID))).visible = "TRUE" Then
                cmd.Parameters("@pCtlValue") = (Me.Controls("chk" & Trim(str(rs!ControlID))))
            ElseIf Me.Controls("cmb" & Trim(str(rs!ControlID))).visible = "TRUE" Then
                cmd.Parameters("@pCtlValue") = (Me.Controls("cmb" & Trim(str(rs!ControlID))))
            End If
        cmd.Parameters("@ptxtValue") = Nz(Me.Controls("Txt" & Trim(str(rs!ControlID))), "")
        cmd.Execute
        
            If cmd.Parameters("@pValidDateMsg") <> "" Then
                MsgBox (cmd.Parameters("@pValidDateMsg"))
            End If
        
        
        'strErrMsg = cmd.Parameters("@pErrMsg")
            If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
                If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_AuditClm_DocReview_DME_Saved_Update"
                Err.Raise 50001, "usp_AuditClm_DocReview_DME_Saved_Update", strErrMsg
            End If
    
          rs.MoveNext
    Next intI

Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    
    
End Sub

Sub ReviewGuidelinesDisplay_Update()
'========================================================================================================
'Will refresh the values for the controls on the 'Review Guidelines' Page.  Needs to be called after the ReviewGuidelines_Update subroutine
'Developed by Kathleen C Flanagan Thursday 1/17/2013

'========================================================================================================

Dim strSQL As String
Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet
Dim intI As Integer
Dim lngPreviousTop As Long

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = " SELECT * from AuditClm_DocReview_DME_saved where cnlyclaimnum = '" & Me.txtCnlyClaimNum & "' Order by ControlID"
    Set rs = MyAdo.OpenRecordSet(strSQL)
     
    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No Template Information for this Claim "
        Exit Sub
    End If
    
    lngPreviousTop = 500
    
'Set the page to expand to fit the data
    Me.Controls("ReviewGuidelines").Height = 20000
    
    For intI = 1 To rs.recordCount
      'For each record, set the value on the page & format the display
      
      Me.Controls("chk" & Trim(str(rs!ControlID))) = rs!CtlValue  'new line
      Me.Controls("txt" & Trim(str(rs!ControlID))) = rs!txtValue

      rs.MoveNext
    Next intI
  
End Sub

Sub Determination_Update()
'===================================================================
'Will call the subroutines to update the Review Guidelines values, then switch to the Determination Page
'GenerateDetermination event will also be invoked when a user selects the 'Determination' page
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================

    ReviewGuidelines_Update 'Call subroutine that will update the database with the user entered values
    ReviewGuidelinesDisplay_Update 'Call subroutine that will refresh the 'Review Guidelines' page with the saved values
    Me.txtRevGuide.SetFocus
    Me.chkConfirmDataEntry.Enabled = False
    Me.chkConfirmDataEntry = False
    
    
    Me.Determination.visible = True
    Me.Determination.SetFocus 'When switch to Determination page, will load the DME_Determination subform
    
    Call Me.frm_AuditClm_DocReviewt_DME_Determination.Form.DMEDetermatination_RefreshData(strCnlyClaimNum)
    
    Me.cmdTakeFocus.SetFocus 'Will force the Determination page to scroll to the top (otherwise will scroll to the subform & hid the page tabs)
    
End Sub

Public Sub Rationale_Populate()
'===================================================================
'Called when the user selects the 'Doc Review - DME' option in the listbox on CA AuditClm_Main form; called after the DocReview_DME_ReviewSetup subroutine
'Will get the Adj_Rationale value from the database for the claim
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================
Dim strSQL As String
Dim MyAdo As clsADO
Dim rs As ADODB.RecordSet

    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = " SELECT adj_Rationale from AuditClm_Hdr where cnlyclaimnum = '" & Me.txtCnlyClaimNum & "'"
    Set rs = MyAdo.OpenRecordSet(strSQL)
     
    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        Me.txtAdj_Rationale = "No Rationale available for this claim"
        Exit Sub
    End If
    
    Me.txtAdj_Rationale = rs.Fields("Adj_Rationale")
    
End Sub



'Private Sub cmdGenerateDetermination_Click()
'
'Determination_Update
'
'Me.Determination.visible = True
'Me.Determination.SetFocus 'When switch to Determination page, will load the DME_Determination subform
'
'    Call Me.frm_AuditClm_DocReviewt_DME_Determination.Form.DMEDetermatination_RefreshData(strCnlyClaimNum)
'
'    Me.cmdTakeFocus.SetFocus 'Will force the Determination page to scroll to the top (otherwise will scroll to the subform & hid the page tabs)
'
'End Sub

Sub cmdGenerateRationale_Click()
    GenerateRationale

End Sub

Private Sub GenerateRationale()
'===================================================================
'Will populate & format the 'Rationale' page of the form with the table values; called when user selects the 'Generate Rationale' button on the 'Determination' Page
'Calls the usp_AuditClm_DocReview_DME_Rationale
'
'Developed December 2012 by Kathleen C. Flanagan
'===================================================================

    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    
    On Error GoTo Err_handler
    
    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "dbo.usp_AuditClm_DocReview_DME_Rationale"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyClaimNum") = Me.txtCnlyClaimNum
    cmd.Execute
    
    'strErrMsg = cmd.Parameters("@pErrMsg")
    If cmd.Parameters("@RETURN_VALUE") <> 0 Or strErrMsg <> "" Then
        If strErrMsg = "" Then strErrMsg = "Error executing stored procedure usp_AuditClm_DocReview_DME_Rationale"
        Err.Raise 50001, "usp_AuditClm_DocReview_DME_Rationale", strErrMsg
    End If
    
    If IsNull(cmd.Parameters("@txtRationale")) Then
        Me.txtAdj_Rationale = "There is not enough information"
    Else
        Me.txtAdj_Rationale = cmd.Parameters("@txtRationale")
    End If

    Me.Rationale.SetFocus
    txtAdj_Rationale_AfterUpdate 'This will set the main AuditClm form to dirty, which will prompt the user to save the record; the auditclm_hdr record will be updated

Exit_Sub:
    Set myCode_ADO = Nothing
    'Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub



Private Sub Form_Load()
    Dim iAppPermission As Integer
    iAppPermission = UserAccess_Check(Me)
    
    Me.chkConfirmDataEntry.Enabled = False
    
    Me.txtRevGuide.SetFocus
     
End Sub

Private Sub FormIsDirty()
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub



Private Sub tabChartReviewDME_Change()
'If user selects the Determination tab, will refresh the Determination

    If tabChartReviewDME.Value = 1 Then
        ReviewGuidelinesDisplay_Update
    ElseIf tabChartReviewDME.Value = 2 Then
        Me.cmdTakeFocus.SetFocus
    End If
        
End Sub


Sub txtAdj_Rationale_AfterUpdate()
'Will set the form to dirty, which will prompt user to save the record
   If Not rsAuditClmHdr.EOF Then
        rsAuditClmHdr.Fields("Adj_Rationale") = Me.txtAdj_Rationale
        FormIsDirty
    End If
End Sub

Function chkConfirmDataEntryEnable()
    
    If Me.chkConfirmDataEntry.Enabled = False Then
        MsgBox ("You have made updates to the Review Guidelines.  You will need to re-generate the Determination.")
    End If
    
    Me.chkConfirmDataEntry.Enabled = True
    Me.Determination.visible = False
    
End Function

Sub Form_Close()

Me.chkConfirmDataEntry.Enabled = False
End Sub



Private Sub chkConfirmDataEntry_Click()
    Determination_Update
End Sub
