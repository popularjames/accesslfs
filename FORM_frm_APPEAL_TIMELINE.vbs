Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 2013-0509 KD: I can't believe this wasn't in here until now

Private strCnlyClaimNum As String
Private strRowSource As String
Private strAppID As String

Private myCode_ADO As clsADO
Private rs As ADODB.RecordSet
Private strSQL As String

' 2013-05-09 KD: Needed to add 2 "drop downs" for this form, so took the opportunity to
'       clean up some code, change the insert to use usp_Appeal_INSERT_SQL instead of
'       dynamically creating the insert statement
'       and removed a couple of linked tables that aren't needed (ok, at least 1)


'2012-08-08 ***DPR ADDED MAIN GRID OBJECT TO EXPOSE EVENTS FROM THE GRID WITH APPEALS DATA
Private WithEvents oMainGrid As Form_frm_GENERAL_Datasheet
Attribute oMainGrid.VB_VarHelpID = -1
'2012-08-08 ***DPR ADDED MAIN GRID OBJECT TO EXPOSE EVENTS FROM THE GRID WITH APPEALS DATA

' 2013-05-09 KD added:
' 2014-06-18 VS: Added ALJ Code
' 2014-07-03 VS: Moved Clear_Exception_ALJ function to ALJ Module
' 2014-07-16 VS: Deleted commented out code
Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Property Let frmAppID(data As String)
    strAppID = data
End Property

Property Get frmAppID() As String
    frmAppID = strAppID
End Property

Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property

Property Let CnlyRowSource(data As String)
     strRowSource = data
End Property

Property Get CnlyRowSource() As String
     CnlyRowSource = strRowSource
End Property

Private Sub Command2_Click()
    RefreshData
End Sub

'This is a public refresh, so we can call it from elsewhere
Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
    'Refresh the grid based on the rowsource passed into the form
    Me.frm_GENERAL_Datasheet.Form.InitData strRowSource, 2
    Me.frm_GENERAL_Datasheet.Form.RecordSource = strRowSource

    '2012-08-08 ***DPR ADDED MAIN GRID OBJECT TO EXPOSE EVENTS FROM THE GRID WITH APPEALS DATA
    Set oMainGrid = Me.frm_GENERAL_Datasheet.Form
    oMainGrid_Current
    '2012-08-08 ***DPR ADDED MAIN GRID OBJECT TO EXPOSE EVENTS FROM THE GRID WITH APPEALS DATA


    Dim ctl As Control
     
    'Loop through the controls and size them correctly.
    For Each ctl In Me.frm_GENERAL_Datasheet.Form.Controls
      If ctl.ControlType = acTextBox Then
          ctl.ColumnWidth = -2
      End If
   Next
   


exitHere:
    Exit Sub
    
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub




Private Function Validation() As Boolean
'  This function will validate each one of the Inputs on the form before the record is inserted into the
' Appeal Hist Prod Table as a manualy Appeal Update Entry
Dim sMsg As String
    sMsg = ""
        ' Make sure required fields have values:
    txtICN.SetFocus
    If Me.txtICN.Text = "" Then
        sMsg = "  Icn blank" & vbCrLf
    End If
    
        'txtAppealIcn.SetFocus
        '   If Me.txtAppealIcn.Text = "" Then
        '       sMsg = "  Appeal Icn blank" & vbCrLf
        '   End If                                     'Commented these out because the Operations Team does not always recieve Appeal ICN
           
    
        ' Make sure required fields have values:
    cmbAppealLevel.SetFocus
    
        ' Accessing the value as opposed to the text of a combo box:
        ' me.AppealLevel.ItemData(me.AppealLevel.ListIndex)

    If Me.cmbAppealLevel.ListIndex = -1 Then
    
        sMsg = sMsg & "  Appeal Level blank" & vbCrLf
    End If
    
    txtAppealDt.SetFocus
    If Me.txtAppealDt.Text = "" Then
        sMsg = sMsg & "  Appeal Date blank" & vbCrLf
    End If
    
        '  txtDecisionDt.SetFocus
        '    If Me.txtDecisionDt.Text = "" Then
        '        sMsg = sMsg & "  Decision Date blank" & vbCrLf
        '    End If
            
    cmbAppealSource.SetFocus
    If Me.cmbAppealSource.ListIndex = -1 Then
        sMsg = sMsg & "  Appeal Source blank" & vbCrLf
    End If
    
   cmbAppealOutcome.SetFocus
    If Me.cmbAppealOutcome.ListIndex = -1 Then
        sMsg = sMsg & "  Appeal Outcome blank" & vbCrLf
    End If
    
   cmbPayer.SetFocus
    If Me.cmbPayer.ListIndex = -1 Then
        sMsg = sMsg & "  Payer blank" & vbCrLf
    End If
    
    If sMsg <> "" Then
        MsgBox "The following errors prevented the Record from being added:" & vbCrLf & sMsg
        Validation = True
    End If
    
End Function

Private Sub chkDefaults_Click()
    oMainGrid_Current
End Sub

Private Sub CmdAddRecord_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim rs As ADODB.RecordSet
Dim sUser As String
Dim dtDate As Date
Dim oCtl As Control

    strProcName = ClassName & ".cmdAddRecord_Click"
     

    If Validation() = True Then
        ' Do nothing because we've already told the user what they need to fix
        Exit Sub
    End If

    ' Confirmation...
    If MsgBox("Would you like to add this Appeal Event to the Appeal History Table?", vbYesNo, "Save") = vbNo Then
        Exit Sub
    End If
    
    Set oAdo = New clsADO
    
    sUser = Identity.UserName
    dtDate = Date
   
                ' 2013-05-09 KD: Following is the original code:
            '    oAdo.ConnectionString = GetConnectString("Appeal_Prod_History")
            '    oAdo.SQLTextType = sqltext
            '    oAdo.SqlString = " Insert into Appeal_Prod_History "
            '    oAdo.SqlString = oAdo.SqlString & "VALUES('" & CnlyClaimNum & "'," & _
            Me.cmbAppealLevel & ",'" & _
            cmbAppealOutcome & "','" & _
            Nz(txtReversalAmt, 0) & "','" & _
            txtReversalReason & "','" & _
            cmbAppealSource & "','" & _
            txtSFileName & "','" & _
            txtIcn & "','" & _
            dtDate & "','" & _
            sUser & "',1,'" & _
            cmbPayer & "','" & _
            txtAppealDt & "','" & _
            txtDecisionDt & "','" & _
            dtDate & "','" & _
            Nz(txtAppealIcn, "NULL") & "','" & _
            Nz(txtAppealComment, "NULL") & _
            "','" & _
            Nz(cmbALJJudgeName, "NULL")
            'oAdo.SqlString = oAdo.SqlString & "','" & _
            Nz(txtQICAppealNum, "NULL") & "','" & _
            Nz(txtALJAppealNum, "NULL") & "','" & _
            Nz(txtPromotionDt, "") & "','" & _
            Nz(cmbPayObsServices, "") & "','" & _
            Nz(cmbPayAncillaryServices, "") & "','" & _
            Nz(cmbOConnor, "") & "','" & _
            Nz(cmbRuledOnReopening, "") & "','" & _
            Nz(cmbWaiverLiab, "") & "','" & _
            Nz(cmbALJDecisionActual, "") & "','" & _
            Nz(cmbEHRClaim, "") & "','" & _
            Nz(txtNOISentDt, "")
            'oAdo.SqlString = oAdo.SqlString & "','" & _
            Nz(txtHearingDt, "") & "','" & _
            Nz(txtNOHRecDt, "") & "','" & _
            Nz(cmbParticpationInHearing, "") & "','" & _
            "00000" & "','" & _
            "00000" & "','" & _
            Nz(cmbParticipant, "") & "', "
            '
            '    If Me.chkValidABN.Value = -1 Then
            '        oAdo.SqlString = oAdo.SqlString & "1)"
            '    ElseIf Me.chkInvalidABN.Value = -1 Then
            '        oAdo.SqlString = oAdo.SqlString & "0)"
            '    Else
            '        oAdo.SqlString = oAdo.SqlString & Nz("NULL") & ")"
            '    End If
                  
            

    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_Appeal_INSERT_SQL"
        .Parameters.Refresh
        
        For Each oCtl In Me.Controls
            If Nz(oCtl.Properties("Tag"), "") <> "" Then
                Select Case UCase(oCtl.Name)
                Case "chkValidABN", "chkInvalidABN"  ' don't know why there are 2 check boxes..
                    If Not (Me.chkValidABN = 0 And Me.chkValidABN = 0) Then
                        If Me.chkInvalidABN <> 0 Then
                            .Parameters("@pABN").Value = 0
                        Else
                            .Parameters("@pABN").Value = 1
                        End If
                    End If
                Case "ckALJDecision"
                    .Parameters("@pALJDecisionEntered").Value = IIf(oCtl.Value = -1, 1, 0)
                Case "ckOTR"
                    .Parameters("@pOTR").Value = IIf(oCtl.Value = -1, 1, 0)
                Case "txtHearingDt"
                    If Nz(oCtl.Value, "") <> "" Then
                        .Parameters("@pHearingDt") = Format(Me.txtHearingDt.Value, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case Else
                    If Nz(oCtl.Value, "") <> "" Then
                        .Parameters("@p" & oCtl.Properties("Tag").Value) = oCtl.Value
                    Else
                    
                    End If
                End Select
            End If
        Next
        
        
        '' Couple others that aren't controls:
        .Parameters("@pCnlyClaimNum") = CnlyClaimNum
        .Parameters("@pLastUpdateDt") = dtDate
        .Parameters("@pLastUpdateUser") = sUser
        .Parameters("@pManualEventEntry") = 1
        .Parameters("@pLoadDt") = Now()
        .Parameters("@pReOpenInd") = "00000"
        .Parameters("@pDismissalInd") = "00000"
        
    End With
                                                                                                                                                                                                       
    oAdo.Execute
    
    If Me.cmbNOIStatus.Value <> "" Then
        Clear_Exception_ALJ (CnlyClaimNum)
    End If
    
    If Me.cmbNOIStatus.Value = NOI_STATUS_SENT Then
        AmendedFlag = False
        Add_ALJ_Package
    End If
    
    If Me.cmbNOIStatus.Value = NOI_STATUS_AMENDED Then
        AmendedFlag = True
        Add_ALJ_Package
    End If
    
    If Me.cmbNOIStatus.Value = NOI_STATUS_WITHDRAW_CLAIM Then
        'Call Delete package
    End If
    
    Me.chkDefaults = False

    If Nz(Me.cmbNOIStatus.Value, "") <> NOI_STATUS_SENT Then
        Me.frm_GENERAL_Datasheet.Form.Requery
    End If
 
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub Form_Close()
    'This form can be instanced, so it is removed from the global collection before it is closed
    RemoveObjectInstance Me
End Sub

Private Sub Command3_Click()
    On Error GoTo Err_Command3_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70

Exit_Command3_Click:
    Exit Sub

Err_Command3_Click:
    MsgBox Err.Description
    Resume Exit_Command3_Click
    
End Sub

Private Sub Form_Load()
    Set myCode_ADO = New clsADO
End Sub

Private Sub oMainGrid_Current()
'2012-08-08 ***DPR ADDED MAIN GRID OBJECT TO EXPOSE EVENTS FROM THE GRID WITH APPEALS DATA
On Error GoTo ErrHandler

    If Me.chkDefaults.Value <> 0 Then
    
        Me.txtICN = oMainGrid.Form.RecordSet.Fields("ICN")
        Me.txtAppealIcn = oMainGrid.Form.RecordSet.Fields("appealicn")
        Me.cmbAppealLevel = oMainGrid.Form.RecordSet.Fields("appeallevel")
        Me.txtAppealDt = oMainGrid.Form.RecordSet.Fields("appealdt")
        Me.txtReversalAmt = oMainGrid.Form.RecordSet.Fields("reversalamt")
        Me.txtReversalReason = oMainGrid.Form.RecordSet.Fields("reversalreason")
        Me.txtDecisionDt = oMainGrid.Form.RecordSet.Fields("decisiondt")
        Me.cmbAppealOutcome = oMainGrid.Form.RecordSet.Fields("appealoutcome")
        Me.cmbAppealSource = oMainGrid.Form.RecordSet.Fields("appealsource")
        Me.cmbPayer = oMainGrid.Form.RecordSet.Fields("payer")
        Me.txtSFileName = oMainGrid.Form.RecordSet.Fields("sfilename")
        
        Me.txtQICAppealNum = oMainGrid.Form.RecordSet.Fields("QICAppealNumber")
        Me.txtALJAppealNum = oMainGrid.Form.RecordSet.Fields("ALJAppealNumber")
        Me.cmbALJJudgeName = oMainGrid.Form.RecordSet.Fields("ALJJudgeName")
        Me.txtPromotionDt = oMainGrid.Form.RecordSet.Fields("NOHRecDt")
        Me.txtNOHRecDt = oMainGrid.Form.RecordSet.Fields("NOHRecDt")
        Me.txtHearingDt = Format(oMainGrid.Form.RecordSet.Fields("HearingDt"), "mm/dd/yyyy hh:mm:ss AMPM")
'oMainGrid.Form.Recordset.Fields ("HearingDt")
        Me.txtNOISentDt = oMainGrid.Form.RecordSet.Fields("NOISentDate")
        Me.cmbEHRClaim = oMainGrid.Form.RecordSet.Fields("EHRFlag")
        Me.txtDecisionDt = oMainGrid.Form.RecordSet.Fields("DecisionDt")
        Me.cmbParticpationInHearing = oMainGrid.Form.RecordSet.Fields("ParticipatedInHearing")
        Me.cmbPayObsServices = oMainGrid.Form.RecordSet.Fields("PayOutpatientObservServices")
        Me.cmbPayAncillaryServices = oMainGrid.Form.RecordSet.Fields("PayAncillaryServices")
        Me.cmbParticipant = oMainGrid.Form.RecordSet.Fields("Participant")
'        Me.cmbOConnor = oMainGrid.Form.Recordset.Fields("FollowOConnorDecision")
'        Me.cmbRuledOnReopening = oMainGrid.Form.Recordset.Fields("RuledOnReopening")
'        Me.cmbWaiverLiab = oMainGrid.Form.Recordset.Fields("WaiverofLiabWithoutFault")
'
'        If Trim(oMainGrid.Form.Recordset.Fields("ABN")) = "Valid ABN" Then
'                Me.chkValidABN.Value = -1
'                Me.chkInvalidABN.Value = 0
'            ElseIf Trim(oMainGrid.Form.Recordset.Fields("ABN")) = "Invalid ABN" Then
'                Me.chkValidABN = 0
'                Me.chkInvalidABN = -1
'            Else
'                Me.chkValidABN.Value = 0
'                Me.chkInvalidABN = 0
'        End If
            
        
        If IsNull(oMainGrid.Form.RecordSet.Fields("ABN")) = True Then
                Me.chkValidABN.Value = 0
                Me.chkInvalidABN = 0
            ElseIf oMainGrid.Form.RecordSet.Fields("ABN") = 0 Then
                Me.chkValidABN = 0
                Me.chkInvalidABN = -1
            ElseIf oMainGrid.Form.RecordSet.Fields("ABN") = 1 Then
                Me.chkValidABN.Value = -1
                Me.chkInvalidABN.Value = 0

        End If


        Me.cmbNOIStatus = oMainGrid.Form.RecordSet.Fields("NOIStatus")
        Me.cmbHearing = oMainGrid.Form.RecordSet.Fields("HearingOTR")
        Me.cmbPaperClinical = oMainGrid.RecordSet.Fields("PS_Summary_Written") 'VS Was missing?
        
        If oMainGrid.Form.RecordSet.Fields("ALJDecisionEntered") = 0 Then
            Me.ckALJDecision.Value = 0
        Else
            Me.ckALJDecision.Value = -1
        End If
                
'        If oMainGrid.Form.Recordset.Fields("OTR") = 0 Then
'            Me.ckOTR.Value = 0
'        Else
'            Me.ckOTR.Value = -1
'        End If
                
'        If oMainGrid.Form.Recordset.Fields("Withdrawn") = 0 Then
'            Me.ckWithdrawn.Value = 0
'        Else
'            Me.ckWithdrawn.Value = -1
'        End If
                
    Else
    
        Me.txtICN = ""
        Me.txtAppealIcn = ""
        Me.cmbAppealLevel = ""
        Me.txtAppealDt = ""
        Me.txtReversalAmt = ""
        Me.txtReversalReason = ""
        Me.txtDecisionDt = ""
        Me.cmbAppealOutcome = ""
        Me.cmbAppealSource = ""
        Me.cmbPayer = ""
        Me.txtSFileName = ""
        
        Me.txtQICAppealNum = ""
        Me.txtALJAppealNum = ""
        Me.cmbALJJudgeName = ""
        Me.txtPromotionDt = ""
        Me.txtNOHRecDt = ""
        Me.txtHearingDt = ""
        Me.txtNOISentDt = ""
        Me.cmbEHRClaim = ""
        Me.txtDecisionDt = ""
        Me.cmbParticpationInHearing = ""
        Me.cmbParticipant = ""
        Me.chkInvalidABN.Value = 0
        Me.chkValidABN.Value = 0
        Me.cmbPayObsServices = ""
        Me.cmbPayAncillaryServices = ""
        
'        Me.cmbOConnor = ""
'        Me.cmbRuledOnReopening = ""
'        Me.cmbWaiverLiab = ""

        Me.cmbNOIStatus = ""
        Me.cmbHearing = ""

        Me.ckALJDecision.Value = 0
        Me.cmbPaperClinical = ""
'        Me.ckOTR.Value = 0
'        Me.ckWithdrawn.Value = 0
    End If




Exit Sub
ErrHandler:
    MsgBox "Error populating smart defaults.", vbCritical + vbOKOnly
End Sub


Public Sub Add_ALJ_Package()
    
    Dim ErrorReturned As String
    Dim PackageID As String
    
    If (Nz(Me.cmbALJJudgeName.Value, "") = "" Or Nz(Me.txtHearingDt.Value, "") = "" _
    Or Nz(Me.txtALJAppealNum.Value, "") = "" Or Nz(Me.cmbEHRClaim.Value, "") = "" _
    Or Nz(Me.cmbParticipant.Value, "") = "") Then
    
    MsgBox ("Please make sure you entered Judge Name, Hearing Date, Appeal Number, Appellant Name and Hearing Participant. Because one of these values is missing, the package can not be created!")
    GoTo Exit_Sub
    End If

    On Error GoTo Err_handler

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
                myCode_ADO.SQLTextType = StoredProc
                myCode_ADO.sqlString = "usp_Add_ALJ_Package"
                myCode_ADO.Parameters("@pJudgeName") = Trim(Me.cmbALJJudgeName.Value)
                myCode_ADO.Parameters("@pHearingDate") = Format(Me.txtHearingDt.Value, "yyyy-mm-dd hh:mm:ss")
                myCode_ADO.Parameters("@pALJAppealNumber") = Trim(Me.txtALJAppealNum.Value)
                myCode_ADO.Parameters("@pCnlyClaimNum") = CnlyClaimNum
                myCode_ADO.Parameters("@pAppellantName") = Me.cmbEHRClaim.Value
                myCode_ADO.Parameters("@pConnollyParticipant") = Me.cmbParticipant.Value
                myCode_ADO.Parameters("@pTestimonyType") = Me.cmbParticpationInHearing.Value
                
                If Nz(Me.cmbPartyStatus.Value, "") <> "" Then
                    myCode_ADO.Parameters("@pPartyStatus") = Me.cmbPartyStatus.Value
                End If
                
                If Nz(Me.txtHearingPhone.Value, "") <> "" Then
                    myCode_ADO.Parameters("@pHearingPhone") = Me.txtHearingPhone.Value
                    myCode_ADO.Parameters("@pHearingPasscode") = Me.txtHearingPasscode.Value
                End If
                
                Set rs = myCode_ADO.ExecuteRS
                
                ErrorReturned = Nz(myCode_ADO.Parameters("@pErrMsg").Value, "")
                If ErrorReturned <> "" Then
                    MsgBox ErrorReturned, vbExclamation
                    Exit Sub
                Else
                    PackageID = Nz(myCode_ADO.Parameters("@pALJPackageId").Value, "")
                    If PackageID <> "" Then
                        MsgBox ("ALJ Hearing Package  " + PackageID + "  had been created.")
                        algPackageName = PackageID
                        algCnlyClaimNum = CnlyClaimNum
                        algAppealNum = Trim(Me.txtALJAppealNum.Value)
                        CreatePackageFolder
                        CopyALJNoticeFile
                        
                        If AmendedFlag = True Then
                           Call Delete_Claim(PackageID, True)
                        End If
                        
                        Set rs = myCode_ADO.ExecuteRS
                        
                    Else
                        
                        algCnlyClaimNum = CnlyClaimNum
                        algHearDate = Format(Me.txtHearingDt.Value, "yyyy-mm-dd hh:mm:ss")
                        algAppealNum = Trim(Me.txtALJAppealNum.Value)
                        algJudgeName = Me.cmbALJJudgeName.Value
                        algAppellant = Me.cmbEHRClaim.Value
                        algParticipant = Me.cmbParticipant.Value
                        algTestimonyType = Me.cmbParticpationInHearing.Value
                        
                        If Nz(Me.cmbPartyStatus.Value, "") <> "" Then
                            algPartyStatus = Me.cmbPartyStatus.Value
                        End If
                        
                        If Nz(Me.txtHearingPhone.Value, "") <> "" Then
                            algHearingPhone = Me.txtHearingPhone.Value
                            algHearingPasscode = Me.txtHearingPasscode.Value
                        End If
                             
                    End If
                    
                        DoCmd.OpenForm "frm_Existing_ALJ_Packages", , , , , , Me.Name
                        Set Forms("frm_Existing_ALJ_Packages").lstExistPkgs.RecordSet = rs
                    
                End If
                               
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub

Err_handler:
        MsgBox Err.Description
End Sub

Private Sub cmbALJJudgeName_NotInList(NewData As String, Response As Integer)
 
 Dim intAnswer As Integer
 Dim strSQL As String
 
 On Error GoTo Err_handler
 
 intAnswer = MsgBox("The judge you selected: " & Chr(34) & NewData & _
 Chr(34) & " is not in the list." & vbCrLf & _
 "Would you like to add this new judge to the list now?" _
 , vbQuestion + vbYesNo, "")
 
If intAnswer = vbYes Then
     myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
     myCode_ADO.SQLTextType = StoredProc
     myCode_ADO.sqlString = "usp_Add_ALJ_Judge"
     myCode_ADO.Parameters("@pJudgeName") = NewData
     Set rs = myCode_ADO.ExecuteRS
    
     MsgBox "New judge had been added successfully." _
     , vbInformation, ""
     Response = acDataErrAdded
 Else
    MsgBox "Please choose the Judge from the list." _
    , vbInformation, ""
    Response = acDataErrContinue
 End If
 
Exit_Sub:
    Set rs = Nothing
    Set myCode_ADO = Nothing
    Exit Sub
 
Err_handler:
        MsgBox Err.Description
 
 End Sub
