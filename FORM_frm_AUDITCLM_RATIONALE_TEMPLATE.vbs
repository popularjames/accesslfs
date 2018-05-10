Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AuditClmRationale"
Public clnZoomWindows As New Collection
Private WithEvents frmRationalezoom As Form_frm_AUDITCLM_Rationale_Zoom
Attribute frmRationalezoom.VB_VarHelpID = -1
Private WithEvents frmObsDisplay As Form_frm_GENERAL_Display
Attribute frmObsDisplay.VB_VarHelpID = -1

Private rsAuditClmHdr As ADODB.RecordSet
Private rsAuditClmProc As ADODB.RecordSet

Private intTemplateID  As Integer
Private strCnlyClaimNum As String
Public Event RationaleConfirmed(strCommitRationale As String, strHwnd As String)
Public Event FormClosed()

Property Let TemplateID(data As Integer)
     intTemplateID = data
End Property
Property Get TemplateID() As Integer
     TemplateID = intTemplateID
End Property
Property Let CnlyClaimNum(data As String)
     strCnlyClaimNum = data
End Property
Property Get CnlyClaimNum() As String
     CnlyClaimNum = strCnlyClaimNum
End Property
Property Set HdrRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmHdr = data
End Property
'3/6/2013 BEGIN KCF: Bring in the recordsets to support the DRG Template Proc & Diag code lookups
Property Set ProcCodeRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmProc = data
End Property

'3/6/2013 END KCF: Bring in the recordsets to support the DRG Template Proc & Diag code lookups


Public Sub RefreshData()
'Called by the form that launches this form. Takes properties and applies them to the data elements
'Revised 3/7 - 3/12/2013 by KCF: Set up for new DRG templates - primarily new lookups included
'Revised 11/25 by KCF:  Fix the error handling for HH claims on the Diagnosis Code (date handling).
    Dim strSQL As String
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim intI As Integer
    Dim strLookupTable As String
    Dim strLookupValue As String
    Dim strFieldValue As String
    Dim strLookupKey As String
    Dim lngPreviousTop As Long
    Dim intSeqNo As Integer ' andrew
    Dim ctrl As Control '12/13/2012 by KCF
    
    Dim strProcCode As String '3/6/2013 by KCF
    
    Dim rsOrigProc As ADODB.RecordSet '3/11/2013
    Dim strSQLOrigProc As String '3/11/2013
    
    Dim rsFirstOrigDiag As ADODB.RecordSet  '3/12/2013 by KCF
    Dim strSQLFirstOrigDiag As String '3/12/2013 by KCF
    Dim rsFirstRevDiag As ADODB.RecordSet '3/12/2013 by KCF
    Dim strSQLFirstRevDiag As String '3/12/2013 by KCF
    
    Dim rsSecDiag As ADODB.RecordSet '3/7/2013 by KCF
    Dim strSQLSecDiag As String '3/7/2013 by KCF
    
    
    On Error GoTo ErrHandler
    
    If IsNull(Me.txtBoxCnlyClaimNum) Then Me.txtBoxCnlyClaimNum = CnlyClaimNum
    If IsNull(Me.txtBoxRationaleID) Then Me.txtBoxRationaleID = TemplateID
    If IsNull(Me.txtBoxFlip) Then Me.txtBoxFlip = 0
    If IsNull(Me.txtBoxLoadSaved) Then Me.txtBoxLoadSaved = 0
    Me.cboFieldData1 = ""

'12/13/2012 BEGIN: KCF Set initial formatting for the fields
    'CtrlFormat tag for the fields set up
    For Each ctrl In Me.Controls
        
        If ctrl.Name = "cmbBoxRationaleID" Then
            ctrl.Tag = ""
        End If
        
        If ctrl.Tag = "CtrlFormat" Then
            ctrl.visible = False
        End If
    Next ctrl
    
    Me.txtBoxRationaleID.visible = True
'12/13/2012 END: KCF Set initial formatting for the fields
               
    'refresh saved rationles
    strSQL = " select distinct seqno, username, RationaleID,cnlyclaimnum, createdt from AuditClm_Rationale_Saved where cnlyClaimnum = '" & Me.txtBoxCnlyClaimNum & "'"
    RefreshListBox strSQL, Me.lstSavedRationale
    
    If Me.txtBoxFlip = 0 Then
        intSeqNo = GetNextSequence() - 1 'andrew
        If intSeqNo > 0 Then  'andrew
            Me.txtBoxRationaleID = GetRationaleID(intSeqNo) 'andrew
        End If 'andrew
            Me.cmbBoxRationaleID = Me.txtBoxRationaleID
    End If
        
    'Keep count on how may comboboxes are repurposed for lookups.
    Dim intComboCounter As Integer
    intComboCounter = 1
    
    'Get the values for the template being applied to this form
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = "SELECT * from AuditClm_Rationale_Template where rationaleID = " & Me.txtBoxRationaleID & " Order by Fieldid" 'KCF 1/22/2013 to put text boxes in proper order
    Set rs = MyAdo.OpenRecordSet(strSQL)
     
    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No Template Information for this Claim "
        Exit Sub
    End If
    
    'Clear out the final rationale
    Me.txtFinalRationale = ""
    'Loop through the template fields
    
      lngPreviousTop = 1000
    For intI = 1 To rs.recordCount
      'Set the text box label
      Me.Controls("lblField" & Trim(str(rs!FieldID))).Caption = rs!FieldDisplay
      Me.Controls("lblField" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
      Me.Controls("lblField" & Trim(str(rs!FieldID))).top = lngPreviousTop
      Me.Controls("lblField" & Trim(str(rs!FieldID))).visible = True '12/13/2012 KCF
      'Set the base text to be used to construct this part of the rationale
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))) = rs!TextRationale
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))).top = lngPreviousTop
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))).visible = True '12/13/2012 KCF
      'Set the data entry control's tag to contain the token that will be replaced in the rationale template
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).Tag = rs!token
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
'12/13/2012 BEGIN: KCF Do not display the Observation FieldText
            If rs!FieldName = "Observation Language" Then
              Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).visible = False
            Else
              Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).visible = True
            End If
'12/13/2012 END: KCF Do not display the Observation FieldText
      'empty out the field to start fresh
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).top = lngPreviousTop
      
      'Andrew Adding prompts 08-30-2012
      Me.Controls("txtFieldPrompt" & Trim(str(rs!FieldID))) = rs!RationalePrompt
      
      'Set the value of the bottom of the last control to dynamically size them
      lngPreviousTop = Me.Controls("lblField" & Trim(str(rs!FieldID))).top + Me.Controls("lblField" & Trim(str(rs!FieldID))).Height + 50
           
      'Populate the default values if they exist in the table
      'I am Losing Sream on making this generic, I am putting some specific things to deal with this on the medical necessity for the moment
      'If the row in the table has a data entry type it means we need to fill it in
       If rs!entryType <> "Combo" Then
            'Get the data element that the textbox needs to be populated with
            strLookupValue = Nz(DLookup("LookupValue", "AuditClm_Rationale_Template", "RationaleID = " & Me.txtBoxRationaleID & " AND FieldID = " & rs!FieldID & ""), "")
                           
            Select Case Nz(strLookupValue, "")
'1/16/2013 BEGIN: KCF Set up Case for Complications for Surgical Template
            Case "Complications"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = rs!R3LookupValues
'1/16/2013 END: KCF Set up Case for Complications for Surgical Template
            Case "DischargeStatus"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = rsAuditClmHdr.Fields("DischargeStatus")
            Case "DischargeStatusCode"
                If Me.txtBoxRationaleID = 1 Or Me.txtBoxRationaleID = 2 Or Me.txtBoxRationaleID = 3 Then  'Because there is text appended to the discharge desc for the MN claims
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = LCase(DLookup("DischargeStatusCodeDesc", "UBDischargeStatus", "DischargeStatusCode = '" & rsAuditClmHdr.Fields("DischargeStatus") & "'")) & " in stable condition with instructions for follow up with the physician. "
                Else
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = LCase(DLookup("DischargeStatusCodeDesc", "UBDischargeStatus", "DischargeStatusCode = '" & rsAuditClmHdr.Fields("DischargeStatus") & "'"))
                End If
            Case "DRG"
                'Get the DRG from the recordser
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = rsAuditClmHdr.Fields("DRG")
            Case "SexCd"
                'Get the sex from the recordset
                If rsAuditClmHdr.Fields("BeneSexCd") = "M" Then
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = "male"
                ElseIf rsAuditClmHdr.Fields("BeneSexCd") = "F" Then
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = "female"
                End If
            Case "Age"
                    'Calculate the age from the recordset
                    'Me.Controls("txtFieldData" & Trim(Str(rs!FieldID))) = DateDiff("yyyy", rsAuditClmHdr.Fields("BeneBirthDt"), rsAuditClmHdr.Fields("IPAdmitDate"))
                    If (Nz(rsAuditClmHdr.Fields("IPAdmitDate"), "") = "" Or Nz(rsAuditClmHdr.Fields("BeneBirthDt"), "") = "") Then
                        Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
                    Else
                        Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Age(rsAuditClmHdr.Fields("BeneBirthDt"), rsAuditClmHdr.Fields("IPAdmitDate"))
                    End If
            Case "Year"
                    If Nz(rsAuditClmHdr.Fields("IPAdmitDate"), "") = "" Then
                        Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
                    Else
                        Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Yyear(rsAuditClmHdr.Fields("IPAdmitDate"))
                   End If
'3/5/2013 BEGIN KCF: Set up lookup fields for DRG templates
            Case "Adj_DischargeStatus"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsAuditClmHdr.Fields("Adj_DischargeStatus"), "")
            Case "Adj_DischargeStatusCode"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = LCase(DLookup("DischargeStatusCodeDesc", "UBDischargeStatus", "DischargeStatusCode = '" & rsAuditClmHdr.Fields("Adj_DischargeStatus") & "'")) & " in stable condition with instructions for follow up with the physician. "
            Case "AdmitDiag"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = LCase(DLookup("DxCodeDesc", "DxCode", "DxCode = '" & rsAuditClmHdr.Fields("AdmitDiag") & "'"))
            Case "ORIGPROC"
                If rsAuditClmProc Is Nothing Then
                    Set MyAdo = New clsADO
                    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                    strSQLOrigProc = "SELECT * from AuditClm_Proc where LineNum = 1 and CnlyClaimNum = '" & Me.txtBoxCnlyClaimNum & "'"
                    Set rsOrigProc = MyAdo.OpenRecordSet(strSQLOrigProc)
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(LCase(DLookup("PxCodeDesc", "PxCode", "PxCode = '" & rsOrigProc.Fields("ProcCd") & "' and PxCodeEffectiveDT < #" & rsOrigProc.Fields("ProcDt") & "#" & " and PxCodeEndDT > #" & rsOrigProc.Fields("ProcDt") & "#")), "")
                Else
                    rsAuditClmProc.Find "LineNum = 1"
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(LCase(DLookup("PxCodeDesc", "PxCode", "PxCode = '" & rsAuditClmProc.Fields("ProcCd") & "' and PxCodeEffectiveDT < #" & rsAuditClmProc.Fields("ProcDt") & "#" & " and PxCodeEndDT > #" & rsAuditClmProc.Fields("ProcDt") & "#")), "")
                End If
            Case "SODIAG"
                Set MyAdo = New clsADO
                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                strSQLSecDiag = "SELECT * from AuditClm_Diag where LineNum = 2 and CnlyClaimNum = '" & Me.txtBoxCnlyClaimNum & "'"
                Set rsSecDiag = MyAdo.OpenRecordSet(strSQLSecDiag)
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsSecDiag.Fields("DiagCd") & " " & UCase(DLookup("DxCodeDesc", "DxCode", "DxCode = '" & rsSecDiag.Fields("DiagCd") & "' and DxCodeEffectiveDt < #" & rsAuditClmHdr.Fields("IPDischargeDt") & "#" & " and DxCodeEndDt > #" & rsAuditClmHdr.Fields("IPDischargeDt") & "#")), "")
            Case "SODIAG2"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsSecDiag.Fields("DiagCd"), "")
            Case "Adj_DRG"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsAuditClmHdr.Fields("Adj_DRG"), "")
            Case "FODIAG"
                Set MyAdo = New clsADO
                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                strSQLFirstOrigDiag = "SELECT * from AuditClm_Diag where LineNum = 1 and CnlyClaimNum = '" & Me.txtBoxCnlyClaimNum & "'"
                Set rsFirstOrigDiag = MyAdo.OpenRecordSet(strSQLFirstOrigDiag)
                'BEGIN KCF 11/25/2013 - Set up to pick up the ClmThruDt when the IPDischargeDate is null
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsFirstOrigDiag.Fields("DiagCd") & " " & UCase(DLookup("DxCodeDesc", "DxCode", "DxCode = '" & rsFirstOrigDiag.Fields("DiagCd") & "' and DxCodeEffectiveDt < #" & Nz(rsAuditClmHdr.Fields("IPDischargeDt"), rsAuditClmHdr.Fields("ClmThruDt")) & "#" & " and DxCodeEndDt > #" & Nz(rsAuditClmHdr.Fields("IPDischargeDt"), rsAuditClmHdr.Fields("ClmThruDt")) & "#")), "")
                'END KCF 11/25/2013
            Case "FRDIAG"
                Set MyAdo = New clsADO
                MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
                strSQLFirstRevDiag = "SELECT * from AuditClm_Revised_Diag where LineNum = 1 and CnlyClaimNum = '" & Me.txtBoxCnlyClaimNum & "'"
                Set rsFirstRevDiag = MyAdo.OpenRecordSet(strSQLFirstRevDiag)
                If rsFirstRevDiag.BOF = True And rsFirstRevDiag.EOF = True Then
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
                Else
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsFirstRevDiag.Fields("DiagCd") & " " & UCase(DLookup("DxCodeDesc", "DxCode", "DxCode = '" & rsFirstOrigDiag.Fields("DiagCd") & "' and DxCodeEffectiveDt < #" & rsAuditClmHdr.Fields("IPDischargeDt") & "#" & " and DxCodeEndDt > #" & rsAuditClmHdr.Fields("IPDischargeDt") & "#")), "")
                End If
            Case "FRDIAG1"
                If rsFirstRevDiag.BOF = True And rsFirstRevDiag.EOF = True Then
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
                Else
                    Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rsFirstRevDiag.Fields("DiagCd"), "")
                End If
            Case "Reference"
                Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = Nz(rs!R3LookupValues, "")
            End Select
'3/5/2013 END KCF: Set up lookup fields for DRG templates
       End If
       
    If Me.txtBoxFlip = 0 Then
        'load latest saved form at start 'andrew
        intSeqNo = GetNextSequence() - 1 'andrew
        If intSeqNo > 0 Then  'andrew
        LoadSavedRationale (intSeqNo) 'andrew
        End If 'andrew
    End If
      'Handle combo boxes where the user needs a choice
      If Nz(rs!entryType, "") = "Combo" Then
          'Set the rowsource for the combo box
          Me.Controls("lblCombo" & Trim(str(intComboCounter))).Caption = rs!FieldDisplay
          Me.Controls("cboFieldData" & Trim(str(intComboCounter))).RowSource = Nz(rs!RowSource)
          'Set the tag of the control to the template portion we are going to update
          Me.Controls("cboFieldData" & Trim(str(intComboCounter))).Tag = Trim(str(rs!FieldID)) 'Me.Controls("txtFieldText" & Trim(Str(rs!FieldID))).Name
          'keep track of how many combo boxes we use.
          'This value is not doing anything at this point
           intComboCounter = intComboCounter + 1
      End If
      rs.MoveNext
    Next intI

If rs.recordCount < 13 Then
    For intI = rs.recordCount + 1 To 13
    Me.Controls("lblField" & intI).visible = False
    Me.Controls("txtFieldData" & intI).visible = False
    Me.Controls("txtFieldText" & intI).visible = False
    Next intI
End If

If Me.txtBoxRationaleID = 1 Or Me.txtBoxRationaleID = 2 Or Me.txtBoxRationaleID = 3 Then
    Me.cboFieldData1.visible = True
Else
    Me.cboFieldData1.visible = False
End If

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "RefreshData"
End Sub
Private Sub SaveRationale()
 Dim strSQL As String
    Dim rs As ADODB.RecordSet
    Dim intSeqNum As Integer
    Dim MyAdo As clsADO
    Dim intI As Integer
    
    Set MyAdo = New clsADO

    On Error GoTo ErrHandler
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
     
    strSQL = "SELECT * from AuditClm_Rationale_Template where rationaleID = " & Me.txtBoxRationaleID
    Set rs = MyAdo.OpenRecordSet(strSQL)

    intSeqNum = GetNextSequence
    'Make sure it exists
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No Template Information for this Claim "
        Exit Sub
    End If

    For intI = 1 To rs.recordCount
'12/10/2012 by KCF included the update to the IsCommittedToClaim and CommittedDAte fields
'3/20/2013 by KCF to handle single quotes in the text
        strSQL = " INSERT INTO AuditClm_Rationale_Saved (cnlyClaimNum,SeqNo,UserName,FieldID,FieldValue, Rationale, RationaleID, IsCommittedToClaim, CommittedDate, CreateDt) "
        strSQL = strSQL & " VALUES ( " & Chr(34) & Nz(Me.txtBoxCnlyClaimNum, "") & Chr(34) & ", " & intSeqNum & " , " & Chr(34) & Identity.UserName & Chr(34) & " , " & intI & " , " & Chr(34) & Nz(Replace(Me.Controls("txtFieldData" & Trim(str(rs!FieldID))), """", "'"), 0) & Chr(34) & "," & Chr(34) & left(Me.txtFinalRationale, 8000) & Chr(34) & "," & Me.txtBoxRationaleID & "," & 0 & ", '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & " ', '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "')"
        CurrentDb.Execute (strSQL)
        rs.MoveNext
    Next intI

Exit Sub

ErrHandler:
    MsgBox "Error saving, please copy and paste the generated rationale to your clipboard prior to closing the form - " & Err.Description, vbOKOnly + vbCritical
End Sub



'This will generate the rationale based on the entered data
Private Sub GenerateRationale()
    Dim strSQL As String
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim intI As Integer
    Dim strFinalRationale As String

    On Error GoTo ErrHandler
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    'Get the template that we are working with
    strSQL = "SELECT * from AuditClm_Rationale_Template where rationaleID = " & Me.txtBoxRationaleID & " Order by FieldID" 'KCF 1/22/2013
    
    Set rs = MyAdo.OpenRecordSet(strSQL)
     
    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No Template Information for this Claim "
        Exit Sub
    End If

    'Set the final rationale to nothing to start fresh
    strFinalRationale = ""
    
    'Go through all the fields
    For intI = 1 To rs.recordCount
      
     'concatenate the  rationale based on the data entered by the users
     '3/20/2103 KCF: Handle inclusion of single quotes in the text
     strFinalRationale = strFinalRationale & Replace(Replace(Nz(Me.Controls("txtFieldText" & Trim(str(rs!FieldID))), ""), Nz(Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).Tag, ""), Nz(Me.Controls("txtFieldData" & Trim(str(rs!FieldID))), "")), """", "'")
     
     If Nz(rs!termline, 0) = -1 Then
          'Insert a carriage return line feed if needed
         strFinalRationale = strFinalRationale & vbCrLf & vbCrLf
     End If
     rs.MoveNext
    
    Next intI
    
    Me.txtFinalRationale = strFinalRationale

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "GenerateRationale"
End Sub


Private Sub cboFieldData1_AfterUpdate()

    Dim ctrl As Control
    Dim strLookupTable As String
    Dim strLookupValue As String
    Dim strFieldValue As String
    Dim strLookupKey As String

    On Error GoTo ErrHandler


    'The tag of the combobox contains the field we are looking at
    'Get the database values we need todo the lookupo with
    strLookupTable = DLookup("TableName", "AuditClm_Rationale_Template", "RationaleID = " & Me.txtBoxRationaleID & " AND FieldID = " & Me.cboFieldData1.Tag & "")
    strLookupValue = DLookup("LookupValue", "AuditClm_Rationale_Template", "RationaleID = " & Me.txtBoxRationaleID & " AND FieldID = " & Me.cboFieldData1.Tag & "")
    strLookupKey = DLookup("KeyValue", "AuditClm_Rationale_Template", "RationaleID = " & Me.txtBoxRationaleID & " AND FieldID = " & Me.cboFieldData1.Tag & "")
    strFieldValue = DLookup(strLookupValue, strLookupTable, strLookupKey & " = " & Me.cboFieldData1 & "")
    
    'Set the control
    Me.Controls("txtFieldText" & Trim(str(Me.cboFieldData1.Tag))) = strFieldValue


Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "cboFieldData1_AfterUpdate"
End Sub

Private Sub cmdSpellCheck_Click()

On Error GoTo ErrHandler

    Me.txtFinalRationale.SetFocus
    DoCmd.RunCommand acCmdSpelling
Exit Sub

ErrHandler:
    MsgBox "Spelling Check Failed - " & Err.Description, vbOKOnly + vbCritical


End Sub



Private Sub cmbBoxRationaleID_Change()

If MsgBox("Are you sure that you want to switch to a different template? The below data will not be saved. ", vbYesNo, "Warning") = vbYes Then
 
    Me.txtBoxFlip = 1
    Me.txtBoxRationaleID = Me.cmbBoxRationaleID
    RefreshData
    Me.txtBoxFlip = 0
Else
Exit Sub
End If
End Sub

Private Sub Command24_Click()

    If Validation() = True Then 'Andrew
        ' Do nothing because we've already told the user what they need to fix
        Exit Sub
    End If

    GenerateRationale
    Me.TabCtl54.Pages.Item(1).SetFocus
    'BEGIN 11/8/2012 KCF - to select all text so that right-click, copy will pick up all text
    Me.txtFinalRationale.SetFocus
    Me.txtFinalRationale.SelStart = 0
    Me.txtFinalRationale.SelLength = Len(Me.txtFinalRationale.Value)
    'END 11/8/2012 KCF - to select all text so that right-click, copy will pick up all text
    
End Sub
Private Sub Command48_Click()

    SaveRationale
    If Me.txtTotalChar > "8000" Then
    MsgBox ("Your Rationale is too long. Please condense your rationale so that it only contains 8000 characters or less.")
    Exit Sub
    End If
    
    
    'If the user commits the rationale, we will send it back to the calling form.
    RaiseEvent RationaleConfirmed(Nz(Me.txtFinalRationale), Me.hwnd)
    RaiseEvent FormClosed
Exit Sub

ErrHandler:
    MsgBox "Error saving, plese copy and paste the generated rationale to your clipboard prior to closing the form - " & Err.Description, vbOKOnly + vbCritical
End Sub


Private Sub Command65_Click()

On Error GoTo ErrHandler

    Me.txtFinalRationale.SetFocus
    DoCmd.RunCommand acCmdSpelling
Exit Sub

ErrHandler:
    MsgBox "Spelling Check Failed - " & Err.Description, vbOKOnly + vbCritical


End Sub

Private Sub Form_Close()
    SaveRationale
    'On Error Resume Next
    'RemoveObjectInstance Me
    RaiseEvent FormClosed
End Sub

Private Sub Form_Load()
'

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRationalezoom = Nothing
    Set frmObsDisplay = Nothing
    RaiseEvent FormClosed
End Sub

Private Sub frmRationalezoom_FormClosed()
    Set frmRationalezoom = Nothing
End Sub
Private Sub frmObsDisplay_FormClosed()
    Set frmObsDisplay = Nothing
End Sub

Private Sub frmRationalezoom_TextConfirmed(strText As String, strControlName As String, bCancel As Boolean)
    
    If bCancel = False Then
        Me.Controls(strControlName) = strText
    End If


End Sub

Private Sub lstSavedRationale_DblClick(Cancel As Integer)
    If Nz(Me.lstSavedRationale, "") <> "" Then
    Me.txtBoxLoadSaved = 1
    LoadSavedRationale Me.lstSavedRationale
    Me.cmbBoxRationaleID = Me.txtBoxRationaleID
    Me.txtBoxLoadSaved = 0
    Me.TabCtl54.Pages.Item(1).SetFocus
    'BEGIN 11/8/2012 KCF - to select all text so that right-click, copy will pick up all text
    Me.txtFinalRationale.SetFocus
    Me.txtFinalRationale.SelStart = 0
    Me.txtFinalRationale.SelLength = Len(Me.txtFinalRationale.Value)
    'END 11/8/2012 KCF - to select all text so that right-click, copy will pick up all text
    End If
End Sub

Private Sub txtFieldData1_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData1.SetFocus
    ZoomText Me.txtFieldData1, Me.LblField1.Caption, False, Me.txtFieldData1.Name
End Sub
Private Sub ZoomText(strText As String, strlabel As String, bLocked As Boolean, strControlName As String)
   
  On Error GoTo ErrHandler
    
    'Dim frmPopup As Form
   
    Dim intControlIndex As Integer
    If IsNumeric(Right(strControlName, 2)) Then
        intControlIndex = CInt(Right(strControlName, 2))
    Else
        intControlIndex = CInt(Right(strControlName, 1))
    End If
   
   If frmRationalezoom Is Nothing Then
        Set frmRationalezoom = New Form_frm_AUDITCLM_Rationale_Zoom
    '    ColObjectInstances.Add Item:=frmRationalezoom, Key:=frmRationalezoom.hwnd & " "
        frmRationalezoom.CnlyClaimNum = Me.txtBoxCnlyClaimNum
        frmRationalezoom.TextData = Me.Controls("txtFieldData" & Trim(str(intControlIndex)))
        frmRationalezoom.TextLabel = strlabel
        frmRationalezoom.TextTemplate = Me.Controls("txtFieldtext" & Trim(str(intControlIndex)))
        frmRationalezoom.TextPrompt = Me.Controls("txtFieldPrompt" & Trim(str(intControlIndex))) ' andrew added 08-31-2012
        frmRationalezoom.ControlName = strControlName
        frmRationalezoom.Locked = bLocked
        frmRationalezoom.RefreshData
        frmRationalezoom.visible = True
'        ShowFormAndWait frmRationalezoom
'        frmRationalezoom.visible = True
    Else
        frmRationalezoom.SetFocus
    End If

ErrHandler_Exit:
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ErrHandler_Exit
    Set frmRationalezoom = Nothing

End Sub

Private Sub cboFieldData1_DblClick(Cancel As Integer)
    On Error GoTo ErrHandler
    
    Me.txtFieldText1.SetFocus
    Me.cboFieldData1.SetFocus
    
    If frmObsDisplay Is Nothing Then
    Set frmObsDisplay = New Form_frm_GENERAL_Display
    frmObsDisplay.TextPrompt = Me.Controls("txtFieldPrompt11")
    frmObsDisplay.RefreshData
        frmObsDisplay.visible = True
    Else
        frmObsDisplay.SetFocus
    End If
    
ErrHandler_Exit:
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    Resume ErrHandler_Exit
    Set frmObsDisplay = Nothing
    
End Sub
Private Sub txtFieldData10_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData10.SetFocus
    ZoomText Me.txtFieldData10, Me.LblField10.Caption, False, Me.txtFieldData10.Name
End Sub


Private Sub txtFieldData11_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData11.SetFocus
    ZoomText Me.txtFieldData11, Me.LblField11.Caption, False, Me.txtFieldData11.Name
End Sub


Private Sub txtFieldData12_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData12.SetFocus
    ZoomText Me.txtFieldData12, Me.LblField12.Caption, False, Me.txtFieldData12.Name
End Sub
Private Sub txtFieldData13_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData13.SetFocus
    ZoomText Me.txtFieldData13, Me.LblField13.Caption, False, Me.txtFieldData13.Name
End Sub

Private Sub txtFieldData14_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData14.SetFocus
    ZoomText Me.txtFieldData14, Me.LblField14.Caption, False, Me.txtFieldData14.Name
End Sub

Private Sub txtFieldData15_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData15.SetFocus
    ZoomText Me.txtFieldData15, Me.LblField15.Caption, False, Me.txtFieldData15.Name
End Sub

Private Sub txtFieldData2_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData2.SetFocus
    ZoomText Me.txtFieldData2, Me.LblField2.Caption, False, Me.txtFieldData2.Name
End Sub

Private Sub txtFieldData3_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData3.SetFocus
    ZoomText Me.txtFieldData3, Me.LblField1.Caption, False, Me.txtFieldData3.Name
End Sub

Private Sub txtFieldData4_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData4.SetFocus
    ZoomText Me.txtFieldData4, Me.LblField4.Caption, False, Me.txtFieldData4.Name
End Sub
Private Sub txtFieldData5_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData5.SetFocus
    ZoomText Me.txtFieldData5, Me.LblField5.Caption, False, Me.txtFieldData5.Name
End Sub
Private Sub txtFieldData6_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData6.SetFocus
    ZoomText Me.txtFieldData6, Me.LblField6.Caption, False, Me.txtFieldData6.Name
End Sub
Private Sub txtFieldData7_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData7.SetFocus
    ZoomText Me.txtFieldData7, Me.LblField7.Caption, False, Me.txtFieldData7.Name
End Sub
Private Sub txtFieldData8_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData8.SetFocus
    ZoomText Me.txtFieldData8, Me.LblField8.Caption, False, Me.txtFieldData8.Name
End Sub
Private Sub txtFieldData9_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldData9.SetFocus
    ZoomText Me.txtFieldData9, Me.LblField9.Caption, False, Me.txtFieldData9.Name
End Sub
Private Sub txtFieldText1_DblClick(Cancel As Integer)
    Me.txtFieldText2.SetFocus
    Me.txtFieldText1.SetFocus
    ZoomText Me.txtFieldData1, Me.LblField1.Caption, False, Me.txtFieldData1.Name
End Sub
Private Sub txtFieldText10_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText10.SetFocus
    ZoomText Me.txtFieldData10, Me.LblField10.Caption, False, Me.txtFieldData10.Name
End Sub


Private Sub txtFieldText11_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText11.SetFocus
    ZoomText Me.txtFieldData11, Me.LblField11.Caption, False, Me.txtFieldData11.Name
End Sub


Private Sub txtFieldText12_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText12.SetFocus
    ZoomText Me.txtFieldData12, Me.LblField12.Caption, False, Me.txtFieldData12.Name
End Sub


Private Sub txtFieldText13_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText13.SetFocus
    ZoomText Me.txtFieldData13, Me.LblField13.Caption, False, Me.txtFieldData13.Name
End Sub


Private Sub txtFieldText2_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText2.SetFocus
    ZoomText Me.txtFieldData2, Me.LblField2.Caption, False, Me.txtFieldData2.Name
End Sub


Private Sub txtFieldText3_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText3.SetFocus
    ZoomText Me.txtFieldData3, Me.LblField3.Caption, False, Me.txtFieldData3.Name
End Sub


Private Sub txtFieldText4_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText4.SetFocus
    ZoomText Me.txtFieldData4, Me.LblField4.Caption, False, Me.txtFieldData4.Name
End Sub


Private Sub txtFieldText5_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText5.SetFocus
    ZoomText Me.txtFieldData5, Me.LblField5.Caption, False, Me.txtFieldData5.Name
End Sub


Private Sub txtFieldText6_Click()
    Me.txtFieldText1.SetFocus
    Me.txtFieldText6.SetFocus
    ZoomText Me.txtFieldData6, Me.LblField6.Caption, False, Me.txtFieldData6.Name
End Sub

Private Sub txtFieldText7_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText7.SetFocus
    ZoomText Me.txtFieldData7, Me.LblField7.Caption, False, Me.txtFieldData7.Name
End Sub

Private Sub txtFieldText8_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText8.SetFocus
    ZoomText Me.txtFieldData8, Me.LblField8.Caption, False, Me.txtFieldData8.Name
End Sub


Private Sub txtFieldText9_DblClick(Cancel As Integer)
    Me.txtFieldText1.SetFocus
    Me.txtFieldText9.SetFocus
    ZoomText Me.txtFieldData9, Me.LblField9.Caption, False, Me.txtFieldData9.Name
End Sub

Public Function Age(dteDOB As Date, Optional SpecDate As Variant) As Integer
    Dim dteBase As Date, intCurrent As Date, intEstAge As Integer
    If IsMissing(SpecDate) Then
        dteBase = Date
    Else
        dteBase = SpecDate
    End If
    intEstAge = DateDiff("yyyy", dteDOB, dteBase)
    intCurrent = DateSerial(Year(dteBase), Month(dteDOB), Day(dteDOB))
    Age = intEstAge + (dteBase < intCurrent)
End Function

Private Function GetNextSequence()
    Dim strSQL As String
    Dim rs As DAO.RecordSet
    Dim intSeqNo As Integer
    
On Error GoTo ErrHandler
    
    
    strSQL = " SELECT max(SeqNo) as MaxSeq FROM AuditClm_Rationale_Saved WHERE cnlyClaimNum = '" & Nz(Me.txtBoxCnlyClaimNum) & "'"
    
    Set rs = CurrentDb.OpenRecordSet(strSQL)
    
    If Not rs.EOF Then
        intSeqNo = Nz(rs!MaxSeq, 0) + 1
    Else
        intSeqNo = 1
    End If
GetNextSequence = intSeqNo

Exit Function

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "GetNextSequence"
    GetNextSequence = 1
End Function

Private Function GetRationaleID(intSeqNo As Integer)
    
    
On Error GoTo ErrHandler
    
    GetRationaleID = DLookup("RationaleID", "AuditClm_Rationale_Saved", "cnlyClaimNum = '" & Me.txtBoxCnlyClaimNum & "' AND Seqno=" & intSeqNo & " AND FieldID=1")
    
Exit Function

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "GetRationaleID"
    
End Function


Private Sub LoadSavedRationale(intSeqNo As Integer)

On Error GoTo ErrHandler

Dim StrSQLd As String
Dim rd As DAO.RecordSet

'''''''''''''''''''''''''''''''''''
If Me.txtBoxLoadSaved = 1 Then

    Me.txtBoxRationaleID = GetRationaleID(intSeqNo)
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim strSQL As String
    Dim lngPreviousTop As Long
    Dim intI As Integer
    Dim intComboCounter As Integer
    'Get the values for the template being applied to this form
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strSQL = "SELECT * from AuditClm_Rationale_Template where rationaleID = " & Me.txtBoxRationaleID
    Set rs = MyAdo.OpenRecordSet(strSQL)
    'Make sure it exists
    'Clear out the final rationale
    Me.txtFinalRationale = ""
    'Loop through the template fields
    lngPreviousTop = 1300
    For intI = 1 To rs.recordCount
      'Set the text box label
      Me.Controls("lblField" & Trim(str(rs!FieldID))).Caption = rs!FieldDisplay
      Me.Controls("lblField" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
      Me.Controls("lblField" & Trim(str(rs!FieldID))).top = lngPreviousTop
      'Set the base text to be used to construct this part of the rationale
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))) = rs!TextRationale
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
      Me.Controls("txtFieldText" & Trim(str(rs!FieldID))).top = lngPreviousTop
      'Set the data entry control's tag to contain the token that will be replaced in the rationale template
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).Tag = rs!token
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).Height = rs!FieldHeight * 1000
      'empty out the field to start fresh
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))) = ""
      Me.Controls("txtFieldData" & Trim(str(rs!FieldID))).top = lngPreviousTop
     'Set the value of the bottom of the last control to dynamically size them
      lngPreviousTop = Me.Controls("lblField" & Trim(str(rs!FieldID))).top + Me.Controls("lblField" & Trim(str(rs!FieldID))).Height + 50
  rs.MoveNext
    Next intI
    Me.cboFieldData1 = ""
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
StrSQLd = " select * from AuditClm_Rationale_Saved where cnlyClaimNUm = '" & Me.txtBoxCnlyClaimNum & "' and seqno = " & intSeqNo '& " AND RationaleID=" & Me.txtBoxRationaleID & ""
Set rd = CurrentDb.OpenRecordSet(StrSQLd)

    If Not (rd.EOF = True And rd.BOF = True) Then
        Me.txtFinalRationale = Nz(rd!Rationale, "")
    End If
    
    
While Not rd.EOF
    Me.Controls("txtFieldData" & Trim(str(rd!FieldID))) = Nz(rd!FieldValue)
    rd.MoveNext
Wend

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly, "LoadSavedRationale"
End Sub

Private Function Validation() As Boolean
'  This function will validate to make sure that the Observation field is populated before the Rationale is generated
Dim sMsg As String
sMsg = ""
        ' Make sure required fields have values:
 If Me.txtBoxRationaleID = 1 Or Me.txtBoxRationaleID = 2 Or Me.txtBoxRationaleID = 3 Then
 cboFieldData1.SetFocus
    If Me.cboFieldData1.ListIndex = -1 Then
        sMsg = "  Observation drop down box blank" & vbCrLf
    End If
End If
    
    If sMsg <> "" Then
        MsgBox "The following errors prevented the Rationale from being generated:" & vbCrLf & sMsg
        Validation = True
    End If

End Function
Public Function Yyear(DtIPDischargeDt As Date) As String
   Dim mMonth As Integer
    
                 mMonth = DatePart("m", DtIPDischargeDt)
    
                  If mMonth > 9 Then
                    Yyear = DatePart("yyyy", DtIPDischargeDt) & "-" & (DatePart("yyyy", DtIPDischargeDt) + 1)
                   Else
                    Yyear = (DatePart("yyyy", DtIPDischargeDt) - 1) & "-" & (DatePart("yyyy", DtIPDischargeDt))
                 End If
    
End Function


Private Sub txtFinalRationale_Change()
Me.txtCharCount = (8000 - Len(Me.txtFinalRationale.Text)) & " Available Characters Remaining"
Me.txtTotalChar = Len(Me.txtFinalRationale.Text)
End Sub


Private Sub txtFinalRationale_GotFocus()
Me.txtCharCount = (8000 - Len(Me.txtFinalRationale.Text)) & " Available Characters Remaining"
Me.txtTotalChar = Len(Me.txtFinalRationale.Text)
End Sub
