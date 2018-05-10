Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 03/28/2012
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
'''  - 03/28/2012 - Locked the synch stuff when the concept has been submitted
'''     already
'''  - 03/26/2012 - Changed it to only auto refresh once every 10 minutes unless the
'''         button is clicked (taking a bit too long)
'''     - also, fixed detail claims - they weren't being imported
'''  - 03/23/2012 - changed it to auto refresh upon load
'''     Need to also make sure it does NOT do so if the concept has been submitted
'''  - 03/21/2012 - Added IdValue to be populated with (in this case) the
'''    concept id we're dealing with
'''  - 03/20/2012 - Created class
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


Private cstrRowSource As String
Private csConceptId As String
Private gblnIsRunning As Boolean



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    Me.txtSelectedId = csConceptId
    
    If mod_Concept_Specific.WasConceptSubmitted(sValue) = True Then
        bNeedToSynch = False
        Me.cmdSynchTaggedClaims.Enabled = False
        GoTo Block_Exit
    End If

    ' Don't synch unless it's been > 10 minutes (or it's the first time)
    Set oSettings = New clsSettings
    sLastTimeSynched = oSettings.GetSetting(Me.IdValue & "_TagClaimSynchTime")
    If IsDate(sLastTimeSynched) Then
        If DateDiff("n", CDate(sLastTimeSynched), Now) > 10 Then
            bNeedToSynch = True
        End If
    Else
        ' this is the first time for this concept
        bNeedToSynch = True
    End If
    
    If bNeedToSynch = True Then
        ' notify user?
        oSettings.SetSetting Me.IdValue & "_TagClaimSynchTime", Format(Now(), "m/d/yyyy hh:nn:ss AM/PM")
        Call SynchTaggedClaims(csConceptId)
        Call Me.RefreshData
    End If



Block_Exit:
    Exit Property
End Property


'' This and RefreshData is used to tie into the framework
Property Get CnlyRowSource() As String
     CnlyRowSource = cstrRowSource
End Property
Property Let CnlyRowSource(data As String)
     cstrRowSource = data
End Property


Public Sub PayerChange()
    cmbPayer_Change
End Sub


'' This and CnlyRowSource is used to tie into the framework
Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
'    Me.RecordSource = "SELECT CnlyTaggedClaimsByConcept.eRacTaggedClaimId, CnlyTaggedClaimsByConcept.PayerNameId, " & _
'        " XREF_PAYERNAMES.PayerName, CnlyTaggedClaimsByConcept.CnlyClaimNum, CnlyTaggedClaimsByConcept.ConceptId, " & _
'        " CnlyTaggedClaimsByConcept.ICN, IIf([HeaderLevelClaim]=1,True,False) AS IsHeaderLevel, " & _
'        " IIf([DetailLevelClaim]=1,True,False) AS IsDetailLevelClaim " & _
'        " FROM CnlyTaggedClaimsByConcept LEFT JOIN XREF_PAYERNAMES " & _
'        " ON CnlyTaggedClaimsByConcept.PayerNameId = XREF_PAYERNAMES.PayerNameId " & _
'        " WHERE CnlyTaggedClaimsByConcept.ConceptId = '" & csConceptId & "' " & _
'        " ORDER BY XREF_PAYERNAMES.PayerName "
    Me.RecordSource = "SELECT * FROM v_CONCEPT_TaggedClaims WHERE ConceptId = '" & csConceptId & "' ORDER BY PayerName "
    'Refresh the grid based on the rowsource passed into the form
'    Me.RecordSource = CnlyRowSource
exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub



Private Sub cmbPayer_Change()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmbPayer_Change"
    
        '' Need to filter or unfilter tagged claims
    
    If cmbPayer.Value = 1000 Then
        ' No filter:
        Me.filter = ""
        Me.FilterOn = False
    Else
        Me.filter = "PayerNameId = " & CStr(cmbPayer.Value)
        Me.FilterOn = True
    End If
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub cmdSynchTaggedClaims_Click()
    Call SynchTaggedClaims(Me.IdValue)
    Call Me.RefreshData
End Sub








Private Sub Form_Load()
Dim sPayers As String

  
    sPayers = GetRelatedPayerNameIDsForFilter(Me.Parent.Form.txtConceptID)
            
'    lblPayersNote.Caption = GetRelatedPayerNamesStr(CStr("" & Me.Parent.Form.txtConceptId))
'    If Trim(lblPayersNote.Caption) = "" Then
'        lblPayersNote.Caption = "This concept does not appear to be a payer specific Concept"
'    End If
    If sPayers <> "" Then
        sPayers = "1000," & sPayers

        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID IN (" & sPayers & ") ORDER BY PayerName"
    
    Else
        ' stop.. what should we do for the source here.. I don't think we should allow them to do anything
        Me.cmbPayer.RowSource = " SELECT PayerNameID, PayerName FROM XREF_PAYERNAMES WHERE PayerNameID = 1000 "
    End If
            
End Sub

Private Sub lblCnlyClaimNum_Click()
    Call SortForm("CnlyClaimNum")
End Sub

Private Sub lblConceptID_Click()
    Call SortForm("PayerName")
End Sub


Private Sub SortForm(sControlName As String)
Dim sAscOrDesc As String
    
    If left(Me.OrderBy, Len(sControlName)) = sControlName Then
        If Me.OrderBy Like "*DESC" Then
            sAscOrDesc = "" ' nothing by default is ASC
        Else
            sAscOrDesc = " DESC"
        End If
    End If
    Me.OrderBy = sControlName & sAscOrDesc
    Me.OrderByOn = True

    Me.Requery
    
End Sub

Private Sub lblDetailLvlClaim_Click()
    Call SortForm("IsHeaderLevel")
End Sub

Private Sub lblHeaderLvlClaim_Click()
    Call SortForm("IsHeaderLevel")
End Sub

Private Sub lblICN_Click()
    Call SortForm("ICN")
End Sub


Private Sub start()
    gblnIsRunning = True
    DoCmd.Hourglass True
    DoCmd.SetWarnings False
    'DoCmd.Echo True, ""
End Sub
Private Sub Done()
    gblnIsRunning = False
    DoCmd.Hourglass False
    DoCmd.SetWarnings True
    DoCmd.Echo True, "Ready..."
    DoEvents
End Sub
