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


Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    Me.txtSelectedId = sValue
    Call Me.RefreshData

Block_Exit:
    Exit Property
End Property


'' This and RefreshData is used to tie into the framework
'Property Get CnlyRowSource() As String
'     CnlyRowSource = cstrRowSource
'End Property
'Property Let CnlyRowSource(data As String)
'     cstrRowSource = data
'End Property

'' This and CnlyRowSource is used to tie into the framework
Public Sub RefreshData()
Dim strError As String
On Error GoTo ErrHandler
Dim sRelatedPayers As String


    sRelatedPayers = GetRelatedPayerNameIDsForFilter(Me.IdValue)

    Me.cmbPayerToCopyFrom.RowSource = "SELECT XREF_PAYERNAMES.PayerNameId, XREF_PAYERNAMES.PayerName FROM XREF_PAYERNAMES WHERE PayerNameId IN (" & sRelatedPayers & ") ORDER BY PayerName"

    '' Need to ffilter OUT the payers that are already assigned to this concept for the subform..
    Me.sfrmPayers.Form.filter = " PayerNameID NOT IN (1000," & sRelatedPayers & ")"
    Me.sfrmPayers.Form.FilterOn = True
    

exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub



Private Sub cmdOk_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim iPayerIdToAdd As Integer
Dim sErrMsg As String
Dim oRs As DAO.RecordSet
Dim bCmsSelected As Boolean
Dim bNonCmsSelected As Boolean
Dim bCGSSelected As Boolean
Dim bNonCGSSelected As Boolean

    strProcName = ClassName & ".cmdOk_Click"

    ' Adding a payer to an established concept involves these steps
    ' add a row to the CONCEPT_Payer_Dtl table from
    
    ' loop through the payers and if they are selected, add 'em
    
    '' But make sure we don't have CMS along with anything else
    
    '' 20130529 KD: And finally, make sure that there is not Concept Status at the header level
    
    Set oRs = Me.sfrmPayers.Form.RecordsetClone
  
    oRs.MoveFirst
    While Not oRs.EOF
    
        If oRs("Selected").Value = True Then
            Select Case oRs("PayerNameID").Value
            Case 1000
                bCGSSelected = True
            Case 1011
                bCmsSelected = True
            Case Else
                bNonCGSSelected = True
                bNonCmsSelected = True
            End Select
        End If
        oRs.MoveNext
    Wend
        
    Dim sMsg As String
    
    If bCGSSelected = True And bNonCGSSelected = True Then
        sMsg = "CGS needs to be selected by itself!"
        MsgBox sMsg, vbOKOnly, sMsg
        GoTo Block_Exit
    End If
    
    If bCmsSelected = True And bCmsSelected = True Then
        sMsg = "CMS needs to be selected by itself. It's primarilly for Medical Necessity and DRG"
        MsgBox sMsg, vbOKOnly, sMsg
        GoTo Block_Exit
    End If
    
    
    oRs.MoveFirst
    While Not oRs.EOF
    
        If oRs("Selected").Value = True Then
            iPayerIdToAdd = oRs("PayerNameID").Value
            
            If Nz(Me.cmbPayerToCopyFrom.Value, 0) = 0 Then
                Stop
                If AddPayerToEstablishedConceptFromPayerCopy(Me.IdValue, iPayerIdToAdd, iPayerIdToAdd) Then
                    LogMessage strProcName, "ERROR", "It looks like there was a problem adding a payer to concept!", CStr(iPayerIdToAdd), False, CStr(Me.IdValue)
                End If
            Else
                Stop
                If AddPayerToEstablishedConceptFromPayerCopy(Me.IdValue, iPayerIdToAdd, Me.cmbPayerToCopyFrom.Value) = False Then
                    LogMessage strProcName, "ERROR", "It looks like there was a problem adding a payer to concept!", CStr(iPayerIdToAdd), False, CStr(Me.IdValue)
                End If
            
            End If
        End If
        oRs.MoveNext
    Wend
    
    Me.Parent.TabSelected = 1
    
    Call Me.Parent.RefreshData

Block_Exit:
    Set oRs = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub

Private Sub Command37_Click()
Stop
End Sub
