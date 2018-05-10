Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private cstrConceptId As String
Private cstrClientIssueId As String
Private cintStep As Integer
Private cblnPayerLevelDocSelected As Boolean
Private cstrFilePathSelected As String
Private cstrFileNewName As String

Private coCurConcept As clsConcept

Public Event AttachmentSelected(strAttachmentType As String, sCnlyDocTypeID As String)
Public Event TaggedClaimSelected(intEracTaggedClaimId As Integer)
Public Event NewNameOfFileGenerated(sNewFileName As String)
Public Event EracRequiredDocTypeFound(oReqdDocType As clsConceptReqDocType)
Public Event PayersSelected(sPayerNameIds As String, sPayerNames As String)


Private ccolPayers As Collection

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get FilePathSelected() As String
    
    FilePathSelected = cstrFilePathSelected
End Property
Public Property Let FilePathSelected(sFilePathSelected As String)
    cstrFilePathSelected = sFilePathSelected
    Call DisplaySelectedFileName(sFilePathSelected)
End Property

Public Property Get FileNewName() As String
    FileNewName = cstrFileNewName
End Property
Public Property Let FileNewName(sFileNewName As String)
    cstrFileNewName = sFileNewName
End Property



Public Property Get ConceptID() As String
    ConceptID = cstrConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    cstrConceptId = sConceptId
    
    Set ccolPayers = GetRelatedPayerNameIDs(sConceptId)
    Set coCurConcept = New clsConcept
    If coCurConcept.LoadFromId(sConceptId) = False Then
        Stop
    End If
    Call FilterPayerSubForm
End Property



Public Property Get ClientIssueId() As String
    ClientIssueId = cstrClientIssueId
End Property
Public Property Let ClientIssueId(sClientIssueId As String)
    cstrClientIssueId = sClientIssueId
End Property


Public Property Get PayerLevelDocSelected() As Boolean
    PayerLevelDocSelected = cblnPayerLevelDocSelected
End Property
Public Property Let PayerLevelDocSelected(bClaimLvlDocSelected As Boolean)
    cblnPayerLevelDocSelected = bClaimLvlDocSelected
End Property




Private Sub cmbAttachmentType_Change()
    DoCmd.Hourglass True
    Call AttachmentTypeWasPicked
    DoCmd.Hourglass False
End Sub



Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub cmdOk_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdOk_Click"
    
    If CStr("" & cmbAttachmentType) = "" Then
        MsgBox "Please select an attachment type first"
        GoTo Block_Exit
    Else
        If FinalizeSelection() = False Then
            LogMessage strProcName, "ERROR", "There was an error finalizing the selection!"
            GoTo Block_Exit
        End If
        
        RaiseEvent NewNameOfFileGenerated(FileNewName)
        RaiseEvent AttachmentSelected(CStr("" & cmbAttachmentType), Me.cmbAttachmentType.Column(3, cmbAttachmentType.ListIndex + 1))
    End If

    DoCmd.Close acForm, Me.Name

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
''' This function checks to see if a claim level document was selected, if so,
'''  this attempts to find a tagged claim with the ICN as the named file
'''  if not, user must select one from the drop down (otherwise, dropdown
'''  isn't even shown)
'''
Private Function AttachmentTypeWasPicked() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oEracDoc As clsConceptReqDocType
Dim oRs As ADODB.RecordSet
Dim SFileName As String
Dim sIcn As String
Dim bFoundClaim As Boolean
Dim bSubmitted As Boolean

    strProcName = ClassName & ".AttachmentTypeWasPicked"

    If cmbAttachmentType = "" Then GoTo Block_Exit

    Set oEracDoc = New clsConceptReqDocType
    If oEracDoc.LoadFromId(CLng("0" & (cmbAttachmentType.Column(0)))) = False Then
        LogMessage strProcName, "WARNING", "Could not load required doc type object", CStr(cmbAttachmentType.Column(0))
        GoTo Block_Exit
    End If

    ' 20120622: KD: If this is a Payer Type, then make sure they pick the payer type
    If ccolPayers.Count > 0 Then    '' There should only be 0 if it's an OLD concept that needs to be converted..
        If oEracDoc.IsPayerDoc = True Then
        
            '' KD COMEBACK: if there's only 1 payer, no need to have them select the payer.. do it for them.
            
        
            Me.sfrmPayers.visible = True
            MsgBox "Please select the associated payer(s) that this document can be sent to", vbOKOnly, "Which payer(s)?"
            GoTo Block_Exit
        Else
            Me.sfrmPayers.visible = False
        End If
    End If

Block_Exit:
    Set oEracDoc = Nothing
    Set oRs = Nothing
    Exit Function
Block_Err:
    Err.Clear
    AttachmentTypeWasPicked = False
    GoTo Block_Exit
End Function




'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
'''''' This function checks to see if a claim level document was selected, if so,
''''''  this attempts to find a tagged claim with the ICN as the named file
''''''  if not, user must select one from the drop down (otherwise, dropdown
''''''  isn't even shown)
''''''
'''Private Function AttachmentTypeWasPicked_LEGACY() As Boolean
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim oEracDoc As clsConceptReqDocType
'''Dim oRs As ADODB.Recordset
'''Dim SFileName As String
'''Dim sIcn As String
'''Dim bFoundClaim As Boolean
'''Dim bSubmitted As Boolean
'''
'''    strProcName = ClassName & ".AttachmentTypeWasPicked"
'''
'''    If cmbAttachmentType = "" Then GoTo Block_Exit
'''
'''    Set oEracDoc = New clsConceptReqDocType
'''    If oEracDoc.LoadFromID(CInt("0" & (cmbAttachmentType.Column(0)))) = False Then
'''        LogMessage strProcName, "WARNING", "Could not load required doc type object", CStr(cmbAttachmentType.Column(0))
'''        GoTo Block_Exit
'''    End If
'''
'''    ' 20120622: KD: If this is a Payer Type, then make sure they pick the payer type
'''    If oEracDoc.IsPayerDoc = True Then
'''        Me.sfrmPayers.visible = True
'''        MsgBox "Please select the associated payer(s) that this document can be sent to", vbOKOnly, "Which payer(s)?"
'''        GoTo Block_Exit
'''    End If
'''
'''    '' pass the required doc type back to our opening form (if they want it!)
'''    RaiseEvent EracRequiredDocTypeFound(oEracDoc)
'''
'''    '' Ok, we need to get the naming convention for the file:
'''    FileNewName = oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, , Me.FilePathSelected)
'''
'''
'''    '' Now we need to see:
'''    ' #1) if this is a Claim Level Document, if so, try to get tagged claim by the ICN (filename)
'''    '       but of course, if there are no tagged claims yet,
'''    '       then bring them in
'''Stop ' kd: didn't do this yet.
'''    If oEracDoc.IsPayerDoc = False Then
'''        '' not a claim level document being attached
'''        PayerLevelDocSelected = False
'''    Else
'''        PayerLevelDocSelected = True
'''
'''        ' do we need to get the tagged claims for this concept?
'''        ' if so, just do it..
'''        ' Note that we aren't going to automatically refresh them.. Just the initial get
'''        '' Ok, changed that: 03232012 KD:        Call GetTaggedClaimsIfNone(Me.ConceptID)
'''
'''        Call SynchTaggedClaims(Me.ConceptID)
'''
'''
'''        '' Ok, now, if we have some claims then prompt the user to select the related claim
'''        ' Question is, what do I need to show them in order for them to know what claim it
'''        ' is, they aren't going to know the claim num right?
'''        Set oRs = GetTaggedClaimsRS(Me.ConceptID)
'''
'''        Set Me.cmbTaggedClaims.Recordset = oRs
'''        Me.Repaint
'''
'''            '' Now, see if we can find the document
'''            ' Get the filename (no extension, no folders) - this is going to be the ICN
'''        If PathInfoFromPath(Me.FilePathSelected, SFileName) = 0 Then
'''            LogMessage strProcName, "WARNING", "There was a problem with the file - couldn't get the filename", SFileName
'''            GoTo Block_Exit
'''        End If
'''
'''        sIcn = Trim(SFileName)
'''
'''        '' Ok, we need to get the naming convention for the file:
'''        FileNewName = oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, sIcn, Me.lblSelectedFileName.Caption)
'''
'''            '' If no claims were found
'''        If HasData(oRs) = False Then
'''            LogMessage strProcName, "ERROR", "No tagged claims were found for this concept. Please tag some claims before attaching claim level documents", , True
'''            GoTo Block_Exit
'''        End If
'''
'''        Do While Not oRs.EOF
'''            If UCase(Trim("" & oRs("ICN").Value)) = sIcn Then
'''
'''                If SelectComboBoxItemFromText(Me.cmbTaggedClaims, sIcn, True) = -1 Then
'''                    ' Couldn't find it.
'''                    bFoundClaim = False
'''                    GoTo NextOne
'''                End If
'''                bSubmitted = IIf(Nz(oRs("SubmitDate"), "") = "", False, True)
'''                bFoundClaim = True
'''                Exit Do
'''            End If
'''NextOne:
'''            oRs.MoveNext
'''        Loop
'''
'''        '' Now, if we found one but it's already been submitted..
'''        If bSubmitted = True Then
'''            LogMessage strProcName, "ERROR", "The document for claim ICN: " & sIcn & " has already been submitted", , True
'''            AttachmentTypeWasPicked_LEGACY = False
'''            GoTo Block_Exit
'''        End If
'''
'''
'''        If bFoundClaim = False Then
'''            LogMessage strProcName, "WARNING", "Could not find a claim", sIcn
'''            lblAttachToClaim.visible = True
'''            Me.cmbTaggedClaims.visible = True
'''        End If
'''
'''
'''    End If
'''
'''    AttachmentTypeWasPicked_LEGACY = True
'''
'''Block_Exit:
'''    Set oEracDoc = Nothing
'''    Set oRs = Nothing
'''    Exit Function
'''Block_Err:
'''    Err.Clear
'''    AttachmentTypeWasPicked_LEGACY = False
'''    GoTo Block_Exit
'''End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' This function checks to see if a claim level document was selected, if so,
'''  this attempts to find a tagged claim with the ICN as the named file
'''  if not, user must select one from the drop down (otherwise, dropdown
'''  isn't even shown)
'''
Private Function FinalizeSelection() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim SFileName As String
Dim sIcn As String
Dim bFoundClaim As Boolean
Dim bSubmitted As Boolean
Dim sPayerNames As String
Dim sPayerNameIds As String
Dim oPayerFrm As Form_frm_PAYERNAMES
Dim sPayers() As String
Dim sPayerNs() As String
Dim iThisPayerId As Integer
Dim sThisPayerName As String
Dim i As Integer
Dim oEracDoc As clsConceptReqDocType

    strProcName = ClassName & ".AttachmentTypeWasPicked"

    If cmbAttachmentType = "" Then GoTo Block_Exit

    Set oEracDoc = New clsConceptReqDocType
    If oEracDoc.LoadFromId(CLng("0" & (cmbAttachmentType.Column(0)))) = False Then
        LogMessage strProcName, "WARNING", "Could not load required doc type object", CStr(cmbAttachmentType.Column(0))
        GoTo Block_Exit
    End If
    
    If oEracDoc.IsPayerDoc = True Then
            ' get the id's and the name(s)
        Set oPayerFrm = Me.sfrmPayers.Form
            ' maybe have to load it.. hmm well, it'll throw an error
            ' set oPayerFrm = Forms("frm_PAYERNAMES")
        
        sPayerNameIds = oPayerFrm.GetSelectedPayerNameIDs(sPayerNames)
 
        If sPayerNameIds = "" Then
            ' no payers selected.. Is this ok? yes (as long as it's an old concept
'            Stop
        End If
        
        RaiseEvent PayersSelected(sPayerNameIds, sPayerNames)
    Else
        RaiseEvent PayersSelected("", "")
    End If

    '' pass the required doc type back to our opening form (if they want it!)
    RaiseEvent EracRequiredDocTypeFound(oEracDoc)

    '' Ok, we need to get the naming convention for the file:
    '' which we store in a property for later retrieval
'    FileNewName = oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, , Me.FilePathSelected, "")    '' hmm. ridiculous huh?

        ' Not going to parse it yet..
    FileNewName = GetFileName(Me.FilePathSelected)
'    FileNewName = Replace(FileNewName, "." & FileExtension(FileNewName), "")

        
        '' Now we need to see:
        ' #1) if this is a Claim Level Document, if so, try to get tagged claim by the ICN (filename)
        '       but of course, if there are no tagged claims yet,
        '       then bring them in
 
    PayerLevelDocSelected = oEracDoc.IsPayerDoc
    
        ' do we need to get the tagged claims for this concept?
        ' if so, just do it..
        ' Note that we aren't going to automatically refresh them.. Just the initial get
        '' Ok, changed that: 03232012 KD:        Call GetTaggedClaimsIfNone(Me.ConceptID)
   

    sPayers = Split(sPayerNameIds, ",")
    sPayerNs = Split(sPayerNames, ",")
    

'''    For i = 0 To UBound(sPayers)
'''        iThisPayerId = CInt(sPayers(i))
'''        sThisPayerName = sPayerNs(i)
'''
'''
'''        If PayerLevelDocSelected = False Then
'''            '' Ok, we need to get the naming convention for the file:
'''            If i = 0 Then
'''                FileNewName = oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, sICN, Me.lblSelectedFileName.Caption, sThisPayerName)
'''            Else    ' multiple, so keep it in a comma separated list
'''                FileNewName = FileNewName & "," & oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, sICN, Me.lblSelectedFileName.Caption, sThisPayerName)
'''            End If
'''        Else
'''            Call SynchTaggedClaims(Me.ConceptID, , sThisPayerName)
'''
'''
'''            Set oRs = GetTaggedClaimsRS(Me.ConceptID, iThisPayerId)
'''
'''            Set Me.cmbTaggedClaims.Recordset = oRs
'''            Me.Repaint
'''
'''                '' Now, see if we can find the document
'''                ' Get the filename (no extension, no folders) - this is going to be the ICN
'''            If PathInfoFromPath(Me.FilePathSelected, SFileName) = 0 Then
'''                LogMessage strProcName, "WARNING", "There was a problem with the file - couldn't get the filename", SFileName
'''                GoTo Block_Exit
'''            End If
'''
'''            sICN = Trim(SFileName)
'''
'''            '' Ok, we need to get the naming convention for the file:
'''            If i = 0 Then
'''                FileNewName = oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, sICN, Me.lblSelectedFileName.Caption, sThisPayerName)
'''            Else    ' multiple, so keep it in a comma separated list
'''                FileNewName = FileNewName & "," & oEracDoc.ParseFileName(Me.ConceptID, Me.ClientIssueId, sICN, Me.lblSelectedFileName.Caption, sThisPayerName)
'''            End If
'''
'''                '' If no claims were found
'''            If HasData(oRs) = False Then
'''                LogMessage strProcName, "ERROR", "No tagged claims were found for this concept. Please tag some claims before attaching claim level documents", , True
'''                GoTo Block_Exit
'''            End If
'''
'''            Do While Not oRs.EOF
'''                If UCase(Trim("" & oRs("ICN").Value)) = sICN Then
'''
'''                    If SelectComboBoxItemFromText(Me.cmbTaggedClaims, sICN, True) = -1 Then
'''                        ' Couldn't find it.
'''                        bFoundClaim = False
'''                        GoTo NextOne
'''                    End If
'''                    bSubmitted = IIf(Nz(oRs("SubmitDate"), "") = "", False, True)
'''                    bFoundClaim = True
'''                    Exit Do
'''                End If
'''NextOne:
'''                oRs.MoveNext
'''            Loop
'''
'''            '' Now, if we found one but it's already been submitted..
'''            If bSubmitted = True Then
'''                LogMessage strProcName, "ERROR", "The document for claim ICN: " & sICN & " has already been submitted", , True
'''                FinalizeSelection = False
'''                GoTo Block_Exit
'''            End If
'''
'''
'''            If bFoundClaim = False Then
'''                LogMessage strProcName, "WARNING", "Could not find a claim", sICN
'''                lblAttachToClaim.visible = True
'''                Me.cmbTaggedClaims.visible = True
'''            End If
'''        End If
'''
'''    Next
    
    '' Ok, now, if we have some claims then prompt the user to select the related claim
    ' Question is, what do I need to show them in order for them to know what claim it
    ' is, they aren't going to know the claim num right?
    
'    End If
    
    FinalizeSelection = True

Block_Exit:
    Set oEracDoc = Nothing
    Set oRs = Nothing
    Exit Function
Block_Err:
    Err.Clear
    FinalizeSelection = False
    GoTo Block_Exit
End Function


Private Sub FilterPayerSubForm()
On Error GoTo Block_Err
Dim strProcName As String
Dim vPayerId As Variant
Dim sFltr As String

    strProcName = ClassName & ".FilterPayerSubForm"

    If ccolPayers.Count = 0 Then
        ' Old concept?
        Me.sfrmPayers.Form.filter = " 1 = 2 "
        Me.sfrmPayers.Form.FilterOn = True
        GoTo Block_Exit
    End If

    ' Make sure we have a 1000 for All in there

    For Each vPayerId In ccolPayers
        sFltr = sFltr & CStr(vPayerId) & ","
    Next

    sFltr = left(sFltr, Len(sFltr) - 1) ' remove final comma
    On Error Resume Next
    Me.sfrmPayers.Form.filter = " PayerNameId IN (" & sFltr & ")"
    Me.sfrmPayers.Form.FilterOn = True
    On Error GoTo 0
    
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
''' This will display just the filename of the selected file..
'''
Private Sub DisplaySelectedFileName(ByVal sFilePathSelected As String)
 On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".DisplaySelectedFileName"
    
    sFilePathSelected = Right(sFilePathSelected, InStr(1, StrReverse(sFilePathSelected), "\") - 1)
    
    Me.lblSelectedFileName.Caption = sFilePathSelected
    

    
Block_Exit:
    Exit Sub
Block_Err:
    Err.Clear
    GoTo Block_Exit
End Sub



Private Function HasData(oRs As ADODB.RecordSet) As Boolean
    If oRs Is Nothing Then Exit Function
    If oRs.recordCount < 1 Then Exit Function
    HasData = True
End Function

Private Sub Form_Close()
    On Error Resume Next
'    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    Set ccolPayers = New Collection
End Sub
