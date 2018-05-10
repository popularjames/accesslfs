Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 10/16/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a CMS Concept at the payer level, basically a "hook" into the
'''     _CLAIMS.dbo.CONCEPT_PAYER_Dtl table
'''  With validation and various other methods..
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 10/16/2012 - added effectivedate and end date
'''  - 09/21/2012 - added PayerStatus properties..
'''  - 06/18/2012 - Created class
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

Public Event ConceptPayerError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean

'Private Const cstr_CONCEPT_ROOT_FOLDER As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\ConceptID\"
'Private Const cstr_CONCEPT_WORK_FOLDER As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\"

Private Const csIDFIELDNAME As String = "ConceptIDPayerID_RowID"
Private Const csTableName As String = "CONCEPT_PAYER_Dtl"
Private coSourceTable As clsTable

    '' The table to use for the connection string to the _ERAC database
Private Const csSP_TABLENAME As String = "ConceptDocTypes"

Private Const csREQUIRED_CONCEPT_PAYER_DTL_FIELDS As String = "ConceptId,PayerNameId,ConceptIdPayerId_RowId,ConceptStatus"

Private coReqRule As clsEracRequirementRule
Private ccolAttachedDocs As Collection

Private ccolTaggedClaims As Collection

Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private coParentConcept As clsConcept

Private csConceptId As String
Private csCnlyReviewTypeCode As String
Private csCnlyDataTypeCode As String
Private ciEracReviewTypeId As Integer
Private clConceptIDPayerNameId_RowId As Long
Private ciPayerNameId As Integer

Private cdctPayerNamesById As Scripting.Dictionary

Private cdtPayerEffectiveDate As Date
Private cdtPayerEndDate As Date

Private coValidateRpt As clsEracValidationRpt


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get Concept() As clsConcept
    If coParentConcept Is Nothing Then
        Set coParentConcept = New clsConcept
        Call coParentConcept.LoadFromId(Me.ConceptID)
    End If
    Set Concept = coParentConcept
End Property
Public Property Let Concept(oParentConcept As clsConcept)
    Set coParentConcept = oParentConcept
End Property


Public Property Get ConceptIDPayerNameId_RowId() As Long
    ConceptIDPayerNameId_RowId = clConceptIDPayerNameId_RowId
End Property
Public Property Let ConceptIDPayerNameId_RowId(iCIDPNId_RID As Long)
    clConceptIDPayerNameId_RowId = iCIDPNId_RID
End Property

Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As Long
        ID = ConceptIDPayerNameId_RowId
    End Property
    Public Property Let ID(lNewId As Long)
        ConceptIDPayerNameId_RowId = lNewId
    End Property

Public Property Get PayerNameId() As Integer
    PayerNameId = ciPayerNameId
End Property
Public Property Let PayerNameId(iPayerNameId As Integer)
    ciPayerNameId = iPayerNameId
End Property



Public Property Get PayerName() As String
    PayerName = CStr("" & cdctPayerNamesById.Item(CStr(ciPayerNameId)))
End Property



Public Property Get DateSubmitted() As Date
Dim sDate As String
    sDate = CStr("" & GetTableValue("DateSubmitted"))
    If IsDate(sDate) Then
        DateSubmitted = CDate(sDate)
    Else
        DateSubmitted = CDate("1/1/1900")
    End If
End Property
Public Property Let DateSubmitted(dSubmitDate As Date)
    SetTableValue "DateSubmitted", Format(dSubmitDate, "mm/dd/yyyy h:Nn:Ss AM/PM"), True
End Property


Public Property Get PayerStatusNum() As String
Dim sRet As String

    sRet = Me.GetField("ConceptStatus")
    
    If IsNumeric(sRet) = True Then
        PayerStatusNum = sRet
    End If
End Property


' Private dtPayerEffectiveDate As Date
' Private dtPayerEndDate As Date

Public Property Get EffectiveDate() As Date
    EffectiveDate = cdtPayerEffectiveDate
End Property
Public Property Let EffectiveDate(dEffectiveDate As Date)
    cdtPayerEffectiveDate = dEffectiveDate
End Property




Public Property Get EndDate() As Date
    EndDate = cdtPayerEndDate
End Property
Public Property Let EndDate(dEndDate As Date)
    cdtPayerEndDate = dEndDate
End Property


'Public Property Get RequirementRuleObj() As clsEracRequirementRule
'    Set RequirementRuleObj = coReqRule
'End Property



Public Property Get AttachedDocuments() As Collection
    Set AttachedDocuments = ccolAttachedDocs
End Property

Public Property Get TaggedClaims() As Collection
    Set TaggedClaims = ccolTaggedClaims
End Property


'
Public Property Get ClientIssueId() As String
    ClientIssueId = CStr("" & GetTableValue("ClientIssueNum"))
End Property
Public Property Let ClientIssueId(sClientIssueId As String)
    SetTableValue "ClientIssueNum", sClientIssueId, True
End Property

''##########################################################
''##########################################################
''##########################################################
'' Class state properties
''##########################################################
''##########################################################
''##########################################################
Public Property Get Dirty() As Boolean
    Dirty = cblnDirtyData
End Property
Public Property Let Dirty(blnDirtyData As Boolean)
    cblnDirtyData = blnDirtyData
    coSourceTable.Dirty = blnDirtyData
End Property


Public Property Get WasInitialized() As Boolean
    WasInitialized = cblnIsInitialized
End Property
Public Property Let WasInitialized(blnWasInit As Boolean)
    cblnIsInitialized = blnWasInit
End Property



''##########################################################
''##########################################################
''##########################################################
'' General properties
''##########################################################
''##########################################################
''##########################################################
'
Public Property Get GetField(sFieldName As String) As String
    GetField = CStr("" & GetTableValue(sFieldName))
End Property
Public Property Let DocName(sFieldName As String, sDocName As String)
    SetTableValue "DocName", sDocName
End Property


Public Function Fields() As Collection
    Set Fields = coSourceTable.Fields
End Function




''##########################################################
''##########################################################
''##########################################################
'' Business logic type functions
''##########################################################
''##########################################################
''##########################################################


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' How many claims are expected for this concept?
'''
Public Function RequiredClaimsNum(Optional ByRef sOutMsg As String) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".RequiredClaimsNum"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracNumOfClaimsRequired_Payer"
        .Parameters("@pConceptId") = Me.ConceptID
        .Parameters("@pPayerNameId") = Me.PayerNameId
        .Parameters("@pRequirementId") = Me.Concept.RequirementRuleObj.ID
        
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Problem finding required claims - look in " & .sqlString, "Req Rule ID: " & CStr(Me.Concept.RequirementRuleObj.ID) & " " & Me.ConceptID & " Payer: " & CStr(Me.PayerNameId)
            GoTo Block_Exit
        End If
    End With
    
    If Not oRs.EOF Then
        If Nz(oRs("ExceptionClaims"), -1) > -1 Then
                ' This is an exception..
            sOutMsg = "This is an exception, it normally requires " & CStr(oRs("NumClaimsPerConcept").Value) & _
                " claims, but is set (with Connolly) to have " & CStr(oRs("ExceptionClaims").Value)
            RequiredClaimsNum = oRs("ExceptionClaims").Value
        Else
            RequiredClaimsNum = oRs("NumClaimsPerConcept").Value
        End If
            ' Should only be 1 row, no need to movenext
    End If
    

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    RequiredClaimsNum = 10  '' default
    GoTo Block_Exit
End Function

'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
'''' sets the date and status to Concept Submitted in the EracConceptStatusLog
''''
'Public Function MarkAsSubmitted(Optional ByRef sErrMessage As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'
'    strProcName = ClassName & ".MarkAsSubmitted"
'Stop    '' modify to fit Payer table
'    ' bottom line, we need to set the date and status in the EracConceptStatusLog
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_EracSetConceptAsSubmitted"
'        .Parameters("@pConceptId") = Me.ConceptID
'        .Parameters("@pSubmitUser") = Identity.Username
'
'        If .Execute() = 0 Then
'            sErrMessage = "Update failed for unknown reason with : " & CStr(.CurrentConnection.Errors.Count) & " Ado errors"
'            MarkAsSubmitted = False
'            GoTo Block_Exit
'        End If
'
'    End With
'
'
'    MarkAsSubmitted = True
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    MarkAsSubmitted = False
'    GoTo Block_Exit
'End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
'Public Function SubmitDocPaths(Optional ByRef sOutMessage As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oReqDoc As clsConceptReqDocType
'Dim oAtchDoc As clsConceptDoc
'Dim sOneDoc As String
'Dim sRet As String
'
'    strProcName = ClassName & ".SubmitDocPaths"
'Stop ' what is this for?
'
'    For Each oAtchDoc In Me.AttachedDocuments
'                '' For now, we are only sending the package (hdr) level documents
'                '' the claim level (dtl lvl) docs will be burned to CD and sent separately
'        If oAtchDoc.GetEracReqDocType.IsHdrLvlDoc = True Then
'            sOneDoc = oAtchDoc.GetEracReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId, oAtchDoc.ICN, _
'                    oAtchDoc.FileName & oAtchDoc.FileName)
'            sOneDoc = ConceptWorkFolder & sOneDoc & "." & LCase(oAtchDoc.GetEracReqDocType.SendAsFileType)
'            sRet = sRet & Replace(sOneDoc, ConceptFolder, ConceptWorkFolder) & ","
'        Else
''            Stop
''            sRet = sRet & oAtchDoc.ConvertedFilePath
'        End If
'
'    Next
'        ' remove final comma
'    If Len(sRet) > 2 Then sRet = left(sRet, Len(sRet) - 1)
'    SubmitDocPaths = sRet
'
'Block_Exit:
'    Set oReqDoc = Nothing
'    Set oAtchDoc = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
'''' Has the concept already been submitted?
''''
'Public Function AlreadySubmitted(Optional ByRef sOutMessage As String) As Date
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim sSubmitUser As String
'
'    strProcName = ClassName & ".AlreadySubmitted"
'
'Stop    ' modify to fit  Payer_Dtl table
'
'    AlreadySubmitted = CDate("1/1/1900")
'
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_EracWasConceptSubmitted"
'        .Parameters.Refresh
'        .Parameters("@pConceptId") = Me.ConceptID
'        .Execute
'        If .Parameters("@pErrMsg").Value <> "" Then
'            GoTo Block_Exit
'        End If
'        AlreadySubmitted = .Parameters("@pSubmitDate").Value
'        sSubmitUser = .Parameters("@pSubmitUser").Value
'        sOutMessage = "Concept was already submitted on " & CStr(AlreadySubmitted) & " by " & sSubmitUser
'    End With
'
'Block_Exit:
'
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    GoTo Block_Exit
'End Function



'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function ValidateForSubmission(Optional ByRef oRs As ADODB.Recordset) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sReturn As String
'Dim oReqRule
'Dim dSubmitDate As Date
'Dim sOutMessage As String
'Dim sReport As String
'
'    strProcName = ClassName & ".ValidateForSubmission"
'Stop    ' modify to fit  Payer_Dtl table
'
'    ValidateForSubmission = True
'    Set coValidateRpt = New clsEracValidationRpt
'
'    If WasInitialized = False Then
'        sReport = "Concept ID not set"
'        ValidateForSubmission = False
'        GoTo Block_Exit
'    End If
'
'        '' eRAC is going away before I deploy this whole thing!
''' KD COMEBACK: THE NEW STORY IS:
'''  we are going to create it ourselves using the sproc I wrote: usp_CMS_Get_New_ClientIssueNum (_ERAC)
'
'    If Me.ClientIssueId = "" Then
'        If IssueClientIssueNum(Me, sReport) = "" Then
'            LogMessage strProcName, "ERROR", "There was a problem generating the Client Issue ID"
'            ValidateForSubmission = False
'            GoTo Block_Exit
'        End If
'
'    End If
'
'        '' Does it have all of the required fields in CONCEPT_hdr?
'    If Me.HasRequiredFields(sOutMessage) = False Then
'        sReport = sReport & "Missing required fields: " & vbCrLf & sOutMessage & vbCrLf
'        ValidateForSubmission = False
'        coValidateRpt.AddNote False, "Required fields", sOutMessage
'        sOutMessage = ""
'    Else
'        coValidateRpt.AddNote True, "Required fields", "All required fields have values"
'    End If
'
'
'        '' Ok, prompt the user if needed
'    If CheckTaggedClaimsAndPromptUserIfNeeded(ConceptID, sOutMessage) = False Then
'        LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
'        ValidateForSubmission = False
'        If Not coValidateRpt Is Nothing Then
'            coValidateRpt.AddNote False, "Number of expected tagged claims", sOutMessage
'        End If
'    Else
'        coValidateRpt.AddNote True, "Number of expected tagged claims", "Ok: all of the expected claims were found"
'    End If
'    sOutMessage = ""
'
'
'        '' Assuming we got here, we have the correct amount of tagged claims
'        '' we need to make sure that no tagged claim has already been submitted to another concept
'    If TaggedClaimsAlreadySubmittedToAnotherConcept(sOutMessage) = True Then
'        LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
'        ValidateForSubmission = False
'        If Not coValidateRpt Is Nothing Then
'            coValidateRpt.AddNote False, "Tagged claims already submitted", sOutMessage
'        End If
'    Else
'        coValidateRpt.AddNote True, "Tagged claims already submitted", "Ok: none of the tagged claims have been submitted to another concept"
'    End If
'    sOutMessage = ""
'
'
'        '' Does it have all of the required documents attached?
'    sOutMessage = Me.GetMissingRequiredDocsMessage()
'    If sOutMessage <> "" Then
'        sReport = sReport & "Missing required documents: " & sOutMessage
'
'        ValidateForSubmission = False
'            '        coValidateRpt.AddNote False, "Required documents", "Missing: " & sOutMessage
'    Else
'        coValidateRpt.AddNote True, "Required documents", "All required documents present"
'    End If
'    sOutMessage = ""
'
'        '' Has it been submitted yet?
'    dSubmitDate = Me.AlreadySubmitted(sOutMessage)
'    If dSubmitDate <> CDate("1/1/1900") Then
'        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'        ValidateForSubmission = False
'        coValidateRpt.AddNote False, "Concept already submitted", "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'    Else
'        coValidateRpt.AddNote True, "Concept already submitted", "OK: Not submitted yet"
'    End If
'    sOutMessage = ""
'
''        '' Is the notification status 3 (or 7)
''    If Me.IsStatusOkForEracSubmission(sOutMessage) = False Then
''        sReport = sReport & sOutMessage & vbCrLf
''        ValidateForSubmission = False
''        coValidateRpt.AddNote False, "Concept status ready 2 submit?", sOutMessage
''        sOutMessage = ""
''    Else
''        coValidateRpt.AddNote True, "Concept status ready 2 submit?", "eRAC status is ready for submission"
''    End If
'
'
'
'Block_Exit:
'    Set oRs = coValidateRpt.GetRecordset
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    ValidateForSubmission = False
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function PreviouslyPassedValidation(Optional ByRef sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oTClaim As clsEracClaim
'Dim oAdo As clsADO
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".PreviouslyPassedValidation"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = sqltext
'        .SQLstring = "SELECT * FROM V_EracConceptHistory WHERE ConceptID = '" & _
'                    Me.ConceptID & "' AND EracActionID = 18 AND ActionResult <> 'Failed' "
'        Set oRs = .ExecuteRS
'        If .GotData = True Then
'            PreviouslyPassedValidation = True
'        End If
'    End With
'
'
'Block_Exit:
'    Set oRs = Nothing
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username()
'    PreviouslyPassedValidation = False
'    GoTo Block_Exit
'End Function

'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function TaggedClaimsAlreadySubmittedToAnotherConcept(Optional ByRef sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oTClaim As clsEracClaim
'Dim oAdo As clsADO
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".TaggedClaimsAlreadySubmittedToAnotherConcept"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = sqltext
''        .SQLstring = "SELECT * FROM CnlyTaggedClaimsByConcept WHERE CnlyClaimNum = '" & _
'                    oTClaim.CnlyClaimNum & "' AND ConceptId = '" & Me.ConceptID & "'"
'    End With
'
'    For Each oTClaim In TaggedClaims
'
'        oAdo.SQLstring = "SELECT * FROM CnlyTaggedClaimsByConcept WHERE CnlyClaimNum = '" & _
'                    oTClaim.CnlyClaimNum & "' AND ConceptId <> '" & Me.ConceptID & "'"
'        Set oRs = oAdo.ExecuteRS
'        If oAdo.GotData = True Then
'                ' problem!!!
'            sReport = sReport & "Claim " & oTClaim.CnlyClaimNum & " was submitted with a different concept: " & oRs("ConceptId") & " already" & vbCrLf
'            If Not coValidateRpt Is Nothing Then
'                coValidateRpt.AddNote False, "Tagged Claim already submitted to different concept!", "Connolly Claim Num: " & oTClaim.CnlyClaimNum & " was submitted to " & oRs("ConceptId") & " already"
'            End If
'        End If
'    Next
'
'
'Block_Exit:
'    Set oRs = Nothing
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username()
'    TaggedClaimsAlreadySubmittedToAnotherConcept = False
'    GoTo Block_Exit
'End Function


'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
''''''
''''''
'''Public Function ValidateForClientIdRequest(Optional ByRef sReport As String) As Boolean
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim sOutMessage As String
'''Dim dSubmitDate As Date
'''
'''    strProcName = ClassName & ".ValidateForClientIdRequest"
'''    ValidateForClientIdRequest = True
'''    Set coValidateRpt = New clsEracValidationRpt
'''
'''    If WasInitialized = False Then
'''        sReport = "Concept ID not set"
'''        GoTo Block_Exit
'''    End If
'''
'''        '' Does it have the ClientIssueID
'''    If Me.ClientIssueId <> "" Then
'''        sReport = sReport & "Concept already has a Client Issue Id: " & Me.ClientIssueId & vbCrLf
'''        ValidateForClientIdRequest = False
'''    End If
'''
'''        '' Does it have all of the required fields in CONCEPT_hdr?
'''    If Me.HasRequiredFields(sOutMessage) = False Then
'''        sReport = sReport & "Missing required fields: " & vbCrLf & sOutMessage & vbCrLf
'''        ValidateForClientIdRequest = False
'''        sOutMessage = ""
'''    End If
'''
'''        '' Has it been submitted yet?
'''
'''    dSubmitDate = Me.AlreadySubmitted(sOutMessage)
'''    If dSubmitDate <> CDate("1/1/1900") Then
'''        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'''        ValidateForClientIdRequest = False
'''        sOutMessage = ""
'''    End If
'''
''''        '' Is the notification status 1??
''''    If Me.IsStatusOkForEracSubmission(sOutMessage) = False Then
''''        sReport = sReport & sOutMessage & vbCrLf
''''    End If
'''
'''Block_Exit:
'''    Exit Function
'''Block_Err:
'''    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'''    ValidateForClientIdRequest = False
'''    GoTo Block_Exit
'''End Function



'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Private Function GetHdrAttachedDocType(sReqDocType As String) As clsConceptDoc
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAttachedDoc As clsConceptDoc
'
'    strProcName = ClassName & ".GetHdrAttachedDocType"
'
'        ' These are only going to have 1 per concept.. (right??)
'    For Each oAttachedDoc In ccolHdrAttached
'        If oAttachedDoc.DocTypeName = sReqDocType Then
'            Set GetHdrAttachedDocType = oAttachedDoc
'            GoTo Block_Exit
'        End If
'    Next
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    GoTo Block_Exit
'End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function CountReqDocsOfType(oRequiredDocType As clsConceptReqDocType) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttachedDoc As clsConceptDoc
Dim iFoundCount As Integer


    strProcName = ClassName & ".CountHdrReqDocsOfType"

    For Each oAttachedDoc In ccolAttachedDocs
        If oAttachedDoc.GetEracReqDocType.CnlyAttachType = oRequiredDocType.CnlyAttachType Then
            iFoundCount = iFoundCount + 1
        End If
    Next

Block_Exit:
    CountReqDocsOfType = iFoundCount
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    GoTo Block_Exit
End Function




'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
''''
'Public Function HasRequiredDocsAttached(Optional ByRef sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sReturn As String
'Dim oReqDocType As clsConceptReqDocType
'Dim oAttachedDoc As clsConceptDoc
'Dim bAtLeastOneNotFound As Boolean
'Dim iFilesNeeded As Integer
'Dim iFilesFound As Integer
'Dim sMsg As String
'
'
'    strProcName = ClassName & ".HasRequiredDocsAttached"
'
'    For Each oReqDocType In coReqRule.RequiredDocs
'            ' If it's a header level doc (package level)
'        If oReqDocType.IsHdrLvlDoc = True Then
'            iFilesNeeded = oReqDocType.NumPerConcept
'        Else    ' must be a Detail level doc
'            ' we are supposed to have oReqDocType.NumPerClaim of these..
'            '' KD COMEBACK, need to change the below to NumPerClaim * Me.TaggedClaims.Count
'            iFilesNeeded = oReqDocType.NumPerClaim * Me.TaggedClaims.Count
'        End If
'
'        If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
'            ' we don't care if we have more - do we?
'            ' also, if we are going to make it, then we don't care..
'            If Nz(oReqDocType.CreateFunctionName, "") = "" Then
'                bAtLeastOneNotFound = True
'            End If
'        End If
'    Next
'    '' KD COMEBACK: Put something in sReport!
'    sReport = ""
'
'    If bAtLeastOneNotFound = True Then
'        ' Fire our error event
'        Dim oErr As ErrObject
'        Set oErr = New ErrObject
'        oErr.Description = ""
'        oErr.Number = 1234
'        oErr.Source = strProcName
'
'        FireError oErr, strProcName, sMsg
'    End If
'
'    HasRequiredDocsAttached = Not bAtLeastOneNotFound
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    HasRequiredDocsAttached = False
'    GoTo Block_Exit
'End Function



'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
'''' Set the number of claims are expected for this concept
''''
'Public Function SetRequiredClaimsNum(iNewAmount As Integer, ByRef sOutMsg As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'
'    strProcName = ClassName & ".RequiredClaimsNum"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_SetTaggedClaimNumException"
'        .Parameters("@pConceptId") = Me.ConceptID
'        .Parameters("@pClaimsToBeSubmitted") = iNewAmount
'
'        Call .Execute
'        sOutMsg = CStr("" & .Parameters("@pErrMsg").Value)
'
'        If sOutMsg <> "" Then
'            LogMessage strProcName, "ERROR", "Problem setting required claims - look in usp_SetTaggedClaimNumException for concept: " & Me.ConceptID, sOutMsg
'            GoTo Block_Exit
'        End If
'    End With
'
'    ' If we get here, we can assume success
'    SetRequiredClaimsNum = True
'
'Block_Exit:
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    SetRequiredClaimsNum = False
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
'''' How many claims are expected for this concept?
''''
'Public Function RequiredClaimsNum(Optional ByRef sOutMsg As String) As Integer
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".RequiredClaimsNum"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = StoredProc
'        .SQLstring = "usp_EracNumOfClaimsRequired"
'        .Parameters("@pConceptId") = Me.ConceptID
''        .Parameters("@pRequirementId") = RequirementRuleObj.ID
'
'        Set oRs = .ExecuteRS
'        If .GotData = False Then
''            LogMessage strProcName, "ERROR", "Problem finding required claims - look in usp_EracNumOfClaimsRequired", "Req Rule ID: " & CStr(RequirementRuleObj.ID) & " " & Me.ConceptID
'            GoTo Block_Exit
'        End If
'    End With
'
'    If Not oRs.EOF Then
'        If Nz(oRs("ExceptionClaims"), -1) > -1 Then
'                ' This is an exception..
'            sOutMsg = "This is an exception, it normally requires " & CStr(oRs("NumClaimsPerConcept").Value) & _
'                " claims, but is set (with Connolly) to have " & CStr(oRs("ExceptionClaims").Value)
'            RequiredClaimsNum = oRs("ExceptionClaims").Value
'        Else
'            RequiredClaimsNum = oRs("NumClaimsPerConcept").Value
'        End If
'            ' Should only be 1 row, no need to movenext
'    End If
'
'
'Block_Exit:
'    Set oRs = Nothing
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    RequiredClaimsNum = 10  '' default
'    GoTo Block_Exit
'End Function

'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function GetMissingRequiredDocsRS(Optional ByRef sReport As String) As ADODB.Recordset
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As ADODB.Recordset
'Dim sMsg As String
'Dim saryRows() As String
'Dim iRow As Integer
'
'    strProcName = ClassName & ".GetMissingRequiredDocsRS"
'
'    Set oRs = New ADODB.Recordset
'    oRs.ActiveConnection = Nothing
'    oRs.LockType = adLockBatchOptimistic
'    oRs.CursorLocation = adUseClient
'
'        ' Add our Field name
'    oRs.Fields.Append "Description", adLongVarChar
'
'    sMsg = Me.GetMissingRequiredDocsMessage()
'    If sMsg = "" Then
'        sMsg = "No documents missing!"
'    End If
'
'    saryRows = Split(sMsg, vbCrLf)
'
'    For iRow = 0 To UBound(saryRows)
'        oRs.AddNew
'        oRs.Fields(0) = saryRows(iRow)
'        oRs.Update
'    Next
'
'Block_Exit:
'    Set GetMissingRequiredDocsRS = oRs
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function HasRequiredFields(Optional ByRef sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim saryReqdFields() As String
Dim iIdx As Integer

    strProcName = ClassName & ".HasRequiredFields"
    HasRequiredFields = True
    saryReqdFields = Split(csREQUIRED_CONCEPT_PAYER_DTL_FIELDS, ",")

    For iIdx = 0 To UBound(saryReqdFields)
        If CStr("" & Me.GetField(saryReqdFields(iIdx))) = "" Then
            HasRequiredFields = False
            sReport = sReport & saryReqdFields(iIdx) & " is missing" & vbCrLf
        End If
    Next

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    HasRequiredFields = False
    GoTo Block_Exit
End Function

'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function GetMissingRequiredDocsMessage() As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sReturn As String
'Dim oReqDocType As clsConceptReqDocType
'Dim oAttachedDoc As clsConceptDoc
'Dim bAtLeastOneNotFound As Boolean
'Dim iFilesNeeded As Integer
'Dim iFilesFound As Integer
'Dim oTagdClaim As clsEracClaim
'Dim SFileName As String
'Dim sFolderPath As String
'Dim iEracClaimId As Integer
'Dim sMsg As String
'
'    strProcName = ClassName & ".GetMissingRequiredDocsMessage"
'
'        '' Ok, so, first, do we have all of the claims we need?
'    If ccolTaggedClaims.Count < Me.RequiredClaimsNum Then
'        sReturn = sReturn & CStr(Me.RequiredClaimsNum - ccolTaggedClaims.Count) & " claims are missing " & _
'                "(should be tagged, but aren't)" & vbCrLf
'        bAtLeastOneNotFound = True
'    End If
'
'    If coReqRule.RequiredDocs Is Nothing Then
'        sMsg = "No required documents"
'        GoTo Block_Exit
'    End If
'
'    For Each oReqDocType In coReqRule.RequiredDocs
'        iFilesNeeded = 0    ' Just make sure it's reset
'            ' If it's a header level doc (package level)
'
'            '' Where are we looking for the files?
'            '' for medical claims, we are doing that ourselves (code) so,
'            '' we could look in the work folder for that concept
''        If oReqDocType.CreateFunctionName <> "" Then     '' "Medical Record/Documentation" for example NIRF
''            sFolderPath = Me.ConceptWorkFolder
''        Else
''            sFolderPath = Me.ConceptFolder
''        End If
'
'            '' Is it a package level document or a claim level doc? each need to be treated differently
'        If oReqDocType.IsHdrLvlDoc = True Then
'            If ValidatePackageLevelDoc(oReqDocType, sReturn) = False Then
'                bAtLeastOneNotFound = True
'                coValidateRpt.AddNote False, "Required Package Document " & oReqDocType.DocName, oReqDocType.DocName & " not found"
'            End If
'
'        Else    ' must be a Claim level doc
'                ' we are supposed to have oReqDocType.NumPerClaim of these..
'
'            If ValidateClaimLevelDoc(oReqDocType, sReturn) = False Then
'                bAtLeastOneNotFound = True
'                coValidateRpt.AddNote False, "Required Claim Document " & oReqDocType.DocName, oReqDocType.DocName & " not found"
'
'            End If
'        End If
'
'
'        If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
'            ' we don't care if we have more - do we?
'            ' also, if we are going to make it, then we don't care..
'            If Nz(oReqDocType.CreateFunctionName, "") = "" Then
'                bAtLeastOneNotFound = True
''Stop
'                coValidateRpt.AddNote False, "Required Package Document " & oReqDocType.DocName, oReqDocType.DocName & " not found, dup?"
'            Else
''Stop
'            End If
'        End If
''NextRequiredDocType:
'    Next
''
''    If bAtLeastOneNotFound = True Then
''        ' Fire our error event
''        Dim oErr As ErrObject
''        Set oErr = New ErrObject
''        oErr.Description = ""
''        oErr.Number = 1234
''        oErr.Source = strProcName
''
''        FireError oErr, strProcName, sMsg
''    End If
'
'
'Block_Exit:
'    GetMissingRequiredDocsMessage = sReturn
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    sReturn = sReturn & "ERROR: " & Err.Description & vbCrLf
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Private Function ValidatePackageLevelDoc(oReqDocType As clsConceptReqDocType, sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim sReturn As String
'Dim oAttachedDoc As clsConceptDoc
'Dim bAtLeastOneNotFound As Boolean
'Dim iFilesNeeded As Integer
'Dim iFilesFound As Integer
'Dim oTagdClaim As clsEracClaim
'Dim SFileName As String
'Dim sFolderPath As String
'Dim iEracClaimId As Integer
'Dim sMsg As String
'Dim oAtchdDoc As clsConceptDoc
'Dim bFoundIt As Boolean
'
'    strProcName = ClassName & ".ValidatePackageLevelDoc"
'
'        ' If we are going to create it on submission,
'        ' then we need to see if we have the 'material' we'll need
'        ' to do so we use the 'CheckExistanceSQL' query
'
'    iFilesNeeded = oReqDocType.NumPerConcept
'
''    sFolderPath = Me.ConceptWorkFolder
'
'    If oReqDocType.CheckExistanceSQL <> "" Then
'        If AllDocExistsForConcept(oReqDocType, iFilesNeeded, sMsg) = False Then
'
'            sReport = sReport & sMsg & vbCrLf
'        End If
'    Else    'If oReqDocType.CreateFunctionName = "" Then
'
'            '' Since I'll be creating the NIRF, we need to add a check here to ignore files
'            '' that don't exist but have a create function
'        If iFilesNeeded > 0 Then
'            SFileName = oReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId, "")
'            If FileExists(sFolderPath & SFileName & "." & oReqDocType.SendAsFileType) = False And oReqDocType.CreateFunctionName = "" Then
''Stop
'                '' KD COMEBACK: if the converted document doesn't exist then look through the attached docs for it..
'
'                For Each oAtchdDoc In Me.AttachedDocuments
'                    If left(LCase(oAtchdDoc.FileName), Len(SFileName)) = LCase(SFileName) Then
'                        coValidateRpt.AddNote True, "Concept Documents", oReqDocType.DocName & " is present!"
'                        bFoundIt = True
'                        GoTo Block_Exit
'                    End If
'                Next
'
'                If bFoundIt = False Then
'                    sReport = sReport & oReqDocType.DocName & " is missing" & vbCrLf
'                    If Not coValidateRpt Is Nothing Then
'                        coValidateRpt.AddNote False, "Concept Documents", oReqDocType.DocName & " is missing"
'
'                        bAtLeastOneNotFound = True
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
'            ' we don't care if we have more - do we?
'            ' also, if we are going to make it, then we don't care..
'        If Nz(oReqDocType.CreateFunctionName, "") = "" Then
'            bAtLeastOneNotFound = True
''            If Not coValidateRpt Is Nothing Then
''                coValidateRpt.AddNote False, "Concept Documents", "Required " & CStr(iFilesNeeded) & " but only have " & CStr(CountReqDocsOfType(oReqDocType))
''            End If
'        End If
'    End If
'
'
'Block_Exit:
'    ValidatePackageLevelDoc = Not bAtLeastOneNotFound
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    sReport = sReport & "ERROR: " & Err.Description & vbCrLf
'    GoTo Block_Exit
'End Function


''''' ##############################################################################
''''' ##############################################################################
''''' ##############################################################################
'''''
'''''
''Private Function ValidateClaimLevelDoc(oReqDocType As clsConceptReqDocType, sReport As String) As Boolean
''On Error GoTo Block_Err
''Dim strProcName As String
''Dim oAttachedDoc As clsConceptDoc
''Dim bAtLeastOneNotFound As Boolean
''Dim iFilesNeeded As Integer
''Dim iFilesFound As Integer
''Dim oTagdClaim As clsEracClaim
''Dim SFileName As String
''Dim sFolderPath As String
''Dim iEracClaimId As Integer
''Dim sMsg As String
''
''    strProcName = ClassName & ".ValidateClaimLevelDoc"
''
''        ' If it's a claim level doc
''
''        '' Where are we looking for the files?
''        '' for medical claims, we are doing that ourselves (code) so,
''        '' we could look in the work folder for that concept
'''    If oReqDocType.CreateFunctionName <> "" Then     '' "Medical Record/Documentation" for example NIRF
'''        sFolderPath = Me.ConceptWorkFolder
'''    Else
'''        sFolderPath = Me.ConceptFolder
'''    End If
''
''
''        ''    iFilesNeeded = oReqDocType.NumPerClaim * Me.TaggedClaims.Count
''    '' This is a bug. We can't use the Me.TaggedClaims.Count because we may not have
''    '' all of the tagged claims that we need.. So
''
''    iFilesNeeded = oReqDocType.NumPerClaim * Me.RequiredClaimsNum
''
''        ' If we are creating them then we should check:
''    If oReqDocType.CreateFunctionName <> "" Then
''        ' b) if we have the 'Materials' we need to create them
''        ' KD COMEBACK
''        If AllDocExistsForConcept(oReqDocType, iFilesNeeded, sMsg) = False Then
''            bAtLeastOneNotFound = True
''            sReport = sReport & sMsg & vbCrLf
''
'''            If Not coValidateRpt Is Nothing Then
'''                coValidateRpt.AddNote False, "Claim Document", "At least 1 " & oReqDocType.DocName & " is missing for this concept!"
'''            End If
''        End If
''        ' a) if we've already created them - eh, who cares.. :P
''
''    Else
''
''        ' we are supposed to have oReqDocType.NumPerClaim of these..
''        ' Look through each of the claims, and see which ones don't have the current document type
''        '' KD COMEBACK: Ok, this is a good idea BUT
''
''        '' First, we have the NEW stuff which is going to have an eRacTaggedClaimId
''        For Each oTagdClaim In ccolTaggedClaims
''
''            iEracClaimId = oTagdClaim.eRacTaggedClaimId '  GetField("eRacTaggedClaimId")
''
''            If iEracClaimId > 0 Then    '' we have one.. it's a "New system" link
''                SFileName = GetAttachmentPathFromEracTgdClaimId(iEracClaimId)
''                If SFileName = "" Then
''                    ' KD COMEBACK Remove this
''                    sReport = sReport & "The " & oReqDocType.DocName & _
''                            " attached doc for claim: " & oTagdClaim.ICN & " is missing" & vbCrLf
''
''                    If Not coValidateRpt Is Nothing Then
''                        coValidateRpt.AddNote False, "Claim Document", "The " & oReqDocType.DocName & _
''                                " attached doc for claim: " & oTagdClaim.ICN & " is missing"
''                    End If
''
''                ElseIf FileExists(sFolderPath & SFileName) = False Then
''                    ' KD COMEBACK Remove this
''                    sReport = sReport & "The " & oReqDocType.DocName & " file for claim: " & _
''                            oTagdClaim.ICN & " is missing" & vbCrLf
''                    If Not coValidateRpt Is Nothing Then
''                        coValidateRpt.AddNote False, "Claim Document", "The " & oReqDocType.DocName & " file for claim: " & _
''                            oTagdClaim.ICN & " is missing"
''                    End If
''
''                End If
''            Else    '' It's an "old system" link LEGACY
''                    '' If it doesn't have the eRacTaggedClaimId so we need to see if the collection contains
''                    '' the doc by getting the parsed filename (without extension)
''                    '' then checking the attached docs collection
''                    '' for that filename (less the extension)
''                    '' not exactly precise
''                SFileName = oReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId, oTagdClaim.ICN)
''
''                SFileName = GetAttachmentPathFromParsedName(SFileName)
''
''                If SFileName = "" Then
''                    ' KD COMEBACK Remove this
''                    sReport = sReport & "The " & oReqDocType.DocName & " attached doc for claim: " & _
''                            oTagdClaim.ICN & " is missing" & vbCrLf
''
''                    If Not coValidateRpt Is Nothing Then
''                        coValidateRpt.AddNote False, "Claim Document", "The " & oReqDocType.DocName & " attached doc for claim: " & _
''                            oTagdClaim.ICN & " is missing"
''                    End If
''                ElseIf FileExists(sFolderPath & SFileName) = False Then
''                    ' KD COMEBACK Remove this
''                    sReport = sReport & "The " & oReqDocType.DocName & " file for claim: " & _
''                            oTagdClaim.ICN & " is missing" & vbCrLf
''                    If Not coValidateRpt Is Nothing Then
''                        coValidateRpt.AddNote False, "Claim Document", "The " & oReqDocType.DocName & " file for claim: " & _
''                            oTagdClaim.ICN & " is missing"
''                    End If
''                End If
''            End If
''
''        Next
''
''
''
''
''    End If
''
''
''    If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
''        ' we don't care if we have more - do we?
''        ' also, if we are going to make it, then we don't care..
''        If Nz(oReqDocType.CreateFunctionName, "") = "" Then
''            bAtLeastOneNotFound = True
''        End If
''    End If
''
''
''Block_Exit:
''    ValidateClaimLevelDoc = Not bAtLeastOneNotFound
''    Exit Function
''Block_Err:
''    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
''    sReport = sReport & "ERROR: " & Err.Description & vbCrLf
''    GoTo Block_Exit
''End Function



'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function GetAttachmentPathFromEracTgdClaimId(iEracTgdClaimId As Integer) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAttach As clsConceptDoc
'
'    strProcName = ClassName & ".GetAttachmentPathFromEracTgdClaimId"
'
'    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
'
'    For Each oAttach In ccolAttachedDocs
'        If oAttach.eRacTaggedClaimId = iEracTgdClaimId Then
'            GetAttachmentPathFromEracTgdClaimId = oAttach.RefFileName
'            Exit For
'        End If
'    Next
'
'
'Block_Exit:
'    Set oAttach = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function NIRF_Exists() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc

    strProcName = ClassName & ".NIRF_Exists"

    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit

    For Each oAttach In ccolAttachedDocs
        If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
            NIRF_Exists = True
            GoTo Block_Exit
        End If
    Next


Block_Exit:
    Set oAttach = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function NIRF_Path() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc

    strProcName = ClassName & ".NIRF_Path"

    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit

    For Each oAttach In ccolAttachedDocs
        If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
            NIRF_Path = oAttach.RefFullPath
Stop
            GoTo Block_Exit
        End If
    Next


Block_Exit:
    Set oAttach = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
'''' NOTE: THis is for LEGACY attached documents
'''' since they won't be named correctly
''''
'Public Function GetAttachmentPathFromParsedName(ByVal SFileName As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAttach As clsConceptDoc
'Dim iFileLen As Integer
'
'    strProcName = ClassName & ".GetAttachmentPathFromParsedName"
'        ' Insure no period and extension:
'    If InStr(1, SFileName, ".", vbTextCompare) > 0 Then
'        SFileName = left(SFileName, InStr(1, SFileName, ".", vbTextCompare) - 1)
'    End If
'        ' Now, since people save stuff like screen prints with JUST the ICN:
'    If InStr(1, SFileName, "_", vbTextCompare) > 0 Then
'        SFileName = left(SFileName, InStr(1, SFileName, "_", vbTextCompare) - 1)
'    End If
'
'    If SFileName = "" Then GoTo Block_Exit
'    SFileName = UCase(SFileName)
'    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
'    iFileLen = Len(SFileName)
'
'    For Each oAttach In ccolAttachedDocs
'            '' a little redundancy here..
'        If left(UCase(oAttach.RefFileNameNoExt), iFileLen) = SFileName Then
'            GetAttachmentPathFromParsedName = oAttach.RefFileName
'            Exit For
'        End If
'    Next
'
'Block_Exit:
'    Set oAttach = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function GetAttachmentPathFromICN(sICN As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAttach As clsConceptDoc
'
'    strProcName = ClassName & ".GetAttachmentPathFromICN"
'    If sICN = "" Then GoTo Block_Exit
'    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
'
'    For Each oAttach In ccolAttachedDocs
'        If oAttach.ICN = sICN Then
'            GetAttachmentPathFromICN = oAttach.RefFileName
'            Exit For
'        End If
'    Next
'
'Block_Exit:
'    Set oAttach = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GoTo Block_Exit
'End Function




'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
'''' Is the status 3 (or 7)
''''
'Private Function ConvertClaims_Code_To_EracCode(sFieldName As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim oRs As ADODB.Recordset
'Dim sVal As String
'
'    strProcName = ClassName & ".ConvertClaims_Code_To_EracCode"
'
'    sVal = CStr("" & GetField(Right(sFieldName, Len(sFieldName) - 5)))    ' - 5 because it begins with CNVT_ - short for convert
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = sqltext
'        .SQLstring = "SELECT ReviewTypeName As ReviewType FROM XrefReviewType WHERE CnlyReviewTypeCode = '" & sVal & "'"
'        Set oRs = .ExecuteRS
'        If .GotData = False Then
'            ConvertClaims_Code_To_EracCode = sVal
'            GoTo Block_Exit
'        End If
'        sVal = CStr("" & oRs(Right(sFieldName, Len(sFieldName) - 5)).Value)
'    End With
'
'
'    ConvertClaims_Code_To_EracCode = sVal
'Block_Exit:
'    Set oAdo = Nothing
'    Set oRs = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    ConvertClaims_Code_To_EracCode = ""
'    GoTo Block_Exit
'End Function

'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
'''' Is the status 3 (or 7)
''''
'Public Function IsStatusOkForEracSubmission(Optional ByRef sErrMsg As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".IsStatusOkForEracSubmission"
'
'    Set oRs = GetRecordsetSP("usp_EracGetNotificationHistDesc", "@pConceptId=" & Me.ConceptID)
'    If RSHasData(oRs) = False Then
'        sErrMsg = "No records found for this concept"
'        GoTo Block_Exit
'    End If
'
'        '' This will likely have more than 1 record but it's ordered in desc order
'        '' so we only need to look at the first to see the current notification status
'    If oRs("NotificationType").Value <> 3 And oRs("NotificationType").Value <> 7 Then
'        sErrMsg = "Current notification type is: " & CStr(oRs("NotificationType").Value) & " (" & Nz(oRs("NotificationTypeName"), "") & ")"
'        GoTo Block_Exit
'    End If
'
'    IsStatusOkForEracSubmission = True
'Block_Exit:
'    Set oRs = Nothing
'
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    IsStatusOkForEracSubmission = False
'    sErrMsg = sErrMsg & " " & Err.Description
'    GoTo Block_Exit
'End Function



'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function MakeWorkCopiesOfFiles() As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'
'    strProcName = ClassName & ".MakeWorkCopiesOfFiles"
'    '' KD COMEBACK: DO THIS!
'
'Block_Exit:
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    MakeWorkCopiesOfFiles = False
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function ParseStringForDetails(sInString As String) As String
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRegEx As RegExp
'Dim oMatches As MatchCollection
'Dim oMatch As Match
'Dim sValue As String
'
'
'    strProcName = ClassName & ".ParseStringForDetails"
'
'    Set oRegEx = New RegExp
'    oRegEx.IgnoreCase = True
'    oRegEx.Pattern = "\[\*([^\*\]]+)\*\]"
'    oRegEx.Global = True
'    oRegEx.MultiLine = True
'
'    Set oMatches = oRegEx.Execute(sInString)
'
'    If oMatches.Count = 0 Then
'        ParseStringForDetails = sInString
'        GoTo Block_Exit
'    End If
'
'    ParseStringForDetails = sInString
'
'    For Each oMatch In oMatches
'        If left(oMatch.SubMatches(0), 5) = "CNVT_" Then
'            ' The code needs to be converted..
'            sValue = ConvertClaims_Code_To_EracCode(oMatch.SubMatches(0))
'
'        Else
'            sValue = CStr("" & GetField(oMatch.SubMatches(0)))
'        End If
'
'
'        ParseStringForDetails = Replace(ParseStringForDetails, "[*" & oMatch.SubMatches(0) & "*]", sValue, , , vbTextCompare)
'    Next
'
'Block_Exit:
'    Set oMatch = Nothing
'    Set oMatches = Nothing
'    Set oRegEx = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    ParseStringForDetails = sInString
'    GoTo Block_Exit
'End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function AllDocExistsForConcept(oRequiredDoc As clsConceptReqDocType, iRequiredNum As Integer, sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As ADODB.Recordset
'Dim sSql As String
'
'    strProcName = ClassName & ".AllDocExistsForConcept"
'    AllDocExistsForConcept = True
'
'    If oRequiredDoc.CheckExistanceSQL = "" Then GoTo Block_Exit
'
'    sSql = Replace(oRequiredDoc.CheckExistanceSQL, "?", Me.ConceptID, , , vbTextCompare)
'
'    Set oRs = GetRecordset(sSql, "ConceptDocTypes")
'
'    If oRs Is Nothing Then
'        AllDocExistsForConcept = False
'        sReport = sReport & "No items found!" & vbCrLf
'        If Not coValidateRpt Is Nothing Then
'            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, "No " & oRequiredDoc.DocName & " documents found!"
'        End If
'        GoTo Block_Exit
'    End If
'    If oRs.RecordCount < 1 Then
'        sReport = sReport & "No items found!" & vbCrLf
'
'        If Not coValidateRpt Is Nothing Then
'            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, "No " & oRequiredDoc.DocName & " documents found!"
'        End If
'        AllDocExistsForConcept = False
'        GoTo Block_Exit
'    End If
'
'    '' Do we have as many as we were expecting - well, that's for another function isn't it?
'    '' and of course that doesn't belong in this object (heck, this function is debatable)
'    If oRs.RecordCount < iRequiredNum Then
'
'        sReport = sReport & CStr(Me.RequiredClaimsNum - oRs.RecordCount) & " items were MISSING (" & _
'            "found " & CStr(oRs.RecordCount) & ")"
'        If Not coValidateRpt Is Nothing Then
'            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, CStr(Me.RequiredClaimsNum - oRs.RecordCount) & " items were MISSING (" & _
'                    "found " & CStr(oRs.RecordCount) & ")"
'        End If
'        AllDocExistsForConcept = False
'    End If
'
'
'Block_Exit:
'    Set oRs = Nothing
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & Me.ConceptID
'    sReport = sReport & Err.Description & vbCrLf
'    AllDocExistsForConcept = False
'    GoTo Block_Exit
'End Function



''##########################################################
''##########################################################
''##########################################################
'' Auditing / Setup data / interacting with the cTable object
'' as well as any generically private methods
''##########################################################
''##########################################################
''##########################################################



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetRecordset(sSql As String, Optional sTableName As String = csTableName) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetRecordset"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(sTableName)
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With

    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    Set GetRecordset = oRs

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, sSql
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetRecordsetSP(sSpName As String, Optional sParamString As String = "", Optional sTableName As String = csSP_TABLENAME) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sParams() As String
Dim iIdx As Integer
Dim sPName As String
Dim sPVal As String

    strProcName = ClassName & ".GetRecordset"

    If sParamString <> "" Then
        sParams = Split(sParamString, ",")
    End If

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(sTableName)
        .SQLTextType = StoredProc
        .sqlString = sSpName
        If sParamString <> "" Then
            For iIdx = 0 To UBound(sParams)
                sPName = Split(sParams(iIdx), "=")(0)
                sPVal = Split(sParams(iIdx), "=")(1)
                .Parameters(sPName) = sPVal
            Next
        End If
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With

    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    Set GetRecordsetSP = oRs

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, sSpName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetTableValue(strFieldName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetTableValue"
'    If coSourceTable.GetTableValue(strFieldName) Is Nothing Then
'        GetTableValue = ""
'    Else
        GetTableValue = Nz(coSourceTable.GetTableValue(strFieldName), "")
'    End If

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & strFieldName
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function SetTableValue(strFieldName As String, varValue As Variant, Optional blnSaveNow As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim blnWorked As Boolean

    strProcName = ClassName & ".SetTableValue"
    SetTableValue = True

    If WasInitialized = False Then
        SetTableValue = False
        GoTo Block_Exit
    End If

    blnWorked = coSourceTable.SetTableValue(strFieldName, varValue, , blnSaveNow)
    If blnWorked = True And blnSaveNow = False Then
        Dirty = True
    End If
    SetTableValue = blnWorked

Block_Exit:
    Exit Function

Block_Err:
    SetTableValue = False
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & strFieldName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' SaveNow (duplicate of Save...)
Public Function SaveNow() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sFieldName As String
Dim iFieldLoop As Integer

    strProcName = ClassName & ".SaveNow"

    If Dirty = False Or WasInitialized = False Then
        SaveNow = False
        GoTo Block_Exit
    End If

    SaveNow = coSourceTable.SaveNow()

    Dirty = Not SaveNow

    SaveNow = False
Block_Exit:

    Exit Function

Block_Err:
    SaveNow = False
    FireError Err, strProcName, "User ID: " & Identity.UserName()
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromId(lConceptIdPayerId_RowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
    ID = lConceptIdPayerId_RowId
    LoadFromId = coSourceTable.LoadFromId(lConceptIdPayerId_RowId)
    WasInitialized = LoadFromId

    Me.ConceptID = GetTableValue("ConceptID")
    Me.PayerNameId = CInt(GetTableValue("PayerNameId"))
    

        ' Get the requirement rule object...
'    If GetReqRule() = False Then
'        ' KD COMEBACK: Deal with this
'    End If

        ' stuff the attached documents into a collection
    If LoadAttachedDocs() = False Then
        ' it's already been logged.. just let it continue as no attached docs isn't a problem
        ' especially for a new concept
    End If

        ' get some basic details of the tagged claims..
    If LoadTaggedClaims() = False Then
        ' it's already been logged, so we'll let this one continue too as not all concepts
        '' will have tagged claims - and especially when they are new
    End If

    Call GetPayerDates
    

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "ID: " & CStr(lConceptIdPayerId_RowId)
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetPayerDates() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim lConceptidPayerIdRowId As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetPayerDates"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT PayerNameID, PayerName, ExcludeFromAll, EffectiveDate, EndDate, ForUserDisplay FROM XREF_Payernames WHERE PayerNameID = " & CStr(Me.PayerNameId)
        Set oRs = .ExecuteRS
        
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Problem getting payer effective dates!", CStr(Me.PayerNameId)
            GoTo Block_Exit
        End If
    End With
    
    If oRs("EffectiveDate").Value = CDate("1/1/1900") Then
        Me.EffectiveDate = CDate("1/2/1900")
    Else
        Me.EffectiveDate = oRs("EffectiveDate").Value
    End If
    
    If oRs("EndDate").Value = CDate("12/31/2999") Then
        Me.EndDate = CDate("12/30/2199")
    Else
        Me.EndDate = oRs("EndDate").Value
    End If
    
    
    GetPayerDates = True
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function

Block_Err:
    GetPayerDates = False
    FireError Err, strProcName, "ID: " & CStr(Me.PayerNameId)
    GoTo Block_Exit
End Function


'GetPayerDates
''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromConceptNPayer(sConceptId As String, lPayerNameId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim lConceptidPayerIdRowId As Long
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet


    strProcName = ClassName & ".LoadFromConceptNPayer"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_DATA_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT TOP 1 ConceptIDPayerID_RowID FROM Concept_Payer_Dtl WHERE ConceptID = '" & sConceptId & "' AND PayerNameID = " & CStr(lPayerNameId)
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Could not get the identifying record for this one!!!"
            GoTo Block_Exit
        End If
        
        lConceptidPayerIdRowId = oRs("ConceptIDPayerID_RowID").Value
    End With
    
    LoadFromConceptNPayer = LoadFromId(lConceptidPayerIdRowId)
    Call GetPayerDates

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function

Block_Err:
    LoadFromConceptNPayer = False
    FireError Err, strProcName, "ConceptID: " & sConceptId & " Payer ID: " & CStr(lPayerNameId)
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function LoadAttachedDocs() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oAttachedFile As clsConceptDoc

    strProcName = ClassName & ".LoadAttachedDocs"

    sSql = "SELECT * FROM v_CONCEPT_References WHERE ConceptID = '" & Me.ID & "' ORDER BY RefSequence ASC "
    Set oRs = GetRecordset(sSql, "V_CODE_Database")
    If oRs Is Nothing Then
'        LogMessage strProcName, "WARNING", "Either no attached documents or there was a problem with the query / connection!", Me.ID
        GoTo Block_Exit
    End If

    While Not oRs.EOF
        Set oAttachedFile = New clsConceptDoc
'   Debug.Assert oRs("RowID").Value <> 7073

        If oAttachedFile.LoadFromId(oRs("RowId")) = False Then
            LogMessage strProcName, "WARNING", "Problem loading an attached file!", "RowID: " & CStr(oRs("RowId").Value)
        Else
                '' Also add it to our ALL collection - though that's redundant. Just started that way before I decided to break them out..
            ccolAttachedDocs.Add oAttachedFile
        End If

        oRs.MoveNext
    Wend

    LoadAttachedDocs = True

Block_Exit:
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    LoadAttachedDocs = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function LoadTaggedClaims() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oClaim As clsEracClaim


    strProcName = ClassName & ".LoadTaggedClaims"

    sSql = "SELECT TOP 20 h.CnlyClaimNum, ICN, MedicalRecordNum " & _
            " FROM CMS_AUDITORS_CLAIMS.dbo.AuditClm_Hdr H INNER JOIN ( " & _
                " SELECT cnlyclaimnum, Adj_ConceptID as 'ConceptID' FROM CMS_AUDITORS_CODE.dbo.v_CONCEPT_ValidationSummary " & _
            " ) as AA ON AA.CnlyClaimNum = h.cnlyClaimNum " & _
            " WHERE AA.ConceptID = '" & Me.ID & "'"

    Set oRs = GetRecordset(sSql, "AuditClm_Hdr")
    If oRs Is Nothing Then
'        LogMessage strProcName, "WARNING", "Either no tagged claims or there was a problem with the query / connection!", Me.ID
        GoTo Block_Exit
    End If


    While Not oRs.EOF
        Set oClaim = New clsEracClaim
        If oClaim.LoadFromId(CStr("" & oRs("CnlyClaimNum").Value)) = False Then
            ' KD COMEBACK deal with this
        End If
            ' just stuff that in our collection
        ccolTaggedClaims.Add oClaim

        oRs.MoveNext
    Wend

    LoadTaggedClaims = True

Block_Exit:
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    LoadTaggedClaims = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Private Function GetReqRule() As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As ADODB.Recordset
'Dim sSql As String
'Dim sException As String
'
'    strProcName = ClassName & ".GetReqRule"
'
'    '' KD COMEBACK: When we have some Data type stuff in there then this needs to be dealt with! (or is this going to work?)
''    sSql = "SELECT ConceptReqId FROM CnlyConceptRequirements WHERE ReviewTypeId = " & _
''            Me.EracReviewTypeId & " AND ISNULL(DataTypeCode,'') = '" & Me.CnlyDataTypeCode & "' "
'
'    sSql = "SELECT ConceptReqId, DataTypeCode FROM CnlyConceptRequirements WHERE ReviewTypeId = " & _
'            Me.EracReviewTypeId & " AND ( ISNULL(DataTypeCode,'') = '" & Me.CnlyDataTypeCode & "' OR ISNULL(DataTypeCode,'') = '' ) "
'
'            '' KD COMEBACK: Note, if we get 2 then we need to use the one that has the DataTypeCode that isn't ''
'
'    sException = " AND ( ISNULL(ExceptionLOB,'') = '" & Me.GetField("LOB") & "' OR ISNULL(ExceptionAuditor,'') = '" & _
'            Me.GetField("Auditor") & "') "
'
'        '' First see if we have one with an exception
'    Set oRs = GetRecordset(sSql & sException, "CnlyConceptRequirements")
'    If Not oRs Is Nothing Then
'        If oRs.RecordCount > 1 Then
'            Do While Not oRs.EOF
'                If Nz(oRs("DataTypeCode"), "") = Me.CnlyDataTypeCode Then
'                    ' this is our record..
'                    Exit Do
'                End If
'                oRs.MoveNext
'            Loop
'        End If
'
'        Set coReqRule = New clsEracRequirementRule
'        GetReqRule = coReqRule.LoadFromID(CInt(oRs("ConceptReqId").Value))
'        GoTo Block_Exit
'    End If
'
'        '' Without the exception
'    Set oRs = GetRecordset(sSql, "CnlyConceptRequirements")
'    If oRs Is Nothing Then GoTo Block_Exit
'
'    If Not oRs.EOF Then
'        If oRs.RecordCount > 1 Then
'            Do While Not oRs.EOF
'                If Nz(oRs("DataTypeCode"), "") = Me.CnlyDataTypeCode Then
'                    ' this is our record..
'                    Exit Do
'                End If
'                oRs.MoveNext
'            Loop
'        End If
'
'        Set coReqRule = New clsEracRequirementRule
'        GetReqRule = coReqRule.LoadFromID(CInt(oRs("ConceptReqId").Value))
'        ''  GetReqRule = coReqRule.LoadFromConceptID(Me.ID)
'
'    End If
'
'
'Block_Exit:
'    Set oRs = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GetReqRule = False
'    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
'End Function




'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
'''' This function performs (or orchastrates) all validation  on a particular concept
'''' to see if it's ready to submit to CMS
''''
'Public Function GetConceptHeaderDetails(sConceptId As String, Optional ByRef sReviewTypeCode As String = "", _
'    Optional ByRef sDataTypeCode As String = "", Optional ByRef sConceptOwner As String = "", _
'    Optional ByRef sClientIssueNum As String = "", Optional iCmsReviewTypeId As Integer = 0) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim oRs As ADODB.Recordset
'
'    strProcName = ClassName & ".GetConceptReviewType"
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
'        .SQLTextType = sqltext
'        .SQLstring = "SELECT CH.Auditor ConceptOwner, CH.ClientIssueNum, CH.ReviewType, CH.DataType FROM " & _
'            " CMS_AUDITORS_CLAIMS.dbo.Concept_Hdr CH WHERE Ch.ConceptID = '" & sConceptId & "'"
'    End With
'
'    Set oRs = oAdo.ExecuteRS
'
'        '' Did we get anything?
'    If oRs Is Nothing Then
'        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
'        GoTo Block_Exit
'    End If
'
'    If oRs.RecordCount < 1 Then
'        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
'        GoTo Block_Exit
'    End If
'
'    If Not oRs.EOF Then
'        sConceptOwner = Nz(oRs("ConceptOwner").Value, "")
'        sReviewTypeCode = Nz(oRs("ReviewType").Value, "")
'        sDataTypeCode = Nz(oRs("DataType").Value, "")
'        sClientIssueNum = Nz(oRs("ClientIssueNum").Value, "")
'            ' We don't expect more than 1, so no need to movenext
'    End If
'
'    iCmsReviewTypeId = TranslateCnlyReviewTypeToCMS(sReviewTypeCode)
'
'    GetConceptHeaderDetails = True
'
'Block_Exit:
'    Set oRs = Nothing
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    GetConceptHeaderDetails = False
'    GoTo Block_Exit
'End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub GetPayerNamesDict()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetPayerNamesDict"

    Set oRs = GetRecordset("SELECT PayerNameId, PayerName FROM XREF_PAYERNAMES")
    
    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount = 0 Then GoTo Block_Exit

    Set cdctPayerNamesById = New Scripting.Dictionary
    
    While Not oRs.EOF
        If cdctPayerNamesById.Exists(CStr(oRs("PayerNameID").Value)) Then
            Stop    ' shouldn't get here, they are unique for crying out loud!
        Else
            cdctPayerNamesById.Add CStr(oRs("PayerNameID").Value), CStr("" & oRs("PayerName").Value)
        End If
        oRs.MoveNext
    Wend


Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Sub




''##########################################################
''##########################################################
''##########################################################
'' Error handling
''##########################################################
''##########################################################
''##########################################################
Private Sub FireError(oErr As ErrObject, sErrSourceProcName As String, Optional sAdditionalDetails As String)

    cbErrorOccurred = True

    ReportError oErr, sErrSourceProcName, , sAdditionalDetails

    If sAdditionalDetails <> "" Then sAdditionalDetails = vbCrLf & sAdditionalDetails

    RaiseEvent ConceptPayerError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub

'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################
'
'       Class Init / Term
'
'########################################################################################################
'########################################################################################################
'########################################################################################################
'########################################################################################################


Private Sub Class_Initialize()

    Set coSourceTable = New clsTable
    coSourceTable.IdFieldName = csIDFIELDNAME
    coSourceTable.TableName = csTableName

    Set ccolAttachedDocs = New Collection

    Set ccolTaggedClaims = New Collection
    Set coValidateRpt = New clsEracValidationRpt

    Call GetPayerNamesDict

    
    cblnIsInitialized = False

End Sub


Private Sub Class_Terminate()
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing

    Set ccolAttachedDocs = Nothing

    Set ccolTaggedClaims = Nothing
    Set coValidateRpt = Nothing
    Set cdctPayerNamesById = Nothing


    cblnIsInitialized = False
End Sub