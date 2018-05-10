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
'''  Represents a CMS Concept, basically a "hook" into the
'''     _CLAIMS.dbo.CONCEPT_Hdr table
'''  With validation and various other methods..
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 06/07/2013 - KD : added ConceptCatId property
'''  - 10/16/2012 - Added some validation around payers effective date and end dates..
'''  - 09/21/2012 - Added concept qa mode and concept status
'''  - 08/29/2012 - fixed SubmitTrackedDate (made it payer specific like I should have originally done)
'''     Also fixed other submit date logic to look in the correct place, leaving CONCEPT_Hdr (and payer dtl)
'''     values JUST for the NIRF
'''  - 08/20/2012 - Added SubmitTrackedDate and changed things to only load sub
'''     objects when needed
'''  - 06/20/2012 - Added Payer details and all kinds of additional things to support
'''  - 06/14/2012 - Fixed DateSubmitted
'''  - 05/07/2012 - Added PassedValidation
'''  - 04/25/2012 - more changes concerning validation and submission
'''  - 03/14/2012 - Created class
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

Public Event ConceptError(ErrMsg As String, ErrNum As Long, ErrSource As String)


Private cbErrorOccurred As Boolean

Private Const cstr_CONCEPT_ROOT_FOLDER As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\ConceptID\"
Private Const cstr_CONCEPT_WORK_FOLDER As String = "\\ccaintranet.com\dfs-cms-fld\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\"

Private Const csIDFIELDNAME As String = "ConceptId"
Private Const csTableName As String = "CONCEPT_Hdr"
Private coSourceTable As clsTable

    '' The table to use for the connection string to the _ERAC database
Private Const csSP_TABLENAME As String = "ConceptDocTypes"

Private Const csREQUIRED_CONCEPT_HDR_FIELDS As String = "ConceptDesc,ConceptLogic,OpportunityType,ReviewType,ConceptIndicator,ErrorCode,ProviderTypeId,ConceptPriority,Auditor,ConceptLevel,DataType,ConceptSource,LOB,ConceptRationale,ConceptStatus"

Private coReqRule As clsEracRequirementRule
Private ccolAttachedDocs As Collection
Private ccolHdrAttached As Collection
Private ccolDtlAttached As Collection

Private ccolTaggedClaims As Collection
Private ccolPayerDetails As Collection

Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private cdctAggregateFieldNames As Scripting.Dictionary

Private csConceptId As String
Private csCnlyReviewTypeCode As String
Private csCnlyDataTypeCode As String
Private ciEracReviewTypeId As Integer

Private cdctInitObjs As Scripting.Dictionary

Private coValidateRpt As clsEracValidationRpt


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ConceptID() As String
    ConceptID = csConceptId
End Property
Public Property Let ConceptID(sConceptId As String)
    csConceptId = sConceptId
End Property
        '' Just an alias for ease of use!
    Public Property Get Id() As String
        Id = ConceptID
    End Property
    Public Property Let Id(sNewId As String)
        ConceptID = sNewId
    End Property


Public Property Get ConceptCatId() As Long
    ConceptCatId = CLng("0" & GetTableValue("ConceptCatId"))
End Property
Public Property Let ConceptCatId(lConceptCatId As Long)
    SetTableValue "ConceptCatId", CStr(lConceptCatId)
End Property




Public Property Get CnlyReviewTypeCode() As String
    CnlyReviewTypeCode = CStr("" & GetTableValue("ReviewType"))
End Property
'Public Property Let CnlyReviewTypeCode(sCnlyReviewTypeCode As String)
'    SetTableValue "ReviewType", sCnlyReviewTypeCode
'End Property



Public Property Get EracReviewTypeId() As Integer
    EracReviewTypeId = EracReviewTypeFromCnlyCode(Me.CnlyReviewTypeCode)
End Property
'Public Property Let EracReviewTypeId(iEracReviewTypeId As Integer)
'    ciEracReviewTypeId = iEracReviewTypeId
'End Property


Public Property Get CnlyDataTypeCode() As String
    CnlyDataTypeCode = GetTableValue("DataType")
End Property
'Public Property Let CnlyDataTypeCode(sCnlyDataTypeCode As String)
'    SetTableValue "DataType", sCnlyDataTypeCode
'End Property


Public Property Get ConceptStatusNum() As String
Dim sRet As String

    sRet = Me.GetField("ConceptStatus")
    
    If IsNumeric(sRet) = True Then
        ConceptStatusNum = sRet
    End If
End Property





Public Property Get SubmitTrackedDate(lPayerNameId As Long) As Date
Dim sDate As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
'Stop

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_DATA_DATABASE")
        .SQLTextType = sqltext
        
        .sqlString = "SELECT MAX(SubmitClickedDt) as SubmitDt FROM CONCEPT_SubmitTracking WHERE ConceptID = '" & Me.ConceptID & "' "
        If lPayerNameId <> 1000 Then
            .sqlString = .sqlString & " AND PayerNameId = " & CStr(lPayerNameId)
        End If

        Set oRs = .ExecuteRS
        If .GotData = False Then
            SubmitTrackedDate = CDate("1/1/1900")
        Else
            sDate = Nz(oRs("SubmitDt").Value, "1/1/1900")
            If IsDate(sDate) = True Then
                SubmitTrackedDate = CDate(sDate)
            Else
                SubmitTrackedDate = CDate("1/1/1900")
            End If
        End If
    End With

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


Public Property Get IsPayerLevelConcept() As Boolean
    If ccolPayerDetails.Count > 0 Then IsPayerLevelConcept = True
End Property


Public Property Get RequirementRuleObj() As clsEracRequirementRule
    Call LoadSubObjects("GETREQRULE")
    If coReqRule Is Nothing Then
        
        RaiseEvent ConceptError("No requirement rule for this concept!", 12321, ClassName)
        Exit Property
    End If
    Set RequirementRuleObj = coReqRule
End Property



Public Property Get AttachedDocuments() As Collection
    Call LoadSubObjects("LOADATTACHEDDOCS")
    Set AttachedDocuments = ccolAttachedDocs
End Property

Public Property Get TaggedClaims() As Collection
    Call LoadSubObjects("LOADTAGGEDCLAIMS")
    Set TaggedClaims = ccolTaggedClaims
End Property



Public Property Get ConceptPayers() As Collection
    Call LoadSubObjects("LOADPAYERDETAILS")
    Set ConceptPayers = ccolPayerDetails
End Property



Public Property Get DesiredAdjOutcome() As String
    DesiredAdjOutcome = GetTableValue("DesiredAdjOutcome")
End Property
Public Property Let DesiredAdjOutcome(sDesiredAdjOutcome As String)
    SetTableValue "DesiredAdjOutcome", sDesiredAdjOutcome
End Property


Public Property Get ConceptPayerIDString() As String
Dim sRet As String
Dim oPayer As clsConceptPayerDtl

    For Each oPayer In ConceptPayers
        sRet = sRet & CStr(oPayer.PayerNameId) & ","
    Next
    If Right(sRet, 1) = "," Then sRet = left(sRet, Len(sRet) - 1)
    
    ConceptPayerIDString = sRet
End Property


Public Property Get ConceptFolder() As String
    ConceptFolder = cstr_CONCEPT_ROOT_FOLDER & Me.ConceptID & "\"
End Property



Public Property Get ConceptWorkFolder() As String
    ConceptWorkFolder = cstr_CONCEPT_WORK_FOLDER & Me.ConceptID & "\"
End Property

Public Property Get ClientIssueId(lPayerNameId As Long) As String
Dim oPayer As clsConceptPayerDtl

    If lPayerNameId = 0 Then
        ClientIssueId = CStr("" & GetTableValue("ClientIssueNum"))
    ElseIf lPayerNameId <> 1000 Then
        For Each oPayer In Me.ConceptPayers
            If oPayer.PayerNameId = lPayerNameId Then
                ClientIssueId = oPayer.ClientIssueId
                GoTo Block_Exit
            End If
        Next
    Else
        ClientIssueId = CStr("" & GetTableValue("ClientIssueNum"))
    End If
Block_Exit:

End Property
Public Property Let ClientIssueId(lPayerNameId As Long, sClientIssueId As String)
    Stop
    SetTableValue "ClientIssueNum", sClientIssueId
End Property

Public Function SetClientIssueNum(iPayerNameId, sClientIssueId) As Boolean
Dim oPayer As clsConceptPayerDtl

    If iPayerNameId = 0 Then

        SetTableValue "ClientIssueNum", sClientIssueId
    Else
        For Each oPayer In ccolPayerDetails
            If oPayer.PayerNameId = iPayerNameId Then
                oPayer.ClientIssueId = sClientIssueId
                oPayer.SaveNow
            End If
        Next
    End If
    
End Function

Public Property Get RiskFactor() As Double
    If IsPayerLevelConcept = True Then
        RiskFactor = AggregatePayerData("RiskFactor")
    Else
        RiskFactor = CDbl(GetTableValue("RiskFactor"))
    End If
End Property


Public Property Get CostFactor() As Double
    If IsPayerLevelConcept = True Then
        CostFactor = AggregatePayerData("CostFactor")
    Else
        CostFactor = CDbl(GetTableValue("CostFactor"))
    End If
End Property


Public Property Get MaxChartRequest() As Double
    If IsPayerLevelConcept = True Then
        MaxChartRequest = AggregatePayerData("MaxChartRequest")
    Else
        MaxChartRequest = CDbl(GetTableValue("MaxChartRequest"))
    End If
End Property


Public Property Get ConceptClaimCount() As Double
    If IsPayerLevelConcept = True Then
        ConceptClaimCount = AggregatePayerData("ConceptClaimCount")
    Else
        ConceptClaimCount = CDbl(GetTableValue("ConceptClaimCount"))
    End If
End Property


Public Property Get ConceptPotentialValue() As Double
    If IsPayerLevelConcept = True Then
        ConceptPotentialValue = AggregatePayerData("ConceptPotentialValue")
    Else
        ConceptPotentialValue = CDbl(GetTableValue("ConceptPotentialValue"))
    End If
End Property



Public Property Get SampleClaimCount() As Double
    If IsPayerLevelConcept = True Then
        SampleClaimCount = AggregatePayerData("SampleClaimClount")  '' hehe, someone mis spelled it and I copied it!
    Else
        SampleClaimCount = CDbl(GetTableValue("SampleClaimClount"))
    End If
End Property



Public Property Get SamplePotentialValue() As Double
    If IsPayerLevelConcept = True Then
        SamplePotentialValue = AggregatePayerData("SamplePotentialValue")
    Else
        SamplePotentialValue = CDbl(GetTableValue("SamplePotentialValue"))
    End If
End Property



Public Property Get CreateDate() As Date
Dim sRet As String
    sRet = GetTableValue("SamplePotentialValue")
    If IsDate(sRet) Then
        CreateDate = CDate(sRet)
    Else
        CreateDate = CDate("1/1/1900")
    End If

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
'''
''' sets the date and status to Concept Submitted in the EracConceptStatusLog
'''
Public Function MarkAsSubmitted(Optional ByRef sErrMessage As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".MarkAsSubmitted"
    
    ' bottom line, we need to set the date and status in the EracConceptStatusLog
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        .sqlString = "usp_EracSetConceptAsSubmitted"
        .Parameters("@pConceptId") = Me.ConceptID
        .Parameters("@pSubmitUser") = Identity.UserName
        
        If .Execute() = 0 Then
            sErrMessage = "Update failed for unknown reason with : " & CStr(.CurrentConnection.Errors.Count) & " Ado errors"
            MarkAsSubmitted = False
            GoTo Block_Exit
        End If
        
    End With
    

    MarkAsSubmitted = True

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    MarkAsSubmitted = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function RefreshPayerCollection() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".RefreshPayerCollection"
    
    RefreshPayerCollection = LoadPayerDetails()
    
Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    RefreshPayerCollection = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function SubmitDocPaths(Optional ByRef sOutMessage As String, Optional sResubmitPath As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oReqDoc As clsConceptReqDocType
Dim oAtchDoc As clsConceptDoc
Dim sOneDoc As String
Dim sRet As String
Dim sPackageFldr As String

    strProcName = ClassName & ".SubmitDocPaths"

    If sResubmitPath <> "" Then
        sPackageFldr = QualifyFldrPath(sResubmitPath)
    Else
        sPackageFldr = ConceptWorkFolder
    End If

    For Each oAtchDoc In Me.AttachedDocuments
                '' For now, we are only sending the package (hdr) level documents
                '' the claim level (dtl lvl) docs will be burned to CD and sent separately
        If oAtchDoc.GetEracReqDocType.IsPayerDoc = True Then
            sOneDoc = oAtchDoc.GetEracReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId(0), oAtchDoc.Icn, _
                    oAtchDoc.FileName & oAtchDoc.FileName)
            sOneDoc = sPackageFldr & sOneDoc & "." & LCase(oAtchDoc.GetEracReqDocType.SendAsFileType)
            sRet = sRet & Replace(sOneDoc, ConceptFolder, sPackageFldr) & ","
        Else
'            sRet = sRet & oAtchDoc.ConvertedFilePath
            Select Case oAtchDoc.CnlyAttachType
            Case "ERAC_ScreenShot"
                ' skip these.. can't go in email
            Case Else
                Stop
                sOneDoc = oAtchDoc.GetEracReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId(0), oAtchDoc.Icn, _
                        oAtchDoc.FileName)
                sOneDoc = sPackageFldr & sOneDoc & "." & LCase(oAtchDoc.GetEracReqDocType.SendAsFileType)
                sRet = sRet & Replace(sOneDoc, ConceptFolder, sPackageFldr) & ","
            End Select

        End If

    Next
        ' remove final comma
    If Len(sRet) > 2 Then sRet = left(sRet, Len(sRet) - 1)
    SubmitDocPaths = sRet
    
Block_Exit:
    Set oReqDoc = Nothing
    Set oAtchDoc = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Has the concept already been submitted?
'''
Public Function AlreadySubmitted(lPayerNameId As Long, Optional ByRef sOutMessage As String) As Date
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSubmitUser As String

    strProcName = ClassName & ".AlreadySubmitted"
    AlreadySubmitted = CDate("1/1/1900")
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = StoredProc
        If lPayerNameId <> 0 Then
            .sqlString = "usp_EracWasConceptSubmitted_Payer"
            .Parameters.Refresh
            .Parameters("@pConceptId") = Me.ConceptID
            .Parameters("@pPayerNameID") = lPayerNameId
        Else
            .sqlString = "usp_EracWasConceptSubmitted"
            .Parameters.Refresh
            .Parameters("@pConceptId") = Me.ConceptID
'            .Parameters("@pPayerNameID") = iPayerNameID
        End If
        .Execute
        If .Parameters("@pErrMsg").Value <> "" Then
            If Me.DateSubmitted > CDate("1/1/1900") And Me.DateSubmitted < CDate("5/1/2012") Then
                AlreadySubmitted = Me.DateSubmitted
            End If
            GoTo Block_Exit
        End If
        AlreadySubmitted = .Parameters("@pSubmitDate").Value
        sSubmitUser = .Parameters("@pSubmitUser").Value
        If AlreadySubmitted > CDate("1/1/1900") Then
            sOutMessage = "Concept was already submitted on " & CStr(AlreadySubmitted) & " by " & sSubmitUser
        End If
    End With

Block_Exit:
    
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''  This will be for validating the entire concept (all payers)
'''
Public Function ValidatePayerForSubmission(oPayerDtl As clsConceptPayerDtl, Optional ByRef oRs As ADODB.RecordSet, Optional bValidateConceptHdr As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oReqRule As clsEracRequirementRule
Dim dSubmitDate As Date
Dim sOutMessage As String
Dim sReport As String

    strProcName = ClassName & ".ValidatePayerForSubmission"
    ValidatePayerForSubmission = True
    Set coValidateRpt = New clsEracValidationRpt

    If WasInitialized = False Then
        sReport = "Concept ID not set"
        ValidatePayerForSubmission = False
        GoTo Block_Exit
    End If

        '' eRAC is going away before I deploy this whole thing!
        '' KD COMEBACK: THE NEW STORY IS:
        ''  we are going to create it ourselves using the sproc I wrote: usp_CMS_Get_New_ClientIssueNum (_ERAC)
    If oPayerDtl.ClientIssueId = "" Then
        If IssueClientIssueNum(Me, oPayerDtl.PayerNameId, sReport) = "" Then
            LogMessage strProcName, "ERROR", "There was a problem generating the Client Issue ID for payer: " & oPayerDtl.PayerName
            ValidatePayerForSubmission = False
            GoTo Block_Exit
        End If
                
    End If
    
    
    If bValidateConceptHdr = True Then
            '' Does it have all of the required fields in CONCEPT_hdr?
        If Me.HasRequiredFields(sOutMessage) = False Then
            sReport = sReport & "Missing Concept level required fields: " & vbCrLf & sOutMessage & vbCrLf
            ValidatePayerForSubmission = False
            coValidateRpt.AddNote False, "Required Concept fields", oPayerDtl.PayerName, sOutMessage
            sOutMessage = ""
        Else
            coValidateRpt.AddNote True, "Required Concept fields", oPayerDtl.PayerName, "All required fields at the concept level have values"
        End If
    End If
    

    If oPayerDtl.HasRequiredFields(sOutMessage) = False Then
        sReport = sReport & "Missing Concept level required fields: " & vbCrLf & sOutMessage & vbCrLf
        ValidatePayerForSubmission = False
        coValidateRpt.AddNote False, "Required payer fields", oPayerDtl.PayerName, sOutMessage
        sOutMessage = ""
    Else
        coValidateRpt.AddNote True, "Required payer fields", oPayerDtl.PayerName, "All required fields at the payer level have values"
    End If


        '' Ok, prompt the user if needed
     If CheckTaggedClaimsAndPromptUserIfNeeded(oPayerDtl, sOutMessage, coValidateRpt) = False Then
         LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims", oPayerDtl.PayerName & " for concept: " & ConceptID
         ValidatePayerForSubmission = False
         If Not coValidateRpt Is Nothing Then
             coValidateRpt.AddNote False, "Number of expected tagged claims", oPayerDtl.PayerName, sOutMessage
         End If
     Else
         coValidateRpt.AddNote True, "Number of expected tagged claims", oPayerDtl.PayerName, "Ok: all of the expected claims were found for " & oPayerDtl.PayerName
     End If
     sOutMessage = ""
       
        '' Assuming we got here, we have the correct amount of tagged claims
        '' we need to make sure that no tagged claim has already been submitted to another concept
    If TaggedClaimsAlreadySubmittedToAnotherConcept(oPayerDtl, sOutMessage) = True Then
        LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
        ValidatePayerForSubmission = False
        If Not coValidateRpt Is Nothing Then
            coValidateRpt.AddNote False, "Tagged claims already submitted", oPayerDtl.PayerName, sOutMessage
        End If
    Else
        coValidateRpt.AddNote True, "Tagged claims already submitted", oPayerDtl.PayerName, "Ok: none of the tagged claims have been submitted to another concept"
    End If
    sOutMessage = ""
    
    
        '' Does it have all of the required documents attached?
    sOutMessage = Me.GetMissingRequiredDocsMessage(oPayerDtl)
    If sOutMessage <> "" Then
        sReport = sReport & "Missing required documents: " & sOutMessage

        ValidatePayerForSubmission = False
            '        coValidateRpt.AddNote False, "Required documents", "Missing: " & sOutMessage
    Else
        coValidateRpt.AddNote True, "Required documents", oPayerDtl.PayerName, "All required documents present"
    End If
    sOutMessage = ""
        
        '' Has it been submitted yet?
    dSubmitDate = Me.AlreadySubmitted(oPayerDtl.PayerNameId, sOutMessage)
    If dSubmitDate <> CDate("1/1/1900") Then
        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
        ValidatePayerForSubmission = False
        coValidateRpt.AddNote False, "Concept already submitted", oPayerDtl.PayerName, "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
    Else
        coValidateRpt.AddNote True, "Concept already submitted", oPayerDtl.PayerName, "OK: Not submitted yet"
    End If
    sOutMessage = ""
    

    
Block_Exit:
    Set oRs = coValidateRpt.GetRecordset
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    ValidatePayerForSubmission = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''  This will be for validating the entire concept (all payers)
'''
Public Function ValidateForSubmission(Optional ByRef oRs As ADODB.RecordSet, Optional lngSpecificPayerNameId As Long = 0) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oReqRule As clsEracRequirementRule
Dim dSubmitDate As Date
Dim sOutMessage As String
Dim sReport As String
Dim oPayerDtl As clsConceptPayerDtl
Dim colPayers As Collection
Dim colInvalidPayers As Collection

    strProcName = ClassName & ".ValidateForSubmission"
    Call LoadSubObjects("")
'    Call LoadSubObjects
'    Call LoadSubObjects
    
    ValidateForSubmission = True
    Set coValidateRpt = New clsEracValidationRpt

    If WasInitialized = False Then
        sReport = "Concept ID not set"
        ValidateForSubmission = False
        GoTo Block_Exit
    End If

        '' First, check that each payer has several things (unless we are only doing 1 payer, in that case
        '' only check the 1 payer
        '' To Simplify this, I'll use a copy of the collection with ONLY our single payer object
    Set colPayers = New Collection
    Set colInvalidPayers = New Collection
    
    For Each oPayerDtl In ccolPayerDetails
        If lngSpecificPayerNameId = 1000 Or oPayerDtl.PayerNameId = lngSpecificPayerNameId Then
            ' Only add the payer if it's enabled and ready to accept submissions
'            If oPayerDtl.EffectiveDate > Now() Or oPayerDtl.EndDate < Now() Then
'                If oPayerDtl.PayerStatusNum <> "990" Then
'                    colInvalidPayers.Add oPayerDtl
'                    coValidateRpt.AddNote False, "Payer not valid!", Nz(oPayerDtl.PayerName, ""), "Check the effective and end dates for this payer!"
'                End If
'            Else
'                colPayers.Add oPayerDtl
'            End If
            If oPayerDtl.EffectiveDate > Now() Or oPayerDtl.EndDate < Now() Or oPayerDtl.PayerStatusNum = "990" Then
                colInvalidPayers.Add oPayerDtl
                coValidateRpt.AddNote False, "Payer not valid!", Nz(oPayerDtl.PayerName, ""), "Check the effective and end dates for this payer!"
            Else
                colPayers.Add oPayerDtl
            End If

        End If
    Next

    If colPayers.Count < 1 Then
        ValidateForSubmission = False

        GoTo Block_Exit
    End If

            '' Start looping over the payers:
    For Each oPayerDtl In colPayers
            ' make sure it has a client issue id
        If oPayerDtl.ClientIssueId = "" Then
            If IssueClientIssueNum(Me, oPayerDtl.PayerNameId, sReport) = "" Then
                LogMessage strProcName, "ERROR", "There was a problem generating the Client Issue ID for payer: " & oPayerDtl.PayerName
                ValidateForSubmission = False
                GoTo Block_Exit
            End If
                    
        End If

        
            '' Does it have all of the required fields in CONCEPT_hdr?
            ' this is complicated now because some fields can be CONCEPT_Hdr specific, but the same
            ' field may be CONCEPT_Payer_Dtl specific.. We have to look at both
        If Me.HasRequiredFields(sOutMessage, oPayerDtl) = False Then
            sReport = sReport & "Missing Concept level required fields: " & vbCrLf & sOutMessage & vbCrLf
            ValidateForSubmission = False
            coValidateRpt.AddNote False, "Required Concept fields", oPayerDtl.PayerName, sOutMessage
            sOutMessage = ""
        Else
            coValidateRpt.AddNote True, "Required fields", oPayerDtl.PayerName, "All required fields at the concept level have values"
        End If
        
            
        ' Now, do we have the required amount of tagged claims for this payer?

        '' Ok, prompt the user if needed
         If CheckTaggedClaimsAndPromptUserIfNeeded(oPayerDtl, sOutMessage, coValidateRpt) = False Then
             LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
             ValidateForSubmission = False
             If Not coValidateRpt Is Nothing Then
                 coValidateRpt.AddNote False, "Number of expected tagged claims", oPayerDtl.PayerName, sOutMessage
             End If
         Else
             coValidateRpt.AddNote True, "Number of expected tagged claims", oPayerDtl.PayerName, "Ok: " & oPayerDtl.PayerName & " has of the expected claims were found"
         End If
         sOutMessage = ""
        
        
            '' Assuming we got here, we have the correct amount of tagged claims
            '' we need to make sure that no tagged claim has already been submitted to another concept
    
        If TaggedClaimsAlreadySubmittedToAnotherConcept(oPayerDtl, sOutMessage) = True Then
            LogMessage strProcName, "ERROR", "Some of the tagged claims have been submitted with another concept", "Payer: " & CStr(oPayerDtl.PayerNameId) & " concept: " & ConceptID
            ValidateForSubmission = False
            If Not coValidateRpt Is Nothing Then
                coValidateRpt.AddNote False, "Tagged claims already submitted", oPayerDtl.PayerName, sOutMessage
            End If
        Else
            coValidateRpt.AddNote True, "Tagged claims already submitted", oPayerDtl.PayerName, "Ok: none of the tagged claims have been submitted to another concept"
        End If
        sOutMessage = ""
        

        
            '' Does it have all of the required documents attached?
        sOutMessage = Me.GetMissingRequiredDocsMessage(oPayerDtl)
        If sOutMessage <> "" Then
            sReport = sReport & "Missing required documents: " & sOutMessage
    
            ValidateForSubmission = False
                '        coValidateRpt.AddNote False, "Required documents", "Missing: " & sOutMessage
        Else
            coValidateRpt.AddNote True, "Required documents", oPayerDtl.PayerName, "All required documents present"
        End If
        sOutMessage = ""
            
            '' Has it been submitted yet?
        dSubmitDate = Me.AlreadySubmitted(oPayerDtl.PayerNameId, sOutMessage)
        If dSubmitDate <> CDate("1/1/1900") Then
            sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
            ValidateForSubmission = False
            coValidateRpt.AddNote False, "Concept already submitted", oPayerDtl.PayerName, "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
        Else
            coValidateRpt.AddNote True, "Concept already submitted", oPayerDtl.PayerName, "OK: Not submitted yet"
        End If
        sOutMessage = ""
        
        
    Next

    Dim sMsg As String
    
    For Each oPayerDtl In colInvalidPayers
        If oPayerDtl.PayerStatusNum <> "990" Then
            sMsg = sMsg & oPayerDtl.PayerName & " is not a valid payer! Effective Date: " & Format(oPayerDtl.EffectiveDate, "mm/dd/yyyy") & " to " & Format(oPayerDtl.EndDate, "mm/dd/yyyy") & vbCrLf
        End If
    Next
    If sMsg <> "" Then
        LogMessage strProcName, "USER NOTE", "Please note that at least 1 payer is invalid! " & vbCrLf & sMsg & vbCrLf & "No validation details will be displayed as these payers should have their status set to void", , True
    End If

    
Block_Exit:
    Set oRs = coValidateRpt.GetRecordset
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    ValidateForSubmission = False
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Moving this to the Concept Object
'''
Public Function CheckTaggedClaimsAndPromptUserIfNeeded(Optional oPayer As clsConceptPayerDtl, Optional sReport As String, Optional coValidateRpt As clsEracValidationRpt) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iTaggedClaims As Integer
Dim iExpectedClaims As Integer
Static dctPayerClaims As Scripting.Dictionary
Dim oPayerItm As clsConceptPayerDtl
Dim oClaim As clsEracClaim
Dim sThisPayerNameID As String
Dim sPayerName As String
Dim oClaimCol As Collection
Static bPromptUser As Boolean
Dim bAtLeastOneFailed As Boolean

    strProcName = ClassName & ".CheckTaggedClaimsAndPromptUserIfNeeded"

    '' Let's do this:
    ' Set up a dictionary: Key = PayerNameID, Value = Collection of Tagged Claim objects
    '   (do this once - hence the dictionary being static)
    CheckTaggedClaimsAndPromptUserIfNeeded = True
    
    If Not oPayer Is Nothing Then
        sPayerName = oPayer.PayerName
    End If

'        Stop
    Call LoadSubObjects("LOADTAGGEDCLAIMS")
        ' Set up our dictionary 1 x
    If dctPayerClaims Is Nothing Then
        Set dctPayerClaims = New Scripting.Dictionary

        
        For Each oClaim In Me.TaggedClaims
'Debug.Assert oClaim.PayerNameId <> 0


            sThisPayerNameID = UCase(oClaim.PayerNameId)
            If dctPayerClaims.Exists(sThisPayerNameID) = True Then
                Set oClaimCol = dctPayerClaims.Item(sThisPayerNameID)
                oClaimCol.Add oClaim
                Set dctPayerClaims.Item(sThisPayerNameID) = oClaimCol
            Else
                Set oClaimCol = New Collection
                oClaimCol.Add oClaim
                dctPayerClaims.Add sThisPayerNameID, oClaimCol
            End If
        Next
    End If
    
    
        '' Now loop over the payers and see if we have enough
    ' Need to start out as true for prompting or we never will..
    bPromptUser = True
    
    For Each oPayerItm In Me.ConceptPayers
        If Not oPayer Is Nothing Then
            If oPayer.PayerNameId <> oPayerItm.PayerNameId Then
                ' then we do not process this one..
                GoTo SkipIt
            End If
        End If
    
            ' How many claims are we expected to get for this payer
            '' the below function checks our Exception table which is populated by the PromptUserForTaggedClaimsException
            '' method call below...
        iExpectedClaims = oPayer.RequiredClaimsNum
    
        If dctPayerClaims.Exists(CStr("" & oPayer.PayerNameId)) = False Then
                '' No tagged claims for this payer
                ' if we prompted them before and they said, the number they entered should be applied for all then don't
                ' prompt again
            If bPromptUser = True Then
                If PromptUserForTaggedClaimsException(Me, oPayer.PayerNameId, iTaggedClaims, iExpectedClaims, bPromptUser) = False Then
                    CheckTaggedClaimsAndPromptUserIfNeeded = False
                        '' canceled or errored.. Do not proceed
                        '' ( details are already logged.. )
                    sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
                    If Not coValidateRpt Is Nothing Then
                        coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Either user canceled when prompted to change or there was a problem saving to DB"
                    End If
                End If
            Else
                CheckTaggedClaimsAndPromptUserIfNeeded = False
                    '' canceled or errored.. Do not proceed
                    '' ( details are already logged.. )
                sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
                
                If Not coValidateRpt Is Nothing Then
                    coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Either user canceled when prompted to change or there was a problem saving to DB"
                End If
            End If
        Else
            ' We have some, now check to see if we have more, or less than we need
            Set oClaimCol = dctPayerClaims.Item(UCase(oPayer.PayerNameId))
            'If oClaimCol.Count < oPayer.RequiredClaimsNum Then
            If oClaimCol.Count < iExpectedClaims Then
                If bPromptUser = True Then
                    If PromptUserForTaggedClaimsException(Me, oPayer.PayerNameId, oClaimCol.Count, iExpectedClaims, bPromptUser) = False Then
                        CheckTaggedClaimsAndPromptUserIfNeeded = False
                            '' canceled or errored.. Do not proceed
                            '' ( details are already logged.. )
                        sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
                        If Not coValidateRpt Is Nothing Then
                            coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Have: " & CStr(oClaimCol.Count) & " expected: " & CStr(iExpectedClaims) & ". Either user canceled when prompted to change or there was a problem saving to DB"
                        End If
                    End If
                Else
                    CheckTaggedClaimsAndPromptUserIfNeeded = False
                        '' canceled or errored.. Do not proceed
                        '' ( details are already logged.. )
                    sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
                    
                    If Not coValidateRpt Is Nothing Then
                        coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Have: " & CStr(oClaimCol.Count) & " expected: " & CStr(iExpectedClaims) & ". Either user canceled when prompted to change or there was a problem saving to DB"
                    End If

                End If
                    '' 20120719: KD: we no longer care if we've got more than we need..
'''            ElseIf oClaimCol.Count > oPayer.RequiredClaimsNum Then
'''                If bPromptUser = True Then
'''                    If PromptUserForTaggedClaimsException(Me, oPayer.PayerNameId, oClaimCol.Count, iExpectedClaims, bPromptUser) = False Then
'''                        CheckTaggedClaimsAndPromptUserIfNeeded = False
'''                            '' canceled or errored.. Do not proceed
'''                            '' ( details are already logged.. )
'''                        sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
'''                        If Not coValidateRpt Is Nothing Then
'''                            coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Either user canceled when prompted to change or there was a problem saving to DB"
'''                        End If
'''                    End If
'''                Else
'''                    CheckTaggedClaimsAndPromptUserIfNeeded = False
'''                        '' canceled or errored.. Do not proceed
'''                        '' ( details are already logged.. )
'''                    sReport = sReport & vbCrLf & "Either user canceled when prompted to change or there was a problem saving to DB"
'''                    If Not coValidateRpt Is Nothing Then
'''                        coValidateRpt.AddNote False, "Required Tagged Claims", sPayerName, "Either user canceled when prompted to change or there was a problem saving to DB"
'''                    End If
'''                End If
            Else    ' we are good:
                If Not coValidateRpt Is Nothing Then
                    coValidateRpt.AddNote True, "Required Tagged Claims", sPayerName, "Ok: We have " & CStr(iExpectedClaims) & " tagged claims! (Expected: " & CStr(iExpectedClaims) & ")"
                End If
                CheckTaggedClaimsAndPromptUserIfNeeded = True
            End If
        End If

SkipIt:
        If bAtLeastOneFailed = False Then
            bAtLeastOneFailed = Not CheckTaggedClaimsAndPromptUserIfNeeded
        End If
    Next
    
    CheckTaggedClaimsAndPromptUserIfNeeded = Not bAtLeastOneFailed
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    CheckTaggedClaimsAndPromptUserIfNeeded = False
    GoTo Block_Exit
End Function





'''
'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
''''''
''''''   (ORIGINAL
'''Public Function ValidateForSubmission(Optional ByRef oRs As ADODB.Recordset) As Boolean
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim sReturn As String
'''Dim oReqRule
'''Dim dSubmitDate As Date
'''Dim sOutMessage As String
'''Dim sReport As String
'''
'''    strProcName = ClassName & ".ValidateForSubmission"
'''    ValidateForSubmission = True
'''    Set coValidateRpt = New clsEracValidationRpt
'''
'''    If WasInitialized = False Then
'''        sReport = "Concept ID not set"
'''        ValidateForSubmission = False
'''        GoTo Block_Exit
'''    End If
'''
'''        '' eRAC is going away before I deploy this whole thing!
''''' KD COMEBACK: THE NEW STORY IS:
'''''  we are going to create it ourselves using the sproc I wrote: usp_CMS_Get_New_ClientIssueNum (_ERAC)
'''
'''    If Me.ClientIssueId = "" Then
'''Stop
'''        If IssueClientIssueNum(Me, 0, sReport) = "" Then
'''            LogMessage strProcName, "ERROR", "There was a problem generating the Client Issue ID"
'''            ValidateForSubmission = False
'''            GoTo Block_Exit
'''        End If
'''
'''    End If
'''
'''        '' Does it have all of the required fields in CONCEPT_hdr?
'''    If Me.HasRequiredFields(sOutMessage) = False Then
'''        sReport = sReport & "Missing required fields: " & vbCrLf & sOutMessage & vbCrLf
'''        ValidateForSubmission = False
'''        coValidateRpt.AddNote False, "Required fields", sOutMessage
'''        sOutMessage = ""
'''    Else
'''        coValidateRpt.AddNote True, "Required fields", "All required fields have values"
'''    End If
'''
'''
'''        '' Ok, prompt the user if needed
'''    If CheckTaggedClaimsAndPromptUserIfNeeded(ConceptID, sOutMessage) = False Then
'''        LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
'''        ValidateForSubmission = False
'''        If Not coValidateRpt Is Nothing Then
'''            coValidateRpt.AddNote False, "Number of expected tagged claims", sOutMessage
'''        End If
'''    Else
'''        coValidateRpt.AddNote True, "Number of expected tagged claims", "Ok: all of the expected claims were found"
'''    End If
'''    sOutMessage = ""
'''
'''
'''        '' Assuming we got here, we have the correct amount of tagged claims
'''        '' we need to make sure that no tagged claim has already been submitted to another concept
'''    If TaggedClaimsAlreadySubmittedToAnotherConcept(sOutMessage) = True Then
'''        LogMessage strProcName, "ERROR", "There was a problem prompting the user for tagged claims"
'''        ValidateForSubmission = False
'''        If Not coValidateRpt Is Nothing Then
'''            coValidateRpt.AddNote False, "Tagged claims already submitted", sOutMessage
'''        End If
'''    Else
'''        coValidateRpt.AddNote True, "Tagged claims already submitted", "Ok: none of the tagged claims have been submitted to another concept"
'''    End If
'''    sOutMessage = ""
'''
'''
'''        '' Does it have all of the required documents attached?
'''    sOutMessage = Me.GetMissingRequiredDocsMessage()
'''    If sOutMessage <> "" Then
'''        sReport = sReport & "Missing required documents: " & sOutMessage
'''
'''        ValidateForSubmission = False
'''            '        coValidateRpt.AddNote False, "Required documents", "Missing: " & sOutMessage
'''    Else
'''        coValidateRpt.AddNote True, "Required documents", "All required documents present"
'''    End If
'''    sOutMessage = ""
'''
'''        '' Has it been submitted yet?
'''    dSubmitDate = Me.AlreadySubmitted(sOutMessage)
'''    If dSubmitDate <> CDate("1/1/1900") Then
'''        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'''        ValidateForSubmission = False
'''        coValidateRpt.AddNote False, "Concept already submitted", "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
'''    Else
'''        coValidateRpt.AddNote True, "Concept already submitted", "OK: Not submitted yet"
'''    End If
'''    sOutMessage = ""
'''
'''
'''
'''Block_Exit:
'''    Set oRs = coValidateRpt.GetRecordset
'''    Exit Function
'''Block_Err:
'''    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'''    ValidateForSubmission = False
'''    GoTo Block_Exit
'''End Function
'''

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function PreviouslyPassedValidation(lPayerNameId As Long, Optional ByRef sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oTClaim As clsEracClaim
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".PreviouslyPassedValidation"
    
    If lPayerNameId = 0 Then lPayerNameId = 1000
    
    
    PreviouslyPassedValidation = mod_Concept_Specific.GetValidationHist(Me.ConceptID, lPayerNameId)
    
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = GetConnectString("ConceptDocTypes")
''        .SQLTextType = sqltext
''        .SQLstring = "SELECT * FROM v_EracConceptHistory WHERE ConceptID = '" & _
''                    Me.ConceptID & "' AND EracActionID = 18 AND ActionResult <> 'Failed' "
''        Set oRs = .ExecuteRS
'
'
'        If .GotData = True Then
'            PreviouslyPassedValidation = True
'        End If
'    End With
    
    
Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName()
    PreviouslyPassedValidation = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function TaggedClaimsAlreadySubmittedToAnotherConcept(oPayer As clsConceptPayerDtl, Optional ByRef sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oTClaim As clsEracClaim
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".TaggedClaimsAlreadySubmittedToAnotherConcept"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
    End With
    
    For Each oTClaim In TaggedClaims
        
        oAdo.sqlString = "SELECT * FROM CnlyTaggedClaimsByConcept WHERE CnlyClaimNum = '" & _
                    oTClaim.CnlyClaimNum & "' AND ConceptId <> '" & Me.ConceptID & _
                    "' AND ISNULL(PayerNameID,0) IN (0, " & CStr(oPayer.PayerNameId) & ")"
        Set oRs = oAdo.ExecuteRS
        If oAdo.GotData = True Then
                ' problem!!!
            sReport = sReport & "Claim " & oTClaim.CnlyClaimNum & " was submitted with a different concept: " & oRs("ConceptId") & " already" & vbCrLf
            If Not coValidateRpt Is Nothing Then
                coValidateRpt.AddNote False, "Tagged Claim already submitted to different concept!", oPayer.PayerName, "Connolly Claim Num: " & oTClaim.CnlyClaimNum & " was submitted to " & oRs("ConceptId") & " already"
            End If
        End If
    Next
    
    
Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName()
    TaggedClaimsAlreadySubmittedToAnotherConcept = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ValidateForClientIdRequest(lPayerNameId As Long, ByRef sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sOutMessage As String
Dim dSubmitDate As Date

    strProcName = ClassName & ".ValidateForClientIdRequest"
' is this ever getting called?
    
    ValidateForClientIdRequest = True
    Set coValidateRpt = New clsEracValidationRpt
    
    If WasInitialized = False Then
        sReport = "Concept ID not set"
        GoTo Block_Exit
    End If

        '' Does it have the ClientIssueID
    If Me.ClientIssueId(lPayerNameId) <> "" Then
        sReport = sReport & "Concept already has a Client Issue Id: " & Me.ClientIssueId(lPayerNameId) & vbCrLf
        ValidateForClientIdRequest = False
    End If
    
        '' Does it have all of the required fields in CONCEPT_hdr?
    If Me.HasRequiredFields(sOutMessage) = False Then
        sReport = sReport & "Missing required fields: " & vbCrLf & sOutMessage & vbCrLf
        ValidateForClientIdRequest = False
        sOutMessage = ""
    End If
    
        '' Has it been submitted yet?

    dSubmitDate = Me.AlreadySubmitted(lPayerNameId, sOutMessage)
    If dSubmitDate <> CDate("1/1/1900") Then
        sReport = sReport & "Concept was submitted on " & dSubmitDate & vbCrLf & sOutMessage & vbCrLf
        ValidateForClientIdRequest = False
        sOutMessage = ""
    End If
    
'        '' Is the notification status 1??
'    If Me.IsStatusOkForEracSubmission(sOutMessage) = False Then
'        sReport = sReport & sOutMessage & vbCrLf
'    End If
    
Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    ValidateForClientIdRequest = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetHdrAttachedDocType(sReqDocType As String) As clsConceptDoc
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttachedDoc As clsConceptDoc

    strProcName = ClassName & ".GetHdrAttachedDocType"
    
        ' These are only going to have 1 per concept.. (right??)
    For Each oAttachedDoc In ccolHdrAttached
        If oAttachedDoc.DocTypeName = sReqDocType Then
            Set GetHdrAttachedDocType = oAttachedDoc
            GoTo Block_Exit
        End If
    Next

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function CountReqDocsOfType(oRequiredDocType As clsConceptReqDocType, Optional intPayerNameId As Integer = 1000) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttachedDoc As clsConceptDoc
Dim iFoundCount As Integer

    strProcName = ClassName & ".CountHdrReqDocsOfType"
    
    ' Just in case, 0 should be the 'ALL'
    If intPayerNameId = 0 Then intPayerNameId = 1000
        
    
    For Each oAttachedDoc In ccolAttachedDocs
        If oAttachedDoc.GetEracReqDocType.CnlyAttachType = oRequiredDocType.CnlyAttachType _
        And (oAttachedDoc.PayerNameId = intPayerNameId Or intPayerNameId = 1000) Then
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




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
'''
Public Function HasRequiredDocsAttached(Optional ByRef sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oReqDocType As clsConceptReqDocType
Dim oAttachedDoc As clsConceptDoc
Dim bAtLeastOneNotFound As Boolean
Dim iFilesNeeded As Integer
Dim iFilesFound As Integer
Dim sMsg As String


    strProcName = ClassName & ".HasRequiredDocsAttached"
    ' kd - didn't do this yet..

    For Each oReqDocType In coReqRule.RequiredDocs
            ' If it's a header level doc (package level)
        If oReqDocType.IsPayerDoc = True Then
            iFilesNeeded = oReqDocType.NumPerConcept
        Else    ' must be a Detail level doc
            ' we are supposed to have oReqDocType.NumPerClaim of these..
            '' KD COMEBACK, need to change the below to NumPerClaim * Me.TaggedClaims.Count
            iFilesNeeded = oReqDocType.NumPerPayer * Me.TaggedClaims.Count
        End If
        
        If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
            ' we don't care if we have more - do we?
            ' also, if we are going to make it, then we don't care..
            If Nz(oReqDocType.CreateFunctionName, "") = "" Then
                bAtLeastOneNotFound = True
            End If
        End If
    Next
    '' KD COMEBACK: Put something in sReport!
    sReport = ""
    
    If bAtLeastOneNotFound = True Then
        ' Fire our error event
        Dim oErr As ErrObject
        Set oErr = New ErrObject
        oErr.Description = ""
        oErr.Number = 1234
        oErr.Source = strProcName
        
        FireError oErr, strProcName, sMsg
    End If
    
    HasRequiredDocsAttached = Not bAtLeastOneNotFound

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    HasRequiredDocsAttached = False
    GoTo Block_Exit
End Function

'' usp_SetTaggedClaimNumException

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Set the number of claims are expected for this concept
'''
Public Function SetRequiredClaimsNum(iNewAmount As Integer, ByRef sOutMsg As String, lngPayerNameId As Long, bApplyToSingle As Boolean) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim colPayers As Collection
Dim oPayer As clsConceptPayerDtl

    strProcName = ClassName & ".RequiredClaimsNum"
    
    Set colPayers = New Collection
    For Each oPayer In Me.ConceptPayers
        If bApplyToSingle = False Or oPayer.PayerNameId = lngPayerNameId Then
            colPayers.Add oPayer
        End If
    Next
    
    For Each oPayer In colPayers
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("ConceptDocTypes")
            .SQLTextType = StoredProc
            .sqlString = "usp_SetTaggedClaimNumException"
            .Parameters("@pConceptId") = Me.ConceptID
            .Parameters("@pPayerNameId") = oPayer.PayerNameId
            .Parameters("@pClaimsToBeSubmitted") = iNewAmount
            
            Call .Execute
            sOutMsg = CStr("" & .Parameters("@pErrMsg").Value)
            
            If sOutMsg <> "" Then
                LogMessage strProcName, "ERROR", "Problem setting required claims - look in usp_SetTaggedClaimNumException for concept: " & Me.ConceptID, sOutMsg
                GoTo Block_Exit
            End If
        End With
    
    Next
    ' If we get here, we can assume success
    SetRequiredClaimsNum = True

Block_Exit:
    Set colPayers = Nothing
    Set oPayer = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    SetRequiredClaimsNum = False
    GoTo Block_Exit
End Function


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
        .sqlString = "usp_EracNumOfClaimsRequired"
        .Parameters("@pConceptId") = Me.ConceptID
        .Parameters("@pRequirementId") = RequirementRuleObj.Id
        
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Problem finding required claims - look in usp_EracNumOfClaimsRequired", "Req Rule ID: " & CStr(RequirementRuleObj.Id) & " " & Me.ConceptID
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

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetMissingRequiredDocsRS(Optional ByRef sReport As String) As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sMsg As String
Dim saryRows() As String
Dim iRow As Integer

    strProcName = ClassName & ".GetMissingRequiredDocsRS"

    Set oRs = New ADODB.RecordSet
    oRs.ActiveConnection = Nothing
    oRs.LockType = adLockBatchOptimistic
    oRs.CursorLocation = adUseClient
    
        ' Add our Field name
    oRs.Fields.Append "Description", adLongVarChar
    ' kd, didn't do this yet..

    sMsg = Me.GetMissingRequiredDocsMessage(Nothing)
    If sMsg = "" Then
        sMsg = "No documents missing!"
    End If
    
    saryRows = Split(sMsg, vbCrLf)
    
    For iRow = 0 To UBound(saryRows)
        oRs.AddNew
        oRs.Fields(0) = saryRows(iRow)
        oRs.Update
    Next

Block_Exit:
    Set GetMissingRequiredDocsRS = oRs
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
Public Function HasRequiredFields(Optional ByRef sReport As String, Optional oPayer As clsConceptPayerDtl) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim saryReqdFields() As String
Dim iIdx As Integer

    strProcName = ClassName & ".HasRequiredFields"
    HasRequiredFields = True
    saryReqdFields = Split(csREQUIRED_CONCEPT_HDR_FIELDS, ",")
    
    For iIdx = 0 To UBound(saryReqdFields)
        If CStr("" & Me.GetField(saryReqdFields(iIdx))) = "" Then
            '' If it's null, then we look at the payer
            If IsMissing(oPayer) = True Then
                HasRequiredFields = False
                sReport = sReport & saryReqdFields(iIdx) & " is missing" & vbCrLf
            Else
                If CStr("" & oPayer.GetField(saryReqdFields(iIdx))) = "" Then
                    HasRequiredFields = False
                    sReport = sReport & saryReqdFields(iIdx) & " is missing" & vbCrLf
                End If
            End If
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

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetMissingRequiredDocsMessage(oPayer As clsConceptPayerDtl) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oReqDocType As clsConceptReqDocType
Dim oAttachedDoc As clsConceptDoc
Dim bAtLeastOneNotFound As Boolean
Dim iFilesNeeded As Integer
Dim iFilesFound As Integer
Dim oTagdClaim As clsEracClaim
Dim SFileName As String
Dim sFolderPath As String
Dim iEracClaimId As Integer
Dim sMsg As String

    strProcName = ClassName & ".GetMissingRequiredDocsMessage"
    
    If coReqRule.RequiredDocs Is Nothing Then
        sMsg = "No required documents"
        GoTo Block_Exit
    End If
    
    For Each oReqDocType In coReqRule.RequiredDocs
        iFilesNeeded = 0    ' Just make sure it's reset
            
            '' Is it a concept level document or a payer level doc? each need to be treated differently
        If oReqDocType.IsPayerDoc = False Then
            If ValidatePackageLevelDoc(oReqDocType, sReturn) = False Then
                bAtLeastOneNotFound = True
                coValidateRpt.AddNote False, "Required Package Document " & oReqDocType.DocName, oPayer.PayerName, oReqDocType.DocName & " not found"
            End If
        
        Else    ' must be a payer level doc
                ' we are supposed to have oReqDocType.NumPerClaim of these..
            
            If ValidatePayerLevelDoc(oReqDocType, oPayer, sReturn) = False Then
                bAtLeastOneNotFound = True
                coValidateRpt.AddNote False, "Required Claim Document " & oReqDocType.DocName, oPayer.PayerName, oReqDocType.DocName & " not found"

            End If
        End If
        
    
        If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
            ' we don't care if we have more - do we?
            ' also, if we are going to make it, then we don't care..
            If Nz(oReqDocType.CreateFunctionName, "") = "" Then
                bAtLeastOneNotFound = True

                coValidateRpt.AddNote False, "Required Package Document " & oReqDocType.DocName, oPayer.PayerName, oReqDocType.DocName & " not found, dup?"
            Else

            End If
        End If
'NextRequiredDocType:
    Next
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
    

Block_Exit:
    GetMissingRequiredDocsMessage = sReturn
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    sReturn = sReturn & "ERROR: " & Err.Description & vbCrLf
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function ValidatePackageLevelDoc(oReqDocType As clsConceptReqDocType, sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oAttachedDoc As clsConceptDoc
Dim bAtLeastOneNotFound As Boolean
Dim iFilesNeeded As Integer
Dim iFilesFound As Integer
Dim oTagdClaim As clsEracClaim
Dim SFileName As String
Dim sFolderPath As String
Dim iEracClaimId As Integer
Dim sMsg As String
Dim oAtchdDoc As clsConceptDoc
Dim bFoundIt As Boolean
                                        
    strProcName = ClassName & ".ValidatePackageLevelDoc"

        ' If we are going to create it on submission,
        ' then we need to see if we have the 'material' we'll need
        ' to do so we use the 'CheckExistanceSQL' query
    
    iFilesNeeded = oReqDocType.NumPerConcept
    
    sFolderPath = Me.ConceptWorkFolder
    
    If oReqDocType.CheckExistanceSQL <> "" Then
        If AllMaterialsExistToCreateDoc(oReqDocType, iFilesNeeded, sMsg) = False Then
            sReport = sReport & sMsg & vbCrLf
        End If
    Else    'If oReqDocType.CreateFunctionName = "" Then

            '' Since I'll be creating the NIRF, we need to add a check here to ignore files
            '' that don't exist but have a create function
        If iFilesNeeded > 0 Then
            SFileName = oReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId(0), "")
            If FileExists(sFolderPath & SFileName & "." & oReqDocType.SendAsFileType) = False And oReqDocType.CreateFunctionName = "" Then

                '' KD COMEBACK: if the converted document doesn't exist then look through the attached docs for it..
                
                For Each oAtchdDoc In Me.AttachedDocuments
                    If oAtchdDoc.CnlyAttachType = oReqDocType.CnlyAttachType Then
                        coValidateRpt.AddNote True, "Concept Documents", "", oReqDocType.DocName & " is present!"
                        bFoundIt = True
                        GoTo Block_Exit
                    End If
                
'                    If left(LCase(oAtchdDoc.FileName), Len(SFileName)) = LCase(SFileName) Then
'                        coValidateRpt.AddNote True, "Concept Documents", "", oReqDocType.DocName & " is present!"
'                        bFoundIt = True
'                        GoTo Block_Exit
'                    End If
                Next
                
                If bFoundIt = False Then
                    sReport = sReport & oReqDocType.DocName & " is missing" & vbCrLf
                    If Not coValidateRpt Is Nothing Then
                        coValidateRpt.AddNote False, "Concept Documents", "", oReqDocType.DocName & " is missing"
                        
                        bAtLeastOneNotFound = True
                    End If
                End If
            End If
        End If
    End If

    If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
            ' we don't care if we have more - do we?
            ' also, if we are going to make it, then we don't care..
        If Nz(oReqDocType.CreateFunctionName, "") = "" Then
            bAtLeastOneNotFound = True
'            If Not coValidateRpt Is Nothing Then
'                coValidateRpt.AddNote False, "Concept Documents", "Required " & CStr(iFilesNeeded) & " but only have " & CStr(CountReqDocsOfType(oReqDocType))
'            End If
        End If
    End If


Block_Exit:
    ValidatePackageLevelDoc = Not bAtLeastOneNotFound
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    sReport = sReport & "ERROR: " & Err.Description & vbCrLf
    GoTo Block_Exit
End Function

'
'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Private Function ValidatePayerLevelDoc(oReqDocType As clsConceptReqDocType, oPayer As clsConceptPayerDtl, sReport As String) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
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
'    strProcName = ClassName & ".ValidatePayerLevelDoc"
'
'        ' If it's a claim level doc
'
'        '' Where are we looking for the files?
'        '' for medical claims, we are doing that ourselves (code) so,
'        '' we could look in the work folder for that concept
'    If oReqDocType.CreateFunctionName <> "" Then     '' "Medical Record/Documentation" for example NIRF
'        sFolderPath = Me.ConceptWorkFolder
'    Else
'        sFolderPath = Me.ConceptFolder
'    End If
'
''Stop
'
'        ''    iFilesNeeded = oReqDocType.NumPerClaim * Me.TaggedClaims.Count
'    '' This is a bug. We can't use the Me.TaggedClaims.Count because we may not have
'    '' all of the tagged claims that we need.. So
'    iFilesNeeded = oReqDocType.NumPerPayer  '   * Me.RequiredClaimsNum
'
'        ' If we are creating them then we should check to see if we've got the stuff we need
'        ' to create them...:
'    If oReqDocType.CreateFunctionName <> "" Then
'
'        ' b) if we have the 'Materials' we need to create them
'        ' KD COMEBACK
'
'        If AllMaterialsExistToCreateDoc(oReqDocType, iFilesNeeded, sMsg) = False Then
'            bAtLeastOneNotFound = True
'            sReport = sReport & sMsg & vbCrLf
'
'            If Not coValidateRpt Is Nothing Then
'                coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "At least 1 " & oReqDocType.DocName & " is missing for this concept!"
'            End If
'        End If
'        ' a) if we've already created them - eh, who cares.. :P
'
'    Else
'
'        ' we are supposed to have oReqDocType.NumPerClaim of these..
'        ' Look through each of the claims, and see which ones don't have the current document type
'        '' KD COMEBACK: Ok, this is a good idea BUT
'Stop
'        '' First, we have the NEW stuff which is going to have an eRacTaggedClaimId
'        For Each oTagdClaim In ccolTaggedClaims
'
'            iEracClaimId = oTagdClaim.eRacTaggedClaimId '  GetField("eRacTaggedClaimId")
'
'            If iEracClaimId > 0 Then    '' we have one.. it's a "New system" link
'                SFileName = GetAttachmentPathFromEracTgdClaimId(iEracClaimId)
'                If SFileName = "" Then
'                    ' KD COMEBACK Remove this
'                    sReport = sReport & "The " & oReqDocType.DocName & _
'                            " attached doc for claim: " & oTagdClaim.ICN & " is missing" & vbCrLf
'
'                    If Not coValidateRpt Is Nothing Then
'                        coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "The " & oReqDocType.DocName & _
'                                " attached doc for claim: " & oTagdClaim.ICN & " is missing"
'                    End If
'
'                ElseIf FileExists(sFolderPath & SFileName) = False Then
'                    ' KD COMEBACK Remove this
'                    sReport = sReport & "The " & oReqDocType.DocName & " file for claim: " & _
'                            oTagdClaim.ICN & " is missing" & vbCrLf
'                    If Not coValidateRpt Is Nothing Then
'                        coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "The " & oReqDocType.DocName & " file for claim: " & _
'                            oTagdClaim.ICN & " is missing"
'                    End If
'
'                End If
'            Else    '' It's an "old system" link LEGACY
'                    '' If it doesn't have the eRacTaggedClaimId so we need to see if the collection contains
'                    '' the doc by getting the parsed filename (without extension)
'                    '' then checking the attached docs collection
'                    '' for that filename (less the extension)
'                    '' not exactly precise
'                SFileName = oReqDocType.ParseFileName(Me.ConceptID, Me.ClientIssueId, oTagdClaim.ICN)
'
'                SFileName = GetAttachmentPathFromParsedName(SFileName)
'
'                If SFileName = "" Then
'                    ' KD COMEBACK Remove this
'                    sReport = sReport & "The " & oReqDocType.DocName & " attached doc for claim: " & _
'                            oTagdClaim.ICN & " is missing" & vbCrLf
'
'                    If Not coValidateRpt Is Nothing Then
'                        coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "The " & oReqDocType.DocName & " attached doc for claim: " & _
'                            oTagdClaim.ICN & " is missing"
'                    End If
'                ElseIf FileExists(sFolderPath & SFileName) = False Then
'                    ' KD COMEBACK Remove this
'                    sReport = sReport & "The " & oReqDocType.DocName & " file for claim: " & _
'                            oTagdClaim.ICN & " is missing" & vbCrLf
'                    If Not coValidateRpt Is Nothing Then
'                        coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "The " & oReqDocType.DocName & " file for claim: " & _
'                            oTagdClaim.ICN & " is missing"
'                    End If
'                End If
'            End If
'
'        Next
'
'
'
'
'    End If
'
'
'    If CountReqDocsOfType(oReqDocType) < iFilesNeeded Then
'        ' we don't care if we have more - do we?
'        ' also, if we are going to make it, then we don't care..
'        If Nz(oReqDocType.CreateFunctionName, "") = "" Then
'            bAtLeastOneNotFound = True
'        End If
'    End If
'
'
'Block_Exit:
'    ValidatePayerLevelDoc = Not bAtLeastOneNotFound
'    Exit Function
'Block_Err:
'    FireError Err, strProcName, "User ID: " & Identity.Username() & " " & csConceptId
'    sReport = sReport & "ERROR: " & Err.Description & vbCrLf
'    GoTo Block_Exit
'End Function
'

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''  NEW Version
'''
Private Function ValidatePayerLevelDoc(oReqDocType As clsConceptReqDocType, oPayer As clsConceptPayerDtl, sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttachedDoc As clsConceptDoc
Dim bAtLeastOneNotFound As Boolean
Dim iFilesNeeded As Integer
Dim iFilesFound As Integer
Dim oTagdClaim As clsEracClaim
Dim SFileName As String
Dim sFolderPath As String
Dim iEracClaimId As Integer
Dim sMsg As String

    strProcName = ClassName & ".ValidatePayerLevelDoc"
    
        ' If it's a claim level doc
        
        '' Where are we looking for the files?
        '' for medical claims, we are doing that ourselves (code) so,
        '' we could look in the work folder for that concept
    If oReqDocType.CreateFunctionName <> "" Then     '' "Medical Record/Documentation" for example NIRF
        sFolderPath = Me.ConceptWorkFolder
    Else
        sFolderPath = Me.ConceptFolder
    End If

        '' How many files do we need for this payer?
    iFilesNeeded = oReqDocType.NumPerPayer  ' Next commented line was for the old way where documents were related to a sample claim:
                                ' * Me.RequiredClaimsNum
        
        ' If we are creating them then we should check to see if we've got the stuff we need
        ' to create them...:
    If oReqDocType.CreateFunctionName <> "" Then

        ' b) if we have the 'Materials' we need to create them

        If AllMaterialsExistToCreateDoc(oReqDocType, iFilesNeeded, sMsg) = False Then
            bAtLeastOneNotFound = True
            sReport = sReport & sMsg & vbCrLf
            
            If Not coValidateRpt Is Nothing Then
                coValidateRpt.AddNote False, "Claim Document", oPayer.PayerName, "At least 1 " & oReqDocType.DocName & " is missing for this concept!"
            End If
        End If
        ' a) if we've already created them - eh, who cares.. :P
        
    End If


        '' Ok, look through the records in CONCEPT_References (the attached documents)
        '' and count how many we have for this type (and payer)
    If iFilesNeeded > 0 Then
        If CountReqDocsOfType(oReqDocType, oPayer.PayerNameId) < iFilesNeeded Then
            ' we don't care if we have more - do we?
            ' also, if we are going to make it, then we don't care..
            If Nz(oReqDocType.CreateFunctionName, "") = "" Then
                bAtLeastOneNotFound = True
            End If
        End If
    End If

Block_Exit:
    ValidatePayerLevelDoc = Not bAtLeastOneNotFound
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    sReport = sReport & "ERROR: " & Err.Description & vbCrLf
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetAttachmentPathFromEracTgdClaimId(iEracTgdClaimId As Integer) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc

    strProcName = ClassName & ".GetAttachmentPathFromEracTgdClaimId"

    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
    
    For Each oAttach In ccolAttachedDocs
        If oAttach.eRacTaggedClaimId = iEracTgdClaimId Then
            GetAttachmentPathFromEracTgdClaimId = oAttach.RefFileName
            Exit For
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
Public Function NIRF_Exists(Optional lPayerNameId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc
    
    strProcName = ClassName & ".NIRF_Exists"
    
    Call LoadSubObjects("LOADATTACHEDDOCS")
    
    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
    
    For Each oAttach In ccolAttachedDocs
        If lPayerNameId <> 0 And lPayerNameId <> 1000 Then
            If oAttach.PayerNameId = lPayerNameId Then
                If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
                    NIRF_Exists = True
                    GoTo Block_Exit
                End If
            End If
        Else
            If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
                NIRF_Exists = True
                GoTo Block_Exit
            End If
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
Public Function NIRF_Path(lPayerNameId As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc

    strProcName = ClassName & ".NIRF_Path"
    
    Call LoadSubObjects("LOADATTACHEDDOCS")
    
    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
    For Each oAttach In ccolAttachedDocs
        If lPayerNameId <> 0 And lPayerNameId <> 1000 Then
            If oAttach.PayerNameId = lPayerNameId Then
                If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
                    NIRF_Path = oAttach.RefFullPath
                    Stop
                    GoTo Block_Exit
                End If
            End If
        Else
            If oAttach.GetEracReqDocType.CnlyAttachType = "ERAC_NIRF" Then
                NIRF_Path = oAttach.RefFullPath
                Stop
                GoTo Block_Exit
            End If
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
''' NOTE: THis is for LEGACY attached documents
''' since they won't be named correctly
'''
Public Function GetAttachmentPathFromParsedName(ByVal SFileName As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc
Dim iFileLen As Integer

    strProcName = ClassName & ".GetAttachmentPathFromParsedName"
        ' Insure no period and extension:
    If InStr(1, SFileName, ".", vbTextCompare) > 0 Then
        SFileName = left(SFileName, InStr(1, SFileName, ".", vbTextCompare) - 1)
    End If
        ' Now, since people save stuff like screen prints with JUST the ICN:
    If InStr(1, SFileName, "_", vbTextCompare) > 0 Then
        SFileName = left(SFileName, InStr(1, SFileName, "_", vbTextCompare) - 1)
    End If
    
    If SFileName = "" Then GoTo Block_Exit
    SFileName = UCase(SFileName)
    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
    iFileLen = Len(SFileName)
    
    For Each oAttach In ccolAttachedDocs
            '' a little redundancy here..
        If left(UCase(oAttach.RefFileNameNoExt), iFileLen) = SFileName Then
            GetAttachmentPathFromParsedName = oAttach.RefFileName
            Exit For
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
Public Function GetAttachmentPathFromICN(sIcn As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAttach As clsConceptDoc

    strProcName = ClassName & ".GetAttachmentPathFromICN"
    If sIcn = "" Then GoTo Block_Exit
    If ccolAttachedDocs.Count = 0 Then GoTo Block_Exit
    
    For Each oAttach In ccolAttachedDocs
        If oAttach.Icn = sIcn Then
            GetAttachmentPathFromICN = oAttach.RefFileName
            Exit For
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
''' Is the status 3 (or 7)
'''
Private Function ConvertClaims_Code_To_EracCode(sFieldName As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sVal As String

    strProcName = ClassName & ".ConvertClaims_Code_To_EracCode"
    
    sVal = CStr("" & GetField(Right(sFieldName, Len(sFieldName) - 5)))    ' - 5 because it begins with CNVT_ - short for convert
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "SELECT ReviewTypeName As ReviewType FROM XrefReviewType WHERE CnlyReviewTypeCode = '" & sVal & "'"
        Set oRs = .ExecuteRS
        If .GotData = False Then
            ConvertClaims_Code_To_EracCode = sVal
            GoTo Block_Exit
        End If
        sVal = CStr("" & oRs(Right(sFieldName, Len(sFieldName) - 5)).Value)
    End With
    
    
    ConvertClaims_Code_To_EracCode = sVal
Block_Exit:
    Set oAdo = Nothing
    Set oRs = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    ConvertClaims_Code_To_EracCode = ""
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
''' Is the status 3 (or 7)
'''
Public Function IsStatusOkForEracSubmission(Optional ByRef sErrMsg As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".IsStatusOkForEracSubmission"
    
    Set oRs = GetRecordsetSP("usp_EracGetNotificationHistDesc", "@pConceptId=" & Me.ConceptID)
    If RSHasData(oRs) = False Then
        sErrMsg = "No records found for this concept"
        GoTo Block_Exit
    End If
    
        '' This will likely have more than 1 record but it's ordered in desc order
        '' so we only need to look at the first to see the current notification status
    If oRs("NotificationType").Value <> 3 And oRs("NotificationType").Value <> 7 Then
        sErrMsg = "Current notification type is: " & CStr(oRs("NotificationType").Value) & " (" & Nz(oRs("NotificationTypeName"), "") & ")"
        GoTo Block_Exit
    End If
    
    IsStatusOkForEracSubmission = True
Block_Exit:
    Set oRs = Nothing
    
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    IsStatusOkForEracSubmission = False
    sErrMsg = sErrMsg & " " & Err.Description
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function MakeWorkCopiesOfFiles() As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".MakeWorkCopiesOfFiles"
    '' KD COMEBACK: DO THIS!

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    MakeWorkCopiesOfFiles = False
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function ParseStringForDetails(sInString As String, Optional lPayerNameId As Long) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oRegEx As RegExp
Dim oMatches As MatchCollection
Dim oMatch As Match
Dim sConceptVal As String
Dim sPayerVal As String
Dim sValue As String
Dim oPayer As clsConceptPayerDtl


    strProcName = ClassName & ".ParseStringForDetails"

    If lPayerNameId <> 0 And lPayerNameId <> 1000 Then
        For Each oPayer In Me.ConceptPayers
            If oPayer.PayerNameId = lPayerNameId Then
                Exit For
            End If
        Next
    End If

    Set oRegEx = New RegExp
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = "\[\*([^\*\]]+)\*\]"
    oRegEx.Global = True
    oRegEx.MultiLine = True
    
    Set oMatches = oRegEx.Execute(sInString)
    
    If oMatches.Count = 0 Then
        ParseStringForDetails = sInString
        GoTo Block_Exit
    End If
    
    ParseStringForDetails = sInString
    
    For Each oMatch In oMatches
        If left(oMatch.SubMatches(0), 5) = "CNVT_" Then
            ' The code needs to be converted..
            sConceptVal = ConvertClaims_Code_To_EracCode(oMatch.SubMatches(0))
            
        Else
            sPayerVal = oPayer.GetField(oMatch.SubMatches(0))
            sConceptVal = CStr("" & GetField(oMatch.SubMatches(0)))
            
            If sPayerVal <> "" Then
                sValue = sPayerVal
            Else
                sValue = sConceptVal
            End If
            
        End If
    
        
    
        
        ParseStringForDetails = Replace(ParseStringForDetails, "[*" & oMatch.SubMatches(0) & "*]", sValue, , , vbTextCompare)
    Next

Block_Exit:
    Set oMatch = Nothing
    Set oMatches = Nothing
    Set oRegEx = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & csConceptId
    ParseStringForDetails = sInString
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' used to be: AllDocExistsForConcept
Public Function AllMaterialsExistToCreateDoc(oRequiredDoc As clsConceptReqDocType, iRequiredNum As Integer, sReport As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String

    strProcName = ClassName & ".AllMaterialsExistToCreateDoc"

    '' 20120719 KD Note:
    '' This was originally written for things like getting the medical charts
    '' but now we don't have anything that requires any documents - the only 2 we are creating
    '' now is the Nirf and the Sample Claim IDs document (which DOES need tagged claims but we already
    ''  validated that)
    '' SO, for now, we are just going to let this puppy be true!
    AllMaterialsExistToCreateDoc = True
    GoTo Block_Exit

    If oRequiredDoc.CheckExistanceSQL = "" Then GoTo Block_Exit
        
    sSql = Replace(oRequiredDoc.CheckExistanceSQL, "?", Me.ConceptID, , , vbTextCompare)
    
    Set oRs = GetRecordset(sSql, "ConceptDocTypes")
    
    If oRs Is Nothing Then
        AllMaterialsExistToCreateDoc = False
        sReport = sReport & "No items found!" & vbCrLf
        If Not coValidateRpt Is Nothing Then
            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, "", "No " & oRequiredDoc.DocName & " documents found!"
        End If
        GoTo Block_Exit
    End If
    If oRs.recordCount < 1 Then
        sReport = sReport & "No items found!" & vbCrLf

        If Not coValidateRpt Is Nothing Then
            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, "", "No " & oRequiredDoc.DocName & " documents found!"
        End If
        AllMaterialsExistToCreateDoc = False
        GoTo Block_Exit
    End If
    
    '' Do we have as many as we were expecting - well, that's for another function isn't it?
    '' and of course that doesn't belong in this object (heck, this function is debatable)
    If oRs.recordCount < iRequiredNum Then
        
        sReport = sReport & CStr(Me.RequiredClaimsNum - oRs.recordCount) & " items were MISSING (" & _
            "found " & CStr(oRs.recordCount) & ")"
        If Not coValidateRpt Is Nothing Then
            coValidateRpt.AddNote False, "Required documents " & oRequiredDoc.DocName, "", CStr(Me.RequiredClaimsNum - oRs.recordCount) & " items were MISSING (" & _
                    "found " & CStr(oRs.recordCount) & ")"
        End If
        AllMaterialsExistToCreateDoc = False
    End If
    

Block_Exit:
    Set oRs = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & Me.ConceptID
    sReport = sReport & Err.Description & vbCrLf
    AllMaterialsExistToCreateDoc = False
    GoTo Block_Exit
End Function



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
Public Function LoadFromId(sConceptId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = True
    Id = sConceptId
    LoadFromId = coSourceTable.LoadFromIDStr(sConceptId)
    WasInitialized = LoadFromId
'
'        ' Get the requirement rule object...
'    If GetReqRule() = False Then
'        ' KD COMEBACK: Deal with this
'    End If
'
'    If LoadPayerDetails() = False Then
'        ' then what? then it's an old one!
'    End If
'
'        ' stuff the attached documents into a collection
'    If LoadAttachedDocs() = False Then
'        ' it's already been logged.. just let it continue as no attached docs isn't a problem
'        ' especially for a new concept
'    End If
'
'        ' get some basic details of the tagged claims..
'    If LoadTaggedClaims() = False Then
'        ' it's already been logged, so we'll let this one continue too as not all concepts
'        '' will have tagged claims - and especially when they are new
'    End If
'
Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "Concept: " & sConceptId
    GoTo Block_Exit
End Function


Public Function RefreshObject() As Boolean
Dim sConceptId As String

    sConceptId = Me.ConceptID

    Call Class_Initialize
    
    Call LoadFromId(sConceptId)
    
    Call LoadSubObjects("")
    
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function LoadSubObjects(sObjsToLoad As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadSubObjects"
    'LogMessage strProcName, , "Function started"

    If sObjsToLoad = "" Then
        '' Everything is requested.. so give it to them:
        LoadSubObjects = LoadSubObjects("GETREQRULE")
        LoadSubObjects = LoadSubObjects("LOADPAYERDETAILS")
        LoadSubObjects = LoadSubObjects("LOADATTACHEDDOCS")
        LoadSubObjects = LoadSubObjects("LOADTAGGEDCLAIMS")
        GoTo Block_Exit
    End If

    Select Case UCase(sObjsToLoad)
    Case "GETREQRULE", ""
        If cdctInitObjs.Exists("GetReqRule") = True Then
            If cdctInitObjs.Item("GetReqRule") = False Then
                    ' Get the requirement rule object...
                If GetReqRule() = False Then
                    ' KD COMEBACK: Deal with this
                    cdctInitObjs.Item("GetReqRule") = False
                    GoTo Block_Exit
                Else
                    cdctInitObjs.Item("GetReqRule") = True
                End If
            End If
        End If
    Case "LOADPAYERDETAILS", ""
        If cdctInitObjs.Exists("LoadPayerDetails") = True Then
            If cdctInitObjs.Item("LoadPayerDetails") = False Then
                    ' Get the requirement rule object...
                If LoadPayerDetails() = False Then
                    ' then what? then it's an old one!
                    cdctInitObjs.Item("LoadPayerDetails") = True
                Else
                    cdctInitObjs.Item("LoadPayerDetails") = True
                End If
    
            End If
        End If
    Case "LOADATTACHEDDOCS", ""
        If cdctInitObjs.Exists("LoadAttachedDocs") = True Then
            If cdctInitObjs.Item("LoadAttachedDocs") = False Then
            '        ' stuff the attached documents into a collection
                If LoadAttachedDocs() = False Then
                    ' it's already been logged.. just let it continue as no attached docs isn't a problem
                    ' especially for a new concept
                    cdctInitObjs.Item("LoadAttachedDocs") = True
                Else
                    cdctInitObjs.Item("LoadAttachedDocs") = True
                End If
    
            End If
        End If
    Case "LOADTAGGEDCLAIMS", ""
        If cdctInitObjs.Exists("LoadTaggedClaims") = True Then
            If cdctInitObjs.Item("LoadTaggedClaims") = False Then
            '        ' get some basic details of the tagged claims..
                If LoadTaggedClaims() = False Then
                    ' it's already been logged, so we'll let this one continue too as not all concepts
                    '' will have tagged claims - and especially when they are new
                    cdctInitObjs.Item("LoadTaggedClaims") = True
                Else
                    cdctInitObjs.Item("LoadTaggedClaims") = True
                End If
            End If
        End If
    Case Else
        LoadSubObjects = False
        Stop
        GoTo Block_Exit
    End Select


    
    LoadSubObjects = True
Block_Exit:
    Exit Function

Block_Err:
    LoadSubObjects = False
    FireError Err, strProcName, "Concept: " & Me.ConceptID
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function LoadPayerDetails() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oPayerDet As clsConceptPayerDtl

    strProcName = ClassName & ".LoadAttachedDocs"

    Set ccolPayerDetails = New Collection

    sSql = "SELECT * FROM CONCEPT_PAYER_Dtl WHERE ConceptID = '" & Me.Id & "' "
    Set oRs = GetRecordset(sSql, "V_Data_Database")
    If oRs Is Nothing Then
        LogMessage strProcName, "WARNING", "Old concept - no payer details!", Me.Id
        GoTo Block_Exit
    End If


    While Not oRs.EOF
        If oRs("PayerNameID") <> 1000 Then  ' skip All
            Set oPayerDet = New clsConceptPayerDtl
            
            If oPayerDet.LoadFromId(oRs("ConceptIDPayerID_RowID")) = False Then
                LogMessage strProcName, "WARNING", "Problem loading an attached file!", "RowID: " & CStr(oRs("RowId").Value)
            Else
                ccolPayerDetails.Add oPayerDet
            End If
        
        End If
        
        oRs.MoveNext
    Wend

    LoadPayerDetails = True
    
Block_Exit:
    Set oRs = Nothing
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    LoadPayerDetails = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
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

'    sSql = "SELECT * FROM V_CONCEPT_References WHERE ConceptID = '" & Me.ID & "' AND SendToPayers = 1 ORDER BY RefSequence ASC "
    sSql = "SELECT * FROM V_CONCEPT_References WHERE ConceptID = '" & Me.Id & "' ORDER BY RefSequence ASC "
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
'Stop ' kd: didn't do this yet.
            If oAttachedFile.GetEracReqDocType.IsPayerDoc = False Then
                ccolHdrAttached.Add oAttachedFile
            Else
                ' must be detail..
                ccolDtlAttached.Add oAttachedFile
            End If
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

'    sSql = "SELECT h.CnlyClaimNum, ICN, MedicalRecordNum " & _
'            " FROM CMS_AUDITORS_CLAIMS.dbo.AuditClm_Hdr H INNER JOIN ( " & _
'                " SELECT cnlyclaimnum, Adj_ConceptID as 'ConceptID' FROM CMS_AUDITORS_CODE.dbo.v_CONCEPT_ValidationSummary " & _
'            " ) as AA ON AA.CnlyClaimNum = h.cnlyClaimNum " & _
'            " WHERE AA.ConceptID = '" & Me.ID & "'"
           
           
    sSql = "SELECT h.CnlyClaimNum, h.ICN, NULL as MedicalRecordNum FROM CMS_AUDITORS_CODE.dbo.v_CONCEPT_TaggedClaims h WHERE h.ConceptId = '" & Me.Id & "' "
    
    
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
'Private Function LoadTaggedClaims() As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oRs As ADODB.Recordset
'Dim sSql As String
'Dim oClaim As clsEracClaim
'
'
'    strProcName = ClassName & ".LoadTaggedClaims"
'
'    sSql = "SELECT h.CnlyClaimNum, ICN, MedicalRecordNum " & _
'            " FROM CMS_AUDITORS_CLAIMS.dbo.AuditClm_Hdr H INNER JOIN ( " & _
'                " SELECT cnlyclaimnum, Adj_ConceptID as 'ConceptID' FROM CMS_AUDITORS_CODE.dbo.v_CONCEPT_ValidationSummary " & _
'            " ) as AA ON AA.CnlyClaimNum = h.cnlyClaimNum " & _
'            " WHERE AA.ConceptID = '" & Me.ID & "'"
'
'    Set oRs = GetRecordset(sSql, "AuditClm_Hdr")
'    If oRs Is Nothing Then
''        LogMessage strProcName, "WARNING", "Either no tagged claims or there was a problem with the query / connection!", Me.ID
'        GoTo Block_Exit
'    End If
'
'
'    While Not oRs.EOF
'        Set oClaim = New clsEracClaim
'        If oClaim.LoadFromID(CStr("" & oRs("CnlyClaimNum").Value)) = False Then
'            ' KD COMEBACK deal with this
'        End If
'            ' just stuff that in our collection
'        ccolTaggedClaims.Add oClaim
'
'        oRs.MoveNext
'    Wend
'
'    LoadTaggedClaims = True
'
'Block_Exit:
'    Set oRs = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    Err.Clear
'    LoadTaggedClaims = False
'    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
'End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetReqRule() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim sException As String

    strProcName = ClassName & ".GetReqRule"

    '' KD COMEBACK: When we have some Data type stuff in there then this needs to be dealt with! (or is this going to work?)
'    sSql = "SELECT ConceptReqId FROM CnlyConceptRequirements WHERE ReviewTypeId = " & _
'            Me.EracReviewTypeId & " AND ISNULL(DataTypeCode,'') = '" & Me.CnlyDataTypeCode & "' "
            
    sSql = "SELECT ConceptReqId, DataTypeCode FROM CnlyConceptRequirements WHERE ReviewTypeId = " & _
            Me.EracReviewTypeId & " AND ( ISNULL(DataTypeCode,'') = '" & Me.CnlyDataTypeCode & "' OR ISNULL(DataTypeCode,'') = '' ) "
            
            '' KD COMEBACK: Note, if we get 2 then we need to use the one that has the DataTypeCode that isn't ''
            
    sException = " AND ( ISNULL(ExceptionLOB,'') = '" & Me.GetField("LOB") & "' OR ISNULL(ExceptionAuditor,'') = '" & _
            Me.GetField("Auditor") & "') "
        
        '' First see if we have one with an exception
    Set oRs = GetRecordset(sSql & sException, "CnlyConceptRequirements")
    If Not oRs Is Nothing Then
        If oRs.recordCount > 1 Then
            Do While Not oRs.EOF
                If Nz(oRs("DataTypeCode"), "") = Me.CnlyDataTypeCode Then
                    ' this is our record..
                    Exit Do
                End If
                oRs.MoveNext
            Loop
        End If
        
        Set coReqRule = New clsEracRequirementRule
        GetReqRule = coReqRule.LoadFromId(CInt(oRs("ConceptReqId").Value))
        GoTo Block_Exit
    End If
    
        '' Without the exception
    Set oRs = GetRecordset(sSql, "CnlyConceptRequirements")
    If oRs Is Nothing Then GoTo Block_Exit
    
    If Not oRs.EOF Then
        If oRs.recordCount > 1 Then
            Do While Not oRs.EOF
                If Nz(oRs("DataTypeCode"), "") = Me.CnlyDataTypeCode Then
                    ' this is our record..
                    Exit Do
                End If
                oRs.MoveNext
            Loop
        End If
        
        Set coReqRule = New clsEracRequirementRule
        GetReqRule = coReqRule.LoadFromId(CInt(oRs("ConceptReqId").Value))
        ''  GetReqRule = coReqRule.LoadFromConceptID(Me.ID)

    End If

    
Block_Exit:
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GetReqRule = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This function performs (or orchastrates) all validation  on a particular concept
''' to see if it's ready to submit to CMS
'''
Public Function GetConceptHeaderDetails(sConceptId As String, Optional ByRef sReviewTypeCode As String = "", _
    Optional ByRef sDataTypeCode As String = "", Optional ByRef sConceptOwner As String = "", _
    Optional ByRef sClientIssueNum As String = "", Optional iCmsReviewTypeId As Integer = 0) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetConceptReviewType"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = "SELECT CH.Auditor ConceptOwner, CH.ClientIssueNum, CH.ReviewType, CH.DataType FROM " & _
            " CMS_AUDITORS_CLAIMS.dbo.Concept_Hdr CH WHERE Ch.ConceptID = '" & sConceptId & "'"
    End With

    Set oRs = oAdo.ExecuteRS
    
        '' Did we get anything?
    If oRs Is Nothing Then
        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
        GoTo Block_Exit
    End If
    
    If oRs.recordCount < 1 Then
        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
        GoTo Block_Exit
    End If
    
    If Not oRs.EOF Then
        sConceptOwner = Nz(oRs("ConceptOwner").Value, "")
        sReviewTypeCode = Nz(oRs("ReviewType").Value, "")
        sDataTypeCode = Nz(oRs("DataType").Value, "")
        sClientIssueNum = Nz(oRs("ClientIssueNum").Value, "")
            ' We don't expect more than 1, so no need to movenext
    End If

    iCmsReviewTypeId = TranslateCnlyReviewTypeToCMS(sReviewTypeCode)

    GetConceptHeaderDetails = True

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GetConceptHeaderDetails = False
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function AggregatePayerData(sFieldName As String) As Double
On Error GoTo Block_Err
Dim strProcName As String
Dim oPayerDtl As clsConceptPayerDtl
Dim dblReturn As Double

    strProcName = ClassName & ".AggregatePayerData"
    
    For Each oPayerDtl In ccolPayerDetails
        If IsNumeric(oPayerDtl.Fields(sFieldName)) = True Then
            dblReturn = dblReturn + oPayerDtl.Fields(sFieldName)
        Else
            Stop
        End If
    Next

Block_Exit:
    AggregatePayerData = dblReturn
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    AggregatePayerData = -1
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function AggregateFields(sFieldName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String
Dim oPayerDtl As clsConceptPayerDtl
Dim dblReturn As Double

    strProcName = ClassName & ".AggregatePayerData"

    If cdctAggregateFieldNames.Exists(UCase(sFieldName)) = False Then
        ' If there is only 1 payer, and the concept's stuff is null then
        ' use the payer, otherwise, what??
        Call LoadSubObjects("LOADPAYERDETAILS")
        If ccolPayerDetails.Count = 1 Then
            If CStr("" & Me.GetField(sFieldName)) = "" Then
                Set oPayerDtl = ccolPayerDetails.Item(1)
                AggregateFields = oPayerDtl.GetField(sFieldName)
                GoTo Block_Exit
            End If
        Else
            AggregateFields = ""
            GoTo Block_Exit
        End If
    Else
        AggregateFields = ""
        GoTo Block_Exit
    End If
    
    For Each oPayerDtl In ccolPayerDetails
        If IsNumeric(oPayerDtl.GetField(sFieldName)) = True Then
            dblReturn = dblReturn + Nz(oPayerDtl.GetField(sFieldName), 0)
        End If
    Next

    AggregateFields = dblReturn


Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    AggregateFields = -1
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Sub GetAggFieldNamesDict()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetAggFieldNamesDict"

    Set oRs = GetRecordset("SELECT UPPER(FieldName) as FieldName FROM XREF_CONCEPT_Aggregate_PayerFields")
    
    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount = 0 Then GoTo Block_Exit

    Set cdctAggregateFieldNames = New Scripting.Dictionary
    
    While Not oRs.EOF
        If cdctAggregateFieldNames.Exists(CStr(oRs("FieldName").Value)) Then
            Stop    ' shouldn't get here, they are unique for crying out loud!
        Else
            cdctAggregateFieldNames.Add CStr(oRs("FieldName").Value), True
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
    
    RaiseEvent ConceptError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

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
    
    Set coReqRule = Nothing
    Set ccolAttachedDocs = New Collection
    Set ccolHdrAttached = New Collection
    Set ccolDtlAttached = New Collection

    Set ccolTaggedClaims = New Collection
    Set ccolPayerDetails = New Collection
    Set coValidateRpt = New clsEracValidationRpt
    
    Set coValidateRpt = Nothing
    
    Set cdctInitObjs = New Scripting.Dictionary
    With cdctInitObjs
        .Add "GetReqRule", False
        .Add "LoadPayerDetails", False
        .Add "LoadAttachedDocs", False
        .Add "LoadTaggedClaims", False
    End With
    
    Call GetAggFieldNamesDict
        
    cblnIsInitialized = False
    
End Sub


Private Sub Class_Terminate()
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing
    Set cdctInitObjs = Nothing
    Set ccolAttachedDocs = Nothing
    Set ccolHdrAttached = Nothing
    Set ccolDtlAttached = Nothing

    Set ccolTaggedClaims = Nothing
    Set coValidateRpt = Nothing

    cblnIsInitialized = False
End Sub