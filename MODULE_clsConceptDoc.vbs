Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 06/22/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a document that is attached to a concept.  This is NOT
'''  Directly related to the ConceptDocTypes table because this data is kept
'''  in _CLAIMS.dbo.CONCEPT_References... Until we are able to update that
'''  table with the Names (and better yet, id's) of the document types
'''  This will be a little confusing.
'''
'''  Important part is that this object has 1 clsConceptReqDocType (mapping table)
'''  as a property..
'''
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  - Resequence method
'''  - USE the resequence method when Detaching it to update the table
'''
'''  HISTORY:
'''  =====================================
'''  - 06/22/2012 - Added IsPayerDoc
'''  - 04/20/2012 - added ConvertedFilePath
'''  - 03/30/2012 - Added Error Occurred and Last Error
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

Public Event ConceptDocError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean
Private csLastError As String

Private Const csIDFIELDNAME As String = "RowId"
Private Const csTableName As String = "v_CONCEPT_References"
Private coSourceTable As clsTable



Private coEracDocType As clsConceptReqDocType


Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean


Private ciRowID As Integer
'Private csCnlyReviewTypeCode As String
'Private csCnlyDataTypeCode As String
'Private ciEracReviewTypeId As Integer
Private ciPayerNameId As Integer
Private ciConceptMgmtRefId As Integer


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get DocID() As Integer
    DocID = ciRowID
End Property
Public Property Let DocID(iDocID As Integer)
    ciRowID = iDocID
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As Integer
        ID = DocID
    End Property
    Public Property Let ID(iNewId As Integer)
        DocID = iNewId
    End Property
        '' Just an alias for ease of use!
    Public Property Get RowID() As Integer
        RowID = DocID
    End Property
    Public Property Let RowID(iNewId As Integer)
        DocID = iNewId
    End Property




Public Property Get eRacTaggedClaimId() As Integer
    eRacTaggedClaimId = CInt("0" & GetTableValue("eRacTaggedClaimId"))
End Property
Public Property Let eRacTaggedClaimId(iEracTaggedClaimId As Integer)
    SetTableValue "eRacTaggedClaimId", iEracTaggedClaimId
End Property




Public Property Get Icn() As String
    Icn = CStr("" & GetTableValue("ICN"))
End Property
Public Property Let Icn(sIcn As String)
    SetTableValue "ICN", sIcn
End Property



Public Property Get ConceptID() As String
    ConceptID = GetTableValue("ConceptId")
End Property
Public Property Let ConceptID(sConceptId As String)
    SetTableValue "ConceptId", sConceptId
End Property


Public Property Get RefSequence() As Integer
    RefSequence = GetTableValue("RefSequence")
End Property
Public Property Let RefSequence(iRefSequence As Integer)
    SetTableValue "RefSequence", iRefSequence
End Property



Public Property Get CreateDt() As Date
    CreateDt = GetTableValue("CreateDt")
End Property
Public Property Let CreateDt(dCreateDt As Date)
    SetTableValue "CreateDt", CStr(dCreateDt)
End Property

    '' Note: this is kinda confusing..
    '' RefType is like DOC, PDF, TIF.. where SubType is the reason type for the file.. So we are going to make some aliases
Public Property Get RefType() As String
    RefType = GetTableValue("RefType")
End Property
Public Property Let RefType(sRefType As String)
    SetTableValue "RefType", sRefType
End Property
    Public Property Get RefFileExtension() As String
        RefFileExtension = RefType
    End Property
    Public Property Let RefFileExtension(sRefFileExtension As String)
        ' No period:
        If left(sRefFileExtension, 1) = "." Then
            sRefFileExtension = Right(sRefFileExtension, Len(sRefFileExtension) - 1)
        End If
        RefType = sRefFileExtension
    End Property


Public Property Get IsPayerDoc() As Boolean
Dim sRet As String
    sRet = Nz(GetTableValue("IsPayerDoc"), "0")
    If sRet = "" Then
        sRet = "0"
    End If
    IsPayerDoc = CBool(sRet)
End Property
Public Property Let IsPayerDoc(bIsPayerDoc As Boolean)
    SetTableValue "IsPayerDoc", IIf(bIsPayerDoc, "1", "0")
End Property


' if it's a payer doc then we'll need to be able to get the payer id
' and maybe payer name
' could be all (1000)

Public Property Get PayerNameId() As Integer
    If ciPayerNameId = 0 Then
        Call GetPayerDetails
    End If
    PayerNameId = ciPayerNameId
End Property
Public Property Let PayerNameId(iPayerNameId As Integer)
    ciPayerNameId = iPayerNameId
End Property


Public Property Get ConceptMgmtRefId() As Integer
    ConceptMgmtRefId = ciConceptMgmtRefId
End Property
Public Property Let ConceptMgmtRefId(iConceptMgmtRefId As Integer)
    ciConceptMgmtRefId = iConceptMgmtRefId
End Property




    '' This is actually the Connolly Reason Type code that we typically want
Public Property Get RefSubType() As String
    RefSubType = GetTableValue("RefSubType")
End Property
Public Property Let RefSubType(sRefSubType As String)
    SetTableValue "RefSubType", sRefSubType
End Property
    Public Property Get DocTypeName() As String
        DocTypeName = RefSubType
    End Property
    Public Property Let DocTypeName(sDocTypeName As String)
        RefSubType = sDocTypeName
    End Property
        ' For our transition to make things a bit easier to follow
    Public Property Get CnlyAttachType() As String
        CnlyAttachType = RefSubType
    End Property
    Public Property Let CnlyAttachType(sCnlyAttachType As String)
        RefSubType = sCnlyAttachType
    End Property




Public Property Get RefPath() As String
    RefPath = GetTableValue("RefPath")
End Property
Public Property Let RefPath(sRefPath As String)
    SetTableValue "RefPath", sRefPath
End Property


Public Property Get RefFileName() As String
    RefFileName = GetTableValue("RefFileName")
End Property
Public Property Let RefFileName(sRefFileName As String)
    SetTableValue "RefFileName", sRefFileName
End Property


Public Property Get RefFileNameNoExt() As String
    RefFileNameNoExt = RefFileName
    If InStr(1, RefFileNameNoExt, ".", vbTextCompare) > 0 Then
        RefFileNameNoExt = left(RefFileNameNoExt, InStr(1, RefFileNameNoExt, ".", vbTextCompare) - 1)
    End If
End Property



Public Property Get RefFullPath() As String
    RefFullPath = QualifyFldrPath(RefPath) & RefFileName
End Property


Public Property Get ImageLink() As String
    ImageLink = GetTableValue("ImageLink")
End Property
Public Property Let ImageLink(sImageLink As String)
    SetTableValue "ImageLink", sImageLink
End Property


Public Property Get RefURL() As String
    RefURL = GetTableValue("RefURL")
End Property
Public Property Let RefURL(sRefURL As String)
    SetTableValue "RefURL", sRefURL
End Property


Public Property Get GetEracReqDocType() As clsConceptReqDocType
    Set GetEracReqDocType = coEracDocType
End Property


Public Property Get LastError() As String
    LastError = csLastError
End Property
Public Property Let LastError(sErrorMessage As String)
    csLastError = sErrorMessage
    cbErrorOccurred = True
End Property


Public Property Get ErrorOccurred() As Boolean
    ErrorOccurred = cbErrorOccurred
End Property
Public Property Let ErrorOccurred(bErrorOccurred As Boolean)
    cbErrorOccurred = bErrorOccurred
End Property



Public Property Get FolderPath() As String
    FolderPath = CStr("" & GetTableValue("RefPath"))
End Property
Public Property Let FolderPath(sFolderPath As String)
    SetTableValue "RefPath", sFolderPath
End Property



Public Property Get FileName() As String
    FileName = CStr("" & GetTableValue("RefFileName"))
End Property
Public Property Let FileName(SFileName As String)
    SetTableValue "RefFileName", SFileName
End Property


Public Property Get ConvertedFilePath() As String
Dim oConcept As clsConcept
Dim sFilePath As String


    Set oConcept = New clsConcept
    If oConcept.LoadFromId(Me.ConceptID) = False Then
        Stop
    End If
    
    sFilePath = QualifyFldrPath(Me.FolderPath) & Me.FileName
    
    sFilePath = GetEracReqDocType.ParseFileName(Me.ConceptID, oConcept.ClientIssueId(0), Me.Icn, sFilePath)
    
    sFilePath = sFilePath & "." & LCase(Me.GetEracReqDocType.SendAsFileType)

    If GetEracReqDocType.IsPayerDoc Then
'        ConvertedFilePath = csCONCEPT_SUBMISSION_WORK_FLDR & Me.ConceptID & "\_BURN\" & sFilePath
        ConvertedFilePath = csCONCEPT_SUBMISSION_WORK_FLDR & Me.ConceptID & "\" & GetPayerNameFromID(Me.PayerNameId) & "\" & Me.FileName ' & "." & coEracDocType.SendAsFileType
        If UCase(Right(ConvertedFilePath, Len(coEracDocType.SendAsFileType))) <> UCase(coEracDocType.SendAsFileType) Then
            ConvertedFilePath = ConvertedFilePath & "." & coEracDocType.SendAsFileType
        End If
    Else
        'ConvertedFilePath = csCONCEPT_SUBMISSION_WORK_FLDR & Me.ConceptID & "\" & sFilePath
        ConvertedFilePath = csCONCEPT_SUBMISSION_WORK_FLDR & Me.ConceptID & "\" & Me.FileName & "." & coEracDocType.SendAsFileType
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
Public Property Let LetField(sFieldName As String, sDocName As String)
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
'''
Public Function ValidateForSubmission() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim oReqRule

    strProcName = ClassName & ".ValidateForSubmission"
'   Stop     ' hammer time!
   ' no code here yet!
    ValidateForSubmission = True

Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & CStr(ciRowID)
    ValidateForSubmission = False
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
Public Function Detach() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
'Dim oColPrms As ADODB.Parameters
Dim prm As ADODB.Parameter
Dim LocCmd As New ADODB.Command
Dim iResult As Integer
Dim strErrMsg As String
Dim oCmd As ADODB.Command
Dim oFso As Scripting.FileSystemObject

    strProcName = ClassName & ".Detach"

    Set oAdo = New clsADO
    oAdo.SQLTextType = StoredProc
    oAdo.ConnectionString = GetConnectString("v_CODE_Database")
    oAdo.sqlString = "usp_CONCEPT_References_Delete"
    
    Set oCmd = New ADODB.Command
    oCmd.ActiveConnection = oAdo.CurrentConnection
    oCmd.commandType = adCmdStoredProc
    oCmd.CommandText = "usp_CONCEPT_References_Delete"
    oCmd.Parameters.Refresh
    
    oCmd.Parameters("@pRowID") = Me.RowID
    iResult = oAdo.Execute(oCmd.Parameters)
    
    '' Now delete the file - we don't really care if this doesn't happen as long as
    '' the record is removed
    Set oFso = New Scripting.FileSystemObject
    
    If oFso.FileExists(Me.RefFullPath) Then
        oFso.DeleteFile Me.RefFullPath, True
    Else
        LogMessage strProcName, "WARNING", "Could not find file where specified!", Me.RefFullPath
    End If
    

    Detach = True
Block_Exit:
    Set oFso = Nothing
    Set oAdo = Nothing
    Set oCmd = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    Detach = False
    GoTo Block_Exit
End Function



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
Public Function GetTableValue(strFieldName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetTableValue"
    
    GetTableValue = coSourceTable.GetTableValue(strFieldName)

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
Public Function LoadFromId(lDocRowId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
'       Debug.Assert iDocRowId <> 7073
    
    ID = lDocRowId
    LoadFromId = coSourceTable.LoadFromId(lDocRowId)
    WasInitialized = LoadFromId

    ' Now, we have to get the document type object which we'll use for
    ' validations.
    ' Currently, because we are between "systems" we are using the
    ' CnlyAttachType field name - at least until we can move
    ' the rest of the concept attachment system over to
    ' use our _ERAC ConceptDocTypes table
    '
    ' So, for now, we load the doc type from the name found in the _CLAIMS, concept table
    
    '' KD COMEBACK.. When we transition to the new system we will want to change this
    '' hopefully to use load from id, (only if we put a new column in _CLAIMS..CONCEPT_References table / view)
    '' Otherwise, we'll just want to change the below to load from me.DocName instead of CnlyAttachType
    Set coEracDocType = New clsConceptReqDocType
    If Me.CnlyAttachType <> "ATTACH" Then
        If coEracDocType.LoadFromDocName(Me.CnlyAttachType) = False Then
            ' not a huge deal.. or is it?
            LastError = "Unknown Attachment type: '" & Me.CnlyAttachType & "'"
            LoadFromId = False
            GoTo Block_Exit
            '' KD COMEBACK: we got an 'ATTACH' RefSubType: select * from CMS_AUDITORS_CODE.dbo.v_CONCEPT_References WHERE ConceptID = 'CM_C0027'
        End If
    End If
    
    '' Is this a payer doc? If so, get / populate the rest of the info
    If Me.IsPayerDoc = True Then
        If GetPayerDetails = False Then
            Stop
        End If
    End If
    

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "DocRowId: " & CStr("" & lDocRowId)
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetPayerDetails() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String

    strProcName = ClassName & ".GetPayerDetails"


    sSql = "SELECT R.ConceptReferenceId, RP.PayerNameID, RP.ConceptReferencePayerId " & _
        "   FROM Concept_Submission_References R LEFT JOIN Concept_Submission_ReferencePayers RP ON R.ConceptReferenceId = RP.ConceptReferenceID" & _
        " Where r.ReferenceRowId = " & CStr(Me.ID)
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("ConceptDocTypes")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "No data returned for this doc type: " & Me.CnlyAttachType & " for concept: " & Me.ConceptID, CStr(Me.ID), , Me.ConceptID
            GoTo Block_Exit
        End If
    End With

    ' we're going to ASSuME that it's only 1 row - should be
    Me.PayerNameId = Nz(oRs("PayerNameID").Value, 0)
    Me.ConceptMgmtRefId = oRs("ConceptReferenceID").Value
    GetPayerDetails = True
    
Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GetPayerDetails = False
    GoTo Block_Exit
End Function

'''
'''
'''''' ##############################################################################
'''''' ##############################################################################
'''''' ##############################################################################
''''''
'''''' This function performs (or orchastrates) all validation  on a particular concept
'''''' to see if it's ready to submit to CMS
''''''
'''Public Function GetConceptHeaderDetails(sConceptId As String, Optional ByRef sReviewTypeCode As String = "", _
'''    Optional ByRef sDataTypeCode As String = "", Optional ByRef sConceptOwner As String = "", _
'''    Optional ByRef sClientIssueNum As String = "", Optional iCmsReviewTypeId As Integer = 0) As Boolean
'''On Error GoTo Block_Err
'''Dim strProcName As String
'''Dim oAdo As clsADO
'''Dim oRs As ADODB.Recordset
'''
'''
'''    strProcName = ClassName & ".GetConceptReviewType"
'''
'''    Set oAdo = New clsADO
'''    With oAdo
'''        .ConnectionString = GetConnectString("ConceptDocTypes")
'''        .SQLTextType = sqltext
'''        .SQLstring = "SELECT CH.Auditor ConceptOwner, CH.ClientIssueNum, CH.ReviewType, CH.DataType FROM " & _
'''            " CMS_AUDITORS_CLAIMS.dbo.Concept_Hdr CH WHERE Ch.ConceptID = '" & sConceptId & "'"
'''    End With
'''
'''
'''    Set oRs = oAdo.ExecuteRS
'''
'''
'''    '' Did we get anything?
'''    If oRs Is Nothing Then
'''        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
'''        GoTo Block_Exit
'''    End If
'''
'''    If oRs.RecordCount < 1 Then
'''        LogMessage strProcName, "WARNING", "Concept not found in Concept Header, please verify and try again"
'''        GoTo Block_Exit
'''    End If
'''
'''    If Not oRs.EOF Then
'''        sConceptOwner = Nz(oRs("ConceptOwner").Value, "")
'''        sReviewTypeCode = Nz(oRs("ReviewType").Value, "")
'''        sDataTypeCode = Nz(oRs("DataType").Value, "")
'''        sClientIssueNum = Nz(oRs("ClientIssueNum").Value, "")
'''        ' We don't expect more than 1, so no need to movenext
'''    End If
'''
'''    iCmsReviewTypeId = TranslateCnlyReviewTypeToCMS(sReviewTypeCode)
'''
'''    GetConceptHeaderDetails = True
'''
'''Block_Exit:
'''    Set oRs = Nothing
'''    Set oAdo = Nothing
'''    Exit Function
'''Block_Err:
'''    ReportError Err, strProcName
'''    Err.Clear
'''    GetConceptHeaderDetails = False
'''    GoTo Block_Exit
'''End Function





''##########################################################
''##########################################################
''##########################################################
'' Error handling
''##########################################################
''##########################################################
''##########################################################
Private Sub FireError(oErr As ErrObject, sErrSourceProcName As String, Optional sAdditionalDetails As String)

    Me.LastError = oErr.Description & sAdditionalDetails
    
    ReportError oErr, sErrSourceProcName, , sAdditionalDetails
    
    If sAdditionalDetails <> "" Then sAdditionalDetails = vbCrLf & sAdditionalDetails
    
    RaiseEvent ConceptDocError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

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
    
    cblnIsInitialized = False
    
End Sub


Private Sub Class_Terminate()
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing
    
    cblnIsInitialized = False
End Sub