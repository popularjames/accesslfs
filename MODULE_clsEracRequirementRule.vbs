Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 03/14/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This object represents a 'Rule' as defined in the _ERAC.dbo.CnlyConceptRequirements
'''  table.  Perhaps I'll create a Concept Validator object which will use this and
'''  shelter the developer from even touching it!
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 03/07/2012 - Created class
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

'Private Const csViewName As String = "v_ConceptRequirements"


Private ciConceptReqId As Integer
Private ccolRequiredDocs As Collection



Public Event EracRequirementError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean


Private Const csIDFIELDNAME As String = "ConceptReqId"
Private Const csTableName As String = "CnlyConceptRequirements"
Private coSourceTable As clsTable


Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean


Private ciDocTypeId As Integer
Private csDocName As String
Private csDescription As String
Private ciIsHdrLvlDoc As Integer
Private ciCmsHdrId As Integer
Private ciIsDtlLvlDoc As Integer
Private ciCmsDtlId As Integer
Private ciNumPerConcept As Integer
Private ciNumPerClaim As Integer
Private csNamingConvention As String
Private csSendAsFileType As String
Private csCreateFunctionName As String



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get ConceptReqId() As Integer
    ConceptReqId = ciConceptReqId
End Property
Public Property Let ConceptReqId(iConceptReqId As Integer)
    ciConceptReqId = iConceptReqId
'    SetTableValue "ConceptReqId", iConceptReqId, True
End Property
        '' Just an alias for ease of use!
    Public Property Get Id() As Integer
        Id = ConceptReqId
    End Property
    Public Property Let Id(intNewId As Integer)
        ConceptReqId = intNewId
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

Public Property Get ReviewTypeId() As Integer
    ReviewTypeId = CInt("0" & GetTableValue("ReviewTypeId"))
End Property
Public Property Let ReviewTypeId(iReviewTypeId As Integer)
    SetTableValue "ReviewTypeId", iReviewTypeId
End Property


Public Property Get CnlyReviewTypeCode() As String
    CnlyReviewTypeCode = EracCnlyReviewTypeFromEracId(Me.ReviewTypeId)
End Property


    '' Read only - set by using the ID!!
Public Property Get ReviewTypeName() As String
    ReviewTypeName = EracReviewTypeNameFromCnlyCode(Me.CnlyReviewTypeCode)
End Property


Public Property Get DataTypeCode() As String
    DataTypeCode = CStr("" & GetTableValue("DataTypeCode"))
End Property
Public Property Let DataTypeCode(sDataTypeCode As String)
    SetTableValue "DataTypeCode", sDataTypeCode
End Property


Public Property Get NumClaimsPerConcept() As Integer
    NumClaimsPerConcept = CInt("0" & GetTableValue("NumClaimsPerConcept"))
End Property
Public Property Let NumClaimsPerConcept(iNumClaimsPerConcept As Integer)
    SetTableValue "NumClaimsPerConcept", iNumClaimsPerConcept
End Property







''##########################################################
''##########################################################
''##########################################################
'' Business logic type functions
''##########################################################
''##########################################################
''##########################################################



''##########################################################
''##########################################################
''##########################################################
'' Collection of document types...
''##########################################################
''##########################################################
''##########################################################


Public Function RequiredDocs() As Collection
    Set RequiredDocs = ccolRequiredDocs
End Function


'Public Function AddField(intUnderlyingFieldID As Integer, intOrder As Integer) As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oField As cUnderlyingField
'Dim oRptField As cReportField
'
'    strProcName = ClassName & ".AddField"
'
'    Set oField = New cUnderlyingField
'    LogMessage "Loading underlying field from id: " & CStr(intUnderlyingFieldID), strProcName, "DEBUG TRAIL"
'    If oField.LoadFromID(intUnderlyingFieldID) = False Then
'        LogMessage "Field ID not defined in AdHoc Fields table", strProcName, "DEBUG TRAIL", CStr(intUnderlyingFieldID)
'        GoTo Block_Exit
'    End If
'
'    Set oRptField = New cReportField
'    With oRptField
''        .LoadFromID intFieldID, True
'        LogMessage "Adding report field", strProcName, "DEBUG TRAIL"
'        .ID = .AddNew(Me.ReportID, intUnderlyingFieldID)
''        .ReportID = Me.ReportID
''        .FieldID = intFieldID
'        .Order = intOrder
'        ccolRptFields.Add oRptField
'        ccolRptFieldsToAdd.Add oRptField
'    End With
'
'    Dirty = True
'    AddField = True
'Block_Exit:
'    Exit Function
'Block_Err:
'    AddField = False
'    ReportError Err, strProcName
'    Resume Block_Exit
'End Function


Public Function Item(intFieldId As Integer) As clsConceptReqDocType
On Error GoTo Block_Err
Dim strProcName As String
Dim intIndex As Integer


    strProcName = ClassName & ".Item"
    intIndex = GetColIndexFromFieldID(intFieldId)
    
    If intIndex > 0 Then
        Set Item = ccolRequiredDocs.Item(intIndex)
    Else
        Set Item = Nothing
    End If

    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function


'Public Function RemoveDoc(intFieldId As Integer, intOrder As Integer) As Boolean
''On Error GoTo Block_Err
''Dim strProcName As String
''Dim oReqdDoc As clsConceptReqDocType
'''Dim oField As cUnderlyingField
''Dim intFieldIndexToRemove As Integer
''
''
''    strProcName = ClassName & ".RemoveField"
''    LogMessage "Removing ReportField: " & CStr(intFieldId), strProcName, "DEBUG TRAIL"
''    Set oField = New cUnderlyingField
'''    With oField
''        If .LoadFromID(intFieldId) = False Then
''            LogMessage "Field ID not defined in AdHoc Fields table", strProcName, "DEBUG TRAIL", CStr(intFieldId)
''            GoTo Block_Exit
''        End If
''        Set oReqdDoc = New cReportField
''
'''        If oReqdDoc.LoadFromRptIDandUnderlyer(Me.ReportID, oField.FieldID) = False Then
'''            LogMessage "Problem loading report field", strProcName, "REPORT FIELD REMOVAL ERROR", CStr(Me.ReportID) & " " & CStr(oField.FieldID)
'''            GoTo Block_Exit
'''        End If
''
''        If RemoveItemFromCollection(ccolRequiredDocs, oReqdDoc) = False Then
''            LogMessage "Problem removing item from collection..", strProcName, "DEBUG TRAIL", oReqdDoc.FieldName
''        End If
'''        ccolRptFields.Remove oField
'''        Else
'''            LogMessage "Could not find index for field: " & oField.FieldName, strProcName
'''        End If
''        ccolRptFieldsToRemove.Add oReqdDoc
'''    End With
''
''    Dirty = True
''    RemoveField = True
''
''Block_Exit:
''    Exit Function
''Block_Err:
''    RemoveField = False
''    ReportError Err, strProcName
''    Resume Block_Exit
'End Function

Private Function RemoveItemFromCollection(oCollection As Collection, varItemToRemove As Variant) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim ovarColItem As Variant
Dim iIndex As Integer


    strProcName = ClassName & ".RemoveItemFromCollection"
    
    For iIndex = 1 To oCollection.Count
        Set ovarColItem = oCollection.Item(iIndex)
        If ovarColItem.Id = varItemToRemove.Id Then
            Set ovarColItem = Nothing
            oCollection.Remove iIndex
            RemoveItemFromCollection = True
            GoTo Block_Exit
        End If
    Next
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function


Public Function GetColIndexFromDoc(oField As clsConceptReqDocType) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oCurDoc As clsConceptReqDocType
Dim intIndex As Integer

    strProcName = ClassName & ".GetColIndexFromDoc"
    GetColIndexFromDoc = -1
    
    For intIndex = 1 To ccolRequiredDocs.Count
        Set oCurDoc = New clsConceptReqDocType
        oCurDoc.LoadFromId (ccolRequiredDocs(intIndex).Id)

'        If oField.FieldID = oCurDoc.DocTypeId Then
'            GetColIndexFromDoc = intIndex
'            GoTo Block_Exit
'        End If
    Next
    
Block_Exit:
    Set oCurDoc = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function


Private Function GetColIndexFromFieldID(intFieldId As Integer) As Integer
On Error GoTo Block_Err
Dim strProcName As String
Dim oCurDoc As clsConceptReqDocType
Dim intIndex As Integer

    strProcName = ClassName & ".GetColIndexFromFieldID"
    GetColIndexFromFieldID = -1
    
    For intIndex = 1 To ccolRequiredDocs.Count
        Set oCurDoc = New clsConceptReqDocType
        oCurDoc.LoadFromId (ccolRequiredDocs(intIndex).Id)
        If oCurDoc.DocTypeId = intFieldId Then
            GetColIndexFromFieldID = intIndex
            GoTo Block_Exit
        End If
    Next
    
Block_Exit:
    Set oCurDoc = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Resume Block_Exit
End Function



''##########################################################
''##########################################################
''##########################################################
'' Audit / Setup data / interacting with the cTable object
''##########################################################
''##########################################################
''##########################################################



Public Function GetTableValue(strFieldName As String) As Variant
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetTableValue"

'    If cintUserID < 1 Then
'        ' not initialized
''        GetTableValue = Nothing
'        GoTo Block_Exit
'    End If

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


' SaveNow (duplicate of Save...)
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




Public Function LoadFromRS(oRs As ADODB.RecordSet) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    
    If isField(oRs, coSourceTable.IdFieldName) = False Then
        GoTo Block_Exit
    End If
    
    Id = oRs(coSourceTable.IdFieldName).Value
    LoadFromRS = coSourceTable.LoadFromId(Id)
    WasInitialized = LoadFromRS

    '' Now, need to get the Required document types and populate our collection
    Call GetReqdDocsForThisRequirement

Block_Exit:
    Exit Function

Block_Err:
    FireError Err, strProcName, "SourceID: " & CStr(Id)
    LoadFromRS = False
    GoTo Block_Exit
End Function



Public Function LoadFromId(lngRSourceId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    
    Id = lngRSourceId
    LoadFromId = coSourceTable.LoadFromId(lngRSourceId)
    WasInitialized = LoadFromId

    '' Now, need to get the Required document types and populate our collection
    
    Call GetReqdDocsForThisRequirement

Block_Exit:
    Exit Function

Block_Err:
    FireError Err, strProcName, "SourceID: " & CStr(lngRSourceId)
    LoadFromId = False
    GoTo Block_Exit
End Function


Public Function LoadFromConceptID(sConceptId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim iCmsReviewTypeId As Integer
Dim sDataTypeCode As String
Dim lngRsSourceId As Long
Dim sSql As String

    strProcName = ClassName & ".LoadFromConceptID"
    
    If GetConceptHeaderDetails(sConceptId, , sDataTypeCode, , , iCmsReviewTypeId) = False Then
        ' KD COMEBACK: What's this?
    End If
    
    '' Now we need to get the rule for the reviewtype and data type
    If sDataTypeCode = "DME" Then
        sSql = "SELECT * FROM " & csTableName & " WHERE ReviewTypeId = " & CStr(iCmsReviewTypeId) & " AND DataTypeCode = 'DME' "
    Else
        sSql = "SELECT * FROM " & csTableName & " WHERE ReviewTypeId = " & CStr(iCmsReviewTypeId) & " AND ISNULL(DataTypeCode, '') = '' "
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(csTableName)
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
    End With
    
    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    
    lngRsSourceId = oRs(csIDFIELDNAME).Value
    
    LoadFromConceptID = LoadFromId(lngRsSourceId)

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Function

Block_Err:
    FireError Err, strProcName, "SourceID: " & sConceptId
    LoadFromConceptID = False
    GoTo Block_Exit
End Function


Private Sub GetReqdDocsForThisRequirement()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim oReqdDocType As clsConceptReqDocType
    
    strProcName = ClassName & ".GetReqdDocsForThisRequirement"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(csTableName)
        .SQLTextType = sqltext
        .sqlString = "SELECT DocTypeId FROM v_ConceptRequirements WHERE ConceptReqId = " & CStr(Id)
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With
    
    Set ccolRequiredDocs = New Collection
    
    While Not oRs.EOF
        If Nz(oRs("DocTypeID"), 0) > 0 Then
            Set oReqdDocType = New clsConceptReqDocType
            If oReqdDocType.LoadFromId(oRs("DocTypeId").Value) = True Then
                ccolRequiredDocs.Add oReqdDocType
            Else
                LogMessage strProcName, "WARNING", "Could not load doctype = " & CStr(oRs("DocTypeID").Value)
            End If
        End If
        oRs.MoveNext
    Wend
    

Block_Exit:
    Set oRs = Nothing
    Set oAdo = Nothing
    Exit Sub

Block_Err:
    FireError Err, strProcName, "SourceID: " & CStr(Me.Id)
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
    
    RaiseEvent EracRequirementError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

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