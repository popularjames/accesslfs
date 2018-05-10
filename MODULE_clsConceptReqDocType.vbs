Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


''' Last Modified: 07/13/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  This doesn't exactly translate to a table.. Instead, it is basically inherited (as inherited as VBA can do)
'''  from the clsConceptDoc class.. It's extended to contain details about the document
'''  that we can use to map back to CMS's id as well as map back to the legacy Attachment types
'''  It's a bit cornfusing at the moment because we're in a transition period..
'''
'''  Bottom line, this is kind of the mapping table
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 06/22/2012 - Added PayerNameId
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



Public Event ConceptReqDocTypeError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean


Private Const csIDFIELDNAME As String = "DocTypeId"
Private Const csTableName As String = "ConceptDocTypes"
Private coSourceTable As clsTable


Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean
Private coPayer As clsConceptPayerDtl

Private ciDocTypeId As Integer



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get DocTypeId() As Integer
    DocTypeId = ciDocTypeId
End Property
Public Property Let DocTypeId(iDocTypeId As Integer)
    ciDocTypeId = iDocTypeId
'    SetTableValue "DocTypeId", iDocTypeId, True
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As Integer
        ID = DocTypeId
    End Property
    Public Property Let ID(intNewId As Integer)
        DocTypeId = intNewId
    End Property


Public Property Get IsPayerDoc() As Boolean
Dim sRet As String
    sRet = Nz(GetTableValue("IsPayerDoc"), "0")
    If sRet = "" Then sRet = "0"
    
    IsPayerDoc = CBool(sRet)
End Property
Public Property Let IsPayerDoc(bIsPayerDoc As Boolean)
    SetTableValue "IsPayerDoc", CStr(IIf(bIsPayerDoc, "1", "0"))
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

Public Property Get DocName() As String
    DocName = CStr("" & GetTableValue("DocName"))
End Property
Public Property Let DocName(sDocName As String)
    SetTableValue "DocName", sDocName
End Property


Public Property Get Description() As String
    Description = CStr("" & GetTableValue("Description"))
End Property
Public Property Let Description(sDescription As String)
    SetTableValue "Description", sDescription
End Property

'
'Public Property Get IsHdrLvlDoc() As Boolean
'Dim iVal As Integer
'    iVal = CInt("0" & GetTableValue("IsHdrLvlDoc"))
'    If iVal = 0 Then
'        IsHdrLvlDoc = False
'    Else
'        IsHdrLvlDoc = True
'    End If
'End Property
'Public Property Let IsHdrLvlDoc(bIsHdrLvlDoc As Boolean)
'    SetTableValue "IsHdrLvlDoc", IIf(bIsHdrLvlDoc, 1, 0)
'End Property

'
'Public Property Get CmsHdrId() As Integer
'    CmsHdrId = CInt(GetTableValue("CmsHdrId"))
'End Property
'Public Property Let CmsHdrId(iCmsHdrId As Integer)
'    SetTableValue "CmsHdrId", iCmsHdrId
'End Property


'
'Public Property Get IsDtlLvlDoc() As Boolean
'Dim iVal As Integer
'    iVal = CInt("0" & GetTableValue("IsDtlLvlDoc"))
'    If iVal = 0 Then
'        IsDtlLvlDoc = False
'    Else
'        IsDtlLvlDoc = True
'    End If
'End Property
'Public Property Let IsDtlLvlDoc(bIsDtlLvlDoc As Boolean)
'    SetTableValue "IsDtlLvlDoc", IIf(bIsDtlLvlDoc, 1, 0)
'End Property

'
'
'Public Property Get CmsDtlId() As Integer
'    CmsDtlId = CInt(GetTableValue("CmsDtlId"))
'End Property
'Public Property Let CmsDtlId(iCmsDtlId As Integer)
'    SetTableValue "CmsDtlId", iCmsDtlId
'End Property



Public Property Get NumPerConcept() As Integer
    NumPerConcept = CInt("0" & GetTableValue("NumPerConcept"))
End Property
Public Property Let NumPerConcept(iNumPerConcept As Integer)
    SetTableValue "NumPerConcept", iNumPerConcept
End Property


Public Property Get NumPerPayer() As Integer
    NumPerPayer = CInt("0" & GetTableValue("NumPerPayer"))
End Property
Public Property Let NumPerPayer(iNumPerPayer As Integer)
    SetTableValue "NumPerPayer", iNumPerPayer
End Property



Public Property Get NamingConvention() As String
    NamingConvention = CStr("" & GetTableValue("NamingConvention"))
End Property
Public Property Let NamingConvention(sNamingConvention As String)
    SetTableValue "NamingConvention", sNamingConvention
End Property


Public Property Get SendAsFileType() As String
    SendAsFileType = CStr("" & GetTableValue("SendAsFileType"))
End Property
Public Property Let SendAsFileType(sSendAsFileType As String)
    SetTableValue "SendAsFileType", sSendAsFileType
End Property


Public Property Get CreateFunctionName() As String
    CreateFunctionName = CStr("" & GetTableValue("CreateFunctionName"))
End Property
Public Property Let CreateFunctionName(sCreateFunctionName As String)
    SetTableValue "CreateFunctionName", sCreateFunctionName
End Property


Public Property Get CheckExistanceSQL() As String
    CheckExistanceSQL = CStr("" & GetTableValue("CheckExistanceSQL"))
End Property
Public Property Let CheckExistanceSQL(sCheckExistanceSQL As String)
    SetTableValue "CheckExistanceSQL", sCheckExistanceSQL
End Property


Public Property Get CnlyAttachType() As String
    CnlyAttachType = CStr("" & GetTableValue("CnlyAttachType"))
End Property
Public Property Let CnlyAttachType(sCnlyAttachType As String)
    SetTableValue "CnlyAttachType", sCnlyAttachType
End Property


Public Property Get Display() As Boolean
Dim sVal As String
    sVal = CStr("" & GetTableValue("Display"))
    If IsNumeric(sVal) Then
        Display = CBool(sVal)
    Else
        Display = True
    End If
End Property
Public Property Let Display(bDisplay As Boolean)
    SetTableValue "Display", IIf(bDisplay, 1, 0)
End Property



Public Property Get Active() As Boolean
Dim sVal As String
    sVal = CStr("" & GetTableValue("Active"))
    If IsNumeric(sVal) Then
        Active = CBool(sVal)
    Else
        Active = True
    End If
End Property
Public Property Let Active(bActive As Boolean)
    SetTableValue "Active", IIf(bActive, 1, 0)
End Property

''##########################################################
''##########################################################
''##########################################################
'' Business logic type functions
''##########################################################
''##########################################################
''##########################################################

Public Function ParseFileName(sConceptId As String, sClientIssueNumber As String, Optional sIcnId As String = "", _
    Optional ByVal SFileName As String, Optional ByVal sPayerName As String, Optional oPayer As clsConceptPayerDtl) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sReturn As String
Dim iDotSpot As Integer

    strProcName = ClassName & ".ParseFileName"

    If Not oPayer Is Nothing Then
        If sPayerName = "" Then sPayerName = oPayer.PayerName
        sClientIssueNumber = oPayer.ClientIssueId
    End If

    sReturn = Me.NamingConvention
    If InStr(1, sReturn, "[*ClientIssueNumber*]", vbTextCompare) > 0 Then
        sReturn = Replace(sReturn, "[*ClientIssueNumber*]", sClientIssueNumber, 1, -1, vbTextCompare)
    End If
    
'    sReturn = Replace(sReturn, "[*ClientIssueNumber*]", sClientIssueNumber, 1, -1, vbTextCompare)
    If InStr(1, sReturn, "[*ConceptId*]", vbTextCompare) > 0 Then
        sReturn = Replace(sReturn, "[*ConceptId*]", sConceptId, 1, -1, vbTextCompare)
    End If
    
    If InStr(1, sReturn, "[*ICN*]", vbTextCompare) > 0 Then
        sReturn = Replace(sReturn, "[*ICN*]", sIcnId, 1, -1, vbTextCompare)
    End If
    
    
    '' Ok, so payer name is weird.. If we have a payername in the string, then we need to put that in
    '' even if we don't have some sort of payername code...
    If sPayerName <> "" Then
        sReturn = sReturn & "_" & SafeFileName(sPayerName)
    End If
    
        '' We'll just leave this in there for now..
    If InStr(1, sReturn, "[*PayerName*]", vbTextCompare) > 0 Then
        sReturn = Replace(sReturn, "[*PAYERNAME*]", sPayerName, 1, -1, vbTextCompare)
    End If
    
    '' Make sure it's not a path:
    If InStr(1, SFileName, "\", vbTextCompare) > 0 Then
        SFileName = Right(SFileName, InStr(1, StrReverse(SFileName), "\") - 1)
    End If
    
    '' If we have sFileName AND we have 'Filename' in the convention,
    '' then we need to incorporate the actual filename into what we return
    '' If we don't have sFileName just return nothing
    If SFileName = "" Then
        '' If it's just 'Filename' then return nothing
        sReturn = Replace(sReturn, "[*FileName*]", "", 1, -1, vbTextCompare)
        iDotSpot = InStr(1, sReturn, ".", vbTextCompare)
        If iDotSpot > 0 Then sReturn = left(sReturn, iDotSpot - 1)
    
    Else
        '' Make sure we don't have the file extension
        iDotSpot = InStr(1, SFileName, ".", vbTextCompare)
        If iDotSpot > 0 Then SFileName = left(SFileName, iDotSpot - 1)
        
        '' SHould we take out any hard coded concept or ICN (Client Issue Number?
        SFileName = Replace(SFileName, sConceptId, "", 1, -1, vbTextCompare)
        If sIcnId <> "" Then
            SFileName = Replace(SFileName, sIcnId, "", 1, -1, vbTextCompare)
        End If
        
        '' Now, insert that as the filename:
        sReturn = Replace(sReturn, "[*FileName*]", SFileName, 1, 1, vbTextCompare)
        
    End If
    

    If sReturn = "" Then sReturn = SFileName
    
Block_Exit:
    ParseFileName = sReturn
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & sConceptId
    GoTo Block_Exit
End Function


Private Function SafeFileName(sUnSafe As String) As String
Dim oRegEx As RegExp

    Set oRegEx = New RegExp
    oRegEx.IgnoreCase = True
    oRegEx.Global = True
    oRegEx.Pattern = "([\s\t\[\]\{\}\(\)\-\+\=\*]+)"
    
    SafeFileName = oRegEx.Replace(sUnSafe, "_")
    SafeFileName = Replace(SafeFileName, "__", "_")
    
End Function

Public Function DocExistsForConcept(sConceptId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String

    strProcName = ClassName & ".DocExistsForConcept"
    DocExistsForConcept = True
    
    If Me.CheckExistanceSQL = "" Then GoTo Block_Exit
        
    sSql = Replace(Me.CheckExistanceSQL, "?", sConceptId, , , vbTextCompare)
    
    Set oRs = GetRecordset(sSql)
    
    If oRs Is Nothing Then
        DocExistsForConcept = False
        GoTo Block_Exit
    End If
    If oRs.recordCount < 1 Then
        DocExistsForConcept = False
        GoTo Block_Exit
    End If
    
    '' Do we have as many as we were expecting - well, that's for another function isn't it?
    '' and of course that doesn't belong in this object (heck, this function is debatable)


Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName, "User ID: " & Identity.UserName() & " " & sConceptId
    DocExistsForConcept = False
    GoTo Block_Exit
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
    If coSourceTable Is Nothing Then GoTo Block_Exit
    
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


'' I don't like doing this because ID is the PK, and technically the Doc Name may not be unique, but
'' I'll put a unique constraint on the table..

Public Function LoadFromDocName(sDocTypeName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Static dctNamesToIds As Scripting.Dictionary
Dim sMsg As String
Dim lngRsSourceId As Long


    strProcName = ClassName & ".LoadFromID"
    
    ' Doc names aren't going to change so we'll cache the name to id
    ' And, while we are in transition, we are going to use the CnlyAttachType names
    ' instead of our DocName
    If dctNamesToIds Is Nothing Then
        Set dctNamesToIds = New Scripting.Dictionary
        '       Set oRs = GetRecordset("SELECT DocTypeId, DocName FROM ConceptDocTypes")      ' WHERE DocName = '" & sDocTypeName & "'")
        Set oRs = GetRecordset("SELECT DocTypeId, CnlyAttachType as DocName FROM ConceptDocTypes")      ' WHERE DocName = '" & sDocTypeName & "'")
        If oRs Is Nothing Then
            sMsg = "Could not find the id for that document name! (" & sDocTypeName & ")"
            LogMessage strProcName, "ERROR", sMsg
            GoTo Block_Exit
        End If
        
        While Not oRs.EOF
            If dctNamesToIds.Exists(CStr("" & oRs("DocName").Value)) = False Then
                dctNamesToIds.Add CStr("" & oRs("DocName").Value), oRs("DocTypeId").Value
            End If
            oRs.MoveNext
        Wend
    End If
    
    If dctNamesToIds.Exists(sDocTypeName) = False Then
        LogMessage strProcName, "ERROR", "Did not find the id for that filename", sDocTypeName
        GoTo Block_Exit
    End If
    
    lngRsSourceId = dctNamesToIds.Item(sDocTypeName)

    LoadFromDocName = coSourceTable.LoadFromId(lngRsSourceId)
    ID = lngRsSourceId
    WasInitialized = LoadFromDocName
    
Block_Exit:
    Set oRs = Nothing
    Exit Function

Block_Err:
    LoadFromDocName = False
    FireError Err, strProcName, sMsg
    GoTo Block_Exit
End Function


Public Function LoadFromId(lngRsSourceId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    
    ID = lngRsSourceId
    LoadFromId = coSourceTable.LoadFromId(lngRsSourceId)
    WasInitialized = LoadFromId

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "SourceID: " & CStr(lngRsSourceId)
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
    
    RaiseEvent ConceptReqDocTypeError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

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