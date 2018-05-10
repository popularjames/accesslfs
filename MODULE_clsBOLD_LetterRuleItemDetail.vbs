Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





'' Last Modified: 04/24/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''  The purpose of this is to
''
''
''
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/13/2013  - Created
''
'' AUTHOR
''  =====================================
'' Kevin Dearing
''
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################


Public Event LetterRuleItemError(ErrMsg As String, ErrNum As Long, ErrSource As String, bHandled As Boolean)

Private codctBoolVal As Scripting.Dictionary
Private codctOperatorTxt As Scripting.Dictionary

Private Const cs_TABLE_NAME As String = cs_TEMP_RULE_TABLE_NAME
Private coListItem As Object

Private clLocalId As Long
Private clRuleId As Long

Private ciBoolean As Integer
Private csItemName As String
Private csItemFieldName As String
Private csOperator As String
Private csItemValue As String

Private cvLookupId As Variant
Private csLkupDisplay As String

Private cbSaved As Boolean

Private csCurTableName As String
'Private csConnString As String

Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean
Private cbErrorOccurred As Boolean
Private csLastError As String

Private Const csIDFIELDNAME As String = "LocalId"
Private coSourceTable As clsTable


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
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


Public Property Get CurrentTableName() As String
    If csCurTableName = "" Then
        CurrentTableName = cs_TABLE_NAME
    Else
        CurrentTableName = csCurTableName
    End If
End Property
Public Property Let CurrentTableName(sTableNameToUse As String)
    csCurTableName = sTableNameToUse
End Property

Public Property Get ListItemObject() As Object
    Set ListItemObject = coListItem
End Property
Public Property Let ListItemObject(oListItem As Object)
    Set coListItem = oListItem
End Property


Public Property Get Id() As Long
    Id = clLocalId
End Property
Public Property Let Id(lId As Long)
    clLocalId = lId
End Property

Public Property Get RuleId() As Long
        ''RemoteId = clRuleId
    RuleId = GetTableValue("RuleId")
End Property
Public Property Let RuleId(lRuleId As Long)
        ''clRuleId = lRemoteId
    SetTableValue "RuleId", lRuleId
End Property


Public Property Get RuleItemId() As Long
        ''RemoteId = clRuleId
    RuleId = GetTableValue("RuleItemId")
End Property
Public Property Let RuleItemId(lRuleId As Long)
        ''clRuleId = lRemoteId
    SetTableValue "RuleItemId", lRuleId
End Property


Public Property Get ItemType() As String
    ItemType = Nz(GetTableValue("ItemType"), "")
End Property
Public Property Let ItemType(sItemType As String)
    SetTableValue "ItemType", sItemType
End Property

Public Property Get LookupId() As Variant
    cvLookupId = Nz(GetTableValue("LkupId"), "")
    LookupId = cvLookupId
End Property
Public Property Let LookupId(vLookupId As Variant)
    cvLookupId = vLookupId
    SetTableValue "LkupId", cvLookupId
End Property


Public Property Get LookupDisplay() As String
    csLkupDisplay = Nz(GetTableValue("LkupDisplay"), "")
    LookupDisplay = csLkupDisplay
End Property
Public Property Let LookupDisplay(sLookupDisplay As String)
    csLkupDisplay = sLookupDisplay
    SetTableValue "LkupDisplay", sLookupDisplay
End Property


Public Property Get BooleanVal() As Integer
    BooleanVal = CInt("0" & GetTableValue("Boolean"))
        ''    BooleanVal = ciBoolean
End Property
Public Property Let BooleanVal(iBooleanVal As Integer)
    ''ciBoolean = iBooleanVal
    SetTableValue "Boolean", iBooleanVal
End Property


Public Property Get BooleanValTxt() As String
    If codctBoolVal Is Nothing Then
        Call PopulateBoolValText
    End If
    BooleanValTxt = Nz(codctBoolVal.Item(Me.BooleanVal), "")
End Property

Public Property Get OperatorTxt() As String
    If codctOperatorTxt Is Nothing Then
        Call PopulateOperatorText
    End If
    OperatorTxt = Nz(codctOperatorTxt.Item(CInt(Nz("0" & Me.Operator, 0))), "")
End Property



Public Property Get ItemName() As String
    'ItemName = csItemName
    ItemName = Nz(GetTableValue("ItemName"), "")
End Property
Public Property Let ItemName(sItemName As String)
'    csItemName = sItemName
    SetTableValue "ItemName", sItemName
End Property


Public Property Get ItemFieldName() As String
    'ItemFieldName = csItemFieldName
    ItemFieldName = Nz(GetTableValue("RelatedFieldName"), "")
End Property
Public Property Let ItemFieldName(sItemFieldName As String)
    'csItemFieldName = sItemFieldName
    SetTableValue "RelatedFieldName", sItemFieldName
End Property



Public Property Get Operator() As String
    'Operator = csOperator
    Operator = Nz(GetTableValue("Operator"), "")
End Property
Public Property Let Operator(sOperator As String)
    'csOperator = sOperator
    SetTableValue "Operator", sOperator
End Property



Public Property Get ItemValue() As String
    'ItemValue = csItemValue
    ItemValue = Nz(GetTableValue("ItemValue"), "")
End Property
Public Property Let ItemValue(sItemValue As String)
'    csItemValue = sItemValue
    SetTableValue "ItemValue", sItemValue
End Property


Private Sub PopulateBoolValText()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".PopulateBoolValText"
    
    If Not codctBoolVal Is Nothing Then
        GoTo Block_Exit
    End If
    Set codctBoolVal = New Scripting.Dictionary
    
    codctBoolVal.Add 0, ""
    codctBoolVal.Add 1, "AND"
    codctBoolVal.Add 2, "OR"
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub PopulateOperatorText()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".PopulateOperatorText"
    
    If Not codctOperatorTxt Is Nothing Then
        GoTo Block_Exit
    End If
    Set codctOperatorTxt = New Scripting.Dictionary
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_LETTER_Automation_Rules_GetOperators"
        .Parameters.Refresh
        Set oRs = .ExecuteRS
    End With
    
    While Not oRs.EOF
        codctOperatorTxt.Add oRs("OperatorId").Value, oRs("OperatorName").Value
    
        oRs.MoveNext
    Wend

    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

'

'Public Property Get Saved() As Boolean
'Stop    ' not sure what I did this for - should this mean it's save to SQL Server or the local table?
'    Saved = cbSaved
'End Property
'Public Property Let Saved(bSaved As Boolean)
'    cbSaved = bSaved
'End Property


''##########################################################
''##########################################################
''##########################################################
'' Business logic type functions
''##########################################################
''##########################################################
''##########################################################

Public Function SecureLocalId(sUniqueInstance As String) As Long
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".SecureLocalId"
    
    SecureLocalId = coSourceTable.AddNewID
    Me.Id = SecureLocalId
    
    If Me.Id > 0 Then
        Me.LoadFromId (Me.Id)
        
    End If
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Private Function GetRecordset(sSql As String, Optional sTableName As String = cs_TABLE_NAME) As ADODB.RecordSet
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
''' SaveNow (duplicate of Save...)
Public Function DeleteNow() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String

    strProcName = ClassName & ".DeleteNow"

    If Dirty = False Or WasInitialized = False Then
        DeleteNow = False
        GoTo Block_Exit
    End If

    DeleteNow = coSourceTable.DeleteID(Me.Id)

    Dirty = Not DeleteNow

    DeleteNow = False
Block_Exit:

    Exit Function

Block_Err:
    DeleteNow = False
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
    
    Id = lDocRowId
    LoadFromId = coSourceTable.LoadFromId(lDocRowId)
    WasInitialized = LoadFromId

'''    ' Now, we have to get the document type object which we'll use for
'''    ' validations.
'''    ' Currently, because we are between "systems" we are using the
'''    ' CnlyAttachType field name - at least until we can move
'''    ' the rest of the concept attachment system over to
'''    ' use our _ERAC ConceptDocTypes table
'''    '
'''    ' So, for now, we load the doc type from the name found in the _CLAIMS, concept table
'''
'''    '' KD COMEBACK.. When we transition to the new system we will want to change this
'''    '' hopefully to use load from id, (only if we put a new column in _CLAIMS..CONCEPT_References table / view)
'''    '' Otherwise, we'll just want to change the below to load from me.DocName instead of CnlyAttachType
'''    Set coEracDocType = New clsConceptReqDocType
'''    If Me.CnlyAttachType <> "ATTACH" Then
'''        If coEracDocType.LoadFromDocName(Me.CnlyAttachType) = False Then
'''            ' not a huge deal.. or is it?
'''            LastError = "Unknown Attachment type: '" & Me.CnlyAttachType & "'"
'''            LoadFromID = False
'''            GoTo Block_Exit
'''            '' KD COMEBACK: we got an 'ATTACH' RefSubType: select * from CMS_AUDITORS_CODE.dbo.v_CONCEPT_References WHERE ConceptID = 'CM_C0027'
'''        End If
'''    End If
'''
'''    '' Is this a payer doc? If so, get / populate the rest of the info
'''    If Me.IsPayerDoc = True Then
'''        If GetPayerDetails = False Then
'''            Stop
'''        End If
'''    End If
    

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
Public Function LoadFromRS(oRs As Object) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromRS"
    coSourceTable.IdIsString = False
    Id = oRs("LocalId").Value
    LoadFromRS = coSourceTable.InitializeFromRS(oRs, True)
    WasInitialized = LoadFromRS

    ' Need to get the address information
'Stop

    ' Kev, if you are actually using this then you need to
    ' add some code here to loop through the fields and populate our static variables...
    ' or, change the properties to something like: if IsField("") then = coSourceTable.GetValue(""

'    If GetAddressDetails() = False Then
'        Stop
'    End If
'    Call LoadQRCodePath

Block_Exit:
    Exit Function

Block_Err:
    LoadFromRS = False
    FireError Err, strProcName, "Instance ID: " & Id
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

    Me.LastError = oErr.Description & sAdditionalDetails
    
    ReportError oErr, sErrSourceProcName, , sAdditionalDetails
    
    If sAdditionalDetails <> "" Then sAdditionalDetails = vbCrLf & sAdditionalDetails
    
    RaiseEvent LetterRuleItemError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName, False)

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
    coSourceTable.IsDAO = True
    coSourceTable.IdFieldName = csIDFIELDNAME
    coSourceTable.TableName = cs_TABLE_NAME
    
    cblnIsInitialized = False
End Sub

Private Sub Class_Terminate()
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing
    
    cblnIsInitialized = False
End Sub