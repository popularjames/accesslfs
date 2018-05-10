Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





'' Last Modified: 5/12/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''  The purpose of this is to be used as a control in another form
''  The other form will be fairly generic so hard to describe but here goes
''  The main form is going to allow the user to select something from a list view and "drop it" in here
''  in here, the user can add criteria to it.. So, say 'Provider' is dropped in. (Attached to that we
''  have the 'CnlyProvId' fieldname).  Criteria may be something like "IN (123321, 1234532)"
''  Or maybe state (ProvStateCd) is dropped in, criteria may be something like " NOT 'PA' "
''      AND NOT LIKE 'M%'
''
''  The way we are going to achie3ve this is with a local temp table...
''  We'll load it up with whatever is in SQL Server. Bind the form and lock the controls. Clicking Edit will unlock that row?
''  Maybe we should create one more form for the individual values where they can edit a single line.. so they can save / cancel?
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
''  - 04/12/2015  - Created
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


Private WithEvents coLetterRuleItems As clsBOLD_LetterRuleItemDetails
Attribute coLetterRuleItems.VB_VarHelpID = -1
Public Event LetterRuleError(ErrMsg As String, ErrNum As Long, ErrSource As String)


Private clRuleId As Long
Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean
Private cbErrorOccurred As Boolean
Private csLastError As String

Private Const csIDFIELDNAME As String = "RuleId"
Private coSourceTable As clsTable

Private Const cs_TABLE_NAME As String = "BOLD_Letter_Automation_Rules"
Private csCurTableName As String

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

Public Property Get RuleName() As String
    RuleName = CStr("" & GetTableValue("RuleName"))
End Property
Public Property Let RuleName(sRuleName As String)
    SetTableValue "RuleName", sRuleName
End Property

Public Property Get Qty() As Long
    Qty = CLng("0" & GetTableValue("Qty"))
End Property
Public Property Let Qty(lQuantity As Long)
    SetTableValue "Qty", lQuantity
End Property


Public Property Get ObjectIdToLimit() As Long
    ObjectIdToLimit = CLng("0" & GetTableValue("ObjectToLimitId"))
End Property
Public Property Let ObjectIdToLimit(lObjectIdToLimit As Long)
    SetTableValue "ObjectToLimitId", lObjectIdToLimit
End Property


Public Property Get FinalFormatId() As Long
    FinalFormatId = CLng("0" & GetTableValue("FinalFormatId"))
End Property
Public Property Let FinalFormatId(lFinalFormatId As Long)
    SetTableValue "FinalFormatId", lFinalFormatId
End Property


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



Public Property Get ID() As Long
    ID = clRuleId
End Property
Public Property Let ID(lId As Long)
    clRuleId = lId
End Property


Public Property Get Details() As clsBOLD_LetterRuleItemDetails
    Set Details = GetRuleItemDetails
End Property
Public Property Let Details(oLtrRuleDetails As clsBOLD_LetterRuleItemDetails)
    Set coLetterRuleItems = oLtrRuleDetails
End Property


Public Function GetRuleItemDetails() As clsBOLD_LetterRuleItemDetails
    Set GetRuleItemDetails = coLetterRuleItems
End Function


Private Sub coLetterRuleItems_ItemChanged()
    Me.Dirty = True
End Sub


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
'''
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
        Stop
            ' secure an id?
'        Call SecureId
        Stop
        
        SaveNow = False
        GoTo Block_Exit
    End If

    SaveNow = coSourceTable.SaveNow()

    '' Now, we need to save the details
    '' but we have to do both: GroupBy and Per - but that's going to be dealt with
    '' inside the object...
    If coLetterRuleItems.SaveNow() = False Then
        Stop
    End If

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

    DeleteNow = coSourceTable.DeleteID(Me.ID)

    Dirty = Not DeleteNow

    DeleteNow = False
Block_Exit:

    Exit Function

Block_Err:
    DeleteNow = False
    FireError Err, strProcName, "User ID: " & Identity.UserName()
    GoTo Block_Exit
End Function
'
'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'Public Function Save() As Boolean
'On Error GoTo Block_Err
'Dim strProcName As String
'
'    strProcName = ClassName & ".Save"
'    ' start optimistically
'    Save = True
'
'    If Me.Dirty = False Then GoTo Block_Exit
'    If Me.WasInitialized = False Then GoTo Block_Exit
'
'    ' So now, we need to save this rule to the database...
'    Call SaveNow
'
'
'Block_Exit:
'    Exit Function
'
'Block_Err:
'    Save = False
'    FireError Err, strProcName, "User ID: " & Identity.UserName()
'    GoTo Block_Exit
'End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromId(lRuleId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
'       Debug.Assert iDocRowId <> 7073
    
    ID = lRuleId
    LoadFromId = coSourceTable.LoadFromId(lRuleId)
    WasInitialized = LoadFromId

    '' Now, we need to get any details - and put them into the local table here for editing..
    
    If coLetterRuleItems Is Nothing Then Set coLetterRuleItems = New clsBOLD_LetterRuleItemDetails
    
    ' populate our local table
'    Dim oAdoRs As ADODB.RecordSet
'    Dim oAdo As clsADO
'
'
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = CodeConnString
'        .SQLTextType = StoredProc
'        .sqlString = "usp_BOLD_LETTER_Automation_RuleItems"
'        .Parameters.Refresh
'        .Parameters("@pRuleId") = lRuleId
'        Set oAdoRs = .ExecuteRS
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            Stop
'        End If
'        If CopyDataToLocalTmpTable(oAdoRs, False, cs_TEMP_RULE_TABLE_NAME) = "" Then
'Stop
'        End If
'    End With

    If coLetterRuleItems.LoadFromRuleId(lRuleId) = False Then
        Stop
        
    End If
    


'    Stop
    ' Now I guess we need to load the forms..
    ' We should make the form a child of the object don't you think?
    

'    Set oDb = CurrentDb()
'
'Dim sFieldList As String
'Dim sValueList As String
'Dim oCurrentItem As clsBOLD_LetterRuleItemDetail
'Dim oFld As DAO.Field
'
'
'    Set oDb = CurrentDb()
'        oDb.Execute "DELETE FROM " & cs_TEMP_RULE_TABLE_NAME
'
'        Set oRs = oDb.OpenRecordSet("SELECT * FROM " & cs_TEMP_RULE_TABLE_NAME & " WHERE 1 = 2", dbOpenDynaset, dbSeeChanges)
'
'
'        For Each oCurrentItem In coLetterRuleItems.Items
'            For Each oFld In oRs.Fields
'                sFieldList = sFieldList & oFld.Name & ", "
'                sValueList = sValueList & QuoteIfNeeded(oCurrentItem.GetField(oFld.Name)) & ", "
'            Next
'            sFieldList = left(sFieldList, Len(sFieldList) - 2)
'            sValueList = left(sValueList, Len(sValueList) - 2)
'
'            sSql = "INSERT INTO " & cs_TEMP_RULE_TABLE_NAME & " (" & sFieldList & ") "
'            sSql = "VALUES (" & sValueList & ")"
'
'Stop
'            oDb.Execute sSql
'        Next
'

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "DocRowId: " & CStr("" & lRuleId)
    GoTo Block_Exit
End Function


Public Function CreateNew(sRuleName As String, lQty As Long, lObjectToLimitID As Long, lFinalFormatId As Long) As Long
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oItmDtls As clsBOLD_LetterRuleItemDetails


    strProcName = ClassName & ".SecureId"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_BOLD_LETTER_Automation_CreateNewRule"
        .Parameters.Refresh
        .Parameters("@pRuleName") = sRuleName
        .Parameters("@pQty") = lQty
        .Parameters("@pObjectToLimitId") = lObjectToLimitID
        .Parameters("@pFinalFormatId") = lFinalFormatId
        
        .Execute
        
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was a problem saving this rule", .Parameters("@pErrStr").Value, True
            Stop
            GoTo Block_Exit
        End If
        CreateNew = .Parameters("@pRuleId").Value
        
        Me.ID = CreateNew
        
    End With
    
    WasInitialized = True

    
Block_Exit:
    Set oAdo = Nothing
    Exit Function

Block_Err:
    CreateNew = False
    FireError Err, strProcName
    GoTo Block_Exit
End Function

''##########################################################
''##########################################################
''##########################################################
'' Class state properties
''##########################################################
''##########################################################
''##########################################################
    
   

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
    
    RaiseEvent LetterRuleError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub









Private Sub Class_Initialize()
    Set coLetterRuleItems = New clsBOLD_LetterRuleItemDetails
    
    Set coSourceTable = New clsTable
    coSourceTable.IsDAO = False
    coSourceTable.IdFieldName = csIDFIELDNAME
    coSourceTable.TableName = cs_TABLE_NAME
    coSourceTable.TableNameForConnection = "v_Data_Database"
    
    cblnIsInitialized = False
End Sub

Private Sub Class_Terminate()
    Set coLetterRuleItems = Nothing
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing
    
    cblnIsInitialized = False
End Sub