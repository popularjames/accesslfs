Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 6/24/2014
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a Manual Override Filter
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 06/24/2014 - Created class
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


Private csFilterName As String
Private clManFilterRID As Long    ' RID
Private clFilterId As Long


Private clUnderlyingOptionId As Long    ' Tie to the LETTER_Automation_Xref_ManualFilters table
Private csPossibleValuesSproc As String
Private csSampleSource As String
Private csSampleFieldName As String


Private clManFilterID As Long
Private clAccountId As Long

Private csOptionID As String
Private csInclusion As String
Private csFilterBy As String
Private csOperator As String
Private csRequiredValueIndex As String
Private csRequiredValue As String


Private cbSave As Boolean   ' should this one be saved or is is only for viewing sample data?


Private Const csIDFIELDNAME As String = "RID"
Private Const csTableName As String = "LETTER_Automation_ManualFilter_Values"

    '' The table to use for the connection string to the _ERAC database
Private Const csSP_TABLENAME As String = "v_Code_Database"

Private coSourceTable As clsTable

Private cblnDirtyData As Boolean
Private cblnIsInitialized As Boolean
'Private cblnIsNew As Boolean


''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

        '' Just an alias for ease of use!
    Public Property Get ID() As Long
        ID = FilterOptionID
    End Property
    Public Property Let ID(lNewId As Long)
        FilterOptionID = lNewId
    End Property


Public Property Get AccountID() As Long
    clAccountId = GetTableValue("AccountId")
    AccountID = clAccountId
End Property
Public Property Let AccountID(lAccountId As Long)
    clAccountId = lAccountId
    SetTableValue "AccountId", lAccountId
End Property

Public Property Get FilterOptionID() As Long
    clManFilterRID = GetTableValue("RID")
    FilterOptionID = clManFilterRID
End Property
Public Property Let FilterOptionID(lFilterOptionID As Long)
    SetTableValue "RID", lFilterOptionID
    clManFilterRID = lFilterOptionID
End Property

Public Property Get Save() As Boolean
    Save = cbSave
End Property
Public Property Let Save(bSave As Boolean)
    cbSave = bSave
End Property


'Public Property Get IsNew() As Boolean
'    IsNew = cblnIsNew
'End Property
'Public Property Let IsNew(bIsNew As Boolean)
'    cblnIsNew = bIsNew
'    If cblnIsNew = True Then
'        Dirty = True
'        WasInitialized = True
'    End If
'End Property


Public Property Get PossibleValuesSproc() As String
'    csPossibleValuesSproc = GetTableValue("PossibleValuesSproc")
    PossibleValuesSproc = csPossibleValuesSproc
End Property
Public Property Let PossibleValuesSproc(sPossibleValuesSproc As String)
    csPossibleValuesSproc = sPossibleValuesSproc
End Property


Public Property Get SampleSource() As String
    SampleSource = csSampleSource
End Property
Public Property Let SampleSource(sSampleSource As String)
    csSampleSource = sSampleSource
End Property

Public Property Get SampleFieldName() As String
    SampleFieldName = csSampleFieldName
End Property
Public Property Let SampleFieldName(sSampleFieldName As String)
    csSampleFieldName = sSampleFieldName
End Property

'' Need this to be loaded at LoadID
Public Property Get UnderlyingOptionId() As Long
    clUnderlyingOptionId = GetTableValue("UnderlyingOptionId")
    UnderlyingOptionId = clUnderlyingOptionId
End Property
Public Property Let UnderlyingOptionId(lUnderlyingOptionId As Long)
    clUnderlyingOptionId = lUnderlyingOptionId
    SetTableValue "UnderlyingOptionId", lUnderlyingOptionId
    Call LoadUnderlyingOptionDetails
End Property


Public Property Get ManFilterID() As Long
    clManFilterID = GetTableValue("ManFilterId")
    ManFilterID = clManFilterID
End Property
Public Property Let ManFilterID(lManFilterID As Long)
    clManFilterID = lManFilterID
    SetTableValue "ManFilterId", lManFilterID
End Property


Public Property Get OptionID() As String
    OptionID = GetTableValue("OptionId")
End Property
Public Property Let OptionID(sOptionId As String)
    SetTableValue "OptionId", sOptionId
End Property


Public Property Get Inclusion() As String
    Inclusion = GetTableValue("Inclusion")
End Property
Public Property Let Inclusion(sInclusion As String)
    SetTableValue "Inclusion", sInclusion
End Property

Public Property Get FilterBy() As String
    FilterBy = GetTableValue("FilterBy")
End Property
Public Property Let FilterBy(sFilterBy As String)
    SetTableValue "FilterBy", sFilterBy
End Property

Public Property Get Operator() As String
    Operator = GetTableValue("Operator")
End Property
Public Property Let Operator(sOperator As String)
    SetTableValue "Operator", sOperator
End Property
'

Public Property Get RequiredValueIdx() As String
    RequiredValueIdx = GetTableValue("RequiredValueIdx")
End Property
Public Property Let RequiredValueIdx(sRequiredValueIdx As String)
    SetTableValue "RequiredValueIdx", sRequiredValueIdx
End Property

Public Property Get RequiredValue() As String
    RequiredValue = GetTableValue("RequiredValue")
End Property
Public Property Let RequiredValue(sRequiredValue As String)
    SetTableValue "RequiredValue", sRequiredValue
End Property

Public Property Get DtAdded() As Date
    DtAdded = GetTableValue("DtAdded")
End Property
Public Property Let DtAdded(dDtAdded As Date)
    SetTableValue "DtAdded", dDtAdded
End Property

Public Property Get AddUser() As String
    AddUser = GetTableValue("AddUser")
End Property
Public Property Let AddUser(sAddUser As String)
    SetTableValue "AddUser", sAddUser
End Property


Public Property Get SourceTableName() As String
    SourceTableName = coSourceTable.TableName
End Property

Public Property Get IdFieldName() As String
    IdFieldName = coSourceTable.IdFieldName
End Property

'Public Function NewId() As Long
'On Error GoTo Block_Err
'Dim strProcName As String
'Dim oAdo As clsADO
'Dim lNewId As Long
'    ' need to insert minimal data to get a new ID
'Stop
'    Set oAdo = New clsADO
'    With oAdo
'        .ConnectionString = CodeConnString
'        .SQLTextType = StoredProc
'        .sqlString = "usp_LETTER_Automation_ManFilterOptionAdd"
'        .Parameters.Refresh
'        .Parameters("@pManFilterid") = Me.ManFilterID
'        .Parameters("@pUnderlyingOptionId") = Me.UnderlyingOptionId
'        .Execute
'        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
'            Stop
'        End If
'        lNewId = .Parameters("@pNewId").Value
'        If lNewId > 0 Then
'            Me.WasInitialized = True
'            clManFilterRID = lNewId
'            coSourceTable.ID = lNewId
'            coSourceTable.WasInitialized = True
'        End If
'    End With
'    NewId = lNewId
'Block_Exit:
'    Set oAdo = Nothing
'    Exit Function
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Function



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


Public Function Fields() As Collection
    Set Fields = coSourceTable.Fields
End Function




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
    ReportError Err, strProcName, sSql
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
    ReportError Err, strProcName, sSpName
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
    ReportError Err, strProcName, "User ID: " & GetUserName() & " " & strFieldName
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
    ReportError Err, strProcName, "User ID: " & GetUserName() & " " & strFieldName
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
    ReportError Err, strProcName, "User ID: " & GetUserName()
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function Delete() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".Delete"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ManualFiltersOptionDel"
        .Parameters.Refresh
        .Parameters("@pRID") = Me.ID
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            GoTo Block_Exit
        End If
    End With
    

    Delete = True
Block_Exit:
    Set oAdo = Nothing
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
Public Function LoadFromId(lManFilterRID As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
    ID = lManFilterRID   ' if we did this we'
    
'    coSourceTable.ConnectionString = GetConnectString("CMS_AUDITORS_CLAIMS")
'    coSourceTable.ConnectionString = DataConnString
    
    LoadFromId = coSourceTable.LoadFromId(lManFilterRID)
    WasInitialized = LoadFromId
    Call LoadUnderlyingOptionDetails
    
Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    ReportError Err, strProcName, "lManFilterId: " & lManFilterRID
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromName(sFilterName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = True
    coSourceTable.IDStr = sFilterName   ' if we did this we'
    

    LoadFromName = coSourceTable.LoadFromIDStr(sFilterName)
    WasInitialized = LoadFromName
    Call LoadUnderlyingOptionDetails

Block_Exit:
    Exit Function

Block_Err:
    LoadFromName = False
    ReportError Err, strProcName, "sFilterName: " & sFilterName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadUnderlyingOptionDetails() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".LoadUnderlyingOptionDetails"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM LETTER_Automation_Xref_ManualFilters WHERE OptionId = " & CStr(Me.UnderlyingOptionId)
        Set oRs = .ExecuteRS
        If oRs.EOF And oRs.BOF Then
Stop
            GoTo Block_Exit
        End If
        
    End With

    Me.PossibleValuesSproc = Nz(oRs("PossibleValuesSproc").Value, "")
    Me.SampleSource = Nz(oRs("SampleSource").Value, "")
    Me.SampleFieldName = Nz(oRs("SampleFieldName").Value, "")
    

    LoadUnderlyingOptionDetails = True
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function

Block_Err:
    LoadUnderlyingOptionDetails = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetWhereClause() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sWhere As String

    strProcName = ClassName & ".GetWhereClause"
    
    
    If Me.SampleFieldName = "" Then
        sWhere = Me.FilterBy & " " & Me.Operator & " '"
    Else
        sWhere = Me.SampleFieldName & " " & Me.Operator & " '"
    End If
    
    If Me.RequiredValueIdx <> "" Then
        sWhere = sWhere & Me.RequiredValueIdx
    Else
        sWhere = sWhere & Me.RequiredValue
    End If
    
    sWhere = sWhere & "' "
    
     
Block_Exit:
    GetWhereClause = sWhere
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

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