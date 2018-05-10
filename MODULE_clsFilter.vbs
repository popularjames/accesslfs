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
Private clFilterId As Long



Private clAccountId As Long

Private csOptionNum As String
Private csInclusion As String
Private csFilterBy As String
Private csOperator As String
Private csRequiredValueIndex As String
Private csRequiredValue As String

Private cdtcFilterOptions As Scripting.Dictionary
Private coFilterOptions As Collection


Private Const csIDFIELDNAME As String = "ManFilterId"
Private Const csTableName As String = "LETTER_Automation_ManualFilters"

    '' The table to use for the connection string to the _ERAC database
Private Const csSP_TABLENAME As String = "v_Code_Database"

Private coSourceTable As clsTable

Private cblnDirtyData As Boolean
Private cblnIsInitialized As Boolean


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
    Public Property Get Id() As Long
        Id = ManFilterID
    End Property
    Public Property Let Id(lNewId As Long)
        ManFilterID = lNewId
    End Property


Public Property Get AccountID() As Long
    AccountID = clAccountId
End Property
Public Property Let AccountID(lAccount As Long)
    clAccountId = lAccount
End Property

Public Property Get ManFilterID() As Long
'    clFilterId = GetTableValue("ManFilterId")
    ManFilterID = clFilterId
End Property
Public Property Let ManFilterID(lFilterId As Long)
    clFilterId = lFilterId
    Call LoadFromId(lFilterId)
End Property

' csFilterName
Public Property Get FilterName() As String
    csFilterName = GetTableValue("FilterName")
    FilterName = csFilterName
End Property
Public Property Let FilterName(sFilterName As String)
    Call SetTableValue("FilterName", sFilterName)
    csFilterName = sFilterName
End Property


Public Property Get Active() As Boolean
    Active = GetTableValue("Active")
End Property
Public Property Let Active(bActive As Boolean)

    SetTableValue "Active", CStr(bActive), True
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
Public Function LoadFromId(lManFilterID As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
    coSourceTable.IdFieldName = csIDFIELDNAME
    clFilterId = lManFilterID   ' if we did this we'
    coSourceTable.LoadSingleRow = True
    
    'coSourceTable.ConnectionString = GetConnectString("CMS_AUDITORS_CLAIMS")
'    coSourceTable.ConnectionString = DataConnString
    
    
    LoadFromId = coSourceTable.LoadFromId(lManFilterID)
    WasInitialized = LoadFromId

    Call LoadFilterOptions

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    ReportError Err, strProcName, "lManFilterId: " & lManFilterID
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

    strProcName = ClassName & ".LoadFromName"
    coSourceTable.IdIsString = True
    coSourceTable.IdFieldName = ""
    'ID = sLetterType   ' if we did this we'
    
'    coSourceTable.ConnectionString = GetConnectString("CMS_AUDITORS_CLAIMS")
'    coSourceTable.ConnectionString = DataConnString
    
    LoadFromName = coSourceTable.LoadFromIDStr(sFilterName)
    WasInitialized = LoadFromName
    
    Call LoadFilterOptions
    
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
Public Function LoadFilterOptions() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFltrOption As clsFilterOption
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".LoadFilterOptions"
    
    Set oFltrOption = New clsFilterOption
    
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = DataConnString
        .SQLTextType = sqltext
        .sqlString = "SELECT * FROM " & oFltrOption.SourceTableName & " WHERE " & csIDFIELDNAME & " = " & Me.Id & " ORDER BY OptionId "
        Set oRs = .ExecuteRS
        If oRs.BOF And oRs.EOF Then
            ' nothing retrieved
'            Stop
            ' possibly just no options added yet?
            GoTo Block_Exit
        End If
    End With
    
    While Not oRs.EOF
        Set oFltrOption = New clsFilterOption
        oFltrOption.LoadFromId (oRs(oFltrOption.IdFieldName).Value)
'        If cdtcFilterOptions.Exists(oRs("OptionId").Value) = True Then
        If cdtcFilterOptions.Exists(oRs("RID").Value) = True Then
            Stop
            GoTo Nxt
'            GoTo Block_Exit
        End If
        cdtcFilterOptions.Add oRs("RID").Value, oFltrOption
Nxt:
        oRs.MoveNext
    Wend
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function

Block_Err:
    LoadFilterOptions = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function NewFilterOption(lUnderlyingOptionId As Long) As clsFilterOption
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim lNewId As Long
Dim oFltrOpt As clsFilterOption
    ' need to insert minimal data to get a new ID

    strProcName = ClassName & ".NewFilterOption"
'Stop
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_ManFilterOptionAdd"
        .Parameters.Refresh
        .Parameters("@pManFilterid") = Me.ManFilterID
        .Parameters("@pUnderlyingOptionId") = lUnderlyingOptionId
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
        lNewId = .Parameters("@pNewId").Value
    End With
    
    Set oFltrOpt = New clsFilterOption
    If oFltrOpt.LoadFromId(lNewId) = False Then
        Stop
    End If
    
    '' Add this to our dictionary...
    Call AddFilterOption(oFltrOpt)
    
    Set NewFilterOption = oFltrOpt
Block_Exit:
    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function FilterOptions() As Collection
On Error GoTo Block_Err
Dim strProcName As String
Dim vKey As Variant

    strProcName = ClassName & ".FilterOptions"
    
    If Not coFilterOptions Is Nothing Then
        GoTo Block_Exit
    End If
    
    Set coFilterOptions = New Collection
    For Each vKey In cdtcFilterOptions.Keys
        coFilterOptions.Add cdtcFilterOptions.Item(vKey)
    Next
    
Block_Exit:
    Set FilterOptions = coFilterOptions
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
Public Function AddFilterOption(oFltrOption As clsFilterOption) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".AddFilterOption"
    
    If oFltrOption.ManFilterID = 0 Then
        oFltrOption.ManFilterID = Me.ManFilterID
    End If
    
    If oFltrOption.FilterOptionID = 0 Then
        oFltrOption.SaveNow
    End If
    
    ' If we don't have an option then we need to assignn one..
    If oFltrOption.OptionID = "" Then
        oFltrOption.OptionID = Format(cdtcFilterOptions.Count + 1, "000")
    '           oFltrOption.OptionId = Format(oFltrOption.ID, "000")
    End If

    If cdtcFilterOptions.Exists(oFltrOption.Id) = True Then
        Stop
        ' reset it
        Set cdtcFilterOptions.Item(oFltrOption.Id) = oFltrOption
    End If
    cdtcFilterOptions.Add oFltrOption.Id, oFltrOption
    
    oFltrOption.SaveNow

    AddFilterOption = True
Block_Exit:
    Exit Function

Block_Err:
    AddFilterOption = False
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function GetFilterFromClause() As String
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".GetFilterFromClause"

    GetFilterFromClause = " FROM v_LETTER_Automation_ManualOverrideSample "
Block_Exit:
    Exit Function

Block_Err:
    GetFilterFromClause = ""
    ReportError Err, strProcName
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
        .sqlString = "usp_LETTER_Automation_ManualFiltersDel"
        .Parameters.Refresh
        .Parameters("@pManFilterId") = Me.Id
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
Public Function GetFilterWhereClause() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim sSql As String
Dim sFrom As String
Dim iItem As Integer
Dim oFltrItm As clsFilterOption
Dim sOptionId As String
Dim sWhere As String
Dim vVar As Variant
Dim bFirst As Boolean


    strProcName = ClassName & ".GetFilterWhereClause"
    bFirst = True

    sWhere = " WHERE ("
'Stop
'    For iItem = 1 To cdtcFilterOptions.Count
'        sOptionId = Format(iItem, "000")
'        If cdtcFilterOptions.Exists(sOptionId) = False Then
'            Stop
'        End If
'
'        Set oFltrItm = cdtcFilterOptions.Item(sOptionId)
'        sWhere = sWhere & oFltrItm.Inclusion & " " & oFltrItm.GetWhereClause
'
'    Next
'    For Each oFltrItm In cdtcFilterOptions.Items
    For Each vVar In cdtcFilterOptions.Keys
        If TypeName(cdtcFilterOptions.Item(vVar)) = "clsFilterOption" Then
            Set oFltrItm = cdtcFilterOptions.Item(vVar)
            If bFirst = True Then
                sWhere = sWhere & " " & oFltrItm.GetWhereClause
                bFirst = False
            Else
                sWhere = sWhere & oFltrItm.Inclusion & " " & oFltrItm.GetWhereClause
            End If

        Else
            Stop
        End If
    Next
    sWhere = sWhere & ")"
    

    GetFilterWhereClause = sWhere
Block_Exit:
    Exit Function

Block_Err:
    GetFilterWhereClause = ""
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function RemoveFilterOption(oFltrOpt As clsFilterOption) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim vVar As Variant
Dim oFltrItm As clsFilterOption

    strProcName = ClassName & ".RemoveFilterOption"
    
    For Each vVar In cdtcFilterOptions.Keys
        If TypeName(cdtcFilterOptions.Item(vVar)) = "clsFilterOption" Then
            Set oFltrItm = cdtcFilterOptions.Item(vVar)
            If oFltrItm.Id = oFltrOpt.Id Then
'                Stop
                cdtcFilterOptions.Remove (vVar)
                ' delete it from the database too
                oFltrOpt.Delete
            End If
        End If
    Next
    
    
Block_Exit:
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
    Set cdtcFilterOptions = New Scripting.Dictionary
    
End Sub


Private Sub Class_Terminate()
    If Dirty = True Then
        SaveNow
    End If
    Set coSourceTable = Nothing
 
    cblnIsInitialized = False
    Set cdtcFilterOptions = Nothing
End Sub