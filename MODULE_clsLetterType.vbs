Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 5/14/2014
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a letter type as found in the
'''     ~Claims.dbo.Letter_Type table
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 05/14/2014 - Created class
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


Private ciAccountId As Integer
Private csLetterType As String
Private csLetterDesc As String
Private csLetterSource As String
Private csAddrType As String
Private csTemplateLoc As String
Private cbForDsOnly As Boolean
Private cbTechDenial As Boolean
Private cbIsAdr As Boolean
Private cbIsFindingRRL As Boolean
Private cbQueueRequiresRelease As Boolean
Private cbIsTimeSensitive As Boolean


Private Const csIDFIELDNAME As String = "LetterType"
Private Const csTableName As String = "LETTER_Type"

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
    Public Property Get Id() As String
        Id = LetterType
    End Property
    Public Property Let Id(sNewId As String)
        LetterType = sNewId
    End Property




Public Property Get AccountID() As Integer
    AccountID = ciAccountId
End Property
Public Property Let AccountID(iAccount As Integer)
    ciAccountId = ciAccountId
End Property

Public Property Get LetterType() As String
    LetterType = csLetterType
End Property
Public Property Let LetterType(sLetterType As String)
    If csLetterType <> sLetterType Then
        csLetterType = sLetterType
        Call LoadFromId(csLetterType)
    End If
End Property

Public Property Get LetterDesc() As String
    LetterDesc = GetTableValue("LetterDesc")
End Property
Public Property Let LetterDesc(sLetterDesc As String)
    SetTableValue "LetterDesc", sLetterDesc
End Property


Public Property Get LetterSource() As String
    LetterSource = GetTableValue("LetterSource")
End Property
Public Property Let LetterSource(sLetterSource As String)
    SetTableValue "LetterSource", sLetterSource
End Property

Public Property Get AddrType() As String
    AddrType = GetTableValue("AddrType")
End Property
Public Property Let AddrType(sAddrType As String)
    SetTableValue "AddrType", sAddrType
End Property

Public Property Get TemplateLoc() As String
    TemplateLoc = GetTableValue("TemplateLoc")
End Property
Public Property Let TemplateLoc(sTemplateLoc As String)
    SetTableValue "TemplateLoc", sTemplateLoc
End Property

Public Property Get ForDsOnly() As Boolean
    ForDsOnly = GetTableValue("ForDsOnly")
End Property
Public Property Let ForDsOnly(bForDsOnly As Boolean)
    SetTableValue "ForDsOnly", IIf(bForDsOnly, 1, 0)
End Property

Public Property Get TechDenial() As Boolean
    TechDenial = GetTableValue("TechDenial")
End Property
Public Property Let TechDenial(bTechDenial As Boolean)
    SetTableValue "TechDenial", IIf(bTechDenial, 1, 0)
End Property

Public Property Get IsAdr() As Boolean
    IsAdr = GetTableValue("IsAdr")
End Property
Public Property Let IsAdr(bIsAdr As Boolean)
    SetTableValue "IsAdr", IIf(bIsAdr, 1, 0)
End Property

Public Property Get IsFindingRRL() As Boolean
    IsFindingRRL = GetTableValue("IsFindingRRL")
End Property
Public Property Let IsFindingRRL(bIsFindingRRL As Boolean)
    SetTableValue "IsFindingRRL", IIf(bIsFindingRRL, 1, 0)
End Property

Public Property Get QueueRequiresRelease() As Boolean
    QueueRequiresRelease = GetTableValue("QueueRequiresRelease")
End Property
Public Property Let QueueRequiresRelease(bQueueRequiresRelease As Boolean)
    SetTableValue "QueueRequiresRelease", IIf(bQueueRequiresRelease, 1, 0)
End Property


' cbIsTimeSensitive
Public Property Get IsTimeSensitive() As Boolean
    IsTimeSensitive = GetTableValue("TimeSensitive")
End Property
Public Property Let IsTimeSensitive(bIsTimeSensitive As Boolean)
    SetTableValue "TimeSensitive", IIf(bIsTimeSensitive, 1, 0)
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
Public Function LoadFromId(sLetterType As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = True
    'ID = sLetterType   ' if we did this we'
    
'    coSourceTable.ConnectionString = GetConnectString("CMS_AUDITORS_CLAIMS")
    
    LoadFromId = coSourceTable.LoadFromIDStr(sLetterType)
    WasInitialized = LoadFromId

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    ReportError Err, strProcName, "LetterType: " & sLetterType
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