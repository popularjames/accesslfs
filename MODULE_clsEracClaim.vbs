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
'''  Represents a tagged claim, basically a "hook" into the
'''
'''  TODO:
'''  =====================================
'''  - Maybe make actual properties instead of using GetField()
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 06/22/2012 - Added PayerNameId
'''  - 04/25/2012 - added LoadFromTaggedClaimI
'''  - 03/6/2012 - Created class
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

Public Event TaggedClaimError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean


Private Const csIDFIELDNAME As String = "CnlyClaimNum"
Private Const csTableName As String = "v_TaggedClaims"
Private coSourceTable As clsTable


Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean


Private csCnlyClaimNum As String


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get CnlyClaimNum() As String
    CnlyClaimNum = csCnlyClaimNum
End Property
Public Property Let CnlyClaimNum(sCnlyClaimNum As String)
    csCnlyClaimNum = sCnlyClaimNum
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As String
        ID = CnlyClaimNum
    End Property
    Public Property Let ID(sNewId As String)
        CnlyClaimNum = sNewId
    End Property



Public Property Get DataTypeCode() As String
    DataTypeCode = GetTableValue("DataType")
End Property



Public Property Get Icn() As String
    Icn = Trim(CStr("" & GetTableValue("ICN")))
End Property
Public Property Let Icn(sIcn As String)
    SetTableValue "ICN", sIcn
End Property



Public Property Get ProvNum() As String
    ProvNum = GetTableValue("ProvNum")
End Property
Public Property Let ProvNum(sProvNum As String)
    SetTableValue "ProvNum", sProvNum
End Property


Public Property Get cnlyProvID() As String
    cnlyProvID = GetTableValue("CnlyProvID")
End Property
Public Property Let cnlyProvID(sCnlyProvId As String)
    SetTableValue "CnlyProvID", sCnlyProvId
End Property



Public Property Get MedicalRecordNum() As String
    MedicalRecordNum = GetTableValue("MedicalRecordNum")
End Property
Public Property Let MedicalRecordNum(sMedicalRecordNum As String)
    SetTableValue "MedicalRecordNum", sMedicalRecordNum
End Property


Public Property Get ConceptID() As String
    ConceptID = GetTableValue("ConceptID")
End Property
Public Property Let ConceptID(sConceptId As String)
    SetTableValue "ConceptID", sConceptId
End Property



Public Property Get eRacTaggedClaimId() As Integer
    eRacTaggedClaimId = Nz(GetTableValue("eRacTaggedClaimId"), 0)
End Property
Public Property Let eRacTaggedClaimId(iEracTaggedClaimId As Integer)
    SetTableValue "eRacTaggedClaimId", iEracTaggedClaimId
End Property



Public Property Get PayerNameId() As Integer
    PayerNameId = Nz(GetTableValue("PayerNameId"), 0)
End Property
Public Property Let PayerNameId(iPayerNameId As Integer)
    SetTableValue "PayerNameId", iPayerNameId
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
Public Function LoadFromTaggedClaimId(lTaggedClaimId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sCnlyClaimNum As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".LoadFromTaggedClaimId"
    
    Set oRs = GetRecordset("SELECT CnlyClaimNum FROM v_TaggedClaims WHERE eRacTaggedClaimId = " & CStr(lTaggedClaimId))
    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    
    sCnlyClaimNum = Nz(oRs("CnlyClaimNum").Value, "")
    
    coSourceTable.IdIsString = True
    ID = sCnlyClaimNum
    LoadFromTaggedClaimId = coSourceTable.LoadFromIDStr(sCnlyClaimNum)
    WasInitialized = LoadFromTaggedClaimId


Block_Exit:
    Exit Function

Block_Err:
    FireError Err, strProcName, "CnlyClaimNum: " & sCnlyClaimNum
    LoadFromTaggedClaimId = False
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromId(sCnlyClaimNum As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = True
    ID = sCnlyClaimNum
    LoadFromId = coSourceTable.LoadFromIDStr(sCnlyClaimNum)
    WasInitialized = LoadFromId


Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "CnlyClaimNum: " & sCnlyClaimNum
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
    
    RaiseEvent TaggedClaimError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

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