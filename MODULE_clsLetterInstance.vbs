Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 9/12/2014
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a CMS Letter
'''
'''  TODO:
'''  =====================================
'''  -
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 09/12/2014 - KD: Added SectionCount and made it load from the table
'''  - 08/06/2013 - Created class
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


Public Event LetterError(ErrMsg As String, ErrNum As Long, ErrSource As String)


Private csInstanceIdQRCodePath As String

Private cbIsDuplexJob As Boolean

Private Const csIDFIELDNAME As String = "InstanceId"
Private Const csTableName As String = "Letter_Xref"
Private coSourceTable As clsTable

Private cbErrorOccurred As Boolean
Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private cstrInstanceId As String
Private clngAccountId As Long

Private ciPageCount As Integer
Private ciSectionCount As Integer
Private csDocType As String
Private csProvName As String

Private csContactTitle As String
Private csContactType As String

Private csContactName As String
Private csStreet1 As String
Private csStreet2 As String
Private csStreet3 As String
Private csCity As String
Private csState As String
Private csZip As String

Private csLetterType As String
Private csProvNum As String

Private cdtLetterCreateDt As Date
Private cdtLetterReqDt As Date

Private csAuditor As String
Private csLetterName As String  ' Cause this is what the table calls it!

Private csCnlyProvId As String
Private clLetterNumInBatch As Long

Private clBatchId As Long

Private clContractId As Long


Private csLetterQueueStatus As String


''##########################################################
''##########################################################
''##########################################################
'' Class state properties
''##########################################################
''##########################################################
''##########################################################

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get InstanceId() As String
    InstanceId = cstrInstanceId
End Property
Public Property Let InstanceId(sInstanceId As String)
    cstrInstanceId = sInstanceId
    WasInitialized = True
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As String
        ID = InstanceId
    End Property
    Public Property Let ID(sNewId As String)
        InstanceId = sNewId
    End Property


Public Property Get AccountID() As Long
    If clngAccountId < 0 Then clngAccountId = 1
    AccountID = clngAccountId
End Property
Public Property Let AccountID(lngAccountId As Long)
    clngAccountId = lngAccountId
End Property

Public Property Get ContractId() As Long
    ContractId = clContractId
End Property
Public Property Let ContractId(lContractId As Long)
    clContractId = lContractId
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
''
'Public Property Get GetField(sFieldName As String) As String
'    GetField = CStr("" & GetTableValue(sFieldName))
'End Property
'Public Property Let SetField(sFieldName As String, sFieldValue As String)
'    SetTableValue sFieldName, sFieldValue
'End Property
'
'
'Public Function Fields() As Collection
'    Set Fields = coSourceTable.Fields
'End Function




''##########################################################
''##########################################################
''##########################################################
'' Properties according to Letter_Xref table
''##########################################################
''##########################################################
''##########################################################
Public Property Get LetterCreateDt() As Date
    LetterCreateDt = cdtLetterCreateDt
End Property
Public Property Let LetterCreateDt(dtLetterCreateDt As Date)
    cdtLetterCreateDt = dtLetterCreateDt
End Property

Public Property Get LetterReqDt() As Date
    LetterReqDt = cdtLetterReqDt
End Property
Public Property Let LetterReqDt(dtLetterReqDt As Date)
    cdtLetterReqDt = dtLetterReqDt
End Property

Public Property Get LetterBatchId() As Long
    LetterBatchId = clBatchId
End Property
Public Property Let LetterBatchId(lBatchId As Long)
    clBatchId = lBatchId
End Property
    ' let's alias it
    Public Property Get BatchID() As Long
        BatchID = LetterBatchId
    End Property
    Public Property Let BatchID(lBatchId As Long)
        LetterBatchId = lBatchId
    End Property

Public Property Get Auditor() As String
    Auditor = csAuditor
End Property
Public Property Let Auditor(sAuditor As String)
    csAuditor = sAuditor
End Property


'Public Property Get LetterType() As String
'    LetterType = GetTableValue("LetterType")
'End Property
'Public Property Let LetterType(sLetterType As String)
'    SetTableValue "LetterType", sLetterType
'End Property

Public Property Get LetterType() As String
    LetterType = csLetterType
End Property
Public Property Let LetterType(sLetterType As String)
    csLetterType = sLetterType
End Property



Public Property Get LetterName() As String
    If csLetterName = "" Then
        csLetterName = GetTableValue("LetterFilePath")
    End If
    LetterName = csLetterName
End Property
Public Property Let LetterName(sLetterName As String)
    csLetterName = sLetterName
End Property
    Public Property Get LetterPath() As String
        LetterPath = LetterName
    End Property
    Public Property Let LetterPath(sLetterPath As String)
        LetterName = sLetterPath
    End Property

Public Property Get LetterNumInBatch() As Long
    LetterNumInBatch = clLetterNumInBatch
End Property
Public Property Let LetterNumInBatch(lLetterNum As Long)
    clLetterNumInBatch = lLetterNum
End Property

Public Property Get LetterFileName() As String
    LetterFileName = GetFileName(LetterName)
End Property

'
'Public Property Get ProvNum() As String
'    ProvNum = GetTableValue("ProvNum")
'End Property
'Public Property Let ProvNum(sProvNum As String)
'    SetTableValue "ProvNum", sProvNum
'End Property
Public Property Get ProvNum() As String
    ProvNum = csProvNum
End Property
Public Property Let ProvNum(sProvNum As String)
    csProvNum = sProvNum
End Property

Public Property Get cnlyProvID() As String
    cnlyProvID = csCnlyProvId
End Property
Public Property Let cnlyProvID(sCnlyProvId As String)
    csCnlyProvId = sCnlyProvId
End Property





''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''

Public Property Get PageCount() As Integer
    PageCount = ciPageCount
End Property
Public Property Let PageCount(iPageCount As Integer)
    ' should we go get it?
    ciPageCount = iPageCount
End Property

Public Property Get SectionCount() As Integer
    SectionCount = 0
    Exit Property
    If ciSectionCount = 0 Then
        ciSectionCount = Nz(GetTableValue("SectionCount"), 0)
    End If
    SectionCount = ciSectionCount
End Property
Public Property Let SectionCount(iSectionCount As Integer)
     ciSectionCount = iSectionCount
End Property

Public Property Get InstanceQRCodePath() As String
    InstanceQRCodePath = csInstanceIdQRCodePath
End Property
Public Property Let InstanceQRCodePath(sInstanceIdQRCodePath As String)
    csInstanceIdQRCodePath = sInstanceIdQRCodePath
End Property

Public Property Get IsDuplexJob() As Boolean
    IsDuplexJob = cbIsDuplexJob
End Property
Public Property Let IsDuplexJob(bIsDuplexJob As Boolean)
    cbIsDuplexJob = bIsDuplexJob
End Property



'' .Doc, pdf
Public Property Get DocumentType() As String
    DocumentType = FileExtension(Me.LetterPath)
End Property

'' address fields that are used on the letter
Public Property Get ProvName() As String
    ProvName = csProvName
End Property
Public Property Let ProvName(sProvName As String)
    csProvName = sProvName
End Property



Public Property Get ContactTitle() As String
    ContactTitle = csContactTitle
End Property
Public Property Let ContactTitle(sContactTitle As String)
    csContactTitle = sContactTitle
End Property

Public Property Get ContactName() As String
    ContactName = csContactName
End Property
Public Property Let ContactName(sContactName As String)
    csContactName = sContactName
End Property

Public Property Get ContactType() As String
    ContactType = csContactName
End Property
Public Property Let ContactType(sContactType As String)
    csContactType = sContactType
End Property

Public Property Get Street1() As String
    Street1 = csStreet1
End Property
Public Property Let Street1(sStreet1 As String)
    csStreet1 = sStreet1
End Property


Public Property Get Street2() As String
    Street2 = csStreet2
End Property
Public Property Let Street2(sStreet2 As String)
    csStreet2 = sStreet2
End Property


Public Property Get Street3() As String
    Street3 = csStreet3
End Property
Public Property Let Street3(sStreet3 As String)
    csStreet3 = sStreet3
End Property




Public Property Get City() As String
    City = csCity
End Property
Public Property Let City(sCity As String)
    csCity = sCity
End Property


Public Property Get State() As String
    State = csState
End Property
Public Property Let State(sState As String)
    csState = sState
End Property

Public Property Get Zip() As String
    Zip = csZip
End Property
Public Property Let Zip(sZip As String)
    csZip = sZip
End Property


Public Property Get LetterQueueStatus() As String
    LetterQueueStatus = csLetterQueueStatus
End Property
Public Property Let LetterQueueStatus(sLetterQueueStatus As String)
    csLetterQueueStatus = sLetterQueueStatus
End Property




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''

Public Function SaveStaticDetails(Optional bNew As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sLtrFileName As String
Dim sExt As String

    strProcName = ClassName & ".SaveStaticDetails"
'Stop

    Call PathInfoFromPath(Me.LetterPath, sLtrFileName, , sExt)
    If Right(LCase(sLtrFileName), Len(sExt)) <> LCase(sExt) Then
        If left(sExt, 1) = "." Then
            sLtrFileName = sLtrFileName & sExt
        Else
            sLtrFileName = sLtrFileName & "." & sExt
        End If
    End If
    
    If Me.ContractId = 0 Then Me.ContractId = 100
    
    
    If bNew = True Then
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_LETTER_Automation_SaveStaticDetails"
            .Parameters.Refresh
            .Parameters("@pAccountId") = Me.AccountID
            .Parameters("@pInstanceId") = Me.InstanceId
            .Parameters("@pPageCount") = Me.PageCount
            .Parameters("@pLetterPath") = Me.LetterPath
            .Parameters("@pLetterFileName") = sLtrFileName
            .Parameters("@pLetterBatchId") = Me.BatchID
            .Parameters("@pMergeRunId") = CurrentProcessor.ThisQueueRunId
            .Parameters("@pDocNumInCurrentBatch") = Me.LetterNumInBatch
            .Parameters("@pSectionCount") = Me.SectionCount
            .Parameters("@pContractId") = Me.ContractId
            .Execute
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
                LogMessage strProcName, "ERROR", "An error occurred saving page count for instance: " & Me.InstanceId, .Parameters("@pErrMsg").Value
                Stop
                GoTo Block_Exit
                
            End If
        End With
    Else
        ' for the old way
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = CodeConnString
            .SQLTextType = StoredProc
            .sqlString = "usp_LETTER_SaveStaticDetails"
            .Parameters.Refresh
'            .Parameters("@pAccountId") = Me.AccountID
            .Parameters("@pInstanceId") = Me.InstanceId
            .Parameters("@pPageCount") = Me.PageCount
            .Parameters("@pLetterPath") = Me.LetterPath
            .Parameters("@pLetterFileName") = sLtrFileName
            .Parameters("@pLetterBatchId") = Me.BatchID
            .Parameters("@pSectionCount") = Me.SectionCount
            .Parameters("@pContractid") = Me.ContractId
            .Execute
            If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
    'Stop
                LogMessage strProcName, "ERROR", "An error occurred saving page count for instance: " & Me.InstanceId, .Parameters("@pErrMsg").Value
                GoTo Block_Exit
                Stop
            End If
        End With
    End If
    '' Should probably reload the details here

    SaveStaticDetails = True

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function UpdateStaticDetails(sCombinedDocFolderPath As String, lCombinedDocNum As Long, Optional sMailRoomFilePath As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".UpdateStaticDetails"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = CodeConnString
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_Automation_UpdateStaticDetails"
        .Parameters.Refresh
        .Parameters("@pAccountId") = gintAccountID
        .Parameters("@pInstanceId") = Me.InstanceId
        .Parameters("@pLetterBatchId") = Me.BatchID
        .Parameters("@pCombineRunId") = CurrentProcessor.ThisQueueRunId
        .Parameters("@pCombinedFolderPath") = sCombinedDocFolderPath
        .Parameters("@pCombinedDocNum") = lCombinedDocNum
        If sMailRoomFilePath <> "" Then
            .Parameters("@pMailRoomFilePath") = sMailRoomFilePath
        End If
'Stop
        .Execute
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
            LogMessage strProcName, "ERROR", "An error occurred saving page count for instance: " & Me.InstanceId, .Parameters("@pErrMsg").Value
            GoTo Block_Exit
            Stop
        End If
    End With

    '' Should probably reload the details here

    UpdateStaticDetails = True

Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function



Private Function GetAddressDetails() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetAddressDetails"
        ' usp_LETTER_GetAddressByInstanceId

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_LETTER_GetAddressByInstanceId"
        .Parameters.Refresh
        .Parameters("@pInstanceId") = Me.InstanceId
        Set oRs = .ExecuteRS
        If .GotData = False Or Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "Could not get the address for instance ID: " & Me.InstanceId, Nz(.Parameters("@pErrMsg").Value, "")
            GoTo Block_Exit
        End If
    End With

    Me.ProvName = oRs("ProvName").Value
    Me.ContactTitle = oRs("ContactTitle").Value
    Me.ContactName = oRs("ContactName").Value
    Me.ContactType = oRs("ContactType").Value
    
    Me.Street1 = oRs("Addr01").Value
    Me.Street2 = oRs("Addr02").Value
    Me.Street3 = oRs("Addr03").Value
    
    Me.City = oRs("City").Value
    Me.State = oRs("State").Value
    
    Me.Zip = oRs("Zip").Value
    
    
    GetAddressDetails = True
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromId(sInstanceId As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = True
    ID = sInstanceId
    LoadFromId = coSourceTable.LoadFromIDStr(sInstanceId)
    WasInitialized = LoadFromId

    ' Need to get the address information
Stop

    ' Kev, if you are actually using this then you need to
    ' add some code here to loop through the fields and populate our static variables...
    ' or, change the properties to something like: if IsField("") then = coSourceTable.GetValue(""

    If GetAddressDetails() = False Then
        Stop
    End If

    Call LoadQRCodePath

Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "Instance ID: " & sInstanceId
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
'''
Public Function LoadFromRS(oRs As ADODB.RecordSet) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromRS"
    coSourceTable.IdIsString = True
    ID = oRs("InstanceId").Value
    LoadFromRS = coSourceTable.InitializeFromRS(oRs)
    WasInitialized = LoadFromRS

    ' Need to get the address information
'Stop

    ' Kev, if you are actually using this then you need to
    ' add some code here to loop through the fields and populate our static variables...
    ' or, change the properties to something like: if IsField("") then = coSourceTable.GetValue(""

'    If GetAddressDetails() = False Then
'        Stop
'    End If
    Call LoadQRCodePath

Block_Exit:
    Exit Function

Block_Err:
    LoadFromRS = False
    FireError Err, strProcName, "Instance ID: " & ID
    GoTo Block_Exit
End Function



Public Function RefreshObject() As Boolean
Dim sInstanceId As String

    sInstanceId = Me.InstanceId

    Call Class_Initialize
    
    Call LoadFromId(sInstanceId)
    
    'Call LoadSubObjects("")
    
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
    FireError Err, strProcName, "User ID: " & GetUserName() & " " & strFieldName
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

'    If WasInitialized = False Then
'        SetTableValue = False
'        GoTo Block_Exit
'    End If

    blnWorked = coSourceTable.SetTableValue(strFieldName, varValue, , blnSaveNow)
    If blnWorked = True And blnSaveNow = False Then
        Dirty = True
    End If
    SetTableValue = blnWorked

Block_Exit:
    Exit Function
    
Block_Err:
    SetTableValue = False
    FireError Err, strProcName, "User ID: " & GetUserName() & " " & strFieldName
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
    FireError Err, strProcName, "User ID: " & GetUserName()
    GoTo Block_Exit
End Function


Public Sub LoadQRCodePath()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim sVal As String

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "SELECT TOP 1 QRPath FROM LETTER_Barcode_Service_Details WHERE InstanceId = '" & Me.InstanceId & "'"
        Set oRs = .ExecuteRS
        If oRs.EOF And oRs.BOF Then
            Debug.Print "no qr code path..."
        Else
            sVal = Nz(oRs("QRPath").Value, "")
'            sVal = Replace(sVal, "\\", "\")
            
            InstanceQRCodePath = sVal
        End If
    End With
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
    
    RaiseEvent LetterError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub


Private Sub Class_Initialize()

    Set coSourceTable = New clsTable
    coSourceTable.IdIsString = True
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