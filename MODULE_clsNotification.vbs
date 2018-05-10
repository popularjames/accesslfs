Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 04/24/2012
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''  Represents a notification "job" which will encompas
'''     stuff like subject, body, to addresses, etc of the email notification
'''
'''  TODO:
'''  =====================================
'''  - Make it possible to create a new notification with this class
'''     (currently it requires an ID to exist)
'''
'''  HISTORY:
'''  =====================================
'''  - 04/24/2012 - added Body (EmailMsg)
'''  - 04/10/2012 - Created class
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

Public Event NotificationError(ErrMsg As String, ErrNum As Long, ErrSource As String)

Private cbErrorOccurred As Boolean


Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private Const csIDFIELDNAME As String = "NotificationId"
Private Const csTableName As String = "CONVERT_Notifications"
Private coSourceTable As clsTable


Private ciNotificationID As Integer



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property




Public Property Get NotificationID() As Integer
    NotificationID = ciNotificationID
End Property
Public Property Let NotificationID(iNotificationId As Integer)
    ciNotificationID = iNotificationId
End Property
        '' Just an alias for ease of use!
    Public Property Get ID() As Integer
        ID = NotificationID
    End Property
    Public Property Let ID(iNewId As Integer)
        NotificationID = iNewId
    End Property





Public Property Get NotificationName() As String
    NotificationName = GetTableValue("NotificationName")
End Property
Public Property Let NotificationName(sNotificationName As String)
    SetTableValue "NotificationName", sNotificationName
End Property



Public Property Get EmailTo() As String
    EmailTo = GetTableValue("EmailTo")
End Property
Public Property Let EmailTo(sEmailTo As String)
    SetTableValue "EmailTo", sEmailTo
End Property


    '' Body
Public Property Get EmailMsg() As String
    EmailMsg = GetTableValue("EmailMsg")
End Property
Public Property Let EmailMsg(sEmailMsg As String)
    SetTableValue "EmailMsg", sEmailMsg
End Property



Public Property Get EmailFrom() As String
    EmailFrom = GetTableValue("EmailFrom")
End Property
Public Property Let EmailFrom(sEmailFrom As String)
    SetTableValue "EmailFrom", sEmailFrom
End Property



Public Property Get EmailSubject() As String
    EmailSubject = GetTableValue("EmailSubject")
End Property
Public Property Let EmailSubject(sEmailSubject As String)
    SetTableValue "EmailSubject", sEmailSubject
End Property




Public Property Get Owner() As String
    Owner = GetTableValue("Owner")
End Property
Public Property Let Owner(sOwner As String)
    SetTableValue "Owner", sOwner
End Property




Public Property Get ModifyUser() As String
    ModifyUser = GetTableValue("ModifyUser")
End Property
Public Property Let ModifyUser(sModifyUser As String)
    SetTableValue "ModifyUser", sModifyUser
End Property





Public Property Get DateCreated() As Date
Dim sDate As String
    
    sDate = CStr("" & GetTableValue("DateCreated"))
    
    If IsDate(sDate) Then
        DateCreated = CDate(sDate)
    Else
        DateCreated = CDate("1/1/1900")
    End If
    
End Property
Public Property Let DateCreated(dDateCreated As Date)
    SetTableValue "DateCreated", Format(dDateCreated, "mm/dd/yyyy Hh:Nn:ss AM/PM")
End Property



Public Property Get DateModified() As Date
Dim sDate As String
    
    sDate = CStr("" & GetTableValue("DateModified"))
    
    If IsDate(sDate) Then
        DateModified = CDate(sDate)
    Else
        DateModified = CDate("1/1/1900")
    End If
    
End Property
Public Property Let DateModified(dDateModified As Date)
    SetTableValue "DateModified", Format(dDateModified, "mm/dd/yyyy Hh:Nn:ss AM/PM")
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
Private Function GetRecordsetSP(sSpName As String, Optional sParamString As String = "", Optional sTableName As String = csTableName) As ADODB.RecordSet
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
    FireError Err, strProcName, sSpName
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
    If TypeName(coSourceTable.GetTableValue(strFieldName)) = "Nothing" Then
        GoTo Block_Exit
    End If
    
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
    If blnWorked = True And blnSaveNow = False Then Dirty = True
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
Public Function LoadFromId(lNotificationId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromID"
    coSourceTable.IdIsString = False
    ID = lNotificationId
    LoadFromId = coSourceTable.LoadFromId(lNotificationId)
    WasInitialized = LoadFromId
    
Block_Exit:
    Exit Function

Block_Err:
    LoadFromId = False
    FireError Err, strProcName, "NotificationId: " & CStr(lNotificationId)
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
    
    RaiseEvent NotificationError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub

  ''##########################################################
Private Sub FireErrorStr(iErrNum As Integer, sDesc As String, sSource As String, Optional sAdditionalDetails As String)
Dim oErr As ErrObject

    Set oErr = New ErrObject
    With oErr
        .Number = iErrNum
        .Description = sDesc
        .Source = sSource
    End With
    FireError oErr, sSource, sAdditionalDetails
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