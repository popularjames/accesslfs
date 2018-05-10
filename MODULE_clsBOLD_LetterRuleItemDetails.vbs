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

Public Event ItemChanged()
Public Event LetterRuleItemsError(ErrMsg As String, ErrNum As Long, ErrSource As String, bHandled As Boolean)

Private cdctItemKeys As Scripting.Dictionary

Private cbErrorOccurred As Boolean
Private csLastError As String
Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private ccolItemDetails As Collection



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get Items() As Collection
    Set Items = ccolItemDetails
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
'' Oh what would I do for __REAL__ inheritance?
''##########################################################
''##########################################################
''##########################################################
Public Function LoadFromItemType(Optional sItemType As String = "") As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim oItm As clsBOLD_LetterRuleItemDetail

    strProcName = ClassName & ".LoadFromItemType"
    
    Set oItm = New clsBOLD_LetterRuleItemDetail
    
    Set oDb = CurrentDb()
    If sItemType = "" Then
        Stop
        Set oRs = oDb.OpenRecordSet("SELECT LocalId FROM " & oItm.CurrentTableName & " ORDER BY LocalId ASC", dbOpenDynaset, dbSeeChanges)
    
    Else
        Set oRs = oDb.OpenRecordSet("SELECT LocalId FROM " & oItm.CurrentTableName & " WHERE ItemType = '" & sItemType & "' ORDER BY LocalId ASC", dbOpenDynaset, dbSeeChanges)
    End If
    
    
    'If ccolItemDetails Is Nothing Then
    Set ccolItemDetails = New Collection
    
    While Not oRs.EOF
        Set oItm = New clsBOLD_LetterRuleItemDetail
        oItm.LoadFromId (oRs("LocalId").Value)
        ccolItemDetails.Add oItm
        oRs.MoveNext
    Wend
    oRs.Close
    
    LoadFromItemType = True
    
Block_Exit:
    Set oRs = Nothing
    Set oDb = Nothing
    Set oItm = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function LoadFromRuleId(lRuleId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oItm As clsBOLD_LetterRuleItemDetail
Dim oDb As DAO.Database
Dim odbRS As DAO.RecordSet

    strProcName = ClassName & ".LoadFromRuleId"
    '' Here we need to copy it down to the local table
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_Letter_Automation_LoadRuleDetails"
        .Parameters.Refresh
        .Parameters("@pRuleId") = lRuleId
        Set oRs = .ExecuteRS
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            Stop
        End If
    End With
    
    Set ccolItemDetails = New Collection
    
    Set oItm = New clsBOLD_LetterRuleItemDetail
    
    If CopyDataToLocalTmpTableFromADORS(oRs, False, oItm.CurrentTableName) = "" Then
        Stop
    End If
    
    Set oDb = CurrentDb
    Set odbRS = oDb.OpenRecordSet("SELECT * FROM " & oItm.CurrentTableName & " WHERE RuleId = " & CStr(lRuleId), dbOpenDynaset, dbSeeChanges)
    
    While Not odbRS.EOF
        Set oItm = New clsBOLD_LetterRuleItemDetail
        If oItm.LoadFromRS(odbRS) = True Then
            ccolItemDetails.Add oItm
        Else
            Stop
        End If
        odbRS.MoveNext
    Wend
    
    LoadFromRuleId = True
    
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oDb = Nothing
    Set odbRS = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function SaveNow() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oItm As clsBOLD_LetterRuleItemDetail

    strProcName = ClassName & ".SaveNow"
    SaveNow = True  ' optimistic aren't we?
    
    Set oItm = New clsBOLD_LetterRuleItemDetail
    
    If ccolItemDetails Is Nothing Then GoTo Block_Exit
    
    For Each oItm In ccolItemDetails
        SaveNow = oItm.SaveNow
        If SaveNow = False Then GoTo Block_Exit
    Next
    
Block_Exit:
    Set oItm = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub RebuildDictionary()
On Error GoTo Block_Err
Dim strProcName As String
Dim oItm As clsBOLD_LetterRuleItemDetail

    strProcName = ClassName & ".RebuildDictionary"

    Set cdctItemKeys = New Scripting.Dictionary
    
    For Each oItm In Me.Items
        If cdctItemKeys.Exists(oItm.ItemName) = True Then
            Set cdctItemKeys.Item(oItm.ItemType & ":" & oItm.ItemName) = oItm
        Else
            cdctItemKeys.Add oItm.ItemType & ":" & oItm.ItemName, oItm
        End If
    Next

Block_Exit:
    Set oItm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Public Function ItemExists(sItemName As String, sDetailItemTypeGroup As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim bRefresh As Boolean

    strProcName = ClassName & ".ItemExists"
    
    If cdctItemKeys Is Nothing Then
        bRefresh = True
    End If
    
    
    If Me.Dirty = True Then
        bRefresh = True
    End If
    
    If bRefresh = True Then
        Call RebuildDictionary
    End If
    
    ItemExists = cdctItemKeys.Exists(sDetailItemTypeGroup & ":" & sItemName)
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function ItemExistsObj(oItem As clsBOLD_LetterRuleItemDetail) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim bRefresh As Boolean

    strProcName = ClassName & ".ItemExistsObj"
    
    If cdctItemKeys Is Nothing Then
        bRefresh = True
    End If
    
    
    If Me.Dirty = True Then
        bRefresh = True
    End If
    
    If bRefresh = True Then
        Call RebuildDictionary
    End If
    
    ItemExistsObj = cdctItemKeys.Exists(oItem.ItemType & ":" & oItem.ItemName)
    
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function AddItemObj(oItem As clsBOLD_LetterRuleItemDetail) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".AddItemObj"
    
    ' Insure we don't already have this one in there
    If Me.ItemExistsObj(oItem) = True Then
        AddItemObj = True
        GoTo Block_Exit
    End If
        
    If oItem.Id = 0 Then
        oItem.SecureLocalId (oItem.ItemType)
    End If
    
    With oItem
        .ItemName = .ItemName
        .ItemFieldName = .ItemFieldName
        .BooleanVal = .BooleanVal
        .ItemValue = .ItemValue
        .ListItemObject = .ListItemObject
        .SaveNow
    End With
    
    ccolItemDetails.Add oItem
    
    AddItemObj = True
    RaiseEvent ItemChanged
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function AddItem(sItemName As String, sItemFieldName As String, sItemType As String, Optional iItemBoolean As Integer, Optional sOperator As String, _
                            Optional sValue As String, Optional lItemId As Long, Optional lRemoteItemId As Long, Optional lRuleItemId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oItem As clsBOLD_LetterRuleItemDetail

    strProcName = ClassName & ".AddItem"
    
    Set oItem = New clsBOLD_LetterRuleItemDetail
    
    
    oItem.ItemType = sItemType
    oItem.ItemName = sItemName
    oItem.ItemFieldName = sItemFieldName
    oItem.BooleanVal = iItemBoolean
    oItem.Operator = sOperator
    oItem.ItemValue = sValue
    oItem.Id = lItemId
    oItem.RuleId = lRemoteItemId
    oItem.RuleItemId = lRuleItemId
    AddItem = AddItemObj(oItem)
    
    
Block_Exit:
    Set oItem = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function RemoveItemObj(oItem As clsBOLD_LetterRuleItemDetail) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iIdx As Integer
Dim oItm As clsBOLD_LetterRuleItemDetail
' ok, so, my brain hurts from sitting on it for too long
' I'm going to do this the lazy way...

    strProcName = ClassName & ".RemoveItemObj"
    
    For Each oItm In ccolItemDetails
        iIdx = iIdx + 1
        If oItm.Id = oItem.Id Then
            Exit For
        End If
    Next
    
    If oItem.Id = 0 Then
        oItem.SecureLocalId (oItem.ItemType)
    End If
    
    With oItem
        .ItemName = .ItemName
        .ItemFieldName = .ItemFieldName
        .BooleanVal = .BooleanVal
        
        .ItemValue = .ItemValue
        .DeleteNow
    End With
    

    ccolItemDetails.Remove iIdx
    RemoveItemObj = True
    RaiseEvent ItemChanged
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function RemoveItem(sItemName As String, sItemFieldName As String, sItemType As String, Optional iItemBoolean As Integer, Optional sOperator As String, _
                            Optional sValue As String, Optional lItemId As Long, Optional lRuleId As Long, Optional lRuleItemId As Long) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oItem As clsBOLD_LetterRuleItemDetail

    strProcName = ClassName & ".RemoveItem"
    
    Set oItem = New clsBOLD_LetterRuleItemDetail
    
    
    oItem.ItemType = sItemType
    oItem.ItemName = sItemName
    oItem.ItemFieldName = sItemFieldName
    oItem.BooleanVal = iItemBoolean
    oItem.Operator = sOperator
    oItem.ItemValue = sValue
    oItem.Id = lItemId
    oItem.RuleId = lRuleId
    oItem.RuleItemId = lRuleItemId
    
    RemoveItem = RemoveItemObj(oItem)
    
    
Block_Exit:
    Set oItem = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
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
    
    RaiseEvent LetterRuleItemsError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName, False)

End Sub

Private Sub Class_Initialize()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Class_Initialize"
    
    Set ccolItemDetails = New Collection
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Class_Terminate()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Class_Terminate"
    
    Set ccolItemDetails = Nothing
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub