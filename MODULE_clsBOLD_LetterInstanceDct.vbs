Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 8/9/2013
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


Private Const csIDFIELDNAME As String = "InstanceId"
Private Const csTableName As String = "Letter_Xref"
Private coSourceTable As clsTable

Private cdctLetters As Scripting.Dictionary ' This holds the clsLetterInstance objects (Key is InstanceID)
Private cdctLetterPageCnt As Scripting.Dictionary ' this will hold the Key: InstanceId, Value = Page count of the physical letter

Private cbErrorOccurred As Boolean
Private cblnIsInitialized As Boolean
Private cblnDirtyData As Boolean

Private clBatchId As Long




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




'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''
'''' ##############################################################################
'''' ##############################################################################
'''' ##############################################################################
''''
''''


Public Property Get Count() As Integer
    Count = cdctLetters.Count
End Property


' clBatchId
Public Property Get BatchID() As Long
    BatchID = clBatchId
End Property
Public Property Let BatchID(lBatchId As Long)
    clBatchId = lBatchId
End Property


''##########################################################
''##########################################################
''##########################################################
'' Class methods
''##########################################################
''##########################################################
''##########################################################

Public Function Letters() As Collection
On Error GoTo Block_Err
Dim strProcName As String
Static colLetters As Collection
Dim vLetterInstance As Variant
Dim oLetter As clsLetterInstance

    strProcName = ClassName & ".Letters"
    
    If Me.Dirty = True Or colLetters Is Nothing Then
        Set colLetters = New Collection
        
        For Each vLetterInstance In cdctLetters.Keys
            If IsEmpty(vLetterInstance) Then
                Stop
            End If
            Set oLetter = cdctLetters.Item(vLetterInstance)
            If Me.BatchID <> 0 Then
                oLetter.BatchID = Me.BatchID
            End If
            colLetters.Add oLetter
        Next
        
    End If
    
    Set Letters = colLetters
    
Block_Exit:
    
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'' Here we are assuming that we have all details including page count
'' in the object that is passed.. so we have to update both dictionaries
Public Function UpdateLetter(oLetter As clsLetterInstance) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".UpdateLetter"
    
    If Me.BatchID <> 0 Then
        oLetter.BatchID = Me.BatchID
    End If
    
    If cdctLetters.Exists(oLetter.InstanceID) = True Then
        Set cdctLetters.Item(oLetter.InstanceID) = oLetter
    Else
        cdctLetters.Add oLetter.InstanceID, oLetter
    End If
    
    If cdctLetterPageCnt.Exists(oLetter.InstanceID) = True Then
        cdctLetterPageCnt.Item(oLetter.InstanceID) = oLetter.PageCount
    Else
        cdctLetterPageCnt.Add oLetter.InstanceID, oLetter.PageCount
    End If
    
    UpdateLetter = True
Block_Exit:
    Exit Function
Block_Err:
    FireError Err, strProcName
    GoTo Block_Exit
End Function


Public Function AddLetterInstance(sInstanceId As String, sQueueStatus As String, sProvNum As String, ByVal sLetterType As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oLtr As clsLetterInstance

    strProcName = ClassName & ".AddLetterInstance"
    
    If cdctLetters.Exists(sInstanceId) = False Then
        Set oLtr = New clsLetterInstance
        With oLtr
            .InstanceID = sInstanceId
            .LetterQueueStatus = sQueueStatus
            .ProvNum = sProvNum
            .LetterType = UCase(Trim(sLetterType))
            .LetterCreateDt = Now()
            .LetterReqDt = Now()    ' not really sure what this is all about.. KD COMEBACK: may need to change this
                                    ' and use whatever is in the database
            If Me.BatchID <> 0 Then
                .BatchID = Me.BatchID
            End If
        End With
        
        ' eventually we'll probably ahve QR codes for all letters..
        oLtr.LoadQRCodePath
        
        cdctLetters.Add sInstanceId, oLtr
        ' if it does exist, then silently do nothing
    Else
        Stop
    End If
    
    AddLetterInstance = True
    
Block_Exit:
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

    cbErrorOccurred = True
    
    ReportError oErr, sErrSourceProcName, , sAdditionalDetails
    
    If sAdditionalDetails <> "" Then sAdditionalDetails = vbCrLf & sAdditionalDetails
    
    RaiseEvent LetterError(oErr.Description & sAdditionalDetails, oErr.Number, sErrSourceProcName)

End Sub


Private Sub Class_Initialize()
    Set cdctLetters = New Scripting.Dictionary
    Set cdctLetterPageCnt = New Scripting.Dictionary ' this will hold the Key: InstanceId, Value = Page count of the physical letter
    Me.BatchID = 0
    
'    Set coSourceTable = New clsTable
'    coSourceTable.IdFieldName = csIDFIELDNAME
'    coSourceTable.TableName = csTableName
    
'    Set coReqRule = Nothing
'    Set ccolAttachedDocs = New Collection
'    Set ccolHdrAttached = New Collection
'    Set ccolDtlAttached = New Collection
'
'    Set ccolTaggedClaims = New Collection
'    Set ccolPayerDetails = New Collection
'    Set coValidateRpt = New clsEracValidationRpt
'
'    Set coValidateRpt = Nothing
'
'    Set cdctInitObjs = New Scripting.Dictionary
'    With cdctInitObjs
'        .Add "GetReqRule", False
'        .Add "LoadPayerDetails", False
'        .Add "LoadAttachedDocs", False
'        .Add "LoadTaggedClaims", False
'    End With
'
'    Call GetAggFieldNamesDict
        
    cblnIsInitialized = False
End Sub

Private Sub Class_Terminate()
    If Dirty = True Then
'        SaveNow
    End If
    Set cdctLetters = Nothing
'    Set coSourceTable = Nothing
'    Set cdctInitObjs = Nothing
'    Set ccolAttachedDocs = Nothing
'    Set ccolHdrAttached = Nothing
'    Set ccolDtlAttached = Nothing
'
'    Set ccolTaggedClaims = Nothing
'    Set coValidateRpt = Nothing
    cblnIsInitialized = False
End Sub