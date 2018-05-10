Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database



''' Last Modified: 04/30/2015
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 04/30/2015 : KD: extended to optionally use DAO instead of just ADO
'''  - 12/14/2012 : Fixed LogMessage calls..
'''  - Fixed SaveNow stuff
'''  - 03/14/2012 - Added Fields() property
'''     - Added ability to have a non integer ID
'''  - 03/07/2012 - Created...
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


Private ciId As Integer
Private csId As String
Private cdctTableVals As Scripting.Dictionary
Private cblnDirtyData As Boolean
Private cblnIsInitialized As Boolean
Private cstrTableName As String
Private cstrTableNameForConnection As String
Private cstrIDFieldName As String
Private cstrRecordNameFieldName As String
Private cintHowManyRecords As Integer
Private ccolFields As Collection
Private cblnStringId As Boolean
Private cblnDao As Boolean
Private cbLoadSingleRow As Boolean


Public Property Get Id() As Integer
'    ciId = cdctTableVals.Item(IdFieldName).Value
    Id = ciId
End Property
Public Property Let Id(intNewId As Integer)
    ciId = intNewId
End Property



Public Property Get IDStr() As String
'    ciId = cdctTableVals.Item(IdFieldName).Value
    IDStr = csId
End Property
Public Property Let IDStr(strNewId As String)
    csId = strNewId
End Property



Public Property Get HowManyRecords() As Integer
    HowManyRecords = cintHowManyRecords
End Property
Public Property Let HowManyRecords(intNumOfRecords As Integer)
    cintHowManyRecords = intNumOfRecords
End Property

Public Property Get IdFieldName() As String
    IdFieldName = cstrIDFieldName
End Property
Public Property Let IdFieldName(strIdFieldName As String)
    cstrIDFieldName = strIdFieldName
End Property


Public Property Get RecordNameFieldName() As String
    RecordNameFieldName = cstrRecordNameFieldName
End Property
Public Property Let RecordNameFieldName(strFieldName As String)
    cstrRecordNameFieldName = strFieldName
End Property

'
Public Property Get LoadSingleRow() As Boolean
    LoadSingleRow = cbLoadSingleRow
End Property
Public Property Let LoadSingleRow(bLoadSingleRow As Boolean)
    cbLoadSingleRow = bLoadSingleRow
End Property

Public Property Get IdIsString() As Boolean
    IdIsString = cblnStringId
End Property
Public Property Let IdIsString(blnStringId As Boolean)
    cblnStringId = blnStringId
End Property


Public Property Get TableName() As String
    TableName = cstrTableName
End Property
Public Property Let TableName(strTableName As String)
    cstrTableName = strTableName
End Property


Public Property Get TableNameForConnection() As String
    If cstrTableNameForConnection <> "" Then
        TableNameForConnection = cstrTableNameForConnection
    Else
        TableNameForConnection = cstrTableName
    End If
End Property
Public Property Let TableNameForConnection(strTableNameForConnection As String)
    cstrTableNameForConnection = strTableNameForConnection
End Property


Public Property Get IsDAO() As Boolean
    IsDAO = cblnDao
End Property
Public Property Let IsDAO(blnDAO As Boolean)
    cblnDao = blnDAO
End Property

Public Property Get WasInitialized() As Boolean
    WasInitialized = cblnIsInitialized
End Property
Public Property Let WasInitialized(blnWasInit As Boolean)
    cblnIsInitialized = blnWasInit
End Property

Public Property Get Dirty() As Boolean
    Dirty = cblnDirtyData
End Property
Public Property Let Dirty(blnDirtyData As Boolean)
    cblnDirtyData = blnDirtyData
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Function Fields() As Collection
    Set Fields = ccolFields
End Function


Public Function InitializeFromRS(oRs As Object, Optional blnLoadOneRecordForSchema As Boolean = False) As Boolean
On Error GoTo Funct_Err
Dim strProcName As String


    strProcName = ClassName & ".InitializeFromRS"
    If oRs Is Nothing Then
        GoTo Funct_Exit
    End If
    
    If TypeOf oRs Is DAO.RecordSet Then
        InitializeFromRS = InitializeFromRSDAO(oRs, blnLoadOneRecordForSchema)
    ElseIf TypeOf oRs Is ADODB.RecordSet Then
        InitializeFromRS = InitializeFromRSADO(oRs, blnLoadOneRecordForSchema)
    End If
    

FinishedNow:
Funct_Exit:
    Exit Function
Funct_Err:
    InitializeFromRS = False
    ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
    Resume Funct_Exit
End Function

    
    Private Function InitializeFromRSADO(oRs As ADODB.RecordSet, Optional blnLoadOneRecordForSchema As Boolean = False) As Boolean
    On Error GoTo Funct_Err
    Dim strProcName As String
    Dim iFieldCount As Integer
    Dim sFieldName As String
    Dim iRecordCount As Integer
    Dim varBkMark As Variant
    
        strProcName = ClassName & ".InitializeFromRSADO"
        If oRs Is Nothing Then
            GoTo Funct_Exit
        End If
        
        If oRs.EOF And oRs.BOF Then
            GoTo Funct_Exit
        End If
        InitializeFromRSADO = True
        
        If cdctTableVals Is Nothing Then Set cdctTableVals = New Scripting.Dictionary
    
        'If oRs.Bookmarkable = True Then
            varBkMark = oRs.Bookmark
        'End If
        
        
        While Not oRs.EOF
            iRecordCount = iRecordCount + 1
            For iFieldCount = 0 To oRs.Fields.Count - 1
                sFieldName = oRs.Fields(iFieldCount).Name
                If UCase(sFieldName) = "CAUSECRITICALERRORIFFAILED" Then
                    cdctTableVals.Add CStr(iRecordCount) & "." & "CRITICALIMPORT", oRs(sFieldName).Value
                Else
                    cdctTableVals.Add CStr(iRecordCount) & "." & UCase(sFieldName), oRs(sFieldName).Value
                End If
                        ''' Debug.Print CStr(iRecordCount) & "." & UCase(sFieldName) & " = (" & oRs(sFieldName).Value & ")"
                If UCase(sFieldName) = UCase(IdFieldName) Then
                    If IdIsString = True Then
                        IDStr = oRs(sFieldName).Value
                    Else
                        Id = oRs(sFieldName).Value
                    End If
                End If
            Next
            If blnLoadOneRecordForSchema = True Then
                GoTo FinishedNow
            End If
            oRs.MoveNext
        Wend
        oRs.MoveFirst
        
    
        HowManyRecords = iRecordCount
    
FinishedNow:
        oRs.MoveFirst
        ' Now, move back to the record we started at:
        'If oRs.Bookmarkable = True Then
            oRs.Bookmark = varBkMark
        'End If
        HowManyRecords = iRecordCount
        WasInitialized = True
Funct_Exit:
        Exit Function
Funct_Err:
        InitializeFromRSADO = False
        ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
        Resume Funct_Exit
    End Function
    
    
    '  LEGACY DAO code
    Private Function InitializeFromRSDAO(oRs As DAO.RecordSet, Optional blnLoadOneRecordForSchema As Boolean = False) As Boolean
    On Error GoTo Funct_Err
    Dim strProcName As String
    Dim iFieldCount As Integer
    Dim sFieldName As String
    Dim iRecordCount As Integer
    Dim varBkMark As Variant
    
        strProcName = ClassName & ".InitializeFromRSDAO"
        If oRs.EOF And oRs.BOF Then
            GoTo Funct_Err
        End If
        InitializeFromRSDAO = True
    
        If cdctTableVals Is Nothing Then Set cdctTableVals = New Scripting.Dictionary
    
        If oRs.Bookmarkable = True Then
            varBkMark = oRs.Bookmark
        End If
    
        With cdctTableVals
            While Not oRs.EOF
                iRecordCount = iRecordCount + 1
                For iFieldCount = 0 To oRs.Fields.Count - 1
                    sFieldName = oRs.Fields(iFieldCount).Name
                    If UCase(sFieldName) = "CAUSECRITICALERRORIFFAILED" Then
                        .Add CStr(iRecordCount) & "." & "CRITICALIMPORT", oRs(sFieldName).Value
                    Else
                        .Add CStr(iRecordCount) & "." & UCase(sFieldName), oRs(sFieldName).Value
                    End If
        Debug.Print CStr(iRecordCount) & "." & UCase(sFieldName) & " = (" & oRs(sFieldName).Value & ")"
                    If UCase(sFieldName) = UCase(IdFieldName) Then
                        Id = oRs(sFieldName).Value
                    End If
                Next
                If blnLoadOneRecordForSchema = True Then
                    GoTo FinishedNow
                End If
                oRs.MoveNext
            Wend
            oRs.MoveFirst
    
    
        End With

    
FinishedNow:
        oRs.MoveFirst
        ' Now, move back to the record we started at:
        If oRs.Bookmarkable = True Then
            oRs.Bookmark = varBkMark
        End If
        HowManyRecords = iRecordCount
        WasInitialized = True
Funct_Exit:
        Exit Function
Funct_Err:
        InitializeFromRSDAO = False
        ReportError Err, strProcName, "SourceID: " & CStr(ciId)
        Resume Funct_Exit
    End Function

' Note: not designed for multiple values (ie, one record in the recordset..
Public Function GetTableValue(strFieldName As String) As Variant
On Error GoTo Funct_Err
Dim strProcName As String
    
    strProcName = ClassName & ".GetTableValue"

    If IdIsString = True Then
        If IDStr = "" Then
            Set GetTableValue = Nothing
            GoTo Funct_Exit
        End If
    Else
        If Id < 1 Then  ' not initialized
            Set GetTableValue = Nothing
            ' or,
            GetTableValue = ""
            GoTo Funct_Exit
        End If
    End If


    
    strFieldName = UCase(strFieldName)
    
    If cdctTableVals.Exists("1." & strFieldName) = True Then
        GetTableValue = cdctTableVals.Item("1." & strFieldName)
    End If
    
    
Funct_Exit:
    Exit Function
Funct_Err:
    
    ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
    Resume Funct_Exit
End Function

'
Public Function SetTableValue(strFieldName As String, varValue As Variant, Optional blnAllRecords As Boolean = True, Optional blnSaveNow As Boolean = False) As Boolean
On Error GoTo Funct_Err
Dim strProcName As String
Dim iRecordLoop As Integer
Dim sFldName As String

    strProcName = ClassName & ".SetTableValue"

    If IdIsString = True Then
        If IDStr = "" Then GoTo Funct_Exit
    Else
        If Id < 1 Then GoTo Funct_Exit
    End If

    SetTableValue = True

    
    strFieldName = UCase(strFieldName)
    
        ' how do I deal with multiple rows here...  cstr(iRecordCount) & "." &


    For iRecordLoop = 1 To HowManyRecords
        If cdctTableVals.Exists(CStr(iRecordLoop) & "." & strFieldName) = True Then
            cdctTableVals.Item(CStr(iRecordLoop) & "." & strFieldName) = varValue
            Dirty = True
        Else
            SetTableValue = False
            LogMessage strProcName, "ERROR", "Field name for object does not exist", CStr(iRecordLoop) & "." & strFieldName
''''            Debug.Assert 1 <> 1
            Debug.Print "Huh"

        End If
    
        If blnAllRecords = False Then
            ' We're done.. exit loop.
            GoTo DoneNow
        End If
    Next
DoneNow:
    If blnSaveNow = True Then
        If SaveNow = False Then
            Dirty = True
        Else
            Dirty = False
        End If
    End If
    
    
Funct_Exit:
    Exit Function
Funct_Err:
    SetTableValue = False
    ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
    Resume Funct_Exit
End Function



' SaveNow
Public Function SaveNow(Optional blnAllRecords As Boolean = True) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".SaveNow"
    SaveNow = True
    
    If Dirty = False Or WasInitialized = False Then
        SaveNow = False
        GoTo Block_Exit
    End If


    If IsDAO = True Then
        SaveNow = SaveNowDAO(blnAllRecords)
    Else
        SaveNow = SaveNowADO(blnAllRecords)
    End If

DoneUpdating:
    Dirty = False

    SaveNow = True
Block_Exit:

    Exit Function

Block_Err:
    SaveNow = False
    ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
    GoTo Block_Exit
End Function
    
    ' SaveNow
    Private Function SaveNowADO(Optional blnAllRecords As Boolean = True) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim sSql As String
    Dim oAdo As clsADO
    Dim oCn As ADODB.Connection
    
    Dim oRs As ADODB.RecordSet
    Dim sFieldName As String
    Dim iFieldLoop As Integer
    Dim asPrimaryKeyFields() As String
    Dim iRecordCount As Integer
    
        strProcName = ClassName & ".SaveNowADO"
        SaveNowADO = True
        
        If Dirty = False Or WasInitialized = False Then
            SaveNowADO = False
            GoTo Block_Exit
        End If
    
        If IdIsString = True Then
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & IDStr & "'"
        Else
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(Id)
        End If
    
    
    
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString(Me.TableNameForConnection)
            .SQLTextType = sqltext
            .sqlString = sSql
        End With
        Set oCn = New ADODB.Connection
        Set oRs = New ADODB.RecordSet
        
        oRs.Open sSql, oAdo.CurrentConnection, adOpenDynamic, adLockOptimistic
        
        If oRs.EOF And oRs.BOF Then
            ' not found
            SaveNowADO = False
            GoTo Block_Exit
        End If
        
        oRs.MoveFirst
        While Not oRs.EOF
            iRecordCount = iRecordCount + 1
    '        oRs.Edit
            LogMessage strProcName, , "Saving record " & CStr(iRecordCount)
            For iFieldLoop = 0 To oRs.Fields.Count - 1
                sFieldName = UCase(oRs.Fields(iFieldLoop).Name)
    ''''Debug.Assert sFieldName <> "ACTIVE"
                If sFieldName <> UCase(IdFieldName) And IsEditableField(TableName, sFieldName) = True Then
                    If IsAuditFieldName(sFieldName) = True Then
                        ' global function to set the various audit field values..
                        oRs(sFieldName).Value = AuditFieldValue(sFieldName)
                    Else
                        If cdctTableVals.Exists(CStr(iRecordCount) & "." & sFieldName) = True Then
                                    ''' Debug.Print "Setting: " & sFieldName & " (" & CStr("" & oRs(sFieldName).Value) & ") to " & cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                            oRs(sFieldName).Value = cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                        End If
                    End If
                End If
            Next
            oRs.Update
            If blnAllRecords = False Then
                ' we're done - exit loop
                GoTo DoneUpdating
            End If
            oRs.MoveNext
        Wend
    '    oRs.Update
        oRs.UpdateBatch
        
        LogMessage strProcName, , "Updated " & CStr(iRecordCount) & " records"
DoneUpdating:
        Dirty = False
    
        SaveNowADO = True
Block_Exit:
        Set oRs = Nothing
        Set oAdo = Nothing
        Exit Function
    
Block_Err:
        SaveNowADO = False
        ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
        GoTo Block_Exit
    End Function
    
    '  :LEGACY DAO code
    '' SaveNow
    Private Function SaveNowDAO(Optional blnAllRecords As Boolean = True) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim sSql As String
    Dim oDb As DAO.Database
    Dim oRs As DAO.RecordSet
    Dim sFieldName As String
    Dim iFieldLoop As Integer
    Dim asPrimaryKeyFields() As String
    Dim iRecordCount As Integer
    
        strProcName = ClassName & ".SaveNowDAO"
        SaveNowDAO = True
        
        If Dirty = False Or WasInitialized = False Then
            SaveNowDAO = False
            GoTo Block_Exit
        End If
    
        If IdIsString = True Then
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & IDStr & "'"
        Else
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(Id)
        End If
    
        Set oDb = CurrentDb()
        
        Set oRs = oDb.OpenRecordSet(sSql)
        
        If oRs.EOF And oRs.BOF Then
            ' not found
            SaveNowDAO = False
            GoTo Block_Exit
        End If
        
        oRs.MoveFirst
        While Not oRs.EOF
            iRecordCount = iRecordCount + 1
            oRs.Edit
            LogMessage strProcName, , "Saving record " & CStr(iRecordCount)
            For iFieldLoop = 0 To oRs.Fields.Count - 1
                sFieldName = UCase(oRs.Fields(iFieldLoop).Name)
    ''''Debug.Assert sFieldName <> "ACTIVE"
'                oRs.Edit
                If sFieldName <> UCase(IdFieldName) And IsEditableField(TableName, sFieldName) = True Then
                    If IsAuditFieldName(sFieldName) = True Then
                        ' global function to set the various audit field values..
                        oRs(sFieldName).Value = AuditFieldValue(sFieldName)
                    Else
                        If cdctTableVals.Exists(CStr(iRecordCount) & "." & sFieldName) = True Then
                                    ''' Debug.Print "Setting: " & sFieldName & " (" & CStr("" & oRs(sFieldName).Value) & ") to " & cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                            oRs(sFieldName).Value = cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                        End If
                    End If
                End If
            Next
            oRs.Update
            If blnAllRecords = False Then
                ' we're done - exit loop
                GoTo DoneUpdating
            End If
            oRs.MoveNext
        Wend
'        oRs.Update
    '    oRs.UpdateBatch
        oRs.Close
        
        LogMessage strProcName, , "Updated " & CStr(iRecordCount) & " records"
DoneUpdating:
        Dirty = False
    
        SaveNowDAO = True
Block_Exit:
        Set oRs = Nothing
        Set oDb = Nothing
        Exit Function
    
Block_Err:
        SaveNowDAO = False
        ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
        GoTo Block_Exit
    End Function


Public Function AddNewID(Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Variant
On Error GoTo Funct_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String
Dim intID As Integer
Dim iFieldLoop As Integer
Dim sFieldName As String
Dim iRecordCount As Integer

    strProcName = ClassName & ".AddNewID"

    If IsDAO = True Then
        AddNewID = AddNewIDDAO(strTableName, strIdFieldName, strRecordNameFieldName)
    Else
        AddNewID = AddNewIDADO(strTableName, strIdFieldName, strRecordNameFieldName)
    End If
    
Funct_Exit:
    Exit Function
Funct_Err:
    AddNewID = False
    ReportError Err, strProcName, "SourceID: " & CStr(intID)
    Resume Funct_Exit
End Function
    
    
    Private Function AddNewIDDAO(Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Variant
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oDb As DAO.Database
    Dim oRs As DAO.RecordSet
    Dim sSql As String
    Dim intID As Integer
    Dim iFieldLoop As Integer
    Dim sFieldName As String
    Dim iRecordCount As Integer
    
        strProcName = ClassName & ".AddNewIDDAO"
    
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
        
        If IsTable(TableName) = False Then
            AddNewIDDAO = False
            GoTo Block_Exit
        End If
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
    
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        intID = 0
        csId = ""
        iRecordCount = 1
    
        If IdIsString = True Then
            IDStr = ""
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & IDStr & "'"
        Else
            Id = 0
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(intID)
        End If
    
        
        Set oDb = CurrentDb
        Set oRs = oDb.OpenRecordSet(sSql)
        
        oRs.AddNew
    
        For iFieldLoop = 0 To oRs.Fields.Count - 1
            sFieldName = UCase(oRs.Fields(iFieldLoop).Name)
    
            If sFieldName <> UCase(IdFieldName) And IsEditableField(TableName, sFieldName) = True Then
                If IsAuditFieldName(sFieldName) = True Then
                    ' global function to set the various audit field values..
                    oRs(sFieldName).Value = AuditFieldValue(sFieldName)
                Else
                    If sFieldName = UCase(RecordNameFieldName) Then
                        ' what do we need to do here?
                        Stop
                    ElseIf cdctTableVals.Exists(CStr(iRecordCount) & "." & sFieldName) = True Then
                                '''     Debug.Print "Setting: " & sFieldName & " (" & CStr("" & oRs(sFieldName).Value) & ") to " & cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                        oRs(sFieldName).Value = cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
'                    Else
'                        cdctTableVals.Add CStr(iRecordCount) & "." & sFieldName, Me.Get
'                        oRs(sFieldName).Value = cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                    End If
                End If
            End If
        Next
    
        oRs.Update
        
        oRs.MoveLast
        If IdIsString = True Then
            IDStr = oRs(IdFieldName).Value
        Else
            Id = oRs(IdFieldName).Value
        End If
        AddNewIDDAO = oRs(IdFieldName).Value
        
        oRs.Close
        
Block_Exit:
        Set oRs = Nothing
        Set oDb = Nothing
        Exit Function
Block_Err:
        AddNewIDDAO = False
        ReportError Err, strProcName, "SourceID: " & CStr(intID)
        GoTo Block_Exit
    End Function



    Private Function AddNewIDADO(Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Variant
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim sSql As String
    Dim intID As Integer
    Dim iFieldLoop As Integer
    Dim sFieldName As String
    Dim iRecordCount As Integer
    
        strProcName = ClassName & ".AddNewIDADO"
    
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
        
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
    
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        intID = 0
        csId = ""
        iRecordCount = 1
    
        If IdIsString = True Then
    Stop    ' wtf - how are we using the recordset before we open it?
            IDStr = ""  '   oRs(sFieldName).Value
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & IDStr & "'"
        Else
    Stop    ' wtf - how are we using the recordset before we open it?
            Id = 0      ''oRs(sFieldName).Value
            sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(intID)
        End If
    
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString(Me.TableNameForConnection)
            .SQLTextType = sqltext
            .sqlString = sSql
            Set oRs = .ExecuteRS
        End With
        
        oRs.AddNew
        
        For iFieldLoop = 0 To oRs.Fields.Count - 1
            sFieldName = UCase(oRs.Fields(iFieldLoop).Name)
    
            If sFieldName <> UCase(IdFieldName) And IsEditableField(TableName, sFieldName) = True Then
                If IsAuditFieldName(sFieldName) = True Then
                    ' global function to set the various audit field values..
                    oRs(sFieldName).Value = AuditFieldValue(sFieldName)
                Else
                    If sFieldName = UCase(RecordNameFieldName) Then
                        ' what do we need to do here?
                        Stop
                    ElseIf cdctTableVals.Exists(CStr(iRecordCount) & "." & sFieldName) = True Then
                                '''     Debug.Print "Setting: " & sFieldName & " (" & CStr("" & oRs(sFieldName).Value) & ") to " & cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                        oRs(sFieldName).Value = cdctTableVals(CStr(iRecordCount) & "." & sFieldName)
                    End If
                End If
            End If
        Next
    
        oRs.Update
        
        oRs.MoveLast
        If IdIsString = True Then
            IDStr = oRs(IdFieldName).Value
        Else
            Id = oRs(IdFieldName).Value
        End If
        AddNewIDADO = oRs(IdFieldName).Value
        
    
Block_Exit:
        If Not oRs Is Nothing Then
            If oRs.State = adStateOpen Then oRs.Close
            Set oRs = Nothing
        End If
        Set oAdo = Nothing
        Exit Function
Block_Err:
        AddNewIDADO = False
        ReportError Err, strProcName, "SourceID: " & CStr(intID)
        GoTo Block_Exit
    End Function



Public Function LoadWholeTable(Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String, Optional blnLoadOneRecordForSchema As Boolean = False) As Boolean
On Error GoTo Funct_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String
    
    strProcName = ClassName & ".LoadWholeTable"
    
    If Len(strTableName) > 0 Then
        TableName = strTableName
    End If
    
    If IsTable(TableName) = False Then
        LoadWholeTable = False
        GoTo Funct_Exit
    End If
    
    If Len(strIdFieldName) > 0 Then
        IdFieldName = strIdFieldName
    End If
    
    If Len(strRecordNameFieldName) > 0 Then
        RecordNameFieldName = strRecordNameFieldName
    End If

    sSql = "SELECT * FROM [" & TableName & "] "
    
    Set oDb = CurrentDb
    Set oRs = oDb.OpenRecordSet(sSql)
    If oRs.EOF And oRs.BOF Then
        LoadWholeTable = False
        GoTo Funct_Exit
    End If
    
    LoadWholeTable = InitializeFromRS(oRs, blnLoadOneRecordForSchema)

Funct_Exit:
    Exit Function
Funct_Err:
    LoadWholeTable = False
    ReportError Err, strProcName
    GoTo Funct_Exit
End Function



Public Function LoadFromId(lngId As Long, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oAdo As clsADO
Dim oField As ADODB.Field

    strProcName = ClassName & ".LoadFromID"

    If Len(strTableName) > 0 Then
        TableName = strTableName
    End If
    
    If Me.IsDAO = True Then
        LoadFromId = LoadFromIDDAO(lngId, strTableName, strIdFieldName, strRecordNameFieldName, cbLoadSingleRow)
    Else
        LoadFromId = LoadFromIDADO(lngId, strTableName, strIdFieldName, strRecordNameFieldName, cbLoadSingleRow)
    End If

Block_Exit:
    Exit Function
Block_Err:
    LoadFromId = False
    ReportError Err, strProcName, "SourceID: " & CStr(lngId)
    GoTo Block_Exit
End Function

    
    Private Function LoadFromIDDAO(lngId As Long, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String, Optional bLoadSingleRow As Boolean = True) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oDb As DAO.Database
    Dim oRs As DAO.RecordSet
    Dim sSql As String
    Dim oField As DAO.Field
    
        strProcName = ClassName & ".LoadFromIDDAO"
        
            ' it's got to be a table in this database
        If strTableName <> "" Then
            TableName = strTableName
        End If
        
        If IsTable(TableName) = False Then
            LoadFromIDDAO = False
            LogMessage strProcName, "ERROR", "Table to load from does not exist!", TableName
            GoTo Block_Exit
        End If
        
        If Len(Trim(strIdFieldName)) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        Id = lngId
    
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(lngId)
        
        Set oDb = CurrentDb
        Set oRs = oDb.OpenRecordSet(sSql)
        
        If oRs.EOF And oRs.BOF Then
            LoadFromIDDAO = False
            GoTo Block_Exit
        End If
        
        LoadFromIDDAO = InitializeFromRSDAO(oRs, bLoadSingleRow)
        
        For Each oField In oRs.Fields
            ccolFields.Add CStr("" & oField.Name)
        Next
        
        HowManyRecords = 1
Block_Exit:
        Exit Function
Block_Err:
        ReportError Err, strProcName
        GoTo Block_Exit
    End Function
    
    
    Private Function LoadFromIDADO(lngId As Long, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String, Optional bLoadSingleRow As Boolean = True) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oRs As ADODB.RecordSet
    Dim sSql As String
    Dim oAdo As clsADO
    Dim oField As ADODB.Field
    
        strProcName = ClassName & ".LoadFromIDADO"
    
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
        
    '    If IsTable(TableName) = False Then
    '        LoadFromIDADO = False
    '        GoTo Block_Exit
    '    End If
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        Id = lngId
    
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(lngId)
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString(Me.TableNameForConnection)
            .SQLTextType = sqltext
            .sqlString = sSql
        End With
        
            '    Set oRs = oDB.OpenRecordSet(sSql)
        Set oRs = oAdo.ExecuteRS
        If oRs.EOF And oRs.BOF Then
            LoadFromIDADO = False
            GoTo Block_Exit
        End If
        
        LoadFromIDADO = InitializeFromRSADO(oRs, bLoadSingleRow)
        
        For Each oField In oRs.Fields
            ccolFields.Add CStr("" & oField.Name)
        Next
        
        HowManyRecords = 1
    
Block_Exit:
        If Not oRs Is Nothing Then
            If oRs.State = adStateOpen Then oRs.Close
            Set oRs = Nothing
        End If
        Set oAdo = Nothing
        Exit Function
Block_Err:
        LoadFromIDADO = False
        ReportError Err, strProcName, "SourceID: " & CStr(lngId)
        GoTo Block_Exit
    End Function


Public Function LoadFromIDStr(strID As String, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".LoadFromIDStr"

    If IsDAO = True Then
        LoadFromIDStr = LoadFromIDStrDAO(strID, strTableName, strIdFieldName, strRecordNameFieldName)
    Else
        LoadFromIDStr = LoadFromIDStrADO(strID, strTableName, strIdFieldName, strRecordNameFieldName)
    End If

Block_Exit:

    Exit Function
Block_Err:
    LoadFromIDStr = False
    ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
    GoTo Block_Exit
End Function

    
    Private Function LoadFromIDStrADO(strID As String, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oRs As ADODB.RecordSet
    Dim sSql As String
    Dim oAdo As clsADO
    Dim oField As ADODB.Field
    
        strProcName = ClassName & ".LoadFromIDStrADO"
    
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
    
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        IDStr = strID
    
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & strID & "'"
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString(Me.TableNameForConnection)
            .SQLTextType = sqltext
            .sqlString = sSql
        End With
        
        Set oRs = oAdo.ExecuteRS
        If oRs.EOF And oRs.BOF Then
            LoadFromIDStrADO = False
            GoTo Block_Exit
        End If
        
        LoadFromIDStrADO = InitializeFromRSADO(oRs)
        
        For Each oField In oRs.Fields
            ccolFields.Add CStr("" & oField.Name)
        Next
        
        HowManyRecords = 1
    
Block_Exit:
        If Not oRs Is Nothing Then
            If oRs.State = adStateOpen Then oRs.Close
            Set oRs = Nothing
        End If
        Set oAdo = Nothing
        Exit Function
Block_Err:
        LoadFromIDStrADO = False
        ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
        GoTo Block_Exit
    End Function
    
    
    Private Function LoadFromIDStrDAO(strID As String, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oRs As DAO.RecordSet
    Dim oDb As DAO.Database
    Dim sSql As String
    Dim oField As DAO.Field
    
        strProcName = ClassName & ".LoadFromIDStrDAO"
    
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
    
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
    
        IDStr = strID
    
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = '" & strID & "'"
        
        Set oDb = CurrentDb()
        Set oRs = oDb.OpenRecordSet(sSql)
       
        If oRs.EOF And oRs.BOF Then
            LoadFromIDStrDAO = False
            GoTo Block_Exit
        End If
        
        LoadFromIDStrDAO = InitializeFromRSDAO(oRs)
        
        For Each oField In oRs.Fields
            ccolFields.Add CStr("" & oField.Name)
        Next
        
        HowManyRecords = 1
    
Block_Exit:
        If Not oRs Is Nothing Then
            Set oRs = Nothing
        End If
        Set oDb = Nothing
        Exit Function
Block_Err:
        LoadFromIDStrDAO = False
        ReportError Err, strProcName, "SourceID: " & IIf(cblnStringId, csId, CStr(ciId))
        GoTo Block_Exit
    End Function




Public Function IsAuditFieldName(strFieldName As String) As Boolean

    If gdctAuditFieldNames Is Nothing Then
        Set gdctAuditFieldNames = New Scripting.Dictionary
        With gdctAuditFieldNames
            .Add "LASTMODIFIED", "now"
            .Add "LASTMODIFIEDDATE", "now"
            .Add "LASTMODIFIEDTIME", "now"
            .Add "MODIFYUSER", Environ("username")
            .Add "MODIFYGetPCName", GetPCName
            .Add "MODIFYCOMPUTER", GetPCName
            .Add "COMPUTER", GetPCName
            .Add "GetPCName", GetPCName
            .Add "WORKSTATIONNAME", GetPCName
            .Add "INACTIVEDATE", "now"
        End With
    End If
    
    IsAuditFieldName = gdctAuditFieldNames.Exists(UCase(strFieldName))

End Function

Public Function AuditFieldValue(strFieldName As String) As Variant
On Error GoTo Funct_Err
Dim strProcName As String

    strProcName = ClassName & ".AuditFieldValue"

    If IsAuditFieldName(strFieldName) = False Then
        GoTo Funct_Exit
    End If

    Select Case LCase(gdctAuditFieldNames(strFieldName))
    Case "now"
        AuditFieldValue = Now()
    Case Else
        AuditFieldValue = gdctAuditFieldNames(strFieldName)
    End Select

Funct_Exit:
    Exit Function
Funct_Err:
    ReportError Err, strProcName, strFieldName
    Resume Funct_Exit
End Function


Private Sub Class_Initialize()
    Set cdctTableVals = New Scripting.Dictionary
    Set ccolFields = New Collection
    cbLoadSingleRow = True
End Sub

Private Sub Class_Terminate()
    If Dirty = True Then
        Stop
    End If
    Set cdctTableVals = Nothing
    Set ccolFields = Nothing
End Sub


Public Function DeleteID(intIDToDelete As Integer, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".DeleteID"
    DeleteID = True
    
    If Len(strTableName) > 0 Then
        TableName = strTableName
    End If
    
    If IsDAO = True Then
        DeleteID = DeleteIDDAO(intIDToDelete, strTableName, strIdFieldName, strRecordNameFieldName)
    Else
        DeleteID = DeleteIDADO(intIDToDelete, strTableName, strIdFieldName, strRecordNameFieldName)
    End If


Block_Exit:
    Exit Function

Block_Err:
    DeleteID = False
    ReportError Err, strProcName, "Id to delete: " & CStr(intIDToDelete)
    GoTo Block_Exit
End Function
    
    
    Private Function DeleteIDADO(intIDToDelete As Integer, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oAdo As clsADO
    Dim oRs As ADODB.RecordSet
    Dim sSql As String
    
        strProcName = ClassName & ".DeleteIDADO"
        DeleteIDADO = True
        
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
        
        If IsTable(TableName) = False Then
            DeleteIDADO = False
            GoTo Block_Exit
        End If
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
        
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(intIDToDelete)
        
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString(Me.TableNameForConnection)
            .SQLTextType = sqltext
            .sqlString = sSql
        End With
        
            'Set oRs = oDB.OpenRecordSet(sSql)
        Set oRs = oAdo.ExecuteRS
        If oRs.EOF And oRs.BOF Then
            DeleteIDADO = False
            GoTo Block_Exit
        End If
        
        oRs.Delete
    
Block_Exit:
        Set oRs = Nothing
        Set oAdo = Nothing
        Exit Function
    
Block_Err:
        DeleteIDADO = False
        ReportError Err, strProcName, "Id to delete: " & CStr(intIDToDelete)
        GoTo Block_Exit
    End Function
    
    
    Private Function DeleteIDDAO(intIDToDelete As Integer, Optional strTableName As String, Optional strIdFieldName As String, Optional strRecordNameFieldName As String) As Boolean
    On Error GoTo Block_Err
    Dim strProcName As String
    Dim oDb As DAO.Database
    Dim oRs As DAO.RecordSet
    Dim sSql As String
    
        strProcName = ClassName & ".DeleteIDDAO"
        DeleteIDDAO = True
        
        If Len(strTableName) > 0 Then
            TableName = strTableName
        End If
        
        If IsTable(TableName) = False Then
            DeleteIDDAO = False
            GoTo Block_Exit
        End If
        
        If Len(strIdFieldName) > 0 Then
            IdFieldName = strIdFieldName
        End If
        
        If Len(strRecordNameFieldName) > 0 Then
            RecordNameFieldName = strRecordNameFieldName
        End If
        
        sSql = "SELECT * FROM [" & TableName & "] WHERE [" & IdFieldName & "] = " & CStr(intIDToDelete)
        
        Set oDb = CurrentDb()
        
        Set oRs = oDb.OpenRecordSet(sSql)
    
        If oRs.EOF And oRs.BOF Then
            DeleteIDDAO = False
            GoTo Block_Exit
        End If
        
        oRs.Delete
'        oRs.Update
    
Block_Exit:
        Set oRs = Nothing
        Set oDb = Nothing
        Exit Function
    
Block_Err:
        DeleteIDDAO = False
        ReportError Err, strProcName, "Id to delete: " & CStr(intIDToDelete)
        GoTo Block_Exit
    End Function




'Public Function DefaultFunction(oRs As DAO.RecordSet) As Boolean
'On Error GoTo Funct_Err
'Dim strProcName As String
'
'    strProcName = ClassName & ".DefaultFunction"
'
'    DefaultFunction = False
'Funct_Exit:
'    Exit Function
'Funct_Err:
'    DefaultFunction = False
'    ReportError Err, strProcName, "SourceID: " & CStr(ciId)
'    Resume Funct_Exit
'End Function