Option Compare Database
Option Explicit


Private Const ClassName As String = "mod_ADO"


Public Function AdoExeTxt(sqlStr As String, Optional TableName As String, Optional TimeoutSeconds As Long = 600, Optional ByVal Server As String, Optional Database As String) As Boolean    ' As Long

'** 06/27/2012 KD: added some functions to copy stuff from an ado rs to a local table

'**  6/14/05 JAC Cleaned up original procedure. Removed unused variables, clarified untyped variable
'**              declarations, added destruction of db and tdf object variables on exit,  added return value of records affected, and removed error handler which gave an
'**              incorrect message referring to an xml object.
'**
'** Uses ADO to execute a SQL string in the database which contains the linked table, ReferenceTable.  It seems that all objects must be in the same database for this to work.
'**
'**  5/19/06 JAC Removed the db & tdf objects.   The tdf.connect is now replaced with a string variable.
'**                 Changed LocConn and LocCmd from variant to Object. Added return parameter and cleaned up other incoherence.
'**
'** 6/19/06 JAC Added Timeout parameter
'** 5/17/07 TOM C Added the optional parameter server to allow calling a server different than the currently linked one.
'** 6/15/07 TOM C Added the optional parameter database to allow calling a database different than the currently linked one.
'** 6/15/07 TOM C Made the table name parameter optional.
'** 6/15/07 TOM C NOTE: The user must supply either: 1) TableName or 2) Server and Database


    On Error GoTo ErrorHappened

    Dim LocConn As Object    'As ADODB.Connection
    Dim LocCmd As Object    'As ADODB.Command
            
    Dim StConnect As String
    Dim strDb As String
    Dim intStartPos As Integer
    Dim intLen As Integer
    Dim bReturn As Boolean
    Dim strTableConnect As String

    bReturn = True

    Set LocConn = CreateObject("ADODB.Connection")
    Set LocCmd = CreateObject("ADODB.Command")
    
    If TableName = "" And (Server = "" Or Database = "") Then
        MsgBox "You must enter either a table or a server and database.", vbCritical, "ExeSql ERROR"
        bReturn = False
        GoTo ExitNow
    End If
   
    If TableName <> "" Then
        strTableConnect = CurrentDb.TableDefs(TableName).Connect
        If Right(strTableConnect, 1) <> ";" Then strTableConnect = strTableConnect & ";"
    End If

    '5/17/07 TOM C Added so that you can call a server different than the currently linked one.
    If Server = "" Then
        'get server from connectstring in workfile linked table
        intStartPos = InStr(strTableConnect, "SERVER=") + 7
        intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
        Server = Mid(strTableConnect, intStartPos, intLen)
    End If

    If Database = "" Then
        'get server and database name from connectstring in workfile linked table
        intStartPos = InStr(strTableConnect, "DATABASE=") + 9
        intLen = InStr(intStartPos, strTableConnect, ";") - intStartPos
        strDb = Mid(strTableConnect, intStartPos, intLen)
    End If

    'Set and open Connection and Command Objects
    StConnect = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                "Data Source=" & Server & ";" & _
                StConnect & "Initial Catalog=" & strDb & ";"

    LocConn.Open StConnect

    With LocCmd

        .ActiveConnection = LocConn
        .CommandTimeout = TimeoutSeconds
        .commandType = 1    '* adCmdText
        .CommandText = sqlStr
        .Execute  ', , 16    'adAsyncExecute

        '* JC 9/29/06 Removed this because most events in the app depend on the execution having finished.

        '        Do Until .State <> 4  'adStateExecuting
        '
        '            DoEvents
        '
        '        Loop

    End With

ExitNow:
    On Error Resume Next

    LocConn.Close

    Set LocCmd = Nothing
    Set LocConn = Nothing

    AdoExeTxt = bReturn

    Exit Function

ErrorHappened:

    MsgBox Err.Number & " (" & Err.Description & ")" & vbCr & "String: " & sqlStr & vbCr & _
           "Table: " & TableName

    bReturn = False
    Resume ExitNow
    Resume

End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Returns true if that field is found in the
''' given recordset..
Public Function isField(ByRef rs As ADODB.RecordSet, strFieldNameToCheck As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFld As ADODB.Field

    strProcName = ClassName & ".isField"
    
    If rs Is Nothing Then GoTo Block_Exit
    
    For Each oFld In rs.Fields
        If LCase(oFld.Name) = LCase(strFieldNameToCheck) Then
            isField = True
            Exit For
        End If
    Next
    
Block_Exit:
    Set oFld = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function



''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Returns true if that paramater is found in the
''' given cmd Object
Public Function isParameter(ByRef oCmd As ADODB.Command, ByVal strParamNameToCheck As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oParam As ADODB.Parameter


    strProcName = ClassName & ".isParameter"
    strParamNameToCheck = UCase(strParamNameToCheck)
    
    If oCmd Is Nothing Then GoTo Block_Exit
    
    For Each oParam In oCmd.Parameters
        If UCase(oParam.Name) = strParamNameToCheck Then
            isParameter = True
            Exit For
        End If
    Next
    
Block_Exit:
    Set oParam = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' This method will retry a query (which must already be set up in the clsAd being passed)
''' for a particular result.. i.e. if you call a batch file or something that runs asynchronously
''' and you need the results (which will show up in this query) to be something in particular
''' then you can call this setting desiredvalue and fieldname.. By default it'll time out
''' after 10 tries, but you can change that
Public Function RetryQueryXTimes(ByRef oAdo As clsADO, ByRef rs As ADODB.RecordSet, strDesiredValue As String, _
       Optional strFieldName As String = "ConceptID", Optional blnNegate As Boolean = False, _
       Optional iXTimesBeforeTimeout As Integer = 10) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim iTimeoutCount As Integer

    strProcName = ClassName & ".RetryQueryXTimes"

    '' KD: COMEBACK: Put some code in here to make sure that oAdo is ready to go
    '' Or, just throw an exception to the caller...
    
    If oAdo.sqlString = "" Or oAdo.ConnectionString = "" Then
        '' Do we thrown an exception? not sure how the rest of the code works so
        '' I won't do this now..
        LogMessage strProcName, "WARNING", "ADO object not properly set up before calling function!"
        GoTo Block_Exit
    End If
    
        '' Get the results..
    Set rs = oAdo.OpenRecordSet
    
        '' If we don't have any records, try it again, but not forever...
    Do While ((rs.recordCount < 1) And (iTimeoutCount < iXTimesBeforeTimeout))
       ''' Let's try to wait a couple seconds then try it again
       iTimeoutCount = iTimeoutCount + 1
       Sleep 2000
       Set rs = oAdo.OpenRecordSet
    Loop
    
        
        ' Try to find our record, if we didn't find it, sleep and then try it again
        ' but time out after iXTimesBeforeTimeout times..
    Do While ((FindARecordInRs(rs, strDesiredValue, strFieldName, blnNegate) = False) And (iTimeoutCount < iXTimesBeforeTimeout))
        iTimeoutCount = iTimeoutCount + 1
        Sleep 2000
        Set rs = oAdo.OpenRecordSet
    Loop
    
        '' Did we find it or time out trying?
    If (FindARecordInRs(rs, strDesiredValue, strFieldName, blnNegate) = False) Then
        ' if we didn't find it then we've timed out - no need to check that..
        RetryQueryXTimes = False
        If rs.recordCount > 0 Then rs.MoveFirst
        GoTo Block_Exit
    End If
    

    RetryQueryXTimes = True

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    RetryQueryXTimes = False
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function




''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' returns true if a given fieldname in the passed recordset contains the desired value
''' Note that because this is byref, if found (one or more times) the current row in the
''' recordset will be the active row (the first one found that is..)
Public Function FindARecordInRs(ByRef rs As ADODB.RecordSet, strDesiredValue As String, Optional strFieldName As String = "ConceptID", _
    Optional blnNegate As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".FindARecordInRs"
  
    If rs Is Nothing Then GoTo Block_Exit
        
    ' Make sure that the fieldname is actually a field:
    If isField(rs, strFieldName) = False Then
        FindARecordInRs = False
        GoTo Block_Exit
    End If
    
  
        ' Find our record..
    rs.MoveFirst
    
    Do While Not rs.EOF
        If blnNegate = True Then
            If rs(strFieldName) <> strDesiredValue Then
                FindARecordInRs = True
                GoTo Block_Exit
                Exit Do
            End If
        Else
            If rs(strFieldName) = strDesiredValue Then
                FindARecordInRs = True
                GoTo Block_Exit
                Exit Do
            End If
        
        End If
        rs.MoveNext
    Loop
    
        '' Move it back to the top since we didn't find it
    rs.MoveFirst
    
    FindARecordInRs = False

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit     ' I don't like resume's because we could have a 'Resume without error' exception
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''

''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function RSHasData(oRs As ADODB.RecordSet) As Boolean
    If oRs Is Nothing Then GoTo Block_Exit
    If oRs.recordCount < 1 Then GoTo Block_Exit
    RSHasData = True
Block_Exit:
    Exit Function
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function AdoTypeToDaoType(objFld As ADODB.Field) As Long
    Select Case objFld.Type
    Case adVarChar
        AdoTypeToDaoType = dbText
    Case adInteger
        'AdoTypeToDaoType = dbInteger
        AdoTypeToDaoType = dbDouble
    Case adSmallInt, adUnsignedTinyInt
        AdoTypeToDaoType = dbInteger
'    Case adDBTimeStamp
'        AdoTypeToDaoType = dbText
    Case adBoolean
        AdoTypeToDaoType = dbBoolean
    Case adVarBinary
        AdoTypeToDaoType = dbBinary
    Case adNumeric
        AdoTypeToDaoType = dbDouble
    Case adCurrency
        AdoTypeToDaoType = dbCurrency
    Case adSingle
        AdoTypeToDaoType = dbDouble
    Case adDouble
        AdoTypeToDaoType = dbDouble
    Case adVarWChar
        AdoTypeToDaoType = dbText
    Case adDate, adDBTimeStamp, adDBTime, adDBDate, adDBTimeStamp
        AdoTypeToDaoType = dbDate
    Case Else
        AdoTypeToDaoType = dbText
    End Select
End Function


''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
Public Function DaoTypeToAdoType(objFld As DAO.Field) As Long
    Select Case objFld.Type
    Case dbText
        DaoTypeToAdoType = adVarChar
    Case dbBoolean
        DaoTypeToAdoType = adBoolean
    Case dbBigInt
        DaoTypeToAdoType = adInteger
    Case dbByte
        DaoTypeToAdoType = adBoolean
    Case dbChar
        DaoTypeToAdoType = adVarChar
    Case dbCurrency
        DaoTypeToAdoType = adCurrency
    Case dbDate
        DaoTypeToAdoType = adDBTimeStamp
    Case dbDecimal
        DaoTypeToAdoType = adNumeric
    Case dbDouble
        DaoTypeToAdoType = adDouble
    Case dbFloat
        DaoTypeToAdoType = adDouble
    Case dbInteger
        DaoTypeToAdoType = adInteger
    Case dbLong
        DaoTypeToAdoType = adDouble
    Case dbLongBinary
        DaoTypeToAdoType = adVarBinary
    Case dbMemo
        DaoTypeToAdoType = adVarWChar
    Case dbNumeric
        DaoTypeToAdoType = adNumeric
    Case dbSingle
        DaoTypeToAdoType = adSingle
    Case dbText
        DaoTypeToAdoType = adVarChar
    Case dbTime
        DaoTypeToAdoType = adDBTimeStamp
    Case dbTimeStamp
        DaoTypeToAdoType = adDBTimeStamp
    Case dbVarBinary
        DaoTypeToAdoType = adVarBinary
    Case Else
        DaoTypeToAdoType = adVarChar
    End Select
End Function



Public Function CopyDataToLocalTmpTable(oRs As ADODB.RecordSet, bForceRemake As Boolean, Optional sTmpTableName As String = "tmp_Local_Copy") As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oFld As ADODB.Field
Dim oDaoRs As DAO.RecordSet


    strProcName = ClassName & ".CopyDataToLocalTmpTable"
    
        ' at some point, I may need to make this bit smarter but only 1 thing is using it now so...
    
    If IsTable(sTmpTableName) = False Or bForceRemake = True Then
        Call CreateTableFromADORS(oRs, sTmpTableName, bForceRemake)
    End If
    
    ' Make sure it's empty
    CurrentDb.Execute "DELETE FROM [" & sTmpTableName & "]"
    
    Set oDaoRs = CurrentDb.OpenRecordSet(sTmpTableName, dbOpenTable)
    
    ' populate it:
    If oRs.BOF And oRs.EOF Then
'        Stop
            oDaoRs.AddNew
            oDaoRs.Update


    Else
        oRs.MoveFirst
        While Not oRs.EOF
            oDaoRs.AddNew
            For Each oFld In oRs.Fields
                oDaoRs(oFld.Name) = oFld.Value
            Next
            oDaoRs.Update
            oRs.MoveNext
        Wend

    End If
    
    
    
Block_Exit:
    Set oDaoRs = Nothing
    Set oFld = Nothing
    CopyDataToLocalTmpTable = sTmpTableName
    Exit Function
Block_Err:
    ReportError Err, strProcName
    sTmpTableName = ""
    GoTo Block_Exit
End Function



Public Function CreateTableFromADORS(oRs As ADODB.RecordSet, sTblName As String, Optional bForceRemake As Boolean = False) As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oTDef As DAO.TableDef
Dim oAdoField As ADODB.Field
Dim oTblFld As DAO.Field

    strProcName = ClassName & ".CreateTableFromADORS"
    
    If bForceRemake = True Then
        If IsTable(sTblName) = True Then
            CurrentDb.TableDefs.Delete (sTblName)
            CurrentDb.TableDefs.Refresh
        End If
    ElseIf IsTable(sTblName) = True Then
            ' already created. nothing to do
        CreateTableFromADORS = sTblName
        GoTo Block_Exit
    End If
    
    Set oTDef = New DAO.TableDef
    With oTDef
        .Name = sTblName
        For Each oAdoField In oRs.Fields
            Set oTblFld = New DAO.Field
            oTblFld.Name = oAdoField.Name
            oTblFld.Type = AdoTypeToDaoType(oAdoField)
            .Fields.Append oTblFld
        Next

    End With
    
    
    CurrentDb.TableDefs.Append oTDef
    CreateTableFromADORS = sTblName
    CurrentDb.TableDefs.Refresh
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Public Function GetDSNLessODBCConnString(sServerName As String, sDbName As String) As String
           
    GetDSNLessODBCConnString = "ODBC;Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & sDbName & ";" & _
           "Trusted_Connection=yes;"
End Function






''' ##############################################################################
''' ##############################################################################
''' ##############################################################################
'''
''' Returns true if that field is found in the
''' given recordset..
Public Function isDAOField(ByRef rs As DAO.RecordSet, strFieldNameToCheck As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFld As DAO.Field

    strProcName = ClassName & ".isDAOField"
    
    If rs Is Nothing Then GoTo Block_Exit
    
    For Each oFld In rs.Fields
        If LCase(oFld.Name) = LCase(strFieldNameToCheck) Then
            isDAOField = True
            Exit For
        End If
    Next
    
Block_Exit:
    Set oFld = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    Err.Clear
    GoTo Block_Exit
End Function


Public Function QuoteIfNeeded(sValue As String) As String
On Error GoTo Block_Err
Dim strProcName As String
Static oRegEx As RegExp

    strProcName = ClassName & ".QuoteIfNeeded"

    If oRegEx Is Nothing Then
        Set oRegEx = New RegExp
        With oRegEx
            .IgnoreCase = True
            .Global = False
            .MultiLine = False
            .Pattern = "^[0-9\.]$"
        End With
    End If

    If oRegEx.test(sValue) = True Then
        QuoteIfNeeded = sValue  ' = numeric
    Else
        QuoteIfNeeded = Replace(sValue, "'", "''")
        QuoteIfNeeded = "'" & QuoteIfNeeded & "'"
    End If

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function