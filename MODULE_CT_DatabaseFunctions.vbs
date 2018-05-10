Option Compare Database
Option Explicit

'SA 11/26/2012 - Added GetServerNameFromConnectionString and GetDatabaseNameFromConnectionString

'Moved from CnlyDtFunctions
Public Function GetLinkedServer(strTableName As String) As String
'Retrieves the server name for the specified linked table
' HC 5/2010 cleaned up unused varaibles
On Error GoTo GetLinkedServer_Error
        
    Dim Server As String
    Dim intStartPos As Integer
    Dim intLen As Integer
    Dim bReturn As Boolean
    Dim sLocInd As String
    
    bReturn = True
    sLocInd = strTableName
    
    'get server and database name from connectstring in workfile linked table ActiveAudit
    intStartPos = InStr(CurrentDb.TableDefs(strTableName).Connect, "SERVER=") + 7
    intLen = InStr(intStartPos, CurrentDb.TableDefs(strTableName).Connect, ";") - intStartPos
    Server = Mid(CurrentDb.TableDefs(strTableName).Connect, intStartPos, intLen)


    GetLinkedServer = Server
GetLinkedServer_Exit:
    On Error Resume Next
    Exit Function
    
GetLinkedServer_Error:
    'MsgBox err.Description
    Resume GetLinkedServer_Exit
End Function

'Moved from CnlyDtFunctions
Public Function GetSQLServer(ByVal strTableName As String) As String
    'SA 03/22/2012 - Deprecated function for backward compatibility
    GetSQLServer = GetLinkedServer(strTableName)
End Function

'Moved from CnlyDtFunctions
Public Function GetLinkedDatabase(strTableName As String) As String
'Retrieves the database name for the specified linked table
    On Error GoTo GetLinkedDatabase_Error
    
    ' HC 5/2010 cleaned up unused variables
    Dim Database As String
    Dim intStartPos As Integer
    Dim intLen As Integer
        
    'get server and database name from connectstring in workfile linked table ActiveAudit
    intStartPos = InStr(CurrentDb.TableDefs(strTableName).Connect, "DATABASE=") + 9
    intLen = InStr(intStartPos, CurrentDb.TableDefs(strTableName).Connect, ";") - intStartPos
    If intLen <= 0 Then
        Database = Mid(CurrentDb.TableDefs(strTableName).Connect, intStartPos)
    Else
        Database = Mid(CurrentDb.TableDefs(strTableName).Connect, intStartPos, intLen)
    End If

    GetLinkedDatabase = Database
GetLinkedDatabase_Exit:
    On Error Resume Next
    Exit Function
    
GetLinkedDatabase_Error:
    'MsgBox err.Description
    Resume GetLinkedDatabase_Exit
    Resume
End Function

'Moved from CnlyDtFunctions
Public Function GetSQLDatabase(ByVal strTableName As String) As String
    'SA 03/22/2012 - Deprecated function for backward compatibility moved from CnlyProjectsFunctions
    GetSQLDatabase = GetLinkedDatabase(strTableName)
End Function

'Moved from CnlyDtFunctions
Public Function ExeSqlDb(sqlStr As String, dbName As String, Optional rst As Object) As Boolean
   
    'Dbname is sql db where the statement will be executed.
    
    On Error GoTo ErrorHappened
    
    Dim StConnect As String
    Dim LocConn 'As ADODB.Connection
    Dim LocCmd 'As ADODB.Command
    Dim Server
    Dim Database

    ' hlp050505
    Dim bReturn As Boolean

    screen.MousePointer = 11
    ' hlp050505
    bReturn = True

    Server = GetLinkedServerDb(dbName)
    Database = dbName
     
'Set and open Connection and Command Objects
    ' HC 5/2010 used the defined constant
    StConnect = LINK_SRC_SQL & "Persist Security Info=False;"
    StConnect = StConnect & "Data Source=" & Server & ";"
    StConnect = StConnect & "Initial Catalog=" & Database & ";"

    Set LocConn = CreateObject("ADODB.Connection")
    LocConn.Open StConnect
    Set LocCmd = CreateObject("ADODB.Command")
    LocCmd.ActiveConnection = LocConn
    LocCmd.CommandTimeout = 300 '5 min
    
    LocCmd.commandType = 1
    LocCmd.CommandText = sqlStr
    
    'debug.print "ExeSQL.Sql: " & SQLStr
    
    If rst Is Nothing Then
        ' HC 5/2010 -- modified this command to execute without returning any records since the rst is nothing
        LocCmd.Execute , , &H90  'H84 is adAsyncExecute (x10) Or'd with adExecuteNoRecords (x80)'16 adAsyncExecute
        
        'we're exectuting async., wait for for the command to execute.
        Do Until LocCmd.State <> 4  'adStateExecuting
            DoEvents
        Loop
        
    Else
        rst.CursorLocation = 3 'adUseClient.  Is neccessary for creating a disconnected recordset
        rst.CursorType = adOpenStatic
        rst.LockType = adLockBatchOptimistic
        rst.Open sqlStr, LocConn, adOpenStatic, adLockBatchOptimistic
        'Set Rst = LocCmd.Execute(, , 16)
    End If

    If Not rst Is Nothing Then
        'need to dissconnect the recordset before closing the connection.
        Set rst.ActiveConnection = Nothing
    End If
ExitNow:
    On Error Resume Next
    screen.MousePointer = 0
    LocConn.Close
    Set LocCmd = Nothing
    Set LocConn = Nothing
    ExeSqlDb = bReturn
    Exit Function
    
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Executing SQL on Linked Table"
    bReturn = False
    Resume ExitNow
    Resume
End Function

'Moved from CnlyDtFunctions
Public Function GetLinkedServerDb(ByVal dbName As String) As String
'Retrieves the server name for the specified linked table
    On Error GoTo GetLinkedServerdb_Error
    ' hc 5/2010 cleaned up unused variables
    Dim Server As String
    Dim i As Integer
    Dim intStartPos As Integer
    Dim intLen As Integer
    Dim bReturn As Boolean
    
    bReturn = True

    'get server and database name from connectstring with matching database name
    For i = 0 To CurrentDb.TableDefs.Count - 1
        'SA 2012-07-16 - Added trailing ; to connect string and dbname to prevent matches on partial db names
        If InStr(1, CurrentDb.TableDefs(i).Connect & ";", "DATABASE=" & Trim(dbName) & ";") > 0 Then
            intStartPos = InStr(CurrentDb.TableDefs(i).Connect, "SERVER=") + 7
            intLen = InStr(intStartPos, CurrentDb.TableDefs(i).Connect, ";") - intStartPos
            Server = Mid(CurrentDb.TableDefs(i).Connect, intStartPos, intLen)

            Exit For
        End If
        
    Next i

    GetLinkedServerDb = Server
GetLinkedServerdb_Exit:
    On Error Resume Next
    Exit Function
    
GetLinkedServerdb_Error:
    MsgBox Err.Description
    Resume GetLinkedServerdb_Exit
End Function

'Moved from CnlyDtFunctions
Public Function ExeSqlSp(sqlStr As String, TblName As String, Optional strControl As String) As String
    'strControl is the name of a control who's caption should show the #of elapsed seconds of execution
    On Error GoTo ErrorHappened
    Dim StConnect As String
    Dim LocConn 'As ADODB.Connection
    Dim ErrMsg As String
    Dim Server As String
    Dim Database As String
    Dim AdoErr 'As ADODB.Error
        
    screen.MousePointer = 11

'get server and database name from connectstring in workfile linked table
    Server = GetLinkedServer(TblName)
    Database = GetLinkedDatabase(TblName)

'Set and open Connection and Command Objects
    ' HC 5/2010 use the defined constant
    StConnect = LINK_SRC_SQL & "Persist Security Info=False;"
    StConnect = StConnect & "Data Source=" & Server & ";"
    StConnect = StConnect & "Initial Catalog=" & Database & ";"

    Set LocConn = CreateObject("ADODB.Connection")
    LocConn.Open StConnect
    Do Until LocConn.State <> 0  'adStateClosed
        DoEvents
    Loop

    On Error Resume Next
    DoCmd.Hourglass True
    
    LocConn.Execute sqlStr, , &H90  'H84 is adAsyncExecute (x10) Or'd with adExecuteNoRecords (x80)
    Do Until LocConn.State = 1  'adStateExecuting
        DoEvents
    Loop
        
    DoCmd.Hourglass False
    'get message resulting from parse (in errors collection).
    ErrMsg = ""
    For Each AdoErr In LocConn.Errors
        ErrMsg = ErrMsg & AdoErr.Description & vbCrLf
    Next
    
ExitNow:
    On Error Resume Next
    screen.MousePointer = 0
    LocConn.Close
    Do Until LocConn.State = 0  'adStateClosed
        DoEvents
    Loop

    'Set LocCmd = Nothing
    Set LocConn = Nothing
    ExeSqlSp = ErrMsg
    Exit Function
    
ErrorHappened:
    For Each AdoErr In LocConn.Errors
        MsgBox AdoErr.Description
    Next
    
    If Err.Number <> 0 Then MsgBox Err.Description
    Resume ExitNow
End Function

'Moved from CnlyDtFunctions
Public Function ExDLookup(db As DAO.Database, Expr As String, Domain As String, Optional Criteria As String = "", Optional OrderClause As String)
    'Mimic the funtionality of DLookup() but also take a target Database as an argument.
    On Error GoTo eTrap
    
    Dim SQL As String
    Dim strSQL As String
    Dim rst As RecordSet
    
    SQL = "SELECT TOP 1 " & Expr & " FROM " & Domain & IIf(Criteria = "", "", " WHERE " & Criteria)
    
    If Not IsMissing(OrderClause) And OrderClause <> "" Then
        SQL = SQL & " ORDER BY " & OrderClause
    End If
    
    strSQL = SQL & ";"

    Set rst = db.OpenRecordSet(SQL, dbOpenForwardOnly, dbReadOnly)
    
    If Not (rst.EOF And rst.BOF) Then
        ExDLookup = rst.Fields(0)
    Else
        ExDLookup = Null
    End If
    
eSuccess:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Function
    
eTrap:
    MsgBox Err.Description, vbCritical, Err.Source
    
    Resume eSuccess
    Resume
    
End Function

Public Function CreateInsertSQL(ByVal TableName As String, ByRef rst As Object) As String
'Create insert SQL for screen import
'SA 11/10/2012 - Changed to use GetForeignKeysDAO instead of GetForeignKeys and fixed up exception handling
On Error GoTo ErrorHappened
    Dim SQL As String
    Dim FieldList As String
    Dim InsString As String
    Dim ValString As String
    Dim ColForeignKeys As Collection
    Dim ColFkey
    Dim i As Integer

    FieldList = GetFieldList(TableName, 0, 1)
    InsString = "INSERT INTO " & TableName & "(" & FieldList & ")"
    Set ColForeignKeys = GetForeignKeysDAO(TableName)
    
    If ColForeignKeys Is Nothing Then
        CreateInsertSQL = "Error"
        Exit Function
    End If
    
    SQL = vbNullString

    With rst
        If Not (.EOF And .BOF) Then
            .MoveFirst
            Do Until .EOF
                ValString = InsString & "VALUES("
                For i = 0 To .Fields.Count - 1
                    'If the field is a FKey, get the updated value from the parent table.
                    'Otherwise use the value in the .
                    If InCollection(.Fields(i).Name, ColForeignKeys) Then 'is Fkey
                        Set ColFkey = ColForeignKeys(.Fields(i).Name)
                        ValString = ValString & DLookup(.Fields(i).Name, ColFkey("RelatedTable"), "RefID = " & .Fields(i).Value) & ","
            
                    Else    'Not an Fkey
                        If IsNumeric(.Fields(i).Value) Then
                            ValString = ValString & .Fields(i).Value & " ,"
                        Else
                            ' HC 11/14/2008 - added to set the value of the user
                            If .Fields(i).Value = "Identity.UserName" Then
                                ValString = ValString & "'" & Identity.UserName & "' ,"
                            Else
                                ValString = ValString & "'" & Replace(.Fields(i).Value & "", "'", "''") & "' ,"
                            End If
                        End If
                    End If 'Field is FKey
                Next
    
                SQL = SQL & left$(ValString, Len(ValString) - 1) & ")" & "<%EOL%>"
                .MoveNext
            Loop
        
            Set ColForeignKeys = Nothing
            
            CreateInsertSQL = SQL
        End If
    End With

ExitNow:
On Error Resume Next

Exit Function
ErrorHappened:
    MsgBox Err.Description, vbCritical, "CT_DatabaseFunctions:CreateInsertSQL"
    Resume ExitNow
    Resume
End Function

'Moved from CnlyDtFunctions
Public Function GetFieldList(strTable As String, Optional fldCount As Integer = 0, Optional fldFirst As Integer = 0) As String
'fldCount specifies the first N fields to parse (1 based). 0 = all.

    On Error GoTo eCatch
    
    Dim db As DAO.Database
    Dim TbDef As TableDef
    Dim strFields As String
    Dim i As Integer
    
    'SQL = "SELECT * FROM " & strTable & " WHERE 1=0"
    Set db = CurrentDb
    'Set Rst = DB.OpenRecordset(SQL, dbOpenSnapshot)
    Set TbDef = db.TableDefs(strTable)
    With TbDef
    
        If fldCount = 0 Then
            fldCount = .Fields.Count - fldFirst
        End If
        
        For i = fldFirst To (fldCount - 1) + fldFirst
            strFields = strFields & "[" & .Fields(i).Name & "]"
            
            If i < (fldCount - 1) + fldFirst Then
                strFields = strFields & ", "
            End If
        Next i
        
    End With
    
    GetFieldList = strFields

eCatch:
    On Error Resume Next
    Set TbDef = Nothing
    Set db = Nothing
    Exit Function
    Resume


End Function

'Moved from CnlyDtFunctions
Public Function GetForeignKeys(TableName As String, Optional dbName As String) As Collection
    'Given a table name, returns a collection populated with
    'all foreign keys and their corresponding tables, with the foreign key column name as the key.
    On Error GoTo eTrap
    
    Dim cat 'As New ADOX.Catalog
    Dim KeyLoop
    
    Set cat = CreateObject("ADOX.Catalog")

    ' Connect the catalog
    If dbName = "" Then
    
        dbName = CurrentDb.Name
        cat.ActiveConnection = CurrentProject.Connection
    Else
        ' HC 5/2010 - changed to use the variable for access
        cat.ActiveConnection = LINK_SRC_ACCESS & "Data Source= " & dbName & ";"
    End If
        
    Dim ForeignKeys As New Collection
    Dim ForeignKeysInfo As Collection

    On Error GoTo eTrap
 
    'Debug.Print "Foreign Keys for table: " & tableName
    'Debug.Print cat.Tables(TableName).Keys.count
    
    For Each KeyLoop In cat.Tables(TableName).Keys
        If KeyLoop.Type = 2 Then 'Foreign
        
            'We only want this foreign key if its related table has a field called "RefID"
            If InStr(1, GetFieldList(KeyLoop.RelatedTable), "RefID") Then
            
                'Debug.Print vbTab & "ForeignTable: " & KeyLoop.RelatedTable, "Keyfield: " & KeyLoop.Columns(0).name
                Set ForeignKeysInfo = New Collection
                ForeignKeysInfo.Add Item:=KeyLoop.Columns(0).Name, Key:="Field"
                ForeignKeysInfo.Add Item:=KeyLoop.RelatedTable, Key:="RelatedTable"
                ForeignKeys.Add Item:=ForeignKeysInfo, Key:=KeyLoop.Columns(0).Name
                
            End If
            
            'Exit For
        End If
        
    Next
    
    Set GetForeignKeys = ForeignKeys
    
eSuccess:
    On Error Resume Next

    Set cat = Nothing
    Exit Function
    
eTrap:
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "GetForeignKeys()"
    Set GetForeignKeys = Nothing

    GoTo eSuccess
    Resume
End Function

Public Function GetForeignKeysDAO(ByVal TableName As String, Optional DatabaseName As String = vbNullString) As Collection
'This function is meant to replace GetForeignKeys which uses ADOX
'ADOX errors when there are any linked Access tables. DAO appears to work better.
On Error GoTo ErrorHappened
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim db As DAO.Database
    Dim ForeignKeys As New Collection
    Dim ForeignKeysInfo As Collection

    If LenB(DatabaseName) = 0 Then
        Set db = CurrentDb
    Else
        Set db = DBEngine.OpenDatabase(DatabaseName)
    End If

    For Each rel In db.Relations
        If TableName = rel.ForeignTable Then
            For Each fld In rel.Fields
                If InStr(1, GetFieldList(rel.Table), "RefID") Then
                    Set ForeignKeysInfo = New Collection
                    ForeignKeysInfo.Add fld.Name, "Field"
                    ForeignKeysInfo.Add rel.Table, "RelatedTable"
                    ForeignKeys.Add ForeignKeysInfo, fld.Name
                End If
            Next
        End If
    Next
    
    #If ccSCR = 1 Then
        Set ForeignKeys = SCR_ScreensImportKeysOverride(TableName, ForeignKeys)
    #ElseIf ccSCR = 2 Then
        Set ForeignKeys = SCR_AppScreensImportKeysOverride(TableName, ForeignKeys)
    #End If
    
    Set GetForeignKeysDAO = ForeignKeys
ExitNow:
On Error Resume Next
    db.Close
    Set db = Nothing
Exit Function
ErrorHappened:
    Set GetForeignKeysDAO = Nothing
    Resume ExitNow
    Resume
End Function

'Moved from CnlyDtFunctions
Public Function TableExists(TableName As String, db As Database) As Boolean
    Dim strTemp As String
    On Error Resume Next
    
    'Try to reference the specified table in the TableDefs Collection.
    strTemp = db.TableDefs(TableName).Name
    TableExists = IIf(Err.Number = 0, True, False)
    
End Function

'Moved from CnlyDtFunctions
Public Function GetPrimKeyName(SrcDb As DAO.Database, TableName As String) As String
    'Given a table name, returns the first field name of the primary key.
    'Returns "" if no primary key for specified table.

    Dim IdxLoop As index
    Dim PrimKey As String

    On Error GoTo eTrap

    For Each IdxLoop In SrcDb.TableDefs(TableName).Indexes
        If IdxLoop.Primary Then
            PrimKey = IdxLoop.Fields(0).Name
            Exit For
        End If
    Next
       
eTrap:
    GetPrimKeyName = PrimKey
End Function

'Moved from CnlyDtFunctions
Public Function BuildStringFromQuery(strQry As String, Optional SourceDB As Database) As String
    'given a query against a native table, will return a comma delimited string of the resulting records.
    'Note, will only return the first field of n fields in the result set.
    On Error GoTo eCatch
    
    Dim db As DAO.Database
    Dim rst As DAO.RecordSet
    Dim strDelimit As String
    
    If SourceDB Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = SourceDB
    End If
    
    Set rst = db.OpenRecordSet(strQry, dbOpenSnapshot)
    With rst
        If .EOF And .BOF Then
            
        Else
        

            Select Case .Fields(0).Type
                
                Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbLongBinary
                    strDelimit = ""
                Case Else
                    strDelimit = "'"
            End Select
        
            Do Until .EOF
                BuildStringFromQuery = BuildStringFromQuery & strDelimit & .Fields(0) & strDelimit & ","
                'SubformCalcs.Form
                .MoveNext
            Loop
            .Close
            'trim the last comma
            BuildStringFromQuery = left(BuildStringFromQuery, Len(BuildStringFromQuery) - 1)

        End If
    End With
    
eCatch:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
End Function

'Moved from CnlyDtFunctions
Public Function ExeSql(ByVal sqlStr As String, ByVal TblName As String, Optional ByRef rst As Object, Optional ByRef returnErrorMessage As String = "NotPassedIn") As Boolean
' DLC 06/26/11 - added ByRef/ByVal to parameters
' DPS 05/26/11 - added optional parameter to to ExeSQL to return error message instead of popuping up a MsgBox
    'table name is a linked table with a connection to the sql db where the statement will be executed.
    'TblName parameter only used to get the Server and Database name.
    On Error GoTo ErrorHappened
    
    Dim StConnect As String
    Dim LocConn 'As ADODB.Connection
    Dim LocCmd 'As ADODB.Command
    Dim Server As String
    Dim Database As String

    ' hlp050505
    Dim bReturn As Boolean

    screen.MousePointer = 11
    ' hlp050505
    bReturn = True

'get server and database name from connectstring in workfile linked table
                        
    Server = GetLinkedServer(TblName)
    Database = GetLinkedDatabase(TblName)
    
'Set and open Connection and Command Objects
    ' HC 5/2010 replaced with constants
    StConnect = LINK_SRC_SQL & "Persist Security Info=False;"
    StConnect = StConnect & "Data Source=" & Server & ";"
    StConnect = StConnect & "Initial Catalog=" & Database & ";"

    Set LocConn = CreateObject("ADODB.Connection")
    LocConn.Open StConnect
    Set LocCmd = CreateObject("ADODB.Command")
    LocCmd.ActiveConnection = LocConn
    LocCmd.CommandTimeout = 300 '5 min
    
    LocCmd.commandType = 1
    LocCmd.CommandText = sqlStr
    
    If rst Is Nothing Then
        LocCmd.Execute , , 16 'adAsyncExecute
        'we're exectuting async., wait for for the command to execute.
        Do Until LocCmd.State <> 4  'adStateExecuting
            DoEvents
        Loop
    Else
        rst.CursorLocation = 3 'adUseClient.  Is neccessary for creating a disconnected recordset
        rst.CursorType = adOpenStatic
        rst.LockType = adLockBatchOptimistic
        rst.Open sqlStr, LocConn, adOpenStatic, adLockBatchOptimistic
        'Set Rst = LocCmd.Execute(, , 16)
    End If

    If Not rst Is Nothing Then
        'need to dissconnect the recordset before closing the connection.
        Set rst.ActiveConnection = Nothing
    End If
ExitNow:
    On Error Resume Next
    screen.MousePointer = 0
    LocConn.Close
    Set LocCmd = Nothing
    Set LocConn = Nothing
    ExeSql = bReturn
    Exit Function
    
ErrorHappened:
    If returnErrorMessage = "NotPassedIn" Then
        MsgBox Err.Description, vbCritical, "Executing SQL on Linked Table"
    Else
        returnErrorMessage = Err.Description
    End If
    bReturn = False
    Resume ExitNow
    Resume
End Function

'Moved from CnlyDtFunctions
Public Function JetSQLFixup(TextIn)
Dim temp
On Error GoTo eTrap
  temp = Replace(TextIn, "'", "''")
  JetSQLFixup = Replace(temp, "|", "' & chr(124) & '")
  
eTrap:
Exit Function
Resume
End Function

Public Function CT_BackupDatabase(Optional ByVal NewFileName As String = vbNullString) As String
'Create a backup copy of the current database with a timestamp
'Returns the name of the backup file if successful or null if there is an error
'SA 10/23/2012 - Added function. (Need to use FSO for a file to make a copy of itself)
On Error GoTo ErrorHappened
    Dim Result As String
    Dim fso As Object
    Dim OldFileName As String
    
    OldFileName = CurrentDb.Name
    If LenB(NewFileName) = 0 Then
        NewFileName = Replace(OldFileName & "_Backup_" & Format(Now, "yyyymmddhhmmss"), ".accdb", vbNullString) & ".accdb"
    End If
    
    'Add current path if only a file name is specified
    If InStr(NewFileName, "\") = 0 Then
        NewFileName = CurrentProject.Path & "\" & NewFileName
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile OldFileName, NewFileName

    Result = NewFileName
ExitNow:
On Error Resume Next
    Set fso = Nothing
    CT_BackupDatabase = Result
Exit Function
ErrorHappened:
    Result = vbNullString
    Resume ExitNow
End Function

Public Function GetServerNameFromConnectionString(ByVal ConString As String) As String
'Return the server name from SQL connection string
'ex: ODBC;DRIVER=SQL Server;SERVER=YourServerName;Trusted_Connection=Yes;APP=Decipher.accdb;DATABASE=DecipherScreens
On Error GoTo ErrorHappened
    Dim tmp As String
    
    tmp = ConString
    tmp = Right(tmp, Len(tmp) - InStr(tmp, "SERVER=") - 6)
    tmp = left(tmp, InStr(tmp, ";") - 1)
    
ExitNow:
On Error Resume Next
    GetServerNameFromConnectionString = tmp
Exit Function
ErrorHappened:
    tmp = vbNullString
    Resume ExitNow
    Resume
End Function

Public Function GetDatabaseNameFromConnectionString(ByVal ConString As String) As String
'Return the database name from SQL connection string
'ex: ODBC;DRIVER=SQL Server;SERVER=YourServerName;Trusted_Connection=Yes;APP=Decipher.accdb;DATABASE=DecipherScreens
On Error GoTo ErrorHappened
    Dim tmp As String
    
    tmp = ConString
    tmp = Right(tmp, Len(tmp) - InStr(tmp, "DATABASE=") - 8)
    
ExitNow:
On Error Resume Next
    GetDatabaseNameFromConnectionString = tmp
Exit Function
ErrorHappened:
    tmp = vbNullString
    Resume ExitNow
    Resume
End Function