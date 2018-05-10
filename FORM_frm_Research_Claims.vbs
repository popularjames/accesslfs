Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mvCalledBy As String
Private mvSql As CnlyScreenSQL
Private mvFieldsTable As String
Public Event UpdateSql()
Public Event QueryFormClose()
Public Event QueryFormRefresh()
Public Property Get SQL() As CnlyScreenSQL
    SQL = mvSql
End Property
Public Property Let SQL(SQL As CnlyScreenSQL)
    mvSql = SQL
End Property
Public Property Let FieldsTable(TableName As String)
    '* This allows using a different table for your fields. Do we need to expand to allow a select ?
    mvFieldsTable = TableName
End Property
Public Sub RefreshSubmissionList()
    On Error GoTo HandlerError
    
    Dim strSQL As String
    strSQL = " SELECT  Description, SS.SubmisionID, SS.CriteriaID , SS.UserID , SubmissionDate , ss.CompleteDate "
    strSQL = strSQL & " from [CONCEPT_CRITERIA_Submission] SS INNER JOIN Criteria_Hdr HH ON ss.CriteriaID = hh.criteriaID "
    
    Me.lstSubmissionQueue.RowSource = strSQL

    Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub RefreshListCriteria()
On Error GoTo HandlerError
Me.lstCriteria.RowSource = vbNullString
Me.lstCriteria.RowSource = " SELECT Sqlstring, fieldname,operator,fieldvalue, Tablename , IncludeFlag, CriteriaTempID from CRITERIA_TEMP WHERE USERID = '" & Identity.UserName & "'"
    
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
  
      
End Sub

Private Sub RefreshScreenListing()
On Error GoTo HandlerError
    Dim strSQL As String
    strSQL = " SELECT ScreenID, ScreenName FROM SCR_SCREENS "
    RefreshComboBox strSQL, cboScreen

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub

Private Function GetJoinCondition(strJoinTable As String, strMainRecordsource As String) As String
'This takes the Decipher Config data and builds a JOIN clause from it
Dim rst As DAO.RecordSet
Dim rstFields As DAO.RecordSet
Dim strSQLFrom As String
Dim intLoopCount As Integer
Dim strMasterField As String
Dim strChildField As String
Dim intTabID As String
Dim strTabRecordSource As String
On Error GoTo HandlerError

    'Based on the user selection, get us the sub table that links to the main
    Set rst = CurrentDb.OpenRecordSet("SELECT RecordSource , TabID FROM SCR_SCREENSTabs  where ScreenID = " & Me.cboScreen & " and Recordsource = '" & strJoinTable & "'", dbOpenSnapshot)
    While Not rst.EOF
        'Now get the information from the linked child table to start constructing the join
        intTabID = rst!TabID
        strTabRecordSource = rst!RecordSource
        Set rstFields = CurrentDb.OpenRecordSet("SELECT MasterField, ChildField FROM SCR_SCREENSTabsFields  where tabID = " & intTabID & " ", dbOpenSnapshot)
        'Start the statement.  We always start with the table we are looking to query from
        strSQLFrom = strSQLFrom & " FROM " & strTabRecordSource
        
        'Set this to know the first time we go in here
        intLoopCount = 1
        While Not rstFields.EOF
            'Link the fields based on the table!!
            If intLoopCount = 1 Then
                strSQLFrom = strSQLFrom & " WHERE " & strMainRecordsource & "."
            Else
                strSQLFrom = strSQLFrom & " AND  " & strMainRecordsource & "."
            End If
            strMasterField = rstFields!MasterField
            strChildField = rstFields!ChildField
            strSQLFrom = strSQLFrom & "" & strMasterField & " = " & strTabRecordSource & "." & strChildField
            intLoopCount = intLoopCount + 1
            rstFields.MoveNext
        Wend
        rst.MoveNext
     Wend
     
     GetJoinCondition = strSQLFrom
     
Exit Function

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GetJoinCondition = ""
      
End Function

Private Sub RefreshTableListing()
On Error GoTo HandlerError
    Dim strSQL As String
    strSQL = " SELECT PrimaryRecordSource, ScreenID FROM SCR_SCREENS  where ScreenID = " & Me.cboScreen & " "
    strSQL = strSQL & "UNION ALL  SELECT RecordSource as PrimaryRecordSource, screenid FROM SCR_SCREENSTabs  where ScreenID = " & Me.cboScreen & " "
    RefreshComboBox strSQL, cboTable

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
End Sub


Private Sub ExistsShow()
On Error GoTo HandlerError
    
    Me.cboExists.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL IN ( 'Exists', 'NOT EXISTS' )"
    'This is a stop gap.  I only want to expose certain operations based on what the user is selecting
    If Me.cboTable <> mvSql.From Then
        Me.cboExists.visible = True
        Me.lblExists.visible = True
    Else
        Me.cboExists.visible = False
        Me.lblExists.visible = False
    End If
   
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub





Private Sub cboScreen_Click()
On Error GoTo HandlerError
    Me.cboOperator.RowSource = vbNullString
    Me.cboTable.RowSource = vbNullString
    Me.cboField.RowSource = vbNullString
    Me.txtValue = ""
    Me.lstCriteria = vbNullString
    Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL NOT IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
    Me.cboField = ""
    Me.cboTable = ""
    ClearListBox Me.lstCriteria
    RefreshTableListing
    Me.txtFrom = ""
    mvCalledBy = Me.cboScreen
    UpdateFilterList
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
End Sub




Private Sub cboTable_Change()
On Error GoTo HandlerError
ExistsShow
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub



Private Sub cboTable_Click()
    Dim rst As DAO.RecordSet
    Dim rstFields As DAO.RecordSet
    Dim intTabID As Integer
    Dim strTabRecordSource As String
    Dim strMasterField As String
    Dim strChildField As String
    Dim strMainRecordsource As String
    Dim strSQLFrom As String
    Dim intLoopCount As String
    On Error GoTo HandlerError

    mvFieldsTable = cboTable
    
    Me.cboField = ""
    Me.cboOperator = ""
    Me.txtValue = ""
       
    'Fill the field listng with the columns from the table the user is querying from
    Me.cboField.RowSource = mvFieldsTable
    Me.cboField.RowSource = "SELECT FieldName , isindex FROM v_XREF_TableFields WHERE TableName = '" & mvFieldsTable & "' ORDER BY FieldName"
    Me.cboField.RowSourceType = "Table/Query"
    
    'Set the SQL Object's From Clause to be the main table that drives the decipher screen
    mvSql.From = DLookup("PrimaryRecordSource", "SCR_SCREENS", "ScreenID = " & Me.cboTable.Column(1) & "")
    
    'This is a stop gap.  I only want to expose certain operations based on what the user is selecting
    'If Me.cboTable <> mvSql.From Then
    '   Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL NOT IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
       'Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
    ''Else
    '   Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL NOT IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
    'End If
   
    ExistsShow
    
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub CmdAdd_Click()
On Error GoTo HandlerError

    Dim sValue As String
    Dim sClause As String
    Dim sFiel7d As String
    Dim X As Integer
    Dim ln As Integer
    Dim sVal As String
    Dim sLval As String
    Dim sRval As String
  Dim sField As String
    Dim strJoinTable As String
    Dim strJoinClause  As String
Dim strSQL As String
    Dim intOperatorID As Integer
    
    Dim strCustomType As String
        
        
    Me.lstCriteria.RowSource = vbNullString
    intOperatorID = Nz(DLookup("OperatorID", "CRITERIA_OPERATOR", "OperatorSQL = '" & Nz(Me.cboOperator, "") & "'"), 0)
    strCustomType = Nz(DLookup("DataType", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & ""), "")
    
    If Nz(Me.cboOperator, "") = "" Then
        MsgBox "Select an Operator"
        Me.cboOperator.SetFocus
    ElseIf (Nz(Me.cboField, "") = "" And strCustomType = "") Then
            MsgBox "Select a field"
            Me.cboField.SetFocus
    ElseIf Nz(Me.txtValue, "") = "" Then
        MsgBox "Enter a Value"
        Me.txtValue.SetFocus
    Else
        
        
        sVal = Me.txtValue.Value
        sField = Nz(Me.cboField, "")

        If Me.cboOperator <> "BETWEEN" Then
            sValue = BuildCondition(sVal, strCustomType)
        Else
            sValue = sVal
        End If
                
                
        If sValue <> "" Then
            strJoinTable = Me.cboTable
            strJoinClause = GetJoinCondition(strJoinTable, mvSql.From)
            Select Case Me.cboOperator
                Case Is = "Between"
                    X = InStr(sValue, "AND")
                    ln = Len(sValue)
                    sLval = BuildCondition(Trim(left(sValue, X - 1)))
                    sRval = BuildCondition(Trim(Right(sValue, ln - (X + 2))))
                    sClause = sField & " " & Me.cboOperator & " " & sLval & " and " & sRval
                 Case Is = "NOT IN"
                    sClause = sField & " " & Me.cboOperator & " (" & sValue & ")"
                 Case Is = "IN"
                    sClause = sField & " " & Me.cboOperator & " (" & sValue & ")"
                 Case Is = "Exists"
                     sClause = " EXISTS (SELECT 1 " & strJoinClause & _
                             " AND " & strJoinTable & "." & sField & " IN (" & sValue & "))"
                 Case Is = "Not Exists"

                         sClause = " NOT EXISTS (SELECT 1 " & strJoinClause & _
                             " AND " & strJoinTable & "." & sField & " IN (" & sValue & "))"
                 Case Is = "Exists LK"
                     sClause = " EXISTS (SELECT 1 " & strJoinClause & _
                             " AND " & strJoinTable & "." & sField & " LIKE " & sValue & ")"
                 Case Is = "N EXSTS LK"
                     sClause = "NOT EXISTS (SELECT 1 " & strJoinClause & _
                             " AND " & strJoinTable & "." & sField & " LIKE " & sValue & ")"
                 Case Else
                    sClause = sField & " " & Me.cboOperator & " " & sValue
            End Select
            'Me.lstCriteria.AddItem Chr(34) & sClause & Chr(34) & ";" & Me.cboField & ";" & Me.cboOperator & ";" & Chr(34) & Me.txtValue & Chr(34)
                
            Dim strExists As String
                
            If strJoinTable <> mvSql.From Then
                If Nz(Me.cboExists.Value, "") = "" Then
                    MsgBox "You must select a subquery type when choosing a related table."
                    Me.cboExists.SetFocus
                    
                    Exit Sub
                Else
                    strExists = Me.cboExists
                End If
            
            Else
                strExists = ""
            End If
                
        
            strSQL = " INSERT INTO CRITERIA_TEMP  ([Operator]           ,[FieldName]           ,[FieldValue]           ,[SQLString]           ,[UserID], TableName, IncludeFlag)"
            strSQL = strSQL & " values ("
            strSQL = strSQL & Chr(34) & Me.cboOperator & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.cboField & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Me.txtValue & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & sClause & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Identity.UserName & Chr(34) & " ,  "
            strSQL = strSQL & Chr(34) & strJoinTable & Chr(34) & " , "
            strSQL = strSQL & Chr(34) & strExists & Chr(34) & " ) "
            
            CurrentDb.Execute strSQL
            

            BuildQuery
            RefreshListCriteria
            RaiseEvent UpdateSql
        Else
            MsgBox "Unable to build condition.  Try again or get help.", vbOKOnly, "Unable to Build Condition"
        End If
            
        Me.cboField.SetFocus
        
    End If
    RefreshListCriteria
    Me.txtFrom = mvSql.SqlAll
exitHere:
    On Error Resume Next
    Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub

Private Sub cmdClear_Click()
ClearListBox Me.lstCriteria
Me.cboField = ""
Me.cboOperator = ""
Me.txtValue = ""
Me.txtFrom = ""

End Sub

Private Sub cmdDeleteFilter_Click()
  On Error GoTo HandlerError
  Dim strSQL As String
        Dim varItem As Variant
    
        If Me.lstFilters.ItemsSelected.Count <> 0 Then
            varItem = Me.lstFilters.ItemsSelected(0)
            strSQL = "DELETE FROM Criteria_Hdr WHERE CriteriaId = " & Me.lstFilters.Column(0, varItem)
            CurrentDb.Execute strSQL, dbSeeChanges
            Me.lstFilters.Requery
        Else
            MsgBox "Select a row to delete", vbOKOnly, "Select a Row"
        End If
        
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub cmdDeleteJob_Click()
Dim lngSubmissionID As Long
Dim varItem As Variant
Dim strSQL As String
On Error GoTo HandlerError
     

varItem = Me.lstSubmissionQueue.ItemsSelected(0)
lngSubmissionID = Me.lstSubmissionQueue.Column(1, varItem)

strSQL = " DELETE   from CONCEPT_CRITERIA_Submission WHERE SubmisionID = " & lngSubmissionID & " AND USERID = '" & Identity.UserName & "' and COmpleteDate is null "

CurrentDb.Execute strSQL, dbSeeChanges

RefreshSubmissionList
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub cmdHeavy_Click()
Dim sMsg As String
On Error GoTo HandlerError
    Dim lEstRows As Long
    isHeavy mvSql.From, mvSql.SqlAll, 25000, lEstRows
    sMsg = "The estimated row count is: " & CStr(lEstRows) & vbCrLf & _
    "The query plan suggests a very large result set. It may take a while to return the actual row count. " & vbCrLf & _
    "Would you like to continue?"
    'If lEstRows > CQT_lThreshold Then
     MsgBox sMsg, vbOKOnly, "Row Count Estimate"
        
    'End If
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub cmdLoad_Click()
On Error GoTo HandlerError
    Me.lstCriteria.RowSource = vbNullString
    LoadFilter
    Me.txtFrom = mvSql.SqlAll
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub




Private Sub cmdRefresh_Click()
    RefreshSubmissionList
End Sub

Private Sub CmDRun_Click()
 On Error GoTo HandlerError
   Dim lngFilterId As Long
    Dim strSQL As String
    Dim strSqlType As String
    Dim varItem As Variant
    Dim rst As DAO.RecordSet
    Dim intItemsinlist As Integer
    Dim intCounter As Integer
    
     

    If Me.lstFilters.ItemsSelected.Count <> 0 Then
        varItem = Me.lstFilters.ItemsSelected(0)
        lngFilterId = Me.lstFilters.Column(0, varItem)
    
        strSQL = " SELECT * from CRITERIA_HDR WHERE CriteriaID = " & lngFilterId
        Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
        While Not rst.EOF
        
        
        
            strSQL = " INSERT INTO CONCEPT_CRITERIA_Submission "
            strSQL = strSQL & "        ([CriteriaID]"
            strSQL = strSQL & "        ,[UserID]"
            strSQL = strSQL & "          ,[SubmissionDate], [SqlString]) "
             
             
            strSQL = strSQL & "         values"
            strSQL = strSQL & "               (" & lngFilterId & ","
            strSQL = strSQL & Chr(34) & Identity.UserName & Chr(34) & ","
            strSQL = strSQL & Chr(34) & Now & Chr(34) & ", " & Chr(34) & rst!SqlAll & Chr(34) & ") "
            CurrentDb.Execute strSQL
            rst.MoveNext
         Wend
    Else
        MsgBox "Choose a filter to execute"
    End If
    
    RefreshSubmissionList
    
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


Private Sub cmdSave_Click()
 On Error GoTo HandlerError
 Dim sFilterName As String

    '** ADO Parameters
Dim rst As DAO.RecordSet
    Dim strConnect As String
    Dim cnn As Variant
    Dim cmd As Variant
    Dim cmd2 As Variant
    Dim strSQL As String
    Dim dtpDatestamp As Date
    Dim lngpCriteriaId As Long
    Dim Lst As listBox
    Dim ctr As Single
    
    
    If Not IsSQLParse Then
        MsgBox "Unable to save filter.  There is a syntax error with the SQL statement!", vbCritical
        Exit Sub
    End If
    
    
    If mvSql.WherePrimary <> "" And Me.lstCriteria.ListCount > 0 Then
        
        sFilterName = EscapeQuotes(Nz(InputBox("Filter Name:", "Save Filter"), ""))
    
        If Nz(DLookup("Description", "Criteria_Hdr", "Description = '" & sFilterName & "'"), "") <> "" Then '* Make sure filter doesn't exist
            
            sFilterName = ""
            
            MsgBox "There is already a filter with that name.  Please choose a different name", vbOKOnly

        End If
        
            
        If sFilterName <> "" Then
            
            '** Create objects & setup connection
            Set Lst = Me.lstCriteria
            Set cnn = CreateObject("ADODB.Connection")
            Set cmd = CreateObject("ADODB.Command")
            Set cmd2 = CreateObject("ADODB.Command")
            
            dtpDatestamp = Now()
            strConnect = GetConnectString("v_CODE_Database")
            cnn.Open strConnect
            
            '*** Header Insert
            cmd.ActiveConnection = cnn
            cmd.CommandTimeout = 30
            cmd.CommandText = "usp_CRITERIA_Hdr_Insert"
            cmd.commandType = adCmdStoredProc
            cmd.Parameters.Refresh
        
            cmd.Parameters("@pDatestamp") = Now
            cmd.Parameters("@pUserID") = Identity.UserName
            cmd.Parameters("@pClmType") = DLookup("NoteText", "SCR_SCREENSNOTES", "ScreenID = " & Me.cboScreen & "")
            cmd.Parameters("@pSourceObject") = Me.cboScreen
            cmd.Parameters("@pDescription") = sFilterName
            cmd.Parameters("@pSQLFrom") = mvSql.From
            cmd.Parameters("@pSQLWhere") = mvSql.WherePrimary
            cmd.Parameters("@pSQLOrderBy") = mvSql.OrderBy
            cmd.Parameters("@pSQLGroupBy") = ""
            cmd.Parameters("@pSQLHaving") = ""
            cmd.Parameters("@pSQLAll") = mvSql.SqlAll
            'TL add account ID logic
            cmd.Parameters("@pAccountID") = gintAccountID
            cmd.Parameters("@pCriteriaId") = 0
            cmd.Parameters("@pErrMsg") = ""
            
            
            cmd.Execute 128 '* adExecuteNoRecords
            
            '* if cmd.recordsaffected <> 1 then msgbox "error"
            
            lngpCriteriaId = Nz(cmd.Parameters("@pCriteriaId"), 0)
            
            If lngpCriteriaId <> 0 Then
            
                '***Detail Insert
                cmd2.ActiveConnection = cnn
                cmd2.CommandTimeout = 30
                cmd2.CommandText = "usp_CRITERIA_Dtl_Insert"
                cmd2.commandType = adCmdStoredProc
                
                cmd2.Parameters.Refresh
                
                
                 strSQL = "SELECT SqlString, FieldName, Operator, FieldValue, TableName, IncludeFlag FROM CRITERIA_TEMP WHERE UserID  = '" & Identity.UserName & "'"
                 Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
           
                
                While Not rst.EOF
                        cmd2.Parameters("@pDateStamp") = dtpDatestamp
                        cmd2.Parameters("@pCriteriaID") = lngpCriteriaId
                        cmd2.Parameters("@pOperator") = rst!Operator
                        
                        cmd2.Parameters("@pFieldName") = rst!FieldName
                        cmd2.Parameters("@pFieldValue") = rst!FieldValue
                        cmd2.Parameters("@pSQLString") = rst!sqlString
                        cmd2.Parameters("@pTableName") = rst!TableName
                        cmd2.Parameters("@pIncludeFlag") = rst!IncludeFlag
                    
                        cmd2.Execute 128 '* adExecuteNoRecords
                        
                    rst.MoveNext
                Wend
            Else
            
                MsgBox "Error Saving Filter"
                
            End If '* lngpCriteriaId <> 0
                       
        End If '*  sFilterName <> ""
        
    Else
    
        MsgBox "There is no filter to save.", vbOKOnly, "No filter"
    
    End If  '* mvSql.WherePrimary <> "" And Me.lstCriteria.ListCount > 0
    
    UpdateFilterList

exitHere:
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub
Private Function IsSQLParse() As Boolean
    Dim oCon As New ADODB.Connection
    On Error GoTo ErrHandler


    oCon.ConnectionString = GetConnectString(mvSql.From)
    oCon.CommandTimeout = 0
    oCon.Open
    oCon.Execute "SET PARSEONLY ON " & mvSql.SqlAll & ""
    oCon.Close
    
    IsSQLParse = True

exitHere:
    Set oCon = Nothing
    Exit Function
ErrHandler:
    IsSQLParse = False
    If oCon.Errors.Count <> 0 Then
        IsSQLParse = False
    Else
        IsSQLParse = False
    End If
    Resume exitHere
End Function


Private Sub cmdSyntax_Click()

'Dim oCon As New ADODB.Connection
'On Error GoTo ErrHandler'''


    'oCon.ConnectionString = GetConnectString(mvSql.From)
    'oCon.CommandTimeout = 0
    'oCon.Open
    'oCon.Execute "SET PARSEONLY ON " & mvSql.SqlAll & ""
    'oCon.Execute "SET PARSEONLY ON " & mvSql.SqlAll & ""
    'oCon.Close
    If IsSQLParse Then
        MsgBox "OK!  No Errors Found!!!"
    Else
        MsgBox "ERROR WITH SQL Statament!!!", vbCritical
    End If

'exitHere:
'    Set oCon = Nothing
'    Exit Sub
'ErrHandler:
'    'IsSQLParse = False
'    If oCon.Errors.count <> 0 Then
'        MsgBox Err.Description
'    Else
'        MsgBox "Error IsSQLParse: " & Err.Description
'    End If
'    Resume exitHere
End Sub

Private Sub Form_Load()
On Error GoTo HandlerError
        Me.lstFilters.RowSource = vbNullString
        
    RefreshScreenListing
    Me.cboOperator.RowSource = "CRITERIA_Operator"
    Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL NOT IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
    Me.lstCriteria.RowSource = vbNullString
    RefreshSubmissionList
    Me.txtValue = ""

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub



Private Function BuildCondition(sValue As String, Optional strCustomType As String = "") As String
 On Error GoTo HandlerError
   '* BuildCondition(ctl As Control) As String was Passing a control to accomodate using the combo-box/text box as entry field
    Dim sWhere As String
    Dim sDataType As String
    Dim sField As String
    'Dim sValue As String

 '   sValue = ctl.Value

    If strCustomType = "" Then
        sDataType = CurrentDb.TableDefs(Me.cboTable).Fields(Me.cboField).Type
    Else
        sDataType = strCustomType
    End If

    Select Case sDataType
        Case 2 To 7 '* Numbers--do nothing
            sField = sValue
        Case Is = 8 '* date
            sField = "'" & sValue & "'" '* Access Syntax
           '*  sField = "'" & sValue & "'" SQL Server Syntax for Pass-Through
        Case Is = 10 'text
            sField = "'" & Replace(sValue, ",", "','") & "'"   '* this handles  IN lists
        Case Is = 15 '* bigint--do nothing
            sField = sValue
        Case Is = 18 '* char--this handles  IN lists
            sField = "'" & Replace(sValue, ",", "','") & "'"   '* this handles  IN lists
        Case 19 To 21 '* numeric, decimal, float--do nothing
            sField = sValue
        Case Else '** Error
           sField = ""
    End Select

    BuildCondition = sField


Exit Function

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
exitHere:
    BuildCondition = sField
    Exit Function
End Function
Private Sub BuildQuery()
On Error GoTo HandlerError
    Dim strTableName As String
    Dim rst As DAO.RecordSet
    Dim strSQL As String
    Dim strCondition As String
    Dim strSubCondition As String
    Dim strIncludeFlag As String
    Dim strPreviousIncludeFlag As String
    Dim X As Single
    Dim strJoinTable As String
    Dim strPreviousJoinTable As String
    Dim strJoinClause As String
        
        strSQL = "SELECT SqlString, FieldName, Operator, FieldValue, TableName FROM CRITERIA_TEMP WHERE UserID  = '" & Identity.UserName & "' and Tablename = '" & mvSql.From & "' "
        Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
        'Build the main where clause with no sub tables
        
        X = 1
        While Not rst.EOF
            strTableName = rst!TableName
            If X = 1 Then
                strCondition = rst!sqlString
            Else
                strCondition = strCondition & " and " & rst!sqlString
            End If

            X = X + 1
            rst.MoveNext
        Wend
        
        
        strSubCondition = ""
        strSQL = "SELECT SqlString, FieldName, Operator, FieldValue, TableName, IncludeFlag  FROM CRITERIA_TEMP WHERE UserID  = '" & Identity.UserName & "' and Tablename <> '" & mvSql.From & "' ORDER BY TABLENAME, IncludeFLag "
        Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
        'Build the main where clause with no sub tables
        
        X = 1
        While Not rst.EOF
            strJoinTable = rst!TableName
            strIncludeFlag = Nz(rst!IncludeFlag)
            
            If ((strPreviousJoinTable <> strJoinTable) Or (strIncludeFlag <> strPreviousIncludeFlag)) Then
                If strSubCondition <> "" Then
                    strSubCondition = strSubCondition & ") and "
                End If
                    strJoinClause = GetJoinCondition(strJoinTable, mvSql.From)
                X = 1
            End If
            
            If X = 1 Then
                strSubCondition = strSubCondition & " " & strIncludeFlag & " ( SELECT 1 " & strJoinClause & " and " & rst!sqlString
            Else
                strSubCondition = strSubCondition & " and " & rst!sqlString
            End If

            X = X + 1
            strPreviousJoinTable = strJoinTable
            strPreviousIncludeFlag = strIncludeFlag
            rst.MoveNext
        Wend
        
        If strSubCondition <> "" Then
            strSubCondition = strSubCondition & " ) "
        End If
        
    With mvSql
        If Nz(strSubCondition, "") = "" Then
            .WherePrimary = strCondition
        Else
            .WherePrimary = strCondition & " AND " & Nz(strSubCondition, "")
        End If
        
        .SqlAll = "SELECT " & IIf(.Select <> "", .Select, "*") & _
                  " FROM " & .From & " WHERE " & .WherePrimary
    End With
    Me.txtFrom = mvSql.SqlAll
exitHere:

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


'DPR - Add table references...
Private Sub lstCriteria_DblClick(Cancel As Integer)
On Error GoTo HandlerError
     Dim lngCriteriaTempID As Integer
     Dim strSQL As String
  Dim varItem As Variant

    
    
    varItem = Me.lstCriteria.ItemsSelected(0)
        
    If Me.lstCriteria.Column(1, varItem) <> "" Then
        
        Me.cboField = Me.lstCriteria.Column(1, varItem)
        Me.cboTable = Me.lstCriteria.Column(4, varItem)
        'Me.cboOperator.RowSource = "SELECT * FROM CRITERIA_Operator WHERE OperatorSQL NOT IN ( 'Exists', 'NOT EXISTS', 'Exists LK', 'N EXSTS LK'  )"
        Me.cboOperator = Me.lstCriteria.Column(2, varItem)
        Me.txtValue = Me.lstCriteria.Column(3, varItem)
        ExistsShow
        Me.cboExists = Nz(Me.lstCriteria.Column(5, varItem), "")
        lngCriteriaTempID = Nz(Me.lstCriteria.Column(6, varItem), "")
        
        Me.lstCriteria.RowSource = vbNullString
        
        strSQL = " DELETE FROM CRITERIA_TEMP WHERE UserID = " & Chr(34) & Identity.UserName & Chr(34) & " and CriteriaTempID = " & lngCriteriaTempID & ""
        CurrentDb.Execute strSQL, dbSeeChanges
        
        RefreshListCriteria
       
        BuildQuery
        RaiseEvent UpdateSql
       
    Else
        MsgBox "This row cannot be reloaded.  Open in the editor to make changes", vbOKOnly, "Row Can't be Loaded"
    End If
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub

Private Sub ClearListBox(lstBox As listBox)
Dim strSQL As String
On Error GoTo HandlerError
    strSQL = " DELETE FROM CRITERIA_TEMP WHERE UserID = '" & Identity.UserName & "'"
    CurrentDb.Execute strSQL, dbSeeChanges
    RefreshListCriteria
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub

Private Sub LoadFilter()
On Error GoTo HandlerError

    Dim lngFilterId As Long
    Dim strSQL As String
    Dim strSqlType As String
    Dim varItem As Variant
    Dim rst As DAO.RecordSet
    Dim intItemsinlist As Integer
    Dim intCounter As Integer
     
    Me.lstCriteria.RowSource = vbNullString
    strSQL = " DELETE FROM CRITERIA_TEMP WHERE UserID = '" & Identity.UserName & "'"
    CurrentDb.Execute strSQL, dbSeeChanges
     
    For intCounter = 0 To intItemsinlist - 1
      Me.lstCriteria.RemoveItem 0
    Next

    If Me.lstFilters.ItemsSelected.Count <> 0 Then
        varItem = Me.lstFilters.ItemsSelected(0)
        lngFilterId = Me.lstFilters.Column(0, varItem)
        strSQL = "SELECT SqlString, FieldName, Operator, FieldValue, TableName, IncludeFlag FROM Criteria_Dtl WHERE CriteriaId = " & lngFilterId
        'RefreshListBox strSQL, Me.lstCriteria
        
        Set rst = CurrentDb.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
        While Not rst.EOF
            
            strSQL = " INSERT INTO CRITERIA_TEMP ([Operator]  ,[FieldName]           ,[FieldValue]           ,[SQLString]           ,[UserID], TableName, IncludeFlag)"
            strSQL = strSQL & " values ("
            strSQL = strSQL & Chr(34) & rst!Operator & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & rst!FieldName & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & rst!FieldValue & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & rst!sqlString & Chr(34) & ", "
            strSQL = strSQL & Chr(34) & Identity.UserName & Chr(34) & " ,  "
            strSQL = strSQL & Chr(34) & Nz(rst!TableName, "") & Chr(34) & " ,  "
            strSQL = strSQL & Chr(34) & Nz(rst!IncludeFlag, "") & Chr(34) & " )  "
            CurrentDb.Execute strSQL, dbSeeChanges
            RefreshListCriteria
            rst.MoveNext
        Wend
        
        strSqlType = "SELECT * FROM CRITERIA_HDR WHERE CriteriaId = " & lngFilterId
        Set rst = CurrentDb.OpenRecordSet(strSqlType, dbOpenSnapshot, dbReadOnly)
        
        With mvSql
            .From = rst("SqlFrom")
            .WherePrimary = rst("SqlWhere")
            .OrderBy = rst("SqlOrderBy")
            .SqlAll = rst("SqlAll")
        End With
    Else
        MsgBox "Choose a filter to load"
    End If
    BuildQuery
exitHere:

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub





Private Sub UpdateFilterList()
On Error GoTo HandlerError


  Dim strSource  As String

    ' TL add account ID logic and logic to filter by SQLFrom
    strSource = "SELECT CriteriaId, Description FROM CRITERIA_Hdr " & _
                " WHERE AccountID = " & gintAccountID & " and SourceObject = '" & mvCalledBy & "'"

    With Me.tglType
        If .Value = -1 Then
            .Caption = "Mine"
            strSource = strSource & " and UserId ='" & Identity.UserName & "'"
        Else
            .Caption = "All"
        End If
     
        'Alex added 7/10/08
        strSource = strSource & " ORDER BY Description"
        
        Me.lstFilters.RowSource = strSource
        Me.lstFilters.Requery
    End With
exitHere:

Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub


'added by Michal Roguski - 1/7/11
'returns estimated number of rows to be returned by query
'based on SQL's estimated execution plan
Private Function isHeavy(strSource As String, ByVal strQuery As String, rowThresh As Long, Optional ByRef lEstRows As Long, Optional pbolSilent As Boolean = False) As Boolean
On Error GoTo HandlerError

    Dim strConnect As String
    Dim cnn As Variant
    Dim cmd As Variant
    Dim oRs As Variant
    Dim oSt As Variant
    Dim sMsg As String
    Dim estNoRows As String
    Dim sqlStr As String
    Dim TxtAggregator As String
    'Rx CUSTOMIZATION
    Dim lbolGoToexitHere As Boolean
    Dim lbolErrorOccured As Boolean
    
    'being optimistic
    isHeavy = False
  
      'added by SG on 12/17/2012: convert to SQL side statement
      If TxtAggregator = "" Then
            strQuery = Replace(strQuery, "#", "'") 'Parses out hashmarks around dates TW 1/14/2011
            strQuery = Replace(strQuery, "NZ", "COALESCE")
            'additional parsing for wildcards
            If (InStr(strQuery, " WHERE ")) > 0 Then
              strQuery = left(strQuery, InStr(strQuery, " WHERE ")) & Replace(Right(strQuery, Len(strQuery) - (InStr(strQuery, " WHERE "))), "*", "%")
            End If
      End If
      'updated version not utilizing XML any more - added by Michal Roguski 1/26/11
      'prepare objects
      Set cnn = CreateObject("ADODB.Connection")
      Set cmd = CreateObject("ADODB.Command")
      Set oRs = CreateObject("ADODB.RecordSet")
      Set oSt = CreateObject("ADODB.Stream")
      
      strConnect = GetConnectString(strSource)
      cnn.Open strConnect
      cmd.ActiveConnection = cnn
      cmd.CommandTimeout = 30
      
      'enable estimated query execution plan
      cmd.CommandText = "SET SHOWPLAN_ALL ON"
      cmd.commandType = 1 'adCmdText
      cmd.Execute
    
      'read estimated number of rows from thread with no parent - this is final select
      estNoRows = "-1"
      With oRs
        .CursorLocation = adUseClient
        .Open strQuery, cnn, adOpenStatic, adLockReadOnly, adCmdText
          
        oRs.MoveFirst
        Do Until oRs.EOF
            If (CInt(oRs!Parent) = 0) Then
                estNoRows = CStr(oRs!EstimateRows)
                Exit Do
            End If
            oRs.MoveNext
        Loop
       
      End With
    
      'verify if query is exceeding treshold
      If (CLng(estNoRows) > rowThresh) Then
          isHeavy = True
      End If
      lEstRows = CLng(estNoRows)
 
    
exitHere:
    On Error Resume Next
    DoCmd.Hourglass False
    'cleaneup
    If oSt.State = 1 Then
        oSt.Close: Set oSt = Nothing
    End If
    If oRs.State = 1 Then
        oRs.Close: Set oRs = Nothing
    End If
    If cnn.State = 1 Then
        cnn.Close: Set cnn = Nothing
    End If
    Exit Function

HandlerError:
    If (estNoRows = "-1") Then
      GoTo HandleError2
    End If
    If pbolSilent = False Then
        MsgBox "Row number estimation error." & vbCr & vbCr & _
           "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    End If
    Resume exitHere
HandleError2:
    If pbolSilent = False Then
        MsgBox "Query did not return appropriate execution plan - please contact your DA", vbOKOnly, "Error"
    End If
        Resume exitHere

End Function

Private Sub tglType_Click()
On Error GoTo HandlerError
    UpdateFilterList
Exit Sub

HandlerError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    
      
End Sub
