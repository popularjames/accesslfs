Option Compare Database
Option Explicit

Public Function RestoreScreenFromXML(SourceDB As String, DestDB As String, ScreenName As String, Optional BuildSilent As Boolean = False) As Boolean
On Error GoTo ErrorHappened
    'Ex. RestoreScreenFromXML Currentdb.name, currentdb.name, "Claims",FALSE
    Dim i As Integer
    Dim xRst As RecordSet
    Dim rst As Object 'ADO recordset
    Dim ColScripts As New Collection
    Dim strSQL As String
    Dim DbSrc As Database
    Dim DbDest As Database
    Dim TableName As String
    Dim StrSqlEa() As String
    Dim InstallScreen As Boolean
    Dim addInManager As New CT_ClsCnlyAddinSupport

    If LenB(SourceDB) = 0 Then
        Set DbSrc = CurrentDb
    Else
        Set DbSrc = DBEngine.OpenDatabase(SourceDB)
    End If
    
    If LenB(DestDB) = 0 Then
        Set DbDest = CurrentDb
    Else
        Set DbDest = DBEngine.OpenDatabase(DestDB)
    End If
    
    'Check for a screen with this name in the destination
    If ExDLookup(DbDest, "ScreenID", "SCR_Screens", "ScreenName='" & Replace(ScreenName, "'", "''") & "'") <> "" Then
        If Not BuildSilent Then
            'Offer to overwrite or cancel.
            If MsgBox("A Screen with that name already exists in the destination database." _
                & vbCrLf & "Would you like to overwrite it?", vbQuestion + vbYesNo, "ImportScreenFromXML") = vbYes Then
                InstallScreen = True
            End If
        Else
            InstallScreen = True
        End If
        
        If InstallScreen Then
            'Delete the screen, rely on cascading deletes to clean up corresponding child table entries.
            strSQL = "Delete FROM SCR_Screens where ScreenName = '" & Replace(ScreenName, "'", "''") & "'"
            DbDest.Execute strSQL, dbFailOnError
        End If
    Else
        InstallScreen = True
    End If
    
    If InstallScreen Then
        'Process XML for tables Matching ScreenName
        strSQL = "SELECT [TableName],[XML],[Type] FROM SCR_ScreensXML WHERE ScreenName='" & Replace(ScreenName, "'", "''") & "' " & "ORDER BY [Level], [Type]"
        
        Set xRst = DbSrc.OpenRecordSet(strSQL, dbOpenSnapshot, dbReadOnly)
        With xRst
            If Not (.EOF And .BOF) Then 'xrst has records
                'Proces Tables in progressing Level order (starting with 0) for Each Table
                .MoveFirst
                
                Do Until .EOF
                    TableName = xRst.Fields("TableName")
                    
                    If xRst.Fields("[Type]") = 0 Then 'Base table
                        'Set all RefID(s) in corresponding base Table to 0
                        strSQL = "UPDATE " & TableName & " SET " & TableName & ".RefID = 0 WHERE RefID<>0"
                        DoCmd.SetWarnings False
                        DbDest.Execute strSQL
                        DoCmd.SetWarnings True
                    End If
                    'Store the Max(PrimaryKey.ID) in the corresponding DbTable
                    'PKey = GetPrimKeyName(TableName)
                    'MaxKeyVal = DLookup("Max(" & PKey & ")", TableName)

                    'Get the Rst for the Table using PutXML()
                    PutXML xRst.Fields("[XML]"), rst
     
                    'Get InsertSQL for the Table using CreateInsertSQL(TableName)
                    strSQL = CreateInsertSQL(TableName, rst)
                                    
                    'Do the import:
                    If strSQL = "Error" Then
                        GoTo ErrorHappened
                    End If

                    StrSqlEa = Split(strSQL, "<%EOL%>")
                    For i = 0 To UBound(StrSqlEa) - (IIf(UBound(StrSqlEa) > 0, 1, 0))
                        StrSqlEa(i) = Replace(StrSqlEa(i), "ScrMainScreens", "SCR_MainScreens")
                        DbDest.Execute StrSqlEa(i)
                    Next i
                    .MoveNext
                Loop
            End If 'xRst has records
        End With 'xRst
        
        If Not BuildSilent Then
            addInManager.BuildRibbonBar
        End If
    End If
    
    RestoreScreenFromXML = True
ExitNow:
On Error Resume Next
    Set ColScripts = Nothing
    xRst.Close
    Set xRst = Nothing
    DbSrc.Close
    Set DbSrc = Nothing
    DbDest.Close
    Set DbDest = Nothing
Exit Function
ErrorHappened:
    RestoreScreenFromXML = False
    MsgBox "Error Restoring table " & TableName & "." & vbCrLf & Err.Description, vbCritical, "RestoreScreenFromXML()"
    Resume ExitNow
    Resume
End Function

Public Function SaveScreenToXML(SourceDB As String, DestDB As String, ScreenName As String) As Boolean
'Save screen and settings to XML table SCR_ScreensXML
'SA 11/13/2012 - Changed from sub to boolean function
On Error GoTo ErrorHappened:
    'DS 11/11/11 - changed method to write XML to use ADO parms so that we can store XML docs >64k
    'Ex: SaveScreenToTable currentdb.Name, currentdb.Name, "Duplicate Payment Analysis"
    Dim Result As Boolean
    Dim ScreenID As Long
    Dim ColScripts As New Collection
    Dim Script As CT_ClsExportScript
    Dim strSQL As String
    Dim strXML As String
    Dim DbSrc As Database, DbDest As Database
    Dim objCon As Object 'ADODB.Connection
    Dim objCom As Object 'ADODB.Command

    If SourceDB = vbNullString Then
        Set DbSrc = CurrentDb
    Else
        Set DbSrc = DBEngine.OpenDatabase(SourceDB)
    End If

    If DestDB = vbNullString Then
        Set DbDest = CurrentDb
    Else
        Set DbDest = DBEngine.OpenDatabase(DestDB)
    End If

    'Given a ScreenName
    'Get the ScreenID for that ScreenName
    ScreenID = ExDLookup(DbSrc, "ScreenID", "SCR_Screens", "ScreenName = '" & ScreenName & "'")

    If ScreenID > 0 Then 'Screen Name exists
        'Get the SQL to for that ScreenID and all releated elements into a collection using CreateExportSQL
        CreateItemExportSQL DbSrc, "SCR_Screens", CStr(ScreenID), ScreenName, ColScripts, 0

        'Use an Access Table: ([PK]ScreenXmlID, ScreenName, TableName, Level, XML, Type)
        'If table for storing XML does not exist, create it.
        If TableExists("SCR_ScreensXML", DbDest) Then
            'If ScreenName exists in Table prompt (Replace/Cancel)
            If (ExDLookup(DbDest, "ScreenName", "SCR_ScreensXML", "ScreenName = '" & ScreenName & "'") <> "") Then
                If MsgBox("A screen called '" & ScreenName & "' already exists in storage." & vbCrLf & "Replace It?", vbOKCancel) = vbCancel Then
                    GoTo ExitNow
                Else
                    'Delete the record for the corresponding ScreenName
                    strSQL = "DELETE FROM SCR_ScreensXML WHERE ScreenName='" & ScreenName & "'"

                    DoCmd.SetWarnings False
                    DbDest.Execute strSQL
                    DoCmd.SetWarnings True
                End If 'Replace/Cancel
            End If 'ScreenName exists
        End If 'Table Exists

        Set objCon = CreateObject("ADODB.Connection")

        ' connect to current database via ADO
        objCon.Open LINK_SRC_ACCESS & "Data Source=" & Application.CurrentProject.Path & "\" & Application.CurrentProject.Name & ";Persist Security Info=False"

        'Put it in the table via parameterized insert
        strSQL = "INSERT INTO SCR_ScreensXML([ScreenName], [TableName], [Level], [XML], [Type]) "
        strSQL = strSQL & "VALUES (@ExpName,@ExpTable,@ExpLevel,@Xml,@ExpTyp)"

        'For Each Item In Collection
        For Each Script In ColScripts
            'Use GetXML to retrieve the XML for each item and insert row in Table
            strXML = GetXML(DbSrc.Name, Script.ExpScript)

            Set objCom = CreateObject("ADODB.Command")
            With objCom
                .CommandText = strSQL
                .commandType = adCmdText  'Type : CommandText
                Set .ActiveConnection = objCon
                .Parameters.Append .CreateParameter("@ExpName", adVarChar, adParamInput, 50, Script.ExpName)
                .Parameters.Append .CreateParameter("@ExpTable", adVarChar, adParamInput, 50, Script.ExpTable)
                .Parameters.Append .CreateParameter("@ExpLevel", adInteger, adParamInput, , Script.ExpLevel)
                .Parameters.Append .CreateParameter("@Xml", adLongVarChar, adParamInput, -1, strXML)
                .Parameters.Append .CreateParameter("@ExpTyp", adUnsignedTinyInt, adParamInput, , Script.ExpTyp) ' Byte
                .Execute 128 ' 0x80 dExecuteNoRecords
                Set .ActiveConnection = Nothing
            End With
            Set objCom = Nothing

        Next Script
        Result = True
    End If 'Screen Name exists

ExitNow:
On Error Resume Next
    Set ColScripts = Nothing
    DbSrc.Close
    Set DbSrc = Nothing
    DbDest.Close
    Set DbDest = Nothing

    If Not objCom Is Nothing Then
        Set objCom.ActiveConnection = Nothing
    End If
    Set objCon = Nothing
    Set objCom = Nothing
    
    SaveScreenToXML = Result
Exit Function
ErrorHappened:
On Error Resume Next
    Result = False
    Resume ExitNow
    Resume
End Function

Public Function SCR_ScreensImportKeysOverride(ByVal TableName As String, ByRef ForeignKeys As Collection) As Collection
'SA 11/9/12 - The relationship for these tables needs to be set manually because
'             they were removed to enable moving the user tables out of the database.
On Error GoTo ErrorHappened
    Dim ForeignKeysInfo As New Collection
    Dim Key As String

    Select Case TableName
        Case "SCR_ScreensLayouts", "SCR_ScreensSorts"
            Key = "ScreenID"
            ForeignKeysInfo.Add Key, "Field"
            ForeignKeysInfo.Add "SCR_Screens", "RelatedTable"
        Case "SCR_ScreensLayoutsFields", "SCR_ScreensLayoutsFormats", "SCR_ScreensLayoutsCalculations"
            Key = "LayoutID"
            ForeignKeysInfo.Add Key, "Field"
            ForeignKeysInfo.Add "SCR_ScreensLayouts", "RelatedTable"
        Case "SCR_ScreensTotalsCalculations", "SCR_ScreensTotalsFields"
            Key = "TotalID"
            ForeignKeysInfo.Add Key, "Field"
            ForeignKeysInfo.Add "SCR_ScreensTotals", "RelatedTable"
    End Select
    
    If ForeignKeysInfo.Count > 0 Then
        Set ForeignKeys = New Collection
        ForeignKeys.Add ForeignKeysInfo, Key
    End If
    
ExitNow:
On Error Resume Next
    Set SCR_ScreensImportKeysOverride = ForeignKeys
Exit Function
ErrorHappened:
    Resume ExitNow
    Resume
End Function

Public Sub SCR_ScreensExportKeysOverride(ByRef ColTable As Collection, ByRef ColForeignTable As Collection)
'SA 11/13/12 - The relationship for these tables needs to be set manually because
'              they were removed to enable moving the user tables out of the database.
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensTotals"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensCalculations"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensCondFormats"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensSorts"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_SaveScreens"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensLayouts"
    ColTable.Add "SCR_ScreensLayouts"
    ColForeignTable.Add "SCR_ScreensLayoutsFormats"
    ColTable.Add "SCR_ScreensLayouts"
    ColForeignTable.Add "SCR_ScreensLayoutsCalculations"
    ColTable.Add "SCR_ScreensLayouts"
    ColForeignTable.Add "SCR_ScreensLayoutsFields"
    ColTable.Add "SCR_Screens"
    ColForeignTable.Add "SCR_ScreensFilters"
End Sub

Public Function SCR_ScreensUserTablesList() As Collection
'List of user tables for import/export
'Note: Order is important - Child tables must be deleted before parent
    Dim TableList As New Collection
    With TableList
        .Add "SCR_TablesVersionUser"
        .Add "SCR_SaveScreens"
        .Add "SCR_ScreensSorts"
        .Add "SCR_ScreensFiltersDetails"
        .Add "SCR_ScreensFilters"
        .Add "SCR_ScreensTotalsFields"
        .Add "SCR_ScreensTotalsCalculations"
        .Add "SCR_ScreensTotals"
        .Add "SCR_ScreensLayoutsCalculations"
        .Add "SCR_ScreensCalculations"
        .Add "SCR_ScreensLayoutsFormats"
        .Add "SCR_ScreensCondFormats"
        .Add "SCR_ScreensLayoutsFields"
        .Add "SCR_ScreensLayouts"
    End With
    
    Set SCR_ScreensUserTablesList = TableList
End Function