Option Compare Database
Option Explicit

Private Const ClassName As String = "mod_Cleanup_Linked_Tables"

Private cdctLinkedTables As Scripting.Dictionary


Public Function GenerateListOfUnusedLinkedTables()
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sCurTbl As String
Dim dctResults As Scripting.Dictionary
Dim sListOfObjectsFoundin As String
Dim sarySqlTblsToQuery(3) As String
Dim sCurSqlTbl As String
Dim iSqlTblIdx As Integer
Dim dctNotUsed As Scripting.Dictionary
Dim sResultsNote As String
Dim sSql As String
Dim sFoundString As String

    strProcName = ClassName & ".GenerateListOfUnusedLinkedTables"
    
    sarySqlTblsToQuery(0) = "General_Tabs"
    sarySqlTblsToQuery(1) = "General_Search"
    sarySqlTblsToQuery(2) = "Report_Parameter"
    sarySqlTblsToQuery(3) = "AuditClm_Rationale_Template"
    
    
    If IsTable("tbl_Not_Used_Tables") = False Then
        Call CreateOurTable
    End If
    
    ' Go through the linked tables found in the table.. that's all we are currently concerned with
    ' as the dynamically linked tables (reporting section) aren't refreshed with the version control system
    ' that is taking forever to run when relinking the tables
    
    sSql = "SELECT L.*, NU.* FROM Link_Table_Config L LEFT JOIN tbl_Not_Used_Tables NU ON L.Table = NU.LnkedTableName WHERE NZ(NU.Found, False) <> True and L.Location = 'CMSPROD' "
    
    Set oDb = CurrentDb()
    Set oRs = oDb.OpenRecordSet(sSql)
    
    Set dctResults = New Scripting.Dictionary
    Set dctNotUsed = New Scripting.Dictionary
    
    '' Ok, we need to do this differently for performance..
    '' get all of the tablenames in memory:
    Set cdctLinkedTables = New Scripting.Dictionary
    
    
    While Not oRs.EOF
        sFoundString = ""
        sCurTbl = oRs("Table").Value
        If cdctLinkedTables.Exists(sCurTbl) = False Then
            sFoundString = IIf(Nz(oRs("MODULES"), False), "M", "") & IIf(Nz(oRs("FORMS"), False), "F", "") & IIf(Nz(oRs("REPORTS"), False), "R", "") & _
                IIf(Nz(oRs("QUERIES"), False), "Q", "") & IIf(Nz(oRs("SQLTABLES"), False), "S", "")
            cdctLinkedTables.Add sCurTbl, sFoundString
        End If
        oRs.MoveNext
    Wend
    
    Set oRs = Nothing

    '' Now, call each of the items for ALL of our linked tables
    '' starting with Modules which should be the quickest
    '' and will probably find the most used
    '' then forms
    '' then reports
    '' then Sql tables

        
            ' so, check to see if it's used in code
'    Call FindModulesUsingObject
            ' check to see if it's a direct control source for a form or control
'    Call FindFormsUsingObject
        
'    Call FindReportsUsingObject
            
            ' what else? How about the SQL Server tables:
                '            General_Tabs
                '            General_Search
                '            Report_Parameter
                '            AuditClm_Rationale_Template
        For iSqlTblIdx = 1 To UBound(sarySqlTblsToQuery)
            sCurSqlTbl = sarySqlTblsToQuery(iSqlTblIdx)
            Call FindSqlTblsSettingRecordSources(sCurSqlTbl)
        Next
        

        MsgBox "Finished looking for unused linked tables!"
    
Block_Exit:
    Exit Function
    
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


''' ############################################################
''' ############################################################
''' ############################################################
''' Finds modules using the keyword in one or more of the lines of code
'''
Private Function FindSqlTblsSettingRecordSources(ByVal sSqlTbl As String) As Integer
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim sSql As String
Dim oFld As ADODB.Field
Dim sFldVal As String
Dim vKey As Variant
Dim bFound As Boolean

    strProcName = ClassName & ".FindSqlTblsSettingRecordSources"

    

    sSql = "SELECT * FROM " & sSqlTbl
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set oRs = .ExecuteRS
        If .GotData = False Then
            Stop
        End If
    End With
    
    While Not oRs.EOF
Debug.Print oRs("RowId").Value
If left(oRs("TabName").Value, 2) = "R0" Then
    GoTo NextRcd
End If
        For Each oFld In oRs.Fields
            sFldVal = CStr(Nz(oRs(oFld.Name).Value, ""))
            
            For Each vKey In cdctLinkedTables.Keys
                  bFound = False  ' reset
                  
                  ' skip any we've already found:
                  If InStr(1, cdctLinkedTables.Item(vKey), "S", vbTextCompare) < 1 Then
                      If InStr(1, sFldVal, CStr(vKey), vbTextCompare) > 0 Then
                          bFound = True
                          FindSqlTblsSettingRecordSources = FindSqlTblsSettingRecordSources + 1
                          LogMessage sSqlTbl & " contains: (" & CStr(vKey) & ")", strProcName
                          cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "S" ' for QUERIES
                      End If
                  End If
                  
                  Call UpdateStatsTable(CStr(vKey), "SQLTABLES", bFound, sSqlTbl)
              Next
            
        Next
NextRcd:
        oRs.MoveNext
    Wend
    

Block_Exit:
    Set oAdo = Nothing
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
End Function


Private Sub AddToDict(dct As Scripting.Dictionary, sKey As String, sValToAdd As String)
    If dct.Exists(sKey) = True Then
        dct.Item(sKey) = dct.Item(sKey) & "," & sValToAdd
    Else
        dct.Item(sKey) = sValToAdd
    End If
End Sub




''' ############################################################
''' ############################################################
''' ############################################################
''' Finds macros using the keyword ... never mind - never finished
'''
Private Function FindMacrosUsingObject(ByVal strObjName As String, Optional bLookInControls As Boolean = False, _
    Optional bStopAtFirstFound As Boolean = True, Optional sFoundInList As String) As Integer
Dim objAO As Object
Dim strName As String
Dim strSQL As String

    sFoundInList = ""   ' make sure we zero it out first

    On Error Resume Next
    For Each objAO In Application.CurrentProject.AllMacros
        DoCmd.OpenForm objAO.Name, acDesign, , , , acHidden
'        Set objForm = Application.Forms(objAO.Name)

        ListControlProps objAO

'        strSQL = objForm.RecordSource

        If InStr(1, strSQL, strObjName, vbTextCompare) > 0 Then
            FindMacrosUsingObject = FindMacrosUsingObject + 1
                sFoundInList = sFoundInList & objAO.Name & ","
                If bStopAtFirstFound = True Then GoTo Block_Exit
        End If

'        Unload objForm

    Next

Block_Exit:
    Set objAO = Nothing
End Function


''' ############################################################
''' ############################################################
''' ############################################################
''' Finds modules using the keyword in one or more of the lines of code
'''
Private Function FindModulesUsingObject() As Integer
Dim strProcName As String
Dim objModule As Module
Dim objOA As AccessObject
Dim strCode As String
Dim bFound As Boolean
    

Dim vKey As Variant
    strProcName = ClassName & ".FindModulesUsingObject"


    On Error Resume Next
    For Each objOA In Application.CurrentProject.AllModules
        DoCmd.OpenModule objOA.Name
        Set objModule = Application.Modules(objOA.Name)
        strCode = objModule.Lines(1, objModule.CountOfLines)
'        Unload objOA

        For Each vKey In cdctLinkedTables.Keys
            bFound = False  ' reset
            
            ' skip any we've already found:
            If InStr(1, cdctLinkedTables.Item(vKey), "M", vbTextCompare) < 1 Then
                If InStr(1, strCode, CStr(vKey), vbTextCompare) > 0 Then
                    bFound = True
                    FindModulesUsingObject = FindModulesUsingObject + 1
                    LogMessage objModule.Name & " contains: (" & CStr(vKey) & ")", strProcName
                    cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "M" ' for Module
                End If
            End If
            
            Call UpdateStatsTable(CStr(vKey), "MODULES", bFound, objOA.Name)
        Next

        DoCmd.Close acModule, objOA.Name
        Set objOA = Nothing
    Next
    
Block_Exit:
    Set objOA = Nothing
End Function



''' ############################################################
''' ############################################################
''' ############################################################
''' Finds reports using the keyword passed in the SQL source for that report
'''
Private Function FindReportsUsingObject() As Integer
Dim strProcName As String
Dim objReport As Report
Dim objAccessObj As AccessObject
Dim strNameLike As String
Dim strSQL As String
Dim vKey As Variant
Dim bFound As Boolean

    strProcName = ClassName & ".FindReportsUsingObject"
    


    On Error Resume Next
    For Each objAccessObj In Application.CurrentProject.AllReports
        If left(objAccessObj.Name, 6) <> "LEGACY" Then
            DoCmd.OpenReport objAccessObj.Name, acViewPreview
            Set objReport = Reports(objAccessObj.Name)
            strSQL = objReport.RecordSource
    
    
            For Each vKey In cdctLinkedTables.Keys
                bFound = False  ' reset
                
                ' skip any we've already found:
                If InStr(1, cdctLinkedTables.Item(vKey), "R", vbTextCompare) < 1 Then
                    If InStr(1, strSQL, CStr(vKey), vbTextCompare) > 0 Then
                        bFound = True
                        FindReportsUsingObject = FindReportsUsingObject + 1
                        LogMessage objAccessObj.Name & " contains: (" & CStr(vKey) & ")", strProcName
                        cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "R" ' for Reports
                    End If
                End If
                
                Call UpdateStatsTable(CStr(vKey), "REPORTS", bFound, objAccessObj.Name)
            Next
    
'            If InStr(1, strSQL, strObjName, vbTextCompare) > 0 Then
'                FindReportsUsingObject = FindReportsUsingObject + 1
'                sFoundInList = sFoundInList & objReport.Name & ","
'                If bStopAtFirstFound = True Then GoTo Block_Exit
'            End If
            DoCmd.Close acReport, objAccessObj.Name
        End If
    Next
Block_Exit:
    Set objReport = Nothing
    Set objAccessObj = Nothing

End Function



''' ############################################################
''' ############################################################
''' ############################################################
''' Finds queries where the keyword is found in the SQL
' Note to self:
' If a query uses several other queries that use this table, we won't count it (yet)
Private Function FindQuerysUsingObject() As Integer
'Dim objTable As TableDef
Dim objQuery As QueryDef
Dim strNameLike As String
Dim strSQL As String
Dim strProcName As String
Dim vKey As Variant
Dim bFound As Boolean


    strProcName = ClassName & ".FindQueriesUsingObject"
    
    
    On Error Resume Next
    For Each objQuery In CurrentDb().QueryDefs
        strSQL = objQuery.SQL

        For Each vKey In cdctLinkedTables.Keys
            bFound = False  ' reset
            
            ' skip any we've already found:
            If InStr(1, cdctLinkedTables.Item(vKey), "Q", vbTextCompare) < 1 Then
                If InStr(1, strSQL, CStr(vKey), vbTextCompare) > 0 Then
                    bFound = True
                    FindQuerysUsingObject = FindQuerysUsingObject + 1
                    LogMessage objQuery.Name & " contains: (" & CStr(vKey) & ")", strProcName
                    cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "M" ' for QUERIES
                End If
            End If
            
            Call UpdateStatsTable(CStr(vKey), "QUERIES", bFound, objQuery.Name)
        Next
        

    Next
    
Block_Exit:
    Set objQuery = Nothing

End Function



''' ############################################################
''' ############################################################
''' ############################################################
''' Looks through forms using the passed keyword in the recordsource
'''
Private Function FindFormsUsingObject() As Integer
Dim strProcName As String
Dim objForm As Form
Dim objAO As Object
Dim strName As String
Dim strSQL As String
Dim oCtl As Control
Dim vKey As Variant
Dim bFound As Boolean

    strProcName = ClassName & ".FindFormsUsingObject"
    
    On Error Resume Next
    
    For Each objAO In Application.CurrentProject.AllForms
        DoCmd.OpenForm objAO.Name, acDesign, , , , acHidden
        Set objForm = Application.Forms(objAO.Name)

        strSQL = objForm.RecordSource

        For Each vKey In cdctLinkedTables.Keys
            bFound = False  ' reset
            
            ' skip any we've already found:
            If InStr(1, cdctLinkedTables.Item(vKey), "F", vbTextCompare) < 1 Then
                If InStr(1, strSQL, CStr(vKey), vbTextCompare) > 0 Then
                    bFound = True
                    FindFormsUsingObject = FindFormsUsingObject + 1
                    LogMessage objForm.Name & " contains: (" & CStr(vKey) & ")", strProcName
                    cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "F" ' for Form
'Stop
                Else
                        ' Not in the RecordSource, check the controls to see if they use it
                        ' may also have to check the forms module in the event that the other function doesn't
                        '  look at form modules
                    For Each oCtl In objForm.Controls
                        If IsProperty(oCtl, "RowSource") = True Then
                            If InStr(1, oCtl.Properties("RowSource").Value, CStr(vKey), vbTextCompare) > 0 Then
                                FindFormsUsingObject = FindFormsUsingObject + 1
                                bFound = True
                                FindFormsUsingObject = FindFormsUsingObject + 1
                                LogMessage objForm.Name & " contains: (" & CStr(vKey) & ")", strProcName
                                cdctLinkedTables.Item(vKey) = cdctLinkedTables.Item(vKey) & "F" ' for Form
                                Exit For
                            End If
                        End If
                    Next
                End If
                
            End If
            
NextLnkTbl:
            Call UpdateStatsTable(CStr(vKey), "FORMS", bFound, objAO.Name)
        Next


        DoCmd.Close acForm, objForm.Name
'        Unload objForm

    Next
    
Block_Exit:
    Set oCtl = Nothing
    Set objForm = Nothing
End Function


Private Sub UpdateStatsTable(sLinkedTableName As String, sCurrentObjectType As String, bFound As Boolean, sFoundInName As String)
On Error GoTo Block_Err
Dim strProcName As String
Static oDb As DAO.Database
Dim oRs As DAO.RecordSet
Dim sSql As String
Dim sUpdateField As String

    strProcName = ClassName & ".AddResultToTbl"
    
    sUpdateField = sCurrentObjectType   ' not going to worry about checking to see if we coded this right
    
    
    If oDb Is Nothing Then Set oDb = CurrentDb()
    sSql = "SELECT * FROM tbl_Not_Used_tables WHERE LnkedTableName = '" & sLinkedTableName & "' "
    
    Set oRs = oDb.OpenRecordSet(sSql)
    
    If oRs.EOF And oRs.BOF Then
'        Stop ' need to add it:
        oRs.AddNew
        oRs("lnkedTableName") = sLinkedTableName
        oRs.Update
        oRs.Requery
    End If
    ' update the correct field:
    oRs.Edit
    oRs.Fields(sUpdateField) = -1
    If bFound = True Then
        oRs.Fields("Found") = -1
        oRs.Fields("FoundInName") = sFoundInName
'    Else
'        oRs.Fields("Found") = 0
    End If
    oRs.Update
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Public Sub CreateOurTable()
Dim oDb As DAO.Database
Dim oTDef As DAO.TableDef
Dim oFld As DAO.Field
Dim oIdx As DAO.index

'' Fields: LinkedTableId -= autonumber PK
''  lnkedTableName text 255
'' dateAdded DateTime -= Default = Now()
'' Modules Yes no
'' Forms Yes no
'' Reports Yes no
'' Queries Yes no
'' SqlTables Yes no
'' Found Yes No Default False
'' FoundInName Text 255

    Set oDb = CurrentDb()
    Set oTDef = New TableDef
    With oTDef
        .Name = "tbl_Not_Used_Tables"
        Set oFld = .CreateField("LinkedTableId", dbLong)
        oFld.Attributes = dbAutoIncrField
        
        .Fields.Append oFld
        
        Set oIdx = .CreateIndex("pk_AutoId")
        oIdx.Clustered = True
        oIdx.Primary = True
        Set oFld = oIdx.CreateField("LinkedTableId")
        
        oIdx.Fields.Append oFld
        .Indexes.Append oIdx
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("lnkedTableName", dbText, 255)
        oFld.AllowZeroLength = False
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("DateAdded", dbDate)
'        oFld.AllowZeroLength = False
        oFld.DefaultValue = "Now()"
        oFld.Required = True
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("Modules", dbBoolean)
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("Forms", dbBoolean)
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("Reports", dbBoolean)
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("Queries", dbBoolean)
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("SqlTables", dbBoolean)
        .Fields.Append oFld
        
        '' Now just the rest of the fields:
        Set oFld = .CreateField("Found", dbBoolean)
        oFld.Required = True
        oFld.DefaultValue = 0
        .Fields.Append oFld

        '' Now just the rest of the fields:
        Set oFld = .CreateField("FoundInName", dbText, 255)
        .Fields.Append oFld
        
    
    End With
    oDb.TableDefs.Append oTDef
    oDb.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    
End Sub


Public Function ListMyStuff()
Dim fldLoop As Field2
Dim oDb As DAO.Database
Dim oProp As DAO.Property
Dim oTDef As DAO.TableDef

    Set oDb = CurrentDb()
    Set oTDef = oDb.TableDefs("tbl_Not_Used_Tables")


    For Each fldLoop In oTDef.Fields
        Debug.Print " " & fldLoop.Name & " = " & fldLoop.Attributes
        For Each oProp In fldLoop.Properties
            Debug.Print "    P: " & oProp.Name & " = " & oProp.Value
        Next
    Next fldLoop

End Function