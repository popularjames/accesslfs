Option Compare Database
Option Explicit

'SA 11/26/2012 - Removed unused type CmboCfg


Public Sub PopulateListFunctions(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String
SQL = "Select DISTINCT '' as ListName, Null as FunctionID, '' as Function FROM SCR_Screens Union ALL "
SQL = SQL & "SELECT ListName, FunctionID, Function "
SQL = SQL & "FROM SCR_ScreensFunctions "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " And Enabled = True "
SQL = SQL & "ORDER BY ListName;"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub

Public Sub PopulateListTotals(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String

SQL = "Select DISTINCT Null as TotalID, '' as TheName, -2 as Sort FROM SCR_Screens Where ScreenID = " & ScreenID & " "
SQL = SQL & "Union ALL "
SQL = SQL & "SELECT TotalID, IIf([Global]=True,'GLB: ','') & [TotalName] AS TheName, Global as Sort "
SQL = SQL & "FROM SCR_ScreensTotals "
SQL = SQL & "Where ScreenID = " & ScreenID & " "
SQL = SQL & "ORDER BY Sort, TheName;"


MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub

Public Sub PopulateListReports(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String
SQL = "Select DISTINCT '' as ListName, '' as ReportID FROM SCR_Screens Union ALL "
SQL = SQL & "SELECT ListName, ReportID "
SQL = SQL & "FROM SCR_ScreensReports "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " And Enabled = True "
SQL = SQL & "ORDER BY ListName;"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub
Public Sub PopulateListFilters(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String
SQL = "Select DISTINCT '' as FilterName, '' as FilterSQL FROM SCR_Screens Union ALL "
SQL = SQL & "SELECT FilterName, FilterSQL "
SQL = SQL & "FROM SCR_ScreensFilters "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " "
SQL = SQL & "ORDER BY FilterName;"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub
Public Sub PopulateListDateFilters(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String

SQL = "Select DISTINCT '' as FieldName FROM SCR_Screens Union ALL "
SQL = SQL & "SELECT FieldName "
SQL = SQL & "FROM SCR_ScreensDateFilters "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " "
SQL = SQL & "ORDER BY FieldName"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub

Public Sub PopulateListLayouts(MyControl As Control, ScreenID As Long)
On Error GoTo PopulateListError
Dim SQL As String

SQL = "Select DISTINCT '' as LayoutID, '' as LayoutName FROM SCR_Screens Union ALL "
SQL = SQL & "SELECT LayoutID, LayoutName "
SQL = SQL & "FROM SCR_ScreensLayouts "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " "
SQL = SQL & "ORDER BY LayoutName;"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub
Public Sub PopulateListSelects(MyControl As Control, ScreenID As Long, ListLevel As Byte)
On Error GoTo PopulateListError
Dim SQL As String

SQL = "SELECT FieldName, Bound, AlternateDisplay, Width, FieldType "
SQL = SQL & "FROM SCR_ScreensListFields "
SQL = SQL & "WHERE ScreenID = " & ScreenID & " "
SQL = SQL & "   AND ListLevel = " & ListLevel & " "
SQL = SQL & "ORDER BY Sort, FieldName"

MyControl.RowSource = SQL

PopulateListExit:
    On Error Resume Next
    Exit Sub
PopulateListError:
    MsgBox Err.Description & String(2, vbCrLf) & "Error populating list for " & MyControl.Name, vbCritical, "Sloppy Code Allert"
    Resume PopulateListExit
End Sub

Public Sub PopulateListsRecordSource(Mylist As MSForms.ComboBox, PopType As Byte, Optional MyConnect As Connection)
'PopType = 1  TABLES
'PopType = 2  QUERIES 'SELECT ONLY
'PopType = 3  BOTH (TABLES + QUERIES)
'PopType = 4  SQL SERVER - YOU MUST PROVIDE A *** CONNECTION ****
'PopType = 5  ACTION QUERIES
On Error GoTo PopulateListError
Dim MyTableDef As TableDef, MyQueryDef As QueryDef, SpecDb As Database, TableRst As RecordSet
Dim tmpStr As String, X As Integer


DoCmd.Hourglass True
Mylist.Clear

If MyConnect Is Nothing Then 'List Local Tables
    Set SpecDb = CurrentDb
    If PopType = 1 Or PopType = 3 Then
        SpecDb.TableDefs.Refresh
        For Each MyTableDef In SpecDb.TableDefs
            If left(MyTableDef.Name, 4) <> "Msys" And left(MyTableDef.Name, 3) <> "Cca" Then
                Mylist.AddItem (MyTableDef.Name)
                If PopType = 3 Then
                    Mylist.Column(1, X) = "Table"
                    X = X + 1
                End If
            End If
        Next MyTableDef
    End If
    If PopType = 2 Or PopType = 3 Or PopType = 5 Then
        SpecDb.QueryDefs.Refresh
        For Each MyQueryDef In SpecDb.QueryDefs
            Select Case PopType
            Case 5 'ACTION ONLY
                If (MyQueryDef.Type And dbQAction) And MyQueryDef.Connect = "" Then Mylist.AddItem (MyQueryDef.Name)
            Case Else
                If Not (MyQueryDef.Type And dbQAction) And left(MyQueryDef.Name, 1) <> "~" Then Mylist.AddItem (MyQueryDef.Name)
            End Select
            If PopType = 3 And left(MyQueryDef.Name, 1) <> "~" Then
                Mylist.Column(1, X) = "Query"
                X = X + 1
            End If
        Next MyQueryDef
    End If
Else
    tmpStr = "SELECT so.name ObjectName, so.id ObjectId FROM sysobjects so, sysusers su WHERE so.uid = su.uid and so.type = 'U' ORDER By so.name"
    Set TableRst = MyConnect.OpenRecordSet(tmpStr, dbOpenSnapshot)
    With TableRst
        If .EOF And .BOF Then GoTo PopulateListExit
        Do Until .EOF
            Mylist.AddItem (!ObjectName)
            .MoveNext
        Loop
    End With
End If

PopulateListExit:
    On Error Resume Next
    Set MyTableDef = Nothing
    Set SpecDb = Nothing
    DoCmd.Hourglass False
    Exit Sub
    
PopulateListError:
    MsgBox Err.Description
    Resume PopulateListExit
End Sub