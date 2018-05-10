Option Compare Database
Option Explicit

Public Type CnlyScreenSQLTotalsCustom
    Select As String
    From As String
    Where As String
    GroupBy As String
    OrderBy As String
    Having As String
End Type

Public Function BuildAll(ByVal TotalID As Long, ByRef frm As Form_SCR_MainScreens) As CnlyScreenSQLTotalsCustom
Dim oOut As CnlyScreenSQLTotalsCustom

With oOut
    .Select = BuildSelect(TotalID)
    .From = BuildFrom(frm)
    .Where = BuildWhere(frm)
    .GroupBy = BuildGroupBy(TotalID)
    .Having = BuildHaving(TotalID)
    .OrderBy = BuildOrderBy(TotalID)
End With


BuildAll = oOut

End Function

Public Function ToSQL(oIN As CnlyScreenSQLTotalsCustom) As String
Dim SQL As String
With oIN
    
    SQL = .Select & " "
    SQL = SQL & .From & " "
    SQL = SQL & .Where & " "
    SQL = SQL & .GroupBy & " "
    SQL = SQL & .Having & " "
    SQL = SQL & .OrderBy & " "
End With
ToSQL = SQL
End Function
Function BuildHaving(ByVal TotalID As Long) As String
Dim sHaving As String

sHaving = "" & DLookup("Having", "SCR_ScreensTotals", "TotalID=" & TotalID)

If "" & sHaving <> "" Then
    sHaving = "Having (" & sHaving & ") "
End If
BuildHaving = sHaving
End Function


Function BuildSelect(TotalID As Long) As String
On Error GoTo ErrorHappend
Dim SQL As String, db As DAO.Database, rs As DAO.RecordSet
Dim SqlOut As String

Set db = CurrentDb

SQL = "SELECT [AggregateName] & '(' & [FldName] & ')' AS FLD, "
SQL = SQL & "Alias,Ordinal, 2 as SRC "
SQL = SQL & "FROM SCR_ScreensTotalsCalculationsAggr INNER JOIN SCR_ScreensTotalsCalculations ON SCR_ScreensTotalsCalculationsAggr.AggregateID = SCR_ScreensTotalsCalculations.AggregateID "
SQL = SQL & "Where TotalID = " & TotalID & " "
SQL = SQL & "UNION ALL "

SQL = SQL & "SELECT FldName as FLD, Alias, Ordinal, 1 as SRC "
SQL = SQL & "From SCR_ScreensTotalsFields "
SQL = SQL & "Where FldType = 1 and TotalID = " & TotalID & " "
SQL = SQL & "ORDER BY SRC, Ordinal, Alias, Fld;"


Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

rs.MoveLast
rs.MoveFirst

SqlOut = "Select "

Do While Not rs.EOF
    SqlOut = SqlOut & rs.Fields("FLD")
    If "" & rs.Fields("Alias") <> "" Then
        SqlOut = SqlOut & " as " & rs.Fields("Alias")
    End If
    rs.MoveNext
    If rs.EOF Then
        SqlOut = SqlOut & " "
    Else
        SqlOut = SqlOut & ", "
    End If
Loop


BuildSelect = SqlOut

rs.Close

ExitNow:
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHappend:
    MsgBox Err.Description, vbCritical, "SCR_TotalsCustom.BuildSelect"
    Resume ExitNow
    Resume
End Function

Function BuildFrom(ByRef frm As Form_SCR_MainScreens)
Dim sFrom As String

sFrom = DLookup("PrimaryRecordSource", "SCR_Screens", "ScreenID=" & frm.ScreenID)

If "" & sFrom <> "" Then
    sFrom = "FROM " & sFrom & " "
End If
    
BuildFrom = sFrom
    
End Function

Function BuildWhere(ByRef frm As Form_SCR_MainScreens)
Dim sWhere As String
sWhere = "" & frm.BuildWhere()

If "" & sWhere <> "" Then
    sWhere = "Where (" & sWhere & ") "
End If

BuildWhere = sWhere
End Function

Function BuildGroupBy(ByVal TotalID As Long)
On Error GoTo ErrorHappend
Dim SQL As String, db As DAO.Database, rs As DAO.RecordSet
Dim SqlOut As String

Set db = CurrentDb


SQL = "SELECT FldName, Alias, Ordinal "
SQL = SQL & "From SCR_ScreensTotalsFields "
SQL = SQL & "Where FldType = 1 and TotalID = " & TotalID & " "
SQL = SQL & "ORDER BY Ordinal, FldName;"


Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot, dbReadOnly)

rs.MoveLast
rs.MoveFirst

SqlOut = "Group By "

Do While Not rs.EOF
    SqlOut = SqlOut & rs.Fields("FldName")

    rs.MoveNext
    If rs.EOF Then
        SqlOut = SqlOut & " "
    Else
        SqlOut = SqlOut & ", "
    End If
Loop


BuildGroupBy = SqlOut

rs.Close

ExitNow:
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHappend:
    MsgBox Err.Description, vbCritical, "SCR_TotalsCustom.BuildGroupBy"
    Resume ExitNow
    Resume
End Function

Function BuildOrderBy(Optional ByVal TotalID As Long = 0) As String
'SA 11/26/2012 - I assume this useless function is still here for backward compatibility
    BuildOrderBy = vbNullString
End Function