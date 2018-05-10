Option Compare Database
Option Explicit

Public Function MUT_GetScreenDateFields(DbPath As String, ScreenID As Long) As String
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    Dim tmpStr As String
    
    SQL = "SELECT FieldName FROM [" & DbPath & "].CnlyScreensDateFilters WHERE ScreenID = " & ScreenID & " ORDER BY SortID"
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot)
    If Not (rs.EOF Or rs.BOF) Then
        Do Until rs.EOF
            tmpStr = tmpStr & rs!FieldName & ";"
            rs.MoveNext
        Loop
        tmpStr = left(tmpStr, Len(tmpStr) - 1)
    End If
    MUT_GetScreenDateFields = tmpStr
End Function