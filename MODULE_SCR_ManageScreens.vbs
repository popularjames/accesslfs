'---------------------------------------------------------------------------------------
' Module    : SCR_ManageScreens
' Author    : SA
' Date      : 10/1/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Function SCR_DeleteScreenByName(ByVal ScreenName As String) As Boolean
'Delete screen based on Screen name
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim ScreenID As Long
    ScreenID = Nz(DLookup("ScreenID", "SCR_Screens", "ScreenName='" & Replace(ScreenName, "'", "''") & "'"), 0)
    
    If ScreenID > 0 Then
        Result = SCR_DeleteScreenByID(ScreenID)
    End If
     
ExitNow:
On Error Resume Next
    
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function

Public Function SCR_DeleteScreenByID(ByVal ScreenID As Long) As Boolean
'Delete screen based on ScreenID
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim db As DAO.Database
    Dim SQL As String
    Set db = CurrentDb
    
    'SCR_Screens
    SQL = "DELETE FROM SCR_Screens WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_SaveScreens
    SQL = "DELETE FROM SCR_SaveScreens WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensCalculations
    SQL = "DELETE FROM SCR_ScreensCalculations WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensCondFormats
    SQL = "DELETE FROM SCR_ScreensCondFormats WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensFilters
    SQL = "DELETE FROM SCR_ScreensFilters WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensLayouts
    SQL = "DELETE FROM SCR_ScreensLayouts WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensSorts
    SQL = "DELETE FROM SCR_ScreensSorts WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    'SCR_ScreensTotals
    SQL = "DELETE FROM SCR_ScreensTotals WHERE ScreenID=" & ScreenID
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    Result = True
ExitNow:
On Error Resume Next
    SCR_DeleteScreenByID = Result
    Set db = Nothing
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function

Public Function SCR_CleanUserTables() As Boolean
'Remove all orphaned records that are not associated with any screen
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.RecordSet
    Dim SQL As String
    Dim IdList As String
    
    SQL = "SELECT ScreenID FROM SCR_Screens"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
    
    Do Until rs.EOF
       IdList = IdList & rs!ScreenID & ","
       rs.MoveNext
    Loop
    IdList = left(IdList, Len(IdList) - 1)
    
    SQL = "DELETE FROM SCR_SaveScreens WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensCalculations WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensCondFormats WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensFilters WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensLayouts WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensSorts WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    SQL = "DELETE FROM SCR_ScreensTotals WHERE ScreenID NOT IN(" & IdList & ")"
    db.Execute SQL, dbFailOnError + dbSeeChanges
    
    Result = True
ExitNow:
On Error Resume Next
    SCR_CleanUserTables = Result
    rs.Close
    Set rs = Nothing
    Set db = Nothing
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function

Public Function SCR_DeleteDataByUserName(ByVal UserName As String) As Boolean
'Deletes all user data for a specified user - Use with caution!
On Error GoTo ErrorHappened
    Dim Result As Boolean
    Dim db As DAO.Database
    Dim SQL As String
    Set db = CurrentDb
    
    If LenB(UserName) > 0 Then
        'Escape
        UserName = Replace(UserName, "'", "''")
        
        'SCR_SaveScreens
        SQL = "DELETE FROM SCR_SaveScreens WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
    
        'SCR_ScreensCondFormats
        SQL = "DELETE FROM SCR_ScreensCondFormats WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        'SCR_ScreensFilters
        SQL = "DELETE FROM SCR_ScreensFilters WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        'SCR_ScreensLayouts
        SQL = "DELETE FROM SCR_ScreensLayouts WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        'SCR_ScreensSorts
        SQL = "DELETE FROM SCR_ScreensSorts WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        'SCR_ScreensTotals
        SQL = "DELETE FROM SCR_ScreensTotals WHERE UserName='" & UserName & "'"
        db.Execute SQL, dbFailOnError + dbSeeChanges
        
        Result = True
    Else
        Result = False
    End If
ExitNow:
On Error Resume Next
    SCR_DeleteDataByUserName = Result
    Set db = Nothing
Exit Function
ErrorHappened:
    Result = False
    Resume ExitNow
    Resume
End Function