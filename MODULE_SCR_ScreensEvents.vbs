Option Compare Database
Option Explicit


Public Function RunEvent(ByVal EventType As String, ByVal ScreenID As Long, ByVal FormID As Long, Optional ByVal vAny As Variant) As Boolean
'SA 05/21/2012 - CR2131/2132 Changed from sub to function that returns true if an event is found
On Error GoTo ErrorHappened
    Dim HasEvent As Boolean
    Dim db As DAO.Database
    Dim rst As DAO.RecordSet
    Dim SQL As String
    
    SQL = "SELECT EventType,Function FROM SCR_ScreensEvents WHERE ScreenID=" & _
            ScreenID & " AND EventType='" & Replace(EventType, "'", "''") & "' ORDER BY Seq, Function"
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    With rst
        If .recordCount > 0 Then
            HasEvent = True
            Do Until .EOF
                Select Case UCase(.Fields("EventType"))
                    Case "REPORT RUN", "KEY PRESSED", "REPORT POST-RUN"
                        Application.Run .Fields("Function"), FormID, vAny
                    Case Else
                        Application.Run .Fields("Function"), FormID
                End Select
                .MoveNext
            Loop
        Else
            HasEvent = False
        End If
    End With
    
ExitNow:
    On Error Resume Next
    RunEvent = HasEvent
    rst.Close
    db.Close
    Set rst = Nothing
    Set db = Nothing
Exit Function
ErrorHappened:
    SQL = Err.Description & vbCrLf & vbCrLf & _
        "ScreenID: " & ScreenID & vbCrLf & _
        "EventType: " & EventType
    MsgBox SQL, vbCritical, "Error Running Screen Event"
    Resume ExitNow
End Function