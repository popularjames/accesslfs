Option Compare Database
Option Explicit
' HC 5/2010 removed link constants
Public Function AutoSync(ByRef strDBPath As String) As Boolean
On Error GoTo ErrorHappened
Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
Dim RemoteDB As String
Dim ErrorCT As Integer
Dim ClsImp As CT_ClsImport
 
If ClsImp Is Nothing Then
    Set ClsImp = New CT_ClsImport
End If
 
ErrorCT = 0
SQL = "Select * From SCR_ScreensVersionsUtilities WHERE UtiltiyID <> 6"
 
RemoteDB = Nz(DLookup("Value", "CT_Options", "OptionName = 'ProductionDb'"), "")
If RemoteDB <> "" Then
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF
            If Not ClsImp.RunUtility(rst!UtiltiyID, RemoteDB) Then
                ErrorCT = ErrorCT + 1
            End If
            rst.MoveNext
        Loop
    End If
    
    If ErrorCT <> 0 Then
        strDBPath = "Error Syncing to Production Database"
    Else
        strDBPath = RemoteDB
        AutoSync = True
    End If
Else
    strDBPath = "Error: Production Database not specified"
    AutoSync = False
End If
 
ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set rst = Nothing
    Exit Function
 
ErrorHappened:
    strDBPath = "Error Syncing to Production Database"
    AutoSync = False
    Resume ExitNow
 
End Function

Public Function GetRemoteUsers(ByRef Msg As String)
On Error GoTo ErrorHappened
Dim SQL As String, db As DAO.Database, rst As DAO.RecordSet
Dim StProductionDb As String

StProductionDb = Nz(DLookup("Value", "CT_Options", "OptionName = 'ProductionDb'"), "")
SQL = "SELECT empCurrentlyLoggedIn, UDate, Computer FROM [" & StProductionDb & "].CT_CurrentlyLoggedIn;"

If StProductionDb <> "" Then
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot)
    If Not (rst.EOF And rst.BOF) Then
        rst.MoveFirst
        Do Until rst.EOF
            Msg = Msg & vbTab & rst!empCurrentlyLoggedIn & vbTab & vbTab & rst!UDate & vbTab & rst!Computer & vbCrLf
            rst.MoveNext
        Loop
    End If
End If

ExitNow:
    On Error Resume Next
    Set db = Nothing
    Set rst = Nothing
    Exit Function

ErrorHappened:
    Resume ExitNow

End Function

Public Function Deploy_Link(ByRef StError As Boolean)
On Error GoTo ErrorHappened
' HC 5/2010 removed the update to set location, should be done with Deploy_linkLocation
Deploy_UnLinkTables
If Deploy_LinkLocation("Production") = True Then
    'SavedLocationSet "Production"
    SetApplicationTitle
Else
    GoTo ExitNow
End If

ExitNow:
    Exit Function

ErrorHappened:
    StError = True
    Resume ExitNow

End Function

Private Sub Deploy_UnLinkTables()
#If ccCFG = 1 Then
' HC 5/2010 replaced with call to config links
    Dim mvCfgLinks As Form_CFG_CfgLink
    
    Set mvCfgLinks = New Form_CFG_CfgLink
    mvCfgLinks.visible = False
    mvCfgLinks.UnLinkTables (CurrentDb.Name)
    Set mvCfgLinks = Nothing
    
    SetApplicationTitle
    DoEvents
    
    CurrentDb.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    DoEvents
#End If
End Sub

Private Function Deploy_LinkLocation(strLocation As String) As Boolean
#If ccCFG = 1 Then
On Error GoTo CATCH
    Dim db 'As DAO.Database
    Dim rst 'As DAO.Recordset
    Dim mvCfgLinks As Form_CFG_CfgLink
    
    Set db = CurrentDb
    Set rst = db.OpenRecordSet("Select * FROM CFG_CfgLink Where Location = " & Chr(34) & strLocation & Chr(34), 8) 'dbOpenForwardOnly
    
    If rst.BOF And rst.EOF Then
        MsgBox "Specified location does not exist.", vbCritical, "Linking Location: " & strLocation
    Else
        Set mvCfgLinks = New Form_CFG_CfgLink
        mvCfgLinks.visible = False
        mvCfgLinks.LinkByLocation (strLocation)
        SetApplicationTitle
        DoEvents
    End If
    
    Deploy_LinkLocation = True
Done:
On Error Resume Next
    rst.Close
    Set mvCfgLinks = Nothing
    Set rst = Nothing
    Set db = Nothing
    Exit Function

CATCH:
    MsgBox Err.Description, vbCritical
    Deploy_LinkLocation = False
    Resume Done
#End If
End Function