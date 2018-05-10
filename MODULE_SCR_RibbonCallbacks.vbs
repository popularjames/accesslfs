Option Compare Database
Option Explicit

'SA 8/6/2012 - Added app name to telemetry calls

Private genUtils As New CT_ClsGeneralUtilities

Public Sub SaveAllScreens(ByVal Control As IRibbonControl)
On Error GoTo ErrorHappened
    'SA 03/22/2012 - CR2708 Changed so 1 message is displayed after all screens are saved.
    '                Not 1 message per screen saved.
    Dim i As Integer
    Dim HadError As Boolean

    For i = 1 To 20
        If Not Scr(i) Is Nothing Then
            If Not Scr(i).SaveOpenScreen Then
                HadError = True
            End If
        End If
    Next i
    
    If Not HadError Then
        MsgBox "All open screens were saved.", vbInformation, "Screens Saved"
    Else
        MsgBox "There was a problem saving screens.", vbCritical, "Error"
    End If

    'JL 1/17/2012 - Added telemetry
    Telemetry.RecordAction "Save Screens", "<P>Save all screens</P>", "Decipher Screens"

AllDone:
    Exit Sub

ErrorHappened:
    MsgBox Err.Description, vbCritical
    Resume AllDone
End Sub

Public Sub RestoreSavedScreens(ByVal Control As IRibbonControl)
On Error GoTo ErrorHappened
    
    'SA 03/22/2012 - CR2708 Restore all screens that the current user last saved
    Dim SQL As String
    Dim UserName As String
    Dim MaxDt As String
    Dim db As DAO.Database
    Dim rst As DAO.RecordSet
    Dim ClsSCR As SCR_ClsScreenData
    Dim ScrName As String
    Dim Msg As String
    UserName = Identity.UserName
    
    SysCmd acSysCmdSetStatus, "Start loading saved screens..."
    
    genUtils.SuspendLayout
    DoCmd.Hourglass True

    'Get the latest save date so old saved screens are not loaded
    MaxDt = Format(DMax("CreatedDte", "SCR_SaveScreens", "UserName=" & Chr(34) & UserName & Chr(34)), "yyyy-mm-dd")

    SQL = "SELECT ScreenID FROM SCR_SaveScreens" & _
           " WHERE UserName=" & Chr(34) & UserName & Chr(34) & _
           " AND Format(CreatedDte,'yyyy-mm-dd')=" & Chr(34) & MaxDt & Chr(34) & _
           " ORDER BY CreatedDte"
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(SQL, dbOpenSnapshot, dbForwardOnly)
    
    If rst.recordCount > 0 Then
        Do Until rst.EOF
            genUtils.SuspendLayout
            Set ClsSCR = New SCR_ClsScreenData
            ScrName = DLookup("ScreenName", "SCR_Screens", "ScreenId=" & rst!ScreenID)
            
            SysCmd acSysCmdSetStatus, "Loading screen: " & ScrName & "..."
            
            ClsSCR.CreateScreen ScrName
            If ClsSCR.NewScreen Then
                ClsSCR.GetConfig
                RunEvent "Screen Load", ClsSCR.ScreenForm.ScreenID, ClsSCR.ScreenForm.FormID
            End If
    
            Scr(ClsSCR.ScreenForm.FormID).CmdScreenLoad_Click
            Msg = Msg & ScrName & vbCrLf
            
            rst.MoveNext
        Loop
    
        Msg = "Finished loading saved screens:" & vbCrLf & Msg
    Else
        Msg = "There are no saved screens to load."
    End If

    SysCmd acSysCmdSetStatus, " "
    DoCmd.Hourglass False
    genUtils.ResumeLayout
    MsgBox Msg, vbInformation, "Screen Restore"
    
AllDone:
On Error Resume Next
    genUtils.ResumeLayout
    DoCmd.Hourglass False
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Load Error"
    Resume AllDone
    Resume
End Sub

Public Sub OpenScreen(ByVal Control As IRibbonControl)
'SA 10/22/2012 - Fixed error handling and added check for screens table version
On Error GoTo ErrorHappened
    Dim ClsSCR As SCR_ClsScreenData
    Dim TableVersionConfig As Integer
    Dim TableVersionUser As Integer
    
    DoEvents
    DoCmd.Hourglass True
    
    If Nz(Control.Tag) <> "" Then
        genUtils.SuspendLayout
        
        TableVersionConfig = Nz(DLookup("VersionNum", "SCR_TablesVersionConfig"), 0)
        TableVersionUser = Nz(DLookup("VersionNum", "SCR_TablesVersionUser"), 0)
        
        If TableVersionConfig > 0 And TableVersionConfig = TableVersionUser Then
            Set ClsSCR = New SCR_ClsScreenData
            DoCmd.Hourglass True 'after the screen is created the hourglass is lost set it back again.
            ClsSCR.CreateScreen Control.Tag
            If ClsSCR.NewScreen = True Then
                ClsSCR.GetConfig
                RunEvent "Screen Load", ClsSCR.ScreenForm.ScreenID, ClsSCR.ScreenForm.FormID
                
                'JL 1/17/2012 - Added telemetry
                Telemetry.RecordOpen "Screen", "Screen Name: " & Control.Tag, "Screen ID: " & ClsSCR.ScreenForm.ScreenID, "Decipher Screens"
            End If
        Else
            MsgBox "Your Screens tables are out of sync!" & vbCrLf & vbCrLf & _
                "Please contact a Data Analyst and ask them to make sure the Screens User Data tables have been updated to match the current version of Screens.", _
                vbCritical, "Update needed"
        End If
    Else
        MsgBox "Please select the screen to open!", vbInformation
    End If
    
ExitNow:
On Error Resume Next
    Set ClsSCR = Nothing
    DoCmd.Hourglass False
    genUtils.ResumeLayout
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Problem opening screen"
    Resume ExitNow
    Resume
End Sub

Public Sub RefreshFilterHistory(ByVal Control As IRibbonControl)
    On Error GoTo RefreshFilterHistoryError

    Dim Title As String
    Title = "Filter History"
    'bound the current screens
    CurrentDb.Execute "SCR_AddFilterHistoryBindEvents"
    CurrentDb.Execute "SCR_AddFilterHistoryUnbindevents"
        
    MsgBox "Filter History Refreshed. ", vbInformation, Title

RefreshFilterHistoryExit:
   
    Exit Sub
RefreshFilterHistoryError:
    MsgBox Err.Description, vbOKOnly, "Error " & Title
    Resume RefreshFilterHistoryExit
    
End Sub
Public Sub ShowFilterHistory(Control As IRibbonControl)
    'show filter history form
    FilterHistoryShow
    
    'JL 1/17/2012 - Added telemetry
    Telemetry.RecordOpen "Form", "Filter History", "Decipher Screens"
End Sub