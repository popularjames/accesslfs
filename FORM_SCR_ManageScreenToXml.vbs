Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'SA 11/14/2012 - Added telemetry and changed the way the list boxes are loaded.

Private Sub Form_Load()
    Telemetry.RecordOpen "Form", Me.Name
    SetSavedRowSource
End Sub

Private Sub CmdBrowseDest_Click()
    Dim StDB As String

    StDB = "" & Me.txtDestDB
    ' HC 6/2010 - changed extension to accdb
    StDB = FileDialog(0, "Select Database", Me.hwnd, "", "Access Database (*.accdb)" & Chr(0) & "*.accdb" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0), StDB, "accdb")
    If StDB <> "" Then
        Me.txtDestDB = StDB
    End If
    txtDestDB_AfterUpdate
End Sub

Private Sub cmdBrowseSrc_Click()
    Dim StDB As String

    StDB = "" & Me.txtSourceDB
    ' HC 6/2010 - changed extension to accdb
    StDB = FileDialog(0, "Select Database", Me.hwnd, "", "Access Database (*.accdb)" & Chr(0) & "*.accdb" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0), StDB, "accdb")
    If StDB <> "" Then
        Me.txtSourceDB = StDB
    End If
    SetLiveRowSource
End Sub

Private Sub cmdDelete_Click()
'Delete seleted screen from xml table
'SA 8/6/2012 - Reworked to use selected db
On Error GoTo ErrorHappened
    Dim db As DAO.Database
    
    If LenB(Me.lstSavedScreens.Value) > 0 Then
        If MsgBox("Are you sure you want to delete " & Me.lstSavedScreens.Value & "?", vbQuestion + vbYesNo, "Confirm delete") = vbYes Then
            If LenB(Nz(Me.txtDestDB, vbNullString)) = 0 Then
                Set db = CurrentDb
            Else
                Set db = DBEngine.Workspaces(0).OpenDatabase(Me.txtDestDB)
            End If
            
            db.Execute "DELETE FROM SCR_ScreensXML WHERE ScreenName='" & Me.lstSavedScreens.Value & "'", dbFailOnError
            
            DoEvents
            Me.lstSavedScreens.Requery
        End If
    Else
        MsgBox "Please select a screen to delete.", vbInformation, "Select screen"
    End If

ExitNow:
On Error Resume Next
    Set db = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Error deleting screen"
    Resume ExitNow
End Sub

Private Sub CmdRename_Click()
'SA 8/6/2012 - Refactored
On Error GoTo ErrorHappened
    Dim stName As String
    Dim StNewName As String
    Dim stXml As String
    Dim stNewXml As String
    Dim db As Database
    
    If LenB(Me.lstSavedScreens.Value) > 0 Then
        If LenB(Nz(Me.txtDestDB, vbNullString)) = 0 Then
            Set db = CurrentDb
        Else
            Set db = DBEngine.Workspaces(0).OpenDatabase(Me.txtDestDB)
        End If
    
        stName = Nz(Me.lstSavedScreens.Value, vbNullString)
        StNewName = InputBox("Please enter a new name for '" & stName & "'", "Rename saved screen", stName)
        If StNewName <> vbNullString And StNewName <> stName Then
            'The actual screen name is stored in the XML field.  Update it using a string replace.
            stXml = ExDLookup(db, "XML", "SCR_ScreensXML", "ScreenName='" & Replace(stName, "'", "''") & "'")
            stNewXml = Replace$(stXml, "'" & stName & "'", "'" & StNewName & "'")
            
            db.Execute "UPDATE SCR_ScreensXml SET ScreenName='" & Replace(StNewName, "'", "''") & "' WHERE ScreenName='" & Replace(stName, "'", "''") & "'"
            
            'Apply the new screen name to the xml.
            db.Execute "UPDATE SCR_ScreensXml SET XML='" & Replace(stNewXml, "'", "''") & "' WHERE ScreenName='" & Replace(StNewName, "'", "''") & "' AND TableName='SCR_Screens'"
            Me.lstSavedScreens.Requery
        End If
    Else
        MsgBox "You must select a saved screen to rename.", vbInformation, "Rename saved screen"
    End If
ExitNow:
On Error Resume Next
    Set db = Nothing
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, Err.Source
    Resume ExitNow
End Sub

Private Sub cmdToLive_Click()
'Restore the selected screen from XML
'SA 8/6/2012 - Added check for selected screen
On Error GoTo eTrap
    Dim dbDataBase As Database
    
    DoCmd.Hourglass True
    
    If "" & Me.txtSourceDB = "" Then
        Set dbDataBase = CurrentDb
    Else
        Set dbDataBase = DBEngine.Workspaces(0).OpenDatabase(Me.txtSourceDB)
    End If

    If LenB(Me.lstSavedScreens.Value) > 0 Then
        'Verify that a screen with that name does not already exist.
        If Nz(ExDLookup(dbDataBase, "ScreenName", "SCR_Screens", "[ScreenName]  = '" & Me.lstSavedScreens.Value & "'"), "") = "" Then
            RestoreScreenFromXML "" & Me.txtDestDB, "" & Me.txtSourceDB, Me.lstSavedScreens.Value
            MsgBox "Screen move complete.", vbInformation, "Complete"
            SetLiveRowSource
        Else
            MsgBox "A live screen with that name already exists.", vbInformation, "Screen exsits"
        End If
    Else
        MsgBox "Please select a screen to move.", vbInformation, "Select screen"
    End If
    
eSuccess:
On Error Resume Next
    Set dbDataBase = Nothing
    DoCmd.Hourglass False
Exit Sub
eTrap:
    MsgBox Err.Description, vbCritical, "Restoring Screen from XML"
    Resume eSuccess
    Resume
End Sub

Private Sub cmdToSaved_Click()
'Persist the selected screen to XML
'SA 8/6/2012 - Added check for selected screen and msgbox so saved screens list is refreshed
On Error GoTo ErrorHappened
    Dim dbDataBase As Database
    
    DoCmd.Hourglass True

    If LenB(Nz(Me.txtDestDB, vbNullString)) = 0 Then
        Set dbDataBase = CurrentDb
    Else
        Set dbDataBase = DBEngine.Workspaces(0).OpenDatabase(Me.txtDestDB)
    End If
    
    'Verify that a screen with that name does not already exist in the XML store.
    If LenB(lstLiveScreens.Value) > 0 Then
        If Nz(ExDLookup(dbDataBase, "ScreenName", "SCR_ScreensXML", "[ScreenName]  = '" & Me.lstLiveScreens.Value & "'"), vbNullString) = vbNullString Then
            
            If SaveScreenToXML(Nz(Me.txtSourceDB, vbNullString), Nz(Me.txtDestDB, vbNullString), Me.lstLiveScreens.Value) Then
                MsgBox "The selected screen was exported into SCR_ScreensXML", vbInformation, "Export Complete"
                SetSavedRowSource
            Else
                MsgBox "There was an error exporting the screen. You may need to restart Decipher.", vbExclamation, "Export failed"
            End If
        Else
            MsgBox "A screen with that name has already been saved.", vbInformation, "Screen already saved"
        End If
    Else
        MsgBox "Please select a screen to save.", vbInformation, "Nothing selected"
    End If
    
ExitNow:
On Error Resume Next
    Set dbDataBase = Nothing
    DoCmd.Hourglass False
Exit Sub
ErrorHappened:
    MsgBox Err.Description, vbCritical, "Saving Live Screen"
    Resume ExitNow
End Sub

Private Sub txtDestDB_AfterUpdate()
    SetSavedRowSource
End Sub

Private Sub SetSavedRowSource()
'Set row source for saved screens listbox
On Error GoTo ErrorHappened
    Dim strSQL As String
    strSQL = "SELECT DISTINCT ScreenName FROM SCR_ScreensXML"
    If LenB(Nz(txtDestDB, vbNullString)) > 0 Then
        If LenB(Dir(txtDestDB)) > 0 Then
            Me.cmdToLive.Enabled = True
            Me.cmdToSaved.Enabled = True
            strSQL = strSQL & " IN " & Chr$(34) & txtDestDB & Chr$(34)
        Else
            MsgBox "Specified database does not exists.", vbCritical, "Manage Screens to XML"
            Me.cmdToLive.Enabled = False
            Me.cmdToSaved.Enabled = False
            strSQL = strSQL & " WHERE 1 = 0;"
        End If
    Else
        'Use local db
        Me.cmdToLive.Enabled = True
        Me.cmdToSaved.Enabled = True
    End If
ExitNow:
On Error Resume Next
    Me.lstSavedScreens.RowSource = vbNullString
    Me.lstSavedScreens.RowSource = strSQL
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub

Private Sub txtSourceDB_AfterUpdate()
    SetLiveRowSource
End Sub

Private Sub SetLiveRowSource()
'Set row source for live screens listbox
On Error GoTo ErrorHappened
    Dim strSQL As String
    strSQL = "SELECT ScreenName FROM SCR_Screens"
    If LenB(Nz(txtSourceDB, vbNullString)) > 0 Then
        If LenB(Dir(txtSourceDB)) > 0 Then
            Me.cmdToLive.Enabled = True
            Me.cmdToSaved.Enabled = True
            strSQL = strSQL & " IN " & Chr$(34) & txtSourceDB & Chr$(34)
        Else
            MsgBox "Specified database does not exists.", vbCritical, "Manage Screens to XML"
            Me.cmdToLive.Enabled = False
            Me.cmdToSaved.Enabled = False
            strSQL = strSQL & " WHERE 1 = 0;"
        End If
    Else
        'Use local db
        Me.cmdToLive.Enabled = True
        Me.cmdToSaved.Enabled = True
    End If
ExitNow:
On Error Resume Next
    Me.lstLiveScreens.RowSource = vbNullString
    Me.lstLiveScreens.RowSource = strSQL
Exit Sub
ErrorHappened:
    Resume ExitNow
    Resume
End Sub
