Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents ClsHelp As CT_ClsCnlyHelp
Attribute ClsHelp.VB_VarHelpID = -1
Private MvTblVersion As Single

Private Sub ClsHelp_StatusMessage(ByVal Src As String, ByVal Status As String, ByVal Msg As String)
    StatusMessage Src, Status, Msg
End Sub

Private Sub CmbAppName_AfterUpdate()
On Error GoTo ErrorHappened
    
    If Nz(CmbAppName.SelText, "") <> "" Then
    
        If ClsHelp.DecipherVersion = "Nill" Then
            StatusMessage "Configuration", "Error", "Unable to find Decipher version from SCR_ScreensVersions table"
            StatusMessage "Configuration", "Error", "Verify that the SCR_ScreensVersions table has its description property value set"
            CmdExport.Enabled = False
            CmdUpdate.Enabled = False
            Exit Sub
        End If
        
        Dim db As DAO.Database
        Dim rs As DAO.RecordSet
        Dim SQL As String
    
        Set db = CurrentDb
        
        SQL = "Select AppName, AppPrefix, HelpTableName, HelpFilePath, HelpFileName "
        SQL = SQL & " From CT_HelpConfig "
        SQL = SQL & " Where AppName = '" & CmbAppName.SelText & "'"
        
        Set rs = db.OpenRecordSet(SQL, dbOpenSnapshot + dbReadOnly)
                      
        If Nz(rs!HelpTableName, "") = "" Then
            StatusMessage "Configuration", "Error", "Help Table name is required in order to Export/Update control properties. "
            StatusMessage "Configuration", "Error", "Verify that the help table name is specified for the selected application in CT_HelpConfig table."
            CmdExport.Enabled = False
            CmdUpdate.Enabled = False
            Exit Sub
        Else
            CmdExport.Enabled = True
            CmdUpdate.Enabled = True
            TxtTableName.Value = rs!HelpTableName
            ClsHelp.AppName = rs!AppName
            ClsHelp.HelpTable = rs!HelpTableName
        End If
        
        If Nz(rs!AppPrefix, "") = "" Or rs!AppPrefix = "-No-" Then

            StatusMessage "Configuration", "Warning", "The selected application doesn't have a valid form object name prefix. "
            StatusMessage "Configuration", "Warning", "All form objects that don't have a predefined prefix in their name will be considered for Export/Update. "
               
        End If
        
        TxtHelpFile.Value = Nz(rs!HelpFileName, "")
        ClsHelp.helpFile = Nz(rs!HelpFileName, "")
        ClsHelp.helpPath = Nz(rs!HelpFilePath, "")
                
        StatusMessage "Help Update/Export", "Information", "Select " & Chr(34) & "Update/Export" & Chr(34) & ". "
      
    End If
            
ExitNow:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHappened:
    StatusMessage "Configuration", "Error", Err.Number & " - " & Err.Description
    Resume ExitNow
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdExport_Click()
On Error GoTo ErrorHappened
    
    If (ClsHelp.ExportObjectList) Then
        MsgBox "All forms and controls context help IDs" & Chr(130) & " control tip text and status bar text are exported to " & ClsHelp.HelpTable & " table successfully!", vbInformation, "Help Update And Export"
    Else
        CmbAppName.SetFocus
        CmdExport.Enabled = False
    End If
       
ExitNow:
    On Error Resume Next
    Exit Sub

ErrorHappened:
    StatusMessage "Export", "Error", Err.Number & " - " & Err.Description
    CmbAppName.SetFocus
    CmdExport.Enabled = False
    Resume ExitNow

End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrorHappened
    
    If Nz(ClsHelp.helpFile, "") = "" Or Nz(ClsHelp.helpPath, "") = "" Then
        StatusMessage "Update", "Error", "Help file name and help path name are required to Update control properties. "
        StatusMessage "Update", "Error", "Verify that a valid help file and path are specified for the selected application in CT_HelpConfig table."
        CmbAppName.SetFocus
        CmdUpdate.Enabled = False
        Exit Sub
    
    Else
        Dim helpFile As String
                  
        'Making sure the file path contains trailing slash. If not the slash is added.
        helpFile = TrailingSlash(ClsHelp.helpPath) & ClsHelp.helpFile
        
        If Dir(helpFile) <> "" Then
            Identity.CCAHelp = helpFile
            
            If (ClsHelp.SetMapIds) Then
                MsgBox "All forms and controls context help IDs are updated successfully", vbInformation, "Help Update And Export"
            Else
                CmbAppName.SetFocus
                CmdUpdate.Enabled = False
            End If
        
        Else
            StatusMessage "Update", "Error", "Cannot locate Help File : " & helpFile
            StatusMessage "Update", "Error", "Verify that a valid help file and path are specified for the selected application in CT_HelpConfig table."
            CmbAppName.SetFocus
            CmdUpdate.Enabled = False
        End If
    End If
    
ExitNow:
    On Error Resume Next
    Exit Sub

ErrorHappened:
    StatusMessage "Update", "Error", Err.Number & " - " & Err.Description
    CmbAppName.SetFocus
    CmdUpdate.Enabled = False
    Resume ExitNow

End Sub

Private Sub Form_Load()
Set ClsHelp = New CT_ClsCnlyHelp

Me.LstMessages.AddItem "Source;Status;Message"

Me.visible = True

DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ClsHelp = Nothing
End Sub

Private Sub LstMessages_DblClick(Cancel As Integer)
    If Me.LstMessages.ListIndex <> -1 Then
        MsgBox "Message: " & LstMessages.Column(2, LstMessages.ListIndex + 1), vbInformation, LstMessages.Column(0, LstMessages.ListIndex + 1) & " - " & LstMessages.Column(1, LstMessages.ListIndex + 1)
    End If
End Sub

Sub StatusMessage(ByVal Src As String, ByVal Status As String, ByVal Msg As String)
Dim str As String
str = Src & ";" & Status & ";" & Msg

On Error Resume Next
    Dim i As Integer
    i = Me.LstMessages.ListCount
    If i >= 1 Then
        Me.LstMessages.AddItem str
    End If
    
    DoEvents
End Sub

Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function
