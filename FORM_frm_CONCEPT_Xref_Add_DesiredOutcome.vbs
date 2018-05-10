Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Private Sub CmdCancel_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdCancel_Click"
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".cmdSave_Click"
    
    If Trim(Nz(Me.txtDesiredOutcomeToAdd, "")) = "" Then
        MsgBox "Please enter some text to add or click the cancel button", vbOKOnly, "Nothing to add!"
        GoTo Block_Exit
    End If
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_ADD_DesiredOutcome"
        .Parameters.Refresh
        .Parameters("@pAdjustedOutcome") = Trim(Me.txtDesiredOutcomeToAdd)
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg").Value, , True
            GoTo Block_Exit
        End If
    End With
    
    DoCmd.Close acForm, Me.Name, acSaveNo
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
