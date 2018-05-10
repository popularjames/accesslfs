Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Private Sub FillCaseHdr()
'Begin sub
On Error GoTo ErrorHandler


'declare variables
    Dim strCaseID As String, strNoteText As String, strProviderNum As String, strAssignedTo As String, strRootCauseDesc As String, strSourceDesc As String
    Dim strStatusDesc As String, strDispositionDesc As String, strActionDesc As String, strCategoryDesc As String, strSubCategoryDesc As String
    Dim ErrorReturned As String
    Dim MyCodeAdo As New clsADO
    Dim rs As ADODB.RecordSet
    Dim oAdo As clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As String
    Dim strSQL As String
    
    'Form.RecordSource = "SELECT * FROM CtsCaseHdr WHERE CaseID = '" & Nz(lngEventID, "") & "'"
    'Form.Requery
    
    strCaseID = Nz(lngEventID, "")

    strSQL = "SELECT * FROM CtsCaseHdr WHERE CaseID = '" & strCaseID & "'"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = strSQL
        Set rs = .ExecuteRS
    End With

'   ' Set rs = MyAdo.OpenRecordSet(strSQL)
'    If rs.recordCount > 0 Then
'
'        Debug.Print "record count for caseid: " & strCaseID & " is "; rs.recordCount
'        With rs
'
'            txtCaseID.Value = !CaseId
'            txtNoteText.Value = !NoteText
'            txtProviderNum.Value = !ProviderNum
'            cbAssignedTo.Value = !AssignedTo
'            cbRootCause.Value = !RootCauseDesc
'            cbSource.Value = !SourceDesc
'            cbStatus.Value = !StatusDesc
'            cbDisposition.Value = !DispositionDesc
'            cbAction.Value = !ActionDesc
'            cbCategory.Value = !CategoryDesc
'            cbSubCategory.Value = !SubCategoryDesc
'
'        End With
'
'        CmdUpdate.Enabled = True
'    Else
'        CmdUpdate.Enabled = False
'        'me.Form.RecordLocks
'    End If
    
    
    
' Release used objects. Such as ado.
CleanupAndExit:
    Set oAdo = Nothing
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    With Err
    MsgBox ("FillCaseHdr subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit

'end sub

End Sub

Private Sub cmdCtsInsert_Click()
'Begin sub
On Error GoTo ErrorHandler
'declare variables
    Dim strProviderNum As String
    Dim strSourceDesc As String
    Dim strRootCauseDesc As String
    Dim strDispositionDesc As String
    Dim strCategoryDesc As String
    Dim strSubCategoryDesc As String
    Dim strActionDesc As String
    Dim strStatusDesc As String
    Dim strNoteText As String
    Dim strAssignedTo As String
    
    Dim strCaseID As String
    Dim ErrorReturned As String
    
    Dim MyCodeAdo As New clsADO
    Dim rs As ADODB.RecordSet
    Dim rs2 As ADODB.RecordSet
    Dim oAdo As clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As String
    
    'check for required data
    
    'get data to record
    strProviderNum = Nz(Me.txtProviderNum.Value, "")
    strSourceDesc = Nz(Me.cbSource.Value, "")
    strRootCauseDesc = Nz(Me.cbRootCause.Value, "")
    strDispositionDesc = Nz(Me.cbDisposition.Value, "")
    strCategoryDesc = Nz(Me.cbCategory.Value, "")
    strSubCategoryDesc = Nz(Me.cbSubCategory.Value, "")
    strActionDesc = Nz(Me.cbAction.Value, "")
    strStatusDesc = Nz(Me.cbStatus.Value, "")
    strNoteText = Nz(Me.txtNoteText.Value, "")
    strAssignedTo = Nz(Me.cbAssignedTo.Value, "")


    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "CtsInsert"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        DoCmd.Hourglass True
        cmd.Parameters.Refresh

        cmd.Parameters("@pProviderNum") = strProviderNum
        cmd.Parameters("@pSourceDesc") = strSourceDesc
        cmd.Parameters("@pRootCauseDesc") = strRootCauseDesc
        cmd.Parameters("@pDispositionDesc") = strDispositionDesc
        cmd.Parameters("@pCategoryDesc") = strCategoryDesc
        cmd.Parameters("@pSubCategoryDesc") = strSubCategoryDesc
        cmd.Parameters("@pActionDesc") = strActionDesc
        cmd.Parameters("@pStatusDesc") = strStatusDesc
        cmd.Parameters("@pNoteText") = strNoteText
        cmd.Parameters("@pAssignedTo") = strAssignedTo
       
        cmd.Execute

        DoCmd.Hourglass False
        strCaseID = Nz(.Parameters("@pEventID"), "")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
    
    End With

    MsgBox "Created ticket CaseID[" & strCaseID & "]", vbOKOnly
    
    
    'DoCmd.Close
    
' Release used objects. Such as ado.
CleanupAndExit:
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    With Err
    MsgBox ("cmdCtsInsert_Click subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit

'end sub
End Sub

Private Sub Form_Close()
    'DoCmd.SetWarnings True
End Sub


Private Sub Form_Load()
'Begin sub
On Error GoTo ErrorHandler
    
      
    
    

    
CleanupAndExit:
    Exit Sub
ErrorHandler:
    With Err
    MsgBox ("Form_Load subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit
End Sub



Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub
Private Sub frmAddrDetail_RecordChanged()
'    RaiseEvent RecordChanged
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub
