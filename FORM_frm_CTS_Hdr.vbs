Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private WithEvents frmAddrDetail As Form_frm_PROV_Addr
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

'Public Event RecordChanged()

'Private strRowSource As String
'
'Private mrsPROVAddr As ADODB.RecordSet
'Private mrsPROVAddrPortal As ADODB.RecordSet
'Private mrsPROVAddrDeleted As ADODB.RecordSet
'
'Private mbRecordLocked As Boolean
'Private miAppPermission As Integer

'Const CstrFrmAppID As String = "ProvAddr"



Private Sub cmdUpdate_Click()
'Begin sub
On Error GoTo ErrorHandler
'declare variables
    Dim strCaseID As String, strNoteText As String, strProviderNum As String, strAssignedTo As String, strRootCauseDesc As String, strSourceDesc As String
    Dim strStatusDesc As String, strDispositionDesc As String, strActionDesc As String, strCategoryDesc As String, strSubCategoryDesc As String
    

    Dim ErrorReturned As String
    Dim MyCodeAdo As New clsADO
    Dim rs As ADODB.RecordSet
    Dim rs2 As ADODB.RecordSet
    Dim oAdo As clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As String
    
    strCaseID = Nz(Me.txtCaseID.Value, "")
    strNoteText = Nz(Me.txtNoteText.Value, "")
    strProviderNum = Nz(Me.txtProviderNum.Value, "")
    strAssignedTo = Nz(Me.cbAssignedTo.Value, "")
    strRootCauseDesc = Nz(Me.cbRootCause.Value, "")
    strSourceDesc = Nz(Me.cbSource.Value, "")
    strStatusDesc = Nz(Me.cbStatus.Value, "")
    strDispositionDesc = Nz(Me.cbDisposition.Value, "")
    strActionDesc = Nz(Me.cbAction.Value, "")
    strCategoryDesc = Nz(Me.cbCategory.Value, "")
    strSubCategoryDesc = Nz(Me.cbSubCategory.Value, "")
        
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "CtsUpdateCase"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        DoCmd.Hourglass True
        cmd.Parameters.Refresh
        cmd.Parameters("@pCaseId") = strCaseID
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
        
        '.Execute
        DoCmd.Hourglass False
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErr"), "")
    
    End With

    MsgBox "Updated CaseID: " & strCaseID, vbOKOnly
    
' Release used objects. Such as ado.
CleanupAndExit:
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    With Err
    MsgBox ("FillCaseHdr subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit

'end sub
End Sub


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

   ' Set rs = MyAdo.OpenRecordSet(strSQL)
    If rs.recordCount > 0 Then
        
        Debug.Print "record count for caseid: " & strCaseID & " is "; rs.recordCount
        With rs
                    
            txtCaseID.Value = !CaseId
            txtNoteText.Value = !NoteText
            txtProviderNum.Value = !ProviderNum
            cbAssignedTo.Value = !AssignedTo
            cbRootCause.Value = !RootCauseDesc
            cbSource.Value = !SourceDesc
            cbStatus.Value = !StatusDesc
            cbDisposition.Value = !DispositionDesc
            cbAction.Value = !ActionDesc
            cbCategory.Value = !CategoryDesc
            cbSubCategory.Value = !SubCategoryDesc
        
        End With
        
        CmdUpdate.Enabled = True
    Else
        CmdUpdate.Enabled = False
        'me.Form.RecordLocks
    End If
    
    
    
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

Private Sub Form_Close()
    DoCmd.SetWarnings True
End Sub


Private Sub Form_Load()
'Begin sub
On Error GoTo ErrorHandler
    
    'get case info
    FillCaseHdr
    
    
    

    
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
