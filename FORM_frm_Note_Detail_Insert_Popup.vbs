Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_Note_Detail_Insert_Popup
' Author:      Barbara Dyroff
' Create Date: 2010-05-17
' Description:
'   Prompt the user to create a new note.
'
' Input:
'  txtNoteID
'  txtAppID
'  txtSeqNo   Current SeqNo + 1 for display purposes
'
' Modification History:
'
' =============================================


Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private strNoteStoredProcName As String

Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

' Call the stored procedure to create the note.
Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    Dim objCmd As New ADODB.Command
    Dim iResult As Integer
    Dim intResult As Integer
    Dim strSQL As String
    Dim strMsg As String
    Dim intNoteID As Integer
    Dim lngNoteID As Long
    Dim varNoteID As Variant
    Dim strUserID As String
    Dim lngNewNoteID As Long
    Dim strNoteID As String
        
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    
    'Validate that a Note Type has been selected.
    If IsNull(Me.cboNoteType) Then
        MsgBox "Please select a NoteType", vbCritical
        GoTo Exit_Sub
    End If
    
    
    ' Create a note.
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    myCode_ADO.sqlString = "usp_NOTE_Detail_Insert"
    myCode_ADO.SQLTextType = StoredProc

    strNoteID = Me.txtNoteID
    lngNoteID = val(strNoteID)
      
    strUserID = Identity.UserName()
    
    objCmd.Parameters.Append _
        objCmd.CreateParameter("RC", adInteger, adParamReturnValue)
        
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pNoteID", adInteger, adParamInputOutput, 4, lngNoteID) '4 bytes for SQL Server Int
        
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pAppID", adVarChar, adParamInput, _
            10, Me.txtAppID)
            
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pNoteType", adVarChar, adParamInput, _
            50, Me.cboNoteType)
            
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pNoteText", adVarChar, adParamInput, _
            5000, Me.txtNoteText)
            
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pNoteUserID", adVarChar, adParamInput, _
            50, strUserID)
            
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pErrMsg", adVarChar, adParamOutput, 255)
        
    myCode_ADO.BeginTrans
    intResult = myCode_ADO.Execute(objCmd.Parameters)

    strMsg = Nz(objCmd.Parameters("pErrMsg").Value, "")
    
    lngNewNoteID = objCmd("pNoteID")
    
    ' Check that the ADOCls method completed successfully.
    If intResult <> 1 Then
        GoTo ErrHandler
    End If
    
    'Check that the Stored Procedure Completed Successfully.
    If objCmd("RC") <> 0 Then
        GoTo ErrHandler
    End If
   
    myCode_ADO.CommitTrans

    DoCmd.Close acForm, Me.Name
   
Exit_Sub:
    Set myCode_ADO = Nothing
    Set objCmd = Nothing
    Exit Sub

ErrHandler:
    If strMsg <> "" Then
      MsgBox strMsg, vbOKOnly + vbCritical
    Else
        MsgBox Err.Description, vbOKOnly + vbCritical
    End If
    
    'Improve clsADO Rollback Trans - Need to check if there is an open transaction in RollbackTrans.
    'If no open trans, do not rollback -- throws an error.
    myCode_ADO.RollbackTrans
    
    Resume Exit_Sub

End Sub


Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & vbCrLf & ErrSource
End Sub
