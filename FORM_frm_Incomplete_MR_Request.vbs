Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public saved As Boolean
Private curSelectedBoxes As Long
Private curNotes As String
Private curOther As String
Private FormID As Integer
Private NoteID As Long
Private maxSelected As Long

Private rsNotes As ADODB.RecordSet

Const CstrFrmAppID As String = "IncMREntry"

Public Property Get SelectedBoxes() As Long
    SelectedBoxes = curSelectedBoxes
End Property
Property Let SelectedBoxes(selection As Long)
     curSelectedBoxes = selection
End Property

Public Property Get FormType() As Integer
    FormType = FormID
End Property
Property Let FormType(currForm As Integer)
     FormID = currForm
End Property

Public Property Get Notes() As String
    Notes = curNotes
End Property
Property Let Notes(Notes As String)
     curNotes = Notes
End Property

Public Property Get other() As String
    other = curOther
End Property
Property Let other(other As String)
     curOther = other
End Property

Property Set NoteRecordSource(data As ADODB.RecordSet)
    Set rsNotes = data
    
    If rsNotes.BOF = True And rsNotes.EOF = True Then
        NoteID = -1
    Else
        rsNotes.MoveFirst
        NoteID = rsNotes("NoteID")
    End If
    
End Property

Property Get NoteRecordSource() As ADODB.RecordSet
     Set NoteRecordSource = rsNotes
End Property

Private Sub cmdSave_Click()
SaveData
End Sub

Private Sub cmdDeleteRequest_Click()
Dim DeleteMessage As String
DeleteMessage = "Are you sure you would like to delete this request for Medical Records? You will not be able to undo your changes!"
 
If MsgBox(DeleteMessage, vbYesNo + vbQuestion) = vbYes Then
           Call DeleteRequest
           DoCmd.Close
End If
        
End Sub

Private Function GetSelectedCheckBoxes() As Long

 maxSelected = Me.frm_MR_Needed_Subform.Form.maxSelected
 
 Dim selectedMR As Long
    Dim i As Long

    selectedMR = 0
    
    If Me.frm_MR_Needed_Subform.Controls.Item("ch1").Value = -1 Then
    selectedMR = 1
    End If

    For i = 2 To maxSelected
    
        If Me.frm_MR_Needed_Subform.Controls.Item(chBoxName & i).Value = -1 Then
        selectedMR = selectedMR + i
        End If

    i = (i * 2) - 1
    Next

GetSelectedCheckBoxes = selectedMR
End Function

Private Sub SaveData()

If GetSelectedCheckBoxes = 0 Then
MsgBox ("No Medical Records had been checked in order to request additional information. Nothing will be saved at this time.")
Exit Sub
End If

Dim selection As String
Dim Notes As Variant
Dim completeNoteToSave As String

    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim iResult As Integer

    On Error GoTo ErrHandler

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_AUDITCLM_Incomplete_MR_Requested_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_AUDITCLM_Incomplete_MR_Requested_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pCnlyClaimNum") = Me.RecordSet("CnlyClaimNum")
    cmd.Parameters("@pattributes") = GetSelectedCheckBoxes
    cmd.Parameters("@pAuditorNotes") = Nz(Me.txtNotes, "")
    cmd.Parameters("@pFormID") = FormID
    cmd.Parameters("@pUserID") = GetUserName
    cmd.Parameters("@pOther") = Nz(Me.frm_MR_Needed_Subform.Form!txtOther, "")
    
    iResult = myCode_ADO.Execute(cmd.Parameters)

    'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        MsgBox "SaveData", "Error Saving Incomplete MR Selection - " & strErrMsg
    Else
        MsgBox ("Incomplete Medical Records Selection had been saved.")
        saved = True
        
        'Notes = GetRequestInfoAndNotes(Me.Recordset("CnlyClaimNum"))
        'completeNoteToSave = "The following medical records had been requested:" & Notes(0)
        'completeNoteToSave = completeNoteToSave & "." & vbCrLf & vbCrLf & "Additional Auditor Notes: " & vbCrLf & Notes(1)
        
        'SaveNote (completeNoteToSave)
    End If
    
Exit_Sub:
    Set cmd = Nothing
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error refreshing list"
    Resume Exit_Sub
End Sub

Private Sub cmdExit_Click()
On Error GoTo Err_cmdExit_Click

If Not saved Then
 
    Dim Message As String
    Dim changed As Boolean
    Message = "Would you like to save your "
    changed = False
    
    If Me.txtNotes <> Me.Notes Or Me.frm_MR_Needed_Subform.Form!txtOther <> Me.other Then
    Message = Message & "notes"
    changed = True
    End If
    
    If GetSelectedCheckBoxes <> Me.SelectedBoxes Then
        If Not (changed) Then
         Message = Message & "selection"
         changed = True
        Else
           Message = Message & " and selection"
        End If
    End If
    
    If changed = True Then
        If MsgBox(Message & "?", vbYesNo + vbQuestion) = vbYes Then
            Call SaveData
        End If
    End If
End If
    
    DoCmd.Close

Exit_cmdExit_Click:
    Exit Sub

Err_cmdExit_Click:
    MsgBox Err.Description
    Resume Exit_cmdExit_Click
    
End Sub

Private Sub SaveNote(NoteToSave As String)

  Dim max_seq As Long
  Dim oRs As ADODB.RecordSet
  
  Set oRs = rsNotes.Clone
  
  With rsNotes
            .AddNew
            
            !NoteID = NoteID
            !AppID = CstrFrmAppID
            !NoteType = "GENERAL"
            !NoteText = NoteToSave
            !NoteUserID = GetUserName
            !NoteDate = Now()
        
        'Record count includes the record being added at this point
            If Nz(rsNotes.recordCount, 0) > 0 Then
                    max_seq = rsNotes.recordCount
            
                    'So an issue with sequence number missing. Can't rely on number of records.
                    If Not (oRs.BOF = True And oRs.EOF = True) Then
                        oRs.MoveFirst
                        While Not oRs.EOF
                        'MsgBox ("NoteID " & mrsNotes("NoteID") & " Seq Num " & mrsNotes("SeqNo") & " APP ID " & mrsNotes("AppID") & " Note Type " & mrsNotes("NoteType") & " Note Text " & mrsNotes("NoteText") & " Note User ID " & mrsNotes("NoteUserID") & " NoteDate " & mrsNotes("NoteDate"))
                            If oRs("SeqNo") > max_seq Then
                                max_seq = oRs("SeqNo")
                            End If
                        oRs.MoveNext
                        Wend
                    End If
            
                    max_seq = max_seq + 1
                    !SeqNo = max_seq
            Else
               !SeqNo = 1
            End If
            
            .UpdateBatch
        End With
       
End Sub

Private Sub DeleteRequest()

Dim MyCodeAdo As clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim deleteNote As String
deleteNote = "Request for Incomplete Medical Records had been deleted."

Set MyCodeAdo = New clsADO

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_AUDITCLM_Incomplete_MR_Delete_Request"
                cmd.Parameters.Refresh
                cmd.Parameters("@pCnlyClaimNum") = Me.RecordSet("CnlyClaimNum")
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@pErrMsg")
Else
    SaveNote (deleteNote)
End If

Set MyCodeAdo = Nothing
Set cmd = Nothing

End Sub
