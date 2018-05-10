Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private frmNotes As Form_frm_GENERAL_Notes_Display
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1


Private Sub cmdExit_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdViewAllNotes_Click()
    Dim rs As ADODB.RecordSet
    Dim strNoteIDs As String
    
    Set MyAdo = New clsADO
    
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    MyAdo.sqlString = "select NoteID from QUEUE_Dtl where CnlyClaimNum = '" & Me.RecordSet("cnlyClaimNum") & "' and NoteID is not null"
    Set rs = MyAdo.OpenRecordSet
    
    If rs.BOF = True And rs.EOF = True Then
        MsgBox "There is no notes to display"
    Else
        rs.MoveFirst
        While rs.EOF <> True
            If strNoteIDs = "" Then
                strNoteIDs = rs(0)
            Else
                strNoteIDs = strNoteIDs & "," & rs(0)
            End If
            rs.MoveNext
        Wend
    
        If strNoteIDs <> "" Then
            strNoteIDs = "(" & strNoteIDs & ")"
    
            Set frmNotes = New Form_frm_GENERAL_Notes_Display
            ColObjectInstances.Add frmNotes, frmNotes.hwnd & ""
            frmNotes.CnlyRowSource = "select * from NOTE_Detail where NoteID in " & strNoteIDs & " order by NoteID asc"
            frmNotes.RefreshData
            ShowFormAndWait frmNotes
            Set frmNotes = Nothing
        End If
    End If
    
    Set rs = Nothing
    Set MyAdo = Nothing
End Sub

Private Sub cmdViewNote_Click()
    Set frmNotes = New Form_frm_GENERAL_Notes_Display
    ColObjectInstances.Add frmNotes, frmNotes.hwnd & ""
    frmNotes.CnlyRowSource = "select * from NOTE_Detail where NoteID = " & Me.RecordSet("NoteID")
    frmNotes.RefreshData
    ShowFormAndWait frmNotes
    Set frmNotes = Nothing
End Sub

Private Sub Form_Current()
    With Me.RecordSet
        If IsNull(![NoteID]) Then
            lblNoteExists.visible = False
            cmdViewNote.Enabled = False
        Else
            lblNoteExists.visible = True
            cmdViewNote.Enabled = True
        End If
    End With
End Sub

Private Sub Form_Load()
    Set Me.QUEUE_Dtl.Form.RecordSet = Me.RecordSet
End Sub


Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub
