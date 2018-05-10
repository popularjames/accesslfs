Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Event NoteAdded()

Private strAppID As String
Private iNoteID As Long
Private strUserName As String

Private mstrDefaultNoteType As String

Private rsNote As ADODB.RecordSet

'This property will let the system know what parent application
'to update with the new note id
Property Let frmAppID(data As String)
    strAppID = data
End Property
Property Let DefaultNoteType(data As String)
    mstrDefaultNoteType = data
End Property

Property Get frmAppID() As String
   frmAppID = strAppID
End Property


Property Set NoteRecordSource(data As ADODB.RecordSet)
    Set rsNote = data
    
    If rsNote.BOF = True And rsNote.EOF = True Then
        iNoteID = -1
    Else
        rsNote.MoveFirst
        iNoteID = rsNote("NoteID")
    End If
    
End Property


Property Get NoteRecordSource() As ADODB.RecordSet
     Set NoteRecordSource = rsNote
End Property


'This is a public refresh, so we can call it from elsewhere
Public Sub RefreshData()
    Dim strErrMsg As String
    Dim strErrSource As String
    
    strErrSource = "GeneralNoteAdd.RefreshData"
    
    On Error GoTo Err_handler
    
    Me.Caption = "Note Add"
    'Refresh the combobox to show the correct notes
    RefreshComboBox "SELECT NoteType, NoteType FROM NOTE_XREF_Type WHERE AppID = '" & strAppID & "'", Me.NoteTypeID, , "NoteType"
   
    If rsNote Is Nothing Then
        strErrMsg = "Note recordset is not defined.  Please set it first before calling this routine"
        GoTo Err_handler
    End If
    
    If Not (rsNote.BOF = True And rsNote.EOF = True) Then
        rsNote.MoveFirst
    End If
    
    txtAllNotes = ""
    While Not rsNote.EOF
         txtAllNotes = txtAllNotes & "Added by " & Trim(UCase(rsNote("NoteUserID"))) & " @ " & rsNote("NoteDate") & " Note Type " & Trim(UCase(rsNote("NoteType"))) & vbCrLf
         txtAllNotes = txtAllNotes & String(100, "-") & vbCrLf
         txtAllNotes = txtAllNotes & Trim(rsNote("NoteText")) & vbCrLf & vbCrLf
         rsNote.MoveNext
    Wend
    txtAllNotes.SetFocus
    If txtAllNotes <> "" Then txtAllNotes.SelLength = 0
    Me.NoteTypeID.SetFocus
'Alex C 3/8/2012 - set the default note type in the combobox, if given
NoteTypeID.Value = mstrDefaultNoteType
    
    If gintAccountID = 1 Then
        If Me.frmAppID = "AuditClm" Then
            Me.chkDiscussion.visible = True
            Me.cboDiscussionCompany.visible = True
            Me.cboDiscussionCompany.RowSource = " SELECT [Xref_DiscussionCompany].[DiscussionCompanyID], [Xref_DiscussionCompany].[DiscussionCompanyDesc] FROM [Xref_DiscussionCompany] "
            If Me.chkDiscussion.Value <> 0 Then
                Me.cboDiscussionCompany.Enabled = True
            Else
                Me.cboDiscussionCompany.Enabled = False
            End If
        Else
            Me.chkDiscussion.visible = False
            Me.cboDiscussionCompany.visible = False
        End If
    Else
        Me.cboDiscussionCompany.RowSource = ""
    End If
    
    
        
Exit_Function:
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then
        Err.Raise vbObjectError + 513, strErrSource, strErrMsg
    Else
        Err.Raise Err.Number, strErrSource, Err.Description
    End If
End Sub
Private Sub chkDiscussion_Click()
    If Me.chkDiscussion.Value <> 0 Then
        Me.cboDiscussionCompany.Enabled = True
    Else
        Me.cboDiscussionCompany.Enabled = False
    End If
End Sub
Private Sub cmdExit_Click()
    If Nz(Me.NoteText, "") <> "" Or Me.NoteTypeID.ListIndex >= 0 Then
        If MsgBox("Record has been changed. Are you sure you want to exit? ", vbYesNo) = vbYes Then
            DoCmd.Close acForm, Me.Name
        End If
    Else
        DoCmd.Close acForm, Me.Name
    End If
End Sub


Private Sub cmdSave_Click()

'BEGIN 4/17/2013: replace the loop with regular expression to remove the special characters
Dim i As Integer '3/27/2013 KCF
Dim iChar As String '3/27/2013 KCF
Dim strNoteText As String '4/17/2013 KCF

''BEGIN 3/27/2013 KCF: Need to check for special characters as they are causing issues when Inserting the Note
For i = 1 To Len(Me.NoteText)
    iChar = Mid(Me.NoteText, i, 1)
    If Asc(iChar) = 63 And Mid(Me.NoteText, i, 1) <> "?" Then
        strNoteText = Me.NoteText
        SpecialCharacterCatch (strNoteText)
        MsgBox ("Special characters like emoticons have been removed from this Note. Please review the Note for accuracy before saving.")
    Exit For
    End If
    Next i
''END 3/27/2013 KCF: Need to check for special characters as they are causing issues when Inserting the Note
'BEGIN 4/17/2013: replace the loop with regular expression to remove the special characters


    If Me.NoteTypeID.ListIndex = -1 Then
        MsgBox "Please select a note type", vbOKOnly + vbInformation
        Exit Sub
    ElseIf Nz(Me.NoteText, "") = "" Then
        MsgBox "Note field is blank", vbOKOnly + vbInformation
        Exit Sub
    Else
        InsertNewNote
        Me.RefreshData
        RaiseEvent NoteAdded
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub InsertNewNote()
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim strSQL As String
    
    '************CMS ONLY DPR added 1/6/2011****************
    Dim intNoteRecordCount As Integer
    '************CMS ONLY DPR added 1/6/2011****************
    
    
    strErrSource = "GeneralNoteAdd.InsertNewNote"
    
    On Error GoTo Err_handler
    
    If iNoteID = -1 Then iNoteID = GetAppKey("NOTE")
    
    '************CMS ONLY DPR added 1/6/2011****************
    intNoteRecordCount = rsNote.recordCount + 1
    '************CMS ONLY DPR added 1/6/2011****************
    
    If Me.NoteText <> "" Then
        With rsNote
            .AddNew
            !NoteID = iNoteID
            !SeqNo = rsNote.recordCount + 1
            !AppID = Me.frmAppID
            !NoteType = Me.NoteTypeID
            !NoteText = Me.NoteText
            !NoteUserID = strUserName
            !NoteDate = Now()
            .UpdateBatch
        End With
    End If
    
    
    If gintAccountID = 1 Then
    '************CMS ONLY DPR added 1/6/2011****************
    'If this is a audit company, we want to attribute it to this note per Chad
    'This note will be added to a cross reference table that tracks the note and the claim
    'If a note is entered and not saved, this will potentially break
    If Me.frmAppID = "AuditClm" Then
        If Me.chkDiscussion.Value <> 0 Then
            If Nz(Me.cboDiscussionCompany, "") <> "" Then
                CurrentDb.Execute ("INSERT INTO AuditClm_DiscussionCompany (DiscussionCompanyID, NoteID, SeqNo ) VALUES ( '" & Me.cboDiscussionCompany & "', " & iNoteID & " , " & intNoteRecordCount & ")")
            End If
        End If
    End If
    '************CMS ONLY DPR added 1/6/2011****************
    End If

Exit_Sub:
    Exit Sub

Err_handler:
    Err.Raise Err.Number, strErrSource, Err.Description
End Sub
Private Sub Form_Load()
    Me.Caption = "Adding Note Form"
    Call Account_Check(Me)
    strUserName = Identity.UserName()
    iNoteID = -1
End Sub
Private Sub TabCtrl_Change()
    If UCase(TabCtrl.Pages(TabCtrl.Value).Name) = "ALLNOTES" Then
        txtAllNotes.SetFocus
        If txtAllNotes <> "" Then txtAllNotes.SelLength = 0
    End If
End Sub


Public Function SpecialCharacterCatch(sText As String) As String
'*************************************************************************
'Added Wednesday 4/17/2013 by Kathleen C Flanagan
'*************************************************************************
Dim oRegEx As RegExp

    Set oRegEx = New RegExp
    oRegEx.Global = True
    oRegEx.Pattern = "([^\40-\176\r\n\t])"
    removespecialcharacters = oRegEx.Replace(sText, "")
    Set oRegEx = Nothing
    
    Debug.Print removespecialcharacters
    
    Me.NoteText = removespecialcharacters
      
      
      
End Function
