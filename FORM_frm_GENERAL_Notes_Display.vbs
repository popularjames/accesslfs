Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private mstrRowSource As String
Private mrsNote As ADODB.RecordSet

Property Let CnlyRowSource(data As String)
     mstrRowSource = data
End Property

Property Get CnlyRowSource() As String
     CnlyRowSource = mstrRowSource
End Property

Property Set NoteRecordSource(data As ADODB.RecordSet)
     Set mrsNote = data
     mstrRowSource = ""
End Property

Property Get NoteRecordSource() As ADODB.RecordSet
     Set NoteRecordSource = mrsNote
End Property

'This is a public refresh, so we can call it from elsewhere
Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
    Me.Caption = "View Notes"
    If mstrRowSource <> "" And mrsNote Is Nothing Then
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = CnlyRowSource
        
        Set mrsNote = MyAdo.OpenRecordSet()
    End If
    
    txtAllNotes = ""
    
    If Not (mrsNote.BOF = True And mrsNote.EOF = True) Then
        mrsNote.MoveFirst
    End If
    
    While Not mrsNote.EOF
         txtAllNotes = txtAllNotes & "Added by " & Trim(UCase(mrsNote("NoteUserID"))) & " @ " & mrsNote("NoteDate") & " Note Type " & Trim(UCase(mrsNote("NoteType"))) & vbCrLf
         txtAllNotes = txtAllNotes & String(100, "-") & vbCrLf
         'BEGIN 3/27/2013 KCF: Change so that html tags created for Return to Auditor QA Notes are dropped - html tags are included for formatting in R3 for subcontrators
         'txtAllNotes = txtAllNotes & Trim(mrsNote("NoteText")) & vbCrLf & vbCrLf
         txtAllNotes = txtAllNotes & Trim(Replace(Replace(Replace(Replace(Replace(Replace(Nz(mrsNote("NoteText"), ""), "<p>", ""), "</b>", ""), "<br/>", "" & vbCrLf & ""), "</html>", ""), "<b>", ""), "</p>", "" & vbCrLf & "")) & vbCrLf & vbCrLf
         'END 3/27/2013 KCF: Change so that html tags created for Return to Auditor QA Notes are dropped - html tags are included for formatting in R3 for subcontrators
         mrsNote.MoveNext
    Wend
    
    txtAllNotes.SelLength = 0
    
    
exitHere:
    Set MyAdo = Nothing
    Exit Sub
    
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub

Private Sub Form_Close()
    On Error Resume Next
    'Instanced form, remove from collection
    RemoveObjectInstance Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Note Display"
    
    Call Account_Check(Me)
End Sub

Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    Err.Raise ErrNum, ErrSource, ErrMsg
End Sub
