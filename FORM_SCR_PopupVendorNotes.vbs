Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private MvVenID1 As String
Private MvVenID2 As String

Public Property Let VenID1(data As String)
MvVenID1 = data
TxtVenID1 = MvVenID1
End Property
Public Property Let VenID2(data As String)
MvVenID2 = data
TxtVenID2 = MvVenID2
End Property
Public Property Get VenID1() As String
VenID1 = MvVenID1
End Property
Public Property Get VenID2() As String
VenID2 = MvVenID2
End Property

Private Sub Form_Load()
Me.TxtAuditor = Identity.Auditor
ClearAllProblemFields
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Me.SubForm
    .Width = Me.InsideWidth - (.left * 2)
    .Height = Me.InsideHeight - (.left) - .top
End With
End Sub
Private Sub cmdAddNote_Click()
On Error GoTo ErrorHappened

'---VALIDATE REQUIRED INFORMATION --
Dim Msg As String
If "" & Me.TxtVenID1 = "" Then
    ProblemField Me.TxtVenID1, True
    Msg = Msg & "You must provide both VenID1 and VenID2 to add the record." & vbCrLf
End If
If "" & Me.TxtVenID2 = "" Then
    ProblemField Me.TxtVenID2, True
    Msg = Msg & "You must provide both VenID1 and VenID2 to add the record." & vbCrLf
End If
If "" & Me.TxtAuditor = "" Then
    ProblemField Me.TxtAuditor, True
    Msg = Msg & "You must provide auditor initials." & vbCrLf
End If
If "" & Me.TxtNote = "" Then
    ProblemField Me.TxtNote, True
    Msg = Msg & "You must provide text for the note." & vbCrLf
End If
'---CHECK THE LENGTH OF THE NOTE
If Len(Me.TxtNote) > 8000 Then
    ProblemField Me.TxtNote, True
    Msg = Msg & "The text in your note field is too long." & vbCrLf
End If

If "" & Msg <> "" Then
    MsgBox Msg, vbInformation, "Unable to add Note"
    Me.Repaint
    GoTo ExitNow
End If

If Me.AddNote(Me.TxtVenID1, Me.TxtVenID2, Me.TxtAuditor, Me.TxtNote) = True Then
    ClearAllProblemFields
    DoEvents
    Identity.Auditor = Me.TxtAuditor
    Me.SubForm.Form.Requery
End If


ExitNow:
    On Error Resume Next
    Exit Sub

ErrorHappened:
    MsgBox "Unable to Add Note" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Add Vendor Note"
    Resume ExitNow
    
End Sub
Private Function ProblemField(objCtrl As TextBox, HighLight As Boolean)
    With objCtrl
        If HighLight = True Then
            .BorderColor = 255
        Else
            .BorderColor = 0
        End If
    End With

End Function
Private Sub ClearAllProblemFields()
    ProblemField Me.TxtVenID1, False
    ProblemField Me.TxtVenID2, False
    ProblemField Me.TxtAuditor, False
    ProblemField Me.TxtNote, False
End Sub

Public Function AddNote(VenID1 As String, VenID2 As String, Auditor As String, Note As String) As Boolean
On Error GoTo ErrorHappened
Dim db As DAO.Database, rst As DAO.RecordSet

screen.MousePointer = 11
'---VALIDATE REQUIRED INFORMATION --
Dim Msg As String
Select Case ""
Case "" & VenID1, "" & VenID2
    Msg = "You must provide both VenID1 and VenID2 to add the record."
Case "" & Auditor
    Msg = "You must provide auditor initials."
Case "" & Note
    Msg = "You must provide text for the note."
End Select
If "" & Msg <> "" Then
    MsgBox Msg, vbInformation, "Unable to add Note"
    AddNote = False
    GoTo ExitNow
End If

'---DO THE WORK --
Set db = CurrentDb
'TODO Find CnlyVendorNotes
Set rst = db.OpenRecordSet("CnlyVendorNotes", dbOpenDynaset, dbAppendOnly + dbSeeChanges)
With rst
    .AddNew
    .Fields("VenID1") = VenID1
    .Fields("VenID2") = VenID2
    .Fields("Notes") = Note
    .Fields("Host") = Identity.Computer
    .Fields("UserName") = Identity.UserName
    .Fields("Auditor") = Auditor
    .Update
    .Close
End With
AddNote = True

ExitNow:
    On Error Resume Next
    screen.MousePointer = 0
    Set rst = Nothing
    Set db = Nothing
    Exit Function

ErrorHappened:
    AddNote = True = False
    MsgBox "Unable to Add Note" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Add Vendor Note"
    Resume ExitNow
    Resume
End Function
