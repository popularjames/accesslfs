Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Property Set NoteRecordSource(data As ADODB.RecordSet)
    Set Me.RecordSet = data
End Property


Private Sub Form_Load()
    Me.RecordSource = ""
    Me.AllowAdditions = False
    Me.AllowEdits = False
End Sub
