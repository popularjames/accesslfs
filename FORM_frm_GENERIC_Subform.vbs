Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mSourceData As String

Private Property Let SourceData(ByVal vData As String)
    mSourceData = vData
End Property

Private Property Get SourceData() As String
    SourceData = mSourceData
End Property

Private Sub Form_Load()

    Me.RecordSource = mSourceData
    Me.Requery
    
End Sub
