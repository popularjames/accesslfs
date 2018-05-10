Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private strText As String
Private strlabel As String

Property Let TextData(data As String)
     strText = data
End Property
Property Get TextData() As String
     TextData = strText
End Property

Property Let TextLabel(data As String)
     strlabel = data
End Property
Property Get TextLabel() As String
     TextLabel = strlabel
End Property
Public Sub RefreshData()

    Me.txtData = Me.TextData
    Me.lblAppTitle.Caption = Me.TextLabel
End Sub
