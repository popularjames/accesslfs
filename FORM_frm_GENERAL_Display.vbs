Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit


Public Event FormClosed()
Private strTextPrompt As String 'ACL 08-31-2012
Property Let TextPrompt(data As String)  'ACL 08-31-2012
     strTextPrompt = data
End Property
Property Get TextPrompt() As String      'ACL 08-31-2012
     TextPrompt = strTextPrompt
End Property



Private Sub Form_Close()
    RaiseEvent FormClosed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent FormClosed
End Sub
Public Sub RefreshData()
Me.txtPrompt = Me.TextPrompt
End Sub
