Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const CstrFrmAppID As String = "AuditClmRationale"

Private strCnlyClaimNum As String
Private strlabel As String
Private strText As String
Private bLocked As Boolean
Private strTextTemplate As String
Private strTextPrompt As String 'ACL 08-31-2012
Public Event TextConfirmed(strText As String, strControlName As String, bCancel As Boolean)
Public Event FormClosed()
Private strControlName As String
Property Let CnlyClaimNum(data As String)
     strCnlyClaimNum = data
End Property
Property Get CnlyClaimNum() As String
     CnlyClaimNum = strCnlyClaimNum
End Property
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
Property Let Locked(data As Boolean)
     bLocked = data
End Property
Property Get Locked() As Boolean
     Locked = bLocked
End Property





Property Let TextTemplate(data As String)
     strTextTemplate = data
End Property
Property Get TextTemplate() As String
     TextTemplate = strTextTemplate
End Property

Property Let TextPrompt(data As String)  'ACL 08-31-2012
     strTextPrompt = data
End Property
Property Get TextPrompt() As String      'ACL 08-31-2012
     TextPrompt = strTextPrompt
End Property



Property Let ControlName(data As String)
     strControlName = data
End Property
Property Get ControlName() As String
     ControlName = strControlName
End Property




Public Sub RefreshData()
Me.txtPrompt = Me.TextPrompt
Me.txtDisplay = Me.TextTemplate
Me.txtZoomText = Me.TextData
Me.lblAppTitle.Caption = "Claim Num :: " & Me.CnlyClaimNum & " Zoom ::  " & Me.TextLabel

If bLocked Then
    Me.txtZoomText.Locked = True
    Me.CmdOK.Enabled = False
Else
    Me.txtZoomText.Locked = False
    Me.CmdOK.Enabled = True
End If

End Sub
Private Sub CmdCancel_Click()
    RaiseEvent TextConfirmed(Nz(Me.txtZoomText), strControlName, True)
    'RemoveObjectInstance Me
    RaiseEvent FormClosed
End Sub

Private Sub cmdOk_Click()
    RaiseEvent TextConfirmed(Nz(Me.txtZoomText), strControlName, False)
    'RemoveObjectInstance Me
    RaiseEvent FormClosed
End Sub

Private Sub Form_Close()
    RaiseEvent FormClosed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent FormClosed
End Sub
