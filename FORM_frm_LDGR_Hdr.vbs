Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'=============================================
' ID:          Form_frm_LDGR_Hdr
' Author:      Kevin Dearing / Barbara Dyroff
' Create Date: 2012-06-27
' Description:
'      Display Transaction Ledger Header Information.
'
' Modification History:
'   2013-05-10 by BJD Change to Fee Balance display name for the Prepayment change.
'
' =============================================

Public Event Activate()
Public Event ApplyFilter(filter As String)
Public Event Current()
Public Event Click()
Public Event Deactivate()
Public Event FocusLost()
Public Event FocusGot()
Public Event Message(Txt As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Unload()
Public Event KeyPressed(AsciiKey As Integer)

Private csCurrentClaimNum As String


Private Const bPrintEvents As Boolean = False

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get ClaimNum() As String
    ClaimNum = Me.txtCnlyClaimNum
End Property

Public Property Get CurrentClaimNum() As String
    CurrentClaimNum = csCurrentClaimNum
End Property
Public Property Let CurrentClaimNum(sCurrentClaimNum As String)
    csCurrentClaimNum = sCurrentClaimNum
End Property


Private Sub Form_AfterFinalRender(ByVal drawObject As Object)
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_AfterFinalRender Event"
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_ApplyFilter Event"
End Sub

Private Sub Form_CommandExecute(ByVal Command As Variant)
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_CommandExecute Event"
End Sub

Private Sub Form_Current()
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_Current Event"
    CurrentClaimNum = CStr("" & Me.CnlyClaimNum)
    RaiseEvent Current
End Sub

Private Sub Form_DataChange(ByVal Reason As Long)
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_DataChange Event"
End Sub

Private Sub Form_DataSetChange()
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_DataSetChange Event"
End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    If bPrintEvents = True Then Debug.Print ClassName & ".Form_Filter Event"
End Sub



Private Sub Form_OnConnect()
    If bPrintEvents = True Then Debug.Print ClassName & ".on connect Event"

End Sub

Private Sub Form_SelectionChange()
    If bPrintEvents = True Then Debug.Print ClassName & ".SelectionChange Event"
End Sub
