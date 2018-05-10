Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'=============================================
' ID:          Form_frm_LDGR_Dtl
' Author:      Kevin Dearing / Barbara Dyroff
' Create Date: 2012-06-27
' Description:
'      Display Transaction Ledger Detail Information.  The detailed displayed will be prepared by frm_LDGR_Main.
' A new Recordset is created with additional derived Ledger information (balance info etc.)
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

Private cdblCollStart As Double
Private cdblFeeStart As Double

Private cdblCollBal As Double
Private cdblFeeBal As Double


Private csCurrentClaimNum As String

Private Const cbPrintEvents As Boolean = True


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property



Public Property Get CurrentClaimNum() As String
    CurrentClaimNum = csCurrentClaimNum
End Property
Public Property Let CurrentClaimNum(sCurrentClaimNum As String)
    csCurrentClaimNum = sCurrentClaimNum
End Property



Private Sub Form_AfterLayout(ByVal drawObject As Object)
    If cbPrintEvents = True Then Debug.Print ClassName & ".Form_AfterLayout"

End Sub

Private Sub Form_AfterRender(ByVal drawObject As Object, ByVal chartObject As Object)
    If cbPrintEvents = True Then Debug.Print ClassName & ".Form_AfterRender"
End Sub

Private Sub Form_Current()
    RaiseEvent Current
End Sub


Private Sub Form_DataChange(ByVal Reason As Long)
    If cbPrintEvents = True Then Debug.Print ClassName & ".Form_DataChange"
End Sub

Private Sub Form_Load()
'    Me.Controls("Country").ScrollBarAlign = 2
  

End Sub

Private Sub Form_Query()
    If cbPrintEvents = True Then Debug.Print ClassName & ".Form_Query"
End Sub
