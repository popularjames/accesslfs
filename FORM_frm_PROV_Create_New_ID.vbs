Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_PROV_Create_NEW_ID
' Description:
'
'
' Modification History:
'
'   2012-01-09 by Andrew Lauer to fix the bug when creating a new payer
' =============================================





Public Event ReturnIDs(cnlyProvID As String, ProvID As String, PayerID As String)

Private Sub cmdExit_Click()
    RaiseEvent ReturnIDs("", "", "")
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strCnlyProvID As String
    Dim strProvID As String
    Dim strPayerID As String
    
    strPayerID = Trim(Me.PayerID) & ""
    strProvID = Trim(Me.ProvID) & ""
    
    If strProvID = "" Then
        MsgBox "Please enter the client prov ID", vbInformation, "Input Error"
        Me.ProvID.SetFocus
        Exit Sub
    End If
    
    If strPayerID = "" Then
        MsgBox "Please select a payer ID", vbInformation, "Input Error"
        Me.PayerID.SetFocus
        Exit Sub
    End If
    
    'strCnlyProvID = strPayerID & strProvID     ' thieu
    strCnlyProvID = strProvID                   ' thieu
    RaiseEvent ReturnIDs(strCnlyProvID, strProvID, strPayerID)
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    Call Account_Check(Me)
    Me.PayerID.RowSource = "SELECT PayerNum, PayerName FROM XREF_Payer WHERE AccountID =  " & gintAccountID
    Me.PayerID.Requery
End Sub
