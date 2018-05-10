Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrCurrRelatedQA As Integer

Public Event UpdateReferences(RelatedQA As Integer)

Const CstrFrmAppID As String = "AuditClmRef"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Property Let CurrRelatedQA(data As Integer)
     mstrCurrRelatedQA = data
End Property



Public Sub RefreshScreen()
    Dim strError As String
    On Error GoTo ErrHandler
    
    If mstrCurrRelatedQA = 1 Then 'pending
        Me.fraRelatedQA = 0
    Else
        Me.fraRelatedQA = mstrCurrRelatedQA
    End If
    

exitHere:
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub


Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strErrMsg As String
    
    If Not (Me.fraRelatedQA >= 2 And Me.fraRelatedQA <= 3) Then
        MsgBox "You did not select a valid Related MR QA value.", vbInformation, "Error"
        Exit Sub
    End If

    

    
    RaiseEvent UpdateReferences(Me.fraRelatedQA)
    DoCmd.Close acForm, Me.Name
End Sub



Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub


Private Sub Form_Load()
    
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    Me.Caption = "AuditClm_References Related Claim Image QA"
    
    Call RefreshScreen
End Sub
