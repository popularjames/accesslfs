Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private frmExceptionClear As Form_frm_QUEUE_Exception_Clear
Private mrsExceptions As ADODB.RecordSet
Private mstrFilter As String

Property Let FormFilter(data As String)
    mstrFilter = data
End Property

Public Sub RefreshData()
    Dim strRowSource As String
    
    strRowSource = "select * from v_QUEUE_Exception_Info where AccountID = " & gintAccountID
    'TL add account ID logic
    If mstrFilter <> "" Then
        strRowSource = strRowSource & " and " & mstrFilter
    End If
    
    Me.RecordSource = strRowSource
    Me.Requery
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    If Not (Me.RecordSet.BOF And Me.RecordSet.EOF) Then
        If IsSubForm(Me) Then
            If Me.Parent.RecordChanged Then
                MsgBox "Please save your record first before proceeding", vbInformation
            Else
                Set frmExceptionClear = New Form_frm_QUEUE_Exception_Clear
                frmExceptionClear.FormFilter = "CnlyClaimNum = '" & Me.RecordSet("CnlyClaimNum") & "' and ExceptionType = '" & Me.RecordSet("ExceptionType") & "'"
                frmExceptionClear.RefreshData
                ShowFormAndWait frmExceptionClear
                Set frmExceptionClear = Nothing
                RefreshData
            End If
        End If
    End If
End Sub
