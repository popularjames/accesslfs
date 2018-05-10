Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:           Form_frm_SCANNING_MR_Invoice_Main
' Author:       Barbara Dyroff
' Date:         2010-02-03
' Description:
'   Display Medical Record Scanning Invoice data for a given Provider.  Display both the totals
' for an Invoice and the details associated with each Invoice.
'
' Modification History:
'
'
' =============================================

Private bolMRInvDtlFormLoaded As Boolean
Private strMRInvPropTableName As String
Private strMRInvPropKey As String
Private strMRInvFrmAppID As String


Public Property Let frmAppID(data As String)
    strMRInvFrmAppID = data
    Call UserAccess_Check(Me)
End Property

Public Property Get frmAppID() As String
    frmAppID = strMRInvFrmAppID
End Property

Public Sub DetailFormLoaded()
    bolMRInvDtlFormLoaded = True
End Sub

Public Property Let PropTableName(data As String)
    strMRInvPropTableName = data
End Property

Public Property Let AppTitle(data As String)
    Me.Caption = data
End Property

Public Property Let PropKey(data As String)
    strMRInvPropKey = data
End Property

Public Property Let cnlyProvID(data As String)
    Me.txtCnlyProvID = data
End Property

' Refresh the data for both the Invoice Properties (Totals) as well as the Detail datasheets.
Public Sub RefreshData()
    Me.sfrm_SCANNING_MR_Invoice_Properties.Form.PropTotTableName = strMRInvPropTableName
    Me.sfrm_SCANNING_MR_Invoice_Properties.Form.PropTotKey = strMRInvPropKey
    Me.sfrm_SCANNING_MR_Invoice_Properties.Form.RefreshData
    RefreshDetail
End Sub

' Check user permissions and load.
Private Sub Form_Load()
    Dim iAppPermission As Integer
    
    Call Account_Check(Me)
    
    If Me.frmAppID <> "" Then
        iAppPermission = UserAccess_Check(Me)
        If iAppPermission = 0 Then Exit Sub
    End If
    
    If IsSubForm(Me) Then
        Me.cmdExit.visible = False
    End If
    Me.sfrm_SCANNING_MR_Invoice_Detail.SourceObject = "frm_SCANNING_MR_Invoice_Detail"
End Sub

'Refresh the detailed data for the Current Invoice.
Public Sub RefreshDetail()
    If bolMRInvDtlFormLoaded Then
        If txtSQLSource & "" <> "" Then
            Me.sfrm_SCANNING_MR_Invoice_Detail.Form.RecordSQL = txtSQLSource
            Me.sfrm_SCANNING_MR_Invoice_Detail.Form.RefreshData
        Else
            'Create an empty record set for display.
            Me.txtSQLSource = "SELECT * FROM " & strMRInvPropTableName & " WHERE 1 = 2 "
            Me.sfrm_SCANNING_MR_Invoice_Detail.Form.RecordSQL = txtSQLSource
        End If
        Me.sfrm_SCANNING_MR_Invoice_Detail.Form.RefreshData
    End If
End Sub

Private Sub cmdExit_Click()
    On Error GoTo Err_cmdExit_Click
    DoCmd.Close acForm, Me.Name

Exit_cmdExit_Click:
    Exit Sub

Err_cmdExit_Click:
    MsgBox Err.Description
    Resume Exit_cmdExit_Click
    
End Sub


Private Sub Form_Close()
    On Error Resume Next
    RemoveObjectInstance Me
End Sub
