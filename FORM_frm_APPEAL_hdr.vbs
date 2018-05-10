Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mstrSearchType As String
Private mstrSearchTable As String

Property Let SearchType(data As String)
    mstrSearchType = data
End Property

Property Get SearchType() As String
    SearchType = mstrSearchType
End Property

Property Let SearchTable(data As String)
    mstrSearchTable = data
End Property

Property Get SearchTable() As String
    SearchTable = mstrSearchTable
End Property

Private Sub Form_BeforeInsert(Cancel As Integer)
   ' Dim Identity As New ClsIdentity
    Me.AppealReceiptDt = Now()
    Me.EnteredBy = Identity.UserName
End Sub

Private Sub Form_Current()
On Error GoTo err_hndlr
    If Me.Count = 0 Then
        Me.Parent.ApDetails.Form.filter = " CnlyClaimNum=''"
    Else
        Me.Parent.ApDetails.Form.filter = " CnlyClaimNum='" & Me.CnlyClaimNum & "'"
    End If
    Me.Parent.ApDetails.Form.FilterOn = True
    
    If Me.Parent.ApDetails.Form.Count = 0 Then
        Me.Parent.PackageDetails.Form.filter = " PackageID=-1"
        Me.Parent.PackageDetails.Form.FilterOn = True
    End If
err_hndlr:
    'Necessary because details load after header!
End Sub

Private Sub Form_DblClick(Cancel As Integer)
On Error GoTo ErrHandler
        
    Dim strParameter As String
    Dim strParameterString As String
    
    Dim strError As String
    Dim strParent As String
    Dim arrParameters() As String
    Dim intI As Integer
    
    Me.SearchType = "AUDITCLM"
    strParameterString = ""
    
    strParent = Me.Name
    
    strParameter = Nz(DLookup("Parameter", "GENERAL_Navigate", "SearchType = '" & Me.SearchType & "' and ActionName = 'dblClick' and parentform = '" & strParent & "'"), "")
    arrParameters = Split(strParameter, "|")
    
    If UBound(arrParameters) > 0 Then
        For intI = 0 To UBound(arrParameters)
           strParameterString = strParameterString & Me.RecordSet(arrParameters(intI)) & "|"
        Next intI
    Else
          Me.CnlyClaimNum.SetFocus
          strParameterString = strParameterString & Nz(DLookup(strParameter, "AUDITCLM_HDR", "CnlyClaimNum = '" & Me.CnlyClaimNum.Text & "'"), "")
          'Debug.Print strParameterString
    End If
    
    If strParameter <> "" Then
        Navigate strParent, Me.SearchType, "DblClick", strParameterString
    End If

Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"
End Sub

Private Sub Form_Load()
   ' Dim Identity As New ClsIdentity
    Identity.UserName
    Me.filter = "EnteredBy = '" & Identity.UserName & "'"
End Sub
