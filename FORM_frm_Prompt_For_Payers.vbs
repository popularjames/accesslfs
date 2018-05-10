Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private cstrSelPayerNameIdsStr As String
Private cstrSelPayerNamesStr As String
Private cbShowPastPayers As Boolean
Public Event PayersSelected(sPayerNames() As String, sPayerIds() As String)
Private cstAllowedPayers As String



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get SelPayerNameIdsString() As String
    SelPayerNameIdsString = cstrSelPayerNameIdsStr
End Property
Public Property Let SelPayerNameIdsString(sSelPayerNameIdString As String)
    cstrSelPayerNameIdsStr = sSelPayerNameIdString
End Property

Public Property Get AllowedPayers() As String
    AllowedPayers = cstAllowedPayers
End Property
Public Property Let AllowedPayers(strAllowedPayers As String)
    cstAllowedPayers = strAllowedPayers
    Call RefreshData
End Property


Public Property Get PromptText() As String
    PromptText = Me.lblPromptText.Caption
End Property
Public Property Let PromptText(sPromptText As String)
    Me.lblPromptText.Caption = sPromptText
End Property

Public Property Get ShowPastPayers() As Boolean
    ShowPastPayers = cbShowPastPayers
End Property
Public Property Let ShowPastPayers(bShowPastPayers As Boolean)
    cbShowPastPayers = bShowPastPayers
End Property


Public Property Get SelPayerNamesString() As String
    SelPayerNamesString = cstrSelPayerNamesStr
End Property
Public Property Let SelPayerNamesString(sSelPayerNamesString As String)
    cstrSelPayerNamesStr = sSelPayerNamesString
End Property



Public Property Get SelPayerNameIdsArray() As Variant
Dim iPayerID As Integer
Dim iIndx As Integer
Dim saryPayers() As String
Dim iaryPayers() As Integer

    saryPayers = Split(SelPayerNameIdsString, ",")
    
    For iIndx = 0 To UBound(saryPayers)
        iPayerID = CInt(saryPayers(iIndx))
        ReDim Preserve iaryPayers(iIndx)
        iaryPayers(iIndx) = iPayerID
    Next
    SelPayerNameIdsArray = iaryPayers
    
'    SelPayerNameIdsArray = CInt(Split(SelPayerNameIdsString, ","))
End Property


Public Property Get SelPayerNamesArray() As Variant
    SelPayerNamesArray = Split(SelPayerNamesString, ",")
End Property


Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub


Private Sub cmdOk_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oSFrm As Form_frm_PAYERNAMES
Dim sSelNameString As String


    strProcName = ClassName & ".cmdOK_Click"

    Set oSFrm = Me.sfrm_PAYERNAMES.Form
'    Set oSFrm = Forms("frm_PAYERNAMES")
    
    If oSFrm.AtLeastOnePayerSelected = False Then
        MsgBox "No payers are selected. Please select at least 1 or click cancel to cancel", vbOKOnly + vbExclamation, "No payers selected!"
        GoTo Block_Exit
    End If

    Me.SelPayerNameIdsString = oSFrm.GetSelectedPayerNameIDs(sSelNameString)
    Me.SelPayerNamesString = sSelNameString
    
    Call UnloadForm(Me)
    
      
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Load()

    If Me.OpenArgs <> "" Then
        Me.PromptText = Me.OpenArgs
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    If Me.OpenArgs <> "" Then
        Me.PromptText = Me.OpenArgs
    End If
    RefreshData
    
End Sub

Public Sub RefreshData()
Dim sFltr As String

    If AllowedPayers <> "" Then
        sFltr = "PayerNameId IN (" & AllowedPayers & ") "
    End If
    
    
    If Me.ShowPastPayers = False Then
        If sFltr <> "" Then sFltr = sFltr & "AND "
        sFltr = sFltr & "EndDate > #" & Format(Now, "m/d/yyyy") & "#"
    End If

    If sFltr <> "" Then
        Me.sfrm_PAYERNAMES.Form.filter = sFltr
        Me.sfrm_PAYERNAMES.Form.FilterOn = True
    Else
        Me.sfrm_PAYERNAMES.Form.filter = ""
        Me.sfrm_PAYERNAMES.Form.FilterOn = False
    End If
    Me.sfrm_PAYERNAMES.Form.Requery
    
End Sub
