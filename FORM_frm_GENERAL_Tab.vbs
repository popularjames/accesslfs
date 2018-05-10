Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private strCnlyClaimNum As String
Private strRowSource As String
Private strAppID As String

Property Let frmAppID(data As String)
    strAppID = data
End Property

Property Get frmAppID() As String
    AppID = strAppID
End Property

Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property

Property Let CnlyRowSource(data As String)
     strRowSource = data
End Property

Property Get CnlyRowSource() As String
     CnlyRowSource = strRowSource
End Property

Private Sub Command2_Click()
    RefreshData
End Sub

'This is a public refresh, so we can call it from elsewhere
Public Sub RefreshData()
    Dim strError As String
    On Error GoTo ErrHandler
    
    'Refresh the grid based on the rowsource passed into the form
    Me.frm_GENERAL_Datasheet.Form.InitData strRowSource, 2
    Me.frm_GENERAL_Datasheet.Form.RecordSource = strRowSource

    Dim ctl As Control
     
    'Loop through the controls and size them correctly.
    For Each ctl In Me.frm_GENERAL_Datasheet.Form.Controls
      If ctl.ControlType = acTextBox Then
          ctl.ColumnWidth = -2
      End If
   Next

exitHere:
    Exit Sub
    
ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume exitHere
End Sub

Private Sub Form_Close()
    'This form can be instanced, so it is removed from the global collection before it is closed
    RemoveObjectInstance Me
End Sub

Private Sub Command3_Click()
    On Error GoTo Err_Command3_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70

Exit_Command3_Click:
    Exit Sub

Err_Command3_Click:
    MsgBox Err.Description
    Resume Exit_Command3_Click
    
End Sub
