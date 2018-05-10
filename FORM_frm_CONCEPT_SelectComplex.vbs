Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const CstrFrmAppID As String = "ConceptLimits"


Private ActionButton As String

Private miAppPermission As Integer
Private mbAllowChange, mbAllowAdd, mbAllowView, mbAllowDelete As Boolean




Private Sub cmdAddLimit_Click()


On Error GoTo ErrHandler

If IsNumeric(Me.txtAddLimit) = False Or Nz(Me.txtAddField, "") = "" Then
MsgBox "'Field' cannot be empty and 'Limit' must be numeric."
Exit Sub
End If

DoCmd.Hourglass (True)


Dim HoldRecordSource As String
Dim HoldFilter As String

HoldRecordSource = Me.ctl_SubForm1.Form.RecordSource
HoldFilter = Me.ctl_SubForm1.Form.filter



    
Dim cmd As ADODB.Command
Dim strSQL As String
Dim myCode_ADO As clsADO
Set myCode_ADO = New clsADO
    
myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
myCode_ADO.SQLTextType = sqltext

Set cmd = New ADODB.Command
Set cmd.ActiveConnection = myCode_ADO.CurrentConnection
cmd.CommandTimeout = 0
cmd.commandType = adCmdText

myCode_ADO.sqlString = "exec usp_RPT_R0142X2 '" & Identity.UserName() & "', '" & Me.cboAddLimit & "', '" & Me.txtAddField & "'," & Me.txtAddLimit & ", '" & Me.txtAddNote & "'"
'MsgBox MyCode_ADO.sqlString
cmd.CommandText = myCode_ADO.sqlString
cmd.Execute


myCode_ADO.sqlString = "exec usp_RPT_R0142X1 '" & Identity.UserName() & "'"
cmd.CommandText = myCode_ADO.sqlString
cmd.Execute

Me.txtAddField = ""
Me.txtAddLimit = ""
Me.txtAddNote = ""
 
'this put it back to the way you were viewing it
'Me.ctl_SubForm1.Form.RecordSource = HoldRecordSource
'Me.ctl_SubForm1.Form.filter = HoldFilter
'Me.ctl_SubForm1.Form.FilterOn = True

    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form"
    Me.ctl_SubForm1.Form.OrderBy = "updatedt desc"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.ctl_SubForm1.Form.FilterOn = False
    Me.lstSetBy = "All"
    Me.lstLimitType = 0
     
DoCmd.Hourglass (False)
MsgBox "Done adding limit."


Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical
End Sub

Private Sub cmdUpdateLimits_Click()

On Error GoTo ErrHandler

DoCmd.Hourglass (True)

Dim HoldRecordSource As String
Dim HoldFilter As String

HoldRecordSource = Me.ctl_SubForm1.Form.RecordSource
HoldFilter = Me.ctl_SubForm1.Form.filter

    
Dim cmd As ADODB.Command
Dim strSQL As String
Dim myCode_ADO As clsADO
Set myCode_ADO = New clsADO
    
myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
myCode_ADO.SQLTextType = sqltext

Set cmd = New ADODB.Command
Set cmd.ActiveConnection = myCode_ADO.CurrentConnection
cmd.CommandTimeout = 0
cmd.commandType = adCmdText

myCode_ADO.sqlString = "exec usp_RPT_R0142X1 '" & Identity.UserName() & "'"
cmd.CommandText = myCode_ADO.sqlString
cmd.Execute

    
Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form"
 
'this puts it back to the way you were viewing it
'Me.ctl_SubForm1.Form.RecordSource = HoldRecordSource
'Me.ctl_SubForm1.Form.filter = HoldFilter
'Me.ctl_SubForm1.Form.FilterOn = True

    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form"
    Me.ctl_SubForm1.Form.OrderBy = "updatedt desc"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.ctl_SubForm1.Form.FilterOn = False
    Me.lstSetBy = "All"
    Me.lstLimitType = 0
    
DoCmd.Hourglass (False)

MsgBox "Done updating limits."

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical


End Sub

Private Sub Form_Load()

'Dim strSQL As String
    
    Me.Caption = "Concept Limits"
   Me.RecordSource = ""
    
   ' Me.frmAppID = CstrFrmAppID
    '
    Call Account_Check(Me)
    Me.lstLimitType = 0
    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form"
    Me.ctl_SubForm1.Form.OrderBy = "updatedt desc"
    Me.ctl_SubForm1.Form.OrderByOn = True
    
    'miAppPermission = UserAccess_Check(Me)
    'If miAppPermission = 0 Then Exit Sub
    
   ' miAppPermission = GetAppPermission(Me.frmAppID)
   ' mbAllowChange = (miAppPermission And gcAllowChange)
   ' mbAllowAdd = (miAppPermission And gcAllowAdd)
   ' mbAllowView = (miAppPermission And gcAllowView)
   ' mbAllowDelete = (miAppPermission And gcAllowDelete)
    
    
    
   ' Me.lstLimitType = "Concept Group"
   ' lstLimitType_Click
   ' Me!ctl_SubForm1.Form.FilterOn = True
    
End Sub



Private Sub lstLimitType_Click()
'MsgBox Me.lstLimitType
'Me.lstLimitType.Selection.ForeColor = vbRed


Dim strSQL As String

Me!ctl_SubForm1.SourceObject = "frm_CONCEPT_SelectComplex_sub1"

'MsgBox Me.lstLimitType

If Me.lstLimitType = 0 Then
    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form"
    Me.ctl_SubForm1.Form.OrderBy = "updatedt desc"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.lstSetBy = "All"
ElseIf Me.lstLimitType >= 80 Then
    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form WHERE LimitID >= " & Me.lstLimitType
    Me.ctl_SubForm1.Form.OrderBy = "LimitField1, LimitField2, LimitField3"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.lstSetBy = "All"
Else
    Me.ctl_SubForm1.Form.RecordSource = "select * FROM SELECT_Limits_Form WHERE LimitID = " & Me.lstLimitType
    Me.ctl_SubForm1.Form.OrderBy = "LimitField1, LimitField2, LimitField3"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.lstSetBy = "All"
End If


End Sub








Private Function ValidLimit(ByRef TextField As TextBox)

    ValidLimit = True
    
    If Nz(TextField, "") = "" Then GoTo Validate_Error
    If Not (val(TextField) >= 0 And val(TextField) <= 99999) Then GoTo Validate_Error
    If Not IsNumeric(TextField) Then GoTo Validate_Error
    
    Exit Function
    
Validate_Error:

    ValidLimit = False
    TextField.SetFocus
    MsgBox "Limit must be between 0 and 99,999", vbInformation, "Validation Error"
        
End Function

Private Function ValidDRG(ByRef TextField As TextBox)

    ValidDRG = True
    
    If Nz(TextField, "") = "" Then GoTo Validate_Error
    If Not (val(TextField) >= 0 And val(TextField) <= 999) Then GoTo Validate_Error
    If Not Len(TextField) = 3 Then GoTo Validate_Error
    If Not IsNumeric(TextField) Then GoTo Validate_Error
    
    Exit Function
    
Validate_Error:

    ValidDRG = False
    TextField.SetFocus
    MsgBox "DRG must be between 000 and 999", vbInformation, "Validation Error"
        
End Function

Private Function ValidPriority(ByRef TextField As TextBox)

    ValidPriority = True
    
    If Not (val(Nz(TextField, 0)) >= 0 And val(Nz(TextField, 0)) <= 99) Then GoTo Validate_Error
    If Not IsNull(TextField) And Not IsNumeric(TextField) Then GoTo Validate_Error
    
    Exit Function
    
Validate_Error:

    ValidPriority = False
    TextField.SetFocus
    MsgBox "Priority must be between 0 and 99", vbInformation, "Validation Error"
        
End Function


Private Sub lstSetBy_Click()

If Me.lstSetBy = "All" Then
    Me.ctl_SubForm1.Form.FilterOn = False
ElseIf Me.lstSetBy = "Sort by Date" Then
    Me.ctl_SubForm1.Form.filter = "[UpdateDt] is not null"
    Me.ctl_SubForm1.Form.OrderBy = "updatedt desc"
    Me.ctl_SubForm1.Form.OrderByOn = True
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Set by Hand" Then
    Me.ctl_SubForm1.Form.filter = "[Setby] = 'Hand'"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Set by Me" Then
    Me.ctl_SubForm1.Form.filter = "[UpdateUser] = '" & Identity.UserName() & "'"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Set by Calculation" Then
    Me.ctl_SubForm1.Form.filter = "[Setby] = 'Calc'"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Not Set" Then
    Me.ctl_SubForm1.Form.filter = "[Setby] = '--'"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Limit > 0" Then
    Me.ctl_SubForm1.Form.filter = "[Limit] > 0"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Limit = 0" Then
    Me.ctl_SubForm1.Form.filter = "[Limit] = 0"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Unlimited" Then
    Me.ctl_SubForm1.Form.filter = "[Limit] is null"
    Me.ctl_SubForm1.Form.FilterOn = True
ElseIf Me.lstSetBy = "Has Ever Limit" Then
    Me.ctl_SubForm1.Form.filter = "[EverCap] is not null"
    Me.ctl_SubForm1.Form.FilterOn = True
End If





End Sub
