Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'=============================================
' ID:          Form_frm_ADMIN_User_Hours_Popup
' Author:
' Create Date:
' Description:
'      Prompt the user to enter hours.
'
' Modification History:
'   2010-04-30 by BJD to allow multiple effective Audit Numbers (use table ADMIN_Audit_Number_Multi).  Initially for Aetna.
'                   Also, default to the minimum Audit Number for the Year on the combo box.
' =============================================

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private dblHours As Double
Private dtmEffectiveDate As Date
Private intAuditNum As Integer
Private mReturnDate As Date
Public Event RecordSaved(AuditNum As Integer, EffectiveDate As Date, Hours As Double, CancelButton As Boolean)
Property Let Hours(data As Double)
    dblHours = data
End Property
Property Get Hours() As Double
    Hours = dblHours
End Property
Property Let EffectiveDate(data As Date)
    dtmEffectiveDate = data
End Property
Property Get EffectiveDate() As Date
    EffectiveDate = dtmEffectiveDate
End Property
Property Let AuditNum(data As Integer)
    intAuditNum = data
End Property
Property Get AuditNum() As Integer
    AuditNum = intAuditNum
End Property
Public Sub RefreshData()
    Dim strSQL As String
    Dim iAuditCnt As Integer
    
    strSQL = "SELECT an.AuditNum, an.AuditDesc " & _
             "FROM ADMIN_Audit_Number as an INNER JOIN ADMIN_User_AuditNum as ua ON an.AuditNum = ua.AuditNum " & _
             "WHERE UCase(ua.UserID)='" & UCase(Identity.UserName()) & "' AND date() between ua.EffDt And ua.TermDt " & _
             " and an.AccountID = " & gintAccountID
    
    
    'RefreshComboBox "SELECT AuditNum, AuditDesc FROM ADMIN_Audit_Number_Multi WHERE AccountID = " & gintAccountID & "", Me.cboAuditNum, Me.AuditNum, "AuditNum"
    RefreshComboBox strSQL, Me.cboAuditNum, Me.AuditNum, "AuditNum"
    Me.txtEffectiveDate = Me.EffectiveDate
    Me.txtHours = Me.Hours
    iAuditCnt = DLookup("Count(AuditNum)", "ADMIN_User_AuditNum", " Date() between EffDt And TermDt and UserID = '" & Identity.UserName() & "'") + 0
    If iAuditCnt = 1 Then
        'Me.cboAuditNum.DefaultValue = DLookup(count("AuditNum"), "ADMIN_User_AuditNum", " Date() between EffDt And TermDt and UserID = '" & Identity.Username() & "'") & ""
        Me.cboAuditNum.DefaultValue = DLookup("AuditNum", "ADMIN_User_AuditNum", " Date() between EffDt And TermDt and Ucase(UserID) = '" & UCase(Identity.UserName()) & "'") & ""
    Else
        Me.cboAuditNum.DefaultValue = ""
    End If
End Sub
Private Sub cmdEffDate_Click()
 On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtEffectiveDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    'Prevent user from entering a future date
    If CDate(mReturnDate) > Date Then
        MsgBox "Date cannot be in the future", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    'To avoid value 12:00:00 AM when closing calendar form
    If mReturnDate <> #12:00:00 AM# Then
        Me.txtEffectiveDate = mReturnDate
        Exit Sub
    End If


Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub

Private Sub cmdIncrementCalendar_Click()
    'Prevents user from entering in dates in the future
    If CDate(Me.txtEffectiveDate.Value) < Date Then
        Me.txtEffectiveDate.Value = CDate(Me.txtEffectiveDate.Value) + 1
    End If
End Sub

Private Sub cmdDecrementCalendar_Click()
    Me.txtEffectiveDate.Value = CDate(Me.txtEffectiveDate.Value) - 1
End Sub

Private Sub CmdCancel_Click()
    RaiseEvent RecordSaved(Nz(Me.cboAuditNum, 0), Nz(Me.txtEffectiveDate, "1/1/1900"), Nz(Me.txtHours, 0), True)
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdSave_Click()
    ' Check for audit number
    If Me.cboAuditNum & "" = "" Then
        MsgBox "Please select an Audit number", vbOKOnly + vbCritical
        Exit Sub
    
    'Check if Hours is numbers
    ElseIf Not IsNumeric(Me.txtHours) Then
        MsgBox "Hours must be numeric", vbOKOnly + vbCritical
        Me.txtHours = 0
        Exit Sub
        
    'Check if hours is NULL or blank
    ElseIf Nz(Me.txtHours, "") = "" Then
        MsgBox "Please enter hours", vbOKOnly + vbCritical
        Exit Sub
    
    'Check if hours is blank or zero
    ElseIf Me.txtHours <= 0 Then
        MsgBox "Hours cannot be zero or less", vbOKOnly + vbCritical
        Exit Sub
    
    'Check if hours > 24
    ElseIf Me.txtHours > 24 Then
        MsgBox "Hours cannot be greater than 24", vbOKOnly + vbCritical
        Me.txtHours = 24
        Exit Sub
    End If
    
    RaiseEvent RecordSaved(Nz(Me.cboAuditNum, 0), Nz(Me.txtEffectiveDate, "1/1/1900"), Nz(Me.txtHours, 0), False)
    DoCmd.Close acForm, Me.Name
End Sub


Private Sub frmCalendar_DateSelected(SelectedDate As Date)
    mReturnDate = SelectedDate
End Sub



Private Sub txtEffectiveDate_Click()
    mReturnDate = Me.txtEffectiveDate.Value
End Sub


Private Sub txtEffectiveDate_LostFocus()
    'Prevent user from entering non-date data
    If Not IsDate(Me.txtEffectiveDate) Then
        MsgBox "Please enter valid date: MM/DD or MM/DD/YYYY", vbOKOnly + vbCritical
        Me.txtEffectiveDate.Value = mReturnDate
        Exit Sub
    End If
        
    'Prevent user from changing FROM DATE that is less than TO DATE
    If CDate(Me.txtEffectiveDate.Value) > Date Then
        MsgBox "To Date cannot be less than From Date", vbOKOnly + vbCritical
        Me.txtEffectiveDate.Value = mReturnDate
    End If
    
End Sub
