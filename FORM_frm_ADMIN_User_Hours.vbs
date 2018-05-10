Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frmPopup As Form_frm_ADMIN_User_Hours_Popup
Attribute frmPopup.VB_VarHelpID = -1

Private mstrUserProfile As String
Private mbRecordChanged As Boolean

Private miAppPermission As Integer
Private mbAllowView As Boolean
Private mbAllowChange As Boolean
Private mbAllowDelete As Boolean
Private mbAllowAdd As Boolean
Private mbLocked As Boolean
Private mReturnDate As Date

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Const CstrFrmAppID As String = "UserHours"
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property
Property Let RecordChanged(data As Boolean)
    mbRecordChanged = data
End Property
Property Get RecordChanged() As Boolean
    RecordChanged = mbRecordChanged
End Property

Private Sub cmdAddNew_Click()

    On Error GoTo ErrHandler
    
    Set frmPopup = New Form_frm_ADMIN_User_Hours_Popup
    frmPopup.Hours = 0
    frmPopup.EffectiveDate = Date
    'frmPopup.AuditNum = Me.lstHours.Column(GetColumnPosition(Me.lstHours, "AuditNum"), 1)
    frmPopup.RefreshData
    ShowFormAndWait frmPopup

ErrHandler_Exit:
    Set frmPopup = Nothing
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ErrHandler_Exit
End Sub

Private Sub cmdRefresh_Click()
    Me.RefreshData
End Sub

Private Sub cmdToDate_Click()
    'To open calendar for date selection
    
On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtToDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    'To avoid value 12:00:00 AM when closing calendar form
    If mReturnDate = #12:00:00 AM# Then
        Exit Sub
    End If
    
    'Prevent user from entering TO DATE that is less than FROM DATE
    If CDate(mReturnDate) < CDate(Me.txtFromDate) Then
        MsgBox "To Date cannot be less than From Date", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    Me.txtToDate = mReturnDate
        
Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim sCmd As String
    MsgBox "Please enter your time in TimePlus", vbOKOnly, "Enter hours in TimePlus now"
    sCmd = """C:\Program Files (x86)\Internet Explorer\iexplore.exe"" ""https://timeplus.myconnolly.com/"""
    Shell sCmd
    
    Cancel = True
End Sub

Private Sub frmCalendar_DateSelected(SelectedDate As Date)
    mReturnDate = SelectedDate
End Sub

Private Sub cmdFromDate_Click()
On Error GoTo Err_btnChkDt_Click
       
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.txtFromDate, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    'To avoid value 12:00:00 AM when closing calendar form
    If mReturnDate = #12:00:00 AM# Then
        Exit Sub
    End If
    
    'Prevent user from entering FROM that is greater than TODATE
    If CDate(mReturnDate) > CDate(Me.txtToDate) Then
        MsgBox "From Date cannot be greater than To Date", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    Me.txtFromDate = mReturnDate
        
Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click
End Sub

Private Sub Form_Load()
    Me.Caption = "User Hour Data Entry Form"
    
    Call Account_Check(Me)
    'miAppPermission = UserAccess_Check(Me)
    'If miAppPermission = 0 Then Exit Sub
    
    'mstrUserProfile = GetUserProfile()
    'CheckPermission

    'If mbAllowAdd = True Then
        Me.cmdAddNew.Enabled = True
    'End If

    'If mbAllowDelete = True Then
        Me.CmdDelete.Enabled = True
    'End If
    
    'Setting default date for txtFromDate
    txtFromDate = Date - 29
    
    'Setting default date for txtToDate
    txtToDate = Date
        
    RefreshData

End Sub
Private Sub CheckPermission()
    If miAppPermission = gcLocked Then mbLocked = True Else mbLocked = False
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    mbAllowDelete = (miAppPermission And gcAllowDelete)
    mbAllowChange = (miAppPermission And gcAllowChange) Or mbAllowAdd Or mbAllowDelete
    mbAllowView = (miAppPermission And gcAllowView) Or mbAllowChange
End Sub
Public Sub RefreshData()
    On Error GoTo ErrHandler
    Dim strSQL As String
    Dim rst As ADODB.RecordSet
    
    'Creating a new instance of ADO-class variable
    Set MyAdo = New clsADO
    
    'Making a Connection call to SQL database?
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    'Setting a string to a SQL query statement, depending on ID
    strSQL = "SELECT UserID, WorkDate, HoursWorked, AuditNum"
    strSQL = strSQL & " FROM Admin_User_Hours"
    strSQL = strSQL & " WHERE UserID = '" & Identity.UserName & "'"
    strSQL = strSQL & " AND WorkDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "'"
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    MyAdo.sqlString = strSQL
    
    'Setting the list record set equal to the specify ADO-class record set
    Set lstHours.RecordSet = MyAdo.OpenRecordSet()

    'Write Sum query
    strSQL = "SELECT Sum(HoursWorked) as TotalHoursWorked"
    strSQL = strSQL & " FROM Admin_User_Hours"
    strSQL = strSQL & " WHERE UserID = '" & Identity.UserName & "'"
    strSQL = strSQL & " AND WorkDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "'"
      
    MyAdo.sqlString = strSQL
    Set rst = MyAdo.OpenRecordSet()
    
    If Not rst.EOF Then
        Me.txtTotalHours = Nz(rst!TotalHoursWorked, 0)
    Else
        Me.txtTotalHours = 0
    End If
   
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Sub

Private Sub frmPopup_RecordSaved(AuditNum As Integer, EffectiveDate As Date, Hours As Double, CancelButton As Boolean)
    If CancelButton = False Then
        If SaveData(Identity.UserName(), AuditNum, Hours, EffectiveDate) = True Then
            Me.RefreshData
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim intSelectedIndex As Integer
    intSelectedIndex = Me.lstHours.ListIndex
    
    'Check if a row is selected
    If intSelectedIndex < 0 Then
        MsgBox "Please select a row to delete "
        Exit Sub
    End If
    
    'Confirm row deletion
    If MsgBox("Do you really want to delete?", vbYesNo) = vbYes Then
        'Code for Yes button Press
         If DeleteData(Identity.UserName(), _
                    Me.lstHours.Column(GetColumnPosition(Me.lstHours, "AuditNum"), intSelectedIndex + 1), _
                    Me.lstHours.Column(GetColumnPosition(Me.lstHours, "HoursWorked"), intSelectedIndex + 1), _
                    Me.lstHours.Column(GetColumnPosition(Me.lstHours, "WorkDate"), intSelectedIndex + 1)) = True Then
            MsgBox "Entry deleted"
            Me.RefreshData
         Else
            MsgBox "Error deleting entry"
        End If
    End If

End Sub

Private Function DeleteData(strUserID As String, intAuditNum As Integer, dblHours As Double, dtEffectiveDate As Date) As Boolean

    Dim colPrms As ADODB.Parameters
    Dim prm As ADODB.Parameter
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
   
    On Error GoTo ErrHandler
    
    
    'Put validation code here?
    Dim cmd As ADODB.Command
    
    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_ADMIN_USER_Hours_Delete"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_ADMIN_USER_Hours_Delete"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pUserID") = strUserID
    cmd.Parameters("@pAuditNum") = intAuditNum
    cmd.Parameters("@pWorkDate") = dtEffectiveDate
    cmd.Parameters("@pHoursWorked") = dblHours
    With cmd.Parameters("@pHoursWorked")
        .Precision = 8
        .NumericScale = 5
    End With
    
    
    iResult = myCode_ADO.Execute(cmd.Parameters)

    'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        DeleteData = False
        MsgBox "DeleteData", "Error updating Hours - " & strErrMsg
    Else
        DeleteData = True
    End If
    
Exit_Function:
    Set myCode_ADO = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    DeleteData = False
    Resume Exit_Function
End Function

Private Sub lstHours_DblClick(Cancel As Integer)
    Dim intSelectedIndex As Integer

    On Error GoTo ErrHandler
    
     Set frmPopup = New Form_frm_ADMIN_User_Hours_Popup
     
     intSelectedIndex = Me.lstHours.ListIndex
     
     frmPopup.Hours = Me.lstHours.Column(GetColumnPosition(Me.lstHours, "HoursWorked"), intSelectedIndex + 1)
     frmPopup.EffectiveDate = Me.lstHours.Column(GetColumnPosition(Me.lstHours, "WorkDate"), intSelectedIndex + 1)
     frmPopup.AuditNum = Me.lstHours.Column(GetColumnPosition(Me.lstHours, "AuditNum"), intSelectedIndex + 1)
     frmPopup.RefreshData
     ShowFormAndWait frmPopup


ErrHandler_Exit:
    Set frmPopup = Nothing
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ErrHandler_Exit
End Sub
Private Function SaveData(strUserID As String, intAuditNum As Integer, dblHours As Double, dtEffectiveDate As Date) As Boolean

    Dim colPrms As ADODB.Parameters
    Dim prm As ADODB.Parameter
    Dim LocCmd As New ADODB.Command
    Dim iResult As Integer
    Dim strErrMsg As String
   
    On Error GoTo ErrHandler
    
    
    'Put validation code here?
    Dim cmd As ADODB.Command
    
    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_ADMIN_USER_Hours_Apply"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_ADMIN_USER_Hours_Apply"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pUserID") = strUserID
    cmd.Parameters("@pAuditNum") = intAuditNum
    cmd.Parameters("@pWorkDate") = dtEffectiveDate
    cmd.Parameters("@pHoursWorked") = dblHours
    With cmd.Parameters("@pHoursWorked")
        .Precision = 8
        .NumericScale = 5
    End With

    
    iResult = myCode_ADO.Execute(cmd.Parameters)

    'Make sure there are no errors
    If Not Nz(strErrMsg, "") = "" Then
        SaveData = False
        'Err.Raise 65000, "SaveData", "Error updating Hours - " & strErrMsg
        MsgBox "SaveData", "Error updating Hours - " & strErrMsg
    Else
        SaveData = True
    End If
    
Exit_Function:
    Set myCode_ADO = Nothing
    Exit Function

ErrHandler:
    'Rollback anything we did up until this point
    'Return false so the whole save event can rollback.
    SaveData = False
    Resume Exit_Function
End Function

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox "Database Error  - " & ErrMsg, vbOKOnly + vbCritical
End Sub


Private Sub txtFromDate_Enter()
    mReturnDate = Me.txtFromDate.Value
End Sub

Private Sub txtFromDate_LostFocus()
    'Prevent user from entering non-date data
    If Not IsDate(Me.txtFromDate) Then
        MsgBox "Please enter valid date: MM/DD or MM/DD/YYYY", vbOKOnly + vbCritical
        Me.txtFromDate.Value = mReturnDate
        Exit Sub
    End If
    
    'Prevent user from changing FROM DATE that is less than TO DATE
    If CDate(Me.txtFromDate.Value) > CDate(Me.txtToDate.Value) Then
        MsgBox "From Date cannot be greater than To Date", vbOKOnly + vbCritical
        Me.txtFromDate.Value = mReturnDate
    End If
End Sub

Private Sub txtToDate_Enter()
    mReturnDate = Me.txtToDate.Value
End Sub

Private Sub txtToDate_LostFocus()
    'Prevent user from entering non-date data
    If Not IsDate(Me.txtToDate) Then
        MsgBox "Please enter valid date: MM/DD or MM/DD/YYYY", vbOKOnly + vbCritical
        Me.txtToDate.Value = mReturnDate
        Exit Sub
    End If
    
    'Prevent user from changing FROM DATE that is less than TO DATE
    If CDate(Me.txtFromDate.Value) > CDate(Me.txtToDate.Value) Then
        MsgBox "To Date cannot be less than From Date", vbOKOnly + vbCritical
        Me.txtToDate.Value = mReturnDate
    End If
End Sub
