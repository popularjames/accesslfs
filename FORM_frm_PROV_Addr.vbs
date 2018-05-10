Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frmCalendar As Form_frm_GENERAL_Calendar
Attribute frmCalendar.VB_VarHelpID = -1
Private WithEvents myMouseWheel As clsMouseWheel
Attribute myMouseWheel.VB_VarHelpID = -1

Public Event RecordChanged()

Private mrsPROVAddr As ADODB.RecordSet
Private mrsPROVAddrPortal As ADODB.RecordSet

Public mrsProvAddrChanged As ADODB.RecordSet
Public mbProvAddrChanged As Boolean


Private mdReturnDate As Date
Private mbNewRecord As Boolean
Private mbDirty As Boolean
Private mstrProvID As String
Private mbDisableMouseWheel As Boolean
Private mbRecordChanged As Boolean
Private mbRecordLocked As Boolean
Private CurrRecord As Long
Private miAppPermission As Integer
Const CstrFrmAppID As String = "ProvAddr"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property


Public Property Let DisableMousewheel(data As Boolean)
    mbDisableMouseWheel = data
    
    If mbDisableMouseWheel Then
        Set myMouseWheel = New clsMouseWheel
        Set myMouseWheel.Form = Me

        'Subclass the current form by calling
        'the SubClassHookForm method in the class
        myMouseWheel.HookForm
    End If
End Property

Public Property Set AddrRecordSource(data As ADODB.RecordSet)
    Set mrsPROVAddr = data
End Property

Public Property Get AddrRecordSource() As ADODB.RecordSet
    Set AddrRecordSource = mrsPROVAddr
End Property

Public Property Set PortalRecordSource(data As ADODB.RecordSet)
    Set mrsPROVAddrPortal = data
End Property

Public Property Get PortalRecordSource() As ADODB.RecordSet
    Set PortalRecordSource = mrsPROVAddrPortal
End Property


Public Property Get ProvID() As String
    ProvID = mstrProvID
End Property

Public Property Let ProvID(data As String)
    mstrProvID = data
End Property

Public Property Let NewAddrRecord(data As Boolean)
    mbNewRecord = data
    If mbNewRecord Then
        If mrsPROVAddr.recordCount > 0 Then
            mrsPROVAddr.MoveLast
        End If
        mrsPROVAddr.AddNew
    End If
End Property

Public Property Get NewAddrRecord() As Boolean
    NewAddrRecord = mbNewRecord
End Property

Property Let RecordLocked(data As Boolean)
    mbRecordLocked = data
    If mbRecordLocked Then
        cmdSave.Enabled = False
    End If
End Property

Public Sub RefreshMain()
    On Error GoTo ErrHandler
    
    
    Set Me.RecordSet = mrsPROVAddr
    Me.Caption = gstrAcctDesc & " - Provider Address Maintenance"
    If mbRecordLocked Then
        mstrProvID = cnlyProvID
        Me.lblAppTitle.Caption = "LOCKED; CnlyProvID: " & Nz(mstrProvID, "")
    ElseIf mbNewRecord Then
        Me.lblAppTitle.Caption = "NEW ADDRESS; CnlyProvID: " & Nz(mstrProvID, "")
        cnlyProvID = mstrProvID
        Me.AddrType.SetFocus
    Else
        mstrProvID = cnlyProvID
        Me.lblAppTitle.Caption = "CnlyProvID: " & Nz(mstrProvID, "")
    End If
    
    check_PortalAddr
    
Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_AUDITCLM_Main : RefreshMain"
End Sub
Private Sub check_PortalAddr()

    Dim MyAdo As clsADO
    Set MyAdo = New clsADO
    Set mrsProvAddrChanged = New ADODB.RecordSet
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    'MsgBox (Me.cnlyProvID & "|" & Me.AddrType & "|" & Me.EffDt & "|" & Me.TermDt)
    MyAdo.sqlString = "exec CMS_Auditors_Code.dbo.usp_GetProvAddress '" & Me.cnlyProvID & "','" & Me.AddrType & "','" & Me.EffDt & "','" & Me.TermDt & "'"
    Set mrsProvAddrChanged = MyAdo.OpenRecordSet
    If mrsProvAddrChanged.recordCount > 0 Then mbProvAddrChanged = True Else mbProvAddrChanged = False
    
    TglAddr.Enabled = mbProvAddrChanged
    TglAddr.visible = mbProvAddrChanged
    If mbProvAddrChanged Then
        TglAddr.Value = False 'Unpressed
        TglAddr.Caption = "Original" 'Showing unchanged record by default
        'Position cursor to the changed record in mrsProvAddrPortal
        mrsPROVAddrPortal.MoveFirst
        While mrsPROVAddrPortal.Fields("PortalAddrID").Value <> mrsProvAddrChanged.Fields("PortalAddrID").Value
            mrsPROVAddrPortal.MoveNext
        Wend
    End If
End Sub
Private Sub cmdExit_Click()
    If mbNewRecord Then
        mrsPROVAddr.Delete
        
        If mrsPROVAddr.recordCount > 0 Then
            mrsPROVAddr.MoveLast
        End If
        mbNewRecord = False     ' set this so that when we exit the form, it won't be trigger again in the form close event
        DoCmd.Close acForm, Me.Name
    ElseIf mbNewRecord = False And Me.RecordSource <> "" Then
        If Me.Dirty Then
            If MsgBox("Record has been changed. Are you sure you want to exit? ", vbYesNo) = vbYes Then
                DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
                DoCmd.Close acForm, Me.Name
            End If
        Else
            DoCmd.Close acForm, Me.Name
        End If
    Else
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub cmdProvAddressDup_Click()
On Error GoTo ErrorHandler
    Dim strCnlyProvID As String
    Dim strAddrTypeToDuplicate As String
    Dim strAddrTypeNew As String
    Dim strErrMsg As String
    
    Dim cmd As ADODB.Command
    Dim myCode_ADO As clsADO

    Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    


    
    strCnlyProvID = Trim(Me.cnlyProvID.Value)
    strAddrTypeToDuplicate = Trim(Me.AddrType.Value)
    strAddrTypeNew = Trim(Replace(Replace(Nz(Me.cbAddrTypeNew.Value, ""), Chr(10), ""), Chr(13), ""))
        Debug.Print "strCnlyProvID:" & strCnlyProvID
        Debug.Print "strAddrTypeToDuplicate:" & strAddrTypeToDuplicate
        Debug.Print "strAddrTypeNew:" & strAddrTypeNew

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "ProvAddressDup"
    cmd.Parameters.Refresh
    cmd.Parameters("@pCnlyProvID").Value = strCnlyProvID
    cmd.Parameters("@pAddrTypeToDuplicate").Value = strAddrTypeToDuplicate
    cmd.Parameters("@pAddrTypeNew").Value = strAddrTypeNew

    cmd.Execute

'    If cmd.Parameters("@pErrMsg") <> "" Then
'        MsgBox (cmd.Parameters("@pErrMsg"))
'    End If

    
    MsgBox ("Done. Close window and select provider address to see update.")
    
EXIT_HERE:
' Release used objects
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
   
ErrorHandler:
    With Err
        MsgBox ("Sub encountered error. Error in  " & .Number & ". " & .Description)
    End With

    Resume EXIT_HERE
End Sub

Private Sub cmdSave_Click()
    If mbDirty Or mbNewRecord Then
        mbRecordChanged = True
    End If
    SaveData
'    If mbProvAddrChanged Then updatePortalsStatus
End Sub

'Private Sub updatePortalsStatus()
''Not taking into account the fact that the main form could be saved without accepting address changes. There's no way to know for sure which addresses were changed/(and then reverted) from the main window
'    Dim myPortalADO As clsADO
'    Set myPortalADO = New clsADO
'    myPortalADO.ConnectionString = GetConnectString("v_DATA_Database")
'    myPortalADO.BeginTrans
'    myPortalADO.SQLstring = "update CMS_AUDITORS_Claims.dbo.PROV_ADDRESS_PortalPending SET ReqStatusCode = '02', LastUpdateDt = current_timestamp where AutoId = '" & mrsPROVAddrChanged.Fields("AutoId").Value & "' AND cnlyProvID = '" & mrsPROVAddrChanged.Fields("cnlyProvID").Value & "'"
'    myPortalADO.SQLTextType = sqltext
'    myPortalADO.Execute
'    myPortalADO.CommitTrans
'
'End Sub

Private Sub SaveData()
    Dim rsTemp As ADODB.RecordSet
    Dim strCriteria As String
    Dim i, iCurrRecord As Integer
    

    On Error GoTo Err_SaveData
    
    If TermDt & "" = "" Then TermDt = "12/31/9999"
    If Nz(Me.AddrType, "") = "" Then
        MsgBox "'Addr Type' field cannot be blank.", vbOKOnly + vbInformation
        Me.AddrType.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.EffDt, "") = "" Then
        MsgBox "'Eff Dt' field cannot be blank.", vbOKOnly + vbInformation
        Me.EffDt.SetFocus
        Exit Sub
        
    ElseIf Me.EffDt > Me.TermDt Then
        MsgBox "'Eff Dt' can not be more than 'Term Dt'", vbOKOnly + vbInformation
        Me.TermDt.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.Addr01, "") = "" Then
        MsgBox "'Addr01' field cannot be blank.", vbOKOnly + vbInformation
        Me.Addr01.SetFocus
        Exit Sub
    
    ElseIf Nz(Me.City, "") = "" Then
        MsgBox "'City' field cannot be blank.", vbOKOnly + vbInformation
        Me.City.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.State, "") = "" Then
        MsgBox "'State' field cannot be blank.", vbOKOnly + vbInformation
        Me.State.SetFocus
        Exit Sub
        
    ElseIf Nz(Me.Zip, "") = "" Then
        MsgBox "'Zip' field cannot be blank.", vbOKOnly + vbInformation
        Me.Zip.SetFocus
        Exit Sub
    End If
    
    
    Set rsTemp = mrsPROVAddr.Clone
    rsTemp.MoveFirst
    iCurrRecord = Me.CurrentRecord
    
    For i = 1 To rsTemp.recordCount
        With rsTemp
            
            'Checks to see if effective period on new or updated Provider Address already exists
            If i <> iCurrRecord And AddrType = !AddrType Then
                If EffDt >= !EffDt And EffDt <= Nz(!TermDt, "#12/31/9999#") Then
                    MsgBox "Error: Another " & AddrType.Column(1) & " address is already effective from " & !EffDt & " to " & Nz(!TermDt, "12/31/9999") & "." & vbCrLf & vbCrLf & _
                    "If you would like to make this address effective from " & EffDt & ", please modify the Term Dt of the other " & AddrType.Column(1) & " address accordingly."
                    Exit Sub
                ElseIf EffDt < !EffDt And Nz(TermDt, "#12/31/9999#") >= !EffDt Then
                    MsgBox "Error: Another " & AddrType.Column(1) & " address is already effective from " & !EffDt & " to " & Nz(!TermDt, "12/31/9999") & "." & vbCrLf & vbCrLf & _
                    "If you would like to terminate this address on " & TermDt & ", please modify the Eff Dt of the other " & AddrType.Column(1) & " address accordingly."
                    Exit Sub
                End If
            End If
            
            .MoveNext
        End With
    Next i

    'If MsgBox("Would you like to save address data for Provider - " & CnlyProvID & "?", vbYesNo + vbQuestion) = vbYes Then
        If Not mrsPROVAddr Is Nothing Then
            Me.LastUpdateDt = Now()
            Me.LastUpdateUser = Identity.UserName
        End If
    'End If
    
    mbNewRecord = False
    mbDirty = False
    DoCmd.Close acForm, Me.Name


Exit_SaveData:
    Set rsTemp = Nothing
    Exit Sub

Err_SaveData:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
    Resume Exit_SaveData
    
End Sub

Private Sub cmdEffDt_Click()
    On Error GoTo Err_btnChkDt_Click
    
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Nz(Me.EffDt, Date)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    If mdReturnDate <> "12:00:00 AM" Then
        Me.EffDt = mdReturnDate
    End If

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click

End Sub

Private Sub cmdTermDt_Click()
    On Error GoTo Err_btnChkDt_Click
    Set frmCalendar = New Form_frm_GENERAL_Calendar
    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
    frmCalendar.DatePassed = Switch(IsNull(Me.TermDt), Date, Me.TermDt = "12/31/9999", Date, 1, Me.TermDt)
    frmCalendar.RefreshData
    ShowFormAndWait frmCalendar
    
    If mdReturnDate <> "12:00:00 AM" Then
        Me.TermDt = mdReturnDate
    End If

Exit_btnChkDt_Click:
    Exit Sub

Err_btnChkDt_Click:
    MsgBox Err.Description
    Resume Exit_btnChkDt_Click

End Sub


Public Sub AddNewAddress()
    If (miAppPermission And gcAllowAdd) Then
        If Nz(Me.ProvID, "") <> "" Then
            Me.NewAddrRecord = True
            RefreshMain
            Me.Detail.visible = True
            Me.cmdSave.Enabled = True
        Else
            MsgBox "Please set the CnlyProvID first"
        End If
    Else
        MsgBox "You don't have permission to add provider address data."
        Me.Detail.visible = False
        Me.cmdSave.Enabled = False
    End If
End Sub



Private Sub Form_Current()
    If Not (Me.RecordSet Is Nothing) Then
        If Me.Dirty Then mbDirty = True
    End If
    'check_PortalAddr
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    mbDirty = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0
End Sub

Private Sub Form_Load()
    Me.Caption = "Provider Address Data Entry Form"

    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
    RefreshComboBox "SELECT AddrType, Description  FROM PROV_XREF_Address_Code ", Me.AddrType, "", ""
    
    mbRecordChanged = False
    mbDisableMouseWheel = False
    mbDirty = False
End Sub


Private Sub Form_Close()
    On Error Resume Next
    If mbNewRecord Then
        mrsPROVAddr.Delete
        mrsPROVAddr.UpdateBatch
        If mrsPROVAddr.recordCount > 0 Then
            mrsPROVAddr.MoveLast
        End If
    ElseIf mbNewRecord = False And mbDirty Then
        DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    End If
    
    If mbDisableMouseWheel Then
        myMouseWheel.UnHookForm
        Set myMouseWheel.Form = Nothing
        Set myMouseWheel = Nothing
    End If
    
    If mbRecordChanged Then
        RaiseEvent RecordChanged
    End If
    
    On Error Resume Next
    RemoveObjectInstance Me
End Sub

Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub myMouseWheel_MouseWheel(Cancel As Integer)
     'This is the event procedure where you can
     'decide what to do when the user rolls the mouse.
     'If setting Cancel = True, we disable the mouse wheel
     'in this form.
     
     Cancel = True
End Sub

Private Sub frmCalendar_DateSelected(ReturnDate As Date)
    mdReturnDate = ReturnDate
End Sub

Private Sub TglAddr_Click()
Dim fld, frmfld 'As Field 'Andrew Commented out to leave these variables undefined so that the Change Address will work
Dim LikeField As Boolean
Dim clrHighlight As Long
Dim ctl As Control
clrHighlight = RGB(255, 255, 0)

'MsgBox mrsPROVAddrChanged.Fields("addrtype").Value
If mbProvAddrChanged Then
    If TglAddr.Value = True Then
        TglAddr.Caption = "Changed"
        mrsPROVAddrPortal.Fields("ReqStatusCode").Value = "02"
    Else
        TglAddr.Caption = "Original"
        For Each ctl In Me.Controls
          If (ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox) Then
            If ctl.BackColor = clrHighlight Then
                ctl.BackColor = RGB(255, 255, 255)
            End If
          End If
        Next ctl
        Me.Undo
        mrsPROVAddrPortal.Fields("ReqStatusCode").Value = "01"
        mbDirty = False
        GoTo Before_Exit
    End If
End If


'Highlight fields that have changed
For Each fld In mrsProvAddrChanged.Fields
    LikeField = False
    For Each frmfld In Me.RecordSet.Fields
        If frmfld.Name = fld.Name Then LikeField = True
    Next frmfld
    If LikeField = True Then
        If Me.RecordSet.Fields(fld.Name).Value <> Trim(fld.Value) And fld.Name <> "AddrId" Then
            Me.Controls(fld.Name).BackColor = RGB(255, 255, 0)
            Me.Controls(fld.Name).Value = fld.Value
        End If
    End If
Next fld
Me.Comments = Me.Comments.Value & ";Provider req from Portal. ID=" & mrsProvAddrChanged.Fields("AutoID") & "|ReceivedOn-" & mrsProvAddrChanged.Fields("ReqDate")
mbDirty = True
Before_Exit:
End Sub
