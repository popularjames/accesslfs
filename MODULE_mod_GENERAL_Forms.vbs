Option Explicit


Private Const ClassName As String = "mod_GENERAL_Forms"


Declare Function GetSystemMetrics32 Lib "user32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long


'' Last modified: 05/29/2012
'' History:
''  - 05/29/2012  - KD: Fixed IsOpen to actually use the object type param


Public Sub NewQuickLook(strSearchType As String, strCaption As String)
    Dim frmNew As New Form_frm_General_QuickLookup
    Set frmNew = New Form_frm_General_QuickLookup
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.SearchType = strSearchType
    
    frmNew.visible = True
    frmNew.lblAppTitle.Caption = "SEARCH - " & strCaption
    frmNew.RefreshData
    
End Sub

Public Function UserAccess_Check(frm As Form, Optional ParentAppID As String = "") As Integer
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim ctrl As Control
    Dim strSubFormCtrlName As String
    Dim strCtrlName As String
    Dim iUserAccess As Integer
    Dim strAction As String
    
    Dim bAllowChange As Boolean
    Dim bAllowAdd As Boolean
    Dim bAllowDelete As Boolean
    Dim bAllowView As Boolean
    
    Dim iAppPermission As Integer
    
    Dim strAppID As String
    
    If ParentAppID <> "" Then strAppID = ParentAppID Else strAppID = frm.frmAppID
    
    iAppPermission = GetAppPermission(strAppID)
    bAllowAdd = (iAppPermission And gcAllowAdd)
    bAllowDelete = (iAppPermission And gcAllowDelete)
    bAllowChange = (iAppPermission And gcAllowChange) Or bAllowAdd Or bAllowDelete
    bAllowView = (iAppPermission And gcAllowView) Or bAllowChange
    
    If IsSubForm(frm) Then
        strSubFormCtrlName = ""
        For Each ctrl In frm.Parent.Form.Controls
            If ctrl.ControlType = acSubform Then
                If ctrl.SourceObject = frm.Name Then
                    strSubFormCtrlName = ctrl.Name
                    Exit For
                End If
            End If
        Next
    End If
    
    
    If (iAppPermission = gcLocked) Or (bAllowView = False) Then
        frm.visible = False
        MsgBox "You do not have permission to view this form.  Please contact your system admin", vbInformation
        If IsSubForm(frm) Then
            If strSubFormCtrlName <> "" Then
                frm.Parent.Form(strSubFormCtrlName).visible = False
            End If
        Else
            On Error Resume Next
            DoCmd.Close acForm, frm.Name
            
        End If
        UserAccess_Check = 0
        Exit Function
    End If
    
    frm.AllowAdditions = bAllowAdd
    frm.AllowDeletions = bAllowDelete
    
    '' 20120821 KD Change to speed things up!
    Set MyAdo = New clsADO
    With MyAdo
        .ConnectionString = GetConnectString("ADMIN_Form_Security")
        .SQLTextType = sqltext
        .sqlString = "select * from ADMIN_Form_Security where FormName = '" & frm.Name & "'"
        Set rs = .ExecuteRS
    End With
'    Set rs = CurrentDb.OpenRecordSet("select * from ADMIN_Form_Security where FormName = '" & frm.Name & "'")
    '' END 20120821 KD Change to speed things up!
    
    While Not rs.EOF
        strCtrlName = rs("CtrlName")
        iUserAccess = rs("UserAccess")
        strAction = rs("Action")
        
        If strAction = "Visible" Then
            frm(strCtrlName).visible = (iUserAccess And iAppPermission)
        ElseIf strAction = "Enable" Then
            'Use ControlType to determine the Type of Control
            Select Case frm(strCtrlName).ControlType
                Case acCommandButton: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acOptionButton: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acCheckBox: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acBoundObjectFrame: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acTextBox: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acListBox: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acOptionGroup: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acComboBox: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acSubform: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acObjectFrame: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acCustomControl: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acToggleButton: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
                Case acTabCtl: frm(strCtrlName).Enabled = (iUserAccess And iAppPermission)
            End Select
        End If
        
        rs.MoveNext
    Wend
    
    For Each ctrl In frm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) And bAllowChange = False Then
            ctrl.Locked = True
        End If
    Next
        
    If bAllowChange Then
        iAppPermission = (iAppPermission Or gcAllowChange)
    End If
    UserAccess_Check = iAppPermission
End Function


Public Function Account_Check(frm As Form, Optional SourceTable As String = "") As Integer
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim frmAcctSelect As Form_frm_ADMIN_Account_Selection
'    Dim lRcdCnt As Long
    
    If gintAccountID = -1 Then
        ' no account setup
        MsgBox "You are not set up for any account.  Please consult with your administrator", vbCritical
    ElseIf gintAccountID = 0 Then
        ' redisplay screen for user to select an account if neccessary
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "select * from ADMIN_User_Account where UserID = '" & Identity.UserName() & "'"
        Set rs = MyAdo.OpenRecordSet
    
'        Set rs = MYADO.ExecuteRS
'        If rs Is Nothing Then
'            lRcdCnt = 0
'        Else
'            lRcdCnt = rs.RecordCount
'        End If
        
    
        If rs.recordCount = 0 Then
            MsgBox "You are not set up for any account.  Please consult with your administrator", vbCritical
            gintAccountID = -1
            gstrAcctAbbrev = ""
            gstrAcctDesc = ""
        ElseIf rs.recordCount = 1 Then
            gintAccountID = rs("AccountID")
            MyAdo.sqlString = "select * from ADMIN_Client_Account where AccountID = " & gintAccountID
            Set rs = MyAdo.OpenRecordSet
            gstrAcctAbbrev = rs("AcctAbbrev")
            gstrAcctDesc = rs("AcctDesc")
        Else
            Set frmAcctSelect = New Form_frm_ADMIN_Account_Selection
            ColObjectInstances.Add frmAcctSelect, frmAcctSelect.hwnd & ""
            ShowFormAndWait frmAcctSelect
        End If
    End If
    
    
    ' double check and set form values
    If gintAccountID > 0 Then
        ' set form caption
        If frm.Caption = "" Then
            frm.Caption = gstrAcctDesc
        Else
            frm.Caption = gstrAcctDesc & " - " & frm.Caption
        End If
    
        ' set form row source
        If SourceTable <> "" Then
            frm.RecordSource = "select * from " & SourceTable & " where AccountID = " & gintAccountID
        End If
    
    
        gstrProfileID = GetUserProfile()
    
    Else
        frm.Section("FormHeader").visible = False
        frm.Section("FormFooter").visible = False
        frm.Section("Detail").visible = False
        gstrProfileID = ""
    End If
    
End Function

Public Sub NewProvider(strCnlyProvID As String, strCaption As String)
    Dim frmNew As New Form_frm_PROV_Hdr
    Set frmNew = New Form_frm_PROV_Hdr
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.cnlyProvID = strCnlyProvID
    
    frmNew.visible = True
    frmNew.LoadProvider
    
End Sub

Public Sub NewManualAdjustment(strCnlyClaimNum As String, strCaption As String)
    Dim frmNew As New Form_frm_COLL_CNLY_Adj_Main
    Set frmNew = New Form_frm_COLL_CNLY_Adj_Main
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.CnlyClaimNum = strCnlyClaimNum
    
    frmNew.visible = True
    frmNew.RefreshData
    
End Sub

Public Sub NewConcept(strConceptID As String, strCaption As String)
    Dim frmNew As New Form_frm_CONCEPT_Hdr
    Set frmNew = New Form_frm_CONCEPT_Hdr
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.FormConceptID = strConceptID
    
    frmNew.visible = True
    frmNew.RefreshData
    
End Sub
Public Sub NewMainSearch(strAppID As String, strGridSource As String, strCaption As String)

    Dim frmNew As Form_frm_GENERAL_Search
    Set frmNew = New Form_frm_GENERAL_Search

    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "

    frmNew.frmAppID = strAppID
    frmNew.GridSource = strGridSource
    frmNew.RefreshMain
    frmNew.Caption = strCaption & ": Search"
    frmNew.visible = True

End Sub


Public Sub NewMainTab(strRowSource As String, strCnlyClaimNum As String, strCaption As String)
    Dim frmNew As Form_frm_GENERAL_Tab
    Set frmNew = New Form_frm_GENERAL_Tab
    
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
    frmNew.Caption = strCaption & ": ClaimNum : " & strCnlyClaimNum
    
    frmNew.CnlyRowSource = strRowSource
       
    frmNew.visible = True
    frmNew.RefreshData
End Sub


Public Sub NewMain(strCnlyClaimNum As String, strCaption As String)
    Dim frmNew As Form_frm_AUDITCLM_Main
    Set frmNew = New Form_frm_AUDITCLM_Main
    
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
    frmNew.Caption = strCaption & ": ClaimNum : " & strCnlyClaimNum
    frmNew.CnlyClaimNum = strCnlyClaimNum
    
    frmNew.visible = True
    If frmNew.CnlyClaimNum <> "" Then frmNew.LoadData
End Sub

'Damon 06/03/08
'General ComboBox Filling
'This adds the functionality of being able to specify a default selected value

'Taken to get around modal debacle
 Public Function ShowFormAndWait_Function(frm As Form) As Boolean
     Dim blnCancelled   As Boolean
     Dim lngLoop As Long
     Dim strName As String
     Const adhcInterval As Long = 1000
    
     strName = frm.Name
    
     ShowFormAndWait_Function = False
     frm.visible = True
     
     Do
         If lngLoop Mod adhcInterval Then
            DoEvents
            'Is it still Open?
            If Not IsOpen(strName) Then
                blnCancelled = True
                Exit Do
            End If
            
            'Is it still visible?
            If Not frm.visible Then
                blnCancelled = False
                Exit Do
            End If
                
           lngLoop = 0
          End If
          lngLoop = lngLoop + 1
     Loop
     ShowFormAndWait_Function = Not blnCancelled
 End Function


Public Function CreateNewForm(FormName As String) As Integer
    Dim bFormExist As Boolean
    
    bFormExist = False
    On Error Resume Next
    If CurrentProject.AllForms(FormName).Name <> "" Then
        bFormExist = True
    End If
    
    If bFormExist Then
        MsgBox "Form " & FormName & " already exists in database.  Do you want to overwrite?", vbYesNo
    End If
        
    CreateNewForm = 1
    
End Function


Public Function IsSubForm(frm As Form) As Boolean
    Dim strName As String
Dim iCurrentSetting As Integer

    ' Since I "live" with 'break on all errors' option turned on,
    ' let's turn it off:
    iCurrentSetting = Application.GetOption("Error Trapping")
    Application.SetOption "Error Trapping", 2
    On Error Resume Next
    strName = frm.Parent.Name
    IsSubForm = (Err.Number = 0)
    Err.Clear
    Application.SetOption "Error Trapping", iCurrentSetting
End Function


Public Function IsLoaded(strFormName As String, Optional lngType As AcObjectType = acForm) As Boolean
    IsLoaded = (SysCmd(acSysCmdGetObjectState, lngType, strFormName) <> 0)
End Function


Public Function IsOpen(strFormName As String, Optional lngType As AcObjectType = acForm) As Boolean
    IsOpen = SysCmd(acSysCmdGetObjectState, lngType, strFormName)
End Function


Public Sub ShowFormAndWait(frm As Form)
     Dim bFormClosed   As Boolean
     Dim strFormName As String
    
     strFormName = frm.Name
     frm.visible = True
     
     Do
        'Is it still Open?
        If IsLoaded(strFormName) Then
            DoEvents
            Wait 1
        Else
            bFormClosed = True
        End If
     Loop Until bFormClosed
End Sub
'Alex C 2/12/2012 - Added to launch Customer Service for an open claim
Public Function LaunchNewCustClaimEvent(strCnlyClaimNum As String) As String
    Dim frmNew As New Form_frm_CUST_Main
    Set frmNew = New Form_frm_CUST_Main
    
    'Does user have administration access to CustomerService?
    If frmNew.AppPermission <> 0 Then
    
        ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
        frmNew.CnlyClaimNum = strCnlyClaimNum
    
        If frmNew.CreateEventFromClaim = True Then
            frmNew.visible = True
            LaunchNewCustClaimEvent = ""
            Exit Function
        End If
    End If
            
    'User doesn't have access - close the form
    DoCmd.Close acForm, "frm_cust_main", acSaveNo
    Set frmNew = Nothing
    LaunchNewCustClaimEvent = ""
    Exit Function
    
End Function
'Alex C 3/6/2012 - Added to launch Customer Service for an open provider
Public Function LaunchNewCustProviderEvent(strCnlyProvID As String) As String
    Dim frmNew As New Form_frm_CUST_Main
    Set frmNew = New Form_frm_CUST_Main
    
    'Does user have administration access to CustomerService?
    If frmNew.AppPermission <> 0 Then
    
        ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
        frmNew.cnlyProvID = strCnlyProvID
    
        If frmNew.CreateEventFromProvider = True Then
            frmNew.visible = True
            LaunchNewCustProviderEvent = ""
            Exit Function
        End If
    End If
            
    'User doesn't have access - close the form
    DoCmd.Close acForm, "frm_cust_main", acSaveNo
    Set frmNew = Nothing
    LaunchNewCustProviderEvent = ""
    Exit Function
    
End Function


Public Sub ReportingAccessForm(strSearchType As String, ReportTable As String, ReportFilter As String)

    Dim frmNew As New Form_frm_RPT_AccessForm
    Set frmNew = New Form_frm_RPT_AccessForm
    
    ColObjectInstances.Add Item:=frmNew, Key:=frmNew.hwnd & " "
   
    frmNew.SearchType = strSearchType
    frmNew.RecordSource = ReportTable
    
    Dim db As DAO.Database
    Dim tdfld As DAO.TableDef
    Dim fld As DAO.Field
    
    Dim FieldCounter As Integer
 
    Set db = CurrentDb()
    Set tdfld = db.TableDefs(ReportTable)
    
    frmNew.Caption = "frm_" & ReportTable & " (AccessForm)"
    
    FieldCounter = 1
    
    For Each fld In tdfld.Fields    'loop through all the fields of the tables
    
        frmNew("Text" & FieldCounter).ControlSource = fld.Name
        frmNew("Label" & FieldCounter).Caption = fld.Name
        FieldCounter = FieldCounter + 1
    Next
    
    Dim i As Integer
    
    For i = FieldCounter To 270
        frmNew("Text" & i).Properties("ColumnHidden") = True
    Next

    frmNew.FilterOn = True
    frmNew.filter = ReportFilter
    frmNew.visible = True

    'DoCmd.Close acForm, Me.Name
    
End Sub

Function FormExists(ByVal FormName As String) As Boolean
    'check if form exists
    Dim frmCurr As AccessObject
    FormExists = False
    For Each frmCurr In Application.CurrentProject.AllForms
        If UCase(frmCurr.Name) = UCase(FormName) Then
            FormExists = True
            Exit Function
        End If
    Next frmCurr
End Function




Public Function AppAccess_Check_Passive(strAppID As String) As Integer
    Dim MyAdo As clsADO
    Dim rs As ADODB.RecordSet
    Dim ctrl As Control
    Dim strSubFormCtrlName As String
    Dim strCtrlName As String
    Dim iUserAccess As Integer
    Dim strAction As String
    
    Dim bAllowChange As Boolean
    Dim bAllowAdd As Boolean
    Dim bAllowDelete As Boolean
    Dim bAllowView As Boolean
    
    Dim iAppPermission As Integer
    
        
    iAppPermission = GetAppPermission(strAppID)
    bAllowAdd = (iAppPermission And gcAllowAdd)
    bAllowDelete = (iAppPermission And gcAllowDelete)
    bAllowChange = (iAppPermission And gcAllowChange) Or bAllowAdd Or bAllowDelete
    bAllowView = (iAppPermission And gcAllowView) Or bAllowChange
    
    
    If (iAppPermission = gcLocked) Or (bAllowView = False) Then
            AppAccess_Check_Passive = 0
            Exit Function
    End If
    
    If bAllowChange Then
        iAppPermission = (iAppPermission Or gcAllowChange)
    End If
    AppAccess_Check_Passive = iAppPermission
End Function

Function MonitorHeight() As Integer
Dim w As Long, h As Long
    MonitorHeight = GetSystemMetrics32(1) ' height in points
End Function


Function MonitorWidth() As Integer
Dim w As Long, h As Long
    MonitorWidth = GetSystemMetrics32(0) ' width in points
End Function


'It retrieves the setting of the particular type from the DB
Public Function GetUserSetting(UserID As String, SettingName As String) As String
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    Dim myCode_ADO As New clsADO
    Dim objparameter  As Object
    
    GetUserSetting = "1"
Exit Function
    
Stop    ' not implemented in CMS yet

    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "USP_ADMIN_GET_USER_SETTING"
    myCode_ADO.SQLTextType = StoredProc

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "USP_ADMIN_GET_USER_SETTING"
    Set objparameter = cmd.CreateParameter("@SETTINGVALUE", adVarChar, adParamOutput, 100, "")
    'cmd.Parameters.Append (objparameter)
    cmd.Parameters("@PSETTINGNAME") = SettingName
    cmd.Parameters("@PUSERID") = UserID
    Set rs = myCode_ADO.ExecuteRS(cmd.Parameters)
    GetUserSetting = Nz(cmd.Parameters("@PSETTINGVALUE"))
End Function

Public Function SetUserSetting(UserID As String, SettingName As String, Value As String) As String
    Dim rs As ADODB.RecordSet
    Dim cmd As ADODB.Command
    Dim myCode_ADO As New clsADO
    Dim objparameter  As Object
    
Stop    ' not implemented in CMS yet
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "USP_ADMIN_SET_USER_SETTING"
    myCode_ADO.SQLTextType = StoredProc

    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "USP_ADMIN_SET_USER_SETTING"

    cmd.CreateParameter "@PERRMSG", adVarChar, adParamOutput, 100, ""
    cmd.Parameters("@PSETTINGVALUE") = Value
    cmd.Parameters("@PSETTINGNAME") = SettingName
    cmd.Parameters("@PUSERID") = UserID
    Set rs = myCode_ADO.ExecuteRS(cmd.Parameters)
    
    SetUserSetting = Nz(cmd.Parameters("@PERRMSG"))
End Function




Public Function PopulateListViewFromRs(oLView As Object, oRs As ADODB.RecordSet, Optional dctColHeaderNames As Scripting.Dictionary, Optional dctColHeaderNums As Scripting.Dictionary) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oLI As ListItem
Dim oLV As ListView
Dim oHdr As ColumnHeader
Dim oFld As ADODB.Field
Dim iCols As Integer
'Dim dctColHeaderNames As Scripting.Dictionary
'Dim dctColHeaderNums As Scripting.Dictionary

    strProcName = ClassName & ".PopulateListViewFromRs"
    
    
    If oRs Is Nothing Then
        GoTo Block_Exit
    End If
    
    oLView.ListItems.Clear
    
    If oRs.EOF And oRs.BOF Then
        GoTo Block_Exit
    End If
    
    Set dctColHeaderNames = New Scripting.Dictionary
    Set dctColHeaderNums = New Scripting.Dictionary
    
    '' Set up our headers:
    ' remember, first one will be the 'Text' the rest will be subitems
    oLView.ColumnHeaders.Clear
    For Each oFld In oRs.Fields
        iCols = iCols + 1
        Set oHdr = oLView.ColumnHeaders.Add(, oFld.Name, oFld.Name, 1500)
        If dctColHeaderNames.Exists(oFld.Name) = False Then
            dctColHeaderNames.Add oFld.Name, iCols
            dctColHeaderNums.Add iCols, oFld.Name
        Else
            Stop ' dup!!
        End If
    Next
    
    While Not oRs.EOF
        Set oLI = oLView.ListItems.Add(, , oRs(dctColHeaderNums.Item(1)).Value)
        
        For iCols = 2 To oRs.Fields.Count
            oLI.SubItems(iCols - 1) = oRs(dctColHeaderNums.Item(iCols)).Value
        Next
        
        oRs.MoveNext
    Wend
    
    PopulateListViewFromRs = True
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function