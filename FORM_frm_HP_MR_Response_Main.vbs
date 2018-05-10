Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents frm_Void_Image As Form_frm_AUDITCLM_References_Update
Attribute frm_Void_Image.VB_VarHelpID = -1
Private frm_HP_Claim_Display As Form_frm_HP_Claim_Display

Private mstrFilter As String
Private mstrOldFilter As String
Private bSQLChange As Boolean
Private mbLoaded As Boolean
Private miRowsSelected As Integer
Private miStartRow As Integer

Public Function Loaded() As Boolean
    Loaded = mbLoaded
End Function

Public Property Get RowsSelected() As Long
    RowsSelected = miRowsSelected
End Property

Public Property Let RowsSelected(ByVal vData As Long)
    miRowsSelected = vData
    Me.txtRecordsSelected = vData
End Property

Public Property Get StartRow() As Long
    StartRow = miStartRow
End Property

Public Property Let StartRow(ByVal vData As Long)
    miStartRow = vData
End Property


Private Sub cmdHoldClaim_Click()
    Dim iWindowHandle As Long
    Dim f As clsWindowHandles
    Dim bFound As Boolean
    
    Dim rs As RecordSet
    Dim i As Long
    Dim strCnlyClaimNum As String
    Dim strClmStatus As String
    Dim strDisplayMsg As String
    
    bFound = False
    
    For Each f In ColWindows
        If f.WindowName = "HP_MR_VOID" Then
            iWindowHandle = f.WindowHandle
            SetForegroundWindow iWindowHandle
            If screen.ActiveForm.hwnd = iWindowHandle Then
                bFound = True
                Exit For
            End If
        End If
    Next
    
    If Not bFound Then
        Set f = New clsWindowHandles
        Set frm_HP_Claim_Display = New Form_frm_HP_Claim_Display
        f.WindowHandle = frm_HP_Claim_Display.hwnd
        f.WindowName = "HP_CLAIM_DISPLAY"
        ColWindows.Add f, f.WindowHandle & ""
        frm_HP_Claim_Display.visible = True
    End If
    
    ' CYCLE THROUGH CLAIMS
    Set rs = Me.frm_HP_MR_Response_GridView.Form.RecordSet
    rs.Move miStartRow
    For i = miStartRow To miRowsSelected
        Debug.Print rs("CnlyClaimNum")
        strCnlyClaimNum = Trim(rs("CnlyClaimNum") & "")
        strClmStatus = Trim(rs("ClmStatus") & "")
        If strCnlyClaimNum <> "" Then
            If InStr(1, "302", strClmStatus) = 0 Then
                strDisplayMsg = "Can not put claim [" & strCnlyClaimNum & "] on hold.  Current status = " & strClmStatus
            Else
                strDisplayMsg = "Claim  [" & strCnlyClaimNum & "] put on hold."
            End If
        Else
            strDisplayMsg = "Error: Claim number is blank"
        End If
                
        frm_HP_Claim_Display.lstStatDisplay.AddItem strDisplayMsg
        DoEvents
        DoEvents
        DoEvents
        
        
        rs.MoveNext
    Next i
    
    Me.RowsSelected = 1
    
    Me.frm_HP_MR_Response_GridView.Form.Requery
    
    Set rs = Nothing
    
    Set f = Nothing
End Sub

Private Sub clearICN_Click()
    Me.ICNFilter = ""
End Sub

Private Sub clearRequestDate_Click()
    Me.LetterReqDtFilter = ""
End Sub

Private Sub clrCnlyClaimNum_Click()
    Me.CnlyClaimNumFilter = ""
End Sub

Private Sub clrImageName_Click()
    Me.ImageNameFilter = ""
End Sub

Private Sub cmdClaimStatus_Click()

DoCmd.OpenForm "frm_CUST_SubStatus"

Forms!frm_CUST_SubStatus.Controls("optSelect").Value = 2
Call Forms("frm_CUST_SubStatus").ControlVisible

Forms!frm_CUST_SubStatus.Controls("optSelect").Locked = True
Forms!frm_CUST_SubStatus.Controls("subFrmSubStatus").Locked = True

End Sub

Private Sub cmdMRRequest_Click()

DoCmd.OpenForm "frm_HP_Create_MR_Request", acNormal

End Sub

Private Sub cmdQuickLog_Click()

DoCmd.OpenForm "frm_QuickLog_Main", acNormal

End Sub

Private Sub cmdTransID_Click()

DoCmd.OpenForm "frm_esMD_Main", acNormal

End Sub

Private Sub cmdVoid_Click()
    Dim iWindowHandle As Long
    Dim f As clsWindowHandles
    Dim bFound As Boolean
    
    bFound = False
    
    For Each f In ColWindows
        If f.WindowName = "HP_MR_VOID" Then
            iWindowHandle = f.WindowHandle
            SetForegroundWindow iWindowHandle
            If screen.ActiveForm.hwnd = iWindowHandle Then
                bFound = True
                Exit For
            End If
        End If
    Next
    
    If Not bFound Then
        Set f = New clsWindowHandles
        Set frm_Void_Image = New Form_frm_AUDITCLM_References_Update
        f.WindowHandle = frm_Void_Image.hwnd
        f.WindowName = "HP_MR_VOID"
        ColWindows.Add f, f.WindowHandle & ""
        ShowFormAndWait frm_Void_Image
    End If
    
    Set f = Nothing
    
    Set frm_Void_Image = Nothing
    
End Sub


Private Sub FileNameFilter_Click()
    Me.FileNameFilter = ""
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim StrAccessRights As String
    
    Me.ICNFilter = ""
    Me.CnlyClaimNumFilter = ""
    Me.ImageNameFilter = ""
    Me.ProviderFilter = ""
    Me.LetterReqDtFilter = ""
    Me.FileNameFilter = ""
    
    
    
    mstrOldFilter = ""
    mstrFilter = ""
    
    


    StrAccessRights = Nz(DLookup("[SupervisorID]", "ADMIN_User", "[UserID] ='" & Identity.UserName & "'"), "")

        If StrAccessRights <> "Data Center" Then
            Me.cmdMRRequest.Enabled = False
        End If
    
    mbLoaded = True
    
        If Not (Me.frm_HP_MR_Response_GridView.Form.RecordSet.EOF Or Me.frm_HP_MR_Response_GridView.Form.RecordSet.BOF) Then
            Me.frm_HP_MR_Response_GridView.Form.RecordSet.MoveFirst
            strSQL = "select * from v_HP_MR_Consolidated_View where MRRID = " & Me.frm_HP_MR_Response_GridView.Form.MRRID
        Else
            strSQL = "select * from v_HP_MR_Consolidated_View where 1=2"
        End If
    
'    If Me.frm_HP_MR_Response_GridView.Form.MRRID & "" <> "" Then
'        strSql = "select * from v_HP_MR_Consolidated_View where MRRID = " & Me.frm_HP_MR_Response_GridView.Form.MRRID
'    Else
'        strSql = "select * from v_HP_MR_Consolidated_View where 1=2"
'    End If


    Me.frm_HP_MR_Response.Form.RecordSource = strSQL
    
End Sub


Private Sub frm_Void_Image_UpdateReferences(ErrorCode As String, NewImageType As String, NewPageCount As Integer, Comment As String)
    Dim myCode_ADO As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim strUserMsg As String
    Dim iResult As Integer
    
    

    On Error GoTo ErrHandler

    Set myCode_ADO = New clsADO
    myCode_ADO.SQLTextType = StoredProc
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_SCANNING_Image_Error_Log_Insert"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_SCANNING_Image_Error_Log_Insert"
    cmd.Parameters.Refresh
    
    cmd.Parameters("@pScannedDt") = ConvertTimeToString(Me.frm_HP_MR_Response.Form.ScannedDt)
    cmd.Parameters("@pCnlyClaimNum") = Me.frm_HP_MR_Response.Form.CnlyClaimNum
    cmd.Parameters("@pNewCnlyClaimNum") = ""
    cmd.Parameters("@pNewImageType") = UCase(NewImageType)
    cmd.Parameters("@pNewPageCnt") = NewPageCount
    cmd.Parameters("@pErrorCd") = ErrorCode
    cmd.Parameters("@pComment") = Comment
    
    cmd.Execute
    
    'Make sure there are no errors
    strErrMsg = cmd.Parameters("@pErrMsg") & ""
    If strErrMsg <> "" Then
        strErrMsg = "Error updating Image - " & strErrMsg
        GoTo ErrHandler
    End If
    
    'check to see if we need to display a user message
    strUserMsg = cmd.Parameters("@pUserMsg") & ""
    If strUserMsg <> "" Then MsgBox strUserMsg, vbInformation
    
    ' refresh grid box view
    Me.frm_HP_MR_Response_GridView.Requery
    
Exit_Function:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub

ErrHandler:
    If strErrMsg <> "" Then
        MsgBox strErrMsg
    Else
        MsgBox Err.Description
    End If
    Resume Exit_Function
End Sub

Private Sub tgCnlyClaimNum_Click()
    SetSQL
End Sub

Private Sub Provider_Click()
    Me.ProviderFilter = ""
End Sub

Private Sub tgFileName_Click()
    SetSQL
End Sub

Private Sub tgICN_Click()
    SetSQL
End Sub

Private Sub tgImageName_Click()
    SetSQL
End Sub

Private Sub tgLetterReqDt_Click()
    SetSQL
End Sub

Private Sub tgProvider_Click()
    SetSQL
End Sub

Private Sub SetSQL()
    Dim strSQL
    Dim tmpFileName As String
    
    
    mstrFilter = ""
    
    If Trim(Me.ICNFilter & "") <> "" Then
        mstrFilter = mstrFilter & " and ICN like '" & Trim(Me.ICNFilter) & "%'"
    End If
    
    If Trim(Me.CnlyClaimNumFilter & "") <> "" Then
        mstrFilter = mstrFilter & " and CnlyClaimNum = '" & Trim(Me.CnlyClaimNumFilter) & "'"
    End If
    
    If Trim(Me.ProviderFilter & "") <> "" Then
        mstrFilter = mstrFilter & " and ProvNum like '" & Trim(Me.ProviderFilter) & "%'"
    End If
    
    If Trim(Me.ImageNameFilter & "") <> "" Then
        mstrFilter = mstrFilter & " and ImageName like '" & Trim(Me.ImageNameFilter) & "%'"
    End If
    
    If Trim(Me.LetterReqDtFilter & "") <> "" Then
        mstrFilter = mstrFilter & " and LetterReqDt = '" & Trim(Me.LetterReqDtFilter) & "'"
    End If
    
    If Trim(Me.FileNameFilter & "") <> "" Then
        
        Select Case InStr(1, Trim(Me.FileNameFilter & ""), ".") - 1
            Case -1
                tmpFileName = Trim(Me.FileNameFilter & "") & "%"
            Case Else
                tmpFileName = left(Trim(Me.FileNameFilter & ""), InStr(1, Trim(Me.FileNameFilter & ""), ".") - 1) & "%"
        End Select
        'mstrFilter = mstrFilter & " and HPResponseFile like '" & Trim(Me.FileNameFilter) & "%'"
        mstrFilter = mstrFilter & " and HPResponseFile like '" & tmpFileName & "'"
    End If
            
    ' set record source
    If mstrFilter <> "" Then
        mstrFilter = Mid(mstrFilter, 5, Len(mstrFilter))
        strSQL = "select * from v_HP_MR_Consolidated_View where " & mstrFilter
    Else
        strSQL = "select * from v_HP_MR_Consolidated_View order by MRRID"
    End If
    
    
    
    Me.frm_HP_MR_Response_GridView.Form.RecordSource = strSQL
    If Me.frm_HP_MR_Response_GridView.Form.RecordSet.recordCount = 0 Then
        strSQL = "select * from v_HP_MR_Consolidated_View where 1=2"
        Me.frm_HP_MR_Response.Form.RecordSource = strSQL
        Me.RowsSelected = 0
    End If
    
    If mstrOldFilter <> mstrFilter Then
        bSQLChange = True
        mstrOldFilter = mstrFilter
        'SetDropDown
    End If

End Sub

Private Sub SetDropDown()
    Dim strSQL As String
    
    If mstrFilter <> "" Then
        strSQL = " from v_HP_MR_Consolidated_View where " & mstrFilter
    Else
        strSQL = " from v_HP_MR_Consolidated_View"
    End If
    
'    Me.ICNFilter.RowSource = "select distinct ICN" & strSql & " order by 1"
'    Me.CnlyClaimNumFilter.RowSource = "select distinct CnlyClaimNum" & strSql & " order by 1"
'    Me.ImageNameFilter.RowSource = "select distinct ImageName" & strSql & " order by 1"
'    Me.ProviderFilter.RowSource = "select distinct ProvNum" & strSql & " order by 1"
'    Me.LetterReqDtFilter.RowSource = "select distinct LetterReqDt" & strSql & " order by 1 desc"
    'Me.FileNameFilter.RowSource = "select distinct Filename" & strSQL & " order by 1"
    'Me.FileNameFilter.RowSource = "select distinct HPResponseFile" & strSQL & " order by 1"
    
End Sub

Public Sub DisplayClaimScreen(CnlyClaimNum As String)
    Dim frm_AUDITCLM_Main As Form_frm_AUDITCLM_Main
    Dim iWindowHandle As Long
    Dim f As clsWindowHandles
    Dim bFound As Boolean
    
   
    bFound = False
    
    If CnlyClaimNum & "" <> "" Then
        For Each f In ColWindows
            'Debug.Print f.WindowName
            'Debug.Print f.WindowHandle
            iWindowHandle = f.WindowHandle
            SetForegroundWindow iWindowHandle
            If f.WindowName = "ClaimMain" & CnlyClaimNum Then
                SetForegroundWindow iWindowHandle
                
                If screen.ActiveForm.hwnd = iWindowHandle Then
                    bFound = True
                    Exit For
                End If
            End If
        Next
    
        If Not bFound Then
            Set f = New clsWindowHandles
            Set frm_AUDITCLM_Main = New Form_frm_AUDITCLM_Main
            f.WindowHandle = frm_AUDITCLM_Main.hwnd
            f.WindowName = "ClaimMain" & CnlyClaimNum
            ColWindows.Add f, f.WindowHandle & ""
            ColObjectInstances.Add Item:=frm_AUDITCLM_Main, Key:=frm_AUDITCLM_Main.hwnd & " "
            
            frm_AUDITCLM_Main.Caption = "CMS: ClaimNum : " & CnlyClaimNum
    
            frm_AUDITCLM_Main.visible = True
            frm_AUDITCLM_Main.CnlyClaimNum = CnlyClaimNum
            frm_AUDITCLM_Main.LoadData
        End If
    End If
    
    Set f = Nothing

End Sub
