Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Const DemoGraphicsFrame = 8
Const WebURL = "http://webstrat.ccaintranet.com/HSS/WebStrat/webstrat.aspx"
    'tk 2014-03-20 test site: WebURL = "http://netprvwebst-004/HSS/WebStrat/WebStrat.aspx"

'=============================================
' ID:          Form_frm_AUDITCLM_ReviewChart
'
' Description:
'   Maintain the claim diagnosis codes, procedure codes and rationale
'
' Modification History:
'   2013-07-23 by KD: Modified this and the parent form to see if the claim is a 'Therapy (Congress)'
'       claim, if so get the error codes that the auditor needs to choose from and set parent form
'       properties for saving.. Also transferred save procedure from the _Click event to the main form
'       Svae Button
'   2010-09-13 by BJD to store the updated rational text when updated from the spell checker.
'   2011-04-12 by Gautam Malhotra: Added Webstrat Button
'   2011-05-13 by Alex C: Added Present on Admission (POA) to Diag codes
' =============================================

Private WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private WithEvents frmRationaleTemplate As Form_frm_AUDITCLM_RATIONALE_TEMPLATE
Attribute frmRationaleTemplate.VB_VarHelpID = -1

Private strCnlyClaimNum As String
Private strDRG As String
Private strAppID As String
Private rsAuditClmHdr As ADODB.RecordSet
Private rsAuditClmDiag As ADODB.RecordSet
Private rsAuditClmProc As ADODB.RecordSet
Private rsAuditClmProcRev As ADODB.RecordSet
Private rsAuditClmDiagRev As ADODB.RecordSet
Private rsAuditClmHdrAdditionalInfo As ADODB.RecordSet
Private Const cHCPCSFrame = 4
'Private rsAuditClmHdr As ADODB.Recordset 'andrew
Const CstrFrmAppID As String = "ChartReview"
Public mbFinished As Boolean
Private strUserName, strPassWord As String
Private mbWebStratQuit, mbWebStratLoggedIn As Boolean

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property
Property Set HdrRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmHdr = data
End Property

Property Set DiagCodeRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmDiag = data
End Property

Property Get DiagCodeRecordsource() As ADODB.RecordSet
     Set DiagCodeRecordsource = rsAuditClmDiag
End Property

Property Set ProcCodeRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmProc = data
End Property

Property Get ProcCodeRecordsource() As ADODB.RecordSet
     Set ProcCodeRecordsource = rsAuditClmProc
End Property
Property Set DiagCodeRevRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmDiagRev = data
End Property

Property Get DiagCodeRevRecordsource() As ADODB.RecordSet
     Set DiagCodeRevRecordsource = rsAuditClmDiagRev
End Property

Property Set ProcCodeRevRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmProcRev = data
End Property


Property Get ProcCodeRevRecordsource() As ADODB.RecordSet
     Set ProcCodeRevRecordsource = rsAuditClmProcRev
End Property

Property Get HdrAddInfoRecordsource() As ADODB.RecordSet
     Set HdrAddInfoRecordsource = rsAuditClmHdrAdditionalInfo
End Property


Property Set HdrAddInfoRecordsource(data As ADODB.RecordSet)
    Set rsAuditClmHdrAdditionalInfo = data
End Property

Property Get HdrRecordsource() As ADODB.RecordSet
     Set HdrRecordsource = rsAuditClmHdr
End Property

'Main property of form.  This drives everything that this object is based on
Property Let CnlyClaimNum(data As String)
    strCnlyClaimNum = data
End Property

Property Get CnlyClaimNum() As String
    CnlyClaimNum = strCnlyClaimNum
End Property

Property Let DRG(data As String) 'andrew
    strDRG = data
End Property

Property Get DRG() As String 'andrew
    DRG = strDRG
End Property

Private Sub Adj_Rationale_AfterUpdate()
    If Not rsAuditClmHdr.EOF Then
        rsAuditClmHdr.Fields("Adj_Rationale") = Me.Adj_Rationale
        FormIsDirty
    End If
End Sub
Function GetOpenIEByTitle(i_Title As String, _
                          Optional ByVal i_ExactMatch As Boolean = True) As SHDocVw.InternetExplorer
Dim objShellWindows As New SHDocVw.ShellWindows

  If i_ExactMatch = False Then i_Title = "*" & i_Title & "*"
  'ignore errors when accessing the document property
  On Error Resume Next
  'loop over all Shell-Windows
  For Each GetOpenIEByTitle In objShellWindows
    'if the document is of type HTMLDocument, it is an IE window
    If TypeName(GetOpenIEByTitle.Document) = "HTMLDocument" Then
      'check the title
      If GetOpenIEByTitle.Document.Title Like i_Title Then
        'leave, we found the right window
        Exit Function
      End If
    End If
  Next
End Function


Private Sub cmdRationale_Click()
    On Error GoTo ErrHandler
    
    'Dim frmPopup As Form
    'Public ColTemplateInstances As New Collection
    
        'ColObjectInstances.Add Item:=frmRationaleTemplate, KeyKey
        
        
     'Andrew hold off on these below
'    Dim myDRG As String
'    myDRG = DLookup("DRGGROUP", "XREF_DRGGROUP", "DRG =" & rsAuditClmHdr.Fields("DRG"))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim myADO As clsADO
'    Dim rs As ADODB.Recordset
'     Set myADO = New clsADO

'     myADO.ConnectionString = GetConnectString("XREF_DRGGROUP")
'     myADO.SQLTextType = sqltext
     'myADO.SQLstring = " SELECT DRGGROUP FROM XREF_DRGGROUP "
     'myADO.SQLstring = myADO.SQLstring & "WHERE DRG =" & Chr(34) & Nz(DRG, "999") & Chr(34)
     If frmRationaleTemplate Is Nothing Then

        Set frmRationaleTemplate = New Form_frm_AUDITCLM_RATIONALE_TEMPLATE
 Dim MyRatID As String
 MyRatID = Nz(DLookup("RationaleID", "XREF_DRGGROUP", "DRG = '" & DRG & "'"), "1")
  frmRationaleTemplate.TemplateID = MyRatID
    
'
'    If myDRG = "S" Then frmRationaleTemplate.TemplateID = 2
'
'    If myDRG <> "S" Then frmRationaleTemplate.TemplateID = 1
 
 
 
 
'     Set rs = myADO.OpenRecordSet(" SELECT DRGGROUP FROM XREF_DRGGROUP WHERE DRG =" & Chr(34) & Nz(DRG, "999") & Chr(34))
'     Set rs = myADO.ExecuteRS
'     myDRG = rs("DRGGROUP")
'     If rs Is Nothing Then myDRG = "M"
'
'       ' Else
'        '  If rs = 0 Then myDRG = "M"
'        End If
'    '("select TrtmtAuth from CMS_Data_NCH.dbo.INT_HDR where CnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "'")
'
'
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '  frmRationaleTemplate.TemplateID = 2 ' original
       
        frmRationaleTemplate.CnlyClaimNum = Me.CnlyClaimNum
        Set frmRationaleTemplate.HdrRecordsource = Me.HdrRecordsource
        frmRationaleTemplate.RefreshData
        frmRationaleTemplate.visible = True
        'frmRationaleTemplate.visible = True
        'ShowFormAndWait frmRationaleTemplate
        'frmRationaleTemplate.visible = True
        'RemoveObjectInstance frmRationaleTemplate
        'Set frmRationaleTemplate = Nothing
    Else
        frmRationaleTemplate.SetFocus
    End If
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Set frmRationaleTemplate = Nothing
End Sub

Private Sub ErrorCode_AfterUpdate()
On Error GoTo Block_Err
Dim strProcName As String
'Dim oAdo As clsADO
'Dim rst As ADODB.Recordset
Dim strSQL As String
Dim strErrorCodeNew
Dim oForm As Form_frm_AUDITCLM_Main
    
    strProcName = ClassName & ".ErrorCode_AfterUpdate"
    
    strErrorCodeNew = Me.ErrorCode.Value
    
    ' 7/16/2013 KD: Don't want to call this unless it's actually set
    If Trim(strErrorCodeNew) = "" Then
        GoTo Block_Exit
    End If
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strSQL = " EXEC cms_auditors_code.dbo.usp_ErrorCodeUpdate '" & Me.Parent.Form.CnlyClaimNum & "', '" & strErrorCodeNew & "'"
    'oAdo.Execute strSQL
    Debug.Print strSQL
    
    '' 20130723 KD: Ok, let's not save it here.. Let's pass the value to the main form so we can set it there IF (and only if) the
    '' user clicks the save button..
        '    Set oAdo = New clsADO
        '    With oAdo
        '        .ConnectionString = GetConnectString("v_Data_Database")
        '        .SQLTextType = SQLTEXT
        '        .sqlString = strSQL
        '        Set rst = .ExecuteRS
        '    End With
    
        ' 7/16/2013 KD: Also, want to keep the value on the main form so we don't have to figure it out again
    If IsSubForm(Me) Then
        
        Set oForm = Me.Parent.Form
        oForm.ErrorCodePrpty = strErrorCodeNew
        oForm.IsTherapyConcept = True
        Set oForm = Nothing
    End If

    
Block_Exit:
'    If Not rst Is Nothing Then
'        If rst.State = adStateOpen Then rst.Close
'    End If
'
'    Set rst = Nothing
'    Set oAdo = Nothing
    Exit Sub


Block_Err:
    ReportError Err, strProcName, , Me.Parent.Form.CnlyClaimNum
    GoTo Block_Exit
End Sub

'VS 11/17/2015 Give the user ability to clear text box - part of RVC updates
Private Sub ErrorCode_LostFocus()
    If Nz(Me.ErrorCode, "") = "" Then
    Dim oForm As Form_frm_AUDITCLM_Main
    Set oForm = Me.Parent.Form
        oForm.ErrorCodePrpty = Nz(Me.ErrorCode, "")
        'FormIsDirty
    End If
End Sub

Private Sub Form_Current()
    If (Me.Parent.DataType = "CARR") Or _
        (Me.Parent.DataType = "DME") _
        Then
        'Or (Me.lstDiagCodesRevised.ListCount = 0 And Me.lstProcCodesRevised.ListCount = 0) Or (Me.Parent.Adj_DRG <> "") Then
        Me.btnWebStrat.Enabled = False
    Else
        Me.btnWebStrat.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRationaleTemplate = Nothing
End Sub

Private Sub frmRationaleTemplate_FormClosed()
    Set frmRationaleTemplate = Nothing
End Sub

Private Sub frmRationaleTemplate_RationaleConfirmed(strCommitRationale As String, strHwnd As String)
    Dim obj As Object
    Dim blnRemove As Boolean
    Dim intI As Integer
    
    Me.Adj_Rationale = strCommitRationale
    If Not rsAuditClmHdr.EOF Then
        rsAuditClmHdr.Fields("Adj_Rationale") = Me.Adj_Rationale
        FormIsDirty
    End If
    Set frmRationaleTemplate = Nothing
End Sub

Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'Debug.Print "Document complete"
End Sub

Private Sub IE_OnQuit()
    'mbWebStratLoggedIn = False
    mbWebStratQuit = True
End Sub

Private Sub IE_TitleChange(ByVal Text As String)
    If IE.Document.Title <> "Web.Strat" Then
        mbWebStratLoggedIn = False
    Else
        mbWebStratLoggedIn = True
    End If
End Sub
'Private Function IE_oncontextmenu() As Boolean
'   IE_oncontextmenu = False
'   'PopupMenu mnu '<---Check the mnu to your own menu name
'   Debug.Print "Custom Right Click"
'End Function

Private Sub btnWebStrat_Click()
On Error Resume Next 'Needed to skip the tmpusername and tmppassword assign statements

If VarType(IE) <> 9 Then 'if IE wasn't closed JS 09/11/2012
    ' do nothing if IE is busy or already has webstrat open
    ' this is not a definite solution but avoids the issues due to user clicking multiple times on webstrat consecutively
    If IE.Busy Or IE.LocationURL = WebURL Then Exit Sub
End If

Dim DiagList As listBox
Dim ProcList As listBox
Dim tmpUsername, tmpPassword As String 'Used to save credentials temporarily till there's a successful login

    mbFinished = False
    Set IE = GetOpenIEByTitle("Web.Strat", False)

If IE Is Nothing Then
    Set IE = CreateObject("InternetExplorer.Application")
    With IE
        .left = 20
        .top = 20
        .Height = 700
        .Width = 1100
        .MenuBar = 0
        .Toolbar = 1
        .StatusBar = 0
        .visible = True
    End With
End If
    mbWebStratQuit = False
    IE.Navigate WebURL
    IE.Document.Focus
    
    'wait until IE has finished loading itself.
    Do While IE.Busy Or Not IE.ReadyState = 4 'Or IE.Document.Title <> "Web.Strat Login"
        DoEvents
        If (VarType(IE) = 9) Then GoTo BreakOutOfLoop 'if the user closed IE - webstrat while loading the data JS 09/11/2012
    Loop
    
BreakOutOfLoop:
    
    Dim credentialsSaved As Boolean
    
    If IE.Document.URL = WebURL Then
        credentialsSaved = True
        mbWebStratLoggedIn = True
    Else
        If strUserName = "" And strPassWord = "" Then
            credentialsSaved = False
        Else
            IE.Document.Forms.Item(, 0).elements("TextBoxUserId").Value = strUserName
            IE.Document.Forms.Item(, 0).elements("TextBoxPassword").Value = strPassWord
            IE.Document.getElementById("ButtonLogin").Click
            credentialsSaved = True
        End If
    End If
    
    'Let the user enter credentials
    Do While credentialsSaved = False
        If IE Is Nothing Then
            'Clear credentials
            'strUserName = ""
            'strPassWord = ""
            Exit Sub 'Exit process if the user closed the IE window without proceeding
        Else
            'Save login credentials and persist till this form gets submitted
            If IE.Document.Title = "Web.Strat login" Then 'We are still on the login page
                tmpUsername = IE.Document.Forms.Item(, 0).elements("TextBoxUserId").Value
                tmpPassword = IE.Document.Forms.Item(, 0).elements("TextBoxPassword").Value
            End If
            If IE.Document.Title = "Web.Strat" Then
                credentialsSaved = True
                mbWebStratLoggedIn = True
                Exit Do
            End If
        End If
        'DoEvents
    Loop
   
    'wait until IE has finished loading itself.
    Do While (IE.Busy Or Not IE.ReadyState = 4) And mbWebStratQuit = False
        If IE Is Nothing Or mbWebStratQuit Or (VarType(IE) = 9) Then Exit Sub 'vartype(IE) = 9 means IE was closed JS 09/11/2012
        'Exit Sub
        'DoEvents
    Loop
    
    'Wait for sometime
    'Wait (10)
    
    'wait until IE has finished loading itself.
    Do While IE.Busy Or Not IE.ReadyState = 4
        DoEvents
    Loop
    '***********LOGGED IN*************
    

    
    'Wait for sometime
    Wait (10)
    
    'Since we are logged in now, save the tmp credentials permanently
    If strUserName = "" And strPassWord = "" Then
        strUserName = tmpUsername
        strPassWord = tmpPassword
    End If
    
     'VS 2/8/2016 ICD Version 10 Fix
    If Me.Parent.IcdversionCDflag = 0 Then
        IE.Document.frames(DemoGraphicsFrame).Document.getElementById("dropdownlistCodeClass").Value = "01"
        IE.Document.frames(DemoGraphicsFrame).Document.getElementById("dropdownlistCodeClass").fireevent ("onchange")
    End If
    
    '******Fill up demographics here except payer ID(comes from header)*******
    If Me.Parent.DataType = "IP" Then IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownPatType").Value = "01"
    If Me.Parent.DataType = "OP" Then IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownPatType").Value = "02"
    If Me.Parent.DataType = "HH" Then IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownPatType").Value = "07"
    If Me.Parent.DataType = "SNF" Then IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownPatType").Value = "06"
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownPatType").fireevent ("onchange")
    
    'Take care of the popup here - NOT A PERMANENT SOLUTION - Why does the IE object loose relevance due to a popup?!!
    While Trim(IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").Value) = ""
        If Trim(Me.Parent.Adj_DischargeStatus) <> "" Then
            Debug.Print IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").Value + " BEFORE"
            IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").Value = Me.Parent.Adj_DischargeStatus
            Debug.Print IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").Value + " AFTER"
        Else
            IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").Value = Me.Parent.DischargeStatus
        End If
        If VarType(IE) = 9 Then Exit Sub 'Get out if user closes IE while processing webstrat page, it was getting stuck here JS 09/11/2012
    Wend
    
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("DropDownDStat").fireevent ("onchange")
    
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxBirth").Value = Me.Parent.BeneBirthDt
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxBirth").fireevent ("onchange")
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxSex").Value = Me.Parent.BeneSexCd
    IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxSex").fireevent ("onchange")
    
    '****FILL up header values here
    IE.Document.getElementById("TextBoxAdmitDate").Value = Me.Parent.ClmFromDt
    IE.Document.getElementById("TextBoxAdmitDate").fireevent ("onchange")
    IE.Document.getElementById("TextBoxDischDate").Value = Me.Parent.ClmThruDt
    IE.Document.getElementById("TextBoxDischDate").fireevent ("onchange")
    IE.Document.getElementById("TextBoxFacilityID").Value = Me.Parent.ProvNum
    IE.Document.getElementById("TextBoxFacilityID").fireevent ("onchange") 'Loads facility id in the demographics frame
    IE.Document.getElementById("TextBoxPtIDMedRec").Value = "01"
    IE.Document.getElementById("txtSumRVDX1").Value = Me.Parent.AdmitDiag
    '****Header finished
    
    If Me.lstDiagCodesRevised.ListCount = 0 Then
        Set DiagList = Me.lstDiagCodes
    Else
        Set DiagList = Me.lstDiagCodesRevised
    End If
    For i = 0 To DiagList.ListCount - 1
        'Debug.Print DiagList.Column(1, i)
        IE.Document.getElementById("dx" + Trim(str(i))).Value = DiagList.Column(1, i)
        IE.Document.getElementById("dx" + Trim(str(i))).fireevent ("onchange")
        If IsNull(DiagList.Column(2, i)) Or DiagList.Column(2, i) = "" Then
            IE.Document.getElementById("Onset" + Trim(str(i))).Value = 1 'Old claims not having a poa cd
        Else
            IE.Document.getElementById("Onset" + Trim(str(i))).Value = DiagList.Column(2, i)
        End If
        IE.Document.getElementById("Onset" + Trim(str(i))).fireevent ("onchange")
        If IE.Document.frames("inPatIFrameL").Document.getElementById("ELeft" + Trim(str(i))).Value = "E" Then
            'Debug.Print IE.Document.frames("inPatIFrameL").Document.getElementById("ELeft0").value
            IE.Document.getElementById("Onset" + Trim(str(i))).Value = "Y"
            IE.Document.getElementById("Onset" + Trim(str(i))).fireevent ("onchange")
        End If
    Next i
    
    
    If Me.lstProcCodesRevised.ListCount = 0 Then
        'Debug.Print "Revised Proc code list is empty"
        Set ProcList = Me.lstProcCodes
    Else
        Set ProcList = Me.lstProcCodesRevised
    End If
    For i = 0 To ProcList.ListCount - 1
        'Debug.Print ProcList.Column(1, i)
        If left(ProcList.Column(1, i), 3) <> "000" Then
            IE.Document.getElementById("px" + Trim(str(i))).Value = ProcList.Column(1, i)
            IE.Document.getElementById("px" + Trim(str(i))).fireevent ("onchange")
            IE.Document.frames(2).Document.getElementById("PxDate" + Trim(str(i))).Value = ProcList.Column(2, i)
            IE.Document.frames(2).Document.getElementById("PxDate" + Trim(str(i))).fireevent ("onchange")
        End If
    Next i

    If Me.Parent.DataType = "HH" Then
    'Home Health specific processes
        IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxBillType").Value = "0329"
        IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxBillType").fireevent ("onchange")
        
        'Fetch treatment authorisation code from INT_Hdr
         Dim MyAdo As clsADO
         Dim rs As ADODB.RecordSet
         
         Set MyAdo = New clsADO
         MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
         Set rs = MyAdo.OpenRecordSet("select TrtmtAuth from CMS_Data_NCH.dbo.INT_HDR where CnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "'")
        
        If rs.recordCount > 0 Then
            IE.Document.frames(DemoGraphicsFrame).Document.getElementById("TextBoxTAC").Value = rs.Fields(0).Value
        Else
            MsgBox "No Claim information in the raw data. Cannot fetch Treatment Authorisation Code", vbInformation
        End If
         
         'Get Value Codes
         IE.Document.frames(DemoGraphicsFrame).Document.getElementById("checkValueCodes").Checked = True
         IE.Document.frames(DemoGraphicsFrame).Document.getElementById("checkValueCodes").fireevent ("onclick")
         
        Set rs = MyAdo.OpenRecordSet("select ValueCds,ValueAmts from CMS_Data_NCH.dbo.INT_Value where CnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "'")

        If rs.recordCount > 0 Then
            Dim arrProfile() As String
            Dim intI As Integer
            arrProfile = Split(rs.Fields(0).Value, "|")
            For intI = 0 To UBound(arrProfile())
                IE.Document.frames(DemoGraphicsFrame).Document.getElementById("VC_" & intI).Value = Trim(arrProfile(intI))
                IE.Document.frames(DemoGraphicsFrame).Document.getElementById("VC_" & intI).fireevent ("onchange")
            Next intI
            
            arrProfile = Split(rs.Fields(1).Value, "|")
            For intI = 0 To UBound(arrProfile())
                IE.Document.frames(DemoGraphicsFrame).Document.getElementById("VC_amount_" & intI).Value = Trim(arrProfile(intI))
                IE.Document.frames(DemoGraphicsFrame).Document.getElementById("VC_amount_" & intI).fireevent ("onchange")
            Next intI
        Else
            MsgBox "No value codes in the raw data.", vbInformation
        End If
         
         
    End If

   If Me.Parent.DataType <> "IP" Then
   ' This loop will fill in the line level details
   Debug.Print Me.Parent.lstTabs.Value
   Dim thisClaim As clsAUDITCLM
   Set thisClaim = New clsAUDITCLM
   If Not (thisClaim.LoadClaim(Me.CnlyClaimNum, False)) Then
        MsgBox "Error loading claim details!", vbCritical
        Exit Sub
   End If
   Dim rstInputDetail As ADODB.RecordSet
   Set rstInputDetail = thisClaim.rsAuditClmDtl
            
            Dim iRecord As Integer
            Dim aModifiers As Variant
            
            iRecord = 0
            rstInputDetail.MoveFirst
            
            'If rstInputDetail.RecordCount > 12 Then
               'MsgBox "Please expand the list of HCPCS codes to allow for " + CStr(rstInputDetail.recordCount) + " HCPCS codes inside WebStrat and then click Okay", vbOKOnly, "Expand HCPCS fields before continuing"
            'End If
            
            While Not rstInputDetail.EOF
            With IE.Document
                If rstInputDetail!RevCd <> "0001" Then
                    .frames(cHCPCSFrame).Document.getElementById("rev" + CStr(iRecord)).Value = rstInputDetail!RevCd & ""
                    .frames(cHCPCSFrame).Document.getElementById("rev" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("hcpcs" + CStr(iRecord)).Value = rstInputDetail!HCPCSCd & ""
                    .frames(cHCPCSFrame).Document.getElementById("hcpcs" + CStr(iRecord)).fireevent ("onload")
                    
                    If rstInputDetail!Adj_Ind = "Y" Then
                        .frames(cHCPCSFrame).Document.getElementById("Units" + CStr(iRecord)).Value = rstInputDetail!Adj_Units & ""
                    Else
                        .frames(cHCPCSFrame).Document.getElementById("Units" + CStr(iRecord)).Value = rstInputDetail!Units & ""
                    End If
                    .frames(cHCPCSFrame).Document.getElementById("Units" + CStr(iRecord)).fireevent ("onload")
                    
                    .frames(cHCPCSFrame).Document.getElementById("charges" + CStr(iRecord)).Value = rstInputDetail!LnChargeAmt & ""
                    .frames(cHCPCSFrame).Document.getElementById("charges" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("date" + CStr(iRecord)).Value = rstInputDetail!LnClmFromDt & ""
                    .frames(cHCPCSFrame).Document.getElementById("date" + CStr(iRecord)).fireevent ("onload")
                    
                    .frames(cHCPCSFrame).Document.getElementById("M1" + CStr(iRecord)).Value = rstInputDetail!Mod01 & ""
                    .frames(cHCPCSFrame).Document.getElementById("M1" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("M2" + CStr(iRecord)).Value = rstInputDetail!Mod02 & ""
                    .frames(cHCPCSFrame).Document.getElementById("M2" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("M3" + CStr(iRecord)).Value = rstInputDetail!Mod03 & ""
                    .frames(cHCPCSFrame).Document.getElementById("M3" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("M4" + CStr(iRecord)).Value = rstInputDetail!Mod04 & ""
                    .frames(cHCPCSFrame).Document.getElementById("M4" + CStr(iRecord)).fireevent ("onload")
                    .frames(cHCPCSFrame).Document.getElementById("M5" + CStr(iRecord)).Value = rstInputDetail!Mod05 & ""
                    .frames(cHCPCSFrame).Document.getElementById("M5" + CStr(iRecord)).fireevent ("onload")
                End If
                If iRecord > 10 Then
                    .frames(cHCPCSFrame).Document.parentWindow.execScript "AddRow(false,'APC');parent.AddQuickRow('hcpcs');", "JavaScript"
                End If
                'IE.Document.Window.execScript "AddRow(false,'APC');parent.AddQuickRow('hcpcs');", "JavaScript"
                iRecord = iRecord + 1
                rstInputDetail.MoveNext
            End With
            Wend
   End If
    
'    iCount = IE.Document.frames(DemoGraphicsFrame).Document.getElementsByTagName("INPUT").length
'    For i = 1 To iCount
'        Debug.Print IE.Document.frames(DemoGraphicsFrame).Document.getElementsByTagName("INPUT")(i).Name
'    Next i
  
    'Choose payer ID now
    '2013:05:02:Gautam: Changed payerID to 09 for medicare. R3 to replicate these changes
    'IE.Document.getElementById("SumDropDownPayerID").Value = IE.Document.getElementById("SumDropDownPayerID").Item(1).Value
    'IE.Document.getElementById("SumDropDownPayerID").fireevent ("onchange")
    IE.Document.getElementById("txtsumDropDownPayerID").Value = "09"
    IE.Document.getElementById("txtSumDropDownPayerID").fireevent ("onchange")
    IE.Document.getElementById("txtSumDropDownPayerID").fireevent ("onblur")

    'Ctrl+G Now
    IE.Document.parentWindow.execScript "MenuGroupAndPrice()", "JScript"
    
    'wait until IE has finished Recalculating everything
    Do While IE.Busy Or Not IE.ReadyState = 4
        DoEvents
    Loop
    
    If IE.Document.frames("InPatIFrameFoot").Document.getElementById("LabelDRGCode").innerText <> "" And IE.Document.frames("InPatIFrameFoot").Document.getElementById("LabelTotalReimbursement").innerText <> "" Then
    '*****Disabled at Auditors' Request
        If ((Me.lstDiagCodesRevised.ListCount > 0 Or Me.lstProcCodesRevised.ListCount > 0 Or Trim(Me.Parent.Adj_DischargeStatus) <> "") And (Me.Parent.Adj_DRG = "" Or IsNull(Me.Parent.Adj_DRG))) Then
            Me.Parent.Adj_ReimbAmt = IE.Document.frames("InPatIFrameFoot").Document.getElementById("LabelTotalReimbursement").innerText
            Me.Parent.Adj_DRG = Right(IE.Document.frames("InPatIFrameFoot").Document.getElementById("LabelDRGCode").innerText, 3)
            Me.Parent.Adj_ProjectedSavings = Me.Parent.ReimbAmt - Me.Parent.Adj_ReimbAmt
            
            If MsgBox("Update Adjusted DRG, Reimb Amt & Projected Savings as shown?", vbYesNo) = vbYes Then
                FormIsDirty
            Else
                Me.Parent.Undo
            End If
        End If
    End If
ErrHandlr:
End Sub


Private Sub cmdAddProc_Click()
    AddProcCode
End Sub

Private Sub SaveData()
    Dim lngLineNum As Long
    Dim strDiagCode As String
    Dim varItem As Variant
    Dim intI As Integer
    Dim bResult As Boolean
    
    bResult = True
    
    If Not rsAuditClmDiagRev Is Nothing Then
        If rsAuditClmDiagRev.recordCount > 0 Then
            'Now wipe the local recordset
            rsAuditClmDiagRev.MoveFirst
            While Not rsAuditClmDiagRev.EOF
                rsAuditClmDiagRev.Delete
                rsAuditClmDiagRev.MoveNext
                rsAuditClmDiagRev.UpdateBatch
            Wend
            bResult = True
        Else
            bResult = True
        End If
            
        'If the above was successful, add the list contents to the recordset
        If bResult Then
            For intI = 0 To Me.lstDiagCodesRevised.ListCount - 1
                rsAuditClmDiagRev.AddNew
                rsAuditClmDiagRev("CnlyClaimNum") = Me.CnlyClaimNum
                rsAuditClmDiagRev("LineNum") = Me.lstDiagCodesRevised.Column(0, intI)
                rsAuditClmDiagRev("DiagCd") = Me.lstDiagCodesRevised.Column(1, intI)
                rsAuditClmDiagRev("PoaCd") = Me.lstDiagCodesRevised.Column(2, intI)
                rsAuditClmDiagRev.MoveNext
                rsAuditClmDiagRev.UpdateBatch
            Next intI
        End If
    End If
        
    If Not rsAuditClmProcRev Is Nothing Then
        If rsAuditClmProcRev.recordCount > 0 Then
            'Now wipe the local recordset
            rsAuditClmProcRev.MoveFirst
            While Not rsAuditClmProcRev.EOF
                rsAuditClmProcRev.Delete
                rsAuditClmProcRev.MoveNext
                rsAuditClmProcRev.UpdateBatch
            Wend
            bResult = True
        Else
            bResult = True
        End If
    End If
        
    'If the above was successful, add the list contents to the recordset
    If bResult Then
        For intI = 0 To Me.lstProcCodesRevised.ListCount - 1
            rsAuditClmProcRev.AddNew
            rsAuditClmProcRev("CnlyClaimNum") = Me.CnlyClaimNum
            rsAuditClmProcRev("LineNum") = Me.lstProcCodesRevised.Column(0, intI)
            rsAuditClmProcRev("ProcCd") = Me.lstProcCodesRevised.Column(1, intI)
            
            rsAuditClmProcRev("ProcDt") = IIf(Nz(Me.lstProcCodesRevised.Column(2, intI), "") = "", "1/1/1900", Nz(Me.lstProcCodesRevised.Column(2, intI), ""))
            rsAuditClmProcRev.MoveNext
            rsAuditClmProcRev.UpdateBatch
        Next intI
    End If

    'MsgBox (rsAuditClmHdr.recordCount)
    If Not rsAuditClmHdr Is Nothing Then
        If Not rsAuditClmHdr.EOF Then
         '   MsgBox (rsAuditClmHdr.recordCount)
            rsAuditClmHdr.Fields("Adj_Rationale") = Me.Adj_Rationale
        End If
    End If
    
    'Move first
    If Not rsAuditClmDiagRev Is Nothing Then
        If rsAuditClmDiagRev.recordCount > 0 Then
            rsAuditClmDiagRev.MoveFirst
        End If
    End If

    If Not rsAuditClmProcRev Is Nothing Then
        If rsAuditClmProcRev.recordCount > 0 Then
            rsAuditClmProcRev.MoveFirst
        End If
    End If
End Sub

Private Sub cmdManualDiagCode_Click()
    
    'JS 20121128 Added validation to the txtDiag field
    'GS 20160707 Changed valid range from 3-5 to 3-7
    If Len(Trim(Me.txtDiag)) < 3 Or Len(Trim(Me.txtDiag)) > 7 Then
        MsgBox "Revised Diagnosis code must be between 3 to 7 characters long.", vbInformation, "Revised Diagnosis Code Entry"
        Me.txtDiag.SetFocus
        Exit Sub
    End If
    

    If Nz(Me.txtDiag, "") <> "" And Nz(Me.cmboPOACd, "") <> "" Then
        Me.lstDiagCodesRevised.AddItem CStr(Me.lstDiagCodesRevised.ListCount + 1) & ";" & Me.txtDiag & ";" & Me.cmboPOACd.Value
        FormIsDirty
        Me.txtDiag = ""
        Me.cmboPOACd = ""
    End If
    
    SaveData
    
End Sub

Private Sub cmdManualProcCode_Click()

    'JS 20121128 Added validation to the txtProc and txtProcDt field
    If Len(Trim(Me.txtProc)) < 3 Or Len(Trim(Me.txtProc)) > 7 Then
        MsgBox "Revised procedure codes must be between 3 to 7 characters long.", vbInformation, "Revised Procedure Code Entry"
        Me.txtProc.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(Me.txtProcDt) Then
        MsgBox "Revised procedure code date is invalid.", vbInformation, "Revised Procedure Code Date Entry"
        Me.txtProcDt.SetFocus
        Exit Sub
    End If
    
    If Not (CDate(Me.txtProcDt) >= CDate(Me.Parent.Form.IPAdmitDate) And CDate(Me.txtProcDt) <= CDate(Me.Parent.Form.IPDischargeDt)) Then
        MsgBox "Revised procedure code date must be between admission and discharge.", vbInformation, "Revised Procedure Code Date Entry"
        Me.txtProcDt.SetFocus
        Exit Sub
    End If
    
    If Nz(Me.txtProc, "") <> "" And Nz(Me.txtProcDt, "") <> "" Then
        If IsDate(Me.txtProcDt) = False Then
            MsgBox "Revised procedure codes must include a valid date."
        Else
            Me.lstProcCodesRevised.AddItem CStr(Me.lstProcCodesRevised.ListCount + 1) & ";" & Me.txtProc & ";" & Me.txtProcDt
            FormIsDirty
            Me.txtProc = ""
            Me.txtProcDt = ""
        End If
    End If

    SaveData
    
End Sub

Private Sub cmdMoveDownProc_Click()
    Dim temp
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim colCnt As Integer
    Dim tempval As String
    Dim Result As String
    
    If Me.lstProcCodesRevised.ItemsSelected.Count = 0 Then
        Exit Sub
    End If
    
    colCnt = lstProcCodesRevised.ColumnCount
    p = lstProcCodesRevised.ListIndex + 1
    
    If p < lstProcCodesRevised.ListCount Then
        temp = Split(lstProcCodesRevised.RowSource, ";")
        
        ' swap rows down
        For j = 1 To colCnt
            tempval = temp(p * colCnt - j)
            temp(p * colCnt - j) = temp((p + 1) * colCnt - j)
            temp((p + 1) * colCnt - j) = tempval
        Next j
        
        For i = 0 To UBound(temp) Step colCnt
            Result = Result & Trim(CStr((i \ colCnt) + 1)) & ";"
            For j = 1 To colCnt - 1
                Result = Result & temp(i + j) & ";"
            Next j
        Next i
        
        Result = left(Result, Len(Result) - 1)
        lstProcCodesRevised.RowSource = Result
        lstProcCodesRevised.Selected(p) = True
        FormIsDirty
    End If
    
    SaveData
    
End Sub

Private Sub cmdMoveUpDiag_Click()
    Dim temp
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim colCnt As Integer
    Dim tempval As String
    Dim Result As String
    
    If Me.lstDiagCodesRevised.ItemsSelected.Count = 0 Then
        Exit Sub
    End If
        
    colCnt = lstDiagCodesRevised.ColumnCount
    p = lstDiagCodesRevised.ListIndex + 1
    
    If p > 1 Then
        temp = Split(lstDiagCodesRevised.RowSource, ";")
        
        For j = 1 To colCnt
            tempval = temp(p * colCnt - j)
            temp(p * colCnt - j) = temp((p - 1) * colCnt - j)
            temp((p - 1) * colCnt - j) = tempval
        Next j
        
        For i = 0 To UBound(temp) Step colCnt
            Result = Result & Trim(CStr((i \ colCnt) + 1)) & ";"
            For j = 1 To colCnt - 1
                Result = Result & temp(i + j) & ";"
            Next j
        Next i
        
        Result = left(Result, Len(Result) - 1)
        lstDiagCodesRevised.RowSource = Result
        lstDiagCodesRevised.Selected(p - 2) = True
        FormIsDirty
    End If

    SaveData

End Sub

Private Sub cmdMoveUpProc_Click()
    Dim temp
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim colCnt As Integer
    Dim tempval As String
    Dim Result As String
    
    If Me.lstProcCodesRevised.ItemsSelected.Count = 0 Then
        Exit Sub
    End If

    colCnt = lstProcCodesRevised.ColumnCount
    p = lstProcCodesRevised.ListIndex + 1
    
    If p > 1 Then
        temp = Split(lstProcCodesRevised.RowSource, ";")
        
        For j = 1 To colCnt
            tempval = temp(p * colCnt - j)
            temp(p * colCnt - j) = temp((p - 1) * colCnt - j)
            temp((p - 1) * colCnt - j) = tempval
        Next j
        
        For i = 0 To UBound(temp) Step colCnt
            Result = Result & Trim(CStr((i \ colCnt) + 1)) & ";"
            For j = 1 To colCnt - 1
                Result = Result & temp(i + j) & ";"
            Next j
        Next i
        
        Result = left(Result, Len(Result) - 1)
        lstProcCodesRevised.RowSource = Result
        lstProcCodesRevised.Selected(p - 2) = True
        FormIsDirty
    End If

    SaveData

End Sub

Private Sub cmdSpellCheck_Click()

On Error GoTo ErrHandler

    Me.Adj_Rationale.SetFocus
    DoCmd.RunCommand acCmdSpelling
    rsAuditClmHdr.Fields("Adj_Rationale") = Me.Adj_Rationale
    FormIsDirty
Exit Sub

ErrHandler:
    MsgBox "Spelling Check Failed - " & Err.Description, vbOKOnly + vbCritical

End Sub

Private Sub Form_Close()
    SaveData
End Sub

Private Sub Form_Load()
    Dim iAppPermission As Integer
    iAppPermission = UserAccess_Check(Me)
    
    GetErrorCode
    

End Sub

Public Sub RefreshData()
'    Me.Adj_Rationale = rsAuditClmHdr.Fields("Adj_Rationale")
    RefreshListBoxCodes rsAuditClmDiag, Me.lstDiagCodes, "Diag"
    RefreshListBoxCodes rsAuditClmDiagRev, Me.lstDiagCodesRevised, "Diag"
    RefreshListBoxCodes rsAuditClmProc, Me.lstProcCodes, "Proc"
    RefreshListBoxCodes rsAuditClmProcRev, Me.lstProcCodesRevised, "Proc"
    Me.Adj_Rationale = rsAuditClmHdr.Fields("Adj_Rationale")
    
    If rsAuditClmHdrAdditionalInfo.EOF And rsAuditClmHdrAdditionalInfo.BOF Then
        Me.ErrorCode.Value = ""
    Else
        Me.ErrorCode.Value = rsAuditClmHdrAdditionalInfo.Fields("ErrorCode")
    End If
    
    
    RefreshProcIPStatusIndicator
End Sub

Private Sub CmdAddDiag_Click()
  AddDiagCode
  SaveData
End Sub

Private Sub Form_LostFocus()
    SaveData
End Sub


'' 20130723: KD, Need to modifiy this to:
''  - See if this is a 'Therapy (Congress)' claim
''  - If so:
''      - Set the visibility of the Drop down
''      - Get the Error codes for the auditor to choose from
''      - SEt the main form's properties
''  - If not:
''      - Set the visibility (In visible)
''

Private Sub GetErrorCode()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim rst As ADODB.RecordSet
Dim strSQL As String
Dim oForm As Form_frm_AUDITCLM_Main
    
    
    strProcName = ClassName & ".GetErrorCode"
    
    Set oForm = Me.Parent.Form
    
    If ClaimOfRightReviewTypeDoesNotHaveLineLevelReason(oForm.CnlyClaimNum) = True Then
        Me.ErrorCode.visible = True
        
        Set oAdo = New clsADO
        oAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
        'Setting the ADO-class sqlstring to the specified SQL query statement
        'VS 11/11/2015 Changed udf_ErrorCode_Available to return records from xref_RecoveryReason table without any filtering.
        strSQL = "select ErrorCode, ErrorCodeDesc from cms_auditors_code.dbo.udf_ErrorCode_Available ('" & oForm.CnlyClaimNum & "', " & gintAccountID & ") "
        Set rst = oAdo.OpenRecordSet(strSQL)
    
        RefreshComboBoxFromRecordset rst, Me.ErrorCode, ""
        oForm.IsTherapyConcept = True
    Else
        oForm.IsTherapyConcept = False
        Me.ErrorCode.visible = False
    End If
    

Block_Exit:
    Set oAdo = Nothing
    Set rst = Nothing
    Set oForm = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName, , "Error getting list of Error Code: " + Err.Description, , Me.Parent.Form.CnlyClaimNum
    GoTo Block_Exit
End Sub



Private Sub lstDiagCodes_DblClick(Cancel As Integer)
  AddDiagCode
  SaveData
End Sub

Private Sub AddDiagCode()
    Dim intI As Long
    For intI = 0 To Me.lstDiagCodes.ListCount
        If Me.lstDiagCodes.Selected(intI) = True Then
            Me.lstDiagCodesRevised.AddItem CStr(Me.lstDiagCodesRevised.ListCount + 1) & ";" & Me.lstDiagCodes.Column(1, intI) & ";" & Me.lstDiagCodes.Column(2, intI)
            FormIsDirty
        End If
    Next
    SaveData
End Sub

Private Sub AddProcCode()
    Dim intI As Long
    For intI = 0 To Me.lstProcCodes.ListCount
        If Me.lstProcCodes.Selected(intI) = True Then
            Me.lstProcCodesRevised.AddItem CStr(Me.lstProcCodesRevised.ListCount + 1) & ";" & Me.lstProcCodes.Column(1, intI) & ";" & Me.lstProcCodes.Column(2, intI)
            FormIsDirty
        End If
    Next
    SaveData
End Sub

Private Sub CmdAddProcAll_Click()
   AddAllProcCodes
End Sub

Private Sub CmdAddDiagAll_Click()
   AddAllDiagCodes
End Sub

Private Sub AddAllDiagCodes()
    Dim intI As Long
    For intI = 0 To Me.lstDiagCodes.ListCount - 1
            Me.lstDiagCodesRevised.AddItem CStr(Me.lstDiagCodesRevised.ListCount + 1) & ";" & Me.lstDiagCodes.Column(1, intI) & ";" & Me.lstDiagCodes.Column(2, intI)
    Next
    
    If Me.lstDiagCodes.ListCount > 0 Then
        FormIsDirty
    End If
    SaveData
End Sub

Private Sub AddAllProcCodes()
    Dim intI As Long
    
    For intI = 0 To Me.lstProcCodes.ListCount - 1
            Me.lstProcCodesRevised.AddItem CStr(Me.lstProcCodesRevised.ListCount + 1) & ";" & Me.lstProcCodes.Column(1, intI) & ";" & Me.lstProcCodes.Column(2, intI)
    Next
    
    If Me.lstProcCodes.ListCount > 0 Then
        FormIsDirty
    End If
    SaveData
End Sub

Private Sub cmdRemoveDiag_Click()
    RemoveDiagCode
End Sub

Private Sub RemoveDiagCode()
    Dim intI As Long
    Dim holdindex As Long
    Dim holdstring As String
    
    holdindex = lstDiagCodesRevised.ListIndex
    
    If lstDiagCodesRevised.ListIndex >= 0 Then
        Me.lstDiagCodesRevised.RemoveItem (lstDiagCodesRevised.ListIndex)
        For intI = 0 To Me.lstDiagCodesRevised.ListCount - 1
            holdstring = Me.lstDiagCodesRevised.Column(1, 0) & ";" & Me.lstDiagCodesRevised.Column(2, 0)
            Me.lstDiagCodesRevised.RemoveItem (0)
            Me.lstDiagCodesRevised.AddItem CStr(intI + 1) & ";" & holdstring
        Next
    
        FormIsDirty
        
    End If
    
    If lstDiagCodesRevised.ListCount < holdindex + 1 Then
        lstDiagCodesRevised.Selected(lstDiagCodesRevised.ListCount - 1) = True
    Else
        lstDiagCodesRevised.Selected(holdindex) = True
    End If
    SaveData
End Sub


Private Sub CmdClearDiag_Click()
    Dim intI As Long
    
    If Me.lstDiagCodesRevised.ListCount > 0 Then
        FormIsDirty
    End If

    While Me.lstDiagCodesRevised.ListCount > 0
        Me.lstDiagCodesRevised.RemoveItem (0)
    Wend

    SaveData

End Sub

Private Sub cmdMoveDownDiag_Click()
    Dim temp
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim colCnt As Integer
    Dim tempval As String
    Dim Result As String
    
    If Me.lstDiagCodesRevised.ItemsSelected.Count = 0 Then
        Exit Sub
    End If
    
        colCnt = lstDiagCodesRevised.ColumnCount
    p = lstDiagCodesRevised.ListIndex + 1
    
    If p < lstDiagCodesRevised.ListCount Then
        temp = Split(lstDiagCodesRevised.RowSource, ";")
        
        ' swap rows down
        For j = 1 To colCnt
            tempval = temp(p * colCnt - j)
            temp(p * colCnt - j) = temp((p + 1) * colCnt - j)
            temp((p + 1) * colCnt - j) = tempval
        Next j
        
        For i = 0 To UBound(temp) Step colCnt
            Result = Result & Trim(CStr((i \ colCnt) + 1)) & ";"
            For j = 1 To colCnt - 1
                Result = Result & temp(i + j) & ";"
            Next j
        Next i
        
        Result = left(Result, Len(Result) - 1)
        lstDiagCodesRevised.RowSource = Result
        lstDiagCodesRevised.Selected(p) = True
        FormIsDirty
    End If

    SaveData

End Sub

Private Sub lstDiagCodesRevised_Click()
    Me.txtDiag = Me.lstDiagCodesRevised.Column(1, Me.lstDiagCodesRevised.ListIndex)
    Me.cmboPOACd = Me.lstDiagCodesRevised.Column(2, Me.lstDiagCodesRevised.ListIndex)
End Sub

Private Sub lstProcCodes_DblClick(Cancel As Integer)
  AddProcCode
End Sub

Private Sub lstProcCodesRevised_Click()
    Me.txtProc = Me.lstProcCodesRevised.Column(1, Me.lstProcCodesRevised.ListIndex)
    Me.txtProcDt = Me.lstProcCodesRevised.Column(2, Me.lstProcCodesRevised.ListIndex)
End Sub

Private Sub cmdRemoveProc_Click()
  RemoveProcCode
End Sub

Private Sub RemoveProcCode()
    Dim intI As Long
    Dim holdindex As Long
    Dim holdstring As String
    
    holdindex = lstProcCodesRevised.ListIndex
    
    If lstProcCodesRevised.ListIndex >= 0 Then
        Me.lstProcCodesRevised.RemoveItem (lstProcCodesRevised.ListIndex)
    
        For intI = 0 To Me.lstProcCodesRevised.ListCount - 1
            holdstring = Me.lstProcCodesRevised.Column(1, 0) & ";" & Me.lstProcCodesRevised.Column(2, 0)
            Me.lstProcCodesRevised.RemoveItem (0)
            Me.lstProcCodesRevised.AddItem CStr(intI + 1) & ";" & holdstring
        Next
        
        FormIsDirty
    
    End If
    
    If lstProcCodesRevised.ListCount < holdindex + 1 Then
        lstProcCodesRevised.Selected(lstProcCodesRevised.ListCount - 1) = True
    Else
        lstProcCodesRevised.Selected(holdindex) = True
    End If
    
    SaveData
    
End Sub

Private Sub CmdClearProc_Click()
    Dim intI As Long
    If Me.lstProcCodesRevised.ListCount > 0 Then
        FormIsDirty
    End If
    
    While Me.lstProcCodesRevised.ListCount > 0
        Me.lstProcCodesRevised.RemoveItem (0)
    Wend
    SaveData
End Sub

Private Sub RefreshListBoxCodes(rst As ADODB.RecordSet, _
                                lstBox As listBox, _
                                strType As String)

    On Error GoTo ErrHandler

    Dim ctr As Long
    Dim strItem As String
    
    lstBox.RowSource = vbNullString
    
    
    If rst.recordCount > 0 Then
        While Not rst.EOF
            'For ctr = 0 To rst.Fields.Count - 1
                
                If strType = "Diag" Then
                    strItem = rst.Fields("LineNum").Value & ";" & rst.Fields("DiagCd").Value & ";" & rst.Fields("PoaCd").Value
                ElseIf strType = "Proc" Then
                    strItem = rst.Fields("LineNum").Value & ";" & Nz(rst.Fields("ProcCd").Value, "") & ";" & IIf(Nz(rst.Fields("ProcDt").Value, "") = "", "1/1/1900", rst.Fields("ProcDt").Value)
                End If
                
            'Next ctr
            lstBox.AddItem strItem
            rst.MoveNext
        Wend
    End If

ExitNow:
    If rst.recordCount > 0 Then
        rst.MoveFirst
    End If

    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshProjectListing"
    GoTo ExitNow

End Sub

Private Sub FormIsDirty()
    If IsSubForm(Me) Then
        Me.Parent.RecordChanged = True
    End If
End Sub



Sub RefreshProcIPStatusIndicator()
         Dim MyAdo As clsADO
         Dim rs_CrossCodes As ADODB.RecordSet
         Dim SQLstr1, SQLstr2, SQLstr3, SQLstr4, SQLstr5, SQLstr6, SQLstr7, sqlString As String
         Dim IPPresent, NoDataPresent As Boolean
         
         Set MyAdo = New clsADO
         MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
         

        sqlString = ""
        
        
        'SQLstring = " SELECT ProcCode, CrossCode FROM GlobalHCCnlyMedicalCodesLibrary.dbo.ING_CrossCoder WHERE CrossCodeType=4 AND GETDATE() BETWEEN EffDt AND EndDt "
        
        'Set rs_xWalkProcDate = myado.OpenRecordSet(SQLstring)
        
        ' " & Me.Parent.Form.IPAdmitDate & "
        ' " & Me.Parent.Form.CnlyClaimNum & "
        
         SQLstr1 = " SELECT ClmProcCodes.LineNum, ClmProcCodes.ProcCd AS ProcCd, Xwalk.CrossCode AS CPTCode, ProcDate, "
         SQLstr2 = "    XCodeDateLoaded = CASE WHEN ISNULL(xwalk.CrossCode,'')='' THEN 'N' ELSE 'Y' END, "
         SQLstr3 = "    SIDateLoaded = CASE WHEN ISNULL(HCPCSSI.SI,'')='' THEN 'N' ELSE 'Y' END, "
         SQLstr4 = "    SI = CASE WHEN Xwalk.CrossCode IS NULL THEN 'NODATA' WHEN HCPCSSI.SI ='C' THEN 'C' Else '-' END "
         SQLstr5 = " FROM (SELECT ProcCd, LineNum, ProcDate = ISNULL(ProcDt,'" & Me.Parent.Form.IPAdmitDate & "') FROM CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_Proc WHERE cnlyClaimNum = '" & Me.Parent.Form.CnlyClaimNum & "') AS ClmProcCodes "
         SQLstr6 = " LEFT JOIN GlobalHCCnlyMedicalCodesLibrary.dbo.ING_CrossCoder AS XWalk ON XWalk.CrossCodeType=4 AND ClmProcCodes.ProcCd = XWalk.ProcCode AND ClmProcCodes.ProcDate BETWEEN XWalk.EffectiveDt AND XWalk.EndDt "
         SQLstr7 = " LEFT JOIN CMS_Auditors_Concepts.dbo.CPTCodeStatusInd AS HCPCSSI ON XWalk.CrossCode = HCPCSSI.HCPCS AND ClmProcCodes.ProcDate BETWEEN HCPCSSI.EffDate  AND HCPCSSI.EndDate "
         SQLstr8 = " ORDER BY ClmProcCodes.LineNum, ClmProcCodes.ProcCd, XWalk.CrossCode "
        
         sqlString = SQLstr1 & SQLstr2 & SQLstr3 & SQLstr4 & SQLstr5 & SQLstr6 & SQLstr7

         Set rs_CrossCodes = MyAdo.OpenRecordSet(sqlString)
        
        If Not rs_CrossCodes Is Nothing And Not (rs_CrossCodes.BOF Or rs_CrossCodes.EOF) Then
            Me.CmboIPIndicator.AddItem "ICD9; CPT; SI"
            rs_CrossCodes.MoveFirst
            While Not rs_CrossCodes.EOF
                Select Case rs_CrossCodes("SI") & rs_CrossCodes("XCodeDateLoaded") & rs_CrossCodes("SIDateLoaded")
                    Case "NODATA" & "N" & "N", "NODATA" & "N" & "Y", "NODATA" & "Y" & "N"
                        NoDataPresent = True
                    Case "C" & "Y" & "Y"
                        IPPresent = True
                        Me.CmboIPIndicator.AddItem rs_CrossCodes("ProcCd") & "; " & rs_CrossCodes("CPTCode") & "; " & rs_CrossCodes("SI")
                End Select
                rs_CrossCodes.MoveNext
            Wend
        End If
        
        If IPPresent Then 'if there was at least one proc code with IP indicator
            Me.lblIPIndicator.Caption = "-IP Indicator Present-"
            Me.lblIPIndicator.ForeColor = 206516
            Me.lblIPIndicator.visible = True
            Me.CmboIPIndicator.visible = True
        ElseIf NoDataPresent Then 'if there was at least one proc code with
            Me.lblIPIndicator.Caption = "XCoder Data NOT Avail !"
            Me.lblIPIndicator.ForeColor = 7425807
            Me.lblIPIndicator.visible = True
            Me.CmboIPIndicator.visible = False
        Else
            Me.lblIPIndicator.visible = False
            Me.CmboIPIndicator.visible = False
        End If
End Sub
