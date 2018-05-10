Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrLocalHoldPath As String
Private mstrLocalPath As String
Private mstrCalledFrom As String
Private mstrMemberName As String
Private mstrMemberDOB As String
Private mstrClmFromDt As String
Private mstrClmThruDt As String

Private Enum FormType
    Master = 1
    PhillyOffice = 2
    HumanaOffice = 3
End Enum

Const CstrFrmAppID As String = "ImageLog"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Private Sub cmdDeleteRecord_Click()
    Dim iAns
    If Me.RecordSet.recordCount > 0 Then
        iAns = MsgBox("Are you sure you want to delete this record '" & CnlyClaimNum & "' ?", vbYesNo)
        If iAns = vbYes Then
            With Me.RecordSet
                .Delete
                .MoveNext
            End With
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    DoCmd.Close
End Sub

Private Sub cmdGetImageName_Click()
    Dim strImageName As String
    
    Call GetImageName(strImageName)
    
    DoCmd.OpenForm "frm_SCANNING_Generic_TextBox", acPreview, , , , acDialog, ("ImageName;" & strImageName)
End Sub

Private Sub cmdPrint_Click()
    Dim strParams As String
    Dim strImageName As String
    
    
    Call GetImageName(strImageName)
    Call GetImageInfo
    
    If ReceivedDt & "" = "" Then
        MsgBox "Please enter received date"
        Me.ReceivedDt.SetFocus
        GoTo Error_Exit
    End If
    
    strParams = CStr(ScannedDt) & ";" & CStr(ReceivedDt) & ";" & cnlyProvID & ";" & Icn & ";" & _
                CnlyClaimNum & ";" & strImageName & ";" & ImageType & ";" & mstrMemberName & ";" & mstrMemberDOB & _
                ";" & mstrClmFromDt & ";" & mstrClmThruDt & ";" & Me.ScanOperator & ";" & Me.ScanStation
                
    DoCmd.OpenReport "rpt_Scanning_Cover_page", acViewPreview, , , acWindowNormal, strParams

Error_Exit:

End Sub

Private Sub cmdValidate_Click()
    Dim fso As FileSystemObject
    Dim strImgPath As String
    Dim strPDFFile As String
    Dim strTIFFile As String
    Dim bFileExists As Boolean
    Dim strImageFile As String
    
    On Error Resume Next
    
    strImgPath = mstrLocalHoldPath & cnlyProvID
    
    Set fso = CreateObject("scripting.filesystemobject")
    strTIFFile = strImgPath & "\" & ImageName & ".tif"
    strPDFFile = strImgPath & "\" & ImageName & ".pdf"
    
    strImageFile = strTIFFile
    bFileExists = fso.FileExists(strTIFFile)
    If bFileExists = False Then
        bFileExists = fso.FileExists(strPDFFile)
        If bFileExists Then strImageFile = strPDFFile
    End If
    
    If bFileExists Then
        If chkDisplayImage Then
            Shell "explorer.exe " & strImageFile, vbNormalFocus
        Else
            MsgBox "Image exists!"
        End If
    Else
        MsgBox "ERROR: " & vbCrLf & Space(5) & "Image " & strImageFile & " does not exists!", vbCritical
    End If
    
    ' display image folder
    Shell "explorer.exe " & strImgPath, vbNormalFocus

    Set fso = Nothing
End Sub

Private Sub cmdViewError_Click()
    DoCmd.OpenForm "frm_SCANNING_View_Error", acPreview, , , , acDialog, (CnlyClaimNum & ";" & ErrMsg)
End Sub

Private Sub CnlyClaimNum_AfterUpdate()
    Dim rs As DAO.RecordSet
    Dim strSQL As String
        
    CnlyClaimNum = UCase(CnlyClaimNum) & ""
    If ImageType <> "PAC" Then
        strSQL = "select * from AUDITCLM_Hdr where CnlyClaimNum = '" & CnlyClaimNum & "' and AccountID = " & gintAccountID
        Set rs = CurrentDb.OpenRecordSet(strSQL)
        If rs.BOF = True And rs.EOF = True Then
            MsgBox "Claim is invalid"
            CnlyClaimNum.SetFocus
        Else
            Icn = rs("ICN")
            ProvNum = rs("ProvNum")
            AuditNum = rs("Adj_AuditNum")
            cnlyProvID = rs("CnlyProvID")
            mstrMemberName = rs("PatFirstInit") & Space(1) & rs("PatSurName")
            mstrMemberDOB = Format(rs("BeneBirthDt"), "mm-dd-yyyy")
            mstrClmFromDt = Format(rs("ClmFromDt"), "mm-dd-yyyy")
            mstrClmThruDt = Format(rs("ClmThruDt"), "mm-dd-yyyy")
            PageCnt.SetFocus
        End If
    End If
    
    If Nz(ImageType, "") <> "" Then
        strSQL = "select 1 from SCANNING_Image_Log_Tmp where ImageType = '" & ImageType & "'" & _
                 " and CnlyClaimNum = '" & CnlyClaimNum & "'" & _
                 " and ScannedDt >= #" & Date & "#"
        Set rs = CurrentDb.OpenRecordSet(strSQL)
        If Not (rs.EOF = True And rs.BOF = True) Then
            MsgBox "Warning: claim " & CnlyClaimNum & " has been entered earlier today.  Please check to make sure this is not a duplicate scanning", vbInformation
        End If
    End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    
    Dim dLetterReqDt As Date
    Dim strChkVal As String
    Dim strImageName As String
    
    If CnlyClaimNum & "" = "" Then
        MsgBox "CnlyClaimNum can not be blank", vbCritical
        CnlyClaimNum.SetFocus
        Exit Sub
        
        strChkVal = DLookup("CnlyClaimNum", "v_SCANNING_Claim_Info", "CnlyClaimNum = '" & Me.CnlyClaimNum & "'") & ""
        If strChkVal = "" Then
            MsgBox "Error: we did not request medical record for this claim. Claim = " & Me.CnlyClaimNum & ".  Please check with IT"
            Cancel = True
        End If
    End If
    
    If Icn & "" = "" And ImageType <> "PAC" Then
        MsgBox "ICN can not be blank", vbCritical
        Icn.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If ReceivedDt & "" = "" Then
        MsgBox "Received date can not be blank"
        ReceivedDt.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    ' check received date can not be before MR request date
    If IsDate(ReceivedDt) And Me.ImageType = "MR" Then
        dLetterReqDt = DLookup("LetterReqDt", "v_SCANNING_Prov_Info", "CnlyClaimNum = '" & Me.CnlyClaimNum & "'")
        If ReceivedDt < dLetterReqDt Then
            MsgBox "Received date can not be prior to medical record request date"
            Cancel = True
        End If
    End If
    
    If PageCnt & "" = "" Then
        MsgBox "Page count can not be blank", vbCritical
        PageCnt.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If ProvNum & "" = "" Then
        MsgBox "ProvNum can not be blank.", vbCritical
        ProvNum.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If cnlyProvID & "" = "" Then
        MsgBox "CnlyProvID can not be blank", vbCritical
        cnlyProvID.SetFocus
        Cancel = True
        Exit Sub
    End If
    
    If Me.NewRecord Then
        ScannedDt = Now()
        AccountID = gintAccountID
    End If
    
    If Nz(mstrLocalPath, "") = "" Then
        mstrLocalPath = "" & DLookup("LocalPath", "SCANNING_Config", "AccountID = " & gintAccountID)
        If Right$(mstrLocalPath, 1) <> "\" Then mstrLocalPath = mstrLocalPath & "\"
    End If
    Me.LocalPath = mstrLocalPath
    
    If ImageType = "PAC" Then
        CnlyClaimNum = cnlyProvID & ImageType & "-" & Format(ScannedDt, "yyyymmddhhmmss")
        Icn = cnlyProvID & ImageType
        ImageName = UCase(CnlyClaimNum)
    Else
        If Nz(ImageName, "") = "" Then
            Call GetImageName(strImageName)
            ImageName = strImageName
        Else
            ImageName = UCase(ImageName)
            CnlyClaimNum = UCase(CnlyClaimNum)
            ImageType = UCase(ImageType)
            
            If Mid(ImageName, Len(Me.cnlyProvID) + 1, Len(ImageType)) <> ImageType Then
                MsgBox "Error: Imagename has different image type.  Please re-check your cover sheet"
                Cancel = True
            End If
        End If
    End If
End Sub


Private Sub Form_Close()
    If mstrCalledFrom <> "" Then
        DoCmd.OpenForm mstrCalledFrom
    End If
End Sub


Private Sub Form_Current()
    If Me.NewRecord And ImageType = "PAC" Then
        CnlyClaimNum.Enabled = False
        CnlyClaimNum = ""
        Icn = ""
    Else
        If ImageType = "PAC" Then
            CnlyClaimNum.Enabled = False
        Else
            CnlyClaimNum.Enabled = True
        End If
    End If
    
    CnlyClaimNum.Requery
End Sub

Private Sub Form_Load()
    Dim strErrMsg As String
    Dim strErrSource As String
    Dim iAppPermission As String
    Dim iTempAccountID As String
    Dim iFormType As Integer
    
    
      
    Me.Caption = "Image Log"
    Me.ImageType.DefaultValue = ""
        
    Call Account_Check(Me)
    
    iAppPermission = UserAccess_Check(Me)
    If iAppPermission = 0 Then Exit Sub

    chkDisplayImage = 0
    
    strErrSource = "SCANNING_Data_Entry.Load"
    
    If IsSubForm(Me) Then
        cmdExit.visible = False
        lblAppTitle.visible = False
    End If
    
    Me.SortGroup = 1                ' default for listing by provider number
    
    ' Master = 1; Philly Office = 2; Humana office = 3
    iFormType = FormType.Master
    Select Case iFormType
        Case 1              ' Master version
            Me.ScannedOffice = 1
            Me.RecordSource = "select * from SCANNING_Image_Log_Tmp where AccountID = " & gintAccountID
        Case 2              ' Philly office
            Me.ScannedOffice = 1
            Me.ScannedOffice.Locked = True
            Me.RecordSource = "select * from SCANNING_Image_Log_Tmp where ScanOperator = '" & Identity.UserName & "' and AccountID = " & gintAccountID
        Case 3              ' Humana office
            Me.ScannedOffice = 2
            Me.ScannedOffice.Locked = True
            Me.RecordSource = "select * from SCANNING_Image_Log_Tmp where ScanOperator = '" & Identity.UserName & "' and AccountID = " & gintAccountID
    End Select
        
    
    Call ScannedOffice_AfterUpdate  ' set folder path
    
    On Error GoTo Err_handler
    
    ' set default drop down selection
    txtProvNum = ""
    txtProvNum.RowSource = "SELECT * FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " order by ProvNum"
    
    lblProvName.visible = False
    txtLetterReqDt = ""
    ScanOperator.DefaultValue = Chr(34) & Identity.UserName() & Chr(34)
    ScanStation.DefaultValue = Chr(34) & GetPCName() & Chr(34)

    If Me.OpenArgs() & "" <> "" Then
        mstrCalledFrom = Me.OpenArgs
    End If
    
    Me.ImageType.RowSource = "select * from SCANNING_XREF_ImageType where Active = 'Y'"
Exit_Sub:
    Exit Sub

Err_handler:
    If strErrMsg <> "" Then
        Err.Raise vbObjectError + 513, strErrSource, strErrMsg
    Else
        Err.Raise Err.Number, strErrSource, Err.Description
    End If
End Sub


Private Sub ImageType_AfterUpdate()
    Dim strSQL As String
    Dim rs As DAO.RecordSet
    
    ImageType.DefaultValue = Chr(34) & ImageType & Chr(34)
    If ImageType = "PAC" Then
        CnlyClaimNum.Enabled = False
        CnlyClaimNum = ""
        Icn = ""
        PageCnt.SetFocus
    Else
        CnlyClaimNum.Enabled = True
    End If
    
    If Nz(CnlyClaimNum, "") <> "" Then
        If Nz(ScannedDt, "") <> "" Then
            ImageName = UCase(CnlyClaimNum) & ImageType & "-" & Format(ScannedDt, "yyyymmddhhmmss")
        End If
        
        strSQL = "select 1 from SCANNING_Image_Log_Tmp where ImageType = '" & ImageType & "'" & _
                 " and CnlyClaimNum = '" & CnlyClaimNum & "'" & _
                 " and ScannedDt >= #" & Date & "#"
        Set rs = CurrentDb.OpenRecordSet(strSQL)
        If Not (rs.EOF = True And rs.BOF = True) Then
            MsgBox "Warning: claim " & CnlyClaimNum & " has been entered earlier today.  Please check to make sure this is not a duplicate scanning", vbInformation
        End If
    End If

End Sub


Private Sub ReceivedDt_Exit(Cancel As Integer)
    If ReceivedDt.Value <> "" Then
        If IsDate(ReceivedDt.Value) Then
            ReceivedDt = Format(ReceivedDt, "mm-dd-yyyy")
            ReceivedDt.DefaultValue = "#" & ReceivedDt & "#"
        Else
            MsgBox Chr(34) & ReceivedDt.Value & Chr(34) & " is not a valid date." & vbCrLf & "Please enter a valid received date"
            Cancel = True
        End If
    End If
End Sub

Private Sub ScannedOffice_AfterUpdate()
    Dim iLocalAccountID As Integer
    
    Select Case ScannedOffice
        Case 1      ' Philly office
            iLocalAccountID = gintAccountID
        Case 2      ' Humana office
            iLocalAccountID = 5
    End Select
    
    mstrLocalHoldPath = "" & DLookup("LocalHoldPath", "SCANNING_Config", "AccountID = " & iLocalAccountID)
    If Right$(mstrLocalHoldPath, 1) <> "\" Then mstrLocalHoldPath = mstrLocalHoldPath & "\"

    mstrLocalPath = "" & DLookup("LocalPath", "SCANNING_Config", "AccountID = " & iLocalAccountID)
    If Right$(mstrLocalPath, 1) <> "\" Then mstrLocalPath = mstrLocalPath & "\"

    If mstrLocalHoldPath = "" Then
        MsgBox "Temporary image hold path is not defined.  Please set the hold path first", vbCritical
    End If
    
    If FolderExist(mstrLocalHoldPath) = False Then
        MsgBox "Image path '" & mstrLocalHoldPath & "' does not exists or in accessible.  Please check.", vbCritical
    End If
    
End Sub

Private Sub SortGroup_AfterUpdate()
    Select Case SortGroup
        Case 1              ' Provider number
            txtProvNum.RowSource = "SELECT ProvNum, CnlyProvID, ProvName, LetterReqDt, ReferenceNum, ClmCnt, AccountID FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " order by ProvNum"
            lblProvider.Caption = "Provider Number"
            txtProvNum.ColumnWidths = "1440;1440;2880;1152;1728;576"
        Case 2              ' Connolly provider number
            txtProvNum.RowSource = "SELECT CnlyProvID, ProvNum, ProvName, LetterReqDt, ReferenceNum, ClmCnt, AccountID FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " order by CnlyProvID"
            lblProvider.Caption = "CnlyProvID"
            txtProvNum.ColumnWidths = "1440;1440;2880;1152;1728;576"
        Case 3              ' Provider name
            txtProvNum.RowSource = "SELECT ProvName, ProvNum, CnlyProvID, LetterReqDt, ReferenceNum, ClmCnt, AccountID FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " order by ProvName"
            lblProvider.Caption = "Provider Name"
            txtProvNum.ColumnWidths = "2880;1440;1440;1152;1728;576"
        Case 4              ' Sent date
            txtProvNum.RowSource = "SELECT LetterReqDt, ProvNum, CnlyProvID, ProvName, ReferenceNum, ClmCnt, AccountID FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " order by LetterReqDt"
            lblProvider.Caption = "Letter Sent Dt"
            txtProvNum.ColumnWidths = "1152;1440;1440;2880;1728;576"
        Case 5              ' Reference Number / InstanceID
            txtProvNum.RowSource = "SELECT ReferenceNum, ProvNum, CnlyProvID, ProvName, LetterReqDt, ClmCnt, AccountID FROM v_SCANNING_Prov_Info WHERE AccountID=" & gintAccountID & " and ReferenceNum is not null order by LetterReqDt"
            lblProvider.Caption = "Reference #"
            txtProvNum.ColumnWidths = "1728;1440;1440;2880;1152;576"
    End Select
    txtProvNum.Requery
End Sub

Private Sub txtProvNum_AfterUpdate()
    Dim strProvNum As String
    
    If txtProvNum.ListIndex >= 0 Then
        Select Case SortGroup
            Case 1
                strProvNum = txtProvNum.Column(0)
                lblProvName.Caption = txtProvNum.Column(2)
                txtLetterReqDt = txtProvNum.Column(3)
            Case 2
                strProvNum = txtProvNum.Column(1)
                lblProvName.Caption = txtProvNum.Column(2)
                txtLetterReqDt = txtProvNum.Column(3)
            Case 3
                strProvNum = txtProvNum.Column(1)
                lblProvName.Caption = txtProvNum.Column(0)
                txtLetterReqDt = txtProvNum.Column(3)
            Case 4
                strProvNum = txtProvNum.Column(1)
                lblProvName.Caption = txtProvNum.Column(3)
                txtLetterReqDt = txtProvNum.Column(0)
            Case 5
                strProvNum = txtProvNum.Column(1)
                lblProvName.Caption = txtProvNum.Column(3)
                txtLetterReqDt = txtProvNum.Column(4)
        End Select
        lblProvName.visible = True
        CnlyClaimNum.RowSource = "select * from v_SCANNING_Claim_Info where ProvNum = '" & strProvNum & "' and LetterReqDt = #" & txtLetterReqDt & "#"
    Else
        lblProvName.visible = False
        txtLetterReqDt = ""
        CnlyClaimNum.RowSource = ""
    End If
End Sub


Private Sub GetImageName(strImageName As String)
    If Nz(Me.ScannedDt, "") = "" Then
        Me.ScannedDt = Now()
    End If
       
    If Me.ImageName & "" = "" Then
        Me.ImageName = UCase(Me.cnlyProvID) & ImageType & "-" & Format(ScannedDt, "yyyymmddhhmmss")
    End If
    
    strImageName = Me.ImageName
End Sub


Private Sub GetImageInfo()

    Dim rs As DAO.RecordSet
    Dim strSQL As String
    
    If Me.ScanOperator & "" = "" Then
        Me.ScanOperator = Identity.UserName()
    End If
    
    If Me.ScanStation & "" = "" Then
        Me.ScanStation = GetPCName()
    End If
        
    If ImageType <> "PAC" Then
        strSQL = "select * from AUDITCLM_Hdr where CnlyClaimNum = '" & CnlyClaimNum & "' and AccountID = " & gintAccountID
        Set rs = CurrentDb.OpenRecordSet(strSQL)
        If rs.BOF = True And rs.EOF = True Then
            MsgBox "Claim is invalid"
            CnlyClaimNum.SetFocus
        Else
            mstrMemberName = rs("PatFirstInit") & Space(1) & rs("PatSurName") & ""
            mstrMemberDOB = Format(rs("BeneBirthDt"), "mm-dd-yyyy") & ""
            mstrClmFromDt = Format(rs("ClmFromDt"), "mm-dd-yyyy") & ""
            mstrClmThruDt = Format(rs("ClmThruDt"), "mm-dd-yyyy") & ""
        End If
    End If

    Set rs = Nothing
End Sub
