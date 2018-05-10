Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public mbFinished As Boolean
Public strMemo, strMailTo, strMailSubject As String

'Private Sub btn_Letters_Click()
'    On Error GoTo Error_Handler
'
'    Dim strICN, strLocalPath, strError, myCnlyClaimNum, ReturnPath As String
'    Dim strPasswd As String
'    strPasswdConfirm = "Dummy"
'
'    Set rsclm = Me.frm_APPEAL_hdr.Form.Recordset
'
'    If MsgBox("Do you want to generate Letter Packages for the " & rsclm.RecordCount & " claims below?", vbYesNo) <> vbYes Then Exit Sub
'
''    Disabled choose password and instead used randomize
''    While strPasswd <> strPasswdConfirm
''        strPasswd = InputBox("Choose a Password", "Password")
''        strPasswdConfirm = InputBox("Confirm Password", "Password")
''        If strPasswdConfirm = "" Then strPasswdConfirm = "Default"
''    Wend
'
'    mbFinished = False
'    DoCmd.OpenForm "frm_APPEAL_popup", acNormal
'    Do
'        DoEvents
'    Loop Until mbFinished
'
'    strPasswd = "medicare" 'RandomString("AN#HN#HNNHN")
'    strLocalPath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_PACKAGE\<payername>\"
'    DoCmd.Hourglass True
'    rsclm.MoveFirst
'
'    While Not (rsclm.EOF)
'        'Loop through each claim to generate package
'        strICN = Trim(rsclm.Fields("ICN"))
'        myCnlyClaimNum = Trim(rsclm.Fields("CnlyClaimNum"))
'        CreateAppeal strICN, myCnlyClaimNum, strLocalPath, strPasswd, strMemo, ReturnPath
'        If ReturnPath <> "" Then
'            SendsqlMail strMailSubject & " at " & ReturnPath, strMailTo, "", "", "*****************" & strMemo
'        End If
'        rsclm.MoveNext
'    Wend
'
'    DoCmd.Hourglass False
'    Exit Sub
'
'Error_Handler:
'    DoCmd.Hourglass False
'    strError = Err.Description
'    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
'
'End Sub

Private Sub btnAppeals_Click()
    On Error GoTo Error_Handler
    
    Dim strICN, strLocalPath, strError, myCnlyClaimNum, ReturnPath As String
    Dim strPasswd As String
    Dim strPasswdconfirm As String
    
    Dim rsclm As DAO.RecordSet '* 12/6/12 jc added declaration
    Dim strPayerName As String '* 12/6/12 jc added declaration
    Dim RtnValue As Long       '* 12/6/12 jc added declaration
    
    strPasswdconfirm = "Dummy"
    
    Set rsclm = Me.frm_APPEAL_hdr.Form.RecordSet
    
    If Me.DocType = "APPEAL" Then If MsgBox("Do you want to generate Appeals Packages for the " & rsclm.recordCount & " claims below?", vbYesNo) <> vbYes Then Exit Sub
    If Me.DocType = "LETTER" Then If MsgBox("Do you want to generate Letter Packages for the " & rsclm.recordCount & " claims below?", vbYesNo) <> vbYes Then Exit Sub
    If Me.DocType = "DOCONLY" Then If MsgBox("Do you want to generate Documentation Packages for the " & rsclm.recordCount & " claims below?", vbYesNo) <> vbYes Then Exit Sub
    
'    Disabled choose password and instead used randomize
'    While strPasswd <> strPasswdConfirm
'        strPasswd = InputBox("Choose a Password", "Password")
'        strPasswdConfirm = InputBox("Confirm Password", "Password")
'        If strPasswdConfirm = "" Then strPasswdConfirm = "Default"
'    Wend

    mbFinished = False
    DoCmd.OpenForm "frm_APPEAL_popup", acNormal
    Do
        DoEvents
    Loop Until mbFinished
    
    rsclm.MoveFirst
    strPayerName = Trim(rsclm.Fields("Payer"))
    If strPayerName = "PINNACLE/TRISPAN" Then strPayerName = "PINNACLE"
    
    strPasswd = "medicare" 'RandomString("AN#HN#HNNHN")
    If Me.DocType = "APPEAL" Then strLocalPath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\APPEALS\<payername>\"
    If Me.DocType = "LETTER" Then strLocalPath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\APPEALS\<payername>\"
    'If Me.DocType = "LETTER" Then strLocalPath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\LETTER_SENT\<payername>\"
    If Me.DocType = "DOCONLY" Then strLocalPath = "\\Ccaintranet.com\dfs-cms-fld\Audits\CMS\CMS MAIL\"
    strLocalPath = Replace(strLocalPath, "<payername>", strPayerName)
    
    DoCmd.Hourglass True
    rsclm.MoveFirst
    Dim flagDocType As Boolean
    
    While Not (rsclm.EOF)
        'Loop through each claim to generate package
        strICN = Trim(rsclm.Fields("ICN"))
        myCnlyClaimNum = Trim(rsclm.Fields("CnlyClaimNum"))
        CreateAppeal strICN, myCnlyClaimNum, strLocalPath, strPasswd, strMemo, ReturnPath, Me.DocType, strPayerName
        If ReturnPath <> "" Then
            SendsqlMail strMailSubject & " at " & ReturnPath, strMailTo, "", "", "*****************" & strMemo
        End If
        rsclm.MoveNext
    Wend
    
    'Zip all into a single package
    If Me.DocType = "LETTER" Then
        strLocalPath = strLocalPath & Date$ & "\"
        RtnValue = ShellWait("c:\program files (x86)\winzip\wzzip.exe -m -s" & strPasswd & " -x*.zip " & """" & strLocalPath & "ITR.zip" & """" & " " & """" & strLocalPath & "*.*" & """", vbMinimizedNoFocus)
    End If
     
'* 12/6/12 JC added close & set = nothing
    rsclm.Close
    Set rsclm = Nothing
    
    DoCmd.Hourglass False
    
Exit Sub
    

Error_Handler:
    DoCmd.Hourglass False
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frmGenericTab : RefreshData"
End Sub

Private Sub btnFilterPending_Click()
    Me.frm_APPEAL_hdr.Form.filter = "(APPEAL_Hdr.cnlyclaimnum not in (select cnlyclaimnum from appeal_dtl) or APPEAL_Hdr.cnlyclaimnum in ('X','Y'))"
    Me.frm_APPEAL_hdr.Form.FilterOn = True
End Sub

Private Sub btnSearchICN_Click()
'On Error GoTo Err_btnSearchICN_Click
    Dim strSQL As String
    'Creating a new instance of ADO-class variable
    Dim TestArray() As String
    Dim srchString As String
    Dim i As Integer
    
    If IsNull(Me.Icn) Then
        MsgBox "Nothing to Search", vbCritical
        Exit Sub
    End If

    Me.lstCnlyClaimNum.RowSource = ""
    
    srchString = Replace(Me.Icn, "'", "")

    TestArray = Split(srchString, ",")

    If Me.DataType = "Part-A" Then srchString = left(TestArray(0), 14)
    If Me.DataType = "Part-B" Then srchString = Right(TestArray(0), 13)

    Application.SysCmd acSysCmdInitMeter, "Searching", UBound(TestArray())
    
    For i = 1 To UBound(TestArray())

    If Me.DataType = "Part-A" Then srchString = srchString + "," + left(TestArray(i), 14)
    If Me.DataType = "Part-B" Then srchString = srchString + "," + Right(TestArray(i), 13)

    'srchString = srchString + "," + TestArray(i)
    'Show Progress Here
    Application.SysCmd acSysCmdUpdateMeter, , i + 1
    
    If Round(i / 7, 0) = i / 7 Then
        'Debug.Print srchString
        If i > 0 And i < UBound(TestArray()) Then
            addrecords (srchString)
            i = i + 1
            If Me.DataType = "Part-A" Then srchString = left(TestArray(i), 14)
            If Me.DataType = "Part-B" Then srchString = Right(TestArray(i), 13)
        End If
    End If

Next i

If srchString <> "" Then addrecords (srchString)

Exit_btnSearchICN_Click:
    Exit Sub

Err_btnSearchICN_Click:
    'If Err.Number = "13" Then
    '    MsgBox "Search is too long. Break it into more pieces"
    'Else
        MsgBox Err.Description
    'End If
    Resume Exit_btnSearchICN_Click
    
End Sub

Private Sub addrecords(srchString As String)
    Dim MyAdo As New clsADO
    Dim rs As RecordSet
    Dim cmd As ADODB.Command
    Dim i As Integer  '* 12/6/12 JC added declaration
    
    'Making a Connection call to SQL database?
    MyAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyAdo.CurrentConnection
    cmd.CommandText = "usp_Payer_Appeal"
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters("@ICNList") = srchString
    cmd.Parameters("@DocType") = Me.DocType

    Set rs = cmd.Execute
    Me.lstCnlyClaimNum.RowSourceType = "Value List"
    
    Dim ListHeader As String
    ListHeader = rs.Fields(0).Name
    
    If Me.lstCnlyClaimNum.RowSource = "" Then
        For i = 1 To rs.Fields.Count - 1
            ListHeader = ListHeader + ";" + rs.Fields(i).Name
        Next i
        Me.lstCnlyClaimNum.AddItem ListHeader
    End If
    Dim strappend As String
    rs.MoveFirst
    'rs.MoveNext
    While Not (rs.EOF Or rs.BOF)
        ListHeader = rs.Fields(0).Value
        
        For i = 1 To rs.Fields.Count - 1
            If i = 7 Then 'cnlyclaimnum
                strappend = IIf(IsNull(rs.Fields(i).Value), "", rs.Fields(i).Value)
            Else
                strappend = Format(IIf(IsNull(rs.Fields(i).Value), "", rs.Fields(i).Value), "")
            End If
            ListHeader = ListHeader + ";" + strappend
        Next i
        Me.lstCnlyClaimNum.AddItem ListHeader
        rs.MoveNext
    Wend
    
    'Refresh the Claim#s listbox
     Me.lstCnlyClaimNum.Requery

End Sub

Private Sub addRecordSet(rsparent As RecordSet, rsnew As ADODB.RecordSet)
Dim fld As ADODB.Field

rsnew.MoveFirst
'rsparent.Open
While Not (rsnew.EOF And rsnew.BOF)
    rsparent.AddNew
    For Each fld In rsparent.Fields
        rsparent.Fields(fld.Name) = rsnew.Fields(fld.Name)
    Next fld
    rsparent.Update
    rsnew.MoveNext
Wend
rsparent.Close
'Debug.Print rsparent.RecordCount
End Sub

Private Sub DocType_Change()
    If InStr(1, Me.frm_APPEAL_hdr.Form.filter, "DocType") > 0 Then
        Me.frm_APPEAL_hdr.Form.filter = Replace(Me.frm_APPEAL_hdr.Form.filter, "'APPEAL'", "'" + Me.DocType + "'")
        Me.frm_APPEAL_hdr.Form.filter = Replace(Me.frm_APPEAL_hdr.Form.filter, "'LETTER'", "'" + Me.DocType + "'")
        Me.frm_APPEAL_hdr.Form.filter = Replace(Me.frm_APPEAL_hdr.Form.filter, "'DOCONLY'", "'" + Me.DocType + "'")
    Else
        Me.frm_APPEAL_hdr.Form.filter = Me.frm_APPEAL_hdr.Form.filter & " and Doctype = '" & Me.DocType & "'"
    End If
    
    Me.frm_APPEAL_hdr.Form.FilterOn = True
End Sub

Private Sub Form_Load()
    RefreshAll
End Sub

Private Sub lstCnlyClaimNum_DblClick(Cancel As Integer)
On Error GoTo err_hndlr
    'Debug.Print Me.lstCnlyClaimNum.ItemData(1)
    'Debug.Print Me.Forms("frm_APPEAL_hdr")
   ' Dim Identity As New ClsIdentity
    Dim rtnCode, intStartPos, intLen As Integer
    Dim strSQL As String
    
    
    '* 12/6/12  jc added declarations
    Dim clm As Variant
    Dim rs As DAO.RecordSet
    
    

    
For Each clm In Me.lstCnlyClaimNum.ItemsSelected

    If left(lstCnlyClaimNum.Column(8, clm), 1) = "A" And Me.DocType = "APPEAL" Then
        Me.frm_APPEAL_hdr.Form.filter = Replace(Me.frm_APPEAL_hdr.Form.filter, "'X',", "'X','" + lstCnlyClaimNum.Column(7, clm) + "',")
        GoTo Exit_Fn
    End If

    If Right(lstCnlyClaimNum.Column(8, clm), 1) = "L" And Me.DocType = "LETTER" Then
        Me.frm_APPEAL_hdr.Form.filter = Replace(Me.frm_APPEAL_hdr.Form.filter, "'X',", "'X','" + lstCnlyClaimNum.Column(7, clm) + "',")
        GoTo Exit_Fn
    End If



    If Me.DocType = "APPEAL" Then 'Move to next queue for appeals only
        If lstCnlyClaimNum.Column(4, clm) <> 402 _
        And lstCnlyClaimNum.Column(4, clm) <> 501 _
        And lstCnlyClaimNum.Column(4, clm) <> 412 _
        And lstCnlyClaimNum.Column(4, clm) <> 413 _
        And lstCnlyClaimNum.Column(4, clm) <> 702 _
        And lstCnlyClaimNum.Column(4, clm) <> 502 _
        And lstCnlyClaimNum.Column(4, clm) <> 983 _
        And lstCnlyClaimNum.Column(4, clm) <> 451 _
        And lstCnlyClaimNum.Column(4, clm) <> 452 _
        And lstCnlyClaimNum.Column(4, clm) <> 453 _
        And lstCnlyClaimNum.Column(4, clm) <> 454 _
        And lstCnlyClaimNum.Column(4, clm) <> 414 _
        And lstCnlyClaimNum.Column(4, clm) <> 415 _
        And lstCnlyClaimNum.Column(4, clm) <> 512 _
        And lstCnlyClaimNum.Column(4, clm) <> 503 Then
            MsgBox "Claim not in correct status for Appeal", vbCritical
            GoTo Exit_Fn
        End If
        
        If lstCnlyClaimNum.Column(4, clm) = 415 Then
            MsgBox "This claim has already been considered a non-recovery", vbInformation
        End If
        
        'Move Claim to next queue below
        'Send a default tablename to let adoexetxt derive the Server. Specify DB since it's different for all Procs/views/codes
'        If lstCnlyClaimNum.Column(5, clm) <> "AP006" Then
'            strsql = "exec cms_auditors_code.dbo.usp_QUEUE_Manual_MoveToNextQueue '" & _
'                                            Me.lstCnlyClaimNum.Column(7, clm) & _
'                                            "','" & Me.lstCnlyClaimNum.Column(5, clm) & "','AP006" & _
'                                            "','" & Me.lstCnlyClaimNum.Column(4, clm) & _
'                                            "','" & Me.lstCnlyClaimNum.Column(4, clm) & _
'                                            "','Claim Appealed'"
'            'MsgBox strsql
'            rtnCode = AdoExeTxt(strsql, "AUDITCLM_Hdr", , , "CMS_AUDITORS_CODE")
'            If rtnCode <> vbTrue Then Exit Sub
'            'End move claim to next queue
'        End If
    End If
      
        Set rs = Me.frm_APPEAL_hdr.Form.RecordSet
        rs.AddNew
        rs!Icn = lstCnlyClaimNum.Column(2, clm)
        rs!Payer = lstCnlyClaimNum.Column(3, clm)
        rs!CnlyClaimNum = Me.lstCnlyClaimNum.ItemData(clm)
        rs!AppealReceiptDt = Now()
        rs!EnteredBy = Identity.UserName
        rs!EntryDt = Now()
        rs!DocType = Me.DocType
        rs.Update
        
Next clm
    
    Me.ApDetails.Requery
    
Exit_Fn:
    '* 12/6/12 JC added close and set = nothing
    rs.Close
    Set rs = Nothing

    Exit Sub
    
err_hndlr:
    Select Case (Err.Number)
    Case 3022
        MsgBox "Claim Already Appealed", vbExclamation
    Case Else
        MsgBox Err.Description & "|" & Err.Number
    End Select
End Sub

Private Sub RefreshAll()
    Me.lstCnlyClaimNum.RowSource = ""
    Me.PackageDetails.Form.filter = "PackageID=-1" 'All package details disappear
    Me.PackageDetails.Form.FilterOn = True
    Me.ApDetails.Form.filter = "CnlyClaimNum=''"
    Me.ApDetails.Form.FilterOn = True
    Me.frm_APPEAL_hdr.Form.filter = "(APPEAL_Hdr.cnlyclaimnum not in (select cnlyclaimnum from appeal_dtl) or APPEAL_Hdr.cnlyclaimnum in ('X','Y'))"
    Me.frm_APPEAL_hdr.Form.FilterOn = True
    tglPending.Value = True
End Sub

Private Sub tglPending_Click()
    If tglPending.Value = True Then
        Me.frm_APPEAL_hdr.Form.filter = "(APPEAL_Hdr.cnlyclaimnum not in (select cnlyclaimnum from appeal_dtl) or APPEAL_Hdr.cnlyclaimnum in ('X','Y'))"
        Me.frm_APPEAL_hdr.Form.FilterOn = True
    Else
        Me.frm_APPEAL_hdr.Form.FilterOn = False
    End If
End Sub
