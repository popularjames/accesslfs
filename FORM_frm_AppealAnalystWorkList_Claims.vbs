Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Public clmLstRowCount As Integer
    Public ClmLstRowPosition As Integer
    Public strCSFilePath As String
    Public strPPFilePath As String
    
    Public strPPCVFilePath As String
    Public strCPMaxFilePath As String
    Public strCPJudgeFilePath As String
    Public strCPProvFilePath As String
    
    Private Const CPMAX As String = "PP_FXP_MAX"
    Private Const CPJudge As String = "PP_FXP_Judge"
    Private Const CPProv As String = "PP_FXP_Prov"
    Private Const CSIRF As String = "CSIRFTemplate.docx"
    Private Const CSHH As String = "CSHHTemplate.docx"
    Private Const PPHH As String = "PP_HH_Template.docx"
    
Private Sub Form_Load()
    Me.cmdUpdateAnalyst.Enabled = False
    
End Sub


Private Sub cmdCS_Click()
'********************************************************************************
'Clinical Summary Templates:  The Template Type, Name, Path are in the table dbo.Appeal_Hearing_Package_Analyst_Templates.
'Clinical Summary templates and hearing files are stored in the \\ccaintranet.com\dfs-cms-fld\Audits\CMS\ALJHearing folder.
'The Clinical Summary Type (Me.CS) is determined using the function dbo.udf_GetALJHearingClinicalSummary
'The CS Type is the link in the template table.

'********************************************************************************
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim strfilePathSave As String
    Dim strFilePathTemplate As String
    Dim strTemplateName As String
    Dim strSQL As String
    
    Dim MyAdo As New clsADO
    Dim CSTemplateRS As ADODB.RecordSet
    
    'IF: Is there a defined Clinical Summary?
    If Nz(Me.CS, "") <> "" Then
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strSQL = "Select TemplateName, TemplateFilePath, SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & Me.CS & "'"
        Set CSTemplateRS = MyAdo.OpenRecordSet(strSQL)
    
        'IF: Is there a template defined?
        If CSTemplateRS.EOF = True And CSTemplateRS.BOF = True Then
            MsgBox ("The Clinical Summary template has not been set up for the Clinical Summary type " & Me.CS & ". Please alert the system administrator.")
        'IF: Is there a template defined?
        Else
            strfilePathSave = CSTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & CSTemplateRS("SaveFilename") & ".docx"
            strFilePathTemplate = CSTemplateRS("TemplateFilePath") & CSTemplateRS("TemplateName")
            strTemplateName = CSTemplateRS("Templatename")
            
            'IF: Is there a Clinical Summary saved for the claim?  (IF: CS saved, ELSEIF: get template, ELSE: cannot find the physical file)
            If FileExists(strfilePathSave) = True Then 'If a file saved, open up Word
                Shell "explorer.exe " & strfilePathSave, vbNormalFocus
                
            'IF: Is there a Clinical Summary saved for the claim?
            ElseIf FileExists(strFilePathTemplate) = True Then 'If no CS yet, pull up the template & populate the fields; will then SaveAs, close and re-open using shell command
            'Need to close & re-open because of focus issues if there are multiple documents open.
                Set wrdApp = CreateObject("Word.Application")
                wrdApp.visible = True
                Set wrdDoc = wrdApp.Documents.Open(strFilePathTemplate)
                
                With wrdDoc
                    .FormFields("fldALJNumber").Result = Me.ALJAppealNumber
                    .FormFields("fldDOH").Result = Nz(Me.HearingDateTime, "")
                    .FormFields("fldJudgeLastName").Result = Nz(Me.ALJudgeName, "")
                    .FormFields("fldICN").Result = Me.Icn
                    .FormFields("fldAppellant").Result = Me.AppellantName
                    .FormFields("fldServiceDateFrom").Result = Me.ServiceFromDate
                    .FormFields("fldServiceDateTo").Result = Me.ServiceToDate
                    .FormFields("fldAge").Result = Me.Age
                    .FormFields("fldSex").Result = Me.BeneSex
                    .FormFields("fldBeneficiary").Result = Me.Beneficiary
                    .FormFields("fldBeneficiaryFI").Result = Me.BeneficiaryFirstInitial
                    
                    If strTemplateName <> CSIRF And strTemplateName <> CSHH Then
                        .FormFields("fldDRG").Result = Me.DRG
                    End If
                    
                End With
                
                wrdApp.Documents(strFilePathTemplate).SaveAs2 (strfilePathSave)
                wrdApp.Documents(strfilePathSave).Activate
                wrdApp.Documents(strfilePathSave).Close SaveChanges:=wdDoNoSaveChanges
                wrdApp.Application.Quit
                
                Set wrdDoc = Nothing
                Set wrdApp = Nothing
                
                Shell "explorer.exe " & strfilePathSave, vbNormalFocus
                
            'IF: Is there a Clinical Summary saved for the claim?
            Else
                MsgBox ("There is an error with the template.  Please contact the system administrator")
            End If
        
        'IF: Is there a template defined?
        End If
    
    'IF: Is there a defined Clinical Summary?
     Else
        MsgBox ("There is no required Clinical Summary defined for this Claim.  Please contact the system administrator.")
     End If
        
Exit_Sub:
    Set MyAdo = Nothing
    Set cmd = Nothing
    
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
    
End Sub



Private Sub cmdPP_Click()
'********************************************************************************
'Position Paper Templates:  The Template Type, Name, Path are in the table dbo.Appeal_Hearing_Package_Analyst_Templates.
'Position Paper templates and hearing files are stored in the \\ccaintranet.com\dfs-cms-fld\Audits\CMS\ALJHearing folder.
'The Position Paper type (Me.PP) is determined by the function dbo.udf_GetALJHearingPositionPaper
'The PP Type is the link in the template table

'********************************************************************************
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim strfilePathSave As String
    Dim strFilePathTemplate As String
    Dim strTemplateName As String
    Dim strSQL As String
    
    Dim MyAdo As New clsADO
    Dim PPTemplateRS As ADODB.RecordSet
    
    'IF: Is there a defined Position Paper?
    If Nz(Me.PP, "") <> "" Then
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strSQL = "Select TemplateName, TemplateFilePath, SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & Me.PP & "'"
        Set PPTemplateRS = MyAdo.OpenRecordSet(strSQL)
    
        'IF: Is there a template defined?
        If PPTemplateRS.EOF = True And PPTemplateRS.BOF = True Then 'The udf_GetALJHearingPositionPaper is defined, but there is no corresponding record in the dbo.Appeal_hearing_Package_Analyst_Templates
            MsgBox ("The Position Paper template has not been set up for the definition " & Me.PP & ". Please alert the system administrator")
        'IF: Is there a template defined?
        Else
            strfilePathSave = PPTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & PPTemplateRS("SaveFilename") & ".docx"
            strFilePathTemplate = PPTemplateRS("TemplateFilePath") & PPTemplateRS("TemplateName")
            strTemplateName = PPTemplateRS("Templatename")
            
            'IF: Is there a PP saved for the claim? (IF: PP saved; ELSEIF: get template; ELSE: cannot find the physical file)
            If FileExists(strfilePathSave) = True Then 'Copy already exists
                Shell "explorer.exe " & strfilePathSave, vbNormalFocus
                
            'IF: Is there a PP saved for the claim?
            ElseIf FileExists(strFilePathTemplate) = True Then
                Set wrdApp = CreateObject("Word.Application")
                wrdApp.visible = True
                Set wrdDoc = wrdApp.Documents.Open(strFilePathTemplate)
                
                With wrdDoc
                    .FormFields("fldJudgeFirstName").Result = Nz(Me.ALJJudgeFirstname, "")
                    .FormFields("fldJudgeLastName").Result = Nz(Me.ALJudgeName, "")
                    .FormFields("fldALJNumber").Result = Me.ALJAppealNumber
                    .FormFields("fldJudgeLastName2").Result = Nz(Me.ALJudgeName, "")
                    .FormFields("fldServiceDateFrom").Result = Me.ServiceFromDate
                    .FormFields("fldServiceDateTo").Result = Me.ServiceToDate
                    .FormFields("fldAge").Result = Me.Age
                    .FormFields("fldSex").Result = Me.BeneSex
                    .FormFields("fldFieldOffice").Result = Me.txtFieldOffice
                    
                    If strTemplateName <> PPHH Then
                        .FormFields("fldDOH").Result = Nz(Me.HearingDateTime, "")
                    Else
                        .FormFields("fldServiceDateFrom2").Result = Me.ServiceFromDate
                        .FormFields("fldICN").Result = Me.Icn
                    End If
                    
                End With
                
                wrdApp.Documents(strFilePathTemplate).SaveAs2 (strfilePathSave)
                wrdApp.Documents(strfilePathSave).Activate
                wrdApp.Documents(strfilePathSave).Close SaveChanges:=wdDoNoSaveChanges
                wrdApp.Application.Quit
                
                Set wrdDoc = Nothing
                Set wrdApp = Nothing
                
                Shell "explorer.exe " & strfilePathSave, vbNormalFocus
            
            'IF: Is there a PP saved for the claim?
            Else
                MsgBox ("There is an error with the template.  Please contact the system administrator")
            End If
        'Is there a template defined?
        End If
        
    'Is there a defined PP?
    Else
        MsgBox ("There is no required Position Paper defined for this Claim.  If you think this is an error, please contact the system administrator.")
    End If
    
Exit_Sub:
    Set MyAdo = Nothing
    Set cmd = Nothing
    
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
    
End Sub


Private Sub CnlyClaimNum_DblClick(Cancel As Integer)

     If Me.CnlyClaimNum & "" <> "" Then
        DisplayAuditClmMainScreen Me.CnlyClaimNum
    End If

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Me.cmdUpdateAnalyst.Enabled = True
    
    Call GetRecords(Me)
    
End Sub

Function GetRecords(frm As Form)
    'Define the starting position and number of records selected.
    'Have to do as didn't initially realize that a continuous form would not retain the records selected
    With frm
        ClmLstRowPosition = SelTop
        clmLstRowCount = SelHeight
    End With

End Function

Private Sub ICN_DblClick(Cancel As Integer)

     If Me.CnlyClaimNum & "" <> "" Then
        DisplayAuditClmMainScreen Me.CnlyClaimNum
    End If

End Sub

Private Sub cmdUpdateAnalyst_Click()
'Allow user to tag multiple claims with an Analyst name
    Dim i As Long
    Dim strCnlyClaimNum As String
    Dim strAnalyst As String
    Dim strDocumentUpdateType As String
    Dim mstrUserName As String
    
    Dim AnalystUpdateAdo As clsADO
    Dim cmd As ADODB.Command
    Dim strErrMsg As String
    Dim strCompleteMsg As String
    
    On Error GoTo Err_handler
    
    strAnalyst = Nz(cmbUpdateAnalyst.Value, "")
    strDocumentUpdateType = "Analyst"
    mstrUserName = GetUserName()
    strErrMsg = ""
    strCompleteMsg = ""
    
    If clmLstRowCount = 0 Then
        MsgBox ("Please select the rows to be updated.")
    Else
        'If the user selects from the bottom up - need to get the top of the selection
        If Me.CurrentRecord > ClmLstRowPosition Then
            DoCmd.GoToRecord , , acGoTo, ClmLstRowPosition
        End If
        
        'strAnalyst = cmbUpdateAnalyst.Value
        'MsgBox strAnalyst
        
        Set AnalystUpdateAdo = New clsADO
        AnalystUpdateAdo.ConnectionString = GetConnectString("v_Code_Database")
        
        'Loop through the selected records to update
        For i = ClmLstRowPosition To ClmLstRowPosition + (clmLstRowCount - 1)
            strCnlyClaimNum = Me.CnlyClaimNum
            
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = AnalystUpdateAdo.CurrentConnection
            cmd.commandType = adCmdStoredProc
            cmd.CommandText = "dbo.usp_AppealHearingPackageAnalystsDocumentsUpdate"
            cmd.Parameters.Refresh
            cmd.Parameters("@pDocumentUpdateType") = strDocumentUpdateType
            cmd.Parameters("@pCnlyClaimNum") = strCnlyClaimNum
            cmd.Parameters("@pAnalyst") = strAnalyst
            cmd.Parameters("@pUser") = mstrUserName
            cmd.Execute
            
            strErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
            If cmd.Parameters("@Return_Value") <> 0 Or Nz(strErrMsg, "") <> "" Then
                If strErrMsg = "" Then strErrMsg = "Error executing stored procedure dbo.usp_AppealHearingPackageAnalystDocumentsUPdate"
            Err.Raise 50001, "dbo.usp_AppealHearingPackageAnalystDocumentsUpdate", strErrMsg
            End If
            
            Me.RecordSet.MoveNext
        Next
    End If
    
    Me.cmdUpdateAnalyst.Enabled = False
    Me.cmbUpdateAnalyst.Value = ""

Exit_Sub:
    Set AnalystUpdateAdo = Nothing
    Set cmd = Nothing
    Exit Sub

Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    
    Me.cmdUpdateAnalyst.Enabled = False
    Me.cmbUpdateAnalyst.Value = ""
    
    Resume Exit_Sub

End Sub

Sub TemplatesExists()
Dim CSAdo As New clsADO
Dim CSTemplateRS As ADODB.RecordSet

Dim PPAdo As New clsADO
Dim PPTemplateRS As ADODB.RecordSet

Dim PPCVAdo As New clsADO
Dim PPCVTemplateRS As ADODB.RecordSet

Dim FXPAdo As New clsADO
Dim FXPTemplateRS As ADODB.RecordSet

Dim FXMAdo As New clsADO
Dim FXMTemplateRS As ADODB.RecordSet

Dim FXJAdo As New clsADO
Dim FXJTemplateRS As ADODB.RecordSet

Dim strCSSQL As String
Dim strPPSQL As String
Dim strPPCVSQL As String
Dim strFXPSQL As String
Dim strFXMSQL As String
Dim strFXJSQL As String

Dim strCnlyClaimNum As String

    Set CSAdo = New clsADO
    CSAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strCSSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & Me.CS & "'"
    Set CSTemplateRS = CSAdo.OpenRecordSet(strCSSQL)
    
    Set PPAdo = New clsADO
    PPAdo.ConnectionString = GetConnectString("v_DATA_Database")
    strPPSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & Me.PP & "'"
    Set PPTemplateRS = CSAdo.OpenRecordSet(strPPSQL)
    
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
      
        Set PPCVAdo = New clsADO
        PPCVAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strPPCVSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & PP_CV_Type() & "'"
        Set PPCVTemplateRS = CSAdo.OpenRecordSet(strPPCVSQL)
        
        Set FXJAdo = New clsADO
        FXJAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strFXJSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & CPJudge & "'"
        Set FXJTemplateRS = CSAdo.OpenRecordSet(strFXJSQL)
        
        Set FXMAdo = New clsADO
        FXMAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strFXMSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & CPMAX & "'"
        Set FXMTemplateRS = CSAdo.OpenRecordSet(strFXMSQL)
        
        Set FXPAdo = New clsADO
        FXPAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strFXPSQL = "Select SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & CPProv & "'"
        Set FXPTemplateRS = CSAdo.OpenRecordSet(strFXPSQL)
        
    End If
    
    If CSTemplateRS.EOF = False And CSTemplateRS.BOF = False Then
        strCSFilePath = CSTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & CSTemplateRS("SaveFileName") & ".docx"
    End If
    
    If Nz(Me.PP, "") <> "" Then
        If PPTemplateRS.EOF = False And PPTemplateRS.BOF = False Then
            strPPFilePath = PPTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & PPTemplateRS("SaveFileName") & ".docx"
        End If
    Else
        strPPFilePath = ""
    End If
    
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
        If PPCVTemplateRS.EOF = False And PPCVTemplateRS.BOF = False Then
            strPPCVFilePath = PPCVTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & PPCVTemplateRS("SaveFileName") & ".docx"
        End If
    Else
        strPPCVFilePath = ""
    End If
    
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
        If FXJTemplateRS.EOF = False And FXJTemplateRS.BOF = False Then
            strCPJudgeFilePath = FXJTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & FXJTemplateRS("SaveFileName") & ".docx"
        End If
    Else
        strCPJudgeFilePath = ""
    End If
    
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
        If FXMTemplateRS.EOF = False And FXMTemplateRS.BOF = False Then
            strCPMaxFilePath = FXMTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & FXMTemplateRS("SaveFileName") & ".docx"
        End If
    Else
        strCPMaxFilePath = ""
    End If
    
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
        If FXPTemplateRS.EOF = False And FXPTemplateRS.BOF = False Then
            strCPProvFilePath = FXPTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & FXPTemplateRS("SaveFileName") & ".docx"
        End If
    Else
        strCPProvFilePath = ""
    End If
    
End Sub


Private Sub cmdSave_Click()
Dim CompleteUpdateADO As clsADO
Dim cmd As ADODB.Command
Dim strErrMsg As String
Dim strCompleteMsg As String


Dim strCnlyClaimNum As String
Dim strPCSFilePath As String
Dim strPPPFilePath As String
Dim strResponseComplete As String
Dim strResponseComplete2 As String
Dim strResponseCompleteMsg As String
Dim strDocumentUpdateType As String
Dim mstrUserName As String

On Error GoTo Err_handler

strDocumentUpdateType = "Complete"
mstrUserName = GetUserName()
strErrMsg = ""
strCompleteMsg = ""

If Me.workcomplete = "-1" Then
    strResponseComplete = MsgBox("This claim has already been marked as 'Complete'.  Do you wish to proceed?", vbOKCancel)
End If

If Nz(strResponseComplete, "") <> 2 Then
    'MsgBox "Will proceed with update"
    
    'VS 11/10/2014 Add cover pages for Position papers that will be completed, but the hearing had not been scheduled for these yet.
    If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
       Call WritePP_CV
    End If
    
    Call TemplatesExists
    'MsgBox strCnlyClaimNum
'    MsgBox Nz(strCSFilePath, "")
'    MsgBox Nz(strPPFilePath, "")
    
    If FileExists(strCSFilePath) = False Then
    'If Nz(strCSFilePath, "") = "" Or (Nz(strPPFilePath, "") = "" And Nz(Me.PP, "") <> "") Then
        strResponseCompleteMsg = "Clinical Summary; "
    Else
        strPCSFilePath = strCSFilePath 'Will need to pass the FilePath only if it exists
    End If
    
    If (FileExists(strPPFilePath) = False And Nz(Me.PP, "") <> "") Then
        strResponseCompleteMsg = strResponseCompleteMsg + "Position Paper for " + Me.PP
    ElseIf FileExists(strPPFilePath) = True Then
        strPPPFilePath = strPPFilePath
    End If
    
    If Nz(strResponseCompleteMsg, "") <> "" Then
        strResponseComplete2 = MsgBox("This Claim is missing the noted expected documentation: " + strResponseCompleteMsg + " Do you wish to procceed?", vbOKCancel)
    End If
    
    'MsgBox strResponseComplete2
    
    If Nz(strResponseComplete2, "") = 2 Then
        'MsgBox "Do Nothing"
    Else
        'MsgBox "Proceed with Update"
        
        Set CompleteUpdateADO = New clsADO
        CompleteUpdateADO.ConnectionString = GetConnectString("v_Code_Database")
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = CompleteUpdateADO.CurrentConnection
        cmd.commandType = adCmdStoredProc
        cmd.CommandText = "dbo.usp_AppealHearingPackageAnalystsDocumentsUpdate"
        cmd.Parameters.Refresh
        cmd.Parameters("@pDocumentUpdateType") = strDocumentUpdateType
        cmd.Parameters("@pCnlyClaimNum") = Me.CnlyClaimNum
        cmd.Parameters("@pPPFilePath") = strPPPFilePath
        cmd.Parameters("@pCSFilePath") = strPCSFilePath
        cmd.Parameters("@pUser") = mstrUserName
        cmd.Parameters("@pBeneficiaryLastName") = Nz(Me.Beneficiary, "BeneficaryNameMissing")
        cmd.Parameters("@pBeneficiaryFirstName") = Nz(Me.BeneficiaryFirstInitial, "_")
        cmd.Parameters("@pALJNumber") = Nz(Me.ALJAppealNumber, "ALJNumberMissing")
        
        If Me.HearingStatus = "Not Scheduled. Judge name known" And Nz(Me.PP, "") <> "" Then
                 cmd.Parameters("@pPPCVFilePath") = strPPCVFilePath
                 cmd.Parameters("@pPPFXJFilePath") = Nz(strCPJudgeFilePath, "")
                 cmd.Parameters("@pPPFXMFilePath") = Nz(strCPMaxFilePath, "")
                 cmd.Parameters("@pPPFXPFilePath") = Nz(strCPProvFilePath, "")
        End If
        cmd.Execute
    
        strErrMsg = Nz(cmd.Parameters("@pErrMsg"), "")
        If cmd.Parameters("@Return_Value") <> 0 Or Nz(strErrMsg, "") <> "" Then
            If strErrMsg = "" Then strErrMsg = "Error executing stored procedure dbo.usp_AppealHearingPackageAnalystDocumentsUPdate"
        Err.Raise 50001, "dbo.usp_AppealHearingPackageAnalystDocumentsUpdate", strErrMsg
        End If

        strCompleteMsg = Nz(cmd.Parameters("@pCompleteMsg"), "No Message Returned")
        'MsgBox strCompleteMsg
        
        DoCmd.OpenForm "frm_AppealAnalystWorkList_Claims_CompleteConfirm" ', , , , , acDialog
        Forms("frm_AppealAnalystWorkList_Claims_CompleteConfirm").txtUserFindings = strCompleteMsg
        
    End If
    
Else
    MsgBox ("No updates were made to this claims.")
End If

Exit_Sub:
    Set CompleteUpdateADO = Nothing
    Set cmd = Nothing
    Exit Sub

Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    
    Resume Exit_Sub

End Sub

Private Sub WritePP_CV()

    Dim judgeId As Integer
    Dim CVType As String
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim strfilePathSave As String
    Dim strFilePathTemplate As String
    Dim strTemplateName As String
    Dim strCopy As String
    Dim strSQL As String
    Dim User As String
    
    Dim MyAdo As New clsADO
    Dim PPTemplateRS As ADODB.RecordSet
 
    Dim fso As Object

    User = GetUserName()

    If fso Is Nothing Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
        CVType = PP_CV_Type()
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strSQL = "Select TemplateName, TemplateFilePath, SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & CVType & "'"
        Set PPTemplateRS = MyAdo.OpenRecordSet(strSQL)
    
        'Write Cover Page 1st
        'IF: Is there a template defined?
        If PPTemplateRS.EOF = True And PPTemplateRS.BOF = True Then 'The udf_GetALJHearingPositionPaper is defined, but there is no corresponding record in the dbo.Appeal_hearing_Package_Analyst_Templates
            MsgBox ("The Position Paper cover page template has not been set up for the definition " & CVType & ". Please alert the system administrator")
        'IF: Is there a template defined?
        Else
            strfilePathSave = PPTemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & PPTemplateRS("SaveFilename") & ".docx"
            strFilePathTemplate = PPTemplateRS("TemplateFilePath") & PPTemplateRS("TemplateName")
            strTemplateName = PPTemplateRS("Templatename")
                
            'IF: Is there a PP saved for the claim?
            If FileExists(strFilePathTemplate) = True Then
                Set wrdApp = CreateObject("Word.Application")
                wrdApp.visible = False
                
                'Won't use original template in case someone else is using it. Make a copy instead.
                strCopy = strFilePathTemplate & GetUserName() & ".doc"
                
                If Not fso.FileExists(strCopy) Then
                    fso.CopyFile strFilePathTemplate, strCopy, True
                End If
                
                Set wrdDoc = wrdApp.Documents.Open(strCopy)
                
                
                With wrdDoc
                    .FormFields("fldJudgeFirstName").Result = Me.ALJJudgeFirstname
                    .FormFields("fldJudgeLastName").Result = Me.ALJudgeName
                    .FormFields("fldALJNumber").Result = Me.ALJAppealNumber
                    .FormFields("fldFieldOffice").Result = Nz(Me.txtFieldOffice, "")
                    .FormFields("fldJudgeLastName2").Result = Nz(Me.ALJudgeName, "")
                    .FormFields("fldALJNumber2").Result = Me.ALJAppealNumber
                    .FormFields("fldDoctor").Result = Doc()
                    .FormFields("fldICN").Result = Me.Icn
                    .FormFields("fldProvName").Result = Me.txtProvName
                End With
                
                If fso.FileExists(strfilePathSave) Then
                      fso.DeleteFile strfilePathSave, True
                End If
                
                'VS 1/28/2015 Overwite existing file if it exists already. Same logic applies to fax cover pages.
                wrdApp.Documents(strCopy).SaveAs2 (strfilePathSave)
                wrdApp.Application.Quit
                
                fso.DeleteFile strCopy, True
            
                    Set PPTemplateRS = Nothing
                    Set MyAdo = Nothing
                    Set cmd = Nothing
                    Set wrdDoc = Nothing
                    Set wrdApp = Nothing
                    Set fso = Nothing
            
                    WritePP_FaxCoverPage (CPJudge)
                    WritePP_FaxCoverPage (CPProv)
                    WritePP_FaxCoverPage (CPMAX)
            'IF: Is there a PP saved for the claim?
            Else
                MsgBox ("There is an error with the template.  Please contact the system administrator")
            End If
        'Is there a template defined?
        End If
          
Exit_Sub:
    
        Set PPTemplateRS = Nothing
        Set MyAdo = Nothing
        Set cmd = Nothing
        Set wrdDoc = Nothing
        Set wrdApp = Nothing
        Set fso = Nothing
    
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub

Private Function PP_CV_Type() As String

PP_CV_Type = ""
    
        judgeId = Nz(DLookup("[Id]", "ALJ_Judges_OTR", "JudgeName = '" & Trim(Replace(Me.ALJudgeName, "'", "''") & "'")), 0)
       
        If judgeId = 0 Then
            PP_CV_Type = "PP_CV"
        Else
            PP_CV_Type = "PP_CV_OTR"
        End If

End Function

Private Function Doc() As String

Doc = ""

        If Me.PP = "IRF PP" Then
            Doc = "Dr. Kenneth Adams, M.D."
        Else
            Doc = "Dr. Joby Varghese, D.O."
        End If

End Function

Private Sub WritePP_FaxCoverPage(CPType As String)
    
    Dim MyAdo As New clsADO
    Dim TemplateRS As ADODB.RecordSet
    Dim strSQL As String
   
    Dim CSName As String
    Dim CSPhone As String
    Dim CSFax As String
    Dim JudgeFax As String
    Dim JudgePhone As String
    Dim JudgeLAName As String
    
    Dim strfilePathSave As String
    Dim strFilePathTemplate As String
    Dim strTemplateName As String
    Dim strCopy As String
 
    Dim fso As Object
    Dim User As String
    
    User = GetUserName()
    CSName = DLookup("FullName", "v_ALJ_CS_PP_Contact")
    CSPhone = DLookup("Phone", "v_ALJ_CS_PP_Contact")
    CSFax = DLookup("Fax", "v_ALJ_CS_PP_Contact")
    
    If CPType = CPJudge Then
        JudgeLAName = Nz(DLookup("ClerkName", "APPEAL_XREF_ALJ_Judges", "JudgeName = '" & Trim(Replace(Me.ALJudgeName, "'", "''") & "'")), "")
        JudgePhone = Nz(DLookup("PhoneNumber", "APPEAL_XREF_ALJ_Judges", "JudgeName = '" & Trim(Replace(Me.ALJudgeName, "'", "''") & "'")), "")
        JudgeFax = Nz(DLookup("FaxNumber", "APPEAL_XREF_ALJ_Judges", "JudgeName = '" & Trim(Replace(Me.ALJudgeName, "'", "''") & "'")), "")
    End If
    
    If fso Is Nothing Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
        CVType = PP_CV_Type()
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        strSQL = "Select TemplateName, TemplateFilePath, SaveFilePath, SaveFilename from Appeal_hearing_Package_Analyst_Templates where Active = 1 and TemplateDescription = '" & CPType & "'"
        Set TemplateRS = MyAdo.OpenRecordSet(strSQL)
    
        'Write Cover Page 1st
        'IF: Is there a template defined?
        If TemplateRS.EOF = True And TemplateRS.BOF = True Then 'The udf_GetALJHearingPositionPaper is defined, but there is no corresponding record in the dbo.Appeal_hearing_Package_Analyst_Templates
            MsgBox ("The Position Paper cover page template has not been set up for the definition " & CVType & ". Please alert the system administrator")
        'IF: Is there a template defined?
        Else
            strfilePathSave = TemplateRS("SaveFilePath") & Me.Beneficiary & "_" & Me.BeneficiaryFirstInitial & "_" & Me.ALJAppealNumber & "_" & TemplateRS("SaveFilename") & ".docx"
            strFilePathTemplate = TemplateRS("TemplateFilePath") & TemplateRS("TemplateName")
            strTemplateName = TemplateRS("Templatename")
                
            'IF: Is there a PP saved for the claim?
            If FileExists(strFilePathTemplate) = True Then
                Set wrdApp = CreateObject("Word.Application")
                wrdApp.visible = False
                
                'Won't use original template in case someone else is using it. Make a copy instead.
                strCopy = strFilePathTemplate & GetUserName() & ".docx"
                
                If Not fso.FileExists(strCopy) Then
                    fso.CopyFile strFilePathTemplate, strCopy, True
                End If
                
                Set wrdDoc = wrdApp.Documents.Open(strCopy)
                
                
                 With wrdDoc
                
                    .FormFields("fldCSName").Result = CSName
                    .FormFields("fldCSPhone").Result = CSPhone
                    .FormFields("fldCSFax").Result = CSFax
                    .FormFields("fldALJNumber").Result = Me.ALJAppealNumber
                    .FormFields("fldJudgeLastName").Result = Nz(Me.ALJudgeName, "")
                    .FormFields("fldALJNumber").Result = Me.ALJAppealNumber
                'End With
                
                If (CPType = CPJudge) Then
                       .FormFields("fldJudgeFirstName").Result = Me.ALJJudgeFirstname
                       .FormFields("fldJudgeLastName2").Result = Me.ALJudgeName
                       .FormFields("fldJudgeLAName").Result = JudgeLAName
                       .FormFields("fldJudgePhone").Result = JudgePhone
                       .FormFields("fldJudgeFax").Result = JudgeFax
                End If
                     
                If (CPType = CPProv) Then
                       .FormFields("fldProvName").Result = Me.txtProvName
                       .FormFields("fldProvFax").Result = Me.txtProvFax
                       .FormFields("fldProvPhone").Result = Me.txtProvPhone
                End If
                
                End With
                
                If fso.FileExists(strfilePathSave) Then
                      fso.DeleteFile strfilePathSave, True
                End If
                
                wrdApp.Documents(strCopy).SaveAs2 (strfilePathSave)
                wrdApp.Application.Quit
                
                fso.DeleteFile strCopy, True
            
            'IF: Is there a PP saved for the claim?
            Else
                MsgBox ("There is an error with the template.  Please contact the system administrator")
            End If
        'Is there a template defined?
        End If
          
Exit_Sub:
    
        Set TemplateRS = Nothing
        Set MyAdo = Nothing
        Set cmd = Nothing
        Set wrdDoc = Nothing
        Set wrdApp = Nothing
        Set fso = Nothing
    
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub

End Sub
