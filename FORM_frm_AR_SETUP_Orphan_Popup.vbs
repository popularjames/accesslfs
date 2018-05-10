Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AR_SETUP_Orphan_Popup
' Author:      Barbara Dyroff
' Create Date: 2014-12-16
' Description:
'   Prompt the user to update AR orphans (AR/AP that have not been matched to an Audit Claim).  Provide the user the option to enter research information
' associated with the AR/AP.  Provide the option to assign a CnlyClaimNum to an AR/AP orphan.  If a CnlyClaimNum is applied, the AR will be Set Up if
' it is an initial outcome AR/AP (not collection additional).
'
' Input:
'  CurrentCnlyClaimARID            Selected AR/AP to work.
'  AR/AP and Research Info         v_AR_SETUP_Orphan_Update_Info provides the current AR/AP and research info.
'
' Output:
'   CMS_AUDITORS_CLAIMS.dbo.AR_SETUP_ORPHAN_Research_Raw    Multiple records for the same CnlyClaimARID can be entered in the raw table for processing.
'                                                           If there was a problem processing the record to the consolidated Orphan Research table, solve the problem and reprocess from the raw.
'                                                           (The Orphan Research table receives input from multiple sources.  If a conflict is detected, the update may not be allowed and manual intervention is needed.)
'
'   CMS_AUDITORS_CLAIMS.dbo.AR_SETUP_ORPHAN_Research        Records the user research for AR update and processing.
'   CMS_AUDITORS_CLAIMS.dbo.AR_SETUP_Stage_Process          Que the AR/AP for processing to set up the AR etc.  The AR/AP is not processed if from Collection additional.
'   CMS_AUDITORS_CLAIMS.dbo.AR_SETUP_Hdr                    The AR Setup is stored in the AR Setup Hdr table.
'   CMS_AUDITORS_CLAIMS.dbo.AUDITCLM_Hdr                    The Audit Claim may be updated for the AR/AP.
'
' Modification History:
'
'
' =============================================

Const CstrFrmAppID As String = "ARSetupM"
Private Const strFILETYPECD_INSUSER As String = "INSUSER"

Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Public mstrCurrentCnlyClaimARID As String  'Current AR/AP selected.

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get CurrentCnlyClaimARID() As String
    CurrentCnlyClaimARID = mstrCurrentCnlyClaimARID
End Property

Public Property Let CurrentCnlyClaimARID(strCurrentCnlyClaimARID As String)
    mstrCurrentCnlyClaimARID = strCurrentCnlyClaimARID
End Property

Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub

' Call the stored procedures to update the AR Orphan research and process the AR.
Private Sub cmdSave_Click()
    On Error GoTo Err_cmdSave_Click
    
    Dim strMsg As String
    Dim strProcName As String
    Dim strDocGUID As String
    Dim blnProcSuccessInd As Boolean
    Dim intARToPrcsNum As Integer
    
    strProcName = ClassName & ".cmdSave_Click"
    intARToPrcsNum = 0
    
    'Insert the raw research record.
    DoCmd.Hourglass True
    DoCmd.Echo True, "Updating ..."
    InsertARSetupOrphanResearchRaw strDocGUID
    
    If strDocGUID = "" Then
        strMsg = "Raw Orphan Insert identifier not returned from the procedure.  Contact application support.  "
        GoTo Err_cmdSave_Click
    End If
    
    ' Add/Update the orphan research to the consolidated Orphan Research table.
    DoCmd.Echo True, "Recording Research Info ..."
    blnProcSuccessInd = False 'Indicates if you can proceed with the next step.
    UpdateARSetupOrphanResearch strDocGUID, blnProcSuccessInd, intARToPrcsNum
    
    If blnProcSuccessInd = False Then
        strMsg = "Problem updating the consolidated AR Orphan Research.  Please contact application support for assistance.  "
        GoTo Err_cmdSave_Click
    End If

    ' Update AR if a claim was matched and process the AR Setup, if needed.
    If intARToPrcsNum > 0 Then
        DoCmd.Echo True, "Updating AR/AP ..."
        blnProcSuccessInd = False 'Indicates if you can proceed with the next step.
        UpdateARSetupOrphan blnProcSuccessInd
        
        If blnProcSuccessInd = False Then
            strMsg = "Problem processing the AR orphan update.  Please contact application support for assistance.  "
            GoTo Err_cmdSave_Click
        End If
    End If

    DoCmd.Close acForm, Me.Name
   
Exit_cmdSave_Click:
    DoCmd.Echo True, "Ready..."
    DoCmd.Hourglass False
    Exit Sub

Err_cmdSave_Click:
    MsgBox strMsg + " " + Nz(Err.Description, ""), vbOKOnly + vbCritical, "Error - " & strProcName
    ReportError Err, strProcName
    GoTo Exit_cmdSave_Click
End Sub


Private Sub Form_Load()
On Error GoTo Err_Form_Load
    Dim strProcName As String
    Dim iAppPermission As Integer

    'Init
    strProcName = ClassName & ".Form_Load"
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)

    If iAppPermission = 0 Then
        Exit Sub
    End If

    DoCmd.Echo True, "Refreshing..."

    RefreshData
    
Exit_Form_Load:
    DoCmd.Echo True, "Ready..."
    Exit Sub
Err_Form_Load:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & strProcName
    ReportError Err, strProcName
    Resume Exit_Form_Load
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & vbCrLf & ErrSource
End Sub


' Refresh the data for the window.
Public Sub RefreshData()
On Error GoTo Err_RefreshData
    Dim strProcName As String
    Dim objAdo As clsADO
    Dim strSQL As String

    strProcName = ClassName & ".RefreshData"

    If (Me.CurrentCnlyClaimARID <> "") Then

        strSQL = "SELECT * FROM v_AR_SETUP_Orphan_Update_Display WHERE CnlyClaimARID = " & Me.CurrentCnlyClaimARID
    
        'Creating a new instance of ADO-class variable
        Set objAdo = New clsADO

        Set Me.RecordSet = Nothing
        
        DoCmd.Hourglass True
        DoCmd.Echo True, "Searching ..."

        ' Select the orphan.
        With objAdo
            .ConnectionString = GetConnectString("v_AR_SETUP_Orphan_Update_Display")   'Making a Connection call to SQL database?
            .SQLTextType = sqltext
            .sqlString = strSQL                                                        'Setting the ADO-class sqlstring to the specified SQL query statement
            Set Me.RecordSet = .OpenRecordSet()
        End With
       
        If Me.RecordSet.recordCount <> 1 Then
            MsgBox "There was a problem retrieving the AR/AP [" & Me.CurrentCnlyClaimARID & "].  ", vbOKOnly + vbCritical, "Error - " & ClassName & ".RefreshData"
            GoTo Err_RefreshData
        End If
        
        'If the orphan research has been processed, change to display only.
        If Me.intProcessFlag = 1 Then
            SetProcessedDisplayOnly
        End If
        
    End If ' End if Refresh after ID identified.

Exit_RefreshData:
    DoCmd.Echo True, "Ready..."
    DoCmd.Hourglass False
    Exit Sub
Err_RefreshData:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".RefreshData"
    ReportError Err, strProcName
    Resume Exit_RefreshData
End Sub


'Insert the orphan research info in the raw table.  Return the GUID assigned on the insert.
Private Sub InsertARSetupOrphanResearchRaw(ByRef rstrDocGUID As String)
    On Error GoTo Err_InsertARSetupOrphanResearchRaw

    Dim strProcName As String
    Dim objCmd As New ADODB.Command
    Dim intResult As Integer
    Dim strMsg As String
    Dim strUserID As String

    strProcName = ClassName & ".InsertARSetupOrphanResearchRaw"
    
    intResult = -1

    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    
    ' Insert Orphan Research to the raw table.
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_AR_SETUP_ORPHAN_Research_Raw_Insert"
    myCode_ADO.SQLTextType = StoredProc

    strUserID = Identity.UserName()

    objCmd.Parameters.Append _
        objCmd.CreateParameter("RC", adInteger, adParamReturnValue)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFileTypeCd", adVarChar, adParamInput, _
            20, strFILETYPECD_INSUSER)
     
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFileTypeVersCd", adVarChar, adParamInput, _
            20, Null)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvCnlyClaimARID", adVarChar, adParamInput, _
            30, Me.txtCnlyClaimARID)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvResearchUserID", adVarChar, adParamInput, _
            50, strUserID)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFoundCnlyClaimNum", adVarChar, adParamInput, _
            30, Me.txtFoundCnlyClaimNum)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFoundCnlyICN", adVarChar, adParamInput, _
            30, Me.txtFoundCnlyICN)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFoundSysICN", adVarChar, adParamInput, _
            30, Me.txtFoundSysICN)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvStateCd", adVarChar, adParamInput, _
            2, Me.cboStateCd)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvCommentTxt", adVarChar, adParamInput, _
            1000, Me.txtCommentTxt)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvDocGUID", adVarChar, adParamOutput, 120)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvMsg", adVarChar, adParamOutput, 1000)

    ' Execute the stored procedure.
    intResult = myCode_ADO.Execute(objCmd.Parameters)

    ' Get the return DocGUID assigned and the general return message.
    strMsg = Nz(objCmd.Parameters("pchvMsg").Value, "")
    rstrDocGUID = Nz(objCmd.Parameters("pchvDocGUID").Value, "")
    
    'Check that the Stored Procedure Completed Successfully.
    If objCmd("RC") <> 0 Then
        GoTo Err_InsertARSetupOrphanResearchRaw
    End If
    
    ' Check that the ADOCls method completed successfully and the record was inserted.
    If intResult <> 1 Then
        GoTo Err_InsertARSetupOrphanResearchRaw
    End If

Exit_InsertARSetupOrphanResearchRaw:
    Set myCode_ADO = Nothing
    Set objCmd = Nothing
    Exit Sub
Err_InsertARSetupOrphanResearchRaw:
    MsgBox strMsg + " " + Nz(Err.Description, ""), vbOKOnly + vbCritical, "Error - " & strProcName
    ReportError Err, strProcName
    GoTo Exit_InsertARSetupOrphanResearchRaw
End Sub


'Update the consolidated AR Orphan research info for the given raw research record.  Return an indicator if the procedure
'completed successfully.  Also, return the total number of records ready for AR update (should be at most one from this user interface).
Private Sub UpdateARSetupOrphanResearch(ByVal vstrDocGUID As String, ByRef rblnProcSuccessInd As Boolean, ByRef rintARToPrcsNum As Integer)
    On Error GoTo Err_UpdateARSetupOrphanResearch

    Dim strProcName As String
    Dim objCmd As New ADODB.Command
    Dim intResult As Integer
    Dim strMsg As String
    Dim intSQLReturn As Integer

    strProcName = ClassName & ".UpdateARSetupOrphanResearch"
    rblnProcSuccessInd = False
    intResult = -1
    rintARToPrcsNum = 0
    
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    Set objCmd = New ADODB.Command
    
    ' Update the consolidated research.
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_AR_SETUP_ORPHAN_Manual_Match"
    myCode_ADO.SQLTextType = StoredProc

    objCmd.Parameters.Append _
        objCmd.CreateParameter("RC", adInteger, adParamReturnValue)
    
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvMsg", adVarChar, adParamOutput, 1000)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvDocGUID", adVarChar, adParamInput, _
            1000, vstrDocGUID)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFileTypeCd", adVarChar, adParamInput, _
            20, strFILETYPECD_INSUSER)
            
    ' Execute the stored procedure.
    intResult = myCode_ADO.Execute(objCmd.Parameters)
    
    ' Get Return Code any return message.
    strMsg = Nz(objCmd.Parameters("pchvMsg").Value, "")
    intSQLReturn = Nz(objCmd("RC").Value, -1)  'Returns the number inserted/updated and ready for processing.
    
    'Check that the Stored Procedure Completed Successfully.
    If intSQLReturn < 0 Then
        GoTo Err_UpdateARSetupOrphanResearch
    Else
        rblnProcSuccessInd = True 'If the procedure completed successfully, the calling procedure should be able to proceed even if something unexpected happened with the common class.
        rintARToPrcsNum = intSQLReturn 'Return the number ready for AR update.
    End If
    
    ' Check that the ADOCls method completed successfully.
    If intResult < 0 Then
        GoTo Err_UpdateARSetupOrphanResearch
    End If

Exit_UpdateARSetupOrphanResearch:
    Set myCode_ADO = Nothing
    Set objCmd = Nothing
    Exit Sub
Err_UpdateARSetupOrphanResearch:
    MsgBox strMsg + " " + Nz(Err.Description, ""), vbOKOnly + vbCritical, "Error - " & strProcName
    ReportError Err, strProcName
    GoTo Exit_UpdateARSetupOrphanResearch
End Sub

'Update and process AR Setup for the current selected AR (CnlyClaimARID).  Return the SQL Procedure Return code.
Private Sub UpdateARSetupOrphan(ByRef rblnProcSuccessInd As Boolean)
    On Error GoTo Err_UpdateARSetupOrphan

    Dim strProcName As String
    Dim objCmd As New ADODB.Command
    Dim intResult As Integer
    Dim strMsg As String

    strProcName = ClassName & ".UpdateARSetupOrphan"
    rblnProcSuccessInd = False
    intResult = -1
    
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    Set objCmd = New ADODB.Command
    
    ' Update the consolidated research.
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    myCode_ADO.sqlString = "usp_AR_SETUP_ORPHAN_Update"
    myCode_ADO.SQLTextType = StoredProc

    objCmd.Parameters.Append _
        objCmd.CreateParameter("RC", adInteger, adParamReturnValue)
    
    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvMsg", adVarChar, adParamOutput, 1000)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvCnlyClaimARID", adVarChar, adParamInput, _
            1000, txtCnlyClaimARID)

    objCmd.Parameters.Append _
        objCmd.CreateParameter("pchvFileTypeCd", adVarChar, adParamInput, _
            20, strFILETYPECD_INSUSER)

    ' Execute the stored procedure.
    intResult = myCode_ADO.Execute(objCmd.Parameters)

    ' Get any return message.
    strMsg = Nz(objCmd.Parameters("pchvMsg").Value, "")
    
    'Check that the Stored Procedure Completed Successfully and processed.  The procedure returns the count processed successfully.
    If Nz(objCmd("RC").Value, -1) <= 0 Then
        GoTo Err_UpdateARSetupOrphan
    Else
        rblnProcSuccessInd = True 'If the procedure completed successfully, the calling procedure should be able to proceed even if something unexpected happened with the common class.
    End If

    ' Check that the ADOCls method completed successfully.
    If intResult < 0 Then
        GoTo Err_UpdateARSetupOrphan
    End If
    
Exit_UpdateARSetupOrphan:
    Set myCode_ADO = Nothing
    Set objCmd = Nothing
    Exit Sub
Err_UpdateARSetupOrphan:
    MsgBox strMsg + " " + Nz(Err.Description, ""), vbOKOnly + vbCritical, "Error - " & strProcName
    ReportError Err, strProcName
    GoTo Exit_UpdateARSetupOrphan
End Sub

'Change to Display Only if the orphan research has already been processed.
Private Sub SetProcessedDisplayOnly()

    Dim strProcName As String

    On Error GoTo Err_SetProcessedDisplayOnly
    
    strProcName = ClassName & ".SetProcessedDisplayOnly"
    
    Me.txtFoundCnlyClaimNum.BackColor = "12632256" 'Grey.
    Me.txtFoundCnlyClaimNum.Locked = True
    
    Me.txtFoundCnlyClaimNum.BackColor = "12632256" 'Grey.
    Me.txtFoundCnlyClaimNum.Locked = True

    Me.txtFoundCnlyICN.BackColor = "12632256" 'Grey.
    Me.txtFoundCnlyICN.Locked = True

    Me.txtFoundSysICN.BackColor = "12632256" 'Grey.
    Me.txtFoundSysICN.Locked = True

    Me.cboStateCd.BackColor = "12632256" 'Grey.
    Me.cboStateCd.Locked = True

    Me.txtCommentTxt.BackColor = "12632256" 'Grey.
    Me.txtCommentTxt.Locked = True

Exit_SetProcessedDisplayOnly:
    Exit Sub
    
Err_SetProcessedDisplayOnly:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    
    'If there is an error, Close the Form.
    DoCmd.Close
    
End Sub
