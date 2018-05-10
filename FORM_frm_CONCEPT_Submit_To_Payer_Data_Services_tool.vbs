Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



''' Last Modified: 01/04/2013
'''
''' ############################################################
''' ############################################################
''' ############################################################
'''  DESCRIPTION:
'''  =====================================
'''
'''
'''  TODO:
'''  =====================================
'''  -
'''
'''  HISTORY:
'''  =====================================
'''  - 1/04/2013 - Created
'''
''' AUTHOR
'''  =====================================
''' Kevin Dearing
'''
'''
'''
'''INSERT INTO General_Tabs
'''(TabName, FormName, RowSource, AccessForm, SearchType, SQLValue, FormValue, SQLCharacter, Launch, OrderBy)
'''
'''Values
'''('Concept 12 DS Send', 'frm_CONCEPT_Submit_To_Payer_Data_Services_Tool', NULL, 'frm_CONCEPT_Main', NULL, NULL, NULL, NULL, NULL, Null)
'''
'''
'''
'''
''' ############################################################
''' ############################################################
''' ############################################################
''' ############################################################


Private csConceptId As String
Private Const csMAIN_FOLDER As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\"
Private Const csARCHIVE_FOLDER As String = "\\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\Zip\"


Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get FormConceptID() As String
    FormConceptID = IdValue
End Property
Public Property Let FormConceptID(sConceptId As String)
    IdValue = sConceptId
End Property

' frmAppID
Public Property Get frmAppID() As String
    frmAppID = 1
End Property

Public Property Get IdValue() As String
    IdValue = csConceptId
End Property
Public Property Let IdValue(sValue As String)
Dim oSettings As clsSettings
Dim sLastTimeSynched As String
Dim bNeedToSynch As Boolean
    
    csConceptId = sValue
    
    Call Me.RefreshData

Block_Exit:
    Exit Property
End Property


Public Sub RefreshData()
Dim strProcName As String
On Error GoTo Block_Err
Dim oRs As ADODB.RecordSet
'Dim oFrmGeneric As Form_frm_GENERAL_Datasheet_ADO


    strProcName = ClassName & ".RefreshData"
    
    ' Refresh only the page selected..
'    Select Case Me.tabConceptStatusRpts
'    Case 0

'        Set oFrmGeneric = Me.sfrm_Generic.Form
        
        Set oRs = GetGenericReport()

        If Not oRs Is Nothing Then
            
            If oRs.recordCount > 0 Then
                Set Me.RecordSet = oRs
            Else
'                Set Me.Recordset = Nothing
                Set Me.RecordSet = oRs
            End If
        Else
            Set Me.RecordSet = Nothing
        End If
        
'    Case 1
'
'        Set oFrmGeneric = Me.sfrm_StatusOutOfLine.Form
'
'        Set oRs = GetOutOfLineReport()
'
'        If Not oRs Is Nothing Then
'            oFrmGeneric.InitDataADO oRs, "v_Data_Database"
'            If oRs.RecordCount > 0 Then
'                Set oFrmGeneric.Recordset = oRs
'            End If
'        End If
'
'
'    Case 2
'        Stop
'    Case Else
'        Stop
'    End Select
    
    

Block_Exit:
'    Set oFrmGeneric = Nothing
    Set oRs = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub





Private Sub cmbAuditor_AfterUpdate()
    Call RefreshData
End Sub




Private Sub cmbConcept_AfterUpdate()
    Call RefreshData
End Sub


Private Sub cmbPayer_AfterUpdate()
    Call RefreshData
End Sub




Private Sub cmbStatus_AfterUpdate()
    Call RefreshData
End Sub



Private Sub cmdClear_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdClear_Click"
    
    
    Me.cmbConcept = ""
'    Me.cmbStatus = ""
    Me.cmbAuditor = ""
    Me.cmbPayer = ""
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub cmdMarkAsSent_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sToAddress As String
Dim sMsg As String
Dim sThisUser As String
Dim oRs As ADODB.RecordSet
Dim lPayerNameId As Long

    strProcName = ClassName & ".cmdMarkAsSent_Click"
Stop    ' kev, test this!!

    
    If Me.RecordSet.recordCount > 1 Then
        If MsgBox("Are you sure you wish to prepare ALL of the records shown?", vbQuestion + vbYesNo, "Confirm processing ALL of these!") = vbNo Then
            GoTo Block_Exit
        End If
    End If


    Set oRs = Me.RecordsetClone
    If oRs.EOF And oRs.BOF Then
        MsgBox "No records found!", vbInformation, "oops!"
        GoTo Block_Exit
    End If
    oRs.MoveFirst
    
    Call StartMethod
    
    While Not oRs.EOF
        lPayerNameId = oRs("PayerNameID").Value

        sToAddress = "Kevin.Dearing@connolly.com;Tuan.Khong@connolly.com;"
        If Right(sToAddress, 1) <> ";" Then sToAddress = sToAddress & ";"
        
        ' now just do it:
        Call IT_Mark_Concept_as_Sent_via_NDM(oRs)
        
        sThisUser = GetUserName()

        oRs.MoveNext
    Wend

    
    LogMessage strProcName, "CONFIRMATION", "These concepts have just been marked as sent to the payer!", , True
    
    Call RefreshData
    
Block_Exit:
    Call FinishMethod
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , Me.FormConceptID
    GoTo Block_Exit
End Sub


Private Sub IT_Mark_Concept_as_Sent_via_NDM(oRs As ADODB.RecordSet)
On Error GoTo Block_Err
Dim strProcName As String
Dim sMsg As String
Dim oAdo As clsADO
Dim sSubmitAuditorEmail As String
Dim sConceptId As String
Dim lPayerNameId As Long
Dim sPayerName As String
Dim sAuditor As String


Const sToAddress As String = "Gautam.Malhotra@connolly.com;Tuan.Khong@connolly.com;Kevin.Dearing@connolly.com;"

    strProcName = ClassName & ".IT_Mark_Concept_as_Sent_via_NDM"
Stop
' fix this kev..

    lPayerNameId = oRs("PayerNameId").Value
    sPayerName = oRs("PayerName").Value
    sConceptId = oRs("ConceptID").Value
    sAuditor = oRs("Auditor").Value
    
'    ' Log it as sent,
    sMsg = sAuditor & "'s Concept: '" & sConceptId & "' and payer: " & sPayerName & " has been physically sent to the payer by " & _
        GetUserName() & vbCrLf & vbCrLf & "Please submit it via NDM or the appropriate means for this payer!"

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkSent"
        .Parameters.Refresh
        .Parameters("@pConceptId") = sConceptId
        .Parameters("@pPayerNameID") = lPayerNameId
        .Execute
        If Nz(.Parameters("@pErrMsg"), "") <> "" Then
            LogMessage strProcName, "ERROR", .Parameters("@pErrMsg"), , True, sConceptId
            GoTo Block_Exit
        End If
        sSubmitAuditorEmail = .Parameters("@pSubmitEmail").Value
    End With


        ' Send an email to Ken and the auditor that it was indeed sent
    SendsqlMail "[CONCEPT MGMT] Concept Sent: " & sConceptId & " : " & sPayerName, sSubmitAuditorEmail, "Kenneth.Turturro@connolly.com;" & sToAddress, "", sMsg

    '' Then, we need to generate the canned email to send to us which can then be fwd to the payer

    LogMessage strProcName, "NOTE TO USER", "Concept " & sConceptId & " has been sent to payer: " & sPayerName, , True, sConceptId

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName, , , sConceptId
    Err.Clear
    GoTo Block_Exit
End Sub




Private Sub cmdRefresh_Click()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".cmdRefresh_Click"
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdSendPayersEmail_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim dctEmails As Scripting.Dictionary
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO
Dim oCn As ADODB.Connection
Dim oEmailRs As ADODB.RecordSet

    strProcName = ClassName & ".cmdSendPayersEmail_Click"
    DoCmd.Hourglass True

    ' We have to loop over the recordset and create an email for each of the payers
    ' What we want to do is:
    ' - If the recordset isn't filtered then ask if they want to do ALL
    If Me.RecordSet.recordCount > 1 Then
        If MsgBox("Are you sure you wish to prepare ALL of the records shown?", vbQuestion + vbYesNo, "Confirm processing ALL of these!") = vbNo Then
            GoTo Block_Exit
        End If
    End If
    
    Set oRs = Me.RecordsetClone
    oRs.MoveFirst
    Set oCn = New ADODB.Connection
    oCn.CursorLocation = adUseClientBatch

        '' Populate our table with the concepts and payer combo's we're doing
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("RAC_IneligibleInvoice")
        oCn.ConnectionString = .ConnectionString
        oCn.Open
        .SQLTextType = sqltext
        .sqlString = "TRUNCATE TABLE CONCEPT_SubmitToPayerEmailWork"
        .Execute
        
        .sqlString = "SELECT * FROM CONCEPT_SubmitToPayerEmailWork WHERE 1 = 2"
        Set oEmailRs = .OpenRecordSet
        
    End With
    
    While Not oRs.EOF
        oEmailRs.AddNew
        oEmailRs("ConceptID") = oRs("ConceptID")
        oEmailRs("PayerNameID") = oRs("PayerNameID")
        oEmailRs.Update
        oRs.MoveNext
    Wend
    
    ' Re-connect the recordset and batch update...
    Set oEmailRs.ActiveConnection = oCn
    oEmailRs.UpdateBatch
    oEmailRs.Close
    oCn.Close
    
    ' Now, get the data:
    
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_GetPayerEmailDetails"
        .Parameters.Refresh
        Set oEmailRs = .ExecuteRS
        If .GotData = False Then
            LogMessage strProcName, "ERROR", "Something went wrong getting the data to email the payer." & .Parameters("@pErrMsg").Value, , True
            GoTo Block_Exit
        End If
    End With
    
Dim sLastPayer As String
Dim sThisPayer As String
Dim oOutlook As clsOutlookEmails
Dim oEmail As Object    'Outlook.MailItem
Const csSUBJECT As String = "[PAYERNAME] - Concept New Issue Submission"
Dim sSubject As String
Dim sBody As String
Dim sConceptId As String
Dim oCol As Collection

    Set oCol = New Collection

    Set oOutlook = New clsOutlookEmails
    
    '' Loop over that, each time we have a different payer then we need to start a new email
    While Not oEmailRs.EOF
        sThisPayer = Nz(oEmailRs("PayerName").Value, "")
        sSubject = Replace(csSUBJECT, "[PAYERNAME]", sThisPayer)
        
        sConceptId = Nz(oEmailRs("ConceptID").Value, "")
'Debug.Assert sThisPayer <> "CAHABA"
        If sLastPayer <> sThisPayer Then
            ' Finish the previous email
            If sLastPayer <> "" Then
                sBody = sBody & "</table>"
                oEmail.Display False
                oEmail.HTMLBody = Replace(oEmail.HTMLBody, "<div class=WordSection1>", "<div class=WordSection1>" & sBody, 1, 1, vbTextCompare)
                oEmail.Save
                oCol.Add oEmail
                oEmail.Close 0  ' olSave
                Set oEmail = Nothing
            End If
            ' start a new email
            
            Set oEmail = oOutlook.CreateNewMailItem(oEmailRs("EmailDistribution").Value, , sSubject, , , IIf(sLastPayer = "", True, False), False)
            
            
            sBody = "<p>All,</p><p>" & vbCrLf & "We have just submitted New Issue Concept documentation "
            
            If Nz(oEmailRs("NDM_Address").Value, "") <> "" Or Nz(oEmailRs("SendViaAppealNdm").Value, 0) <> 0 Then
                sBody = sBody & " through the NDM line."
            ElseIf Nz(oEmailRs("SFTP_Directory").Value, "") <> "" Then
                sBody = sBody & " via SFTP to <b>SFTP to sftp://157.154.39.60/FTP0191/Connolly Prod Dropoff/New Issue Validation Response</b>."
            ElseIf Nz(oEmailRs("SendThroughEmail").Value, 0) = 1 Then
                sBody = "<p>All,</p><p>" & vbCrLf & "Please find New Issue Concept documentation as attached."
            End If
            sBody = sBody & vbCrLf & vbCrLf & "<br /><br />" & vbCrLf
            
            sBody = sBody & "</p><table border=1 cellspacing=0 cellpadding=3><tr><th>ConceptID</th><th>Data Type</th><th>Review Type</th><th>Concept Desc.</th></tr>" & vbCrLf
            
        End If
    
        ' Need to build the HTML table of data
        sBody = sBody & "<tr><td>" & sConceptId & "</td>" & vbCrLf
        sBody = sBody & "<td>" & Nz(oEmailRs("DataType").Value, "") & "</td>" & vbCrLf
        sBody = sBody & "<td>" & Nz(oEmailRs("ReviewType").Value, "") & "</td>" & vbCrLf
        sBody = sBody & "<td>" & Nz(oEmailRs("ConceptDesc").Value, "") & "</td></tr>" & vbCrLf

        
        ' If it's sendThroughEmail = 1 then we need to
        ' add the attachments..
        If oEmailRs("SendThroughEmail").Value = 1 Then
'            Stop
Dim sAtachPath As String
            sAtachPath = csARCHIVE_FOLDER & Format(Now, "yyyymmdd") & "\" & sThisPayer & "\" & sConceptId & "\CONCEPT_" & Replace(sThisPayer, " ", "") & ".zip"
'            sAtachPath = csARCHIVE_FOLDER & "20130107" & "\" & sThisPayer & "\" & sConceptId & "\CONCEPT_" & Replace(sThisPayer, " ", "") & ".zip"
            
            oEmail.Attachments.Add (sAtachPath)
        End If

        
        sLastPayer = sThisPayer
        oEmailRs.MoveNext
    Wend
    ' and finish the last email
    If Not oEmail Is Nothing Then
        sBody = sBody & "</table>"
'Debug.Print oEmail.HTMLBody
'Dim oFso As Scripting.FileSystemObject
'Dim oTxt As Scripting.TextStream
'Set oFso = New Scripting.FileSystemObject
'Set oTxt = oFso.CreateTextFile("Y:\Data\CMS\AnalystFolders\KevinD\_Concept_Mgmt\email_html.txt", True, False)
'oTxt.Write oEmail.HTMLBody
'oTxt.Close
'Set oTxt = Nothing
'Set oFso = Nothing

        oEmail.Display False
        oEmail.HTMLBody = Replace(oEmail.HTMLBody, "<div class=WordSection1>", "<div class=WordSection1>" & sBody, 1, 1, vbTextCompare)

        
        oEmail.Save
'        oEmail.Close 0  '   OlInspectorClose.olSave
        oCol.Add oEmail
        Set oEmail = Nothing

    End If
    
    For Each oEmail In oCol
        oEmail.Display False

    Next
    oOutlook.BringToTopWindow
    
    
    ''         Call SetAsArchived(sConceptId, lPayerNameId)
Block_Exit:
    DoCmd.Hourglass False
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdStartTransfers_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".cmdStartTransfers_Click"
    
    ' What we want to do is:
    ' - If the recordset isn't filtered then ask if they want to do ALL
    If Me.RecordSet.recordCount > 1 Then
        If MsgBox("Are you sure you wish to prepare ALL of the records shown?", vbQuestion + vbYesNo, "Confirm processing ALL of these!") = vbNo Then
            GoTo Block_Exit
        End If
    End If

    Set oRs = Me.RecordsetClone
    oRs.MoveFirst
    While Not oRs.EOF
        Call SetAsTransferred(oRs("ConceptID").Value, oRs("PayerNameId").Value)
        oRs.MoveNext
    Wend


Block_Exit:
    RefreshData
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub cmdZipAndArchiveDocuments_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim sBatFilePath As String
Dim sConceptId As String
Dim sPayerName As String
Dim lPayerNameId As Long
Dim sZipSource As String
Dim iCurRecord As Integer

Dim sArchiveFldr As String

    strProcName = ClassName & ".cmdZipAndArchiveDocuments_Click"
    
    ' What we want to do is:
    ' - If the recordset isn't filtered then ask if they want to do ALL
    If Me.RecordSet.recordCount > 1 Then
        If MsgBox("Are you sure you wish to prepare ALL of the records shown?", vbQuestion + vbYesNo, "Confirm processing ALL of these!") = vbNo Then
            GoTo Block_Exit
        End If
    End If

    ' - loop over the recordset
    Set oRs = Me.Form.RecordsetClone
    oRs.MoveFirst
    
    While Not oRs.EOF
        iCurRecord = iCurRecord + 1
        '   - See if there is a Concept_Payername.bat file, in
        '       \\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\CONCEPTID\
        
        sPayerName = oRs("PayerName").Value
        sConceptId = oRs("ConceptId").Value
        lPayerNameId = oRs("PayerNameId").Value
        
        sBatFilePath = csMAIN_FOLDER & sConceptId & "\" & "Concept_" & Replace(sPayerName, " ", "") & ".bat"
        
        
        If FileExists(sBatFilePath) = False Then
            LogMessage strProcName, "ERROR", "No bat file found to zip package together!" & vbCrLf & "Concept: " & sConceptId & " And payer: " & sPayerName, sBatFilePath, True, sConceptId
            If oRs.recordCount > iCurRecord Then
                If MsgBox("Do you wish to continue to process the rest (Ok) or stop now (Cancel)?", vbQuestion + vbOKCancel, "Continue?") = vbOK Then
                    GoTo NextRow
                Else
                    GoTo Block_Exit
                End If
            End If
        End If

        '   - if there is, execute it which will produce a .zip file
        ShellWait sBatFilePath, vbMinimizedFocus
        
        Sleep 2000
        

        sZipSource = Replace(sBatFilePath, ".bat", ".zip")
        
        If FileExists(sZipSource) = False Then
            LogMessage strProcName, "ERROR", "We cannot find the resultant zip file!" & vbCrLf & sConceptId & " Payer: " & sPayerName, sZipSource, True, sConceptId
            
            If oRs.recordCount > iCurRecord Then
                If MsgBox("Do you wish to continue to process the rest (Ok) or stop now (Cancel)?", vbQuestion + vbOKCancel, "Continue?") = vbOK Then
                    GoTo NextRow
                Else
                    GoTo Block_Exit
                End If
            End If
        End If
        
        
        ' Create the destination archive folder
        '   - copy the zip file named CONCEPT_payername.zip into
        '       \\ccaintranet.com\DFS-CMS-FLD\Audits\CMS\AUDITCLM_ATTACHMENT\E-RAC\ZIP\{YYYYMMDD}\{PAYERNAME}\{CONCEPT_ID}
        
        sArchiveFldr = csARCHIVE_FOLDER & Format(Now, "yyyymmdd") & "\" & sPayerName & "\" & sConceptId & "\"
        CreateFolders sArchiveFldr
        
        If CopyFile(sZipSource, sArchiveFldr, False) = False Then
            LogMessage strProcName, "ERROR", "There was a problem copying the zip file to the archive directory for " & sConceptId & " payer: " & sPayerName & vbCrLf & "From: " & sZipSource & vbCrLf & "to: " & sArchiveFldr, sArchiveFldr, True, sConceptId
            If oRs.recordCount > iCurRecord Then
                If MsgBox("Do you wish to continue to process the rest (Ok) or stop now (Cancel)?", vbQuestion + vbOKCancel, "Continue?") = vbOK Then
                    GoTo NextRow
                Else
                    GoTo Block_Exit
                End If
            End If
        End If
    
        Call SetAsArchived(sConceptId, lPayerNameId)
NextRow:
        oRs.MoveNext
    Wend
    

    ' Next step is to zip the resultant files into payer specific zips..
    ' however, Palmetto has 2 depending on if it's part a or part b
    ' first coast is sent via email and that's going to be 1 email per
    ' concept
    ' So we'll only do this for NDM_Address <> "" OR SendViaAppealNdm = 1
    
    sArchiveFldr = csARCHIVE_FOLDER & Format(Now, "yyyymmdd")
    Call FinalZip(sArchiveFldr)
    

    
        '   - Open Windows explorer to that folder ~\ZIP\{YYYYMMDD}
    Shell "explorer.exe """ & sArchiveFldr & """", vbNormalFocus
    
    
Block_Exit:
    RefreshData
    
    MsgBox "Finished", vbOKOnly, "Complete!"
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub FinalZip(sStartFolder As String)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet
Dim oFso As Scripting.FileSystemObject
Dim oFldr As Scripting.Folder
Dim oPayerFldr As Scripting.Folder
Dim oConceptFldr As Scripting.Folder
Dim sConceptString As String
Dim sPayerName As String
Dim lPayerName As Long

    strProcName = ClassName & ".FinalZip"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_GetNewIssueFinalArchiveDetails"
        .Parameters.Refresh
    End With
    
Dim sDestZipFileName As String

    Set oFso = New Scripting.FileSystemObject
    Set oFldr = oFso.GetFolder(sStartFolder)
    For Each oPayerFldr In oFldr.SubFolders
            ' Clear the concept string for this payer
        sConceptString = ""
        
        sPayerName = oPayerFldr.Name
        lPayerName = GetPayerNameIDFromName(sPayerName)
        
        For Each oConceptFldr In oPayerFldr.SubFolders
            sConceptString = sConceptString & oConceptFldr.Name & ","
        Next
            ' Trim the final comma off
        If Right(sConceptString, 1) = "," Then sConceptString = left(sConceptString, Len(sConceptString) - 1)
        
        sDestZipFileName = Nz(oRs("SendZipFileName").Value, sPayerName & "_New_Issue.zip")
        
        oAdo.Parameters("@pPayerNameId") = lPayerName
        oAdo.Parameters("@pConceptIdString") = sConceptString
        
        Set oRs = oAdo.ExecuteRS
        If oAdo.GotData = False Then
            LogMessage strProcName, "ERROR", "An error occurred in " & oAdo.sqlString & vbCrLf & oAdo.Parameters("@pErrMsg").Value, , True, sConceptString
            GoTo Next_Payer
        End If


            ' Loop over the recordset and make up the zip command
            
        While Not oRs.EOF
            
            oRs.MoveNext
        Wend


Next_Payer:

    Next oPayerFldr
    
    
Block_Exit:
    RefreshData
    
    MsgBox "Finished", vbOKOnly, "Complete!"
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub SetAsArchived(sConceptId As String, lPayerNameId As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetAsArchived"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkAsArchived"
        .Parameters.Refresh
        .Parameters("@pConceptID") = sConceptId
        .Parameters("@pPayerNameId") = lPayerNameId
        .Execute
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub SetAsTransferred(sConceptId As String, lPayerNameId As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetAsArchived"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkAsTransferred"
        .Parameters.Refresh
        .Parameters("@pConceptID") = sConceptId
        .Parameters("@pPayerNameId") = lPayerNameId
        .Execute
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub SetAsPayerEmailed(sConceptId As String, lPayerNameId As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetAsArchived"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkAsPayerEmailed"
        .Parameters.Refresh
        .Parameters("@pConceptID") = sConceptId
        .Parameters("@pPayerNameId") = lPayerNameId
        .Execute
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub SetAsAuditorEmailed(sConceptId As String, lPayerNameId As Long)
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO

    strProcName = ClassName & ".SetAsAuditorEmailed"
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_MarkAsAuditorEmailed"
        .Parameters.Refresh
        .Parameters("@pConceptID") = sConceptId
        .Parameters("@pPayerNameId") = lPayerNameId
        .Execute
    End With
    
Block_Exit:
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String
Dim sUsername As String
Dim sUserProfile As String
Dim sSuperName As String


    strProcName = ClassName & ".Form_Load"
    Me.Form.RecordSource = ""
    Me.ckShowSent = 0
    
    '' Only data services / CM_Admin profile should get this
    '
    sUsername = GetUserName()
    sUserProfile = GetUserProfile()
    sSuperName = Identity.UserSupervisorId()
    
    If UCase(sSuperName) <> "DATA CENTER" Or UCase(sUserProfile) <> "CM_ADMIN" Then
        ' no permission to use this form..
        Me.txtSecurityBreachNote = "Only Data Services + CM_Admin are allowed to use this form!"
        ' make all of the other controls invisible..
        ' eh, not now..
        Call VisiblizeCtls(False)
        GoTo Block_Exit
    End If
'    VisiblizeCtls True

    Select Case UCase(sUserProfile)
    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
'        Me.cmdCreateEmail.Enabled = True
    Case Else
'        Me.cmdCreateEmail.Enabled = False
    End Select
    
    Me.AllowFilters = True
    
    ' refresh all of the combo boxes..
    Call RefreshComboBoxes
    
    
    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit


End Sub


Public Sub VisiblizeCtls(blnVisible As Boolean)
On Error Resume Next    ' Shame on you Kev!
Dim oCtl As Control
    
    Me.txtSecurityBreachNote.visible = Not blnVisible
    If blnVisible = False Then
        Me.txtSecurityBreachNote.SetFocus
    End If
    
    For Each oCtl In Me.Controls
        If oCtl.Name <> "txtSecurityBreachNote" Then
            oCtl.visible = blnVisible
        End If
    Next
End Sub

Private Sub RefreshComboBoxes()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As ADODB.RecordSet
Dim oAdo As clsADO

    strProcName = ClassName & ".RefreshComboBoxes"
    
    ''-- Concept:
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Data_Database")
        .SQLTextType = sqltext
        .sqlString = "select distinct ConceptId, ConceptDesc from CONCEPT_Hdr"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbConcept.ColumnCount = 2
            Me.cmbConcept.ColumnWidths = "1000;2880;"
            Set Me.cmbConcept.RecordSet = oRs
        End If
    
    
'        .SqlString = "SELECT ConceptStatus, StatusDescription FROM CONCEPT_XREF_Status"
'        Set oRs = .ExecuteRS
'        If .GotData Then
''            Me.cmbStatus.ColumnCount = 2
''            Me.cmbStatus.ColumnWidths = "1000;2880;"
'            Set Me.cmbStatus.Recordset = oRs
'        End If
    
    
        .sqlString = "SELECT DISTINCT Auditor FROM CONCEPT_Hdr WHERE Auditor IS NOT NULL ORDER BY Auditor"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbAuditor.ColumnCount = 1
            Me.cmbAuditor.ColumnWidths = "2880;"
            Set Me.cmbAuditor.RecordSet = oRs
        End If
    
    
    
        .sqlString = "SELECT PayerNameId, PayerName FROM XREF_Payernames WHERE PayerNameId > 999 ORDER BY PayerName"
        Set oRs = .ExecuteRS
        If .GotData Then
            Me.cmbPayer.ColumnCount = 2
            Me.cmbPayer.ColumnWidths = "1000;2880;"
            Set Me.cmbPayer.RecordSet = oRs
        End If
    
    
    
    
    
    End With
    
    
    
    

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Function GetGenericReport() As ADODB.RecordSet
On Error GoTo Block_Err
Dim strProcName As String
Dim oAdo As clsADO
Dim sSql As String
Dim oRs As ADODB.RecordSet

    strProcName = ClassName & ".GetGenericReport"
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_CODE_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_Ready_2_Send_2_Payers"
        .Parameters.Refresh
        
'        If IsSubForm(Me) Then
'            Me.cmbConcept = Nz(Me.Parent.Form.txtConceptID, "")
'        End If
        
        If Nz(Me.cmbConcept, "") <> "" Then
            .Parameters("@pConceptId") = Me.cmbConcept
        End If
        
        If Nz(Me.cmbPayer, "") <> "" Then
            .Parameters("@pPayerNameID") = Me.cmbPayer
        End If
        
        If Nz(Me.ckShowSent, 0) <> 0 Then
            .Parameters("@pShowSent") = 1
        End If
        
'        If Nz(Me.cmbStatus, "") <> "" Then
'            .Parameters("@pConceptStatus") = Me.cmbStatus
'        End If
        
        If Nz(Me.cmbAuditor, "") <> "" Then
            .Parameters("@pAuditor") = Me.cmbAuditor
        End If
     
        Set oRs = .ExecuteRS
        If .GotData = False Then
'            GoTo Block_Exit
        End If

    End With
    Debug.Print oRs.recordCount
    
    Set GetGenericReport = oRs
Block_Exit:
    Exit Function

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function






'Private Sub Form_Resize()
'Dim lWidth As Long
'Dim lHeight As Long
'
'    '' Resize stuff..
'    ' just in the details section
'    lWidth = Me.InsideWidth - 200
'    lHeight = Me.InsideHeight - 700
'
'
'    Me.sfrm_Generic.width = lWidth - 100
'    Me.sfrm_Generic.Height = lHeight - 300
''    Me.sfrm_StatusOutOfLine.width = Me.tabConceptStatusRpts.Pages(1).width - 100
''    Me.sfrm_StatusOutOfLine.Height = Me.tabConceptStatusRpts.Pages(1).Height - 300
'
'
'End Sub

Private Sub tabConceptStatusRpts_Change()
    Call RefreshData
End Sub



Private Sub cmdCreateEmail_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sConceptId As String
Dim lPayerNameId As Long
Dim oConcept As clsConcept
Dim oRs As ADODB.RecordSet
Dim sUsername As String
Dim sUserProfile As String

    strProcName = ClassName & ".cmdCreateEmail_Click"
    sUserProfile = GetUserProfile()
    
    '' Only Ken should be able to do this:
'    sUsername = GetUserName()
'    Select Case UCase(sUsername)
'    Case "KENNETH.TURTURRO" ' , "KEVIN.DEARING"
'        Stop
'    Case Else
'        MsgBox "You do not have adequate permissions to do this!"
'        GoTo Block_Exit
'    End Select
    
    Select Case UCase(sUserProfile)
    Case "CM_ADMIN", "CM_AUDIT_MANAGERS"
                
    Case Else
        MsgBox "You do not have adequate permissions to do this!"
        GoTo Block_Exit
    End Select
    
    
        ''' So, we need the concept id
        ''  payernameid
        '' and that should be about it..
    
    Set oRs = Me.RecordSet
    
    
    sConceptId = Nz(oRs("ConceptID").Value, "")
    lPayerNameId = Nz(oRs("PayerNameId").Value, 1000)
    
    Set oConcept = New clsConcept
    If oConcept.LoadFromId(sConceptId) = False Then
        LogMessage strProcName, "ERROR", "There was a problem loading the concept object!", , True, sConceptId
        GoTo Block_Exit
    End If
    
    
    Call PrepConceptSubmitEmail(oConcept, lPayerNameId)
    
    
    Call mod_Concept_Specific.MarkConceptAsSentToCms(sConceptId, lPayerNameId)
    
    ' So, now let's just assume that the email is sent and we'll mark the database as sent..
    Call Me.RefreshData

Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub




Private Sub SortForm(strFieldToSortOn As String)
On Error GoTo Block_Err
Dim strProcName As String
Static bAscending As Boolean
Dim sFilter As String
Dim oAdoRs As ADODB.RecordSet
Dim oDaoRs As DAO.RecordSet


    strProcName = ClassName & ".SortForm"
        ' flip it
    bAscending = Not bAscending
    
    sFilter = strFieldToSortOn & IIf(bAscending, " ASC", " DESC")

    
    If TypeOf Me.RecordSet Is ADODB.RecordSet Then
        Set oAdoRs = Me.RecordSet
        oAdoRs.Sort = sFilter
        Set Me.RecordSet = Nothing
        Set Me.RecordSet = oAdoRs
    Else
        Set oDaoRs = Me.RecordSet
        oDaoRs.Sort = sFilter
        Set Me.RecordSet = Nothing
        Set Me.RecordSet = oDaoRs
    End If
    
    


Block_Exit:
    Exit Sub

Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub






Private Sub Form_Open(Cancel As Integer)
    Call VisiblizeCtls(True)
End Sub

Private Sub lblAuditor_Click()
    SortForm ("Auditor")
End Sub

Private Sub lblConceptID_Click()
    SortForm ("ConceptId")
End Sub

Private Sub lblConceptLevel_Click()
    SortForm ("ConceptLevel")
End Sub

Private Sub lblConceptStatus_Click()
    SortForm ("CStatus")
End Sub

Private Sub lblDataType_Click()
    SortForm ("DataType")
End Sub

Private Sub lblDateFinalized_Click()
    SortForm ("DateFinalized")
End Sub

Private Sub lblDateSentToPayer_Click()
    SortForm ("DateSentToPayer")
End Sub

Private Sub lblDtSubmittedonNIRF_Click()
    SortForm ("DtSubmittedonNIRF")
End Sub

Private Sub lblLastUpDt_Click()
    SortForm ("LastUpDt")
End Sub

Private Sub lblLastUpUser_Click()
    SortForm ("LastUpUser")
End Sub

Private Sub lblLOB_Click()
    SortForm ("LOB")
End Sub

Private Sub lblNirfSentToCMSDt_Click()
    SortForm ("NirfSentToCMSDt")
End Sub

Private Sub lblNirfSentToCmsUser_Click()
    SortForm ("NirfSentToCmsUser")
End Sub

Private Sub lblPackageCreatedDt_Click()
    SortForm ("PackageCreatedDt")
End Sub

Private Sub lblPackageCreateUser_Click()
    SortForm ("PackageCreateUser")
End Sub

Private Sub lblPayerName_Click()
    SortForm ("PayerName")
End Sub

Private Sub lblPayerNameID_Click()
    SortForm ("PayerNameId")
End Sub

Private Sub lblQAUser_Click()
    SortForm ("QAUser")
End Sub

Private Sub lblReviewType_Click()
    SortForm ("ReviewType")
End Sub

Private Sub lblStatusDescription_Click()
    SortForm ("StatusDescr")
End Sub

Private Sub lblUserWhoSentToPayer_Click()
    SortForm ("UserWhoSentToPayer")
End Sub
