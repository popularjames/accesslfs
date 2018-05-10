Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit ' 20130515 KD: How does anything work around here? How do people get work done!?!?!


Private Sub FormSource(sClient As String)
Dim strRecSource As String
            
    Select Case sClient
         Case "RECON"
    
             'VS 3/19/2015 Join on DocID and CnlyClaimNum
             strRecSource = "SELECT DV.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                          " INNER JOIN Queue_RECON_Review_Results AS DV" & _
                          " ON FWQ.CnlyClaimNum = DV.CnlyClaimNum and FWQ.DocID = DV.DocID" & _
                          " WHERE FWQ.Client_ext_Ref_ID IN ('1','4','5','6')" & _
                          " Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
                          
         Case "CUSTSERV"
             'MG 9/13/2013 load all converted documents into fax_queue table
             Call getConvertedDocuments
         
             'MG 6/28/2013 CHANGED TO ALLOW DOCUMENTS SHOWING INSTANCE ID
             strRecSource = " SELECT * FROM v_FAX_Work_Queue_CustServ FWQ" & _
                           " WHERE FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
                            
         Case "INC_MR"
         
              strRecSource = "SELECT Hdr.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                          " INNER JOIN AUDITCLM_Hdr AS Hdr" & _
                          " ON FWQ.CnlyClaimNum = Hdr.CnlyClaimNum" & _
                          " WHERE FWQ.Client_ext_Ref_ID = ""3""" & _
                          " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
         Case "ADHOC"
    
             strRecSource = "SELECT Hdr.ICN, FWQ.* from FAX_Work_Queue AS FWQ" & _
                          " INNER JOIN AUDITCLM_Hdr AS Hdr" & _
                          " ON FWQ.CnlyClaimNum = Hdr.CnlyClaimNum" & _
                          " WHERE FWQ.Client_ext_Ref_ID = ""1000""" & _
                          " AND FWQ.UpdateUser Like '" & gbl_sysUser & "' Order by IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate)desc"
                          
                          
         Call DeleteMRToRefTable
    End Select
          
    Me.RecordSource = strRecSource
    Me.Refresh

End Sub

Private Sub getConvertedDocuments()

Dim cmd As New ADODB.Command

Dim MyCodeAdo As New clsADO
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim myDocID As String
    
On Error GoTo ErrHandler

    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_FAX_ConverterQueueJob"
    cmd.Parameters.Refresh
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@ErrMsg"), "")

    If spReturnVal <> 0 Then
        GoTo ErrHandler
    End If

Cleanup:
    Set cmd = Nothing
    Set MyCodeAdo = Nothing
    
Exit Sub

ErrHandler:
    If ErrMsg = "" Then
        ErrMsg = Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source
    End If
    
            MsgBox "Error: " & ErrMsg, vbOKOnly + vbCritical, "Error Processing Request"
     Me.Refresh
    GoTo Cleanup

End Sub

Private Sub cmdDelFromQueue_Click()


Dim cmd As New ADODB.Command

Dim MyCodeAdo As New clsADO
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim myDocID As String
    
On Error GoTo ErrHandler


    If MsgBox("You are about to delete claim number '" & Me.Icn & " ' from the fax queue. Would you like to continue?", vbYesNo + vbQuestion, "Remove From Queue") = vbNo Then
        Exit Sub
    End If
    
    myDocID = Me.DocID
     
    '**********************************************************************************************************************************************
    'Delete from the FAX_Detail, FAX_Header, FAX_Review_Worktable and FAX_Work_Queue tables
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_FAX_DelFromQueue"
    cmd.Parameters.Refresh
    cmd.Parameters("@DocID") = myDocID
    cmd.Execute
    spReturnVal = cmd.Parameters("@Return_Value")
    ErrMsg = Nz(cmd.Parameters("@ErrMsg"), "")

    If spReturnVal <> 0 Then
        GoTo ErrHandler
    End If
'*********************************************************************************************************************************************
 
    Call FormSource(Me.OpenArgs)
 
Cleanup:
    Set cmd = Nothing
    Set MyCodeAdo = Nothing
    
Exit Sub

ErrHandler:
    If ErrMsg = "" Then
        ErrMsg = Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source
    End If
    
            MsgBox "Error: " & ErrMsg, vbOKOnly + vbCritical, "Error Processing Request"
     Me.Refresh
    GoTo Cleanup

End Sub

Private Sub cmdRefresh_Click()

Dim sOptFilter As String

    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    sOptFilter = Me.OptFilter.Value
    Call FormSource(Me.OpenArgs)
    Me.OptFilter.Value = 1
    Me.OptFilter.Value = sOptFilter
    FilterSelection (sOptFilter)

End Sub

Function CheckFormRecord()

    CheckFormRecord = Me.RecordSet.recordCount

End Function



Private Sub cmdSendFax_Click()

Dim sClient As String
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    
    Me.Refresh
    
    If MsgBox("You are about to fax all selected documents from your queue. Would you like to continue?", vbYesNo + vbQuestion, "Send Fax") = vbNo Then
        Exit Sub
    End If

    Select Case Me.OpenArgs
        Case "RECON"
            sClient = "1"
        Case "CUSTSERV"
            sClient = "2"
        Case "INC_MR"
            sClient = "3"
        Case "ADHOC"
            sClient = "1000"
            
    End Select

    Me.Refresh
    Call faxDocuments(sClient, "All", Me.OpenArgs)
    
    'Me.Refresh
    Call FormSource(Me.OpenArgs)
End Sub

Private Sub cmdSendFaxSingle_Click()

Dim FaxIcn As String
Dim sClient As String
Dim sDocID As String
    
    If CheckFormRecord = 0 Then
        Exit Sub
    End If
    
    FaxIcn = Me.txtFaxICN
    sDocID = Me.DocID
    
    If MsgBox("The document with claim number '" & FaxIcn & "' will be sent. Please ensure that the document is checked as queued." & vbCrLf & vbCrLf & "Would you like to continue?", vbYesNo + vbQuestion, "Send Fax") = vbNo Then
        Exit Sub
    End If

    Select Case Me.OpenArgs
        Case "RECON"
            sClient = "1"
        Case "CUSTSERV"
            sClient = "2"
        Case "INC_MR"
            sClient = "3"
        Case "ADHOC"
            sClient = "1000"
    End Select

    Me.Refresh
    Call faxDocuments(sClient, FaxIcn, Me.OpenArgs, sDocID)  'Customer Sevice
    'Me.Refresh
    Call FormSource(Me.OpenArgs)

End Sub

Private Sub cmdViewImage_Click()

    If CheckFormRecord = 0 Then
        Exit Sub
    End If

Dim fso As New FileSystemObject
Dim strFileLoction As String

Dim strFileName As String
 
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFileLoction = Me.DocImage
 
    If Not fso.FileExists(strFileLoction) Then
            MsgBox "The File you are looking for was renamed or moved. Check the file name and try again", vbCritical, "File Does Not Exists"
            GoTo Cleanup
    End If
    
    strFileName = Me.RecordSet("DocImage")
    SetFileReadOnly (strFileName)
    If UCase(Right(strFileName, 3)) = "TIF" Then
        If UCase(left(GetPCName(), 9)) = "TS-FLD-03" Then
            Shell "explorer.exe " & strFileName, vbNormalFocus
        Else
            Shell "C:\Program Files (x86)\IrfanView\i_view32.exe " & strFileName, vbNormalFocus
        End If
    Else
        Shell "explorer.exe " & strFileName, vbNormalFocus
    End If

Cleanup:

    Set fso = Nothing

End Sub



Private Sub DocImage_DblClick(Cancel As Integer)
    
    DoCmd.OpenForm "frm_AUDITCLM_References_Grid_View"
    Forms!frm_AUDITCLM_References_Grid_View.Controls("cmdAttach").visible = False
    Forms!frm_AUDITCLM_References_Grid_View.Controls("btn_comment").visible = False
    
    Forms!frm_AUDITCLM_References_Grid_View.RecordSource = "SELECT * FROM v_AUDITCLM_References WHERE cnlyClaimNum = '" & Me.CnlyClaimNum & "'"
    Forms!frm_AUDITCLM_References_Grid_View.Requery

End Sub


Private Sub FaxInQueue_AfterUpdate()
    Me.Requery

End Sub



Private Sub Form_Load()

    Select Case Me.OpenArgs
    Case "RECON"
        Me.Caption = "Fax Status/Queue (Reconsideration)"
    Case "CUSTSERV"
        Me.Caption = "Fax Status/Queue (Customer Service)"
    Case "INC_MR"
        Me.Caption = "Fax Status/Queue (Incomplete MR)"
    Case "ADHOC"
        Me.Caption = "Fax Status/Queue (Adhoc Projects)"
        
    End Select

    Call FormSource(Me.OpenArgs)

End Sub


Private Sub FilterSelection(strFilter As String)

'Me.Filter = ""
'Me.FilterOn = False

    Select Case strFilter

    Case 2
        DoCmd.ApplyFilter , "FaxInQueue = True"
    
    Case 3
        DoCmd.ApplyFilter , "Status = 'Sent'"
        
    Case 4
        DoCmd.ApplyFilter , "Status Like 'Fail*' OR Status = 'Cancelled'"
    
    Case 5
        DoCmd.ApplyFilter , "Status = 'In Progress'"
    
    Case Else
        Me.filter = ""
        Me.FilterOn = False
        
    End Select
End Sub



Private Sub OptFilter_AfterUpdate()

    FilterSelection (Me.OptFilter.Value)

End Sub
