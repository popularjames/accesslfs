Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim UserID As String

Private Sub cmdExit_Click()
    DoCmd.Close
End Sub

Private Sub cmdImportExcelFile_Click()
    
    Dim objxlsWbk As Excel.Workbook
    Dim objApp As New Excel.Application
    
    'Set objApp = New Excel.Application
    
    Dim dlg As clsDialogs
    Set dlg = New clsDialogs
    'Dim sFilePath As String
        
    Dim xlsRs As New ADODB.RecordSet
    Dim xlsConn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim xlsSheetName As Variant
    Dim i As Integer
        
    Me.lstAllWorksheets.RowSource = ""
    
    On Error GoTo ErrHandler
        
    With dlg
        
            sFilePath = .OpenPath("C:\", xlsf, , "Pick a Excel Workbook to load!")
               If sFilePath = "" Then
                Exit Sub
           End If
            
        End With
    
    'MsgBox sFilePath
    Me.TxtFileName = sFilePath
    
    Set objxlsWbk = objApp.Workbooks.Open(sFilePath, , True)
     
    For i = 1 To objxlsWbk.Sheets.Count
    '    Debug.Print
        Me.lstAllWorksheets.AddItem (objxlsWbk.Sheets.Item(i).Name)
    Next
    
    
Cleanup:
    
    objxlsWbk.Close False
    Set objxlsWbk = Nothing
    objApp.Quit
    Set objApp = Nothing
    
    Exit Sub
    
ErrHandler:
                MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, msgboxtitle
        
    GoTo Cleanup
      
End Sub


Private Sub cmdProcessBatch_Click()

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
        
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = MyCodeAdo.CurrentConnection
    cmd.commandType = adCmdStoredProc
    cmd.CommandText = "usp_PROV_MR_Extension_Claims_Batch_Process"
    cmd.Parameters.Refresh
    cmd.Execute
    
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    
    MsgBox "All claims processed."
    
End Sub

Private Sub Form_Load()

    UserID = Identity.UserName
    
End Sub

Private Sub clearBatch()
    'mg 10/18/2013 clear previous batch from spreadsheet load
    strSQL = "Delete * from PROV_MR_Extension_Batch_Temp where userID = '" & UserID & "'"
    
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings (True)

End Sub

Private Sub refreshBatch()
    'MG 10-18-2013 filter based on user ID
    Dim sqlString As String
    sqlString = " UserID = " & Chr(34) & UserID & Chr(34)
               
    'MG refresh data sheet
    frm_PROV_MR_Extension_Batch_Temp_subform.Form.filter = sqlString
    frm_PROV_MR_Extension_Batch_Temp_subform.Form.FilterOn = True
    frm_PROV_MR_Extension_Batch_Temp_subform.Form.Requery
    frm_PROV_MR_Extension_Batch_Temp_subform.Form.Refresh

    'Me.frm_PROV_MR_Extension_Batch_Temp_subform.SetFocus
End Sub


Private Sub lstAllWorksheets_Click()
    
    clearBatch 'clear data from previous run

    
    Dim xlsRs As New ADODB.RecordSet
    Dim xlsConn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim xlsSheetName As Variant
    
    Dim MyCodeAdo As New clsADO
    Dim spReturnVal As Variant
    Dim ErrMsg As String
    Dim strSQL As String
    Dim strProcCd As String
    Dim strUser As String
    
        
    On Error GoTo ErrHandler
    
    With xlsConn
       .Provider = "Microsoft.ACE.OLEDB.12.0"
       .ConnectionString = "Data Source='" & TxtFileName.Value & "';" & " Extended Properties=Excel 12.0"
       .Open
    End With
     
    xlsSheetName = Me.lstAllWorksheets.Value
    
    Set cmd.ActiveConnection = xlsConn
    cmd.commandType = adCmdText
    cmd.CommandText = "SELECT *  FROM [" & xlsSheetName & "$]"
    xlsRs.CursorLocation = adUseClient
    xlsRs.CursorType = adOpenStatic
    xlsRs.LockType = adLockReadOnly
    xlsRs.Open cmd

    While Not xlsRs.EOF
        DoCmd.SetWarnings (False)
        'MG 10/17/2013 safer to wrap column name in bracket as some words are reserved words in ms access
        'Using the doCMD sql is very tricky as it must have certain amount of spaces between certain words to work
        strSQL = "Insert into PROV_MR_Extension_Batch_Temp ([CnlyClaimNum], [Note], [DaysExtend], [UserID]) " & _
                 "Values ('" & xlsRs.Fields.Item(0) & "', '" & xlsRs.Fields.Item(1) & "', '" & xlsRs.Fields.Item(2) & "', '" & UserID & "')"
        
        'Debug.Print strSQL
        
        DoCmd.RunSQL (strSQL)
        DoCmd.SetWarnings (True)
        
        xlsRs.MoveNext
    Wend
     

    refreshBatch 'refresh data sheet
    
    Me.Refresh
    MsgBox "The worksheet was loaded successfully ", vbInformation, msgboxtitle
    
Cleanup:
    Set xlsRs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set xlsConn = Nothing
    Set MyCodeAdo = Nothing
    
    Me.Refresh
    Exit Sub
    
ErrHandler:
        If ErrMsg = "" Then
            ErrMsg = Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source
        End If
        
        If sFilePath = "" Or Err.Number = 3265 Then
                MsgBox "Unable to load file. The file you selected is invalid or missing the correct fields needed for processing. Please review the file and try again.", vbCritical, msgboxtitle
        Else
                MsgBox "Error: " & ErrMsg, vbOKOnly, msgboxtitle
        End If
    GoTo Cleanup

End Sub
