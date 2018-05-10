Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const msgboxtitle = "Load Sub Status"
Public sFilePath As String
Dim UserID As String


Public Sub ControlVisible()

'
'If Me.optSelect = 2 Then
'    Me.txtNotes.visible = True
'    Me.TxtStatusCd.visible = True
'    Me.Label32.visible = True
'    Me.Label34.visible = True
'    Me.subFrmSubStatus.Form.Controls("SubStatus").ColumnHidden = True
'    Me.subFrmSubStatus.Form.Controls("Comment").ColumnHidden = True
'    Me.subFrmSubStatus.Form.Controls("UpdateDate").ColumnHidden = True
'    Me.subFrmSubStatus.Form.Controls("Note").ColumnHidden = True
' Else
'    Me.txtNotes.visible = False
'    Me.TxtStatusCd.visible = False
'    Me.Label32.visible = False
'    Me.Label34.visible = False
'    Me.subFrmSubStatus.Form.Controls("SubStatus").ColumnHidden = False
'    Me.subFrmSubStatus.Form.Controls("Comment").ColumnHidden = False
'    Me.subFrmSubStatus.Form.Controls("UpdateDate").ColumnHidden = False
'    Me.subFrmSubStatus.Form.Controls("Note").ColumnHidden = False
' End If


End Sub

Private Sub updateCtl(StrLock As String)

Dim ctl As Control

    For Each ctl In Me.Controls
        If (ctl.ControlType <> acLabel) And (ctl.ControlType <> acRectangle) And (ctl.ControlType <> acLine) Then
        'Debug.Print ctl.Name
        ctl.Enabled = StrLock
    End If
    Next ctl
    
End Sub

Private Sub RecordFilter()

Dim strSQL As String

Select Case Me.txtCnlyClaimNumLkUp
    
    Case ""
        strSQL = "select * from MassClaimUpdate_Worktable WHERE UpdateUser = '" & UserID & "'"
    Case Else
        strSQL = "select * from MassClaimUpdate_Worktable WHERE UpdateUser = '" & UserID & "'" & " AND  cnlyClaimNum Like '" & Me.txtCnlyClaimNumLkUp & "%'"
End Select

Me.frm_CUST_MassClaimUpdate_SubForm.Form.RecordSource = strSQL
Me.frm_CUST_MassClaimUpdate_SubForm.Form.Requery
Me.Refresh

End Sub

Private Sub cmd_MassLookUp_Click()
'Begin sub
On Error GoTo ErrorHandler

    Dim dbs As Database
    Dim rs As RecordSet
    Dim qdf As QueryDef
    Dim productName As String
    Dim strSQL As String
    
    
    '2014-06-14 TK: getting cnlyclaimnum
    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As Variant
    Dim strProcCd As String
    Dim ErrMsg As String
    Dim ErrorReturned As String
    
    'tk test
    Debug.Print "CurrentProject.Connection.CommandTimeout = " & CurrentProject.Connection.CommandTimeout
    CurrentProject.Connection.CommandTimeout = 600
    Debug.Print "CurrentProject.Connection.CommandTimeout = " & CurrentProject.Connection.CommandTimeout
        
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
        'tk test
        Debug.Print "cmd.CommandTimeout = " & cmd.CommandTimeout
    cmd.CommandTimeout = 600
        Debug.Print "cmd.CommandTimeout = " & cmd.CommandTimeout
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_MassClaimUpdate_GetCnlyClaimNum"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pUserID") = UserID
        
        DoCmd.Hourglass True
        .Execute
            DoEvents
            DoEvents
            DoEvents
        
        DoCmd.Hourglass False
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
    
    End With


    '2014-06-14 TK report summary
    Set dbs = CurrentDb()
    
    productName = "ADSL"
    
    strSQL = "Select * from v_MassClaimLookUp WHERE UpdateUser = '" & UserID & "'"
    
    Set rs = dbs.OpenRecordSet(strSQL)
    
    With dbs
        Set qdf = .CreateQueryDef("tmpMassClaimLookUp", strSQL)
        DoCmd.OpenQuery "tmpMassClaimLookUp"
        .QueryDefs.Delete "tmpMassClaimLookUp"
    End With
    dbs.Close
    qdf.Close

' Release used objects. Such as ado.
CleanupAndExit:

    Exit Sub

ErrorHandler:
    With Err
        MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit
'Exit

End Sub

Private Sub cmdClear_Click()

    Me.txtCnlyClaimNumLkUp.Value = ""
    Call RecordFilter

End Sub

Private Sub cmdExportResult_Click()

'declare variables
Dim bResult As Boolean
Dim strResultFilePath As String
Dim strSQL As String

'Begin sub
On Error GoTo ErrorHandler

' main code
    bResult = False
        
    strResultFilePath = Me.TxtFileName.Value
    
    If Right(strResultFilePath, 5) = ".xlsx" Then
        strResultFilePath = Replace(strResultFilePath, ".xlsx", "_RESULT.xlsx")
    ElseIf Right(strResultFilePath, 4) = ".xls" Then
        strResultFilePath = Replace(strResultFilePath, ".xls", "_RESULT.xls")
    End If
    
    strSQL = "SELECT * FROM MassClaimUpdate_Worktable WHERE UpdateUser = '" & UserID & "'"
    bResult = ExportToExcel(strSQL, strResultFilePath)


' Release used objects. Such as ado.
CleanupAndExit:

    Exit Sub


ErrorHandler:
    With Err
    MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit
'Exit
End Sub


Private Function ExportToExcel(strSQL As String, strFilePath As String) As Boolean
    Dim dlg As clsDialogs
    Dim cie As clsImportExport

    Set cie = New clsImportExport
    Set dlg = New clsDialogs



    With dlg
    
        strFilePath = .SavePath(Identity.CurrentFolder, xlsf, strFilePath)
        strFilePath = .CleanFileName(strFilePath, CleanPath)
     
        If strFilePath <> "" Then
        
            If .FileExists(strFilePath) = True Then
            
                If MsgBox("Overwrite existing file?", vbYesNo) = vbYes Then
                    .DeleteFile strFilePath
                Else
                    GoTo exitHere
                End If
            
            End If
             
        Else
            GoTo exitHere
        End If
     
        With cie
            .ExportExcelSql strSQL, strFilePath
        End With
         
    End With
    
    ExportToExcel = True

exitHere:
    Set cie = Nothing
    Set dlg = Nothing
    Exit Function
    
HandleError:
    ExportToExcel = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function



Private Sub cmdLoadXLSFile_Click()
'Begin sub
On Error GoTo ErrorHandler

    
    'declare variables
    Dim xlsRs As New ADODB.RecordSet
    Dim xlsConn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim xlsSheetName As Variant
    
    Dim MyCodeAdo As New clsADO
    Dim spReturnVal As Variant
    Dim ErrMsg As String
    Dim strSQL As String
    Dim strProcCd As String
    
    Dim MyAdo As clsADO
    Dim rst As ADODB.RecordSet
    
    'disable button after load
    cmdLoadXLSFile.Enabled = False
        
    'TK clearing data on each load
    ClearWorkTable
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")

    Me.TxtFileName = sFilePath
    
    With xlsConn
       .Provider = "Microsoft.ACE.OLEDB.12.0"
       .ConnectionString = "Data Source='" & sFilePath & "';" & " Extended Properties=Excel 12.0"
       .Open
    End With
     
    xlsSheetName = Me.lstSheetName.Value
    If Nz(xlsSheetName, "") = "" Then
       MsgBox "No worksheet found. Please select a worksheet.", vbInformation, msgboxtitle
       GoTo CleanupAndExit
    End If

    Set cmd.ActiveConnection = xlsConn
    cmd.commandType = adCmdText
    cmd.CommandText = "SELECT *  FROM [" & xlsSheetName & "$]"
    xlsRs.CursorLocation = adUseClient
    xlsRs.CursorType = adOpenStatic
    xlsRs.LockType = adLockReadOnly
    xlsRs.Open cmd
     
  
       
    'TK loop and load excel data to SQL table
    While Not xlsRs.EOF
    DoCmd.SetWarnings (False)
    strSQL = "Insert into MassClaimUpdate_Worktable (" & _
            " CnlyClaimNum,ICN,NEW_ClaimNotes,NEW_ClaimStatus,NEW_ClaimQueue," & _
            " NEW_AdjICN,NEW_Substatus,NEW_Comments,NEW_ARcomments,NEW_AdjCan,NEW_AdjBic,UpdateUser) " & _
            " Select '" & Replace(Nz(xlsRs.Fields.Item(0), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(1), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(2), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(3), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(4), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(5), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(6), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(7), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(8), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(9), ""), "'", "") & _
            "','" & Replace(Nz(xlsRs.Fields.Item(10), ""), "'", "") & _
            "','" & Replace(Nz(UserID, ""), "'", "") & "'"
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings (True)
    
     xlsRs.MoveNext
     Wend
    '
    ''**********************************************************************************************************************************************
    ''Update the note field in the MassClaimUpdate_Worktable table
    'myCodeADO.ConnectionString = GetConnectString("v_CODE_Database")
    '
    '                Set cmd = New ADODB.Command
    '                cmd.ActiveConnection = myCodeADO.CurrentConnection
    '                cmd.commandType = adCmdStoredProc
    '                cmd.CommandText = "usp_AUDITCLM_UpdateSubStatus_Note"
    '                cmd.Parameters.Refresh
    '                cmd.Execute
    '                spReturnVal = cmd.Parameters("@Return_Value")
    '                ErrMsg = Nz(cmd.Parameters("@ErrMsg"), "")
    '
    '                If spReturnVal <> 0 Then
    '                    GoTo ErrHandler
    '                End If
    ''*********************************************************************************************************************************************
     
    Me.Refresh
    RefreshWorkTable
    
    
    'MsgBox "The worksheet was loaded successfully ", vbInformation, msgboxtitle
    
    'TK: Enable button after process data. (Re-enable upon loading new file load)
    cmdRunUpdate.Enabled = True
    cmdExportResult.Enabled = True
 
CleanupAndExit:
    Set xlsRs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set xlsConn = Nothing
    Set MyCodeAdo = Nothing
    Set MyAdo = Nothing
    Set rst = Nothing

Me.Refresh
Exit Sub


ErrorHandler:
    If ErrMsg = "" Then
        ErrMsg = Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source
    End If
    
    If sFilePath = "" Or Err.Number = 3265 Then
        MsgBox "Unable to load file. The file you selected is invalid or missing the correct fields needed for processing. Please review the file and try again.", vbCritical, msgboxtitle
    Else
        MsgBox "Error: " & ErrMsg, vbOKOnly, msgboxtitle
    End If
    Resume CleanupAndExit

End Sub

Private Sub cmdPickXlsFile_Click()

On Error GoTo ErrorHandler

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
        
    Me.lstSheetName.RowSource = ""

    'TK clearing data on each load
    ClearWorkTable
    
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
        Me.lstSheetName.AddItem (objxlsWbk.Sheets.Item(i).Name)
    Next


    'TK enable load functionality
    cmdLoadXLSFile.Enabled = True

' Release used objects. Such as ado.
CleanupAndExit:
    objxlsWbk.Close False
    Set objxlsWbk = Nothing
    objApp.Quit
    Set objApp = Nothing
    Exit Sub


ErrorHandler:
    With Err
        MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With

    Resume CleanupAndExit
End Sub

Private Sub cmdRunUpdate_Click()

    Dim MyCodeAdo As New clsADO
    Dim cmd As ADODB.Command
    Dim spReturnVal As Variant
    Dim strSQL As String
    Dim strProcCd As String
    Dim ErrMsg As String
    Dim ErrorReturned As String
            
    Dim objDB As Database
    Set objDB = CurrentDb
    objDB.QueryTimeout = 360


    'TK: Disable button after process data. (Re-enable upon loading new file load)
    cmdRunUpdate.Enabled = False
    
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")

    Set cmd = New ADODB.Command
    '2014-09-24 tk extend timeout timer
    CurrentProject.Connection.CommandTimeout = 600
    cmd.CommandTimeout = 600
    'Debug.Print "cmd.CommandTimeout = " & cmd.CommandTimeout
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_MassClaimUpdate_Worktable_Process"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pUserID") = UserID
        
        DoCmd.Hourglass True
        .Execute
            DoEvents
            DoEvents
            DoEvents
        
        DoCmd.Hourglass False
        spReturnVal = .Parameters("@Return_Value")
        ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
    
    End With


    RefreshWorkTable
    

    
    MsgBox "Processing complete. Please check results comments below.", vbInformation, msgboxtitle

' Release used objects. Such as ado.
CleanupAndExit:
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    Exit Sub


ErrorHandler:
    With Err
        MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit

'end sub
End Sub

Private Sub cmdSearch_Click()

Exit Sub

Call RecordFilter

End Sub

Private Sub RefreshWorkTable()

'Begin sub
On Error GoTo ErrorHandler


'declare variables
Dim strSource As String

    strSource = "select * from MassClaimUpdate_Worktable WHERE UpdateUser = '" & UserID & "'"
    'TK to change into view.
    'strSource = "Select * from v_MassClaimUpdate_Worktable where UpdateUser = '" & UserID & "'"
    Me.frm_CUST_MassClaimUpdate_SubForm.Form.RecordSource = strSource
        
            
    ' main code
    'TK refresh data sheet
    'Me.frm_CUST_MassClaimUpdate_SubForm.Form.FilterOn = True
    Me.frm_CUST_MassClaimUpdate_SubForm.Form.Requery
    Me.frm_CUST_MassClaimUpdate_SubForm.Form.Refresh

' Release used objects. Such as ado.
CleanupAndExit:

    Exit Sub


ErrorHandler:
    With Err
        MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With

    Resume CleanupAndExit

'end sub
End Sub


Private Sub Form_Load()
On Error GoTo ErrorHandler

    Dim strSQL As String
    Dim StrAccessRights As String
    Dim StrAccessRightsFINAL As String
    Dim StrAccessRightsUSER As String
    
    
    Dim rs As ADODB.RecordSet
    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    
    'TK one time variable defined.
    UserID = Identity.UserName
    
    'TK disable run until data load
    cmdLoadXLSFile.Enabled = False
    cmdRunUpdate.Enabled = False
    cmdExportResult.Enabled = False
    
    'TK clearing data on each load
    ClearWorkTable
    
    'TK refresh subform
    RefreshWorkTable
    
    
    Me.lstSheetName.RowSource = ""
    Me.TxtFileName = ""
    


CleanupAndExit:
   
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, msgboxtitle
    Resume CleanupAndExit

End Sub


Private Sub Form_Unload(Cancel As Integer)

'Dim strSQL As String
'
'strSQL = "Delete * from MassClaimUpdate_Worktable where UpdateUser = '" & UserID & "'"
'DoCmd.SetWarnings (False)
'DoCmd.RunSQL (strSQL)
'DoCmd.SetWarnings (True)

End Sub

Private Sub ClearWorkTable()
Dim strSQL As String
'Begin sub
On Error GoTo ErrorHandler

'declare variables
    Dim MyCodeAdo As clsADO
    Dim cmd As ADODB.Command
    
' main code
    Set MyCodeAdo = New clsADO
    MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = MyCodeAdo.CurrentConnection
        .CommandText = "usp_MassClaimUpdate_Worktable_ClearData"
        .commandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters("@pUserID") = UserID
        DoCmd.Hourglass True
        .Execute
        DoCmd.Hourglass False
        'ErrorReturned = Nz(.Parameters("@pErrMsg"), "")
    End With

'refresh data
''Me.frm_CUST_MassClaimUpdate_SubForm.Form.RecordSource = strSQL
Me.frm_CUST_MassClaimUpdate_SubForm.Form.Requery
Me.Refresh



' Release used objects. Such as ado.
CleanupAndExit:
    Set MyCodeAdo = Nothing
    Set cmd = Nothing
    Exit Sub


ErrorHandler:
    With Err
    MsgBox ("Subroutine error: " & .Number & ". " & .Description)
    End With
    Resume CleanupAndExit

'end sub
End Sub

Private Sub optSelect_AfterUpdate()

ControlVisible

End Sub

Private Sub txtCnlyClaimNumLkUp_AfterUpdate()

Call RecordFilter

End Sub
