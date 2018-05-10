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

    
If Me.optSelect = 2 Then
    Me.txtNotes.visible = True
    Me.TxtStatusCd.visible = True
    Me.Label32.visible = True
    Me.Label34.visible = True
    Me.subFrmSubStatus.Form.Controls("SubStatus").ColumnHidden = True
    Me.subFrmSubStatus.Form.Controls("Comment").ColumnHidden = True
    Me.subFrmSubStatus.Form.Controls("UpdateDate").ColumnHidden = True
    Me.subFrmSubStatus.Form.Controls("Note").ColumnHidden = True
       
 Else
    Me.txtNotes.visible = False
    Me.TxtStatusCd.visible = False
    Me.Label32.visible = False
    Me.Label34.visible = False
    Me.subFrmSubStatus.Form.Controls("SubStatus").ColumnHidden = False
    Me.subFrmSubStatus.Form.Controls("Comment").ColumnHidden = False
    Me.subFrmSubStatus.Form.Controls("UpdateDate").ColumnHidden = False
    Me.subFrmSubStatus.Form.Controls("Note").ColumnHidden = False
       
 End If
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
        strSQL = "select * from AUDITCLM_SubStatus_Tmp"
    Case Else
        strSQL = "select * from AUDITCLM_SubStatus_Tmp where cnlyClaimNum Like '" & Me.txtCnlyClaimNumLkUp & "%'"
End Select

Me.subFrmSubStatus.Form.RecordSource = strSQL
Me.Refresh

End Sub

Private Sub cmdClear_Click()

Me.txtCnlyClaimNumLkUp.Value = ""
Call RecordFilter

End Sub

Private Sub cmdLoadXLSFile_Click()
    
Dim xlsRs As New ADODB.RecordSet
Dim xlsConn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim xlsSheetName As Variant

Dim MyCodeAdo As New clsADO
Dim spReturnVal As Variant
Dim ErrMsg As String
Dim strSQL As String
Dim strProcCd As String
Dim UserID As String

    
On Error GoTo ErrHandler

UserID = Identity.UserName

Me.TxtFileName = sFilePath

With xlsConn
   .Provider = "Microsoft.ACE.OLEDB.12.0"
   .ConnectionString = "Data Source='" & sFilePath & "';" & " Extended Properties=Excel 12.0"
   .Open
End With
 
xlsSheetName = Me.lstSheetName.Value

Set cmd.ActiveConnection = xlsConn
cmd.commandType = adCmdText
cmd.CommandText = "SELECT *  FROM [" & xlsSheetName & "$]"
xlsRs.CursorLocation = adUseClient
xlsRs.CursorType = adOpenStatic
xlsRs.LockType = adLockReadOnly
xlsRs.Open cmd
 
If Me.optSelect = 1 Then
   strProcCd = "SubStat"
 Else
   strProcCd = "ClmStat"
End If

strSQL = "Delete * from AUDITCLM_SubStatus_Tmp where UpdateUser = '" & UserID & "'"

DoCmd.SetWarnings (False)
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings (True)
 
While Not xlsRs.EOF
DoCmd.SetWarnings (False)
DoCmd.RunSQL ("Insert into AUDITCLM_SubStatus_Tmp(CnlyClaimNum, SubStatus, Comment, Process, UpdateUser) Select '" & xlsRs.Fields.Item(2) & "','" & xlsRs.Fields.Item(3) & "','" & xlsRs.Fields.Item(4) & "','" & strProcCd & "','" & UserID & "'")
DoCmd.SetWarnings (True)

 xlsRs.MoveNext
 Wend
 
'**********************************************************************************************************************************************
'Update the note field in the AUDITCLM_SubStatus_Tmp table
MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                cmd.CommandText = "usp_AUDITCLM_UpdateSubStatus_Note"
                cmd.Parameters.Refresh
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")
                ErrMsg = Nz(cmd.Parameters("@ErrMsg"), "")
        
                If spReturnVal <> 0 Then
                    GoTo ErrHandler
                End If
'*********************************************************************************************************************************************
 
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

Private Sub cmdPickXlsFile_Click()

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
    Me.lstSheetName.AddItem (objxlsWbk.Sheets.Item(i).Name)
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

Private Sub cmdRunUpdate_Click()

Dim MyCodeAdo As New clsADO
Dim cmd As ADODB.Command
Dim spReturnVal As Variant
Dim strSQL As String
Dim strProcCd As String
Dim ErrMsg As String

MyCodeAdo.ConnectionString = GetConnectString("v_CODE_Database")
    
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = MyCodeAdo.CurrentConnection
                cmd.commandType = adCmdStoredProc
                    
If Me.optSelect = 1 Then
                cmd.CommandText = "usp_AUDITCLM_UpdateSubStatus"
                cmd.Parameters.Refresh
                cmd.Parameters("@StrRunID") = "SubStat"
                cmd.Parameters("@UserFile") = Me.TxtFileName
Else
                cmd.CommandText = "usp_AUDITCLM_UpdateClaimStatus"
                cmd.Parameters.Refresh
                cmd.Parameters("@StrRunID") = "ClmStat"
                cmd.Parameters("@strNotes") = Me.txtNotes
                cmd.Parameters("@strStatusCd") = Me.TxtStatusCd
                cmd.Parameters("@UserFile") = Me.TxtFileName
                                             
End If
               
                cmd.Execute
                spReturnVal = cmd.Parameters("@Return_Value")

If spReturnVal <> 0 Then
    ErrMsg = cmd.Parameters("@ErrMsg")
    MsgBox ErrMsg, vbCritical, msgboxtitle
Else
    MsgBox "Processing Completed", vbInformation, msgboxtitle
End If
 
If Me.optSelect = 1 Then
    strProcCd = "SubStat"
Else
    strProcCd = "ClmStat"
End If

'MG 11/5/2013 Don't need below sql bc it already previous records in Load Worksheet Button
'strSQL = "Delete * from AUDITCLM_SubStatus_Tmp where UpdateUser = '" & userID & "'"

'DoCmd.SetWarnings (False)
'DoCmd.RunSQL (strSQL)
'DoCmd.SetWarnings (True)

Set MyCodeAdo = Nothing
Set cmd = Nothing

'MG 11-05-2013 filter based on user ID
Dim sqlString As String
sqlString = " UpdateUser = " & Chr(34) & UserID & Chr(34)
           
'MG refresh data sheet
Me.subFrmSubStatus.Form.filter = sqlString
Me.subFrmSubStatus.Form.FilterOn = True
Me.subFrmSubStatus.Form.Requery
Me.subFrmSubStatus.Form.Refresh

End Sub

Private Sub cmdSearch_Click()

Call RecordFilter

End Sub


Private Sub Form_Load()

Dim strSQL As String
Dim strSource As String
Dim StrAccessRights As String
Dim StrAccessRightsFINAL As String
Dim StrAccessRightsUSER As String

UserID = Identity.UserName
strSQL = "Delete * from AUDITCLM_SubStatus_Tmp where UpdateUser = '" & UserID & "'"
strSource = "Select * from AUDITCLM_SubStatus_Tmp where UpdateUser = '" & UserID & "'"

Me.lstSheetName.RowSource = ""
Me.TxtFileName = ""

DoCmd.SetWarnings (False)
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings (True)

Me.optSelect = 1
ControlVisible
Me.Refresh


'***************************
StrAccessRights = Nz(DLookup("[SupervisorID]", "ADMIN_User", "[UserID] ='" & Identity.UserName & "'"), "")

StrAccessRightsFINAL = Nz(DLookup("[SupervisorID]", "AUDITCLM_SubStatus_Rights", "[SupervisorID] ='" & StrAccessRights & "'"), "")

StrAccessRightsUSER = Nz(DLookup("[SupervisorID]", "AUDITCLM_SubStatus_Rights", "[SupervisorID] ='" & UserID & "'"), "")

If StrAccessRightsFINAL = "" And StrAccessRightsUSER = "" Then
    updateCtl False
End If

Me.subFrmSubStatus.Form.RecordSource = strSource

End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim strSQL As String

strSQL = "Delete * from AUDITCLM_SubStatus_Tmp where UpdateUser = '" & UserID & "'"
DoCmd.SetWarnings (False)
DoCmd.RunSQL (strSQL)
DoCmd.SetWarnings (True)

End Sub


Private Sub optSelect_AfterUpdate()

ControlVisible

End Sub

Private Sub txtCnlyClaimNumLkUp_AfterUpdate()

Call RecordFilter

End Sub
