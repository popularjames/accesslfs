Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'=============================================
' ID:          Form_frm_RPT_Generic_Report_Simple
'
' Description:
'      Prompt the user to enter a Date Range for the calling report.  Select the data and
' create a user specific reporting table.  Display the results.
'
' Modification History:
'   2013-01-25 by BJD to use the ADODB command to execute the report stored procedure.  This was to
'       resolve the query timeout issue.
'
' =============================================


Private mstrUserProfile As String
Private mbRecordChanged As Boolean
Private miAppPermission As Integer
Private mbLocked As Boolean
Private mReturnDate As Date
Private ColReSize As clsAutoSizeColumns
Private mstrStoredProcName As String

Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1

Const CstrFrmAppID As String = "QueueProductivity"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Let StoredProcName(data As String)
    mstrStoredProcName = data
End Property

Public Property Get StoredProcName() As String
    StoredProcName = mstrStoredProcName
End Property

Private Sub cmdExportToExcel_Click()
    Dim bExport As Boolean
    
    If lstDetail.RecordSet Is Nothing Then Exit Sub     'nothing to do
    If lstDetail.ListCount = 1 Then Exit Sub     'only row headers, nothing to do
    
    bExport = ExportDetails(Me.lstDetail.RecordSet, "Queue_Productivity.xls")
    
    'If bExport = False Then
    '     MsgBox "An error was encountered while attempting to export Detail data to Excel.", vbCritical
    'End If
End Sub


Private Function ExportDetails(rst As ADODB.RecordSet, strFilePath As String) As Boolean
'Function to export recordset to excel file

    Dim dlg As clsDialogs
    Dim cie As clsImportExport

    Set cie = New clsImportExport
    Set dlg = New clsDialogs

    
    If rst Is Nothing Then Exit Function     ' nothing to save
    
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
            
            If rst.recordCount > 65535 Then
                MsgBox "Warning: Your recordset contains more than 65535 rows, the maximum number of rows allowed in Excel.  " & _
                Trim(str(rst.recordCount - 65535)) & " rows will not be displayed.", vbCritical
            End If
                        
        Else
            GoTo exitHere
        End If
     
        With cie
            .ExportExcelRecordset rst, strFilePath, True
        End With
         
    End With
    
    ExportDetails = True

exitHere:
    Set cie = Nothing
    Set dlg = Nothing
    Exit Function
    
HandleError:
    ExportDetails = False
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly, "Error in clsImportExport"
    GoTo exitHere

End Function



Public Sub RefreshData()
    On Error GoTo ErrHandler
    
    'Use the to execute the reporting stored procedure.
    Dim cmd As ADODB.Command
    
    Dim rst As ADODB.RecordSet
    Dim iResult As Integer
    Dim strRPTTableName As String
    Dim strSQL As String
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strRPTTableName = Replace(GetUserName(), ".", "_") & "_Adhoc_Report"
    myCode_ADO.SQLTextType = sqltext
        
    'myCode_ADO.sqlString = "exec " & Me.StoredProcName & " '" & GetUserName() & "', '" & Format(FromDt, "mm-dd-yyyy") & "','" & Format(ThruDt, "mm-dd-yyyy") & "'"
    'Changing the reports to be triggered by a common stored proc run as DBO
    myCode_ADO.sqlString = "exec [dbo].[usp_Run_UniversalReport] '" & Me.StoredProcName & "', '" & GetUserName() & "', '" & Format(FromDt, "mm-dd-yyyy") & "','" & Format(ThruDt, "mm-dd-yyyy") & "'"


    ' Old version replaced with the following ADODB.Command to resolve a query timeout issue.
    ' iResult = myCode_ADO.Execute
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = myCode_ADO.CurrentConnection
    cmd.CommandTimeout = 0
    cmd.commandType = adCmdText
    cmd.CommandText = myCode_ADO.sqlString
    cmd.Execute
    
    ' Query the data for display from the user specific reporting table created by the reporting stored procedure.
    strSQL = "select * from " & strRPTTableName
    Set rst = myCode_ADO.OpenRecordSet(strSQL)
    
    'Setting the list record set equal to the specic ADO-class record set
    lstDetail.ColumnCount = rst.Fields.Count
    Set lstDetail.RecordSet = rst

    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstDetail
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstDetail.ListCount - 1 > 0 Then
        ColReSize.AutoSize
    End If
    Set ColReSize = Nothing

   
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical

End Sub

Private Sub cmdGetReport_Click()
    If FromDt & "" = "" Then
        MsgBox "Please enter a from date"
    ElseIf ThruDt & "" = "" Then
        MsgBox "Please enter a through date"
    Else
        RefreshData
    End If
End Sub

Private Sub FromDt_Exit(Cancel As Integer)
    If FromDt & "" <> "" Then
        If Not IsDate(FromDt) Then
            MsgBox "Please enter a valid date"
            Cancel = True
        End If
    End If
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & vbCrLf & ErrSource
End Sub

Private Sub ThruDt_Exit(Cancel As Integer)
    If ThruDt & "" <> "" Then
        If Not IsDate(ThruDt) Then
            MsgBox "Please enter a valid date"
            Cancel = True
        End If
    End If
End Sub
