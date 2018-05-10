Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrUserProfile As String
Private mbRecordChanged As Boolean
Private miAppPermission As Integer
Private mbLocked As Boolean
Private mReturnDate As Date
Private ColReSize As clsAutoSizeColumns

Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Const CstrFrmAppID As String = "QueueProductivity"


Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
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
    Dim rst As ADODB.RecordSet
    Dim iResult As Integer
    Dim strRPTTableName As String
    Dim strSQL As String
    'Creating a new instance of ADO-class variable
    Set myCode_ADO = New clsADO
    
    'Making a Connection call to SQL database?
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    
    'Setting the ADO-class sqlstring to the specified SQL query statement
    strRPTTableName = Replace(Identity.UserName(), ".", "_") & "_Adhoc_Report"
    myCode_ADO.SQLTextType = sqltext
    myCode_ADO.sqlString = "usp_RPT_Queue_Productivity_Summary '" & Identity.UserName() & "'"
    iResult = myCode_ADO.Execute
    
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


Private Sub Form_Load()
    RefreshData
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & vbCrLf & ErrSource
End Sub
