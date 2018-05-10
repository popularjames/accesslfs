Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CMS_ExportToExcel()
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As DAO.RecordSet
Dim strScreenID As String

    strProcName = TypeName(Me) & ".cmdExportToExcel_Click"
    strScreenID = Me.OpenArgs
'    Set oRs = Me.GridForm.Recordset
    
        ' for the main grid all the time.
    If Not Scr(strScreenID).SubForm.Form.RecordSet Is Nothing Then
        Set oRs = Scr(strScreenID).SubForm.Form.RecordsetClone
        'Set oRs = Me.SubForm.Form.RecordsetClone
         Call mod_CMS_General.ExportRsToExcel(oRs)
    End If

    DoCmd.Close acForm, Me.Name

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Load()
    CMS_ExportToExcel
End Sub
