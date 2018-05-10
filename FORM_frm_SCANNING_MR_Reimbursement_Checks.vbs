Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1
Private ColReSize As clsAutoSizeColumns

Private Sub cmdExport_Click()

    Dim bExport As Boolean
    Dim strNow As String
    Dim iResults As Integer
    Dim strSQL As String
    
    strNow = Format(Now(), "yyyymmdd")
    
    If lstChks.RecordSet Is Nothing Then Exit Sub     'nothing to do
    If lstChks.ListCount = 1 Then
        Exit Sub     'only row headers, nothing to do
    Else
        bExport = ExportDetails(Me.lstChks.RecordSet, "ProviderInvoice_" & strNow & ".xls")
        'Export to send to accounting for checks to be generated
        If bExport = True Then
            Set myCode_ADO = New clsADO
            myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
            'Update records from requested to sent.
            strSQL = "Update CMS_AUDITORS_Claims.dbo.AP_Invoice Set ChkStatusCd = 'SENT' WHERE ChkStatusCd = 'PEND'"
            myCode_ADO.sqlString = strSQL
            myCode_ADO.SQLTextType = sqltext
            iResults = myCode_ADO.Execute()
            MsgBox "Number of records exported: " & iResults, vbOKOnly
            Set MyAdo = Nothing
            Set myCode_ADO = Nothing
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdGenChk_Click()
    
    Dim strSQL As String
    Dim sprocReimbursements As clsAdoSproc
    Dim sprocAdjustments As clsAdoSproc
    Dim sprocInvoice As clsAdoSproc
    Dim myCode_ADO As clsADO
    
    On Error GoTo HandleError
    
    'Add records to SCANNING_MR_Invoice
    Set sprocReimbursements = New clsAdoSproc
    sprocReimbursements.RefTable = "v_CODE_Database"
    sprocReimbursements.CommandText = "usp_Scanning_MR_Reimbursements"
    sprocReimbursements.Setup
    sprocReimbursements.Exec
        
    If sprocReimbursements.ReturnValue = 0 Then
        'Calculate and insert adjustments into SCANNING_MR_Invoice from Scanning_MR_Adjustments
        Set sprocAdjustments = New clsAdoSproc
        sprocAdjustments.RefTable = "v_CODE_Database"
        sprocAdjustments.CommandText = "usp_Scanning_MR_Adjustments"
        sprocAdjustments.Setup
        sprocAdjustments.Exec
        If sprocAdjustments.ReturnValue = 0 Then
            'Insert all records into AP_Invoice from SCANNING_MR_Invoice
            Set sprocInvoice = New clsAdoSproc
            sprocInvoice.RefTable = "v_CODE_Database"
            sprocInvoice.CommandText = "usp_Scanning_MR_Invoice"
            sprocInvoice.Setup
            sprocInvoice.Exec
            If sprocInvoice.ReturnValue = 0 Then
                'Display records from AP_Invoice
                Set myCode_ADO = New clsADO
                myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
                strSQL = "SELECT SellerNum VendorID, SellerNm VendorName, SellerNm VendorCheckName, SellerAddr01 Address1, SellerAddr02 Address2, SellerAddr03 Address3, SellerCity City, SellerState State, SellerZip Zip, CONVERT(VARCHAR,GETDATE(),110) PostingDate, 'PROVIDER_REIMB' CheckbookID, InvoiceNum InvoiceNumber, RequestingDocId Description, CONVERT(VARCHAR,InvoiceDate,110) InvDate, CONVERT(VARCHAR,InvoiceDate,110) DueDate, ROUND(InvoiceTotAmt,2) Amount, SpecialInstructions01Txt DistReference, '531330-107-HCG-320-GA' PurchAccount, '201000-000-000-000-GA' PayAccount, '03085' AuditNumber "
                strSQL = strSQL & "FROM CMS_AUDITORS_Claims.dbo.AP_Invoice WHERE ChkStatusCd = 'PEND' order by SellerNum"
                lstChks.RowSource = strSQL
                myCode_ADO.sqlString = lstChks.RowSource
                Set lstChks.RecordSet = myCode_ADO.OpenRecordSet()
                lstChks.ColumnCount = lstChks.RecordSet.Fields.Count
                'count of records
                txtCount = lstChks.ListCount - 1
            Else
                MsgBox "Error Moving Records to AP_Invoice. " & vbCr & sprocInvoice.GetParam("@pErrMsg")
            End If
        Else
            MsgBox "Error Moving Records from Scanning_MR_Adjustments . " & vbCr & sprocAdjustments.GetParam("@pErrMsg")
        End If
    Else
        MsgBox "Error Moving Records to SCANNING_MR_Invoice. " & vbCr & sprocReimbursements.GetParam("@pErrMsg")
    End If

    
    Set ColReSize = New clsAutoSizeColumns
    ColReSize.SetControl Me.lstChks
    'don't resize if lstChks is null
    On Error Resume Next
    If Me.lstChks.ListCount > 0 Then
        ColReSize.AutoSize
    End If
    Set ColReSize = Nothing
    
    cmdGenChk.Enabled = False
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Exit Sub
     
HandleError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GoTo Exit_Sub
    
End Sub
