Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:           Form_frm_SCANNING_MR_Invoice_Properties
' Author:       Barbara Dyroff
' Date:         2010-02-03
' Description:
'   Display Invoice properties and totals for Medical Record Scanning for a given Provider.
' Refresh the detail datasheet for the current Invoice.
'
' Modification History:
'
'20100706 Added GrossInvoiceAmt, MailShipTotAmt, ChkStatusCd, SpecialInstructions01Txt, SpecialInstructions01Type, SpecialInstructions02Txt, SpecialInstructions02Type,
'   RequestingDocID, InvoiceType, Reference01Txt, Reference01Type, SellerNm, SellerNumAgency, SellerAddr01, SellerAddr02, SellerCity, SellerState, SellerZip to display by Rob Hall
'
' =============================================

Private strMRInvPropTotTableName As String
Private strMRInvPropTotKey As String


Public Property Let PropTotTableName(data As String)
    strMRInvPropTotTableName = data

End Property

Public Property Let PropTotKey(data As String)
    strMRInvPropTotKey = data
End Property


' Refresh the Invoice properties and totals.
Public Sub RefreshData()
    Dim MyAdo As clsADO
    Dim strSQL As String
    Dim strSellerNum As String
    
    Set MyAdo = New clsADO
    MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
    
    strSellerNum = Right(strMRInvPropTotKey, Len(strMRInvPropTotKey) - (InStr(strMRInvPropTotKey, "=")))

    If strMRInvPropTotTableName <> "" Then
        strSQL = "SELECT AP.InvoiceNum, AP.InvoiceDate, CnlyProvID, ProvNum, SUM(PageCnt) AS TotPageCnt, InvoiceTotAmt, GrossInvoiceAmt, MailShipTotAmt, ChkStatusCd, SpecialInstructions01Txt, SpecialInstructions01Type, SpecialInstructions02Txt, SpecialInstructions02Type, RequestingDocID, InvoiceType, Reference01Txt, Reference01Type, SellerNm, SellerNumAgency, SellerAddr01, SellerAddr02, SellerCity, SellerState, SellerZip " & _
        "FROM CMS_AUDITORS_Claims..AP_Invoice AP Left Join " & strMRInvPropTotTableName & " MR  on AP.InvoiceNum = MR.InvoiceNum"
          
        If strMRInvPropTotKey <> "" Then
            strSQL = strSQL & " WHERE SellerNum =" & strSellerNum
        End If
        
        strSQL = strSQL & " GROUP BY AP.InvoiceDate, AP.InvoiceNum, CnlyProvID, ProvNum, InvoiceTotAmt, GrossInvoiceAmt, MailShipTotAmt, ChkStatusCd, SpecialInstructions01Txt, SpecialInstructions01Type, SpecialInstructions02Txt, SpecialInstructions02Type, RequestingDocID, InvoiceType, Reference01Txt, Reference01Type, SellerNm, SellerNumAgency, SellerAddr01, SellerAddr02, SellerCity, SellerState, SellerZip"
        
        strSQL = strSQL & " ORDER BY AP.InvoiceDate DESC, AP.InvoiceNum, CnlyProvID, ProvNum"
        
        MyAdo.sqlString = strSQL
        Set Me.RecordSet = MyAdo.OpenRecordSet
        
        MyAdo.DisConnect
    End If
    
    Set MyAdo = Nothing
End Sub

'Refresh the Invoice Detail datasheet for the current Invoice.
Private Sub Form_Current()
    Dim strSQL As String
    If IsSubForm(Me) And strMRInvPropTotTableName <> "" Then
        If Not (Me.RecordSet Is Nothing) Then
            strSQL = "SELECT * FROM " & strMRInvPropTotTableName

            If strMRInvPropTotKey <> "" Then
                strSQL = strSQL & " WHERE " & _
                    strMRInvPropTotKey & " AND InvoiceNum = '" & Me.txtInvoiceNum & "'" & _
                    " ORDER BY ScannedDt DESC"
            End If

            Me.Parent.Form.txtSQLSource = strSQL
            Me.Parent.RefreshDetail
        End If
    End If
End Sub
