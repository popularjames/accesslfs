Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'=============================================
' ID:          Form_frm_LDGR_Main
' Author:      Barbara Dyroff / Kevin Dearing
' Create Date: 2012-06-27
' Description:
'      Display Transaction Ledger Information.  For each Connolly Claim display the Claim totals for
' the AR Setup, Collectiions, Invoice, ID (Projected Savings) and balance info.  Display the detailed
' Transaction for a selected claim, sorted by the Transaction Date and providing the balance info
' on the given Transaction date.
'
' Modification History:
'  2012-10-26 by BJD to add Connolly Manual Adjustments for Collections.
'  2012-12-21 by BJD to call the mod_GENERAL_Navigate sub Navigate to instantiate the Connolly Adj Form.
'  2013-02-12 by BJD to change the default view to the last 100 updated Claims.
'  2013-05-09 by BJD to add the Prepayment AR/AP calculations.
'  2014-02-19 by BJD to add Connolly Manual Adjustments for AR Setup.
'
' =============================================

Private Const ccsFrmAppID As String = "Ldgr"  'Used for form security

'Constants for CreditOrDebitCd
Private Const strCREDIT_CD As String = "C"
Private Const strDEBIT_CD As String = "D"
'Constants for LdgrTransGroupCd
Private Const strLDGRTRANSGROUPCD_ARSETUP As String = "ARSETUP"
Private Const strLDGRTRANSGROUPCD_ARSETUP_CNLY As String = "ARSETUP-CNLY"
Private Const strLDGRTRANSGROUPCD_COLLECTION As String = "COLLECTION"
Private Const strLDGRTRANSGROUPCD_INVOICE As String = "INVOICE"
Private Const strLDGRTRANSGROUPCD_COLLECTION_CNLY As String = "COLLECTION-CNLY"
Private Const strLDGRTRANSGROUPCD_ARSETUP_PREPAY As String = "ARSETUP-PREPAY"

'Constants for LdgrTransTypeCd
Private Const strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_UP As String = "COLL-ARSETUP-OVERPAY-ADJUST-UP"
Private Const strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_DOWN As String = "COLL-ARSETUP-OVERPAY-ADJUST-DOWN"
Private Const strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_UP As String = "COLL-CNLY-ARSETUP-OVERPAY-ADJUST-UP"
Private Const strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_DOWN As String = "COLL-CNLY-ARSETUP-OVERPAY-ADJUST-DOWN"

'Constant for Review Type.
Private Const strREVIEWTYPE_PRP As String = "PRP"

Private WithEvents frmCOLLCNLYAdj As Form_frm_COLL_CNLY_Adj
Attribute frmCOLLCNLYAdj.VB_VarHelpID = -1
Private WithEvents sfrmHeader As Form_frm_LDGR_Hdr
Attribute sfrmHeader.VB_VarHelpID = -1
Private WithEvents sfrmDetail As Form_frm_LDGR_Dtl
Attribute sfrmDetail.VB_VarHelpID = -1
'Private WithEvents sfrmMain As Form_frm_LDGR_Main

Private WithEvents coDetailRs As ADODB.RecordSet
Attribute coDetailRs.VB_VarHelpID = -1
Private WithEvents coDetailRsExt As ADODB.RecordSet
Attribute coDetailRsExt.VB_VarHelpID = -1
Private WithEvents coHeaderRs As ADODB.RecordSet
Attribute coHeaderRs.VB_VarHelpID = -1

Private csCurrentClaimNum As String
Private strHdrClaimNumList As String
Private strLastHdrClaimNumList As String

Private csTmpTableName As String

Private miAppPermission As Integer
Private mbAllowChange As Boolean
Private mbAllowAdd As Boolean
Private mbAllowView As Boolean

Public Property Get frmAppID() As String
    frmAppID = ccsFrmAppID
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get CurrentClaimNum() As String
    CurrentClaimNum = csCurrentClaimNum
    
End Property

Public Property Let CurrentClaimNum(sCurrentClaimNum As String)
    csCurrentClaimNum = sCurrentClaimNum
End Property

Public Property Get HdrClaimNumList() As String
    HdrClaimNumList = strHdrClaimNumList
End Property

Public Property Let HdrClaimNumList(strClaimNumList As String)
    strHdrClaimNumList = strClaimNumList
End Property

Public Property Get LastHdrClaimNumList() As String
    LastHdrClaimNumList = strLastHdrClaimNumList
End Property

Public Property Let LastHdrClaimNumList(strClaimNumList As String)
    strLastHdrClaimNumList = strClaimNumList
End Property

Private Property Let TempTableNm(strUserNm As String)
    Dim strUserNmFormatted As String
    Dim strSearchChar As String
    
    strSearchChar = "."
    strUserNmFormatted = Replace(strUserNm, strSearchChar, "")
    csTmpTableName = "tmp_Local_LDGR_Dtl_" & strUserNmFormatted
End Property


'Call the Connolly Manual Adjustments (display/insert/update) for the current claim.
Private Sub cmdManualAdjustment_Click()
    
On Error GoTo ErrHandler
        
    Dim strError As String
    Dim strParameterString As String
    Dim strParent As String
    Dim strSearchType As String
    Dim strAction As String
    
    strParameterString = CurrentClaimNum
    strParent = "frm_GENERAL_QuickLookup"
    strSearchType = "COLLMANUAL"
    strAction = "DblClick"
    
    If strParameterString <> "" Then
        Navigate strParent, strSearchType, strAction, strParameterString
    End If

Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub

'Private Sub coHeaderRs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Stop
'
'End Sub


Private Sub Form_Load()
On Error GoTo Err_Form_Load
    Dim strProcName As String
    Dim iAppPermission As Integer

    'Init
    strProcName = ClassName & ".sfrmHeader_Current"
    TempTableNm = Environ("UserName") ' Assign the name of the temp Table to create on SQL Server.
    Me.Caption = "AR Transaction Ledger"
    Me.HdrClaimNumList = ""
    Me.LastHdrClaimNumList = ""
    
    Call Account_Check(Me)
    iAppPermission = UserAccess_Check(Me)
    
    If iAppPermission = 0 Then
        Exit Sub
    End If
    
    Set sfrmHeader = Me.frm_LDGR_Hdr.Form
    Set sfrmDetail = Me.frm_LDGR_Dtl.Form
    
    DoCmd.Echo True, "Refreshing grids"

    RefreshData
    
Exit_Form_Load:
    Exit Sub
Err_Form_Load:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    Resume Exit_Form_Load
End Sub


Public Sub RefreshData()
On Error GoTo Err_RefreshData
    Dim strProcName As String
    Dim oAdo As clsADO
    Dim sSql As String
'    Dim dtAged As Date

    strProcName = ClassName & ".RefreshData"
    
    ' Set the Header Form Recordset.
'    dtAged = FormatDateTime(DateAdd("d", -47, Date), vbShortDate) 'May use later on.
    If (StrComp(strHdrClaimNumList, "") = 0) Or (StrComp(strHdrClaimNumList, "") = Null) Then
        sSql = "SELECT TOP 100 * FROM v_LDGR_Hdr_Entries ORDER BY LdgrLastUpDt DESC, AutoID DESC"
'        sSql = "SELECT * FROM v_LDGR_Hdr_Entries WHERE ABS(CollFeeBalanceAmt) > 5.00 AND ClmStatus LIKE '5%' ORDER BY ABS(CollFeeBalanceAmt) DESC"
'        sSql = "SELECT * FROM v_LDGR_Hdr_Entries WHERE CnlyClaimNum = '515101720184500008189390104402'"  ' Test with one.
    Else
        sSql = "SELECT * FROM v_LDGR_Hdr_Entries WHERE CnlyClaimNum IN " & Me.HdrClaimNumList & " ORDER BY CnlyClaimNum"
    End If

    ' select the bottom one
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_LDGR_Hdr_Entries")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set coHeaderRs = .ExecuteRS
    End With
    
    'MsgBox "Hdr Record Count[" & coHeaderRs.RecordCount & "]"
    
    ' If no records are found.  Select the last set.
    If coHeaderRs.recordCount = 0 Then
        MsgBox "No Ledger records found.", vbOKOnly + vbInformation
        Me.HdrClaimNumList = Me.LastHdrClaimNumList
        If (StrComp(strHdrClaimNumList, "") = 0) Or (StrComp(strHdrClaimNumList, "") = Null) Then
            sSql = "SELECT TOP 100 * FROM v_LDGR_Hdr_Entries ORDER BY LdgrLastUpDt DESC, AutoID DESC"
            'sSql = "SELECT * FROM v_LDGR_Hdr_Entries WHERE ABS(CollFeeBalanceAmt) > 5.00 AND ClmStatus LIKE '5%' ORDER BY CollFeeBalanceAmt DESC"
        Else
            sSql = "SELECT * FROM v_LDGR_Hdr_Entries WHERE CnlyClaimNum IN " & Me.HdrClaimNumList & " ORDER BY CnlyClaimNum"
        End If
        Set oAdo = New clsADO
        With oAdo
            .ConnectionString = GetConnectString("v_LDGR_Hdr_Entries")
            .SQLTextType = sqltext
            .sqlString = sSql
            Set coHeaderRs = .ExecuteRS
        End With
    End If
    
    Set sfrmHeader.RecordSet = coHeaderRs

    DoCmd.Hourglass True
    DoCmd.Echo True, "Searching ..."
    
    sfrmHeader_Current
   
Exit_RefreshData:
    DoCmd.Echo True, "Ready..."
    DoCmd.Hourglass False
    Exit Sub
Err_RefreshData:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    Resume Exit_RefreshData
End Sub


' Clean up on Unload.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload
    Dim strProcName As String

    strProcName = ClassName & ".RefreshData"

    ' Drop the Temp Table used for processing.
    If IsTable(csTmpTableName) = True Then
        Set sfrmDetail.RecordSet = coDetailRs
        
        CurrentDb.TableDefs.Delete (csTmpTableName)
        CurrentDb.TableDefs.Refresh

    End If
    
    'Erase the last instance of the derived extended Recordset.
    If Not (coDetailRsExt Is Nothing) Then
        Set coDetailRsExt = Nothing
    End If
    
Exit_Form_Unload:
    Exit Sub

Err_Form_Unload:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    Resume Exit_Form_Unload
End Sub

' Refresh the Ledger data for the current claim.
Private Sub sfrmHeader_Current()
On Error GoTo Err_sfrmHeader_Current
    Dim strProcName As String
    Dim oAdo As clsADO
    Dim sSql As String

    strProcName = ClassName & ".sfrmHeader_Current"

    ' Get the headers ID
    CurrentClaimNum = sfrmHeader.CurrentClaimNum
    
    ' and then set the detail forms recordset..
    sSql = "SELECT * FROM v_LDGR_Dtl_Entries WHERE DeleteInd = 0 AND CnlyClaimNum = '" & Me.CurrentClaimNum & "' ORDER BY LdgrTransDt"

    ' select the bottom one
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_LDGR_Dtl_Entries")
        .SQLTextType = sqltext
        .sqlString = sSql
        Set coDetailRs = .ExecuteRS
    End With
    
'    MsgBox "Record Count[" & coDetailRs.RecordCount & "]"
    'Prepare Ledger Display Info
    If Not (coDetailRs Is Nothing) And (coDetailRs.recordCount = 0) Then
        MsgBox "There are not any Detail records to display."
    ElseIf Not (coDetailRs Is Nothing) And (coDetailRs.recordCount > 0) Then
        ' Copy to a working Recordset to derive data.
        If Not (coDetailRsExt Is Nothing) Then
            Set coDetailRsExt = Nothing 'Erase the last instance.
        End If
        
        Set coDetailRsExt = CopyRecordset(coDetailRs)

        ' Derive the Ledger Balance Info etc. for each Transaction date.
        DeriveLdgrDtlInfo coDetailRsExt

        ' Test Display
'        coDetailRsExt.MoveFirst
'        MsgBox "RecordCount[" & coDetailRsExt.RecordCount & "]"
'        Do Until coDetailRsExt.EOF
'            MsgBox "LdgrTransTypeCd[" & coDetailRsExt![LdgrTransTypeCd] & "]  " _
'                    & "CreditOrDebitCd[" & coDetailRsExt![CreditOrDebitCd] & "]  " _
'                    & "LdgrTransAmt[" & coDetailRsExt![LdgrTransAmt] & "]  " _
'                    & "LdgrTransDt[" & coDetailRsExt![LdgrTransDt] & "]  " _
'                    & "FeeRt[" & coDetailRsExt![FeeRt] & "]  " _
'                    & "PrincRecovOrPaidAmt[" & coDetailRsExt![PrincRecovOrPaidAmt] & "]  " _
'                    & "LdgrEntryFeeAmt[" & coDetailRsExt![LdgrEntryFeeAmt] & "]  " _
'                    & "CollBalanceAmt[" & coDetailRsExt![CollBalanceAmt] & "]  " _
'                    & "CollFeeBalanceAmt[" & coDetailRsExt![CollFeeBalanceAmt] & "]  " _
'                    & "AutoId[" & coDetailRsExt![AutoId] & "]  "
'            coDetailRsExt.MoveNext
'        Loop
        ' End Test Display


    ' Derived record set assignment  -- ***This does not work. Form display data problem.  Use the following temp Table instead.
'    Set sfrmDetail.Recordset = coDetailRsExt
    
        ' Copy the data into a local table that we don't care about
        Call CopyDataToLocalTmpTable(coDetailRsExt, False)

        sfrmDetail.RecordSource = csTmpTableName
    
    Else
        MsgBox "An error occurred displaying the Detailed Transactions"
        Err.Raise adErrFeatureNotAvailable
    End If

Exit_sfrmHeader_Current:
    Exit Sub
Err_sfrmHeader_Current:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    Resume Exit_sfrmHeader_Current
End Sub

' Prompt the user to select a Claim.  Provide the general CnlyClaimNum lookup.
Private Sub cmdGetClaim_Click()
On Error GoTo Err_cmdGetClaim_Click
    Dim frmLDGRQuicklookup As Form_frm_LDGR_QuickLookup
    
    Me.LastHdrClaimNumList = Me.HdrClaimNumList  'Retain the last selection.
    
    Set frmLDGRQuicklookup = New Form_frm_LDGR_QuickLookup
    
    frmLDGRQuicklookup.SearchType = "AUDITCLM"
     
    frmLDGRQuicklookup.CallingForm = Me.Form
    
    frmLDGRQuicklookup.RefreshData
    ShowFormAndWait frmLDGRQuicklookup
    
    ' frmLDGRQuicklookup will update the Claim selection HdrClaimNumList in this Main module.
    Set frmLDGRQuicklookup = Nothing

    Me.RefreshData
    
Exit_cmdGetClaim_Click:
    Exit Sub

Err_cmdGetClaim_Click:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Resume Exit_cmdGetClaim_Click
    
End Sub

' Copy ADODB field structure to a new ADODB Recordset
Private Function CopyFields(rs As ADODB.RecordSet) As ADODB.RecordSet
On Error GoTo Err_CopyFields
    Dim newRS As New ADODB.RecordSet, fld As ADODB.Field
    
    For Each fld In rs.Fields
        newRS.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes 'Note Attribute 104 changes to 108
    Next
    Set CopyFields = newRS
    
Exit_CopyFields:
    Exit Function

Err_CopyFields:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Resume Exit_CopyFields
End Function

' Copy an ADODB Recordset to a new Recordset instance.
Private Function CopyRecordset(rs As ADODB.RecordSet) As ADODB.RecordSet
On Error GoTo Err_CopyRecordset
    Dim newRS As New ADODB.RecordSet, fld As ADODB.Field
    Set newRS = CopyFields(rs)
    newRS.Open  'You must open the Recordset before adding new records.
    
    rs.MoveFirst
    Do Until rs.EOF
        newRS.AddNew
        For Each fld In rs.Fields
            newRS(fld.Name) = fld.Value  'Assumes no BLOB fields
        Next
        rs.MoveNext
    Loop
    Set CopyRecordset = newRS

Exit_CopyRecordset:
    Exit Function

Err_CopyRecordset:
    MsgBox Err.Description, vbOKOnly + vbCritical
    Resume Exit_CopyRecordset
End Function

'Populate the SQL Server temp Table for the current claim.
Private Function CopyDataToLocalTmpTable(oRs As ADODB.RecordSet, bForceRemake As Boolean) As String
On Error GoTo Block_Err
    Dim strProcName As String
    Dim oFld As ADODB.Field
    Dim oDaoRs As DAO.RecordSet

    strProcName = ClassName & ".CopyDataToLocalTmpTable"
    
    If IsTable(csTmpTableName) = False Or bForceRemake = True Then
        Call CreateTableFromADORS(oRs, csTmpTableName, bForceRemake)
    End If
    
    ' Make sure it's empty
    CurrentDb.Execute "DELETE FROM [" & csTmpTableName & "]"
    
    Set oDaoRs = CurrentDb.OpenRecordSet(csTmpTableName, dbOpenTable)
    
    ' populate it:
    oRs.MoveFirst
    While Not oRs.EOF
        oDaoRs.AddNew
        For Each oFld In oRs.Fields
            If oFld.Name = "AutoId" Then
                oDaoRs(oFld.Name) = CStr(oFld.Value)
            Else
                oDaoRs(oFld.Name) = oFld.Value
            End If
        Next
        oDaoRs.Update
        oRs.MoveNext
    Wend
    
Block_Exit:
    Set oDaoRs = Nothing
    Set oFld = Nothing
    Exit Function
Block_Err:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


' Create a SQL Server temp Table for the current claim Recordset.
Private Function CreateTableFromADORS(oRs As ADODB.RecordSet, sTblName As String, Optional bForceRemake As Boolean = False) As String
On Error GoTo Block_Err
    Dim strProcName As String
    Dim oTDef As DAO.TableDef
    Dim oAdoField As ADODB.Field
    Dim oTblFld As DAO.Field

    strProcName = ClassName & ".CreateTableFromADORS"
    
    If bForceRemake = True Then
        If IsTable(sTblName) = True Then
            CurrentDb.TableDefs.Delete (sTblName)
            CurrentDb.TableDefs.Refresh
        End If
    ElseIf IsTable(sTblName) = True Then
            ' already created. nothing to do
        CreateTableFromADORS = sTblName
        GoTo Block_Exit
    End If
    
    Set oTDef = New DAO.TableDef
    With oTDef
        .Name = sTblName
        For Each oAdoField In oRs.Fields
            Set oTblFld = New DAO.Field
            oTblFld.Name = oAdoField.Name

            If oAdoField.Name = "AutoId" Then 'Trouble displaying dbBigInt on form and then trouble using ado adInteger with larger numbers.
                oTblFld.Type = dbText ' Convert to a string for display.
            ElseIf (oAdoField.Type = adVarChar) Then 'Allow Zero Length strings.
                oTblFld.Type = AdoTypeToDaoType(oAdoField)
                oTblFld.AllowZeroLength = True
            Else
                oTblFld.Type = AdoTypeToDaoType(oAdoField)
            End If
            .Fields.Append oTblFld
        Next

    End With
    
    CurrentDb.TableDefs.Append oTDef
    CreateTableFromADORS = sTblName
    CurrentDb.TableDefs.Refresh
    
Block_Exit:
    Exit Function
Block_Err:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


' Derive the additional Ledger Transaction detail on each Transaction date.
Private Sub DeriveLdgrDtlInfo(rsDtlExt As ADODB.RecordSet)
    On Error GoTo Err_DeriveLdgrDtlInfo

    Dim strProcName As String
    Dim curLdgrTransAmt As Currency
    Dim curPrincRecovOrPaidAmt As Currency
    Dim curInvoiceFeeAmt As Currency
    Dim curTotAROrAPAmt As Currency
    Dim curTotARAdjAmt As Currency
    Dim curNetAROrAPAmt As Currency
    Dim curTotPrincRecovOrPaidAmt As Currency
    Dim curTotPrincRecovOrPaidFeeAmt As Currency
    Dim curTotInvoiceFeeAmt As Currency
    Dim curLdgrEntryFeeAmt As Currency
    Dim curLastTotPrincRecovOrPaidFeeAmt As Currency
    Dim curNetAROrAPFeeAmt As Currency
    Dim curLastNetAROrAPFeeAmt As Currency

    strProcName = ClassName & ".DeriveLdgrDtlInfo"

    ' Init Totals for the Claim
    curTotAROrAPAmt = 0
    curTotARAdjAmt = 0
    curNetAROrAPAmt = 0
    curTotPrincRecovOrPaidAmt = 0
    curTotPrincRecovOrPaidFeeAmt = 0
    curTotInvoiceFeeAmt = 0
    curNetAROrAPFeeAmt = 0
    
    rsDtlExt.MoveFirst
    Do Until rsDtlExt.EOF
    
        curLdgrTransAmt = rsDtlExt![LdgrTransAmt]
        curPrincRecovOrPaidAmt = 0 'Applies to the current Transaction if it is a Collection Activity.
        curInvoiceFeeAmt = 0 'Applies to the current Transaction if it is an Invoice Activity.
        curLdgrEntryFeeAmt = 0
        curLastTotPrincRecovOrPaidFeeAmt = curTotPrincRecovOrPaidFeeAmt ' Retain the last to calculate the fee entry for Collections.
        curLastNetAROrAPFeeAmt = curNetAROrAPFeeAmt ' Retain the last to calculate the fee entry for Prepayment AR/AP.
        
        'Validate that that Credit or Debit identifier has been defined.
        If (rsDtlExt![CreditOrDebitCd] <> strCREDIT_CD) And (rsDtlExt![CreditOrDebitCd] <> strDEBIT_CD) Then
            MsgBox "Unable to derive additional Transaction detail data for the current claim.  Credit/Debit not defined for this Transaction.  Please contact application support.  See " & strProcName, vbOKOnly + vbCritical
            Exit Sub
        End If
        
        ' AR Setup Activity.
        If (rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_ARSETUP) Or (rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_ARSETUP_PREPAY) Or (rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_ARSETUP_CNLY) Then
            If rsDtlExt![CreditOrDebitCd] = strDEBIT_CD Then
                curTotAROrAPAmt = curTotAROrAPAmt + (curLdgrTransAmt * -1)
            Else
                curTotAROrAPAmt = curTotAROrAPAmt + (curLdgrTransAmt)
            End If
            curNetAROrAPAmt = curTotAROrAPAmt + curTotARAdjAmt
            curNetAROrAPFeeAmt = Abs(curNetAROrAPAmt) * rsDtlExt![FeeRt]
            ' Assign Fee entry if for Prepay (Invoice Fee based on AR/AP since there are no collections.)
            If (rsDtlExt![Adj_ReviewType] = strREVIEWTYPE_PRP) Then
                curLdgrEntryFeeAmt = curNetAROrAPFeeAmt - curLastNetAROrAPFeeAmt  'This is a little tricky.  Using math to determine the amt to post.
            End If
        ' Collection Activity that adjusts the AR Setup.
        ElseIf (((rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_COLLECTION) _
                And ((rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_UP) _
                    Or (rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_DOWN) _
                    ) _
                ) _
                Or _
                ((rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_COLLECTION_CNLY) _
                And ((rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_UP) _
                    Or (rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_DOWN) _
                    ) _
                ) _
            ) Then
            If rsDtlExt![CreditOrDebitCd] = strDEBIT_CD Then
                curTotARAdjAmt = curTotARAdjAmt + (curLdgrTransAmt * -1)
            Else
                curTotARAdjAmt = curTotARAdjAmt + (curLdgrTransAmt)
            End If
            curNetAROrAPAmt = curTotAROrAPAmt + curTotARAdjAmt
            curNetAROrAPFeeAmt = Abs(curNetAROrAPAmt) * rsDtlExt![FeeRt]
            ' Assign Fee entry if for Prepay (Invoice Fee based on AR/AP since there are no collections.)
            If (rsDtlExt![Adj_ReviewType] = strREVIEWTYPE_PRP) Then
                curLdgrEntryFeeAmt = curNetAROrAPFeeAmt - curLastNetAROrAPFeeAmt  'This is a little tricky.  Using math to determine the amt to post.
            End If
        ' Collection Activity (No collection activity for Prepay.)
        ElseIf (((rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_COLLECTION) _
                And Not ((rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_UP) _
                        Or (rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_ARSETUP_OVERPAY_ADJUST_DOWN) _
                        )) _
                Or _
                ((rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_COLLECTION_CNLY) _
                And Not ((rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_UP) _
                        Or (rsDtlExt![LdgrTransTypeCd] = strTRANS_COLL_CNLY_ARSETUP_OVERPAY_ADJUST_DOWN) _
                        )) _
            ) Then
            If rsDtlExt![CreditOrDebitCd] = strDEBIT_CD Then
                curPrincRecovOrPaidAmt = curLdgrTransAmt * -1
                curTotPrincRecovOrPaidAmt = curTotPrincRecovOrPaidAmt + (curLdgrTransAmt * -1)
            Else
                curPrincRecovOrPaidAmt = curLdgrTransAmt
                curTotPrincRecovOrPaidAmt = curTotPrincRecovOrPaidAmt + (curLdgrTransAmt)
            End If
            curTotPrincRecovOrPaidFeeAmt = Abs(curTotPrincRecovOrPaidAmt) * rsDtlExt![FeeRt]
            curLdgrEntryFeeAmt = curTotPrincRecovOrPaidFeeAmt - curLastTotPrincRecovOrPaidFeeAmt  'This is a little tricky.  Using math to determine the amt to post.
        ' Invoice Activity
        ElseIf (rsDtlExt![LdgrTransGroupCd] = strLDGRTRANSGROUPCD_INVOICE) Then
            If rsDtlExt![CreditOrDebitCd] = strDEBIT_CD Then
                curInvoiceFeeAmt = curLdgrTransAmt * -1
                curTotInvoiceFeeAmt = curTotInvoiceFeeAmt + (curLdgrTransAmt * -1)
            Else
                curInvoiceFeeAmt = curLdgrTransAmt
                curTotInvoiceFeeAmt = curTotInvoiceFeeAmt + (curLdgrTransAmt)
            End If
            curLdgrEntryFeeAmt = curInvoiceFeeAmt * -1 'Reverse the fee for entry in the Ledger.
        Else
            MsgBox "Unable to derive additional Transaction detail data for the current claim.  Unexpected condition encountered.  Please contact application support.  See " & strProcName, vbOKOnly + vbCritical
            Exit Sub
        End If
        
        ' Assign the current Collection Ledger Entry Amt (enter the signed value to indicate over/under payment activity)
        rsDtlExt![PrincRecovOrPaidAmt] = curPrincRecovOrPaidAmt
        
        ' Assign the current Fee Entry Amt (Invoice fees will reverse the Collection Fee (or AR/AP for Prepay) forecasted.)
        rsDtlExt![LdgrEntryFeeAmt] = curLdgrEntryFeeAmt

        ' Assign the current Collection Balance Amt
        If (rsDtlExt![Adj_ReviewType] = strREVIEWTYPE_PRP) Then
            rsDtlExt![CollBalanceAmt] = 0   'No Collections for Prepayment Reviews.  The Claim has not yet been reimbursed.
        Else
            rsDtlExt![CollBalanceAmt] = curNetAROrAPAmt - curTotPrincRecovOrPaidAmt
        End If
         
        ' Assign the current Fee Balance Amt
        If (rsDtlExt![Adj_ReviewType] = strREVIEWTYPE_PRP) Then
            ' Final Prepayment Fee balance is based on the AR/AP only.
            rsDtlExt![CollFeeBalanceAmt] = curNetAROrAPFeeAmt - curTotInvoiceFeeAmt
        Else
            ' Otherwise, based on Collections.
            rsDtlExt![CollFeeBalanceAmt] = curTotPrincRecovOrPaidFeeAmt - curTotInvoiceFeeAmt
        End If
        
            
        rsDtlExt.MoveNext
    Loop
    
Exit_DeriveLdgrDtlInfo:
    Exit Sub
Err_DeriveLdgrDtlInfo:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    Resume Exit_DeriveLdgrDtlInfo
End Sub


Private Sub cmdRefresh_Click()
On Error GoTo Err_cmdRefresh_Click
    Me.HdrClaimNumList = ""
    Me.LastHdrClaimNumList = ""
    Me.RefreshData
    
Exit_cmdRefresh_Click:
    Exit Sub

Err_cmdRefresh_Click:
    MsgBox Err.Description
    Resume Exit_cmdRefresh_Click
    
End Sub
