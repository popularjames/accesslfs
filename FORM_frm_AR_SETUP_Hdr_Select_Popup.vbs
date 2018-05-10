Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AR_SETUP_Hdr_Select_Popup
' Author:      Barbara Dyroff
' Create Date: 2014-07-28
' Description:
'   Prompt the user to provide selection criteria for the AR Setup to display.
'
' Input:
'   frmARSetup      Assign current instance to pass back the SQL selection string.
'
' Output:
'   ARSetupSelect   Calls Form_frm_AR_SETUP_Hdr ARSetupSelect to return the selection string.
'
' Modification History:
'   2014-12-19 by BJD to add selection based on the CnlyClaimARID.  This is added to facilitate the AR Orphan maintenance.
'       Also, for the ProvNum select based on either the Claim Prov Num or the AR ProvNum (this will pick up AR orphans for the
'       Provider as well - at least for Part A).  Keep in mind that Part B ProvNum is often different than the Claim ProvNum
'  2015-01-30 BY BJD to add selection of orphans only.
'
' =============================================


Private frmARSetup As Form_frm_AR_SETUP

Public Property Set FormARSetup(ByRef data As Form_frm_AR_SETUP)
    Set frmARSetup = data
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Private Sub cmdApply_Click()
    On Error GoTo Err_Apply_Click
    
    Dim strWhere As String
    Dim ctlSelectList As Control
    Dim strClmStatusList As String
    Dim varListItem As Variant
    Dim bolValidEntriesInd As Boolean
    
    'Validate the entries.
    bolValidEntriesInd = False
    ValidateEntries bolValidEntriesInd
    If bolValidEntriesInd = False Then
        GoTo Exit_Apply_Click
    End If
    
    'Create the SQL Query.
    frmARSetup.ARSetupSelect = "SELECT AR.* FROM AR_SETUP_Hdr AS AR LEFT JOIN AUDITCLM_HDR AS CLM ON AR.CnlyClaimNum = CLM.CnlyClaimNum "
    strWhere = ""
    
    'Provider
    ' Add option to lookup Provider later.
'    If Me.txtProvNum <> "" Then
'        If strWhere <> "" Then strWhere = strWhere & "AND "
'        strWhere = strWhere + " CLM.ProvNum = '" & Me.txtProvNum + "' "
'    End If
    '(Select for either the Claim Provider or the AR Provider (this will pick up orphans for the provider as well).
    If Me.txtProvNum <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " ((CLM.ProvNum = '" & Me.txtProvNum + "') OR (AR.ProvNum = '" & Me.txtProvNum + "')) "
    End If
    
    'Claim Status List
    If Me.lstClmStatus.ItemsSelected.Count <> 0 Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strClmStatusList = ""
        
        Set ctlSelectList = Me!lstClmStatus
        For Each varListItem In ctlSelectList.ItemsSelected
           If strClmStatusList <> "" Then strClmStatusList = strClmStatusList & ", "
           strClmStatusList = strClmStatusList + "'" + ctlSelectList.ItemData(varListItem) + "'"
        Next varListItem

        strWhere = strWhere + "CLM.ClmStatus IN (" + strClmStatusList + ") "
    End If
    
    'AR/AP Start Date
    If Me.txtARAdjClosedStartDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.Adj_ClosedDt >= #" & Me.txtARAdjClosedStartDt & "# "
    End If
    
    'AR/AP End Date
    If Me.txtARAdjClosedEndDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.Adj_ClosedDt <= #" & Me.txtARAdjClosedEndDt & "# "
    End If
    
    'Claim Last Demand Start Dt
    If Me.txtClmAdjClosedStartDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "CLM.Adj_ClosedDt >= #" & Me.txtClmAdjClosedStartDt & "# "
    End If
    
    'Claim Last Demand End Dt
    If Me.txtClmAdjClosedEndDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "CLM.Adj_ClosedDt <= #" & Me.txtClmAdjClosedEndDt & "# "
    End If
    
    'File Start Date
    If Me.txtFileStartDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.FileDate >= #" & Me.txtFileStartDt & "# "
    End If
    
    'File End Date
    If Me.txtFileEndDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.FileDate <= #" & Me.txtFileEndDt & "# "
    End If
    
    'Batch Start Date
    If Me.txtBatchPrcsStartDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.BatchPrcsDt >= #" & Me.txtBatchPrcsStartDt & "# "
    End If
    
    'Batch End Date
    If Me.txtBatchPrcsEndDt <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + "AR.BatchPrcsDt <= #" & Me.txtBatchPrcsEndDt & "# "
    End If
    
    'Payer
    ' Defer.  Add later
    
    'CnlyClaimARID
    If Me.txtCnlyClaimARID <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.CnlyClaimARID = '" & Me.txtCnlyClaimARID + "' "
    End If
    
    'CnlyClaimNum
    ' Add option to lookup CnlyClaimNum later.
    If Me.txtCnlyClaimNum <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " CLM.CnlyClaimNum = '" & Me.txtCnlyClaimNum + "' "
    End If
    
    'Claim ICN
    If Me.txtClmICN <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " CLM.ICN = '" & Me.txtClmICN + "' "
    End If
    
    'AR Original ICN
    If Me.txtARICN <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.ARICN = '" & Me.txtARICN + "' "
    End If
    
    'AR Adjusted ICN
    If Me.txtARICNAdjTo <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.ARICN_Adj_To = '" & Me.txtARICNAdjTo + "' "
    End If
    
    'AR Num
    If Me.txtARNum <> "" Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.ARNum = '" & Me.txtARNum + "' "
    End If
    
    ' Include Collection Additional records?
    If OptCollAddl.Value = 0 Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.AROrAPStatusCd <> 'COLL-ADDL' "
    End If
    
    ' Select only orphans (no CnlyClaimNum assigned)?
    If OptOrphanOnlyInd.Value = 1 Then
        If strWhere <> "" Then strWhere = strWhere & "AND "
        strWhere = strWhere + " AR.CnlyClaimNum IS NULL "
    End If
    
    'Add the Where Clause
    If strWhere <> "" Then
        frmARSetup.ARSetupSelect = frmARSetup.ARSetupSelect + "WHERE " + strWhere
    End If
    
    'Add the Order By clause.
    frmARSetup.ARSetupSelect = frmARSetup.ARSetupSelect + "ORDER BY AR.BatchPrcsDt DESC "

    DoCmd.Close acForm, Me.Name
    
Exit_Apply_Click:
    Exit Sub
    
Err_Apply_Click:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".CmdApply_Click"
    Resume Exit_Apply_Click
    
End Sub


Private Sub CmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


' Clear all the current selection criteria.
Private Sub cmdClear_Click()
    On Error GoTo Err_Clear_Click
    Dim varItem As Variant
    Dim ctl As Control
    
    'Clear the Claim Status List
    For Each varItem In Me!lstClmStatus.ItemsSelected
        Me!lstClmStatus.Selected(varItem) = False
    Next

    ' Clear the text box controls
    For Each ctl In Me.Controls
        Select Case ctl.ControlType
            Case acTextBox
                ctl = ""
        End Select
    Next ctl
    
Exit_Clear_Click:
    Exit Sub

Err_Clear_Click:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & ClassName & ".CmdClear_Click"
    
    Resume Exit_Clear_Click

End Sub

' Validate the entries.
Private Sub ValidateEntries(ByRef bolValidInd As Boolean)
    On Error GoTo Err_ValidateEntries

    Dim strProcName As String

    strProcName = ClassName & ".ValidateEntries"
    
    bolValidInd = False 'Init validation to false.
    
    'Validate the Dates
    
    'AR/AP Start Dt
    If Me.txtARAdjClosedStartDt <> "" And IsDate(Me.txtARAdjClosedStartDt) = False Then
        MsgBox "Invalid date for AR/AP Start Date [" & Me.txtARAdjClosedStartDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtARAdjClosedStartDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'AR/AP End Dt
    If Me.txtARAdjClosedEndDt <> "" And IsDate(Me.txtARAdjClosedEndDt) = False Then
        MsgBox "Invalid date for AR/AP End Date [" & Me.txtARAdjClosedEndDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtARAdjClosedEndDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'Claim (Last Initial Demand) AR/AP Start Dt
    If Me.txtClmAdjClosedStartDt <> "" And IsDate(Me.txtClmAdjClosedStartDt) = False Then
        MsgBox "Invalid date for Claim Last Demand Start Date [" & Me.txtClmAdjClosedStartDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtClmAdjClosedStartDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'Claim (Last Initial Demand) AR/AP End Dt
    If Me.txtClmAdjClosedEndDt <> "" And IsDate(Me.txtClmAdjClosedEndDt) = False Then
        MsgBox "Invalid date for Claim Last Demand End Date [" & Me.txtClmAdjClosedEndDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtClmAdjClosedEndDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'File Start Dt
    If Me.txtFileStartDt <> "" And IsDate(Me.txtFileStartDt) = False Then
        MsgBox "Invalid date for File Start Date [" & Me.txtFileStartDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtFileStartDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'File End Dt
    If Me.txtFileEndDt <> "" And IsDate(Me.txtFileEndDt) = False Then
        MsgBox "Invalid date for File End Date [" & Me.txtFileEndDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtFileEndDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'Batch Process Start Dt
    If Me.txtBatchPrcsStartDt <> "" And IsDate(Me.txtBatchPrcsStartDt) = False Then
        MsgBox "Invalid date for Batch Process Start Date [" & Me.txtBatchPrcsStartDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtBatchPrcsStartDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    'Batch Process End Dt
    If Me.txtBatchPrcsEndDt <> "" And IsDate(Me.txtBatchPrcsEndDt) = False Then
        MsgBox "Invalid date for Batch Process End Date [" & Me.txtBatchPrcsEndDt & "].  Please enter a valid date (Example mm/dd/yyyy).  ", vbOKOnly + vbCritical, "Error Date Entry"
        Me.txtBatchPrcsEndDt = ""
        GoTo Exit_ValidateEntries
    End If
    
    bolValidInd = True  'All valid.
    
Exit_ValidateEntries:
    Exit Sub
Err_ValidateEntries:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error - " & strProcName
    Resume Exit_ValidateEntries
End Sub
