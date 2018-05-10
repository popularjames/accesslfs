Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_AR_SETUP_Hdr_Error_Filter_Popup
' Author:      Barbara Dyroff
' Create Date: 2014-08-19
' Description:
'   Prompt the user to provide filter for the AR Setup Error records.
'
' Input:
'   frmARSetup      Assign current instance to pass back the filter string.
'
' Output:
'   ARSetupErrorFilter   Calls Form_frm_AR_SETUP ARSetupErrorFilter to return the Filter string.
'
' Modification History:
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
    
    Dim strFilter As String
    Dim bolValidEntriesInd As Boolean
    
    'Validate the entries.
    bolValidEntriesInd = False
    ValidateEntries bolValidEntriesInd
    If bolValidEntriesInd = False Then
        GoTo Exit_Apply_Click
    End If
    
    'Init Filter
    strFilter = ""
    
    'Claim Provider
    ' Add option to lookup Provider later.
    If Me.txtProvNum <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + " ProvNum = '" & Me.txtProvNum + "' "
    End If

    'AR/AP Start Date
    If Me.txtARAdjClosedStartDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "Adj_ClosedDt >= #" & Me.txtARAdjClosedStartDt & "# "
    End If

    'AR/AP End Date
    If Me.txtARAdjClosedEndDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "Adj_ClosedDt <= #" & Me.txtARAdjClosedEndDt & "# "
    End If
    
    'File Start Date
    If Me.txtFileStartDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "FileDate >= #" & Me.txtFileStartDt & "# "
    End If

    'File End Date
    If Me.txtFileEndDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "FileDate <= #" & Me.txtFileEndDt & "# "
    End If

    'Batch Start Date
    If Me.txtBatchPrcsStartDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "BatchPrcsDt >= #" & Me.txtBatchPrcsStartDt & "# "
    End If

    'Batch End Date
    If Me.txtBatchPrcsEndDt <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + "BatchPrcsDt <= #" & Me.txtBatchPrcsEndDt & "# "
    End If

    'CnlyClaimNum
    ' Add option to lookup CnlyClaimNum later.
    If Me.txtCnlyClaimNum <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + " CnlyClaimNum = '" & Me.txtCnlyClaimNum + "' "
    End If

    'AR Original ICN
    If Me.txtARICN <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + " ARICN = '" & Me.txtARICN + "' "
    End If

    'AR Adjusted ICN
    If Me.txtARICNAdjTo <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + " ARICN_Adj_To = '" & Me.txtARICNAdjTo + "' "
    End If

    'AR Num
    If Me.txtARNum <> "" Then
        If strFilter <> "" Then strFilter = strFilter & "AND "
        strFilter = strFilter + " ARNum = '" & Me.txtARNum + "' "
    End If

    frmARSetup.ARSetupErrorFilter = strFilter

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
    Dim ctl As Control

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
