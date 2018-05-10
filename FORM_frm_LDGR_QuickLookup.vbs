Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_LDGR_QuickLookup
' Author:      Barbara Dyroff
' Create Date: 2012-07-03
' Description:
'      Select a Claim to display for the Transaction Ledger.
'
' Note:  The generic claim lookup leverages the routines that go against the operational tables.  May want to
'   add routines to do the same against the Ledger Warehouse tables instead to keep it all within the Warehouse.
'
' Modification History:
'   2012-12-19 by Barbara Dyroff to add the resizing of the listbox for the 2010 upgrade.
'
' =============================================

Private mstrSearchType As String
Private mstrSearchTable As String
Private ColReSize3 As clsAutoSizeColumns

Private sfrmMain As Form_frm_LDGR_Main

Property Let CallingForm(frmCallingForm As Form_frm_LDGR_Main)
    If Not (frmCallingForm Is Nothing) Then
        Set sfrmMain = frmCallingForm
    End If
End Property

Property Let SearchType(data As String)
    mstrSearchType = data
End Property

Property Get SearchType() As String
    SearchType = mstrSearchType
End Property

Property Let SearchTable(data As String)
    mstrSearchTable = data
End Property

Property Get SearchTable() As String
    SearchTable = mstrSearchTable
End Property
Public Sub RefreshData()
On Error GoTo ErrHandler
'
    Dim strDefaultField As String
    Dim strError As String

    Me.SearchTable = DLookup("SQLFROM", "GENERAL_SEARCH", "SearchType = '" & mstrSearchType & "'")
    strDefaultField = DLookup("DefaultField", "GENERAL_SEARCH", "SearchType = '" & mstrSearchType & "'")
    
    Me.cboSearchBy.RowSource = ""
    Me.cboSearchBy.RowSource = "SELECT FieldName FROM v_XREF_TableFields WHERE TableName = '" & Me.SearchTable & "' ORDER BY FieldName"
    Me.cboSearchBy.RowSourceType = "Table/Query"
    Me.cboSearchBy = strDefaultField
    
Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub


   
Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    Dim strError As String
    
    If Nz(Me.txtSearchFor, "") = "" Then
        Err.Raise 65000, "cmdSearch_Click", "Search cannot be blank."
    End If
    
    If Nz(Me.cboSearchBy, "") = "" Then
        Err.Raise 65000, "cmdSearch_Click", "Search column not specified."
    End If
    
    'Code to initialize the list
    Set Me.lstClaims.RecordSet = Nothing
    Me.lstClaims.RowSource = vbNullString
    'Set our listbox columns to be the same as letter_selection_temp
    Me.lstClaims.ColumnCount = CurrentDb.TableDefs(Me.SearchTable).Fields.Count
    'Set the data in the list equal to what is in the Selection table with any additional filters applied
    Me.lstClaims.RowSource = " SELECT * from " & Me.SearchTable & " WHERE " & Me.cboSearchBy & " LIKE '" & Me.txtSearchFor & "*'"
 
    ' If recordset is empty then there was a problem.
    If Me.lstClaims.RecordSet Is Nothing Or Nz(Me.lstClaims.ItemData(1), "") = "" Then
        MsgBox "The claim could not be found. "
    End If
    
    'resizing of the listbox.  max columns is 58 columns to be reformatted.  Code from Ron D 11/20/12
    Set ColReSize3 = New clsAutoSizeColumns
    ColReSize3.SetControl Me.lstClaims
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstClaims.ListCount - 1 > 0 Then
        ColReSize3.AutoSize
    End If

Exit Sub
ErrHandler:
    MsgBox "Search Error - " & Err.Description, vbOKOnly + vbCritical, Err.Source
End Sub
Private Sub Form_Resize()
    ResizeControls Me.Form
End Sub

Private Sub lstClaims_DblClick(Cancel As Integer)

On Error GoTo ErrHandler
        
    Dim strParameter As String
    Dim strParameterString As String
    
    Dim strError As String
    Dim strParent As String
    Dim arrParameters() As String
    Dim intI As Integer
    Dim strCnlyClaimNum As String

    strCnlyClaimNum = "" + lstClaims.Column(0)
    
    sfrmMain.HdrClaimNumList = "('" & strCnlyClaimNum & "')"

    DoCmd.Close acForm, "frm_LDGR_QuickLookup", acSaveNo

Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub
Private Sub lstClaims_KeyPress(KeyAscii As Integer)

On Error GoTo ErrHandler
    Dim strError As String

    If KeyAscii = 13 Then
        cmdSearch_Click
    End If
    
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub
