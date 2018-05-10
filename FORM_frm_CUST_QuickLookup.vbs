Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mstrSearchType As String
Private mstrSearchTable As String
Private mCUSTService As clsCUSTSERVICE
Private miEventID As String

Property Let CustService(data As clsCUSTSERVICE)
    Set mCUSTService = data
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

    Dim strDefaultField As String

    'Me.SearchType = "ProvHdr"
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
Dim strSQL As String

    On Error GoTo ErrHandler
    
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
    'Alex C 2/12/2012 - added an exclusion to the SQL to prevent selecting claims already associated with the event
    If Me.SearchTable = "v_CUST_EVENT_Related_Claims" Then
        miEventID = Me.txtSearchFor
    Else
        miEventID = mCUSTService.EventID
    End If
    
    If Me.SearchTable = "v_CUST_EVENT_Related_Claims" Then
       strSQL = " SELECT CnlyClaimNum,ICN,Can,ProvNum from " & Me.SearchTable & " WHERE " & Me.cboSearchBy & " LIKE '" & Me.txtSearchFor & "'"
       strSQL = strSQL + " and CnlyClaimNum in (select CnlyClaimNum from CUST_Event_Related_Claim where EventID = " & miEventID & ")"
       Me.lstClaims.RowSource = strSQL
    Else
       strSQL = " SELECT CnlyClaimNum,ICN,Can,ProvNum from " & Me.SearchTable & " WHERE " & Me.cboSearchBy & " LIKE '" & Me.txtSearchFor & "*'"
       strSQL = strSQL + " and CnlyClaimNum not in (select CnlyClaimNum from CUST_Event_Related_Claim where EventID = " & miEventID & ")"
       Me.lstClaims.RowSource = strSQL
    End If
    
    'added last minute to exclude error messag 04/24/08 purpose if recordset is empty then there was a problem with
    ' the report vs just an empty dataset
    If Me.lstClaims.RecordSet Is Nothing Or Nz(Me.lstClaims.ItemData(1), "") = "" Then
        strErrMsg = "No Records Returned for this query. "
        'GoTo Error_encountered
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
    
If GblParentEvent = "Event" Then
    intEventID = lstClaims.Column(0)
    LaunchNewCustClaimEvent lstClaims.Column(7)
Else
    strCnlyClaimNum = " "
    strCnlyClaimNum = lstClaims.Column(0)
    
    'Send the claim to the class to add it to the list of related claims
    mCUSTService.AddClaimToEvent (strCnlyClaimNum)
    
End If
    DoCmd.Close acForm, "frm_CUST_QuickLookup", acSaveNo
    
    Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub
Private Sub lstClaims_KeyPress(KeyAscii As Integer)

On Error GoTo ErrHandler

If KeyAscii = 13 Then
    cmdSearch_Click
End If
Exit Sub

ErrHandler:
    strError = Err.Description
    MsgBox "Error: " & Err.Number & " (" & strError & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, "frm_General_Search : Navigate"

End Sub
