Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mstrSearchType As String
Private mstrSearchTable As String
Private ColReSize3 As clsAutoSizeColumns


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
    Dim strError As String
    

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
    On Error GoTo ErrHandler
    
    Dim strErrMsg As String
    
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
    If Me.cboSearchBy = "BeneBirthDt" Then
        Me.lstClaims.RowSource = " SELECT * from " & Me.SearchTable & " WHERE " & Me.cboSearchBy & " = #" & Me.txtSearchFor & "#"
    Else
        Me.lstClaims.RowSource = " SELECT * from " & Me.SearchTable & " WHERE " & Me.cboSearchBy & " LIKE '" & Me.txtSearchFor & "*'"
    End If
    
    'added last minute to exclude error messag 04/24/08 purpose if recordset is empty then there was a problem with
    ' the report vs just an empty dataset
    If Me.lstClaims.RecordSet Is Nothing Or Nz(Me.lstClaims.ItemData(1), "") = "" Then
        strErrMsg = "No Records Returned for this query. "
        'GoTo Error_encountered
    End If
    
    
'    Call ListControlProps2(Me.lstClaims)
    
       'resizing of the listbox.  max columns is 58 columns to be reformatted.  Code from Ron D 11/20/12
    Set ColReSize3 = New clsAutoSizeColumns
    ColReSize3.SetControl Me.lstClaims
    'don't resize if lstclaims is null
    On Error Resume Next
    If Me.lstClaims.ListCount - 1 > 0 Then
        ColReSize3.AutoSize
    End If
    

Exit Sub
    
    
    

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
    
    strParameterString = ""
    
    strParent = Me.Name
    
    strParameter = Nz(DLookup("Parameter", "GENERAL_Navigate", "SearchType = '" & Me.SearchType & "' and ActionName = 'dblClick' and parentform = '" & strParent & "'"), "")
    arrParameters = Split(strParameter, "|")
    
    If UBound(arrParameters) > 0 Then
        For intI = 0 To UBound(arrParameters)
           strParameterString = strParameterString & Me.RecordSet(arrParameters(intI)) & "|"
        Next intI
    Else
          strParameterString = strParameterString & Me.lstClaims.Column(GetColumnPosition(Me.lstClaims, strParameter))
    End If
    
    If strParameter <> "" Then
        Navigate strParent, Me.SearchType, "DblClick", strParameterString
    End If

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
