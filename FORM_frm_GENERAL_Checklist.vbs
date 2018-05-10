Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private coRs As ADODB.RecordSet

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property



Public Property Get ListRecordset() As ADODB.RecordSet
    Set ListRecordset = coRs
End Property
Public Property Let ListRecordset(oRs As ADODB.RecordSet)

Dim oFld As ADODB.Field
'Dim oLItem As ListItem
'Dim oLView As ListView

Dim oLItem As Object
Dim oLView As Object
Dim iRow As Integer
Dim i As Integer
Dim oRsCpy As ADODB.RecordSet

    Set coRs = oRs

    Call SetupListView(oRs)
    Set oRsCpy = oRs.Clone
    
    oRsCpy.MoveFirst
    Set oLView = Me.lvwCheckList
    
    While Not oRsCpy.EOF
        iRow = iRow + 1
        Set oLItem = oLView.ListItems.Add(, CStr("" & iRow) & " " & CStr("" & oRsCpy.Fields(0).Value), CStr("" & oRsCpy.Fields(0).Value))
        
        For Each oFld In oRsCpy.Fields
            If oFld.Name <> oRs.Fields(0).Name Then
                i = i + 1
                oLItem.SubItems(i) = Nz(oFld.Value, "")
            End If
            
        Next
        i = 0
        oRsCpy.MoveNext
    Wend
    Set oRsCpy = Nothing
End Property



Public Function SetupListView(oRs As ADODB.RecordSet)
On Error GoTo Block_Err
Dim strProcName As String
Dim oFld As ADODB.Field

'Dim oLItem As ListItem
'Dim oLView As ListView

'Dim oLItem As Object
Dim oLView As Object

    strProcName = ClassName & ".SetupListView"
    
    Set oLView = Me.lvwCheckList
    oLView.ListItems.Clear
    oLView.ColumnHeaders.Clear
    
    
    For Each oFld In oRs.Fields
        oLView.ColumnHeaders.Add , oFld.Name, oFld.Name
    Next
    
    Call ResetColumnWidths
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function ResetColumnWidths() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oFld As ADODB.Field
Dim oRs As ADODB.RecordSet
Dim iCurNum As Integer
Dim oColumn As Object
' ColumnHeader??
'Dim oLView As ListView
Dim oLView As Object
Dim bytType As Byte
Dim dblWidth As Double

    strProcName = ClassName & ".ResetColumnWidths"
    
    Set oRs = Me.ListRecordset
    
    Set oLView = Me.lvwCheckList
    
    
    For Each oFld In oRs.Fields
        iCurNum = iCurNum + 1
        
        Set oColumn = oLView.ColumnHeaders(iCurNum)
        bytType = AdoTypeToDaoType(oFld)
        
        '' hmm.. seems like there are about 12 characters in about an inch
        
        
'        dblWidth = GetFieldWidth(bytType, IIf(bytType = 10, oFld.DefinedSize, 0), RecordSource, oFld.Name, 0)
    
'Stop

        If left(LCase(oFld.Name), 4) = "note" Then
            dblWidth = Len(CStr("" & oFld.Name)) + 20
        Else
            dblWidth = Len(CStr("" & oFld.Name)) '+ 12
        End If
'        dblWidth = Len(CStr("" & oFld.Name)) '+ 12
'        oColumn.width = oFld.DefinedSize * 1440
        oColumn.Width = (dblWidth * 1440) / 8
        
        
    Next oFld

    

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub lvwCheckList_ColumnClick(ByVal ColumnHeader As Object)
Dim oLView As Object
Dim oLItem As Object
    
    Set oLView = Me.lvwCheckList
    oLView.Sorted = False
    oLView.SortKey = ColumnHeader.index - 1
    oLView.SortOrder = IIf(oLView.SortOrder = lvwDescending, lvwAscending, lvwDescending)
    oLView.Sorted = True
End Sub
