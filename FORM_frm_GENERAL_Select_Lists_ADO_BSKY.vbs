Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




'' Last Modified: 04/24/2015
''
'' ############################################################
'' ############################################################
'' ############################################################
''  DESCRIPTION:
''  =====================================
''
''
''  TODO:
''  =====================================
''  - Lots, clean up, better commenting, etc..
''
''  HISTORY:
''  =====================================
''  - 05/13/2013  - Created
''
'' AUTHOR
''  =====================================
'' Kevin Dearing
''
''
'' ############################################################
'' ############################################################
'' ############################################################
'' ############################################################



Private Const cs_LEVEL_SOURCE_SQL As String = "SELECT * FROM BOLD_LETTER_Automation_Req_Xref_Levels "

Private csMouseX As Single  ' to track which item was double clicked
Private csMouseY As Single
Private coRs As ADODB.RecordSet
Private csIdColumnName As String
Private csDisplayColumnName As String



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property

Public Property Get IdColumnName() As String
    IdColumnName = csIdColumnName
End Property
Public Property Let IdColumnName(sIdColumnName As String)
    csIdColumnName = sIdColumnName
End Property


Public Property Get MainCaption() As String
    MainCaption = Me.lblCaption.Caption
End Property
Public Property Let MainCaption(sCaption As String)
    Me.lblCaption.Caption = sCaption
End Property

Public Property Get DisplayColumnName() As String
    DisplayColumnName = csDisplayColumnName
End Property
Public Property Let DisplayColumnName(sDisplayColumnName As String)
    csDisplayColumnName = sDisplayColumnName
End Property


Public Function GetSelectedKeysCollection() As Collection
On Error GoTo Block_Err
Dim strProcName As String
Dim oColReturn As Collection
'Dim oLV As Object
Dim oLV As ListView
'Dim oLITm As Object
Dim oLItm As ListItem

    strProcName = ClassName & ".GetSelectedKeysCollection"
    Set oColReturn = New Collection
    
    Set oLV = Me.lstvSelected
    For Each oLItm In oLV.ListItems
        If oLItm.Selected = True Then
            oColReturn.Add CStr(oLItm.Key)
        End If
    Next
    
Block_Exit:
    Set oLItm = Nothing
    Set oLV = Nothing
    Set GetSelectedKeysCollection = oColReturn
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function



Public Function InitData(oRs As ADODB.RecordSet, Optional strIdColName As String, Optional strDisplayColumnName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim sKey As String
Dim sDisplay As String

    strProcName = ClassName & ".InitData"
    
    If Not coRs Is Nothing Then
        If coRs.State = adStateOpen Then coRs.Close
        Set coRs = Nothing
    End If
    
    If oRs Is Nothing Then
        Stop
        GoTo Block_Exit
    End If
    
    If oRs.Fields.Count < 1 Then
        ' nothing to do.
        GoTo Block_Exit
    End If

    Set coRs = oRs

        ' Figure out our ID and Display columns
    If strIdColName = "" Then
        If isField(oRs, strIdColName) = True Then
            sKey = strIdColName
        Else            ' Use the first field in the recordset
            sKey = oRs.Fields(0).Name
        End If
    Else            ' Use the first field in the recordset
        sKey = oRs.Fields(0).Name
    End If
    
    If strDisplayColumnName = "" Then
        If isField(oRs, strDisplayColumnName) = True Then
            sDisplay = strDisplayColumnName
        Else            ' Use the second field in the recordset
            If oRs.Fields.Count > 1 Then
                sDisplay = oRs.Fields(1).Name
            Else
                sDisplay = oRs.Fields(0).Name
            End If
        End If
    Else            ' Use the second field in the recordset
        If oRs.Fields.Count > 1 Then
            sDisplay = oRs.Fields(1).Name
        Else
            sDisplay = oRs.Fields(0).Name
        End If
    End If

        ' set our globals:
    Me.IdColumnName = sKey
    Me.DisplayColumnName = sDisplay


    Call RefreshData

    InitData = True
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub cmdMoveToSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".cmdMoveToSelected_Click"
    
    Call MoveSelectedItems(Me.lstvUnSelected, Me.lstvSelected)

    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub



Private Sub MoveSelectedItems(oFromLv As Object, oToLv As Object)
On Error GoTo Block_Err
Dim strProcName As String

'Dim oLItemFrom As ListItem
'Dim oLItemTo As ListItem
Dim oLItemFrom As Object
Dim oLItemTo As Object

Dim iCol As Integer
Dim iLItemIdx As Integer

    strProcName = ClassName & ".MoveSelectedItems"
    

    For iLItemIdx = oFromLv.ListItems.Count To 1 Step -1
        Set oLItemFrom = oFromLv.ListItems(iLItemIdx)
        'If oLItemFrom.Checked = True Then
        If oLItemFrom.Selected = True Then
            Set oLItemTo = oToLv.ListItems.Add(, , oLItemFrom.Text)
            For iCol = 1 To oFromLv.ColumnHeaders.Count - 1
                oLItemTo.SubItems(iCol) = oLItemFrom.SubItems(iCol)
            Next
            oFromLv.ListItems.Remove iLItemIdx
        End If
    Next

    '' move it over!
    
Block_Exit:
    Set oLItemTo = Nothing
    Set oLItemFrom = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdMoveToUnSelected_Click()
On Error GoTo Block_Err
Dim strProcName As String


    strProcName = ClassName & ".cmdMoveToUnSelected_Click"
    
    Call MoveSelectedItems(Me.lstvSelected, Me.lstvUnSelected)

    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdRefresh_Click()
    If MsgBox("Are you sure you wish to refresh? You will loose all of your selections?", vbOKCancel, "Refresh?") = vbCancel Then
        Exit Sub
    End If
    Call RefreshData
End Sub


Private Sub Form_Load()
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
    
    
    
'    Call RefreshData
    
Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub


Private Sub lstvSelected_DblClick()
    Call SelectItemClickedOver(Me.lstvSelected)
    Call cmdMoveToUnSelected_Click
End Sub



Private Sub lstvSelected_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
    csMouseX = X
    csMouseY = Y
End Sub

Private Sub lstvUnSelected_DblClick()
    Call SelectItemClickedOver(Me.lstvUnSelected)
    Call cmdMoveToSelected_Click
End Sub


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'Dim oAdo As clsADO
'Dim oCn As ADODB.Connection
'Dim oRs As ADODB.Recordset

    strProcName = ClassName & ".RefreshData"
    
'    Set oCn = New ADODB.Connection
'    With oCn
'        .ConnectionString = CurrentProject.AccessConnection
'        .CursorLocation = adUseNone
'        .Open
'    End With
'
'    Set oRs = oCn.Execute(cs_LEVEL_SOURCE_SQL)
'    Call LoadObjectsLV(oRs, Me.lstvUnSelected, Me.lstvSelected)
    Call LoadObjectsLV(Me.lstvUnSelected, Me.lstvSelected)

    
Block_Exit:
'    If Not oRs Is Nothing Then
'        If oRs.State = adStateOpen Then oRs.Close
'        Set oRs = Nothing
'    End If
'    If Not oCn Is Nothing Then
'        If oCn.State = adStateOpen Then oCn.Close
'        Set oCn = Nothing
'    End If
'    Set oAdo = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'' Ok, this will be a generic function.
'' pass the ADO RS and the listView Object to load
'' If the 'IdColName' parameter is blank, the second field in the RS will be used
'' If the 'DisplayName' parameter is blank, the second field in the RS will be used
''
Private Sub LoadObjectsLV(oLV As Object, oPartnerLv As Object)
'Private Sub LoadObjectsLV(oRs As ADODB.Recordset, oLV As ListView, oPartnerLv As ListView, Optional strIdColName As String, Optional strDisplayName As String)
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLItm As ListItem
Dim oLItm As Object

Dim iColHdr As Integer

Dim oFld As ADODB.Field
Dim sKey As String
Dim sDisplay As String
Dim lWidth As Long
Dim oRs As ADODB.RecordSet
Dim iSubItmIdx As Integer

    strProcName = ClassName & ".LoadObjectsLV"
    
        ' Make sure the list view is set up with the correct amount of columns and such..
        ' Just delete them all and rebuild to make this truely dynamic
    oLV.ListItems.Clear
    oPartnerLv.ListItems.Clear
    
    oLV.ColumnHeaders.Clear
    oPartnerLv.ColumnHeaders.Clear
    
    oLV.View = 3    ' = lvwReport
    oLV.GridLines = True
    oLV.MultiSelect = True
    
    oPartnerLv.View = oLV.View
    oPartnerLv.GridLines = oLV.GridLines
    oPartnerLv.MultiSelect = oLV.MultiSelect
    
    If coRs.recordCount < 1 Then
        GoTo Block_Exit
    End If
    If coRs.Fields.Count < 1 Then        ' nothing to do.
        GoTo Block_Exit
    End If

        ' lets not mess with the main RS
    Set oRs = coRs.Clone

        ' Figure out our ID and Display columns
    sKey = Me.IdColumnName
    sDisplay = Me.DisplayColumnName
    

        ' Now rebuild the column headers:
    For Each oFld In oRs.Fields
        Debug.Print oFld.Name & " = " & oFld.ActualSize
        lWidth = 300 * Len(oFld.Name)
        
        If oFld.Name = sKey Then
            Call oLV.ColumnHeaders.Add(, sKey, sDisplay, lWidth)
            Call oPartnerLv.ColumnHeaders.Add(, sKey, sDisplay, lWidth)
                '        ElseIf oFld.Name = sDisplay Then
                '            ' nothing really..
        ElseIf oFld.Name <> sDisplay Then
            Call oLV.ColumnHeaders.Add(, oFld.Name, oFld.Name, lWidth)
            Call oPartnerLv.ColumnHeaders.Add(, oFld.Name, oFld.Name, lWidth)
        End If
    Next
    
    
        '' Now populate the data
    While Not oRs.EOF
        Set oLItm = oLV.ListItems.Add(, "K" & CStr(Nz(oRs(sKey).Value, "")), CStr(Nz(oRs(sDisplay).Value, "")))
        iSubItmIdx = 1
        For Each oFld In oRs.Fields
            If oFld.Name <> sKey And oFld.Name <> sDisplay Then
                oLItm.SubItems(iSubItmIdx) = CStr(Nz(oRs(oFld.Name).Value, ""))
                iSubItmIdx = iSubItmIdx + 1
            End If
        Next
        oRs.MoveNext
    Wend
   
    
Block_Exit:
    Set oRs = Nothing   ' it's just a clone but if we close it, it's still going to close the main one
    Set oFld = Nothing
    Set oLItm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

''' Legacy version before I went really generic
''' Ok, this will be a generic function.
''' pass the ADO RS and the listView Object to load
''' If the 'IdColName' parameter is blank, the second field in the RS will be used
''' If the 'DisplayName' parameter is blank, the second field in the RS will be used
'''
'Private Sub LoadObjectsLV(oRs As ADODB.Recordset, oLV As Object, oPartnerLv As Object, Optional strIdColName As String, Optional strDisplayName As String)
''Private Sub LoadObjectsLV(oRs As ADODB.Recordset, oLV As ListView, oPartnerLv As ListView, Optional strIdColName As String, Optional strDisplayName As String)
'On Error GoTo Block_Err
'Dim strProcName As String
''Dim oLItm As ListItem
'Dim oLITm As Object
'
'Dim iColHdr As Integer
'
'Dim oFld As ADODB.Field
'Dim sKey As String
'Dim sDisplay As String
'Dim lWidth As Long
'
'Dim iSubItmIdx As Integer
'
'    strProcName = ClassName & ".LoadObjectsLV"
'
'        ' Make sure the list view is set up with the correct amount of columns and such..
'        ' Just delete them all and rebuild to make this truely dynamic
'    oLV.ListItems.Clear
'    oPartnerLv.ListItems.Clear
'
'    oLV.ColumnHeaders.Clear
'    oPartnerLv.ColumnHeaders.Clear
'
'    oLV.View = 3    ' = lvwReport
'    oLV.GridLines = True
'    oLV.MultiSelect = True
'
'    oPartnerLv.View = oLV.View
'    oPartnerLv.GridLines = oLV.GridLines
'    oPartnerLv.MultiSelect = oLV.MultiSelect
'
'
'
'    If oRs.Fields.Count < 1 Then
'        ' nothing to do.
'        GoTo Block_Exit
'    End If
'
'        ' Figure out our ID and Display columns
'    If strIdColName = "" Then
'         If isField(oRs, strIdColName) = True Then
'             sKey = strIdColName
'         Else            ' Use the first field in the recordset
'             sKey = oRs.Fields(0).Name
'         End If
'     Else            ' Use the first field in the recordset
'         sKey = oRs.Fields(0).Name
'     End If
'
'     If strDisplayName = "" Then
'         If isField(oRs, strDisplayName) = True Then
'             sDisplay = strDisplayName
'         Else            ' Use the second field in the recordset
'            If oRs.Fields.Count > 1 Then
'                sDisplay = oRs.Fields(1).Name
'            Else
'                sDisplay = oRs.Fields(0).Name
'            End If
'         End If
'     Else            ' Use the second field in the recordset
'        If oRs.Fields.Count > 1 Then
'            sDisplay = oRs.Fields(1).Name
'        Else
'            sDisplay = oRs.Fields(0).Name
'        End If
'
'     End If
'
'    ' Now rebuild the column headers:
'    For Each oFld In oRs.Fields
'
'        Debug.Print oFld.Name & " = " & oFld.ActualSize
'
''        lWidth = 300 * oFld.ActualSize  ' just guessing at that.. Maybe I should do the number of characters in the field name
'        lWidth = 300 * Len(oFld.Name)
'
'        If oFld.Name = sKey Then
'            Call oLV.ColumnHeaders.Add(, sKey, sDisplay, lWidth)
'            Call oPartnerLv.ColumnHeaders.Add(, sKey, sDisplay, lWidth)
'        ElseIf oFld.Name = sDisplay Then
'            ' nothing really..
'        Else
'            Call oLV.ColumnHeaders.Add(, oFld.Name, oFld.Name, lWidth)
'            Call oPartnerLv.ColumnHeaders.Add(, oFld.Name, oFld.Name, lWidth)
'        End If
'    Next
'
'
'
'    '' Now populate the data
'    While Not oRs.EOF
'        Set oLITm = oLV.ListItems.Add(, "K" & CStr(Nz(oRs(sKey).Value, "")), CStr(Nz(oRs(sDisplay).Value, "")))
'        iSubItmIdx = 1
'        For Each oFld In oRs.Fields
'            If oFld.Name <> sKey And oFld.Name <> sDisplay Then
'                oLITm.SubItems(iSubItmIdx) = CStr(Nz(oRs(oFld.Name).Value, ""))
'                iSubItmIdx = iSubItmIdx + 1
'            End If
'        Next
'        oRs.MoveNext
'    Wend
'
'
'Block_Exit:
'    Set oLITm = Nothing
'    Exit Sub
'Block_Err:
'    ReportError Err, strProcName
'    GoTo Block_Exit
'End Sub



Private Function SelectItemClickedOver(oLV As Object) As Boolean
'Private Function SelectItemClickedOver(oLV As ListView) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLItm As ListItem
Dim oLItm As Object

    strProcName = ClassName & ".SelectItemClickedOver"
    
    If oLV.ListItems.Count = 0 Then GoTo Block_Exit
    
    Set oLItm = oLV.HitTest(csMouseX, csMouseY)
'    Set oLItm = oLV.HitTest(csMouseX, csMouseY)
    If Not oLItm Is Nothing Then
        Debug.Print oLItm
    
    Else
        Set oLItm = oLV.HitTest(125, csMouseY)


    End If
    
    If oLItm Is Nothing Then
        ' clicked somewhere that there isn't an item I guess
        GoTo Block_Exit
    End If
        Debug.Print oLItm
    Call UnselectAll(oLV)
    oLV.SelectedItem = oLItm
    
    
Block_Exit:
    Set oLItm = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub lstvUnSelected_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
    csMouseX = X
    csMouseY = Y
End Sub

Private Sub UnselectAll(oLV As Object)
'Private Sub UnselectAll(oLv As ListView)
On Error GoTo Block_Err
Dim strProcName As String
'Dim oLItm As ListItem
Dim oLItm As Object

    strProcName = ClassName & ".UnselectAll"
    
    If oLV.ListItems.Count < 1 Then GoTo Block_Exit
        
    For Each oLItm In oLV.ListItems
        oLItm.Selected = False
    Next
    
Block_Exit:
    Set oLItm = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub
