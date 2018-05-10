Option Compare Database
Option Explicit


Private Const ClassName As String = "mod_GENERAL_Controls"



Public Function GetColumnPosition(Lst As listBox, _
                                  strField As String) As Integer

    Dim iMaxCol As Integer
    Dim i As Integer

    ' set default value
    GetColumnPosition = -1
    
    iMaxCol = Lst.RecordSet.Fields.Count - 1
    For i = 0 To iMaxCol
        If UCase(Lst.Column(i, 0)) = UCase(strField) Then
            GetColumnPosition = i
            Exit For
        End If
    Next i

End Function


Public Function GetListBoxSQL(FormName As String, Optional ProfileID As String) As String
    
    
    
    If ProfileID & "" <> "" Then
        
        GetListBoxSQL = "select TabName, T.RowID from GENERAL_Tabs T INNER JOIN GENERAL_Tabs_Linked_ProfileIDs P ON P.RowID = T.RowID where AccessForm = """ & FormName & """" & _
                        " and P.ProfileID = " & Chr(34) & ProfileID & Chr(34) & " order by tabname"
    Else
        GetListBoxSQL = "select TabName, RowID from GENERAL_Tabs where AccessForm = """ & FormName & """ order by tabname"
    End If
End Function


Public Function GetListBoxRowSQL(RowID As Long, FormName As String) As String
    GetListBoxRowSQL = "select * from GENERAL_Tabs where AccessForm = '" & FormName & "'" & _
                              " and RowID = " & RowID
             
End Function


Public Sub RefreshComboBox(strSQL As String, _
                            cboBox As ComboBox, _
                            Optional varDefaultSelection As Variant = "", _
                            Optional strField As String = "")
    On Error GoTo ErrHandler

    Dim strProviderID As String
    Dim rst As DAO.RecordSet
    Dim db As Database

    cboBox.RowSource = vbNullString

    Set db = CurrentDb
        
    Set rst = db.OpenRecordSet(strSQL, dbOpenDynaset, dbSeeChanges)
    While Not rst.EOF
        cboBox.AddItem (Trim(Nz(rst.Fields(0).Value, ""))) & ";" & Replace(Nz(rst.Fields(1).Value, ""), ",", "")
        
        If strField <> "" Then
            If Trim(CStr(Nz(rst.Fields(strField).Value, ""))) = Trim(CStr(Nz(varDefaultSelection, ""))) Then
                cboBox = Trim(CStr(varDefaultSelection))
            End If
        End If
        
        rst.MoveNext
    Wend
    
    Set rst = Nothing
   
ExitNow:
    Set rst = Nothing
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshComboBox"
    GoTo ExitNow
End Sub


Public Sub RefreshComboBoxADO(strSQL As String, _
        cboBox As ComboBox, _
        Optional varDefaultSelection As Variant = "", _
        Optional strField As String = "", Optional strTableForConnection As String = "v_Data_Database")
On Error GoTo Block_Err
Dim strProcName As String
Dim strProviderID As String
Dim oAdo As clsADO
Dim oRs As ADODB.RecordSet

    strProcName = "mod_GENERAL_Controls.RefreshComboBoxADO"

    cboBox.RowSource = vbNullString

    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString(strTableForConnection)
        .SQLTextType = sqltext
        .sqlString = strSQL
        Set oRs = .ExecuteRS
        If .GotData = False Then
            GoTo Block_Exit
        End If
    End With
    cboBox.RowSourceType = "Value List"
        
    While Not oRs.EOF
        cboBox.AddItem (Trim(Nz(oRs.Fields(0).Value, ""))) & ";" & Replace(Nz(oRs.Fields(1).Value, ""), ",", "")
        
        If strField <> "" Then
            If Trim(CStr(Nz(oRs.Fields(strField).Value, ""))) = Trim(CStr(Nz(varDefaultSelection, ""))) Then
                cboBox = Trim(CStr(varDefaultSelection))
            End If
        End If
        
        oRs.MoveNext
    Wend
    
    
   
Block_Exit:
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then oRs.Close
        Set oRs = Nothing
    End If
    Set oAdo = Nothing
    Exit Sub

Block_Err:
    ReportError Err, strProcName
'    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshComboBox"
    GoTo Block_Exit
End Sub

'Damon 06/03/08
'General ListBox Filling
'This adds the functionality of being able to specify a default selected value
Public Sub RefreshListBox(strSQL As String, lstBox As listBox, _
                            Optional varDefaultSelection As Variant = "", _
                            Optional strField As String = "")

    On Error GoTo ErrHandler

    Dim strProviderID As String
    Dim rst As DAO.RecordSet
    Dim db As DAO.Database
    Dim ctr As Long
    Dim strItem As String
    lstBox.RowSource = vbNullString
    Set db = CurrentDb
    Set rst = db.OpenRecordSet(strSQL, dbOpenDynaset, dbSeeChanges)
    
    While Not rst.EOF
        For ctr = 0 To rst.Fields.Count - 1
            strItem = strItem & rst.Fields(ctr).Value & ";"
        Next ctr
        lstBox.AddItem strItem
        strItem = ""
        If strField <> "" Then
            If Trim(CStr(Nz(rst.Fields(strField).Value, ""))) = Trim(CStr(Nz(varDefaultSelection, ""))) Then
                lstBox = Trim(CStr(varDefaultSelection))
            End If
        End If
        rst.MoveNext
    Wend
ExitNow:
    '*On Error Resume Next
    Set db = Nothing
    Set rst = Nothing

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshProjectListing"
    GoTo ExitNow

End Sub


'Damon 06/03/08
'General ListBox Filling
'This adds the functionality of being able to specify a default selected value
Public Sub RefreshListBoxFromRecordset(rst As ADODB.RecordSet, _
                            lstBox As listBox, _
                            Optional varDefaultSelection As Variant = "", _
                            Optional strField As String = "")

    On Error GoTo ErrHandler

    Dim ctr As Long
    Dim strItem As String
    
    lstBox.RowSource = vbNullString
    While Not rst.EOF
        For ctr = 0 To rst.Fields.Count - 1
            strItem = strItem & Trim(Nz(rst.Fields(ctr).Value, "")) & ";"
        Next ctr
        lstBox.AddItem strItem
        strItem = ""
        If strField <> "" Then
            If Trim(CStr(Nz(rst.Fields(strField).Value, ""))) = Trim(CStr(Nz(varDefaultSelection, ""))) Then
                lstBox = Trim(CStr(varDefaultSelection))
            End If
        End If
        rst.MoveNext
    Wend
ExitNow:

rst.MoveFirst


Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshProjectListing"
    GoTo ExitNow

End Sub

Public Sub RefreshComboBoxFromRecordset(rst As ADODB.RecordSet, _
                            cboBox As ComboBox, _
                            Optional varDefaultSelection As Variant = "", _
                            Optional strField As String = "")
    On Error GoTo ErrHandler
    Dim ctr As Long
    Dim strItem As String

    cboBox.RowSource = vbNullString
    cboBox.RowSourceType = "Value List"

    While Not rst.EOF
        'cboBox.AddItem (Trim(Nz(rst.Fields(0).Value, ""))) & ";" & Replace(Nz(rst.Fields(1).Value, ""), ",", "")
        
        For ctr = 0 To rst.Fields.Count - 1
            strItem = strItem & Trim(Replace(Nz(rst.Fields(ctr).Value, ""), ";", ",")) & ";"
            strItem = Replace(strItem, ",", " -")
        Next ctr
        
        'strItem = Replace(rst.Fields(ctr).Value, ";", ",")
        
        cboBox.AddItem strItem
        strItem = ""
        
        If strField <> "" Then
            If Trim(CStr(Nz(rst.Fields(strField).Value, ""))) = Trim(CStr(Nz(varDefaultSelection, ""))) Then
                cboBox = Trim(CStr(varDefaultSelection))
            End If
        End If
        
        rst.MoveNext
    Wend
    
    Set rst = Nothing
   
ExitNow:
    Set rst = Nothing
Exit Sub

ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "RefreshComboBox"
    GoTo ExitNow
End Sub




Public Function IsControl(oForm As Form, sName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control

    strProcName = ClassName & ".IsControl"
    
    For Each oCtl In oForm.Controls
        If UCase(oCtl.Name) = UCase(sName) Then
            IsControl = True
            Exit For
        End If
    Next
    
Block_Exit:
    Set oCtl = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function