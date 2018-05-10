Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------------------------------------------------
'Author :
'Description:
'Create Date:
'Last Modified:
' PD,Oct 22 2011 - CR # 2572 fix - Erroneous SQL being created while setting up filters for text values with spaces
' PD,Oct 25 2011 - CR # 2578 fix - Filter for “Is Not Null or Blank” generates incorrect SQL
' PD,Oct 25 2011 - CR # 2579 fix - Invoking 'Add Query Criteria' removes data from main grid but leaves data in tabs
'---------------------------------------------------------------------------------------------------------------------------------


Public Event FinishedFilter(ByVal FilterID As Long)
Public Event Cancelled()
Public Event LoadError()
Private genUtils As New CT_ClsGeneralUtilities
Public oFrmScreen As Form_frm_GENERAL_Datasheet_DAO

Private ApplyFilter As Boolean
Private bNew As Boolean
Private bDCUser As Boolean
Private TogValue As Boolean
Private Enabled As Boolean
Private saveEnabled As Boolean
Private saveAsEnabled As Boolean
Private deleteEnabled As Boolean

Public Sub SetParent(frm As Form_frm_CONCEPT_Main)
    Set oFrmScreen = frm
End Sub

Public Sub Initialize()
    Dim strForm As String
    
    bDCUser = False
    
    Me.visible = True
    genUtils.SuspendLayout Me
    
    DoEvents
    'LOOKUP THE VALUES IN THE CURRENT SCREEN
    With oFrmScreen
Stop
    
'        If "" & oFrmScreen.PrimaryCriteria = "" Then
            RaiseEvent LoadError
            MsgBox "The Primary Record Source for this screen has not been defined.", vbInformation
            Exit Sub
'        End If
Stop
'        strForm = "" & oFrmScreen.PrimaryCriteria
    End With
    
    cboField.RowSource = strForm

    'Set FrmScreen.GridForm.Recordset = Nothing 'Commented line for CR # 2579 fix
    bDCUser = isDcUser
    
    TogValue = True
    Me.lstFilters.RowSourceType = "Table/Query"
    UpdateFilterList
    
    genUtils.ResumeLayout Me
    
End Sub

Private Sub CmdAddRow_Click()
On Error GoTo HandleError

Dim sValue As String
Dim sClause As String
Dim sField As String
Dim ValID As Boolean
Dim sDataType As Integer
            
    If Not Enabled Then
        Exit Sub
    End If
    
    ValID = True
    If Nz(Me.cboField, "") = "" Then
        MsgBox "Please select a field.", vbOKOnly + vbExclamation, "Add Row"
        Me.cboField.SetFocus
        ValID = False
    ElseIf Nz(Me.cboOperator, "") = "" Then
        MsgBox "Please select an operator.", vbOKOnly + vbExclamation, "Add Row"
        Me.cboOperator.SetFocus
        ValID = False
        'CR # 2578 fix - Updated the incorrect “Is Not Null or Blank” condition string to “Is Not Null AND Is Not Blank” condition.
        'PD Nov 17 2011, Refactored from if to elseIf to ensure condition was evaluated.
    ElseIf UCase(Me.cboOperator) = "IS NULL" Or UCase(Me.cboOperator) = "IS NULL OR BLANK" Or _
        UCase(Me.cboOperator) = "IS NOT NULL" Or UCase(Me.cboOperator) = "IS NOT NULL AND IS NOT BLANK" Then
        txtValue = ""
    ElseIf Nz(Me.txtValue, "") = "" Then
        If UCase(Me.cboOperator) <> "IS NULL" And UCase(Me.cboOperator) <> "IS NULL OR BLANK" And _
            UCase(Me.cboOperator) <> "IS NOT NULL" And UCase(Me.cboOperator) <> "IS NOT NULL AND IS NOT BLANK" Then 'CR # 2578 fix
            MsgBox "Please enter a value.", vbOKOnly + vbExclamation, "Add Row"
            Me.txtValue.SetFocus
            ValID = False
        End If
    End If
    
    If ValID Then
        sField = Me.cboField.Value
        sDataType = CInt("0" & RetrieveDataType(sField))
        sValue = Nz(txtValue, "")
        If UCase(Me.cboOperator) = "BETWEEN" Or UCase(Me.cboOperator) = "NOT BETWEEN" Then
            ' convert he string to upper temporarily to get determine if there is an AND in the string
            sValue = txtValue
            If InStr(UCase(sValue), " AND ") = 0 Then
                MsgBox "The operators Between and Not Between require the word AND between the 2 choices.", vbOKOnly + vbExclamation, "Value Entry Error"
                txtValue.SetFocus
            Else
                sValue = ParseValue(" AND ", sValue, sField)
                If sValue = "" Then
                    txtValue.SetFocus
                Else
                    sClause = sField & " " & Me.cboOperator & " " & sValue
                End If
            End If
        ElseIf UCase(Me.cboOperator) = "IN" Or UCase(Me.cboOperator) = "NOT IN" Then
            sValue = txtValue
            If InStr(sValue, ",") = 0 Then
                MsgBox "The comma (,) is required to separate the values in the list.", vbOKOnly + vbExclamation, "Value Entry Error"
                txtValue.SetFocus
            Else
                sValue = ParseValue(",", sValue, sField)
                If sValue = "" Then
                    txtValue.SetFocus
                Else
                    sClause = sField & " " & Me.cboOperator & " (" & sValue & ")"
                End If
            End If
        ElseIf UCase(Me.cboOperator) = "IS NULL OR BLANK" Then
            sClause = "(" & sField & " IS NULL)"
            Select Case sDataType
                Case 2 To 7 'Numbers
                    sClause = "(" & sClause & " OR (" & sField & " = 0))"
                Case Is = 10 'text
                    sClause = "(" & sClause & " OR (" & sField & " = ''))"
                Case Is = 15 'bigint
                    sClause = "(" & sClause & " OR (" & sField & " = 0))"
                Case Is = 18 'char
                    sClause = "(" & sClause & " OR (" & sField & " = ''))"
                Case 19 To 21 'numeric, decimal, float--do nothing
                    sClause = "(" & sClause & " OR (" & sField & " = 0))"
                Case Else '** Error
            End Select
        ElseIf UCase(Me.cboOperator) = "IS NULL" Then
            sClause = sField & " IS NULL"
        ElseIf UCase(Me.cboOperator) = "IS NOT NULL" Then
            sClause = sField & " IS NOT NULL"
        ElseIf UCase(Me.cboOperator) = "IS NOT NULL AND IS NOT BLANK" Then 'CR # 2578 fix
            sClause = "(" & sField & " IS NOT NULL)"
            Select Case sDataType
                Case 2 To 7 'Numbers
                    sClause = "(" & sClause & " AND (" & sField & " <> 0))" 'CR # 2578 fix - Changed 'OR' to 'AND'
                Case Is = 10 'text
                    sClause = "(" & sClause & " AND (" & sField & " <> ''))"
                Case Is = 15 'bigint
                    sClause = "(" & sClause & " AND (" & sField & " <> 0))"
                Case Is = 18 'char
                    sClause = "(" & sClause & " AND (" & sField & " <> ''))"
                Case 19 To 21 'numeric, decimal, float--do nothing
                    sClause = "(" & sClause & " AND (" & sField & " <> 0))"
                Case Else '** Error
            End Select
        Else
            'CR # 2572 fix - Removed code that was parsing sValue with space as a seperator and invoked BuildCondition routine directly
            sValue = BuildCondition(sValue, sDataType)
            If sValue = "" Then
                txtValue.SetFocus
            Else
                sClause = sField & " " & Me.cboOperator & " " & sValue
            End If
        End If
         
        If sClause <> "" Then
            Me.lstCriteria.AddItem Chr(34) & sClause & Chr(34) & ";" & Me.cboField & ";" & Me.cboOperator & ";" & Chr(34) & Me.txtValue & Chr(34)
            BuildQuery
            Me.cboField = ""
            Me.cboOperator = ""
            txtValue = ""
            Me.cboField.SetFocus
        End If
    End If
    
exitHere:
    
    On Error Resume Next
    Exit Sub

HandleError:

    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GoTo exitHere
    
End Sub

Private Sub CmdCancel_Click()
    RaiseEvent Cancelled
End Sub

Private Sub cmdClearAll_Click()
    If Not Enabled Then
        Exit Sub
    End If
    
    Me.lstCriteria.RowSource = ""
End Sub

Private Sub cmdDeleteRow_Click()
    If Not Enabled Then
        Exit Sub
    End If
    
    If Me.lstCriteria.ItemsSelected.Count > 0 Then
        Me.lstCriteria.RemoveItem (Me.lstCriteria.ItemsSelected(0))
    Else
        MsgBox "Please select a row to delete.", vbOKOnly + vbExclamation, "Delete Row"
    End If
End Sub

Private Sub CmdEditRow_Click()
    If Not Enabled Then
        Exit Sub
    End If
    
    If lstCriteria.ItemsSelected.Count > 0 Then
        If UCase(lstCriteria.Column(2, lstCriteria.ItemsSelected(0))) = "CUSTOM" Then
            ZoomText "Row Edit", lstCriteria.ItemsSelected(0), True
        Else
            cboField = lstCriteria.Column(1, lstCriteria.ItemsSelected(0))
            cboOperator = lstCriteria.Column(2, lstCriteria.ItemsSelected(0))
            txtValue = lstCriteria.Column(3, lstCriteria.ItemsSelected(0))
            Me.lstCriteria.RemoveItem (lstCriteria.ItemsSelected(0))
        End If
    Else
        MsgBox "Please select a row to edit.", vbOKOnly + vbExclamation, "Edit Row"
    End If
End Sub

Private Sub cmdNew_Click()
    Dim hold As Boolean
    hold = saveAsEnabled
    txtFocus.SetFocus
    bNew = True
    saveAsEnabled = True
    cmdSaveAs_Click
    saveAsEnabled = hold
    bNew = False
End Sub

Private Sub cmdApply_Click()
    Dim hold As Boolean
    hold = saveEnabled
    ApplyFilter = False
    saveEnabled = True
    cmdSave_Click
    saveEnabled = hold
    If ApplyFilter Then
        If Nz(Me.lstFilters.Value, "") <> "" Then
            RaiseEvent FinishedFilter(Me.lstFilters.Value)
        Else
            RaiseEvent Cancelled
        End If
    End If
        
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    Dim varItem As Integer
    Dim FilterID As Long
    
    If Not deleteEnabled Then
        Exit Sub
    End If
    
    If Nz(lstFilters.Column(0), "") = "" Then
        MsgBox "Please select the filter to delete.", vbOKOnly, "Delete Filter"
        Exit Sub
    End If
    
    If Nz(lstFilters.Value, "") <> "" Then
        varItem = lstFilters.ItemsSelected(0)
        If bDCUser Or Identity.UserName = lstFilters.Column(2, varItem) Then
            FilterID = lstFilters.Value
            strSQL = "DELETE FROM CA_ScreensFilters WHERE FilterId = " & FilterID
            RunDAO strSQL
            ' force the criteria to refresh
            strSQL = "SELECT SqlString, FieldName, Operator, FieldValue FROM CA_ScreensFiltersDetails WHERE FilterId = " & FilterID
            RefreshListBox strSQL, Me.lstCriteria
            lstFilters.Requery
            lstFilters_AfterUpdate
        Else
            MsgBox "You can only delete filters you have created.", vbOKOnly + vbInformation, "Cannot Delete Filter"
        End If
    Else
        MsgBox "Please select the row to delete.", vbOKOnly, "Select a Row"
    End If
    
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim strCriteria As String

    If Not saveEnabled Then
        Exit Sub
    End If
    
    If Nz(lstFilters.Column(0), "") = "" Then
        MsgBox "Please select the filter to save.", vbOKOnly, "Select Filter"
        Exit Sub
    End If

    strCriteria = BuildQuery

    If Nz(strCriteria, "") <> "" Then
        'SA 10/1/2012 - Added CreatPK for SQL Server
        genUtils.CreatePK "CA_ScreensFilters", "FilterID"
        strSQL = " UPDATE CA_ScreensFilters" & _
                    " SET FilterSql = '" & Replace(strCriteria, "'", "''") & "'" & _
                    " WHERE FilterId = " & lstFilters.Column(0)
        RunDAO strSQL
        SaveFilterDetails lstFilters.Column(0)
        Me.lstFilters.Requery
        Me.lstFilters.Value = lstFilters.Column(0)
        lstFilters_Click
        ApplyFilter = True
    Else
        If MsgBox("The filter does not have any criteria." & _
            vbCrLf & vbCrLf & "Would you like to delete the filter?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete empty filter.") = vbYes Then
            Dim hold As Boolean
            hold = deleteEnabled
            deleteEnabled = True
            cmdDelete_Click
            deleteEnabled = hold
        End If
    End If

End Sub

Private Sub cmdSaveAs_Click()
    Dim strSQL As String
    Dim strCriteria As String
    Dim sFilterName As String
    Dim FilterID As Integer
    
    If Not saveAsEnabled Then
        Exit Sub
    End If
    
    sFilterName = EscapeQuotes(Nz(InputBox("Filter Name:", "Save Filter"), ""))
    sFilterName = Replace(sFilterName, "'", "''")
    ' Make sure the filter name doesn't exist
    If Nz(sFilterName, "") <> "" Then
        If Nz(DLookup("FilterName", "CA_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
            "ScreenId = " & oFrmScreen.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34)), "") <> "" Then
            sFilterName = ""
            MsgBox "There is already a filter with that name." & vbCrLf & "Please choose a different name.", vbOKOnly + vbExclamation, "Filter Name"
        Else
            If Not bNew Then
                strCriteria = BuildQuery
            Else
                strCriteria = ""
            End If
            
            'SA 11/15/2012 - Fixed insert SQL to prevent too many single quotes.
            strSQL = "INSERT INTO CA_ScreensFilters(ScreenID,FilterName,FilterSQL,UserName) " & _
                    "VALUES(" & oFrmScreen.ScreenID & ",'" & Replace(sFilterName, "'", "''") & "','" & _
                    Replace(strCriteria, "'", "''") & "','" & Replace(Identity.UserName, "'", "''") & "')"
                    
            RunDAO strSQL
            Me.lstFilters.Requery
            FilterID = DLookup("FilterId", "CA_ScreensFilters", "FilterName = " & Chr(34) & sFilterName & Chr(34) & " AND " & _
            "ScreenId = " & oFrmScreen.ScreenID & " AND Username = " & Chr(34) & Identity.UserName & Chr(34))
            lstFilters.Value = FilterID
            If Not bNew Then
                SaveFilterDetails FilterID
            End If
            lstFilters_Click
            SetButtonSettings
        End If
    Else
        MsgBox "Please enter a filter name.", vbOKOnly, "Save Filter"
        cmdSaveAs.SetFocus
    End If

End Sub

Private Sub ZoomText(Title As String, selectedItemNumber As Integer, ByVal Update As Boolean)
Dim FrmZoom As Form_CT_Text
Dim Txt As String

If selectedItemNumber > -1 Then
    Txt = "" & Me.lstCriteria.Column(0, selectedItemNumber)
Else
    Txt = ""
End If

Set FrmZoom = New Form_CT_Text
FrmZoom.Move lstCriteria.left
With FrmZoom
    .Text = Txt
    .Title = Title
    .visible = True
    If Update Then
        .Txt.Locked = False
    Else
        .Txt.Locked = True
    End If
    Do Until .Results <= 0
        DoEvents
    Loop
    If Update Then
        If .Results = True Then
            If selectedItemNumber > -1 Then
                If Me.lstCriteria.ItemsSelected.Count > 0 Then
                    Me.lstCriteria.RemoveItem (selectedItemNumber)
                Else
                    Me.lstCriteria.RowSource = ""
                End If
            End If
            .Text = Replace(.Text, Chr(34), "'")
            Me.lstCriteria.AddItem Chr(34) & .Text & Chr(34) & ";;CUSTOM;"
        End If
    End If
End With
End Sub

Private Sub ZoomToControl(ctl As TextBox, Optional ByVal Title As String = "")
' bring up a zoom window to show the text
Dim FrmZoom As Form_CT_Text
Dim Txt As String

    Txt = "" & ctl
    Set FrmZoom = New Form_CT_Text
    FrmZoom.Move ctl.left
    With FrmZoom
        .Text = Txt
        .Title = Title
        .visible = True
        Do Until .Results <= 0
            DoEvents
        Loop
        If .Results = True Then
            ctl = "" & .Text
        End If
    End With

End Sub

Private Sub CmdSQL_Click()
    If Not Enabled Then
        Exit Sub
    End If
    
    ZoomText "Enter SQL String", -1, True
End Sub

Private Sub CmdText_Click()
    ZoomToControl txtValue, "Filter Value Criteria"
End Sub

Private Sub cmdToggle_Click()
    TogValue = Not TogValue
    UpdateFilterList
End Sub


Private Sub cmdViewSQL_Click()
    If lstCriteria.ItemsSelected.Count > 0 Then
        ZoomText "View SQL", lstCriteria.ItemsSelected(0), False
    Else
        MsgBox "Please select a row to view.", vbOKOnly + vbExclamation, "View Row"
    End If

End Sub

Private Sub lstFilters_AfterUpdate()
    SetButtonSettings
    EnableFilterCriteria False
    Enabled = False
    
End Sub

Private Sub lstFilters_Click()
    If lstFilters.ListCount > 0 And Nz(lstFilters.Value, "") <> "" Then
        If lstFilters.Column(2) = Identity.UserName Or bDCUser Then
            EnableFilterCriteria True
            Enabled = True
        Else
            EnableFilterCriteria False
            Enabled = False
        End If
        LoadFilter
    Else
        EnableFilterCriteria False
        Enabled = False
    End If
End Sub

Private Sub EnableFilterCriteria(ByVal Setting As Boolean)
    txtFocus.SetFocus
    cboField.Enabled = Setting
    cboOperator.Enabled = Setting
    txtValue.Enabled = Setting
    
'    CmdAddRow.Enabled = setting
'    cmdDeleteRow.Enabled = setting
'    cmdClearAll.Enabled = setting
'    CmdEditRow.Enabled = setting
'    CmdSQL.Enabled = setting
    
    cmdViewSQL.Enabled = True
    CmdApply.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Function LoadFilter()
    Dim lngFilterId As Long
    Dim strSQL As String

    lngFilterId = CInt(Me.lstFilters)
    strSQL = "SELECT SqlString, FieldName, Operator, FieldValue FROM CA_ScreensFiltersDetails WHERE FilterId = " & lngFilterId
    RefreshListBox strSQL, Me.lstCriteria

End Function


Public Sub RefreshListBox(strSQL As String, lstBox As listBox, _
                            Optional varDefaultSelection As Variant = "", _
                            Optional strField As String = "")

    On Error GoTo ErrHandler

    Dim rst As RecordSet
    Dim db As Database
    Dim ctr As Long
    Dim strItem As String
    
    lstBox.RowSource = vbNullString
    For ctr = 0 To lstBox.ListCount - 1
        lstBox.RemoveItem (ctr)
    Next

    Set db = CurrentDb
    Set rst = db.OpenRecordSet(strSQL, dbOpenDynaset, dbSeeChanges)
    While Not rst.EOF
        For ctr = 0 To rst.Fields.Count - 1
            strItem = strItem & Chr(34) & rst.Fields(ctr).Value & Chr(34) & ";"
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
    lstBox.SetFocus

ExitNow:
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error refreshing list"
    GoTo ExitNow
End Sub

Private Function BuildQuery()
    On Error GoTo HandleError

Dim strCondition As String
Dim X As Single
    If Me.lstCriteria.ListCount > 0 Then
        strCondition = Me.lstCriteria.Column(0, 0)
        X = 1
        Do While X < Me.lstCriteria.ListCount
            strCondition = strCondition & " AND " & Me.lstCriteria.Column(0, X)
            X = X + 1
        Loop
    End If

exitHere:
    BuildQuery = strCondition
    Exit Function
HandleError:
    strCondition = ""
    GoTo exitHere
End Function

Private Sub SaveFilterDetails(ByVal FilterID As Long)
Dim i As Integer
Dim sqlString As String
Dim valueString As String
    ' first remove all the items in the table for this filter
    sqlString = "DELETE FROM CA_ScreensFiltersDetails WHERE FilterId = " & FilterID
    RunDAO sqlString

    sqlString = " INSERT INTO CA_ScreensFiltersDetails(FilterId,Operator,FieldName,FieldValue,SqlString) VALUES(" & FilterID & ","
    
    For i = 0 To lstCriteria.ListCount - 1
        valueString = Chr(34) & lstCriteria.Column(2, i) & Chr(34) & "," & Chr(34) & lstCriteria.Column(1, i) & _
                Chr(34) & "," & Chr(34) & lstCriteria.Column(3, i) & Chr(34) & "," & Chr(34) & _
                lstCriteria.Column(0, i) & Chr(34) & ")"
        RunDAO sqlString & valueString
    Next i

End Sub

Private Sub SetButtonSettings()
Dim Setting As Boolean

    Setting = False
    txtFocus.SetFocus
    If Nz(Me.lstFilters.Value, "") <> "" Then
        If lstFilters.Column(2) = Identity.UserName Then
            Setting = True
        End If
    End If
    
    CmdNew.Enabled = True
    'CmdSave.Enabled = setting
    saveEnabled = Setting
    
    If Me.lstFilters.ListCount > 0 Then
        deleteEnabled = Setting Or bDCUser
        'CmdDelete.Enabled = setting Or bDCUser
        saveAsEnabled = True
        'cmdSaveAs.Enabled = True
   Else
        deleteEnabled = Setting
        saveAsEnabled = Setting
'        CmdDelete.Enabled = False
'       cmdSaveAs.Enabled = False
    End If

End Sub

Private Sub RunDAO(SQL As String)
    On Error GoTo HandleErrors

    'SA 10/1/2012 - Added dbSeeChanges for SQL Server
    CurrentDb.Execute SQL, dbFailOnError + dbSeeChanges

exitHere:
    Exit Sub

HandleErrors:
    MsgBox "Error " & Err.Number & " ( " & Err.Description & ")" & vbCr & "String: " & SQL
    GoTo exitHere

End Sub

Private Function ParseValue(ByVal separator As String, ByVal inputString As String, ByVal FieldName As String) As String
Dim sDataType As String
Dim ReturnValue As String
Dim workingString As String
Dim workingSeparator As String
Dim currentValue As String
Dim workingValue As String
Dim Count As Integer
Dim i As Integer

ReturnValue = ""
sDataType = RetrieveDataType(FieldName)
' convert the input string to upper case so we can find the separator
workingString = UCase(inputString)
workingSeparator = UCase(separator)

Count = 0
i = InStr(workingString, workingSeparator)
While i > 0
    workingValue = Trim(Mid(workingString, 1, i - 1))
    currentValue = BuildCondition(workingValue, sDataType)
    If Nz(currentValue, "") <> "" Then
        If Count = 0 Then
            ReturnValue = currentValue
        Else
            ReturnValue = ReturnValue & separator & currentValue
        End If
        Count = Count + 1
        workingString = Mid(workingString, i + Len(separator))
        i = InStr(workingString, workingSeparator)
    Else
        ReturnValue = ""
        i = 0
    End If
Wend

If workingString <> "" Then
    currentValue = BuildCondition(Trim(workingString), sDataType)
    If Nz(currentValue, "") <> "" Then
        If Count = 0 Then
            ReturnValue = currentValue
        Else
            ReturnValue = ReturnValue & separator & currentValue
        End If
        Count = Count + 1
    Else
        ReturnValue = ""
    End If
End If
    
ParseValue = ReturnValue

End Function

Private Function BuildCondition(ByVal inputValue As String, ByVal sDataType As String) As String

Dim sField As String
Dim testDate As Date
Dim sValue As String
sField = ""
sValue = inputValue
    Select Case sDataType
        Case 2 To 7 'Numbers--do nothing
            If Not IsNumeric(sValue) Then
                MsgBox "Please enter a numeric value.", vbOKOnly + vbExclamation, "Value Entry Error."
            Else
                sField = sValue
            End If
        Case Is = 8 'date
            If Not IsDate(sValue) Then
                MsgBox "Please enter a date.", vbOKOnly + vbExclamation, "Value Entry Error."
            Else
                testDate = sValue
                If Not IsDate(testDate) Then
                    MsgBox "Please enter a date.", vbOKOnly + vbExclamation, "Value Entry Error."
                Else
                    sField = "#" & testDate & "#"  '* Access Syntax
                End If
            End If
        Case Is = 10 'text
            sField = Replace(sValue, "'", "''")
            sField = EscapeQuotes(sField)
            sField = "'" & sField & "'"
        Case Is = 15 'bigint--do nothing
            If Not IsNumeric(sValue) Then
                MsgBox "Please enter a numeric value.", vbOKOnly + vbExclamation, "Value Entry Error."
            Else
                sField = sValue
            End If
        Case Is = 18 'char--this handles  IN lists
            sField = Replace(sValue, "'", "''")
            sField = EscapeQuotes(sField)
            sField = "'" & sField & "'"
        Case 19 To 21 'numeric, decimal, float--do nothing
            If Not IsNumeric(sValue) Then
                MsgBox "Please enter a numeric value.", vbOKOnly + vbExclamation, "Value Entry Error."
            Else
                sField = sValue
            End If
        Case Else '** Error
            MsgBox "Error determining data type for field.", vbOKOnly + vbExclamation, "Value Entry Error."
           sField = ""
    End Select

exitHere:
    BuildCondition = sField
    Exit Function
End Function

Private Function RetrieveDataType(ByVal FieldName As String) As String
    On Error GoTo ErrorHappened
    Dim db As Database
    Set db = CurrentDb
    Dim bResult As String

    bResult = ""
Stop
'    bResult = CurrentDb.TableDefs(oFrmScreen.PrimaryCriteria).Fields(FieldName).Type '* 11/25/08 JAC

Done:
    db.Close
    Set db = Nothing
    RetrieveDataType = bResult
    Exit Function
    
ErrorHappened:
    bResult = ""
    On Error Resume Next
Stop
        'bResult = CurrentDb.QueryDefs(oFrmScreen.PrimaryCriteria).Fields(FieldName).Type
    Resume Done
    
End Function

Private Function UpdateFilterList()
Dim strSource  As String
    strSource = "SELECT FilterId, FilterName, UserName FROM CA_ScreensFilters WHERE ScreenId = " & oFrmScreen.ScreenID

    If TogValue Then
        cmdToggle.Caption = "Mine"
        strSource = strSource & " and UserName ='" & Identity.UserName & "'"
    Else
        cmdToggle.Caption = "All"
    End If
        
    strSource = strSource & " ORDER BY FilterName"
    Me.lstFilters.RowSource = strSource
    lstFilters_AfterUpdate

End Function
