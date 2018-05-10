Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'***************************************************************************************************************
'* This form is used as a standard query builder through the application and
'* uses the CnlyScreenSQL type.
'*
'* Setup checks for a FormID().  If one exists, it's in a screen and the
'* Refresh button is available; if FormID() = 0 then it's not in a screen.
'* Decihper use:
'*     Me.Subform.Form.sql = MyCnlyScreenSql
'*     Me.Subform.form.setup
'*
'*
'* The sql string is built up as the user clicks "Add Row".  The string is exposed through:
'*     Me.Subform.Form.sql.SqlAll
'*
'***************************************************************************************************************

Private mvCalledBy As String
Private mvSql As CnlyScreenSQL
Private mvFieldsTable As String
Private WithEvents mvPopup As Form_frm_General_Filter_Popup
Attribute mvPopup.VB_VarHelpID = -1

Private mvFormId  As Long

Public Event UpdateSql()
Public Event QueryFormClose()
Public Event QueryFormRefresh()

Public Property Get SQL() As CnlyScreenSQL
    SQL = mvSql
End Property

Public Property Let SQL(SQL As CnlyScreenSQL)
    mvSql = SQL
End Property

Public Property Let FieldsTable(TableName As String)
    '* This allows using a different table for your fields. Do we need to expand to allow a select ?
    mvFieldsTable = TableName
End Property

Public Property Let CalledbyScreen(FormID As Long)
    mvFormId = FormID
End Property

Public Property Let CalledBy(CallingForm As String)
    '* THis allows piggy-backing form-specific code
    mvCalledBy = CallingForm
End Property

Public Function Setup()
    If mvFormId > 0 Then    '* If there's a FormId, it's in Decipher
        Me.CmdRefresh.Caption = "Refresh"
    Else
        Me.CmdRefresh.Caption = "Update Main"
    End If

    If mvSql.From = "" And mvFieldsTable <> "" Then
        mvSql.From = mvFieldsTable
    ElseIf mvSql.From <> "" And mvFieldsTable = "" Then
        mvFieldsTable = mvSql.From
    End If
    
    If mvSql.Select = "" Then
        mvSql.Select = mvSql.From & ".*"
    End If
    
    Me.cboField.RowSource = mvFieldsTable
    
    
    'Pull fields and order from v_XREF_TableFields
    Me.cboField.RowSource = "SELECT FieldName FROM v_XREF_TableFields WHERE TableName = '" & mvFieldsTable & "' ORDER BY FieldName"
    Me.cboField.RowSourceType = "Table/Query"

    Me.cboOperator.RowSource = "CRITERIA_Operator"
    
    Me.lstCriteria.RowSource = ""
    Me.lstCriteria.RowSourceType = "value list"
    
    Me.txtValue = ""
    
    Me.tglType.Value = -1
    Me.tglType.Caption = "Mine"
    
    UpdateFilterList
    
End Function

Private Function UpdateFilterList()

  Dim strSource  As String

    ' TL add account ID logic and logic to filter by SQLFrom
    strSource = "SELECT CriteriaId, Description FROM CRITERIA_Hdr " & _
                " WHERE AccountID = " & gintAccountID & " and SourceObject = '" & mvCalledBy & "'" & _
                " AND SQLFrom = '" & mvFieldsTable & "'"

    With Me.tglType
        If .Value = -1 Then
            .Caption = "Mine"
            strSource = strSource & " and UserId ='" & Identity.UserName & "'"
        Else
            .Caption = "All"
        End If
     
        'Alex added 7/10/08
        strSource = strSource & " ORDER BY Description"
        
        Me.lstFilters.RowSource = strSource
        Me.lstFilters.Requery
    End With
End Function

Private Function BuildCondition(sValue As String, Optional strCustomType As String = "") As String
    '* BuildCondition(ctl As Control) As String was Passing a control to accomodate using the combo-box/text box as entry field
    Dim sWhere As String
    Dim sDataType As String
    Dim sField As String
    'Dim sValue As String

 '   sValue = ctl.Value

    If strCustomType = "" Then
        sDataType = CurrentDb.TableDefs(mvSql.From).Fields(Me.cboField).Type
    Else
        sDataType = strCustomType
    End If

    Select Case sDataType
        Case 2 To 7 '* Numbers--do nothing
            sField = sValue
        Case Is = 8 '* date
            sField = "#" & sValue & "#" '* Access Syntax
           '*  sField = "'" & sValue & "'" SQL Server Syntax for Pass-Through
        Case Is = 10 'text
            sField = "'" & Replace(sValue, ",", "','") & "'"   '* this handles  IN lists
        Case Is = 15 '* bigint--do nothing
            sField = sValue
        Case Is = 18 '* char--this handles  IN lists
            sField = "'" & Replace(sValue, ",", "','") & "'"   '* this handles  IN lists
        Case 19 To 21 '* numeric, decimal, float--do nothing
            sField = sValue
        Case Else '** Error
           sField = ""
    End Select


exitHere:
    BuildCondition = sField
    Exit Function
End Function


Private Sub cboOperator_AfterUpdate()

    Select Case cboOperator.Value
        Case Is = "cDRGMed"
            Me.cboField = "DRG"
            Me.txtValue = "True"
        Case Is = "xDRGMed"
            Me.cboField = "DRG"
            Me.txtValue = "False"
        Case Is = "cDRGSurg"
            Me.cboField = "DRG"
            Me.txtValue = "True"
        Case Is = "xDRGSurg"
            Me.cboField = "DRG"
            Me.txtValue = "False"
        Case Is = "cPDiagCC"
            Me.cboField = "PrincipalDiag"
            Me.txtValue = "True"
        Case Is = "xPDiagCC"
            Me.cboField = "PrincipalDiag"
            Me.txtValue = "False"
        Case Is = "cPDiagMCC"
            Me.cboField = "PrincipalDiag"
            Me.txtValue = "True"
        Case Is = "xPDiagMCC"
            Me.cboField = "PrincipalDiag"
            Me.txtValue = "False"
        Case Is = "c2DiagCC"
            Me.cboField = "CnlyClaimNum"
            Me.txtValue = "True"
        Case Is = "x2DiagCC"
            Me.cboField = "CnlyClaimNum"
            Me.txtValue = "False"
        Case Is = "c2DiagMCC"
            Me.cboField = "CnlyClaimNum"
            Me.txtValue = "True"
        Case Is = "x2DiagMCC"
            Me.cboField = "CnlyClaimNum"
            Me.txtValue = "False"
    End Select
    
End Sub

Private Sub CmdAdd_Click()
    On Error GoTo HandleError

    Dim sValue As String
    Dim sClause As String
    Dim sField As String
    Dim X As Integer
    Dim ln As Integer
    Dim sVal As String
    Dim sLval As String
    Dim sRval As String
  
    Dim strJoinTable As String
    Dim intOperatorID As Integer
    Dim strCustomType As String
        
    
        
    
    'check search type here
    Select Case mvFieldsTable
        Case Is = "AuditClm_Hdr"
          '  strJoinHcpcs = "AuditClm_Dtl"
          '  strJoinProc = "AuditClm_Proc"
          '  strJoinDIag = "AuditClm_Diag"
    End Select
  
    intOperatorID = Nz(DLookup("OperatorID", "CRITERIA_OPERATOR", "OperatorSQL = '" & Nz(Me.cboOperator, "") & "'"), 0)
    strCustomType = Nz(DLookup("DataType", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & ""), "")
    
    If Nz(Me.cboOperator, "") = "" Then
        MsgBox "Select an Operator"
        Me.cboOperator.SetFocus
    ElseIf (Nz(Me.cboField, "") = "" And strCustomType = "") Then
            MsgBox "Select a field"
            Me.cboField.SetFocus
    ElseIf Nz(Me.txtValue, "") = "" Then
        MsgBox "Enter a Value"
        Me.txtValue.SetFocus
    Else
        
        
        sVal = Me.txtValue.Value
        sField = Nz(Me.cboField, "")

        If Me.cboOperator <> "BETWEEN" Then
            sValue = BuildCondition(sVal, strCustomType)
        Else
            sValue = sVal
        End If
                
                
        If sValue <> "" Then
            Select Case Me.cboOperator
                Case Is = "Between"
                    X = InStr(sValue, "AND")
                    ln = Len(sValue)
                    sLval = BuildCondition(Trim(left(sValue, X - 1)))
                    sRval = BuildCondition(Trim(Right(sValue, ln - (X + 2))))
                    sClause = sField & " " & Me.cboOperator & " " & sLval & " and " & sRval
                Case Is = "NOT IN"
                    sClause = sField & " " & Me.cboOperator & " (" & sValue & ")"
                Case Is = "IN"
                    sClause = sField & " " & Me.cboOperator & " (" & sValue & ")"
                Case Is = "cDiag"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                             " AND qq.DiagCd IN (" & sValue & "))"
                Case Is = "xDiag"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " NOT EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                                 " AND qq.DiagCd IN (" & sValue & "))"
                Case Is = "cHCPCS"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                             " AND qq.HcpcsCd IN (" & sValue & "))"
                Case Is = "xHCPCS"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " NOT EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                              " AND qq.HcpcsCd IN (" & sValue & "))"
                Case Is = "cProc"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                             " AND qq.PRocCd IN (" & sValue & "))"
                Case Is = "xProc"
                    strJoinTable = DLookup("RefTable", "CRITERIA_OPERATOR_DTL", "OperatorID = " & intOperatorID & "")
                    sClause = " NOT EXISTS (SELECT 1 FROM " & strJoinTable & " qq WHERE qq.CnlyClaimNum = " & mvSql.From & ".CnlyClaimNum " & _
                             " AND qq.PRocCd IN (" & sValue & "))"
                
                Case Else
                    sClause = sField & " " & Me.cboOperator & " " & sValue
            End Select
            
            
            
            
            
            
            
            Me.lstCriteria.AddItem Chr(34) & sClause & Chr(34) & ";" & Me.cboField & ";" & Me.cboOperator & ";" & Chr(34) & Me.txtValue & Chr(34)
            BuildQuery
            RaiseEvent UpdateSql
        Else
            MsgBox "Unable to build condition.  Try again or get help.", vbOKOnly, "Unable to Build Condition"
        End If
            
        Me.cboField.SetFocus
        
    End If
    
exitHere:
    On Error Resume Next
    Exit Sub

HandleError:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    GoTo exitHere
      
End Sub

Private Sub cmdAddSql_Click()
    Dim varItem As Variant
 
    Set mvPopup = New Form_frm_General_Filter_Popup
    
    mvPopup.visible = True
    mvPopup.txtFilter = Me.lstCriteria.Column(0, varItem)
    mvPopup.Modal = True
    mvPopup.Caption = "Enter SQL String"

End Sub

Private Sub cmdClearAll_Click()
    
    Me.lstCriteria.RowSource = ""
    
    BuildQuery
    RaiseEvent UpdateSql

End Sub

Private Sub cmdDeleteFilter_Click()
    
    Dim strSQL As String
    Dim varItem As Variant

    If Me.lstFilters.ItemsSelected.Count <> 0 Then
        varItem = Me.lstFilters.ItemsSelected(0)
        strSQL = "DELETE FROM Criteria_Hdr WHERE CriteriaId = " & Me.lstFilters.Column(0, varItem)
        CurrentDb.Execute strSQL, dbSeeChanges
        Me.lstFilters.Requery
    Else
        MsgBox "Select a row to delete", vbOKOnly, "Select a Row"
    End If
    
End Sub

Private Sub cmdDeleteRow_Click()
    
    If Me.lstCriteria.ItemsSelected.Count > 0 Then
        Me.lstCriteria.RemoveItem (Me.lstCriteria.ItemsSelected(0))
        
        BuildQuery
        
        RaiseEvent UpdateSql
    End If
    
End Sub

Private Sub cmdEdit_Click()
    
    Dim varItem As Variant
    
    If Me.lstCriteria.ItemsSelected.Count <> 0 Then
    
        Set mvPopup = New Form_frm_General_Filter_Popup
        
        varItem = Me.lstCriteria.ItemsSelected(0)
             
        mvPopup.visible = True
        mvPopup.txtFilter = Me.lstCriteria.Column(0, varItem)
        mvPopup.Modal = True
        mvPopup.Caption = "Edit Row"
     
    End If

    
End Sub

Private Sub cmdLoad_Click()
    LoadFilter

End Sub

Private Sub lstFilters_DblClick(Cancel As Integer)
    LoadFilter
End Sub

Private Function LoadFilter()

    Dim lngFilterId As Long
    Dim strSQL As String
    Dim strSqlType As String
    Dim varItem As Variant
    Dim rst As DAO.RecordSet

    If Me.lstFilters.ItemsSelected.Count <> 0 Then
        varItem = Me.lstFilters.ItemsSelected(0)
        lngFilterId = Me.lstFilters.Column(0, varItem)
        strSQL = "SELECT SqlString, FieldName, Operator, FieldValue FROM Criteria_Dtl WHERE CriteriaId = " & lngFilterId
        RefreshListBox strSQL, Me.lstCriteria
        strSqlType = "SELECT * FROM CRITERIA_HDR WHERE CriteriaId = " & lngFilterId
        Set rst = CurrentDb.OpenRecordSet(strSqlType, dbOpenDynaset, dbSeeChanges)
        
        With mvSql
            .From = rst("SqlFrom")
            .WherePrimary = rst("SqlWhere")
            .OrderBy = rst("SqlOrderBy")
            .SqlAll = rst("SqlAll")
        End With
    Else
        MsgBox "Choose a filter to load"
    End If

End Function

Private Sub cmdRefresh_Click()
    On Error GoTo HandleError
    
    If mvFormId > 0 Then
    
    ''*Replace Decipher Screen Refresh
    '
    ''*Skipping BuildDetail function.  This is done by this form's BuildCondition
    '
    '    Dim oScr As Form_ScrMainScreens
    '
    '    Set oScr = Scr(mvFormId)
    '
    '    oScr.Sql = mvSql
    '    oScr.SubForm.Form.RecordSource = oScr.Sql.SqlAll
    '
    '    oScr.BuildTotals
    '    oScr.BuildTotalsCustom False
    '    'Code with error has been removed
    '
    '    RunEvent "Screen Refresh", oScr.ScreenID, mvFormId
    '
    ' Set oScr = Nothing
    Else
        '* Running Outside of screen
        '* Should be no need to update here, it's done as each row is updated
                
        If IsSubForm(Me) = False Then
            DoCmd.Close acForm, Me.Name
        End If

    End If

    RaiseEvent QueryFormRefresh


exitHere:

    On Error Resume Next
    DoCmd.Hourglass False

    Exit Sub

HandleError:

    MsgBox "There has been an error. Check to make sure your criteria is correctly formatted or start again." & vbCr & vbCr & _
           "Error: " & Err.Number & " (" & Err.Description & ")", vbOKOnly, "Error"
    Resume exitHere
    
    Resume Next
End Sub


Private Sub cmdSave_Click()

    Dim sFilterName As String

    '** ADO Parameters

    Dim strConnect As String
    Dim cnn As Variant
    Dim cmd As Variant
    Dim cmd2 As Variant

    Dim dtpDatestamp As Date
    Dim lngpCriteriaId As Long
    Dim Lst As listBox
    Dim ctr As Single
    
    If mvSql.WherePrimary <> "" And Me.lstCriteria.ListCount > 0 Then
        
        sFilterName = EscapeQuotes(Nz(InputBox("Filter Name:", "Save Filter"), ""))
    
        If Nz(DLookup("Description", "Criteria_Hdr", "Description = '" & sFilterName & "'"), "") <> "" Then '* Make sure filter doesn't exist
            
            sFilterName = ""
            
            MsgBox "There is already a filter with that name.  Please choose a different name", vbOKOnly

        End If
        
            
        If sFilterName <> "" Then
            
            '** Create objects & setup connection
            Set Lst = Me.lstCriteria
            Set cnn = CreateObject("ADODB.Connection")
            Set cmd = CreateObject("ADODB.Command")
            Set cmd2 = CreateObject("ADODB.Command")
            
            dtpDatestamp = Now()
            strConnect = GetConnectString("v_CODE_Database")
            cnn.Open strConnect
            
            '*** Header Insert
            cmd.ActiveConnection = cnn
            cmd.CommandTimeout = 30
            cmd.CommandText = "usp_CRITERIA_Hdr_Insert"
            cmd.commandType = adCmdStoredProc
            cmd.Parameters.Refresh
        
            cmd.Parameters("@pDatestamp") = dtpDatestamp
            cmd.Parameters("@pUserID") = Identity.UserName
            cmd.Parameters("@pClmType") = "" '** WHat's this?
            cmd.Parameters("@pSourceObject") = mvCalledBy
            cmd.Parameters("@pDescription") = sFilterName
            cmd.Parameters("@pSQLFrom") = mvSql.From
            cmd.Parameters("@pSQLWhere") = mvSql.WherePrimary
            cmd.Parameters("@pSQLOrderBy") = mvSql.OrderBy
            cmd.Parameters("@pSQLGroupBy") = ""
            cmd.Parameters("@pSQLHaving") = ""
            cmd.Parameters("@pSQLAll") = mvSql.SqlAll
            'TL add account ID logic
            cmd.Parameters("@pAccountID") = gintAccountID
            cmd.Parameters("@pCriteriaId") = 0
            cmd.Parameters("@pErrMsg") = ""
            
            cmd.Execute 128 '* adExecuteNoRecords
            
            '* if cmd.recordsaffected <> 1 then msgbox "error"
            
            lngpCriteriaId = Nz(cmd.Parameters("@pCriteriaId"), 0)
            
            If lngpCriteriaId <> 0 Then
            
                '***Detail Insert
                cmd2.ActiveConnection = cnn
                cmd2.CommandTimeout = 30
                cmd2.CommandText = "usp_CRITERIA_Dtl_Insert"
                cmd2.commandType = adCmdStoredProc
                
                cmd2.Parameters.Refresh
                
                ctr = 0
            
                Do While ctr <= Lst.ListCount
                
                        cmd2.Parameters("@pDateStamp") = dtpDatestamp
                        cmd2.Parameters("@pCriteriaID") = lngpCriteriaId
                        cmd2.Parameters("@pOperator") = Lst.Column(2, ctr)
                        
                        cmd2.Parameters("@pFieldName") = Lst.Column(1, ctr)
                        cmd2.Parameters("@pFieldValue") = Lst.Column(3, ctr)
                        cmd2.Parameters("@pSQLString") = Lst.Column(0, ctr)
                    
                        cmd2.Execute 128 '* adExecuteNoRecords
                        
                    ctr = ctr + 1
                Loop
            Else
            
                MsgBox "Error Saving Filter"
                
            End If '* lngpCriteriaId <> 0
                       
        End If '*  sFilterName <> ""
        
    Else
    
        MsgBox "There is no filter to save.", vbOKOnly, "No filter"
    
    End If  '* mvSql.WherePrimary <> "" And Me.lstCriteria.ListCount > 0
    
    UpdateFilterList

exitHere:
    Exit Sub

End Sub

Private Sub Form_Close()
    
    If mvFormId = 0 Then
        BuildQuery
        RaiseEvent UpdateSql
        Set mvPopup = Nothing
    End If

End Sub

Private Function BuildQuery()

    Dim strCondition As String
    Dim X As Single

    If Me.lstCriteria.ListCount > 0 Then

            strCondition = Me.lstCriteria.Column(0, X)
        
        X = 1
                
        Do While X < Me.lstCriteria.ListCount
            
            strCondition = strCondition & " and " & Me.lstCriteria.Column(0, X)
                    
            X = X + 1
            
        Loop

    End If

    With mvSql

        .WherePrimary = strCondition
        .SqlAll = "SELECT " & IIf(.Select <> "", .Select, "*") & _
                  " FROM " & .From & " WHERE " & .WherePrimary

    End With

exitHere:

End Function


Private Sub lstCriteria_DblClick(Cancel As Integer)

    Dim varItem As Variant

    varItem = Me.lstCriteria.ItemsSelected(0)
        
    If Me.lstCriteria.Column(1, varItem) <> "" Then
   
        Me.cboField = Me.lstCriteria.Column(1, varItem)
        
        Me.cboOperator = Me.lstCriteria.Column(2, varItem)
        Me.txtValue = Me.lstCriteria.Column(3, varItem)
        
        
       Me.lstCriteria.RemoveItem (varItem)
       
        BuildQuery
       RaiseEvent UpdateSql
       
    Else
        MsgBox "This row cannot be reloaded.  Open in the editor to make changes", vbOKOnly, "Row Can't be Loaded"
    End If

End Sub


Private Sub mvPopup_UpdateRow()

    Dim varItem As Variant

    If Me.lstCriteria.ItemsSelected.Count > 0 Then
    
        varItem = Me.lstCriteria.ItemsSelected(0)
        Me.lstCriteria.RemoveItem (varItem)
    
    Else
       Me.lstCriteria.RowSource = ""
    
    End If

    Me.lstCriteria.AddItem Chr(34) & mvPopup.txtFilter & Chr(34)
    
    BuildQuery
    
    RaiseEvent UpdateSql
    
End Sub

Private Sub tglType_Click()
    UpdateFilterList
End Sub
