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
''  The purpose of this is to be used as a control in another form
''  The other form will be fairly generic so hard to describe but here goes
''  The main form is going to allow the user to select something from a list view and "drop it" in here
''  in here, the user can add criteria to it.. So, say 'Provider' is dropped in. (Attached to that we
''  have the 'CnlyProvId' fieldname).  Criteria may be something like "IN (123321, 1234532)"
''  Or maybe state (ProvStateCd) is dropped in, criteria may be something like " NOT 'PA' "
''      AND NOT LIKE 'M%'
''
''  The way we are going to achie3ve this is with a local temp table...
''  We'll load it up with whatever is in SQL Server. Bind the form and lock the controls. Clicking Edit will unlock that row?
''  Maybe we should create one more form for the individual values where they can edit a single line.. so they can save / cancel?
''
''
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


Private Const cs_TEMP_TABLE_NAME As String = cs_TEMP_RULE_TABLE_NAME
Private coCurrentItem As clsBOLD_LetterRuleItemDetail
Private coCurrentItems As clsBOLD_LetterRuleItemDetails
Private csItemType As String


Public Event ItemRemoved(oItemRemoved As clsBOLD_LetterRuleItemDetail)
Public Event ItemSaveError(sItemName As String, sErrorMsg As String)
Public Event CriteriaChanged()



Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Public Property Get EnglishWhere() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oItem As clsBOLD_LetterRuleItemDetail
Dim sReturn As String
Dim iIdx As Integer

    strProcName = ClassName & ".EnglishWhere"


    For Each oItem In coCurrentItems.Items
        iIdx = iIdx + 1
        If iIdx = 1 Then
            If coCurrentItems.Items.Count > 1 Then
                sReturn = "( "
            End If
            sReturn = sReturn & oItem.ItemName & " " & oItem.OperatorTxt & " " & oItem.ItemValue
        Else
            sReturn = sReturn & " " & oItem.BooleanValTxt & " " & oItem.ItemName & " " & oItem.OperatorTxt & " " & oItem.ItemValue
        End If
        
    Next
    
    If sReturn <> "" Then
        If coCurrentItems.Items.Count > 1 Then
            sReturn = sReturn & " )"
        End If
    End If
    
    EnglishWhere = sReturn
    
Block_Exit:
    Set oItem = Nothing
    Exit Property
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Property

Public Property Get SqlWhere() As String

End Property

Public Property Let CurrentItem(oItem As clsBOLD_LetterRuleItemDetail)
     Set coCurrentItem = oItem
End Property
Public Property Get CurrentItem() As clsBOLD_LetterRuleItemDetail
    Set CurrentItem = coCurrentItem
End Property

Public Property Get ItemType() As String
    ItemType = csItemType
End Property
Public Property Let ItemType(sItemType As String)
    csItemType = sItemType
End Property


Public Function AddItemObj(oItemObj As clsBOLD_LetterRuleItemDetail) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".AddItemObj"
    If coCurrentItems Is Nothing Then
        Set coCurrentItems = New clsBOLD_LetterRuleItemDetails
    End If
    If coCurrentItems.Items.Count > 0 Then
        oItemObj.BooleanVal = 1
    End If
    
    AddItemObj = coCurrentItems.AddItemObj(oItemObj)
    
    RaiseEvent CriteriaChanged
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function AddItem(sItemName As String, sItemFieldName As String) As Boolean
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".AddItem"
    If coCurrentItems Is Nothing Then
        Set coCurrentItems = New clsBOLD_LetterRuleItemDetails
    End If
    AddItem = coCurrentItems.AddItem(sItemName, sItemFieldName, Me.ItemType)
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Public Function RemoveItem(oItem As clsBOLD_LetterRuleItemDetail)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".RemoveItem"
    
    If coCurrentItems Is Nothing Then
        Stop
        Set coCurrentItems = New clsBOLD_LetterRuleItemDetails
    End If
    RemoveItem = coCurrentItems.RemoveItem(oItem.ItemName, oItem.ItemFieldName, oItem.ItemType)
    
    RaiseEvent CriteriaChanged
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


'' If the table doesn't exist that we need, create it:
Private Function CreateTempTable() As Boolean
On Error GoTo Block_Exit
Dim strProcName As String
Dim oDb As DAO.Database
Dim oTDef As DAO.TableDef
Dim oFld As DAO.Field
Dim oIdx As DAO.index

    strProcName = ClassName & ".CreateTempTable"
    
    If IsTable(cs_TEMP_TABLE_NAME) = True Then
        CreateTempTable = True
        GoTo Block_Exit
    End If
    
    Set oDb = CurrentDb
    Set oTDef = New DAO.TableDef
    With oTDef
        .Name = cs_TEMP_TABLE_NAME
        
        Set oFld = New DAO.Field
        oFld.Name = "LocalId"
        oFld.Type = dbLong
        oFld.Attributes = oFld.Attributes Or dbAutoIncrField

        .Fields.Append oFld
        Set oFld = Nothing
        
        Set oIdx = .CreateIndex("PrimaryKey")
        Set oFld = oIdx.CreateField("LocalId")
        oIdx.Fields.Append oFld
        
        oIdx.Primary = True
        oIdx.Clustered = True
        .Indexes.Append oIdx
        
        Set oFld = New DAO.Field
        oFld.Name = "ItemType"
        oFld.Type = dbText
        oFld.Size = 255
        
     
        .Fields.Append oFld
        Set oFld = Nothing
        
        
        Set oFld = New DAO.Field
        oFld.Name = "RuleId"
        oFld.Type = dbLong
     
        .Fields.Append oFld
        Set oFld = Nothing

 
        Set oFld = New DAO.Field
        oFld.Name = "RuleItemId"
        oFld.Type = dbLong
     
        .Fields.Append oFld
        Set oFld = Nothing

        Set oFld = New DAO.Field
        oFld.Name = "Boolean"
        oFld.Type = dbInteger
     
        .Fields.Append oFld
        Set oFld = Nothing


        Set oFld = New DAO.Field
        oFld.Name = "LkupId"
        oFld.Type = dbLong
     
        .Fields.Append oFld
        Set oFld = Nothing


        Set oFld = New DAO.Field
        oFld.Name = "LkupDisplay"
        oFld.Type = dbText
        oFld.Size = 255
     
        .Fields.Append oFld
        Set oFld = Nothing

        Set oFld = New DAO.Field
        oFld.Name = "ItemName"
        oFld.Type = dbText
        oFld.Size = 255
     
        .Fields.Append oFld
        Set oFld = Nothing


        Set oFld = New DAO.Field
        oFld.Name = "RelatedFieldName"
        oFld.Type = dbText
        oFld.Size = 255
     
        .Fields.Append oFld
        Set oFld = Nothing


        Set oFld = New DAO.Field
        oFld.Name = "Operator"
        oFld.Type = dbText
        oFld.Size = 255
     
        .Fields.Append oFld
        Set oFld = Nothing


        Set oFld = New DAO.Field
        oFld.Name = "ItemValue"
        oFld.Type = dbText
        oFld.Size = 255
     
        .Fields.Append oFld
        Set oFld = Nothing

        
        
'        .Fields.Append
    End With
    
    oDb.TableDefs.Append oTDef
    oDb.TableDefs.Refresh
    
    CreateTempTable = IsTable(cs_TEMP_TABLE_NAME)
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Private Sub CmdAdd_Click()
    RaiseEvent CriteriaChanged
End Sub

Private Sub cmdDel_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim sItemRemoved As String
Dim oRs As DAO.RecordSet
Dim lId As Long
Dim sName As String


    Set oRs = Me.RecordsetClone
    
    If oRs.AbsolutePosition < 0 Then
    Stop
'        DoCmd.RunCommand acCmdSaveRecord
    End If

    lId = oRs("LocalId").Value
    sName = oRs("ItemName").Value
    
    If Not coCurrentItem Is Nothing Then
        If coCurrentItem.Id <> lId Then
            Set coCurrentItem = New clsBOLD_LetterRuleItemDetail
            Stop ' KD: Comeback - what about the list item object..
            Call coCurrentItem.LoadFromId(lId)
            
        End If
    End If

    If Not coCurrentItems Is Nothing Then
        Call coCurrentItems.RemoveItemObj(coCurrentItem)
    End If

    Me.Requery

    ' need to delete it from the temp table here
    ' remove it from the objects
    ' and then call the items removed event
    RaiseEvent ItemRemoved(coCurrentItem)
    
Block_Exit:
    Set oRs = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdEdit_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_BOLD_Letter_Rule_Required_Val_Edit
Dim oRs As DAO.RecordSet

    strProcName = ClassName & ".cmdEdit_Click"
    Set oFrm = New Form_frm_BOLD_Letter_Rule_Required_Val_Edit
'Stop

    ' Can I cound on the current object to be the correct one?
    Debug.Print coCurrentItem.Id
'    Stop
    ' maybe I need to use txtLocalId

'    Set oRs = Me.RecordsetClone

    oFrm.InitData (coCurrentItem.Id)
    oFrm.visible = True
    
    While oFrm.visible = True
        DoEvents
    Wend
    
    Call RefreshData(True)
    RaiseEvent CriteriaChanged
    
Block_Exit:
    Call DoCmd.Close(acForm, oFrm.Name, acSaveNo)
    Set oFrm = Nothing

    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub Form_Current()
Dim oRs As DAO.RecordSet
    Set oRs = Me.RecordSet
    
    If oRs.recordCount > 0 Then
'    Stop
        If oRs.AbsolutePosition > -1 Then
            If Not coCurrentItem Is Nothing Then
                If coCurrentItem.Id = oRs("LocalId").Value Then
                    GoTo Block_Exit ' already loaded..
                End If
            End If
            Set coCurrentItem = New clsBOLD_LetterRuleItemDetail
            coCurrentItem.LoadFromId (oRs("LocalId").Value)
        Else
            ' save it????
'            DoCmd.RunCommand acCmdSaveRecord
        End If
    End If
Block_Exit:
    Set oRs = Nothing
End Sub

Private Sub Form_Load()
Dim sSql As String
    ' load the operator combo here
    
    
    sSql = "SELECT OperatorId, OperatorName FROM BOLD_Letter_Automation_XREF_Operators WHERE Active <> 0  "
    
    '' Now need to get the combo box:
    Me.cmbOperator.ColumnCount = 2
    Me.cmbOperator.ColumnWidths = "0;20"
    Me.cmbOperator.ColumnCount = 2
    
    Call RefreshComboBoxADO(sSql, Me.cmbOperator, , , "v_Data_Database")
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo Block_Err
Dim strProcName As String

    strProcName = ClassName & ".Form_Load"
    If CreateTempTable = False Then
        LogMessage strProcName, "ERROR", "Could not create local work table!", , True
        Cancel = True
    Else
        Me.RecordSource = cs_TEMP_TABLE_NAME
    End If
    
    
    
Block_Exit:

    Exit Sub
Block_Err:
    ReportError Err, strProcName
    Cancel = True
    GoTo Block_Exit
End Sub

'' If we are loading from the table then we cannot truncate it first,
'' we want to populate our coCurrentItems object FROM the table, otherwise
'' we load the table from the coCurrentItems
Public Function RefreshData(Optional bLoadFromTable As Boolean = False) As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oDb As DAO.Database
Dim sSql As String
Dim oRs As DAO.RecordSet
Dim oFld As DAO.Field
Dim oCtl As Control
Dim sFieldList As String
Dim sValueList As String



    strProcName = ClassName & ".RefreshData"
    
    ' truncate our table
    If bLoadFromTable = False Then
        Set oDb = CurrentDb()
        oDb.Execute "TRUNCATE TABLE " & cs_TEMP_TABLE_NAME
        
        Set oRs = oDb.OpenRecordSet("SELECT * FROM " & cs_TEMP_TABLE_NAME & " WHERE 1 = 2", dbOpenDynaset, dbSeeChanges)

        
        For Each coCurrentItem In coCurrentItems.Items
            For Each oFld In oRs.Fields
                sFieldList = sFieldList & oFld.Name & ", "
                sValueList = sValueList & QuoteIfNeeded(coCurrentItem.GetField(oFld.Name)) & ", "
            Next
            sFieldList = left(sFieldList, Len(sFieldList) - 2)
            sValueList = left(sValueList, Len(sValueList) - 2)

                sSql = "INSERT INTO " & cs_TEMP_TABLE_NAME & " (" & sFieldList & ") "
                sSql = "VALUES (" & sValueList & ")"
                
Stop
            oDb.Execute sSql
        Next
        
    Else
        Stop    ' need to populate the coCurrentItems object FROM the table (should probably just have a method there since it should really always be attached to that table
                ' heck, the whole items object should have the method to load either way
                ' ok, I'll do it here then move it once I'm done.
                
                Me.txtFieldName.ControlSource = "RelatedFieldName"
                Me.txtItemName.ControlSource = "ItemName"
                Me.cmbOperator.ControlSource = "Operator"
                Me.txtValue.ControlSource = "ItemValue"
                Me.frmAndOr.ControlSource = "Boolean"
                Me.txtLocalId.ControlSource = "LocalID"
                
                Me.RecordSource = cs_TEMP_TABLE_NAME
                Me.filter = "ItemType = '" & Me.ItemType & "'"
                Me.FilterOn = True
                
                If coCurrentItem Is Nothing Then
                    Set coCurrentItem = New clsBOLD_LetterRuleItemDetail
                End If
                If coCurrentItems Is Nothing Then
                    Set coCurrentItems = New clsBOLD_LetterRuleItemDetails
                End If
                Me.Requery
                coCurrentItems.LoadFromItemType (Me.ItemType)
                
'                cojub
                

    End If

    '

Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function
