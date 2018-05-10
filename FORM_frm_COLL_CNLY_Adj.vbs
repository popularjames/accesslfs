Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=============================================
' ID:          Form_frm_COLL_CNLY_Adj
' Author:      Damon
' Create Date:
' Description:
'      Maintain Connolly manual adjustments for the given claim.  Once an Adjustment is APPROVED it can
' no longer be updated/deleted -- it will be posted to the Ledger.  A reversing entry is needed to modify the entry
' once it is posted.
'
' Modification History:
'   2012-12-26 by BJD to format the Window display and add additional business rules for initial deployment.
'
'
' =============================================

Public Event RecordChanged()

Private Const strAPPROVAL_STATUS_CD_APPROVED = "APPROVED"  'Approved status will trigger the update to the Ledger.
Private Const strACTIVITY_TYPE_CD_ADJUSTUP = "ADJUST +"
Private Const strACTIVITY_TYPE_CD_ADJUSTDOWN = "ADJUST -"

Private mrsCollCnlyAdjustment As ADODB.RecordSet
Private mbRecordChanged As Boolean
Private mstrCnlyClaimNum As String
Private mstrCnlyARCollID As String
Private mbInsert As Boolean
Private mbDirty As Boolean
Private mbAllowChange As Boolean
Private mbAllowAdd As Boolean

Private WithEvents MyAdo As clsADO
Attribute MyAdo.VB_VarHelpID = -1
Private WithEvents myCode_ADO As clsADO
Attribute myCode_ADO.VB_VarHelpID = -1

Public Event RecordSaved()

Const CstrFrmAppID As String = "LdgrCnlyM"  'Used for form security
Private miAppPermission As Integer
Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property
Property Let Insert(data As Boolean)
    mbInsert = data
End Property
Property Get Insert() As Boolean
    Insert = mbInsert
End Property

Public Property Set CollRecordSource(data As ADODB.RecordSet)
    Set mrsCollCnlyAdjustment = data
End Property

Public Property Get CollRecordSource() As ADODB.RecordSet
    Set CollRecordSource = mrsCollCnlyAdjustment
End Property
Public Property Get FormCnlyClaimNum() As String
    FormCnlyClaimNum = mstrCnlyClaimNum
End Property

Public Property Let FormCnlyClaimNum(data As String)
    mstrCnlyClaimNum = data
End Property

Public Property Get FormCnlyARCollID() As String
    FormCnlyARCollID = mstrCnlyARCollID
End Property

Public Property Let FormCnlyARCollID(data As String)
    mstrCnlyARCollID = data
End Property
Public Sub RefreshData()
    On Error GoTo ErrHandler
    
    Me.Caption = gstrAcctDesc & " - Collection Manual Adjustment"
    Dim strSQL As String
    Dim ctl As Variant
    Set mrsCollCnlyAdjustment = CreateObject("ADODB.Recordset")
    
    If mstrCnlyARCollID = "" Then 'Insert
        Me.Insert = True
        'Me.FormCnlyARCollID = "09211743470001W18003D20121024TRECOUPMENTA737SCNLY00"
        'Me.FormCnlyARCollID = "NEW"
        
        Me.lblAppTitle.Caption = "NEW ADJUSTMENT: " & Nz(mstrCnlyClaimNum, "")
    Else 'Edit
        Me.lblAppTitle.Caption = "Edit ADJUSTMENT: " & Nz(mstrCnlyARCollID, "")
        Me.SequenceCd.Locked = True  'User input only on insert.  It is used for the Primary Key when the exact same Activity occurs on the same day.
        Me.SequenceCd.BackColor = "12632256" 'Grey.
        Me.SequenceCd.TabStop = False
        Me.AllowAdditions = False
        Me.AllowDeletions = False

    End If
    
        Set MyAdo = New clsADO
        MyAdo.ConnectionString = GetConnectString("v_DATA_Database")
        MyAdo.sqlString = "select * from COLL_CNLY_Adj where CnlyARCollID = '" & Me.FormCnlyARCollID & "'"
        Set mrsCollCnlyAdjustment = MyAdo.OpenRecordSet
        Set MyAdo = Nothing
       
    Set Me.RecordSet = Nothing
    Set Me.RecordSet = mrsCollCnlyAdjustment



    'Loop through the controls setting their control source to the recordset
    For Each ctl In Me.Controls
    'MsgBox ctl.Name, vbOKOnly
        If ctl.Tag = "R" Then
             Me.Controls(ctl.Name).ControlSource = mrsCollCnlyAdjustment.Fields(ctl.Name).Name
        End If
    Next
    
    ' Set the form to display only once it is Approved.
    SetApprovedDisplayOnly
                
    If (mbAllowChange = False) Then
        Me.cmdSave.Enabled = False
    End If
   
    
Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " (" & Err.Description & ")" & vbCr & vbCr & "Source: " & Err.Source, vbOKOnly + vbCritical, " RefreshMain"
End Sub

Private Sub ActivityTypeCd_BeforeUpdate(Cancel As Integer)
'Me.ActivityTypeDomainCd = Me.ActivityTypeCd.Column(1)value defaulted by the database.
End Sub


Private Sub cmdSave_Click()

    'Perform Validation.
    If Not ValidCnlyCollRec Then
        Exit Sub
    End If

    'Save the Record.
    SaveData
    
    If Me.Insert Then
             DoCmd.Close
    Else
        ' Set the form to display only once it is Approved.
        SetApprovedDisplayOnly
    End If
        
    
End Sub

Private Sub CnlyAdjustmentCd_BeforeUpdate(Cancel As Integer)
'Me.CnlyAdjustmentDomainCd = Me.CnlyAdjustmentCd.Column(4) value defaulted by the database.
End Sub

Private Sub Form_Load()
    Me.Caption = "Collection Manual Adjustment"

    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
    mbAllowChange = (miAppPermission And gcAllowChange)
    mbAllowAdd = (miAppPermission And gcAllowAdd)
    
    mbRecordChanged = False
    mbDirty = False
    
'    Refreshdata

End Sub

Private Sub SaveData()

Dim bResult As Boolean
Dim strErrMsg As String
  
  
  Set myCode_ADO = New clsADO
    myCode_ADO.ConnectionString = GetConnectString("v_CODE_Database")
    
    If mrsCollCnlyAdjustment.recordCount > 0 Then
        mrsCollCnlyAdjustment.MoveFirst
    End If
    
    bResult = False
    If Not (mrsCollCnlyAdjustment.BOF And mrsCollCnlyAdjustment.EOF) Then
        Set mrsCollCnlyAdjustment = Me.RecordSet
        mrsCollCnlyAdjustment.MoveFirst
        ' The following two fields are assigned by the Database Trigger.
'        mrsCollCnlyAdjustment("LastUpDt") = Now
'        mrsCollCnlyAdjustment("LastUPUser") = Identity.UserName
        
        If Me.Insert Then
            mrsCollCnlyAdjustment("cnlyCLaimNum") = mstrCnlyClaimNum
            mrsCollCnlyAdjustment("CreateDt") = Now
        End If
        bResult = myCode_ADO.Update(mrsCollCnlyAdjustment, "usp_COLL_CNLY_Adj_Apply")
        If bResult = False Then
            strErrMsg = "Error: can not save record"
            GoTo Err_handler
        End If
    End If
    
    mbRecordChanged = False
    
Exit Sub

Err_handler:
    'Rollback anything we did up until this point
    strErrMsg = strErrMsg & vbCrLf & Err.Description
    'RaiseEvent ProvError(strErrMsg, Err.Number, strErrSource)
    'myCode_ADO.RollbackTrans  ' Not needed with this implementation
    'SaveProv = False
    
End Sub

'Change to Display Only if the Adjustment has been approved. No updates allowed once it has been approved.
'It will be posted to the Ledger.  A reversing entry is needed once it is posted to the Ledger.
'This is called when the Form is opened or an update is saved.
Private Sub SetApprovedDisplayOnly()
    Dim ctl As Variant

    On Error GoTo Err_SetApprovedDisplayOnly
    If Me.ApprovalStatusCd.Column(0) = "APPROVED" Then

        'Lock and Grey the background on all the controls.
        For Each ctl In Me.Controls
            If ctl.Tag = "R" Then
                Me.Controls(ctl.Name).BackColor = "12632256" 'Grey.
                Me.Controls(ctl.Name).Locked = True
            End If
        Next
    
    End If
    
Exit_SetApprovedDisplayOnly:
    Exit Sub
    
Err_SetApprovedDisplayOnly:
    MsgBox Err.Description, vbOKOnly + vbCritical
    'If there is an error, Close the Form to ensure there are not updates to a record that may
    'have been posted to the Ledger.
    DoCmd.Close
    
End Sub

'Validate the record values.
Private Function ValidCnlyCollRec() As Boolean
    On Error GoTo Err_ValidCnlyCollRec

    ' Init to False
    ValidCnlyCollRec = False

    If mrsCollCnlyAdjustment.recordCount > 0 Then
        mrsCollCnlyAdjustment.MoveFirst
    End If
    If Not (mrsCollCnlyAdjustment.BOF And mrsCollCnlyAdjustment.EOF) Then
        Set mrsCollCnlyAdjustment = Me.RecordSet
        mrsCollCnlyAdjustment.MoveFirst
        
        'Validate that the appropriate field value is used for AR Priciple Adjustment vs Recovered/Paid Adjustment.
        If ((mrsCollCnlyAdjustment("ActivityTypeCd") = strACTIVITY_TYPE_CD_ADJUSTUP) Or (mrsCollCnlyAdjustment("ActivityTypeCd") = strACTIVITY_TYPE_CD_ADJUSTDOWN)) _
        And Not (mrsCollCnlyAdjustment("CnlyPrincRecovOrPaidAmt") = 0) Then
            MsgBox "Adjust +/- is to the AR Principle only (use AR Adjust Amt). Recov/Paid Amt must be zero.  ", vbOKOnly + vbCritical, "Connolly Collection Adjustment Validation - AR Principle Adjustment"
            Exit Function
        End If

        If Not ((mrsCollCnlyAdjustment("ActivityTypeCd") = strACTIVITY_TYPE_CD_ADJUSTUP) Or (mrsCollCnlyAdjustment("ActivityTypeCd") = strACTIVITY_TYPE_CD_ADJUSTDOWN)) _
        And Not (mrsCollCnlyAdjustment("CnlyARAdjAmt") = 0) Then
            MsgBox "AR Adjust Amt must be zero except for Activity Type Adjust +/-.  Use the Recov/Paid Amt.  ", vbOKOnly + vbCritical, "Connolly Collection Adjustment Validation - AR Principle Adjustment"
            Exit Function
        End If
    End If
    
    ValidCnlyCollRec = True
    
Exit_ValidCnlyCollRec:
    Exit Function
    
Err_ValidCnlyCollRec:
    MsgBox Err.Description, vbOKOnly + vbCritical
    GoTo Exit_ValidCnlyCollRec
    
End Function


' Set the Record Changed indicator to True when a change to the data is detected.
Private Sub Form_Dirty(Cancel As Integer)
    mbRecordChanged = True
End Sub


' Prompt the user to save the record if it has changed.
Private Sub Form_Unload(Cancel As Integer)
    
    'Do not prompt to save if the user is not allowed to make changes.
    If (mbAllowChange = False) And (mbAllowAdd = False) Then
        GoTo Exit_Form_Unload
    End If

    If (Me.RecordSource <> "") And (mbRecordChanged Or Me.Dirty) Then

        If MsgBox("Record has changed.   Would you like to save it?", vbYesNo) = vbYes Then
            
            'Perform Validation.
            If Not ValidCnlyCollRec Then
                Cancel = 1
                Exit Sub
            End If


            'Save the Record.
            SaveData

        End If
    End If
    
        
Exit_Form_Unload:
    Exit Sub
End Sub



Private Sub myADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & " - " & Err.Description
End Sub

Private Sub myCODE_ADO_ADOError(ErrMsg As String, ErrNum As Long, ErrSource As String)
    MsgBox ErrMsg & " - " & Err.Description
End Sub
