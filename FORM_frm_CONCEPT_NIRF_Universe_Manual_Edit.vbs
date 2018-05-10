Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private coRs As ADODB.RecordSet
Private Const csTmpTableName As String = "tmp_NIRF_Editor_Universe"
Private cbIsDirty As Boolean


Public Property Get IsDirty() As Boolean
    IsDirty = cbIsDirty
End Property
Public Property Let IsDirty(bIsDirty As Boolean)
    cbIsDirty = bIsDirty
End Property

Public Property Get ClassName() As String
    ClassName = TypeName(Me)
End Property


Private Sub CmdCancel_Click()
    ' Clean out the table
Dim oDb As DAO.Database
    Set oDb = CurrentDb
    oDb.Execute "DELETE FROM " & csTmpTableName
    Set oDb = Nothing
    DoCmd.Close acForm, Me.Name, acSaveYes
    
End Sub

Private Sub cmdRevert_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oFrm As Form_frm_CONCEPT_NIRF_Editor_Universe_Revert_List
Dim bFormClosed   As Boolean
Dim strFormName As String
    
    strProcName = ClassName & ".cmdRevert_Click"
    
    If Nz(Me.txtManualEditId, 0) = 0 Then
        LogMessage strProcName, "USER NOTICE", "There are no previous edits for this NIRF.", , True, Me.txtConceptID.Value

        GoTo Block_Exit
    End If
    
    Set oFrm = New Form_frm_CONCEPT_NIRF_Editor_Universe_Revert_List
    ColObjectInstances.Add oFrm, oFrm.hwnd & ""
    oFrm.ConceptID = Me.txtConceptID.Value
    oFrm.PayerNameId = Me.PayerNameId
    oFrm.RefreshData
    
    
     strFormName = oFrm.Name
     oFrm.visible = True

     Do
        'Is it still Open?
        If IsLoaded(strFormName) Then
            DoEvents
            Wait 1
        ElseIf oFrm.visible = False Then
            bFormClosed = True
        Else
            bFormClosed = True
        End If
        
        If oFrm.ManualEditIdSelected <> 0 Then
            bFormClosed = True
        End If
        If oFrm.Canceled = True Then
            bFormClosed = True
        End If
        
     Loop Until bFormClosed = True
    
'    ShowFormAndWait oFrm
    
    Call RefreshData
    
'    Set frmCalendar = New Form_frm_GENERAL_Calendar
'    ColObjectInstances.Add frmCalendar, frmCalendar.hwnd & ""
'    frmCalendar.DatePassed = Nz(Me.txtThroughDate, Date)
'    frmCalendar.RefreshData

Block_Exit:
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Sub cmdSave_Click()
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim oAdo As clsADO
Dim sXmlParam As String
Dim sXml As String

    strProcName = ClassName & ".cmdSave_Click"
    
    ' If we already have a ManualEditId then we need to UPDATE, or, should we archive, delete then insert?
    ' we'll let the stored proc decide that but I think we are always going to insert
    ' and I'll have to change the view to get the use the most recent one (but then again, we're giving them a
    ' revert option..
    
Call SaveNow
    
GoTo Block_Exit
    
    Set oAdo = New clsADO
    With oAdo
        .ConnectionString = GetConnectString("v_Code_Database")
        .SQLTextType = StoredProc
        .sqlString = "usp_CONCEPT_NIRF_Save_Universe_Manual_Changes"
        .Parameters.Refresh
        
    End With
    
    For Each oCtl In Me.Controls
        If oCtl.Tag <> "" Then
            sXmlParam = sXmlParam & "p" & oCtl.Tag & "=" & Nz(oCtl.Value, "") & "|"
        End If
    Next
    
    sXml = BuildXmlParams(sXmlParam)
    oAdo.Parameters("@pXmlParams") = sXml
Debug.Print sXml
    
    oAdo.Execute
    If Nz(oAdo.Parameters("@pErrMsg"), "") <> "" Then
        LogMessage strProcName, "ERROR", "There was an error saving the NIRF", oAdo.Parameters("@pErrMsg").Value, True, Me.ConceptID
    Else
        DoCmd.Close acForm, Me.Name
    End If
    
    
    
Block_Exit:
    Set oCtl = Nothing
    Set oAdo = Nothing
    Exit Sub
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Sub

Private Function RecalcSums() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oRs As DAO.RecordSet
Dim lClaimCount As Long
Dim lClaimVal As Long

'
'    strProcName = ClassName & ".RecalcSums"
'    If Me.Recordset Is Nothing Then
'        GoTo Block_Exit
'    End If
'    Set oRs = Me.RecordsetClone
'
'    oRs.MoveFirst
'
'
'    While Not oRs.EOF
'        lClaimCount = lClaimCount + oRs("ClaimCount").Value
'        lClaimVal = lClaimVal + oRs("ClaimValue").Value
'
'        oRs.MoveNext
'    Wend
'
'
'    Me.txtClaimCountTTL = lClaimCount
'    Me.txtClaimValTTL = lClaimVal
'
    
Block_Exit:
    Set oRs = Nothing
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function


Public Function RefreshData() As Boolean
On Error GoTo Block_Err
Dim strProcName As String
Dim oCtl As Control
Dim sFilter As String
Dim oCn As ADODB.Connection
Dim sSql As String

    strProcName = ClassName & ".RefreshData"
    
    Me.RecordSource = ""
    If Not coRs Is Nothing Then
        If coRs.State = adStateOpen Then coRs.Close
    End If
    
    Set Me.RecordSet = Nothing
    
    Set coRs = Nothing
    
'    Me.OpenArgs = "CM_C2052 AND PayerNameid = 1008 "
'    sFilter = " WHERE ConceptId = 'CM_C2052' AND PayerNameId = 1008 "

    If Me.OpenArgs <> "" Then
        sFilter = "WHERE " & Me.OpenArgs
        '' kD you need to fix this to accept payernameid's also!
    End If
    
    
    sSql = "SELECT * FROM v_CONCEPT_NIRF_Universe_W_Edits " & sFilter
    
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = GetConnectString("v_Code_Database")
    oCn.CursorLocation = adUseClientBatch
    oCn.Open
    

    Set coRs = New ADODB.RecordSet
    
    coRs.CursorLocation = adUseClientBatch
    coRs.CursorType = adOpenKeyset
    coRs.LockType = adLockBatchOptimistic
    coRs.Open sSql, oCn
    
    Call CopyDataToLocalTmpTable(coRs, False, csTmpTableName)

    ' disconnect:
    Set coRs.ActiveConnection = Nothing

    Set Me.RecordSet = Nothing
    Me.RecordSource = "tmp_NIRF_Editor_Universe"
    
    
'

'    'Loop through the controls setting their control source to the recordset
'    For Each ctl In Me.Controls
'    'MsgBox ctl.Name, vbOKOnly
'        If ctl.Tag = "R" Then
'             Me.Controls(ctl.Name).ControlSource = mrsCollCnlyAdjustment.Fields(ctl.Name).Name
'        End If
'    Next
        
    
'    Stop
    
    
'    Debug.Print TypeName(oCtl.co)
'
    For Each oCtl In Me.Controls
        If oCtl.Tag <> "" Then
            If isField(coRs, oCtl.Tag) = True Then
                Me.Controls(oCtl.Name).ControlSource = oCtl.Tag
'                oCtl.Value = coRS(oCtl.Tag).Value
            End If
        End If
    Next
    
    Me.txtConceptID.DefaultValue = """" & Me.ConceptID & """"
    Me.ConceptID.DefaultValue = """" & Me.ConceptID & """"
    Me.txtManualEditId.DefaultValue = Nz(Me.txtManualEditId, "")
    If Nz(Me.PayerNameId, 0) = 0 Then
        If Me.OpenArgs <> "" Then
            Me.PayerNameId = Right(Me.OpenArgs, 4)
        End If
    Else
        Me.PayerNameId.DefaultValue = Me.PayerNameId
    End If
    
Block_Exit:
    Exit Function
Block_Err:
    ReportError Err, strProcName
    GoTo Block_Exit
End Function




Private Function SaveNow() As String
On Error GoTo Block_Err
Dim strProcName As String
Dim oCn As ADODB.Connection
Dim oRs As ADODB.RecordSet
Dim oDb As DAO.Database
Dim oFld As ADODB.Field
Dim oDaoRs As DAO.RecordSet

    strProcName = ClassName & ".SaveNow"

    If IsTable(csTmpTableName) = False Then
        ' nothing to save..
        LogMessage strProcName, "USER ERROR?", "Nothing to save... If this is in error, please reload the form and try again.", , True
        GoTo Block_Exit
    End If

    RunCommand acCmdSaveRecord

    Set oDb = CurrentDb
    
    
    Set oDaoRs = oDb.OpenRecordSet(csTmpTableName, dbOpenTable)

    '' open our RS in batch mode so we can populate it:
    Set oCn = New ADODB.Connection
    With oCn
        .ConnectionString = GetConnectString("v_Data_Database")
        .CursorLocation = adUseClientBatch
        .Open
    End With
Dim lRet As Long

Dim oCmd As ADODB.Command

    oCn.BeginTrans
Dim oCn2 As ADODB.Connection
Dim lManualEditId As Long

   Set oCn2 = New ADODB.Connection
    With oCn2
        .ConnectionString = GetConnectString("v_Code_Database")
        .CursorLocation = adUseServer
        .Open
    End With
    
    Set oCmd = New ADODB.Command
    With oCmd
        .commandType = adCmdStoredProc
        .CommandText = "usp_CONCEPT_NIRF_Save_Universe_Manual_Changes"
        Set .ActiveConnection = oCn2
        .Parameters.Refresh
        .Parameters("@pConceptId") = Me.ConceptID
        .Parameters("@pPayerNameId") = Me.PayerNameId
        If Nz(Me.txtManualEditId, "") <> "" Then
            .Parameters("@pManualEditId") = Me.txtManualEditId
        End If
        .Execute (lRet)
        If Nz(.Parameters("@pErrMsg").Value, "") <> "" Then
            LogMessage strProcName, "ERROR", "There was an error trying to save the details", .Parameters("@pErrMsg").Value, True, Me.ConceptID
            Stop
            oCn.RollbackTrans
            GoTo Block_Exit
        End If
        lManualEditId = .Parameters("@pManualEditId")
    End With
    oCn2.Close
    Set oCmd = Nothing
    Set oCn2 = Nothing
'
'    oCn.Execute "EXEC CMS_AUDITORS_CODE.dbo.usp_CONCEPT_NIRF_Save_Universe_Manual_Changes @pConceptId = '" & Me.ConceptID & "', @pPayerNameId = " & CStr(Me.PayerNameID)
    

    Set oRs = New ADODB.RecordSet
    With oRs
        .CursorLocation = adUseClientBatch
        .LockType = adLockBatchOptimistic
        Set .ActiveConnection = oCn
        .Open ("SELECT ManualEditId, ConceptId, PayerNameId, PayerName, ConceptState, StateName, ClaimCount, ClaimValue, DataType, LastUser FROM CONCEPT_NIRF_Universe_Edits WHERE 1 = 2")
        
        ''-- disconnect:
        Set .ActiveConnection = Nothing
    End With
    
    ' populate it:
    
    oDaoRs.MoveFirst
    
    While Not oDaoRs.EOF
        oRs.AddNew
        For Each oFld In oRs.Fields
            Select Case UCase(oFld.Name)
            Case "MANUALEDITID"
                ' not going to do anything with this field..
                oRs(oFld.Name) = lManualEditId
            Case "LASTUSER"
                oRs(oFld.Name) = Identity.UserName
            Case Else
                oRs(oFld.Name) = oDaoRs(oFld.Name).Value
            End Select
        Next
        oDaoRs.MoveNext
    Wend

    ' Ok, re-connect and
    Set oRs.ActiveConnection = oCn
    oRs.UpdateBatch
    oRs.Close
    Set oRs = Nothing
    oCn.CommitTrans

    
Block_Exit:
    
    Set oDaoRs = Nothing
    Set oFld = Nothing
    
    Call RefreshData
    Exit Function
Block_Err:
    MsgBox Err.Description, vbOKOnly + vbCritical
    ReportError Err, strProcName
    GoTo Block_Exit
End Function

Private Sub Form_Dirty(Cancel As Integer)
    IsDirty = True
End Sub

'
'Private Sub Form_Current()
'    ' As much as I hate using _Current....
'    Call RecalcSums
'End Sub

Private Sub Form_Load()
    Me.InsideHeight = 5000
    '' need to bind this to the recordset..
    Call RefreshData
    
End Sub
