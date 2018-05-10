Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mbDetailFormLoaded As Boolean
Private mbGridFormLoaded As Boolean
Private miRowsSelected As Long

Public mstrFilter As String
Public mstrSort As String

Const CstrFrmAppID As String = "QAMain"

Public Property Get frmAppID() As String
    frmAppID = CstrFrmAppID
End Property

Public Property Get DetailFormLoaded() As Boolean
    DetailFormLoaded = mbDetailFormLoaded
End Property

Public Property Let DetailFormLoaded(ByVal vData As Boolean)
    mbDetailFormLoaded = vData
End Property

Public Property Get GridFormLoaded() As Boolean
    GridFormLoaded = mbGridFormLoaded
End Property

Public Property Let GridFormLoaded(ByVal vData As Boolean)
    mbGridFormLoaded = vData
End Property

Public Property Get RowsSelected() As Long
    RowsSelected = miRowsSelected
End Property

Public Property Let RowsSelected(ByVal vData As Long)
    miRowsSelected = vData
End Property


Private Sub cmdFindICN_Click()
    If Me.txtICN & "" = "" Then
        MsgBox ("Please provide an ICN to search")
    Else
        Me.frm_AppealAnalystWorkList_Claims.Form.RecordSource = "select * from v_AppealHearingAnalystWorkList_CSTemplate where ICN like '" & Me.txtICN & "%'"
        If Me.frm_AppealAnalystWorkList_Claims.Form.RecordSet.recordCount = 0 Then
            MsgBox ("Requested ICN was not found.")
        End If
    End If
End Sub




Private Sub Form_Load()
Dim strCnlyClaimNum As String
    Me.Caption = "Appeal Analyst Work List"
    
    Call Account_Check(Me)
    miAppPermission = UserAccess_Check(Me)
    If miAppPermission = 0 Then Exit Sub
    
    Call cmdApplyFilters_Click
        
'    'If mbDetailFormLoaded And mbGridFormLoaded Then
'        strCnlyClaimNum = Me.frm_AppealAnalystWorkList_Claims.Form.CnlyClaimNum
'        Me.frm_AppealAnalystWorkList_Claims.Form.RecordSource = "select * from Appeal_Hearing_Package_Analyst_Documents where CnlyClaimNum = '" & strCnlyClaimNum & "' "
'    'End If

End Sub

Public Sub cmdApplyFilters_Click()
Dim strHearingDateTo As String
Dim strRecordCount As Integer
    
    mstrFilter = ""
    
    If Me.cmbJudge & "" <> "" Then
        mstrFilter = mstrFilter & " and ALJudgeName = '" & Me.cmbJudge & "'"
    End If
    
    If Me.cmbAnalyst & "" <> "" Then
        mstrFilter = mstrFilter & " and Analyst = '" & Me.cmbAnalyst & "'"
        
    End If
    
    If Me.cmbParticipant & "" <> "" Then
        mstrFilter = mstrFilter & " and Participant = '" & Me.cmbParticipant & "'"
    End If
    
    If Me.cmbPositionPaper & "" <> "" Then
        mstrFilter = mstrFilter & " and PP = '" & Me.cmbPositionPaper & "'"
    End If
    
    If (Nz(Me.txtHearingDateFrom, "") <> "" And Nz(Me.txtHearingDateTo, "") = "") Or (Nz(Me.txtHearingDateFrom, "") = "" And Nz(Me.txtHearingDateTo, "") <> "") Then
        MsgBox ("Please provide both a Date From and Date To value to search for Hearings")
    ElseIf (Nz(Me.txtHearingDateFrom, "") <> "" And Nz(Me.txtHearingDateTo, "") <> "") Then
        strHearingDateTo = CDate(Me.txtHearingDateTo) + 1
        mstrFilter = mstrFilter & " and (HearingDateTime >= #" & Me.txtHearingDateFrom & "# and HearingDateTime < #" & strHearingDateTo & "#) "
    End If
    
    If Me.cmbStatus & "" = "" Then
        mstrFilter = mstrFilter & " and WorkComplete <> '-1'"
    ElseIf Me.cmbStatus = "Incomplete" Then
        mstrFilter = mstrFilter & " and WorkComplete <> '-1'"
    ElseIf Me.cmbStatus = "Complete" Then
        mstrFilter = mstrFilter & " and WorkComplete = '-1'"
    ElseIf Me.cmbStatus = "All" Then
        mstrFilter = mstrFilter & " and WorkComplete in ('0', '-1', '')"
    End If
    
    
    'Set the SORT options via the UI 'Sort Claims' combo boxes
    mstrSort = ""
    If Me.cbSort1 & "" <> "" Then
        mstrSort = mstrSort & Me.cbSort1
    End If
    If Me.cbSort2 & "" <> "" Then
        If mstrSort & "" <> "" Then
            mstrSort = mstrSort & ", " & Me.cbSort2
        Else
            mstrSort = mstrSort & Me.cbSort2
        End If
    End If
    If Me.cbSort3 & "" <> "" Then
        If mstrSort & "" <> "" Then
            mstrSort = mstrSort & ", " & Me.cbSort3
        Else
            mstrSort = mstrSort & Me.cbSort3
        End If
    End If
    
    If mstrSort & "" <> "" Then
        mstrSort = "ORDER BY " & mstrSort '& ", ICN"
    Else
        mstrSort = "ORDER BY 1, 2"
    End If

    
    'Construct the SQL for the Claim List
    If mstrFilter <> "" Then
        mstrFilter = Right(mstrFilter, Len(mstrFilter) - 4)
        
        Me.frm_AppealAnalystWorkList_Claims.Form.RecordSource = "select * from v_AppealHearingAnalystWorkList_CSTemplate where " & mstrFilter & mstrSort
        Me.frm_AppealAnalystWorkList_Claims.Form.Requery
    Else
        Me.frm_AppealAnalystWorkList_Claims.Form.RecordSource = "select * from v_AppealHearingAnalystWorkList_CSTemplate " & mstrSort
        Me.frm_AppealAnalystWorkList_Claims.Form.Refresh
    End If
        
End Sub



Public Sub cmdClearFilter_Click()
  
    'Claim Filters
    mstrFilter = ""
    Me.cmbJudge = ""
    Me.cmbAnalyst = ""
    Me.cmbParticipant = ""
    Me.cmbPositionPaper = ""
    Me.txtHearingDateFrom = ""
    Me.txtHearingDateTo = ""
    Me.cmbStatus = ""
    
    'Sort Filters
    mstrSort = ""
    Me.cbSort1 = ""
    Me.cbSort2 = ""
    Me.cbSort3 = ""
    
    'ICN Search Filter
    Me.txtICN = "" 'KCF 2/22/2013 to clear the ICN search box
    
    Me.frm_AppealAnalystWorkList_Claims.Form.RecordSource = "select * from v_AppealHearingAnalystWorkList_CSTemplate where WorkComplete <> '-1' order by 1, 2"
    Me.frm_AppealAnalystWorkList_Claims.Form.Requery
    
    'Set_Filters_ComboBoxes
    
    If Me.frm_AppealAnalystWorkList_Claims.Form.RecordSet.recordCount > 0 Then
        Me.RowsSelected = 1
    Else
        Me.RowsSelected = 0
    End If

End Sub

Public Sub CnlyClaimNumToFind(strCnlyClaimNumToFind)
    Me.frm_AppealAnalystWorkList_Claims.SetFocus
    Me.frm_AppealAnalystWorkList_Claims.Form.CnlyClaimNum.SetFocus
    DoCmd.FindRecord strCnlyClaimNumToFind
    Me.SetFocus
    
Exit_Sub:
    Set myCode_ADO = Nothing
    Set cmd = Nothing
    Exit Sub
    
Err_handler:
    If Err.Number = 3265 Then Resume Exit_Sub 'Recordset error that occurs because the status is updated, but the list form is querying based upon the status
    If strErrMsg = "" Then strErrMsg = Err.Description
    MsgBox strErrMsg
    Resume Exit_Sub
    
End Sub
